#!/usr/bin/env python3
"""build_pptx.py -- レイアウト中間表現(JSON)から .pptx を組む.

CLI : build_pptx.py layout.json out.pptx [theme_id]
API : build(layout: dict, out, T: dict)   # out はパスでもBytesIOでも可

px -> EMU 変換だけが仕事。レイアウト判断は一切しない(ブラウザ側で終わっている)。
"""
import base64
import io
import json
import pathlib
import re
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import theme as theme_mod

ROOT = pathlib.Path(__file__).resolve().parent
BLANK_HINTS = ("blank", "白紙")

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
         "justify": PP_ALIGN.JUSTIFY, "start": PP_ALIGN.LEFT, "end": PP_ALIGN.RIGHT}
CHART_TYPE = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}
STACKED = {"column_stacked"}
BAR_LIKE = {"column", "column_stacked", "bar"}      # gapWidth を持てるのはこれだけ

# データラベルの位置はグラフ種ごとに許される値が違う。
# 不正な値を書くとPowerPointが「修復が必要」と言ってファイルを開けない。
#   棒(クラスタ) : ctr / inEnd / inBase / outEnd
#   棒(積み上げ) : ctr / inEnd / inBase   ← outEnd は不正
#   折れ線       : ctr / l / r / t / b     ← outEnd は不正
#   円           : bestFit / ctr / inEnd / outEnd
LABEL_POS = {
    "column": XL_LABEL_POSITION.OUTSIDE_END,
    "bar": XL_LABEL_POSITION.OUTSIDE_END,
    "column_stacked": XL_LABEL_POSITION.INSIDE_END,
    "line": XL_LABEL_POSITION.ABOVE,
    "pie": XL_LABEL_POSITION.BEST_FIT,
}
LEGEND_POS = {"bottom": XL_LEGEND_POSITION.BOTTOM, "top": XL_LEGEND_POSITION.TOP,
              "right": XL_LEGEND_POSITION.RIGHT, "left": XL_LEGEND_POSITION.LEFT}
LN_TAGS = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")   # tcPr内での正しい出現順
SHAPE_MAP = {"rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE, "rect": MSO_SHAPE.RECTANGLE,
             "diamond": MSO_SHAPE.DIAMOND, "oval": MSO_SHAPE.OVAL,
             "chevron": MSO_SHAPE.CHEVRON}


def rgb(css):
    if not css:
        return RGBColor(0, 0, 0)
    if css.startswith("#"):
        h = css.lstrip("#")
        return RGBColor.from_string(h if len(h) == 6 else "000000")
    m = re.findall(r"\d+", css)
    return RGBColor(*(int(v) for v in m[:3])) if len(m) >= 3 else RGBColor(0, 0, 0)


def hexof(css):
    c = rgb(css)
    return "{:02X}{:02X}{:02X}".format(c[0], c[1], c[2])


def is_transparent(css):
    return not css or "rgba(0, 0, 0, 0)" in css or css == "transparent"


def set_cell_borders(cell, borders, px):
    """python-pptx は罫線APIを持たないので tcPr に a:lnL/R/T/B を直接生やす.

    tcPr の子要素順序は ln* -> fill -> headers なので必ず先頭へ挿入する。
    順序を破るとPowerPointが「修復が必要」ダイアログを出す。
    """
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in LN_TAGS:
        for old in tcPr.findall(qn(tag)):
            tcPr.remove(old)
    for tag in reversed(LN_TAGS):        # 逆順に先頭挿入 -> 結果として L,R,T,B の順
        side = (borders or {}).get(tag[-1])
        if not side or side.get("style") in (None, "none", "hidden"):
            continue
        w = max(int(side["w"] * px), 3175)          # 0.25pt未満は描画されない
        ln = tcPr.makeelement(qn(tag),
                              {"w": str(w), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        fill.append(ln.makeelement(qn("a:srgbClr"), {"val": hexof(side["color"])}))
        ln.append(fill)
        if side.get("style") == "dashed":
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
        tcPr.insert(0, ln)


def bullet(paragraph):
    """python-pptx に箇条書きAPIが無いので pPr に直接ぶら下げる."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("indent", "-228600")
    pPr.set("marL", "285750")
    for tag, attrs in (("a:buFont", {"typeface": "Arial"}), ("a:buChar", {"char": "•"})):
        pPr.append(pPr.makeelement(qn(tag), attrs))


def style_run(run, e, T):
    f = run.font
    f.size = Pt(e["size"] * 0.75)          # CSS px -> pt (96dpi -> 72pt)
    f.bold = e.get("bold", False)
    # **計測に使ったフォント名ではなく theme.font を書く。**
    # font_stack の先頭は字幅互換の代替(Liberation Sans等)で、閲覧側のWindowsには
    # 存在しないことが多い。宣言すべきは元のフォント名(Arial等)。
    # 「Arialと宣言しつつLiberation Sansで測る」のがこの設計の目的
    f.name = T["font"]
    f.color.rgb = rgb(e.get("color"))


def box(e, px):
    return tuple(Emu(int(e[k] * px)) for k in ("x", "y", "w", "h"))


def add_text(slide, e, px, T):
    tf = slide.shapes.add_textbox(*box(e, px)).text_frame
    tf.word_wrap = True
    # ブラウザ側で縦中央に寄せている要素は同じにする（ヘッダ帯の中のタイトル等）
    tf.vertical_anchor = (MSO_ANCHOR.MIDDLE if e.get("valign") == "middle"
                          else MSO_ANCHOR.TOP)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate((e.get("text") or "").split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN.get(e.get("align"), PP_ALIGN.LEFT)
        p.line_spacing = e.get("lineHeight") or 1.2
        style_run(p.add_run(), e, T)
        p.runs[0].text = line
        if e["kind"] == "list":
            bullet(p)


def add_raster(slide, e, px, T):
    if not e.get("png_b64"):
        return
    slide.shapes.add_picture(io.BytesIO(base64.b64decode(e["png_b64"])), *box(e, px))


def add_table(slide, e, px, T):
    tbl = slide.shapes.add_table(len(e["rows"]), len(e["colWidths"]), *box(e, px)).table
    # テーマ由来の青い縞スタイルを殺す。これを忘れると全部青くなる
    tbl.first_row = False
    tbl.horz_banding = False
    for i, cw in enumerate(e["colWidths"]):
        tbl.columns[i].width = Emu(int(cw * px))
    for j, r in enumerate(e["rows"]):
        # 行高は「最小値」でPowerPointが勝手に伸ばす。フォント差分の保険で+8%
        tbl.rows[j].height = Emu(int(r["h"] * px * 1.08))
        for i, c in enumerate(r["cells"]):
            cell = tbl.cell(j, i)
            for attr, v in zip(("margin_top", "margin_right", "margin_bottom",
                                "margin_left"), c["pad"]):
                setattr(cell, attr, Emu(int(v * px)))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if not is_transparent(c.get("bg")):
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(c["bg"])
            set_cell_borders(cell, c.get("borders"), px)
            p = cell.text_frame.paragraphs[0]
            p.alignment = ALIGN.get(c.get("align"), PP_ALIGN.LEFT)
            style_run(p.add_run(), c, T)    # cell.text= は書式が飛ぶので使わない
            p.runs[0].text = c["text"]


def add_autoshape(slide, e, px, T):
    """フローチャートのボックス。ネイティブのオートシェイプなので手で編集できる."""
    sh = slide.shapes.add_shape(SHAPE_MAP.get(e.get("shape"), MSO_SHAPE.ROUNDED_RECTANGLE),
                               *box(e, px))
    # 角の丸みをCSSのborder-radiusに合わせる(既定値は箱の高さの16.7%で丸すぎる)
    try:
        if e.get("radius") and len(sh.adjustments):
            sh.adjustments[0] = min(0.5, e["radius"] / min(e["w"], e["h"]))
    except (IndexError, ValueError):
        pass

    if is_transparent(e.get("bg")):
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(e["bg"])
    if e.get("borderW"):
        sh.line.color.rgb = rgb(e.get("borderColor"))
        sh.line.width = Pt(e["borderW"] * 0.75)
    else:
        sh.line.fill.background()

    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pad = e.get("pad") or [0, 0, 0, 0]
    for attr, v in zip(("margin_top", "margin_right", "margin_bottom", "margin_left"),
                       pad):
        setattr(tf, attr, Emu(int(v * px)))
    for i, ln in enumerate(e.get("lines") or [{"text": e.get("text", "")}]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN.get(e.get("align"), PP_ALIGN.LEFT)
        p.line_spacing = ln.get("lineHeight") or 1.3
        style_run(p.add_run(), {**e, **ln}, T)
        p.runs[0].text = ln["text"]
    return sh


def add_quote(slide, e, px, T):
    """引用（逐語）。背景の矩形と左端の色帯、その上にテキストを置く.

    左の色帯は borderLeft の実測値から別図形として作る。
    オートシェイプに「左だけ線」という指定は無いため。
    図形は編集できるネイティブ要素として置く（画像化しない）。
    """
    box_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *box(e, px))
    if is_transparent(e.get("bg")):
        box_shape.fill.background()
    else:
        box_shape.fill.solid()
        box_shape.fill.fore_color.rgb = rgb(e["bg"])
    box_shape.line.fill.background()
    box_shape.shadow.inherit = False          # 既定の影を消す

    bw = e.get("borderW") or 0
    if bw:
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Emu(int(e["x"] * px)), Emu(int(e["y"] * px)),
            Emu(int(bw * px)), Emu(int(e["h"] * px)))
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(e.get("borderColor"))
        bar.line.fill.background()
        bar.shadow.inherit = False

    tf = box_shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pad = e.get("pad") or [0, 0, 0, 0]
    for attr, v in zip(("margin_top", "margin_right", "margin_bottom",
                        "margin_left"), pad):
        setattr(tf, attr, Emu(int(v * px)))
    tf.margin_left = Emu(int((pad[3] + bw) * px))   # 色帯の分だけ内側へ寄せる
    for i, ln in enumerate(e.get("lines") or []):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = ln.get("lineHeight") or 1.5
        style_run(p.add_run(), {**e, **ln}, T)
        p.runs[0].text = ln["text"]
    return box_shape


def add_flow_connector(slide, e, px, T, nodes):
    """図形に「接続」した矢印。接続しておくと、PowerPointで箱を動かしても追従する.

    単に線を置くだけでは動かした瞬間に矢印が取り残されるため、
    手で調整する用途では begin_connect / end_connect が必須。
    """
    a, b = nodes.get(e.get("from")), nodes.get(e.get("to"))
    if not (a and b):
        return
    if e.get("dir") == "vertical":
        bi, ei = 2, 0                      # 下辺 -> 上辺
        x1, y1 = a.left + a.width // 2, a.top + a.height
        x2, y2 = b.left + b.width // 2, b.top
    else:
        bi, ei = 3, 1                      # 右辺 -> 左辺
        x1, y1 = a.left + a.width, a.top + a.height // 2
        x2, y2 = b.left, b.top + b.height // 2

    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    conn.begin_connect(a, bi)
    conn.end_connect(b, ei)

    cfg = T["flow"]
    conn.line.color.rgb = rgb(f"#{cfg['connector_color']}")
    conn.line.width = Pt(cfg["connector_w"] * 0.75)
    # 矢印の先端はpython-pptxにAPIが無いのでXMLを直挿しする。
    # a:ln の子要素順序は fill -> prstDash -> ... -> headEnd -> tailEnd なので最後に足す
    ln = conn.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"),
                             {"type": "triangle", "w": "med", "len": "med"}))
    return conn


def add_chart(slide, e, px, T):
    spec, cfg = e["spec"], T["chart"]
    ctype = spec.get("type", "column")
    cd = CategoryChartData()
    cd.categories = spec["categories"]
    for s in spec["series"]:
        cd.add_series(s["name"], tuple(s["values"]))
    chart = slide.shapes.add_chart(
        CHART_TYPE.get(ctype, XL_CHART_TYPE.COLUMN_CLUSTERED), *box(e, px), cd).chart

    chart.font.size = Pt(12)
    chart.font.color.rgb = rgb(f"#{T['text']}")
    plot = chart.plots[0]
    # 円グラフは系列が1つでカテゴリごとに色を変える必要がある。
    # それ以外で True にすると単系列が虹色になる
    plot.vary_by_categories = (ctype == "pie")
    if ctype in BAR_LIKE:                  # gapWidth は棒グラフ専用の要素
        plot.gap_width = cfg["gap_width"]
    plot.has_data_labels = cfg["data_labels"]
    if cfg["data_labels"]:
        dl = plot.data_labels
        dl.number_format = T["number_format"]
        dl.number_format_is_linked = False
        # ラベル位置はグラフ種ごとに許される値が違う。不正だとファイルが壊れる
        dl.position = LABEL_POS.get(ctype, XL_LABEL_POSITION.CENTER)

    pal = T["palette"]
    for i, ser in enumerate(chart.series):      # 系列数がパレットより多い場合は循環
        color = RGBColor.from_string(pal[i % len(pal)])
        if ctype == "line":
            # 折れ線は塗りではなく線に色を付ける。マーカーも揃える
            ser.format.line.color.rgb = color
            ser.format.line.width = Pt(2.25)
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = color
            ser.marker.format.line.color.rgb = color
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = color

    chart.has_legend = len(chart.series) > 1 and cfg["legend"] not in (None, "none")
    if chart.has_legend:
        chart.legend.position = LEGEND_POS.get(cfg["legend"], XL_LEGEND_POSITION.BOTTOM)
        chart.legend.include_in_layout = False   # 忘れると凡例がプロット領域を食う

    if ctype != "pie":
        va = chart.value_axis
        va.has_major_gridlines = cfg["gridlines"]
        if cfg["gridlines"]:
            va.major_gridlines.format.line.color.rgb = rgb(f"#{T['grid']}")
        va.tick_labels.number_format = T["number_format"]
        va.tick_labels.number_format_is_linked = False
        chart.category_axis.format.line.color.rgb = rgb(f"#{T['axis']}")


DISPATCH = {"title": add_text, "text": add_text, "list": add_text,
            "raster": add_raster, "table": add_table, "chart": add_chart,
            "quote": add_quote}


def clear_slides(prs):
    """既存スライドを全削除する.

    python-pptx に削除APIが無いので sldIdLst と関連付けを直接操作する。
    サンプル資料をベースに使う場合これが無いと、既存の何十枚かの後ろに
    生成分が追記されてしまう。
    """
    lst = prs.slides._sldIdLst
    for sldId in list(lst):
        prs.part.drop_rel(sldId.rId)
        lst.remove(sldId)


def pick_layout(prs, name):
    """レイアウトを名前で選ぶ。添字での決め打ちはテンプレート次第で破綻する.

    prs.slide_layouts は先頭マスターの分しか返さないため全マスターを見る。
    マスターが複数ある資料(Google Slides由来など)では、実際に多用されている
    レイアウトが2番目以降のマスターにいることがある。
    """
    layouts = [lo for m in prs.slide_masters for lo in m.slide_layouts]
    if name:
        for lo in layouts:
            if lo.name == name:
                return lo
        print(f"WARNING: layout '{name}' not found; falling back", file=sys.stderr)
    for lo in layouts:
        if any(h in (lo.name or "").lower() for h in BLANK_HINTS):
            return lo
    return layouts[6] if len(layouts) > 6 else layouts[-1]


def open_base(T):
    """テンプレートをベースに開く。指定が無ければ素の新規プレゼンテーション."""
    base = T.get("base_pptx")
    if not base:
        prs = Presentation()
        w_in, h_in = T["canvas"]["slide_in"]
        prs.slide_width, prs.slide_height = Inches(w_in), Inches(h_in)
        return prs
    path = pathlib.Path(base)
    if not path.is_absolute():
        path = ROOT / path
    if not path.is_file():
        raise FileNotFoundError(f"base_pptx not found: {path}")
    prs = Presentation(str(path))
    clear_slides(prs)
    # スライドサイズはテンプレート側が正。テーマの canvas と食い違う場合は警告する
    tw, th = prs.slide_width // 9525, prs.slide_height // 9525
    if (tw, th) != (T["canvas"]["w"], T["canvas"]["h"]):
        print(f"WARNING: canvas {T['canvas']['w']}x{T['canvas']['h']} != "
              f"template {tw}x{th}; using template size", file=sys.stderr)
    return prs


def strip_placeholders(slide):
    """継承した空のプレースホルダを削除する.

    全要素を絶対配置するのでプレースホルダは使わない。空のまま残すと
    編集時にプロンプト枠が表示される。
    """
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)


def build(layout, out, T):
    prs = open_base(T)
    # md2html(CSS) と extract(viewport) が同じキャンバスを見ているかを確認する。
    # ここが食い違うと px->EMU が二重スケールして全体がズレる
    if layout["canvas"][0] != T["canvas"]["w"]:
        print(f"WARNING: layout canvas {layout['canvas'][0]}px != theme canvas "
              f"{T['canvas']['w']}px; extract側にテーマを渡し忘れていない?",
              file=sys.stderr)
    # px -> EMU はテンプレートの実寸から算出する(テーマの値ではなく実物が正)
    px = prs.slide_width / layout["canvas"][0]
    blank = pick_layout(prs, T.get("base_layout"))
    for s in layout["slides"]:
        slide = prs.slides.add_slide(blank)
        if T.get("strip_placeholders", True):
            strip_placeholders(slide)
        # コネクタは接続先の図形が必要なので2パスにする
        nodes, connectors = {}, []
        for e in s["elements"]:
            if e["kind"] == "shape":
                nodes[e.get("node")] = add_autoshape(slide, e, px, T)
            elif e["kind"] == "connector":
                connectors.append(e)
            else:
                fn = DISPATCH.get(e["kind"])
                if fn:
                    fn(slide, e, px, T)
        for e in connectors:
            add_flow_connector(slide, e, px, T, nodes)
    prs.save(out)
    return out


def main():
    layout = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
    T = theme_mod.load(sys.argv[3] if len(sys.argv) > 3 else "default")
    build(layout, sys.argv[2], T)
    print(f"wrote {sys.argv[2]} ({len(layout['slides'])} slides)", file=sys.stderr)


if __name__ == "__main__":
    main()
