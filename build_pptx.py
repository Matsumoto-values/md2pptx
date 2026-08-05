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
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import theme as theme_mod

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
LEGEND_POS = {"bottom": XL_LEGEND_POSITION.BOTTOM, "top": XL_LEGEND_POSITION.TOP,
              "right": XL_LEGEND_POSITION.RIGHT, "left": XL_LEGEND_POSITION.LEFT}
LN_TAGS = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")   # tcPr内での正しい出現順


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
    f.name = e.get("font") or T["font"]
    f.color.rgb = rgb(e.get("color"))


def box(e, px):
    return tuple(Emu(int(e[k] * px)) for k in ("x", "y", "w", "h"))


def add_text(slide, e, px, T):
    tf = slide.shapes.add_textbox(*box(e, px)).text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
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
    plot.vary_by_categories = False        # 単系列が虹色になる悲劇を防ぐ
    if ctype != "pie":
        plot.gap_width = cfg["gap_width"]
    plot.has_data_labels = cfg["data_labels"]
    if cfg["data_labels"]:
        dl = plot.data_labels
        dl.number_format = T["number_format"]
        dl.number_format_is_linked = False
        # 積み上げで OUTSIDE_END は不正値 -> ファイルが壊れる
        dl.position = (XL_LABEL_POSITION.INSIDE_END if ctype in STACKED
                       else XL_LABEL_POSITION.OUTSIDE_END)

    for ser, color in zip(chart.series, T["palette"]):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(color)

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
            "raster": add_raster, "table": add_table, "chart": add_chart}


def build(layout, out, T):
    w_in, h_in = T["canvas"]["slide_in"]
    px = Inches(w_in) / layout["canvas"][0]      # 1 CSS px あたりの EMU

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(w_in), Inches(h_in)
    blank = prs.slide_layouts[6]
    for s in layout["slides"]:
        slide = prs.slides.add_slide(blank)
        for e in s["elements"]:
            fn = DISPATCH.get(e["kind"])
            if fn:
                fn(slide, e, px, T)
    prs.save(out)
    return out


def main():
    layout = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
    T = theme_mod.load(sys.argv[3] if len(sys.argv) > 3 else "default")
    build(layout, sys.argv[2], T)
    print(f"wrote {sys.argv[2]} ({len(layout['slides'])} slides)", file=sys.stderr)


if __name__ == "__main__":
    main()
