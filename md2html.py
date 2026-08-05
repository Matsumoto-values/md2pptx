#!/usr/bin/env python3
"""md2html.py -- Markdown を「PPTX変換コントラクト」準拠のHTMLスライドに変換する.

CLI : md2html.py deck.md [theme_id] > deck.html
API : to_html(markdown: str, T: dict) -> str

コントラクト:
  <section class="slide">          : スライド1枚
  data-pptx="title|text|list"      : ネイティブテキストボックス
  data-pptx="table"                : ネイティブ表(セル座標はブラウザ実測)
  data-pptx="chart" data-chart=... : ネイティブグラフ(数値は属性のJSONが正)
  data-pptx="raster"               : 要素スクショを画像として貼る
  注釈なしの要素                   : 無視

md記法: "# "=タイトル / "- "=箇条書き / "!!"=raster / "| a | b |"=表
        "```chart" ... "```" =グラフ / "---"=スライド区切り
"""
import html as html_mod
import json
import pathlib
import sys

import yaml

import chart_svg
import theme as theme_mod


def css(T):
    tb, c = T["table"], T["canvas"]
    m = T["margin"]
    stack = ",".join(f"'{f}'" if " " in f else f for f in T["font_stack"])
    fl = T["flow"]
    var = "".join(f"--pal{i}:#{p};" for i, p in enumerate(T["palette"]))
    zebra = (f'[data-pptx="table"] tbody tr:nth-child(even) td '
             f'{{ background:#{tb["row_alt_bg"]}; }}') if tb["zebra"] else ""
    return f"""
:root {{ {var} --text:#{T['text']}; --head:#{T['heading']}; }}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#4a4a4a; font-family:{stack}; }}
.slide {{ position:relative; width:{c['w']}px; height:{c['h']}px;
  background:#fff; overflow:hidden; margin:0 auto 24px; }}
.slide > .body {{ position:absolute; left:{m['x']}px; top:{m['body_top']}px;
  width:{c['w'] - m['x'] * 2}px; }}
[data-pptx="title"] {{ position:absolute; left:{m['x']}px; top:{m['top']}px;
  width:{c['w'] - m['x'] * 2}px;
  font-size:44px; font-weight:700; color:var(--head); line-height:1.25; }}
[data-pptx="text"] {{ font-size:22px; line-height:1.6; color:var(--text); margin-bottom:18px; }}
[data-pptx="list"] {{ font-size:24px; line-height:1.7; color:var(--text); padding-left:1.2em; }}
[data-pptx="list"] li {{ margin-bottom:10px; }}
[data-pptx="raster"] {{ display:inline-block; padding:16px 24px; border-radius:12px;
  background:linear-gradient(135deg,var(--pal0),var(--pal1));
  color:#fff; font-size:20px; font-weight:700; }}
[data-pptx="table"] {{ border-collapse:collapse; width:100%; margin-bottom:18px;
  font-size:20px; color:var(--text); }}
[data-pptx="table"] th, [data-pptx="table"] td {{
  padding:{tb['pad'][0]}px {tb['pad'][1]}px {tb['pad'][2]}px {tb['pad'][3]}px;
  text-align:left; border:{tb['border_w']}px solid #{tb['border']}; }}
[data-pptx="table"] th {{ background:#{tb['header_bg']}; color:#{tb['header_fg']};
  font-weight:700; }}
{zebra}
[data-pptx="table"] td.num, [data-pptx="table"] th.num {{ text-align:right; }}
[data-pptx="chart"] {{ width:1000px; height:380px; margin-bottom:18px; }}

/* フローチャート: flexで並べてブラウザに寸法を測らせる。
   ボックスは中の文字量に応じて自動サイズになる（python-pptx直書きでは不可能な部分） */
.flow {{ display:flex; align-items:stretch; margin-bottom:18px; }}
.flow-vertical {{ flex-direction:column; align-items:flex-start; }}
[data-pptx="shape"] {{ display:flex; flex-direction:column; justify-content:center;
  padding:{fl['box_pad'][0]}px {fl['box_pad'][1]}px {fl['box_pad'][2]}px {fl['box_pad'][3]}px;
  border-radius:{fl['box_radius']}px; background:#{fl['box_bg']};
  border:{fl['box_border_w']}px solid #{fl['box_border']};
  min-width:{fl['box_min_w']}px; max-width:{fl['box_max_w']}px; }}
.flow-vertical [data-pptx="shape"] {{ width:{fl['box_w_vertical']}px; max-width:none; }}
[data-pptx="shape"] .cap {{ font-size:{fl['cap_size']}px; font-weight:700;
  color:#{fl['cap_color']}; margin-bottom:4px; }}
[data-pptx="shape"] .body {{ font-size:{fl['body_size']}px; line-height:1.5;
  color:var(--text); }}
[data-pptx="connector"] {{ flex:0 0 {fl['gap']}px; display:flex;
  align-items:center; justify-content:center; }}
.flow-vertical [data-pptx="connector"] {{ flex-basis:44px;
  width:{fl['box_w_vertical']}px; }}
/* 菱形は内接矩形が小さいので余白を増やす。ブラウザ上の表示は矩形のままだが、
   PPTX側で菱形になった時に文字がはみ出さない寸法をここで確保する */
.shape-diamond {{ padding:{fl['box_pad'][0] * 2}px {fl['box_pad'][1] * 2}px;
  text-align:center; align-items:center; }}
"""


def parse_slides(md):
    slides, cur = [], []
    for line in md.splitlines():
        if line.strip() == "---":
            slides.append(cur); cur = []
        else:
            cur.append(line)
    slides.append(cur)
    return [s for s in slides if any(l.strip() for l in s)]


def is_num(v):
    return v.replace(",", "").replace(".", "").replace("-", "").replace("%", "").isdigit()


def render_table(rows):
    """rows[0]=ヘッダ, rows[1]は区切り行なので捨てる."""
    head, body = rows[0], rows[2:]
    ncol = len(head)
    # 数値列の判定はデータ行で行う。ヘッダ文字("Q1"等)で判定すると揃えがズレる
    numeric = [all(is_num(r[i]) for r in body if i < len(r)) and bool(body)
               for i in range(ncol)]

    def cell(tag, text, i):
        cls = ' class="num"' if i < ncol and numeric[i] else ""
        return f"<{tag}{cls}>{html_mod.escape(text)}</{tag}>"

    ths = "".join(cell("th", c, i) for i, c in enumerate(head))
    trs = "".join("<tr>" + "".join(cell("td", c, i) for i, c in enumerate(r)) + "</tr>"
                  for r in body)
    return (f'<table data-pptx="table" data-pptx-header="1">'
            f"<thead><tr>{ths}</tr></thead><tbody>{trs}</tbody></table>")


def render_chart(spec_text, T):
    spec = yaml.safe_load(spec_text)
    attr = html_mod.escape(json.dumps(spec, ensure_ascii=False), quote=True)
    # 中のSVGはプレビュー用の使い捨て。PPTX側は data-chart の数値だけを見る。
    # 実測されるのは外側のdivの矩形だけなので、中身の描画は成果物に影響しない
    return (f'<div data-pptx="chart" data-chart="{attr}">'
            f'{chart_svg.render(spec, T)}</div>')


def arrow_svg(direction, T):
    """プレビュー用の矢印。PPTX側は接続付きネイティブコネクタとして作り直される."""
    c, w = T["flow"]["connector_color"], T["flow"]["connector_w"]
    if direction == "vertical":
        return (f'<svg width="24" height="40" viewBox="0 0 24 40">'
                f'<line x1="12" y1="2" x2="12" y2="28" stroke="#{c}" '
                f'stroke-width="{w}"/>'
                f'<path d="M12,38 L6,26 L18,26 Z" fill="#{c}"/></svg>')
    return (f'<svg width="48" height="24" viewBox="0 0 48 24">'
            f'<line x1="2" y1="12" x2="34" y2="12" stroke="#{c}" stroke-width="{w}"/>'
            f'<path d="M46,12 L32,6 L32,18 Z" fill="#{c}"/></svg>')


def render_flow(spec_text, T):
    """```flow ブロック -> flexで並んだボックス列 + コネクタ.

    記法: 1行1ボックス。"見出し: 本文" と書くと見出しが小さく太字で乗る。
          先頭行に "direction: vertical" を書くと縦並びになる。
    """
    direction = T["flow"]["direction"]
    nodes = []
    for line in spec_text.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith("direction:"):
            direction = s.split(":", 1)[1].strip()
            continue
        shape = "rounded_rect"
        if s.startswith("?"):
            shape, s = "diamond", s[1:].strip()   # 行頭"?"は判断(菱形)
        cap, _, body = s.partition(":")
        nodes.append((cap.strip(), body.strip(), shape) if body.strip()
                     else (None, cap.strip(), shape))

    parts = []
    for i, (cap, body, shape) in enumerate(nodes):
        nid = f"n{i + 1}"
        inner = (f'<span class="cap">{html_mod.escape(cap)}</span>' if cap else "")
        inner += f'<span class="body">{html_mod.escape(body)}</span>'
        parts.append(f'<div data-pptx="shape" data-shape="{shape}" '
                     f'class="shape-{shape}" data-node="{nid}">{inner}</div>')
        if i < len(nodes) - 1:
            parts.append(f'<div data-pptx="connector" data-from="{nid}" '
                         f'data-to="n{i + 2}" data-dir="{direction}">'
                         f'{arrow_svg(direction, T)}</div>')
    return f'<div class="flow flow-{direction}">{"".join(parts)}</div>'


def render_slide(lines, T):
    out, buf_list, buf_tbl, buf_chart = [], [], [], None
    buf_flow = None
    title = ""

    def flush():
        if buf_list:
            items = "".join(f"<li>{html_mod.escape(i)}</li>" for i in buf_list)
            out.append(f'<ul data-pptx="list">{items}</ul>')
            buf_list.clear()
        if buf_tbl:
            out.append(render_table(list(buf_tbl)))
            buf_tbl.clear()

    for line in lines:
        s = line.strip()
        if buf_chart is not None:
            if s == "```":
                out.append(render_chart("\n".join(buf_chart), T))
                buf_chart = None
            else:
                buf_chart.append(line)
            continue
        if buf_flow is not None:
            if s == "```":
                out.append(render_flow("\n".join(buf_flow), T))
                buf_flow = None
            else:
                buf_flow.append(line)
            continue
        if s.startswith("```chart"):
            flush(); buf_chart = []
        elif s.startswith("```flow"):
            flush(); buf_flow = []
        elif not s:
            flush()
        elif s.startswith("# "):
            title = s[2:]
        elif s.startswith("|"):
            buf_tbl.append([c.strip() for c in s.strip("|").split("|")])
        elif s.startswith("- "):
            if buf_tbl: flush()
            buf_list.append(s[2:])
        elif s.startswith("!!"):
            flush()
            out.append(f'<div data-pptx="raster">{html_mod.escape(s[2:].strip())}</div>')
        else:
            flush()
            out.append(f'<p data-pptx="text">{html_mod.escape(s)}</p>')
    flush()

    head = f'<h1 data-pptx="title">{html_mod.escape(title)}</h1>' if title else ""
    return f'<section class="slide">{head}<div class="body">{"".join(out)}</div></section>'


def to_html(markdown, T):
    body = "\n".join(render_slide(s, T) for s in parse_slides(markdown))
    return f"<!doctype html><meta charset='utf-8'><style>{css(T)}</style>\n{body}"


def main():
    md = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    T = theme_mod.load(sys.argv[2] if len(sys.argv) > 2 else "default")
    print(to_html(md, T))


if __name__ == "__main__":
    main()
