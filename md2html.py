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

import theme as theme_mod


def css(T):
    tb, c = T["table"], T["canvas"]
    m = T["margin"]
    var = "".join(f"--pal{i}:#{p};" for i, p in enumerate(T["palette"]))
    zebra = (f'[data-pptx="table"] tbody tr:nth-child(even) td '
             f'{{ background:#{tb["row_alt_bg"]}; }}') if tb["zebra"] else ""
    return f"""
:root {{ {var} --text:#{T['text']}; --head:#{T['heading']}; }}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ background:#4a4a4a; font-family:'{T['font']}','Helvetica Neue',sans-serif; }}
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


def render_chart(spec_text):
    spec = yaml.safe_load(spec_text)
    attr = html_mod.escape(json.dumps(spec, ensure_ascii=False), quote=True)
    # canvas はプレビュー用の使い捨て。PPTX側は data-chart の数値だけを見る
    return (f'<div data-pptx="chart" data-chart="{attr}">'
            f'<canvas class="preview"></canvas></div>')


def render_slide(lines):
    out, buf_list, buf_tbl, buf_chart = [], [], [], None
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
                out.append(render_chart("\n".join(buf_chart)))
                buf_chart = None
            else:
                buf_chart.append(line)
            continue
        if s.startswith("```chart"):
            flush(); buf_chart = []
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
    body = "\n".join(render_slide(s) for s in parse_slides(markdown))
    return f"<!doctype html><meta charset='utf-8'><style>{css(T)}</style>\n{body}"


def main():
    md = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    T = theme_mod.load(sys.argv[2] if len(sys.argv) > 2 else "default")
    print(to_html(md, T))


if __name__ == "__main__":
    main()
