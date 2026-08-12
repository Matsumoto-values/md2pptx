#!/usr/bin/env python3
"""extract_stub.py -- ブラウザが無い環境用のダミー抽出器(CI/開発用).

本番は extract.py (Playwright)。こちらは CSS 定数決め打ちの近似なので
折り返しやセル幅の正確さは保証しない。EXTRACT_BACKEND=stub で有効化される。

CLI : extract_stub.py deck.html > layout.json
API : extract_stub(html, canvas, T) -> dict
"""
import base64
import html as html_mod
import io
import json
import pathlib
import re
import sys

from PIL import Image, ImageDraw

import theme as theme_mod

SIZES = {"title": (44, 1.25, True), "text": (22, 1.6, False), "list": (24, 1.7, False),
         "raster": (20, 1.2, True), "table": (20, 1.4, False), "chart": (20, 1.2, False),
         "shape": (20, 1.5, False), "connector": (20, 1.2, False)}


def unesc(t):
    return html_mod.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def badge_png(text, w, h, T):
    img = Image.new("RGB", (w * 2, h * 2))
    d = ImageDraw.Draw(img)
    # パレットが1色のテーマは実際に存在する（デザイン層に1色しかない資料）ので
    # 2色目が無い場合は同色を使う
    pal = T["palette"]
    a, b = (tuple(int(pal[min(n, len(pal) - 1)][i:i + 2], 16) for i in (0, 2, 4))
            for n in (0, 1))
    for i in range(w * 2):
        r = i / (w * 2)
        d.line([(i, 0), (i, h * 2)],
               fill=tuple(int(a[k] + (b[k] - a[k]) * r) for k in range(3)))
    d.text((32, h - 8), text, fill=(255, 255, 255))
    buf = io.BytesIO(); img.save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode()


def build_table(inner, y, T, width):
    tb = T["table"]
    pad = tb["pad"]
    row_h = 20 * 1.4 + pad[0] + pad[2]
    grid = []
    for tr in re.findall(r"<tr>(.*?)</tr>", inner, re.S):
        grid.append([dict(text=unesc(body), th=(tag == "th"), num=("num" in attrs))
                     for tag, attrs, body
                     in re.findall(r"<(th|td)([^>]*)>(.*?)</\1>", tr, re.S)])
    ncol = max(len(r) for r in grid)
    # 列幅は最長テキスト長で按分（本番はブラウザ実測値）
    weights = [max(len(r[i]["text"]) if i < len(r) else 0 for r in grid) + 2
               for i in range(ncol)]
    col_w = [round(width * w / sum(weights)) for w in weights]
    borders = {k: dict(w=tb["border_w"], color=f"#{tb['border']}", style="solid")
               for k in "LRTB"}
    rows = []
    for j, r in enumerate(grid):
        cells = [dict(
            text=c["text"], colspan=1, rowspan=1, w=col_w[i], h=row_h, pad=pad,
            size=20, bold=c["th"], font=T["font"],
            color=f"#{tb['header_fg']}" if c["th"] else f"#{T['text']}",
            bg=f"#{tb['header_bg']}" if c["th"]
               else (f"#{tb['row_alt_bg']}" if tb["zebra"] and j % 2 == 0
                     else f"#{tb['row_bg']}"),
            align="right" if c["num"] else "left", lineHeight=1.4, borders=borders)
            for i, c in enumerate(r)]
        rows.append(dict(h=row_h, cells=cells))
    return dict(x=T["margin"]["x"], y=y, w=width, h=row_h * len(rows),
                rows=rows, colWidths=col_w, header=True)


def extract_stub(html, canvas, T):
    m, (cw, ch) = T["margin"], canvas
    width = cw - m["x"] * 2
    slides = []
    for sec in re.findall(r'<section class="slide">(.*?)</section>', html, re.S):
        elems, y = [], m["body_top"]
        for tag, kind, attrs, inner in re.findall(
                r'<(\w+) data-pptx="(\w+)"([^>]*)>(.*?)</\1>', sec, re.S):
            size, lh, bold = SIZES[kind]
            e = dict(kind=kind, font=T["font"], size=size, bold=bold, align="left",
                     lineHeight=lh, pad=T["table"]["pad"],
                     color=f"#{T['heading']}" if kind == "title"
                           else (f"#{T['text']}" if kind != "raster" else "#FFFFFF"))
            if kind == "title":
                e.update(x=m["x"], y=m["top"], w=width,
                         h=T["title"]["h"], valign="middle", text=unesc(inner))
            elif kind == "table":
                e.update(build_table(inner, y, T, width)); y += e["h"] + 18
            elif kind == "chart":
                spec = json.loads(html_mod.unescape(
                    re.search(r'data-chart="(.*?)"', attrs, re.S).group(1)))
                e.update(x=m["x"], y=y, w=1000, h=380, spec=spec); y += 398
            elif kind in ("shape", "connector"):
                # 近似: 横並び固定。flowは実測前提の機能なのでstubでは配線確認のみ
                fl = T["flow"]
                prev = [x for x in elems if x["kind"] in ("shape", "connector")]
                x = m["x"] + sum(p["w"] for p in prev)
                if kind == "shape":
                    cap = re.search(r'class="cap">(.*?)</span>', inner, re.S)
                    body = re.search(r'class="body">(.*?)</span>', inner, re.S)
                    lines = []
                    if cap:
                        lines.append(dict(text=unesc(cap.group(1)), size=fl["cap_size"],
                                          bold=True, color=f"#{fl['cap_color']}",
                                          lineHeight=1.3))
                    lines.append(dict(text=unesc(body.group(1)) if body else unesc(inner),
                                      size=fl["body_size"], bold=False,
                                      color=f"#{T['text']}", lineHeight=1.5))
                    e.update(x=x, y=y, w=fl["box_max_w"], h=96, shape="rounded_rect",
                             node=re.search(r'data-node="(\w+)"', attrs).group(1),
                             radius=fl["box_radius"], borderW=fl["box_border_w"],
                             borderColor=f"#{fl['box_border']}", bg=f"#{fl['box_bg']}",
                             pad=fl["box_pad"], lines=lines)
                else:
                    e.update(x=x, y=y + 36, w=fl["gap"], h=24,
                             **{k: re.search(rf'data-{k}="([\w-]+)"', attrs).group(1)
                                for k in ("from", "to", "dir")})
            elif kind == "raster":
                text = unesc(inner)
                e.update(x=m["x"], y=y + 8, w=min(width, 48 + len(text) * size),
                         h=size + 32, text=text)
                e["png_b64"] = badge_png(text, e["w"], e["h"], T)
                y += e["h"] + 18
            elif kind == "list":
                items = [unesc(i) for i in re.findall(r"<li>(.*?)</li>", inner, re.S)]
                h = round(size * lh * len(items)) + 20
                e.update(x=m["x"], y=y, w=width, h=h, text="\n".join(items)); y += h + 18
            else:
                h = round(size * lh)
                e.update(x=m["x"], y=y, w=width, h=h, text=unesc(inner)); y += h + 18
            elems.append(e)
        # flowの直後に続く要素のyを進める
        for e in elems:
            if e["kind"] == "shape":
                y = max(y, e["y"] + e["h"] + 18)
        slides.append({"w": cw, "h": ch, "elements": elems})
    return {"canvas": [cw, ch], "slides": slides}


def main():
    T = theme_mod.load()
    html = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    c = T["canvas"]
    json.dump(extract_stub(html, (c["w"], c["h"]), T), sys.stdout,
              ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
