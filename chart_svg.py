#!/usr/bin/env python3
"""chart_svg.py -- HTMLプレビュー用のグラフをインラインSVGで描く.

これは**プレビュー専用**。PPTX側はネイティブグラフとして再描画されるので、
ここでの描画結果が成果物に影響することはない（外枠の座標だけが実測対象）。

CDNやJSに依存しないのは意図的:
  - Cloud Runのコンテナが実行時に外部ネットワークへ出るのを避ける
  - オフラインでも同じプレビューが再現できる
  - トークンのパレットを共有してPPTXと見た目を揃える
"""
import html

W, H = 1000, 380
PAD = {"l": 74, "r": 24, "t": 24, "b": 56}     # b は凡例とカテゴリ名の分


def nice_max(v):
    """目盛の上限を切りのいい値に丸める."""
    if v <= 0:
        return 1
    import math
    exp = math.floor(math.log10(v))
    base = 10 ** exp
    for m in (1, 2, 2.5, 5, 10):
        if v <= m * base:
            return m * base
    return 10 * base


def fmt(v):
    return f"{v:,.0f}" if abs(v - round(v)) < 0.05 else f"{v:,.1f}"


def esc(s):
    return html.escape(str(s))


def _axes(out, plot, vmax, ticks, grid, axis):
    x0, y0, w, h = plot
    for i in range(ticks + 1):
        v = vmax * i / ticks
        y = y0 + h - h * i / ticks
        out.append(f'<line x1="{x0}" y1="{y:.1f}" x2="{x0 + w}" y2="{y:.1f}" '
                   f'stroke="#{grid}" stroke-width="1"/>')
        out.append(f'<text x="{x0 - 10}" y="{y + 4:.1f}" text-anchor="end" '
                   f'font-size="13">{fmt(v)}</text>')
    out.append(f'<line x1="{x0}" y1="{y0 + h}" x2="{x0 + w}" y2="{y0 + h}" '
               f'stroke="#{axis}" stroke-width="1"/>')


def _legend(out, names, palette, y):
    # 中央寄せの凡例。文字幅は概算で足りる(プレビューなので厳密さは不要)
    items = [(n, palette[i % len(palette)]) for i, n in enumerate(names)]
    widths = [14 + 6 + len(n) * 12 + 18 for n, _ in items]
    x = (W - sum(widths)) / 2
    for (n, c), wd in zip(items, widths):
        out.append(f'<rect x="{x:.1f}" y="{y - 9}" width="12" height="12" fill="#{c}"/>')
        out.append(f'<text x="{x + 18:.1f}" y="{y + 1}" font-size="13">{esc(n)}</text>')
        x += wd


def render(spec, T):
    kind = spec.get("type", "column")
    cats = [str(c) for c in spec["categories"]]
    series = spec["series"]
    pal, cfg = T["palette"], T["chart"]
    show_legend = len(series) > 1 and cfg["legend"] not in (None, "none")
    stack = kind == "column_stacked"

    x0, y0 = PAD["l"], PAD["t"]
    w = W - PAD["l"] - PAD["r"]
    h = H - PAD["t"] - (PAD["b"] if show_legend else PAD["b"] - 26)
    out = [f'<svg viewBox="0 0 {W} {H}" width="100%" height="100%" '
           f'xmlns="http://www.w3.org/2000/svg" fill="#{T["text"]}" '
           f'font-family="{T["font_stack"][0]}">']

    if kind == "pie":
        out += _pie(series, cats, pal, T)
    else:
        if stack:
            vmax = nice_max(max(sum(s["values"][i] for s in series)
                                for i in range(len(cats))))
        else:
            vmax = nice_max(max(max(s["values"]) for s in series))
        _axes(out, (x0, y0, w, h), vmax, 5, T["grid"], T["axis"])

        cw = w / len(cats)
        gap = cfg["gap_width"] / 100                     # PPTXのgap_widthと同じ意味
        bw = cw / (1 + gap) / (1 if stack else len(series))
        for ci, c in enumerate(cats):
            cx = x0 + cw * ci + cw / 2
            out.append(f'<text x="{cx:.1f}" y="{y0 + h + 20}" text-anchor="middle" '
                       f'font-size="13">{esc(c)}</text>')
            acc = 0
            for si, s in enumerate(series):
                v = s["values"][ci]
                bh = h * v / vmax
                if kind == "line":
                    continue
                bx = (cx - bw / 2 if stack
                      else cx - bw * len(series) / 2 + bw * si)
                by = y0 + h - bh - (h * acc / vmax if stack else 0)
                out.append(f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bw:.1f}" '
                           f'height="{bh:.1f}" fill="#{pal[si % len(pal)]}"/>')
                if cfg["data_labels"]:
                    ly = by + 15 if stack else by - 6
                    out.append(f'<text x="{bx + bw / 2:.1f}" y="{ly:.1f}" '
                               f'text-anchor="middle" font-size="12">{fmt(v)}</text>')
                acc += v

        if kind == "line":
            for si, s in enumerate(series):
                pts = [(x0 + cw * i + cw / 2, y0 + h - h * v / vmax)
                       for i, v in enumerate(s["values"])]
                d = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
                out.append(f'<polyline points="{d}" fill="none" '
                           f'stroke="#{pal[si % len(pal)]}" stroke-width="2.5"/>')
                for (px, py), v in zip(pts, s["values"]):
                    out.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="4" '
                               f'fill="#{pal[si % len(pal)]}"/>')
                    if cfg["data_labels"]:
                        out.append(f'<text x="{px:.1f}" y="{py - 10:.1f}" '
                                   f'text-anchor="middle" font-size="12">{fmt(v)}</text>')

    if show_legend:
        _legend(out, [s["name"] for s in series], pal, H - 14)
    out.append("</svg>")
    return "".join(out)


def _pie(series, cats, pal, T):
    import math
    vals = series[0]["values"]
    total = sum(vals) or 1
    cx, cy, r = W / 2, H / 2 - 10, min(W, H) / 2 - 60
    out, a0 = [], -math.pi / 2
    for i, v in enumerate(vals):
        a1 = a0 + 2 * math.pi * v / total
        large = 1 if (a1 - a0) > math.pi else 0
        x1, y1 = cx + r * math.cos(a0), cy + r * math.sin(a0)
        x2, y2 = cx + r * math.cos(a1), cy + r * math.sin(a1)
        out.append(f'<path d="M{cx:.1f},{cy:.1f} L{x1:.1f},{y1:.1f} '
                   f'A{r:.1f},{r:.1f} 0 {large} 1 {x2:.1f},{y2:.1f} Z" '
                   f'fill="#{pal[i % len(pal)]}"/>')
        am = (a0 + a1) / 2
        lx, ly = cx + (r + 24) * math.cos(am), cy + (r + 24) * math.sin(am)
        out.append(f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
                   f'font-size="13">{esc(cats[i])} {fmt(v)}</text>')
        a0 = a1
    return out
