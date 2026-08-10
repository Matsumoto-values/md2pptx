#!/usr/bin/env python3
"""pptx2theme.py -- 既存の会社資料/テンプレートからデザイントークンを実測抽出する.

usage: python3 tools/pptx2theme.py corp.pptx > themes/corp.json

**テーマ部品(a:clrScheme / a:fontScheme)は読まない。**
実際の営業資料2件を調べた結果、テーマ部品はOffice既定のまま放置されており
(accent1=4472C4 / Calibri)、ブランドデザインは図形の直接書式に入っていた。
テーマ部品を信じると generic な Office カラーを拾って完全に外す。

そのため配色・フォント・余白はすべて実スライドからの実測で決める。
出力JSONは人間のレビューを前提とする(特に font のLinux側での解決可否)。
"""
import collections
import colorsys
import json
import pathlib
import sys

from pptx import Presentation
from pptx.enum.dml import MSO_FILL

# 無彩色(白・黒・グレー)はブランド色ではないので候補から外す。
# 明度の上限は設けない: FFC000のような明度1.0の色は正当なブランド色であり、
# 白は彩度0なのでMIN_SATで既に除外される
MIN_SAT, MIN_VAL = 0.18, 0.12
COLOR_DIST = 60          # これ以下の距離は同じ色と見なす(濃淡の重複を防ぐ)
LAYOUT_HINTS = ("default", "content", "blank", "白紙", "本文", "ユーザー設定")

# 会社テンプレートのフォントはLinux側に無いことが多い。字幅が一致する代替が
# あるものは font_stack に入れる（pptxに書く名前は元のフォントのまま）。
# メトリクス互換なら、ブラウザでの実測値がPowerPoint側でもそのまま通用する。
METRIC_COMPATIBLE = {
    "arial": "Liberation Sans",
    "helvetica": "Liberation Sans",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
    "calibri": "Carlito",
    "cambria": "Caladea",
}
# 日本語フォントには字幅互換の代替が存在しない。BIZ UDPGothicは
# 「Windows標準搭載 かつ Linux導入可能」という条件を満たす唯一の実用解なので
# フォールバックには置くが、元フォントと字幅は違う（要注意として報告する）
JP_FALLBACK = "BIZ UDPGothic"

# 日本語を持たないフォント。資料が日本語主体なのにこれを選ぶと本文が化ける。
# (run.font.name はラテン用の指定を返すため、Latinフォントが最頻出になりやすい)
LATIN_ONLY = ("arial", "helvetica", "calibri", "cambria", "times new roman",
              "courier new", "lexend", "roboto", "open sans", "lato",
              "montserrat", "verdana", "tahoma", "segoe ui")


def font_stack_for(font):
    """(font_stack, 警告文) を返す."""
    sub = METRIC_COMPATIBLE.get(font.strip().lower())
    if sub:
        return [sub, font, JP_FALLBACK, "Noto Sans CJK JP", "sans-serif"], (
            f"'{font}' は字幅互換の '{sub}' で計測する。"
            f"Linux側に fonts-liberation 等が必要")
    return [font, JP_FALLBACK, "Noto Sans CJK JP", "sans-serif"], (
        f"'{font}' に字幅互換の代替が無い。Linux側で解決できない場合、"
        f"実測値がPowerPointで通用しない（extract.pyが警告を出す）")


def rgb_of(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def hsv_of(h):
    r, g, b = (v / 255 for v in rgb_of(h))
    return colorsys.rgb_to_hsv(r, g, b)


def is_brandish(h):
    _, s, v = hsv_of(h)
    return s >= MIN_SAT and v >= MIN_VAL


def far_enough(h, chosen):
    a = rgb_of(h)
    return all(sum((a[i] - rgb_of(c)[i]) ** 2 for i in range(3)) ** 0.5 > COLOR_DIST
               for c in chosen)


def solid_fills(shapes, counter, area):
    for sh in shapes:
        try:
            fc = sh.fill
            if fc.type == MSO_FILL.SOLID and str(fc.fore_color.type).startswith("RGB"):
                c = str(fc.fore_color.rgb)
                counter[c] += 1
                if sh.width and sh.height:
                    area[c] += (sh.width // 9525) * (sh.height // 9525)
        except Exception:      # noqa: BLE001  塗りを持たない図形は多い
            pass


def all_layouts(prs):
    """全マスターのレイアウトを返す.

    prs.slide_layouts は**先頭マスターの分しか返さない**。
    Google Slides由来の資料などはマスターが複数あり、実際に多用されている
    レイアウトが2番目以降のマスターにいることがある(実データで確認)。
    """
    return [lo for m in prs.slide_masters for lo in m.slide_layouts]


def design_layer_fills(prs, used_names):
    """マスターと「実際に使われているレイアウト」の塗りを集める.

    ブランド色はここにある。スライド側の塗りは注釈(赤い矢印など)や
    コンテンツの色が混ざるため、デザイン層を優先する。
    実データ3件で検証: スライド側だとHubSpot資料で赤い注釈を拾って外すが、
    デザイン層だとブランドのオレンジが正しく出る。
    """
    cnt, area = collections.Counter(), collections.Counter()
    for m in prs.slide_masters:
        solid_fills(m.shapes, cnt, area)
        for lo in m.slide_layouts:
            if lo.name in used_names:      # 未使用の表紙等は除外
                solid_fills(lo.shapes, cnt, area)
    return cnt, area


def cjk_ratio(prs):
    total = cjk = 0
    for sl in prs.slides:
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            for ch in sh.text_frame.text:
                if ch.strip():
                    total += 1
                    if "\u3040" <= ch <= "\u30ff" or "\u4e00" <= ch <= "\u9fff":
                        cjk += 1
    return cjk / total if total else 0


def pick_font(fonts, cjk):
    """日本語主体の資料でラテン専用フォントを選ばないようにする."""
    ranked = [f for f, _ in fonts.most_common()]
    if not ranked:
        return JP_FALLBACK
    if cjk >= 0.3:
        for f in ranked:
            if not any(l in f.lower() for l in LATIN_ONLY):
                return f
    return ranked[0]


def collect(prs):
    """図形の塗り・文字色・フォント・内容領域を実測で集計する."""
    fills = collections.Counter()
    fcolors = collections.Counter()
    fonts = collections.Counter()
    lefts, rights, tops = [], [], []
    W = prs.slide_width // 9525

    for sl in prs.slides:
        for sh in sl.shapes:
            try:
                fc = sh.fill
                if fc.type == MSO_FILL.SOLID and str(fc.fore_color.type).startswith("RGB"):
                    fills[str(fc.fore_color.rgb)] += 1
            except Exception:          # noqa: BLE001  塗りを持たない図形は多い
                pass
            if not sh.has_text_frame or sh.width is None:
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.name:
                        fonts[r.font.name] += len(r.text)
                    try:
                        if str(r.font.color.type).startswith("RGB"):
                            fcolors[str(r.font.color.rgb)] += len(r.text)
                    except Exception:  # noqa: BLE001
                        pass
            # 本文ブロックとみなせる幅の図形だけで内容領域を推定する
            w = sh.width // 9525
            top = sh.top // 9525 if sh.top is not None else -1
            if W * 0.45 < w < W * 0.99 and sh.text_frame.text.strip() and top >= 0:
                lefts.append(sh.left // 9525)
                rights.append((sh.left + sh.width) // 9525)
                tops.append(top)
    return fills, fcolors, fonts, lefts, rights, tops


def pick_palette(design_area, slide_fills):
    """デザイン層の色を先頭に置き、足りない分をスライド側から補う.

    先頭の色は表のヘッダ・フロー図の枠線・グラフ第1系列に使われるので、
    ブランドの主色が来るべき場所。
    """
    out = []
    for src in (design_area.most_common(), slide_fills.most_common()):
        for h, _ in src:
            if is_brandish(h) and far_enough(h, out):
                out.append(h)
            if len(out) >= 5:
                return out
    return out or ["4C6FFF"]


def pick_layout(prs):
    """新規スライドを載せるレイアウトを選ぶ.

    **実スライドが最も多く使っているレイアウトを採る。** これが本文用レイアウト。
    「プレースホルダが少ないもの」を選ぶ方式は誤り(実データで検証済み):
    セクション表紙が最少になり、右半分が巨大グラフィックのレイアウトを
    本文に使ってしまう。並び順での決め打ちも当然危険。
    スライドが無いテンプレート(.potx)の場合だけ名前のヒントに頼る。
    """
    used = collections.Counter(s.slide_layout.name for s in prs.slides)
    if used:
        return used.most_common(1)[0][0]
    scored = sorted((not any(h in (lo.name or "").lower() for h in LAYOUT_HINTS),
                     len(lo.placeholders), lo.name) for lo in all_layouts(prs))
    return scored[0][2] if scored else None


def percentile(xs, q):
    if not xs:
        return None
    xs = sorted(xs)
    return xs[min(len(xs) - 1, int(len(xs) * q))]


def build(src):
    prs = Presentation(str(src))
    fills, fcolors, fonts, lefts, rights, tops = collect(prs)
    W, H = prs.slide_width // 9525, prs.slide_height // 9525
    used = {s.slide_layout.name for s in prs.slides}
    d_cnt, d_area = design_layer_fills(prs, used)
    palette = pick_palette(d_area, fills)

    # 文字色は「最も使われている暗い色」を採用(黒に近いものが本文)
    dark = [c for c, _ in fcolors.most_common() if hsv_of(c)[2] < 0.5]
    text = dark[0] if dark else "333333"
    font = pick_font(fonts, cjk_ratio(prs))

    def clamp(v, lo, hi, default):
        return default if v is None else max(lo, min(hi, v))

    # 実測はあくまで手がかり。画面外の図形や全画面背景などの外れ値が混ざるので
    # 常識的な範囲にクランプする。最終調整は人間がやる前提
    left = clamp(percentile(lefts, 0.2), round(W * 0.03), round(W * 0.12),
                 round(W * 0.06))
    body_top = clamp(percentile([t for t in tops if t >= H * 0.08], 0.3),
                     round(H * 0.15), round(H * 0.38), round(H * 0.24))
    title_top = round(max(8, body_top - H * 0.145))

    stack, warn = font_stack_for(font)
    return {
        "_source": pathlib.Path(src).name,
        "_note": "実測抽出。人間のレビュー前提",
        "_font_warning": warn,
        "base_pptx": f"templates/{pathlib.Path(src).stem}.pptx",
        "base_layout": pick_layout(prs),
        "canvas": {"w": W, "h": H,
                   "slide_in": [round(prs.slide_width / 914400, 3),
                                round(prs.slide_height / 914400, 3)]},
        "font": font,
        "font_stack": stack,
        "palette": palette,
        "text": text,
        "heading": text,
        "margin": {"x": left, "top": title_top, "body_top": body_top},
        "table": {"header_bg": palette[0], "header_fg": "FFFFFF"},
        "flow": {"box_border": palette[0], "cap_color": palette[0]},
    }, (fills, fonts, prs, d_area.most_common(4))


def main():
    theme, (fills, fonts, prs, fills_dbg) = build(sys.argv[1])
    json.dump(theme, sys.stdout, ensure_ascii=False, indent=2)
    print(file=sys.stdout)
    e = sys.stderr
    print(f"# デザイン層 : {fills_dbg}", file=e)
    print(f"# スライド層 : {fills.most_common(4)}", file=e)
    print(f"# フォント   : {fonts.most_common(3)}", file=e)
    print(f"# レイアウト : {[lo.name for lo in prs.slide_layouts]}", file=e)
    print(f"# 選択       : {theme['base_layout']}", file=e)
    print(f"# フォント注意: {theme['_font_warning']}", file=e)


if __name__ == "__main__":
    main()
