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

import lxml.etree as etree
from pptx import Presentation
from pptx.enum.dml import MSO_FILL

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
REL_THEME = ("http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/theme")

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


def luminance(h):
    """相対輝度。白文字が乗せられるかの判定に使う."""
    r, g, b = (v / 255 for v in rgb_of(h))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


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


def theme_fonts(prs):
    """テーマのフォント定義を読む -> {"+mn-ea": "ＭＳ Ｐゴシック", ...}

    **色はテーマ部品から採らない**（Office既定のまま放置されているため）が、
    フォント参照の解決にはテーマ部品が必要。実資料で run のフォント名が
    `+mn-ea` のようなテーマ参照になっているケースがあり、そのままでは
    フォント名として使えない。
    """
    out = {}
    try:
        blob = prs.slide_masters[0].part.part_related_by(REL_THEME).blob
        x = etree.fromstring(blob)
    except Exception:          # noqa: BLE001 テーマが辿れない資料もある
        return out
    for kind, prefix in (("majorFont", "+mj"), ("minorFont", "+mn")):
        node = x.find(f".//a:{kind}", NS)
        if node is None:
            continue
        latin = node.find("a:latin", NS)
        latin = latin.get("typeface") if latin is not None else None
        ea = node.find("a:ea", NS)
        ea = ea.get("typeface") if ea is not None else None
        jpan = [f.get("typeface") for f in node.findall("a:font", NS)
                if f.get("script") == "Jpan"]
        # ea が空文字のことが多いので Jpan の指定を優先して拾う
        east = (jpan[0] if jpan else None) or ea or latin
        out[f"{prefix}-ea"] = east
        out[f"{prefix}-lt"] = latin or east
        out[f"{prefix}-cs"] = latin or east
    return {k: v for k, v in out.items() if v}


def title_boxes(prs):
    """実スライドからタイトルの位置・高さ・文字サイズを実測する.

    タイトル位置を「本文位置からの計算値」で決めるのは誤り。
    テンプレートには上部にヘッダ帯があり、タイトルはその中に置く設計のことが多い。
    計算値だと帯の下に落ちて、帯が空のまま残る（実資料で確認）。

    タイトル＝上部にあるテキスト図形のうち最も文字が大きいもの、として拾う。
    プレースホルダが定義されていないテンプレートが多いため、実物から測るしかない。
    """
    H = prs.slide_height // 9525
    rows = []
    for sl in prs.slides:
        best = None
        for sh in sl.shapes:
            if (not sh.has_text_frame or sh.top is None or sh.width is None
                    or not sh.text_frame.text.strip()):
                continue
            top = sh.top // 9525
            if top > H * 0.4:                 # 上部に無いものはタイトルではない
                continue
            size = 0
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size:
                        size = max(size, r.font.size.pt)
            if not size:
                continue
            cand = (size, top, sh.left // 9525, sh.height // 9525)
            if best is None or cand[0] > best[0]:
                best = cand
        if best:
            rows.append(best)
    return rows


def body_size(prs, H):
    """本文の文字サイズを文字数で重み付けして実測する.

    タイトルだけ実測して本文を固定値にすると、見出しと本文が同じ大きさになって
    階層が崩れる（実資料でタイトル18pt・本文9.5ptなのに、生成物は
    タイトル24px・本文22pxになった）。
    単純な中央値ではなく文字数で重み付けするのは、注釈や凡例のような
    短い小さな文字に引っ張られないようにするため。
    """
    rows = []
    for sl in prs.slides:
        for sh in sl.shapes:
            if not sh.has_text_frame or sh.top is None:
                continue
            if sh.top // 9525 < H * 0.2:      # タイトル帯は除外
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size and len(r.text.strip()) > 4:
                        rows.append((r.font.size.pt, len(r.text)))
    if not rows:
        return None
    rows.sort()
    half = sum(n for _, n in rows) / 2
    acc = 0
    for pt, n in rows:
        acc += n
        if acc >= half:
            return pt
    return rows[-1][0]


def footer_top(prs, H):
    """フッタ・ロゴ・ページ番号の上端を実測する.

    下端に確保すべき余白は固定値ではなくテンプレートごとに違う。
    マスターとレイアウトの下部にある要素の上端がその境界になる。
    """
    tops = []
    for m in prs.slide_masters:
        for holder in [m] + list(m.slide_layouts):
            for sh in holder.shapes:
                if sh.top is None:
                    continue
                t = sh.top // 9525
                if t > H * 0.75:
                    tops.append(t)
    return min(tops) if tops else None


def collect(prs, refs=None):
    """図形の塗り・文字色・フォント・内容領域を実測で集計する."""
    refs = refs or {}
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
                    name = r.font.name
                    if name and name.startswith("+"):
                        name = refs.get(name)          # テーマ参照を実名へ
                    if name:
                        fonts[name] += len(r.text)
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


def rank_primary(colors):
    """主色の候補を並べ替える.

    蛍光ペン的な明るい色(FFFF00, CCFFFF など)が先頭に来ると、表ヘッダに
    白文字を乗せた時に読めなくなる。彩度が高く輝度が中程度の色を優先する。
    実資料で FFFF00 や FFFFCC が主色になるケースを確認したため。
    """
    def score(h):
        _, sat, _ = hsv_of(h)
        lum = luminance(h)
        # 輝度0.25付近を理想として離れるほど減点、彩度は加点
        return sat * 1.5 - abs(lum - 0.25) * 2
    return sorted(colors, key=score, reverse=True)


def pick_palette(design_area, slide_fills):
    """デザイン層の色を先頭に置き、足りない分をスライド側から補う.

    先頭の色は表のヘッダ・フロー図の枠線・グラフ第1系列に使われるので、
    ブランドの主色が来るべき場所。
    """
    design, extra = [], []
    for src, bucket in ((design_area.most_common(), design),
                        (slide_fills.most_common(), extra)):
        for h, _ in src:
            if is_brandish(h) and far_enough(h, design + extra):
                bucket.append(h)
    # デザイン層の色を優先し、その中で主色になるべきものを先頭へ
    out = rank_primary(design) + extra
    return out[:5] or ["4C6FFF"]


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
        # 使用頻度が同数のことがある（実データで4対4を確認）。
        # その場合は図形が多い＝装飾を持っているレイアウトを選ぶ
        shapes = {lo.name: len(lo.shapes) for lo in all_layouts(prs)}
        return max(used.items(), key=lambda kv: (kv[1], shapes.get(kv[0], 0)))[0]
    scored = sorted((not any(h in (lo.name or "").lower() for h in LAYOUT_HINTS),
                     len(lo.placeholders), lo.name) for lo in all_layouts(prs))
    return scored[0][2] if scored else None


def percentile(xs, q):
    if not xs:
        return None
    xs = sorted(xs)
    return xs[min(len(xs) - 1, int(len(xs) * q))]


def build(src, source_name=None):
    """src はパスでもファイルライクでも可（アップロードされたバイト列を直接渡せる）."""
    prs = Presentation(src if hasattr(src, "read") else str(src))
    stem = (pathlib.Path(source_name or str(src)).stem
            if not hasattr(src, "read") or source_name else "template")
    label = pathlib.Path(source_name).name if source_name else (
        pathlib.Path(str(src)).name if not hasattr(src, "read") else "uploaded.pptx")
    refs = theme_fonts(prs)
    fills, fcolors, fonts, lefts, rights, tops = collect(prs, refs)
    W, H = prs.slide_width // 9525, prs.slide_height // 9525
    used = {s.slide_layout.name for s in prs.slides}
    d_cnt, d_area = design_layer_fills(prs, used)
    palette = pick_palette(d_area, fills)

    # 文字色は「最も使われている暗い色」を採用(黒に近いものが本文)
    dark = [c for c, _ in fcolors.most_common() if hsv_of(c)[2] < 0.5]
    text = dark[0] if dark else "333333"
    font = pick_font(fonts, cjk_ratio(prs))
    if font.startswith("+"):                  # 解決できなかった場合の保険
        font = refs.get(font) or JP_FALLBACK

    def clamp(v, lo, hi, default):
        return default if v is None else max(lo, min(hi, v))

    # 実測はあくまで手がかり。画面外の図形や全画面背景などの外れ値が混ざるので
    # 常識的な範囲にクランプする。最終調整は人間がやる前提
    titles = title_boxes(prs)
    left = clamp(percentile(lefts, 0.2), round(W * 0.03), round(W * 0.12),
                 round(W * 0.06))
    if titles:
        # タイトルの左端は本文と別に持つテンプレートがある
        left = clamp(percentile([t[2] for t in titles], 0.5),
                     round(W * 0.02), round(W * 0.12), left)
    body_top = clamp(percentile([t for t in tops if t >= H * 0.08], 0.3),
                     round(H * 0.15), round(H * 0.38), round(H * 0.24))
    if titles:
        # 実測値を使う。ヘッダ帯の中に置く設計だと y=0 付近になる
        title_top = clamp(percentile([t[1] for t in titles], 0.5),
                          0, round(H * 0.35), round(H * 0.09))
        title_h = clamp(percentile([t[3] for t in titles], 0.5),
                        round(H * 0.06), round(H * 0.2), round(H * 0.1))
        title_size = clamp(percentile([t[0] for t in titles], 0.5), 12, 40, 24)
    else:
        title_top = round(max(8, body_top - H * 0.145))
        title_h = round(H * 0.1)
        title_size = 24

    stack, warn = font_stack_for(font)
    # 実測値どうしの整合を取る。タイトルと本文の判定はそれぞれ独立に行うため、
    # タイトルが本文より下に来る資料がある（実測で2件）。
    # その場合はタイトルを上に、本文をタイトルの下へ寄せる
    if title_top + title_h > body_top:
        if title_top >= body_top:                 # 完全に逆転している
            title_top = clamp(round(body_top - title_h - H * 0.02), 0,
                              round(H * 0.35), 0)
        body_top = clamp(round(title_top + title_h + H * 0.03),
                         round(H * 0.12), round(H * 0.45), body_top)

    # 本文サイズ。実測できない場合はタイトルとの比率で決める
    bsize = body_size(prs, H)
    # 元資料は文字を詰め込んでいることが多く、そのまま採ると小さすぎることがある。
    # 忠実さと可読性の折衷として下限を設ける（14px = 10.5pt）。
    # 元資料どおりにしたい場合はテーマの body を手で書き換える
    body_px = clamp(round(bsize / 0.75) if bsize else None, 14, 30,
                    max(14, round(title_size / 0.75 * 0.5)))
    # 見出しと本文が同じ大きさにならないよう最低限の差をつける
    title_px = max(round(title_size / 0.75), round(body_px * 1.4))

    return {
        "_source": label,
        "_note": "実測抽出。人間のレビュー前提",
        "_font_warning": warn,
        "base_pptx": f"templates/{stem}.pptx",
        "base_layout": pick_layout(prs),
        "canvas": {"w": W, "h": H,
                   "slide_in": [round(prs.slide_width / 914400, 3),
                                round(prs.slide_height / 914400, 3)]},
        "font": font,
        "font_stack": stack,
        "palette": palette,
        "text": text,
        "heading": text,
        "margin": {"x": left, "top": title_top, "body_top": body_top,
                   # フッタの実測位置から下端余白を決める（固定値だと下に偏る）
                   "bottom": clamp(H - (footer_top(prs, H) or H), round(H * 0.03),
                                   round(H * 0.18), round(H * 0.06)),
                   "distribute": "space-evenly"},
        # タイトルは実測した枠に縦中央で収める（ヘッダ帯に合わせるため）
        "title": {"h": title_h, "size": title_px},
        # 本文の文字サイズも実測する（固定値だと見出しとの階層が崩れる）
        "body": {"text": body_px, "list": body_px,
                 "table": max(11, round(body_px * 0.92))},
        # ヘッダ文字色は背景の輝度で決める。明るい背景に白文字だと読めない
        "table": {"header_bg": palette[0],
                  "header_fg": "FFFFFF" if luminance(palette[0]) < 0.45 else text},
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
