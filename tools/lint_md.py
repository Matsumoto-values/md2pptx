#!/usr/bin/env python3
"""lint_md.py -- md2pptx に渡す md を検査する.

usage: python3 tools/lint_md.py deck.md [theme_id]
exit : 0=問題なし / 1=ERRORあり

目検に頼らず機械的に判定するためのツール。実際にブラウザで描画して座標を測るので、
「収まるかどうか」は推測ではなく実測で判定する。
ERROR は必ず直す。WARN は読みやすさの目安で、意図があれば無視してよい。
"""
import asyncio
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import md2html                      # noqa: E402
import theme as theme_mod           # noqa: E402
from extract import BrowserPool     # noqa: E402
from extract_stub import extract_stub  # noqa: E402

# 読みやすさの目安。物理的な上限（溢れる境界）とは別。
# 実測では箇条書き14行でも「収まる」が、資料としては読めない
SOFT = {"list_items": 6, "text_chars": 120, "title_chars": 40,
        "flow_boxes": 5, "table_rows": 8, "slides": 20, "kinds_per_slide": 3,
        "quote_chars": 90, "quotes_per_slide": 2}
CHART_TYPES = {"column", "column_stacked", "bar", "line", "pie"}


class Report:
    def __init__(self):
        self.errors, self.warns = [], []

    def error(self, slide, msg):
        self.errors.append((slide, msg))

    def warn(self, slide, msg):
        self.warns.append((slide, msg))


def check_source(md, rep):
    """描画前に md の書き方を検査する."""
    slides = md2html.parse_slides(md)
    if len(slides) > SOFT["slides"]:
        rep.warn(0, f"スライドが{len(slides)}枚（目安{SOFT['slides']}枚以下）")
    for i, lines in enumerate(slides, 1):
        text = "\n".join(lines)
        if not re.search(r"^# .+", text, re.M):
            rep.error(i, "見出し（# ）がない。全スライドに見出しを付ける")
        title = re.search(r"^# (.+)", text, re.M)
        if title and len(title.group(1)) > SOFT["title_chars"]:
            rep.warn(i, f"見出しが{len(title.group(1))}字（目安{SOFT['title_chars']}字以下）"
                        "。長いと見出し枠から溢れる")
        n_list = len(re.findall(r"^- ", text, re.M))
        if n_list > SOFT["list_items"]:
            rep.warn(i, f"箇条書きが{n_list}行（目安{SOFT['list_items']}行以下）。分割を検討")
        for para in re.findall(r"^(?![-#|!`\s])(.+)$", text, re.M):
            if len(para) > SOFT["text_chars"]:
                rep.warn(i, f"本文が{len(para)}字（目安{SOFT['text_chars']}字以下）")
        rows = [l for l in lines if l.strip().startswith("|")]
        if rows:
            widths = {len(r.strip().strip("|").split("|")) for r in rows}
            if len(widths) > 1:
                rep.error(i, f"表の列数が揃っていない（{sorted(widths)}）")
            if len(rows) - 2 > SOFT["table_rows"]:
                rep.warn(i, f"表のデータ行が{len(rows)-2}行（目安{SOFT['table_rows']}行以下）")
        for spec in re.findall(r"```chart\n(.*?)```", text, re.S):
            import yaml
            try:
                y = yaml.safe_load(spec) or {}
            except yaml.YAMLError as e:
                rep.error(i, f"chart のYAMLが壊れている: {e}")
                continue
            if y.get("type") not in CHART_TYPES:
                rep.error(i, f"chart の type が不正: {y.get('type')!r} "
                             f"（{sorted(CHART_TYPES)} のいずれか）")
            if not y.get("categories") or not y.get("series"):
                rep.error(i, "chart に categories と series が必要")
            for sr in y.get("series") or []:
                if len(sr.get("values") or []) != len(y.get("categories") or []):
                    rep.error(i, f"系列 {sr.get('name')!r} の値の個数が categories と違う")
        quotes = re.findall(r"^>\s*(.+)$", text, re.M)
        if len(quotes) > SOFT["quotes_per_slide"]:
            rep.warn(i, f"引用が{len(quotes)}件（目安{SOFT['quotes_per_slide']}件以下）。"
                        "代表的なものに絞る")
        for qt in quotes:
            if len(qt) > SOFT["quote_chars"]:
                rep.warn(i, f"引用が{len(qt)}字（目安{SOFT['quote_chars']}字以下）。"
                            "逐語のまま切り出す位置を見直す（要約はしない）")
        # 引用には話者の属性を添える（n やサンプル特性が読み取れるようにする）
        for m in re.finditer(r"^>\s*.+$", text, re.M):
            rest = text[m.end():].lstrip("\n")
            if not rest.startswith("--"):
                rep.warn(i, "引用に話者の属性がない。次の行に `-- 30代女性・解約者` "
                            "の形式で添える")
        if re.search(r"```chart", text) and not re.search(r"n\s*=\s*\d", text):
            rep.warn(i, "グラフのある枚に n の記載がない。母数を明記する")
        if re.search(r"^\|", text, re.M) and not re.search(r"n\s*=\s*\d", text):
            rep.warn(i, "表のある枚に n の記載がない。母数を明記する")

        for spec in re.findall(r"```flow\n(.*?)```", text, re.S):
            boxes = [l for l in spec.splitlines()
                     if l.strip() and not l.strip().startswith("direction:")]
            if len(boxes) > SOFT["flow_boxes"]:
                rep.warn(i, f"フローの箱が{len(boxes)}個（目安{SOFT['flow_boxes']}個以下）")
            if len(boxes) < 2:
                rep.error(i, "フローは箱が2個以上必要")


def check_layout(lay, rep, T):
    """実際に描画した座標で溢れを判定する（推測ではなく実測）."""
    cw, ch = lay["canvas"]
    if lay.get("font_ok") is False:
        rep.warn(0, f"フォント '{lay.get('font')}' が描画側で解決できていない。"
                    "座標が信用できないので、テーマのフォントを見直すこと")
    for i, s in enumerate(lay["slides"], 1):
        kinds = {e["kind"] for e in s["elements"]} - {"title"}
        if len(kinds) > SOFT["kinds_per_slide"]:
            rep.warn(i, f"1枚に{len(kinds)}種類の要素（{sorted(kinds)}）。分割を検討")
        for e in s["elements"]:
            if e["y"] + e["h"] > ch + 1:
                rep.error(i, f"{e['kind']} が下端を超える "
                             f"(y={e['y']:.0f}+h={e['h']:.0f} > {ch})")
            if e["x"] + e["w"] > cw + 1:
                rep.error(i, f"{e['kind']} が右端を超える "
                             f"(x={e['x']:.0f}+w={e['w']:.0f} > {cw})")
            if e["kind"] == "title" and e["h"] > T["title"]["h"] + 2:
                rep.error(i, f"見出しが枠を超えて折り返している "
                             f"(h={e['h']:.0f} > 枠{T['title']['h']})。短くする")


async def main():
    md = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    T = theme_mod.load(sys.argv[2] if len(sys.argv) > 2 else "default")
    rep = Report()
    check_source(md, rep)

    if rep.errors:
        # 描画すると例外で落ちる種類のエラーがあるので、先に報告して終わる
        print(f"theme={sys.argv[2] if len(sys.argv) > 2 else 'default'} "
              f"(記法エラーがあるため描画検査は省略)")
        for slide, msg in rep.errors:
            print(f"  ERROR slide{slide or '-'}: {msg}")
        for slide, msg in rep.warns:
            print(f"  WARN  slide{slide or '-'}: {msg}")
        print(f"\nERRORあり (error={len(rep.errors)} warn={len(rep.warns)})")
        return 1

    html = md2html.to_html(md, T)
    canvas = (T["canvas"]["w"], T["canvas"]["h"])
    pool = BrowserPool()
    try:
        await pool.start()
        lay = await pool.extract(html, canvas, T["font"])
        backend = "browser"
    except Exception:                     # noqa: BLE001 ブラウザが無い環境
        lay = extract_stub(html, canvas, T)
        backend = "stub(近似)"
    finally:
        try:
            await pool.stop()
        except Exception:                 # noqa: BLE001
            pass
    check_layout(lay, rep, T)

    print(f"theme={sys.argv[2] if len(sys.argv) > 2 else 'default'} "
          f"canvas={canvas[0]}x{canvas[1]} slides={len(lay['slides'])} backend={backend}")
    for slide, msg in rep.errors:
        print(f"  ERROR slide{slide or '-'}: {msg}")
    for slide, msg in rep.warns:
        print(f"  WARN  slide{slide or '-'}: {msg}")
    print(f"\n{'OK' if not rep.errors else 'ERRORあり'} "
          f"(error={len(rep.errors)} warn={len(rep.warns)})")
    return 1 if rep.errors else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
