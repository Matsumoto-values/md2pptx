#!/usr/bin/env python3
"""pptx_check.py -- 生成物の構造検証。CIのゲートとして使う.

python3 tests/pptx_check.py

httpxを必要としないので本番イメージ内でそのまま実行できる。
ブラウザバックエンドで動かすため、コンテナ内のフォント解決も同時に検証される。

ここで落とすべきもの:
  - フォントが解決できていない（座標が全部信用できない状態）
  - 要素がスライド外に溢れている
  - XMLの要素順序違反（PowerPointが「修復が必要」と言う状態）
  - 罫線・接続コネクタ・矢印先端の欠落
"""
import asyncio
import io
import pathlib
import sys
import zipfile

from lxml import etree
from pptx import Presentation

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import theme as theme_mod          # noqa: E402
from renderer import Renderer      # noqa: E402

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
FLOW_MD = """# フロー
```flow
before: 全案件を手作業でパワポ化している
after: mdからバッチ生成する
```
"""
fails = []


def check(cond, label, detail=""):
    print(f"  {'PASS' if cond else 'FAIL'}  {label}{'  ' + detail if detail else ''}")
    if not cond:
        fails.append(label)


def slide_xml(data, n=1):
    return etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read(f"ppt/slides/slide{n}.xml"))


def check_layout(lay, T):
    check(lay.get("font_ok") is not False, "フォント解決",
          f"font={lay.get('font')} ok={lay.get('font_ok')}")
    cw, ch = lay["canvas"]
    over = [f"{e['kind']}({e['x']:.0f},{e['y']:.0f},{e['w']:.0f}x{e['h']:.0f})"
            for s in lay["slides"] for e in s["elements"]
            if e["x"] + e["w"] > cw + 1 or e["y"] + e["h"] > ch + 1]
    check(not over, "要素がスライド内に収まっている", ", ".join(over[:3]))


def check_deck(data, T):
    prs = Presentation(io.BytesIO(data))
    check(len(prs.slides) == 3, "スライド数", f"{len(prs.slides)}枚")
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    check(run.font.name == T["font"], "pptxのフォント名", str(run.font.name))

    tables = [sh for s in prs.slides for sh in s.shapes if sh.has_table]
    check(len(tables) == 1, "表の数", str(len(tables)))
    if tables:
        tcPr = tables[0].table.cell(0, 0)._tc.find("{*}tcPr")
        order = [q.tag.split("}")[1] for q in tcPr]
        lns = [t for t in order if t.startswith("ln")]
        # tcPrの子要素順序は ln* -> fill。破るとPowerPointが修復を要求する
        check(lns == ["lnL", "lnR", "lnT", "lnB"], "罫線の要素順序", str(order))
        check(order.index("lnB") < order.index("solidFill"), "罫線がfillより前", "")

    charts = [sh for s in prs.slides for sh in s.shapes if sh.has_chart]
    check(len(charts) == 1, "グラフの数", str(len(charts)))
    if charts:
        ch = charts[0].chart
        check(len(ch.series) == 2, "系列数", str(len(ch.series)))
        check(ch.plots[0].vary_by_categories is False, "vary_by_categories=False", "")


def check_flow(data):
    root = slide_xml(data)
    geoms = [g.get("prst") for g in root.findall(".//a:prstGeom", NS)]
    check("roundRect" in geoms, "角丸図形が存在", str(geoms))
    cxn = root.findall(".//p:cxnSp", NS)
    check(len(cxn) == 1, "コネクタの数", str(len(cxn)))
    if cxn:
        st, en = cxn[0].find(".//a:stCxn", NS), cxn[0].find(".//a:endCxn", NS)
        check(st is not None and en is not None, "コネクタが図形に接続されている",
              f"{st.get('idx') if st is not None else '-'}"
              f"->{en.get('idx') if en is not None else '-'}")
        ln = cxn[0].find(".//a:ln", NS)
        tail = ln.find("a:tailEnd", NS) if ln is not None else None
        check(tail is not None, "矢印の先端がある",
              tail.get("type") if tail is not None else "")
        order = [q.tag.split("}")[1] for q in ln] if ln is not None else []
        check(order and order[-1] == "tailEnd", "a:lnの要素順序", str(order))


async def main():
    T = theme_mod.load()
    backend = "browser"
    print(f"backend={backend} theme.font={T['font']}")
    async with Renderer(backend) as r:
        print("[deck]")
        _, lay = await r.layout((ROOT / "examples/deck.md").read_text(encoding="utf-8"), T)
        check_layout(lay, T)
        data, _ = await r.render((ROOT / "examples/deck.md").read_text(encoding="utf-8"), T)
        check_deck(data, T)

        print("[flow]")
        _, lay = await r.layout(FLOW_MD, T)
        check_layout(lay, T)
        data, _ = await r.render(FLOW_MD, T)
        check_flow(data)

    print("\n" + ("ALL PASSED" if not fails else f"FAILED: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
