#!/usr/bin/env python3
"""extract.py -- HTMLをChromiumで描画し、レイアウト中間表現(JSON)を吐く.

CLI : extract.py deck.html [theme_id] > layout.json
API : await extract(browser, html, canvas) -> dict

ブラウザに折り返し計算をやらせるのがこのツールの存在意義。
data-pptx="raster" の要素だけは要素スクショをbase64で同梱する。
"""
import asyncio
import base64
import json
import pathlib
import sys

# Cloud Run(gVisor)では sandbox と /dev/shm が使えないため必須の引数
LAUNCH_ARGS = ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
               "--font-render-hinting=none"]

COLLECT_JS = """
(slide) => {
  const px = v => parseFloat(v) || 0;
  const base = slide.getBoundingClientRect();
  const rectOf = el => {
    const r = el.getBoundingClientRect();
    return { x: r.left - base.left, y: r.top - base.top, w: r.width, h: r.height };
  };
  const bordersOf = el => {
    const cs = getComputedStyle(el);
    const S = { L: 'Left', R: 'Right', T: 'Top', B: 'Bottom' };
    const out = {};
    for (const k in S) {
      out[k] = { w: px(cs['border' + S[k] + 'Width']),
                 color: cs['border' + S[k] + 'Color'],
                 style: cs['border' + S[k] + 'Style'] };
    }
    return out;
  };
  const styleOf = el => {
    const cs = getComputedStyle(el);
    return {
      font: cs.fontFamily.split(',')[0].replace(/["']/g, ''),
      size: px(cs.fontSize),
      bold: px(cs.fontWeight) >= 600,
      color: cs.color,
      bg: cs.backgroundColor,
      align: cs.textAlign,
      // flexで縦中央に寄せている要素はPPTX側でも縦中央にする。
      // これを拾わないと、枠だけ実測して中身が上寄せになる
      valign: (cs.display.indexOf('flex') >= 0 && cs.alignItems === 'center')
              ? 'middle' : 'top',
      lineHeight: px(cs.lineHeight) / (px(cs.fontSize) || 1),
      pad: [cs.paddingTop, cs.paddingRight, cs.paddingBottom, cs.paddingLeft].map(px)
    };
  };
  return [...slide.querySelectorAll('[data-pptx]')].map(el => {
    const kind = el.dataset.pptx;
    const e = Object.assign({ kind }, rectOf(el), styleOf(el));
    if (kind === 'table') {
      // セル単位で実測する。列幅・行高はブラウザの計算結果が唯一の正解
      e.header = el.dataset.pptxHeader === '1';
      e.rows = [...el.rows].map(tr => ({
        h: tr.getBoundingClientRect().height,
        cells: [...tr.cells].map(td => Object.assign(
          { text: td.innerText, colspan: td.colSpan, rowspan: td.rowSpan,
            borders: bordersOf(td) },
          rectOf(td), styleOf(td)))
      }));
      e.colWidths = [...el.rows[0].cells].map(td => td.getBoundingClientRect().width);
    } else if (kind === 'shape') {
      e.shape = el.dataset.shape || 'rounded_rect';
      e.node = el.dataset.node;
      e.radius = px(getComputedStyle(el).borderTopLeftRadius);
      e.borderW = px(getComputedStyle(el).borderTopWidth);
      e.borderColor = getComputedStyle(el).borderTopColor;
      // 見出しと本文でフォントが違うので行ごとにスタイルを採る
      const spans = [...el.querySelectorAll('.cap,.body')];
      e.lines = (spans.length ? spans : [el]).map(n => {
        const cs = getComputedStyle(n);
        return { text: n.innerText, size: px(cs.fontSize),
                 bold: px(cs.fontWeight) >= 600, color: cs.color,
                 lineHeight: px(cs.lineHeight) / (px(cs.fontSize) || 1) };
      });
    } else if (kind === 'quote') {
      // 発言と属性でフォントサイズ・色が違うので行ごとにスタイルを採る
      const cs = getComputedStyle(el);
      e.borderW = px(cs.borderLeftWidth);
      e.borderColor = cs.borderLeftColor;
      e.lines = [...el.querySelectorAll('.said,.attr')].map(n => {
        const s2 = getComputedStyle(n);
        return { text: n.innerText, size: px(s2.fontSize),
                 bold: px(s2.fontWeight) >= 600, color: s2.color,
                 lineHeight: px(s2.lineHeight) / (px(s2.fontSize) || 1) };
      });
    } else if (kind === 'connector') {
      // 座標より「どの図形とどの図形を繋ぐか」が本質。PPTX側で接続して作り直す
      e.from = el.dataset.from;
      e.to = el.dataset.to;
      e.dir = el.dataset.dir || 'horizontal';
    } else if (kind === 'chart') {
      // 数値はDOMから逆算できないので属性のJSONを唯一の正とする。外枠だけ実測
      e.spec = JSON.parse(el.dataset.chart);
    } else if (kind === 'list') {
      e.text = [...el.querySelectorAll('li')].map(li => li.innerText).join('\\n');
    } else {
      e.text = el.innerText;
    }
    return e;
  });
}
"""


# 指定フォントが実際に使われたかを幅比較で判定する。
# document.fonts.check はシステムフォントに対して信用できないため canvas で実測する
FONT_PROBE_JS = """
(family) => {
  const c = document.createElement('canvas').getContext('2d');
  const probe = '日本語Wwmiil 0123';
  c.font = '40px monospace';
  const base = c.measureText(probe).width;
  c.font = '40px "' + family + '", monospace';
  const test = c.measureText(probe).width;
  return Math.abs(test - base) > 0.5;
}
"""


async def extract(browser, html, canvas=(1280, 720), font=None):
    w, h = canvas
    page = await browser.new_page(viewport={"width": w, "height": h},
                                 device_scale_factor=2)
    try:
        await page.set_content(html, wait_until="load")
        await page.wait_for_timeout(150)          # webfont適用待ち
        font_ok = None
        if font:
            font_ok = await page.evaluate(FONT_PROBE_JS, font)
            if not font_ok:
                # 意図したフォントで測れていない = 座標が全部信用できない状態
                print(f"WARNING: font '{font}' is not available; "
                      f"measurements fell back to another font", file=sys.stderr)
        slides = []
        for handle in await page.query_selector_all("section.slide"):
            elems = await handle.evaluate(COLLECT_JS)
            cells = await handle.query_selector_all("[data-pptx]")
            for e, cell in zip(elems, cells):
                if e["kind"] == "raster":
                    e["png_b64"] = base64.b64encode(await cell.screenshot()).decode()
            slides.append({"w": w, "h": h, "elements": elems})
        return {"canvas": [w, h], "font": font, "font_ok": font_ok,
                "slides": slides}
    finally:
        await page.close()


class BrowserPool:
    """Chromiumはプロセス起動に3〜5秒かかるので常駐させて全リクエストで共有する.

    ページは同時実行数を絞る(メモリを食うため)。Cloud Runでは
    concurrency と max_pages を揃えておくと安定する。
    """

    def __init__(self, max_pages=4):
        self._sem = asyncio.Semaphore(max_pages)
        self._pw = self._browser = None

    async def start(self):
        from playwright.async_api import async_playwright
        self._pw = await async_playwright().start()
        self._browser = await self._pw.chromium.launch(args=LAUNCH_ARGS)

    async def stop(self):
        if self._browser:
            await self._browser.close()
        if self._pw:
            await self._pw.stop()

    @property
    def alive(self):
        return bool(self._browser and self._browser.is_connected())

    async def extract(self, html, canvas, font=None):
        async with self._sem:
            return await extract(self._browser, html, canvas, font)


async def main():
    html = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    import theme as theme_mod
    # テーマを渡さないと viewport と font が既定値になり、md2html側と食い違って
    # px->EMU 換算が二重にスケールする
    T = theme_mod.load(sys.argv[2] if len(sys.argv) > 2 else "default")
    pool = BrowserPool()
    await pool.start()
    try:
        layout = await pool.extract(html, (T["canvas"]["w"], T["canvas"]["h"]),
                                    T["font"])
    finally:
        await pool.stop()
    json.dump(layout, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    asyncio.run(main())
