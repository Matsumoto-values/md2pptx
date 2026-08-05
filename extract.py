#!/usr/bin/env python3
"""extract.py -- HTMLをChromiumで描画し、レイアウト中間表現(JSON)を吐く.

CLI : extract.py deck.html > layout.json
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


async def extract(browser, html, canvas=(1280, 720)):
    w, h = canvas
    page = await browser.new_page(viewport={"width": w, "height": h},
                                 device_scale_factor=2)
    try:
        await page.set_content(html, wait_until="load")
        await page.wait_for_timeout(150)          # webfont適用待ち
        slides = []
        for handle in await page.query_selector_all("section.slide"):
            elems = await handle.evaluate(COLLECT_JS)
            cells = await handle.query_selector_all("[data-pptx]")
            for e, cell in zip(elems, cells):
                if e["kind"] == "raster":
                    e["png_b64"] = base64.b64encode(await cell.screenshot()).decode()
            slides.append({"w": w, "h": h, "elements": elems})
        return {"canvas": [w, h], "slides": slides}
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

    async def extract(self, html, canvas):
        async with self._sem:
            return await extract(self._browser, html, canvas)


async def main():
    html = pathlib.Path(sys.argv[1]).read_text(encoding="utf-8")
    pool = BrowserPool()
    await pool.start()
    try:
        layout = await pool.extract(html, (1280, 720))
    finally:
        await pool.stop()
    json.dump(layout, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    asyncio.run(main())
