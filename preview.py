#!/usr/bin/env python3
"""preview.py -- 生成したpptxを画像化して人間に見せる.

テンプレート登録の確認用。**HTMLプレビューではテンプレートの装飾が見えない**
（ロゴやバナーはベースpptxのマスター側にあり、HTMLには存在しない）。
「ロゴと本文が一緒にどう見えるか」「本文が装飾に被っていないか」を確認するには
実物のpptxをラスタライズする必要がある。

soffice が無い環境では PreviewUnavailable を投げる。呼び出し側はHTMLへ退避する。
"""
import glob
import hashlib
import json
import os
import pathlib
import shutil
import subprocess
import tempfile
import uuid

SOFFICE = os.getenv("SOFFICE_BIN", "soffice")
PDFTOPPM = os.getenv("PDFTOPPM_BIN", "pdftoppm")
CACHE_DIR = pathlib.Path(tempfile.gettempdir()) / "md2pptx-preview"

# プレビュー用のサンプル。全要素を1枚に詰めず、テンプレートの装飾と
# 本文の衝突が見えるように配置する
SAMPLE_MD = """# テンプレート適用の確認
この資料はテンプレート登録時の確認用に自動生成されています。
- ロゴ・バナー・背景が継承されているか
- 本文が装飾に被っていないか
- 配色とフォントが意図通りか
!! 主色のサンプル
---
# 表とグラフの見え方
| 指標 | 前期 | 今期 |
| --- | --- | --- |
| 売上 | 1200 | 1480 |
| 利益率 | 12 | 15 |
```chart
type: column
categories: [Q1, Q2, Q3, Q4]
series:
  - {name: 実績, values: [1200, 1480, 1690, 2050]}
  - {name: 目標, values: [1300, 1500, 1800, 2100]}
```
---
# 図形の見え方
```flow
before: 現状の課題
after: 導入後の姿
```
"""


class PreviewUnavailable(RuntimeError):
    """ラスタライズ手段が無い（LibreOfficeが未導入）."""


def available():
    return bool(shutil.which(SOFFICE) and shutil.which(PDFTOPPM))


def rasterize(pptx_bytes, pages=3, dpi=110, timeout=180):
    """pptx -> PNGのリスト（先頭pages枚）.

    soffice はプロファイルディレクトリを排他ロックするため、呼び出しごとに
    使い捨てのプロファイルを渡す。これをしないと同時実行で固まる。
    """
    if not available():
        raise PreviewUnavailable(f"{SOFFICE} / {PDFTOPPM} が見つからない")
    work = pathlib.Path(tempfile.mkdtemp(prefix="preview-"))
    try:
        src = work / "deck.pptx"
        src.write_bytes(pptx_bytes)
        profile = work / f"profile-{uuid.uuid4().hex[:8]}"
        subprocess.run(
            [SOFFICE, "--headless", "--norestore", "--invisible",
             f"-env:UserInstallation=file://{profile}",
             "--convert-to", "pdf", "--outdir", str(work), str(src)],
            capture_output=True, timeout=timeout, check=False,
            env={**os.environ, "HOME": str(work)})
        pdf = work / "deck.pdf"
        if not pdf.is_file():
            raise PreviewUnavailable("pptx -> pdf の変換に失敗した")
        subprocess.run(
            [PDFTOPPM, "-png", "-r", str(dpi), "-f", "1", "-l", str(pages),
             str(pdf), str(work / "page")],
            capture_output=True, timeout=timeout, check=True)
        return [pathlib.Path(p).read_bytes()
                for p in sorted(glob.glob(str(work / "page-*.png")))]
    finally:
        shutil.rmtree(work, ignore_errors=True)


def cache_key(base_sha, theme):
    """ベースの内容とテーマの内容でキャッシュする.

    テーマを編集して再確認する使い方をするので、テーマもキーに含める。
    """
    h = hashlib.sha256()
    h.update((base_sha or "").encode())
    h.update(json.dumps(theme, sort_keys=True, ensure_ascii=False,
                        default=str).encode())
    return h.hexdigest()[:20]


def cached(key):
    d = CACHE_DIR / key
    if not d.is_dir():
        return None
    imgs = sorted(d.glob("page-*.png"))
    return [p.read_bytes() for p in imgs] or None


def store(key, images):
    d = CACHE_DIR / key
    d.mkdir(parents=True, exist_ok=True)
    for i, b in enumerate(images, 1):
        (d / f"page-{i}.png").write_bytes(b)
    return images
