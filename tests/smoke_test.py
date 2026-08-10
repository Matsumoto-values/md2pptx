#!/usr/bin/env python3
"""smoke_test.py -- ブラウザ無しでAPI経路とバッチ経路を通すスモークテスト.

python3 tests/smoke_test.py
"""
import os
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ.setdefault("EXTRACT_BACKEND", "stub")

from fastapi.testclient import TestClient   # noqa: E402

import app as service                        # noqa: E402

MD = (ROOT / "examples/deck.md").read_text(encoding="utf-8")


def main():
    fails = []
    with TestClient(service.app) as c:
        r = c.get("/healthz")
        print("healthz  ", r.status_code, r.json())
        fails += [] if r.status_code == 200 else ["healthz"]

        r = c.get("/themes")
        print("themes   ", r.status_code, r.json())

        r = c.post("/preview", json={"markdown": MD})
        ok = r.status_code == 200 and 'data-pptx="table"' in r.text
        print("preview  ", r.status_code, f"{len(r.text)} bytes", "OK" if ok else "NG")
        fails += [] if ok else ["preview"]

        r = c.post("/layout", json={"markdown": MD})
        n = len(r.json()["slides"])
        print("layout   ", r.status_code, f"{n} slides")
        fails += [] if n == 3 else ["layout"]

        # テーマの部分上書きが効くか（ブランド色だけ差し替える想定）
        r = c.post("/render", json={
            "markdown": MD, "filename": "themed.pptx",
            "theme": {"palette": ["FF6B00", "004E98"],
                      "table": {"zebra": False, "header_bg": "004E98"}}})
        ct = r.headers.get("content-type", "")
        ok = r.status_code == 200 and "presentation" in ct and len(r.content) > 20000
        print("render   ", r.status_code, f"{len(r.content)} bytes",
              f"slides={r.headers.get('X-Slide-Count')}", "OK" if ok else "NG")
        fails += [] if ok else ["render"]

        r = c.post("/render", json={"markdown": MD, "theme_id": "nope"})
        print("bad theme", r.status_code, "(400期待)")
        fails += [] if r.status_code == 400 else ["bad-theme"]

        r = c.post("/render", json={"markdown": ""})
        print("empty md ", r.status_code, "(422期待)")
        fails += [] if r.status_code == 422 else ["empty"]

    fails += batch_check()
    print("\n" + ("ALL PASSED" if not fails else f"FAILED: {fails}"))
    return 1 if fails else 0


def run_job(manifest, out):
    import subprocess
    r = subprocess.run(
        [sys.executable, str(ROOT / "job.py")], cwd=ROOT, capture_output=True,
        env={**os.environ, "EXTRACT_BACKEND": "stub",
             "MANIFEST": str(ROOT / manifest), "OUT_PREFIX": out})
    return r.returncode


def batch_check():
    """終了コードの分類を検証する.

    データ起因(リトライ無意味)は正常終了し、環境起因だけ非ゼロで終わること。
    ここを間違えると、絶対に成功しない処理をCloud Run Jobsが繰り返す。
    """
    import shutil
    import tempfile
    fails, out = [], tempfile.mkdtemp()
    try:
        # 1. 全件成功 -> exit 0
        code = run_job("examples/manifest.jsonl", out + "/ok")
        made = len(list(pathlib.Path(out + "/ok").glob("*.pptx")))
        ok = code == 0 and made == 4
        print("batch ok  ", f"exit={code}", f"{made} files", "OK" if ok else "NG")
        fails += [] if ok else ["batch-ok"]

        # 2. データ起因の失敗のみ -> exit 0 + 記録が残る
        code = run_job("examples/manifest-failures.jsonl", out + "/data")
        made = len(list(pathlib.Path(out + "/data").glob("*.pptx")))
        rec = pathlib.Path(out + "/data/_failures/0.jsonl")
        ok = code == 0 and made == 4 and rec.is_file() and "data" in rec.read_text()
        print("batch data", f"exit={code}", f"{made} files",
              f"failures={rec.is_file()}", "OK" if ok else "NG")
        fails += [] if ok else ["batch-data"]

        # 3. 環境起因の失敗 -> exit 1 (Cloud Run Jobsにリトライさせる)
        code = run_job("examples/manifest.jsonl", "/proc/nowrite")
        ok = code == 1
        print("batch env ", f"exit={code}", "(1期待)", "OK" if ok else "NG")
        fails += [] if ok else ["batch-env"]
        return fails
    finally:
        shutil.rmtree(out, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
