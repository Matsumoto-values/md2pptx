#!/usr/bin/env python3
"""template_flow_test.py -- 登録 -> 発行 -> 生成 の一連の流れを検証する.

python3 tests/template_flow_test.py [template.pptx]

テンプレートファイルはリポジトリに含めない（顧客資料のため）ので、
引数で渡されなければ python-pptx の既定テンプレートを合成して使う。
"""
import io
import os
import pathlib
import sys
import tempfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ.setdefault("EXTRACT_BACKEND", "stub")
os.environ["TEMPLATE_ROOT"] = tempfile.mkdtemp(prefix="tstore-")

from fastapi.testclient import TestClient   # noqa: E402

import app as service                       # noqa: E402

MD = (ROOT / "examples/deck.md").read_text(encoding="utf-8")
fails = []


def check(cond, label, detail=""):
    print(f"  {'PASS' if cond else 'FAIL'}  {label}{'  ' + detail if detail else ''}")
    if not cond:
        fails.append(label)


def sample_template():
    """検証用テンプレートを合成する.

    python-pptx はマスター/レイアウトへの図形追加をサポートしないため、
    装飾はスライド側に置く（デザイン層からの抽出は実資料3件で別途検証済み。
    ここでは登録〜生成の流れを確認する）。
    """
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Emu, Pt
    prs = Presentation()
    for _ in range(3):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Emu(1280 * 9525),
                                 Emu(12 * 9525))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor.from_string("00B0F0")   # ブランド色
        bar.line.fill.background()
        tb = s.shapes.add_textbox(Emu(80 * 9525), Emu(170 * 9525), Emu(1000 * 9525),
                                  Emu(60 * 9525))
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = "サンプル本文テキストです"
        r.font.size, r.font.name = Pt(18), "BIZ UDPGothic"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def main():
    data = (pathlib.Path(sys.argv[1]).read_bytes() if len(sys.argv) > 1
            else sample_template())
    src = sys.argv[1] if len(sys.argv) > 1 else "(合成テンプレート)"
    print(f"template = {src}  {len(data)} bytes")

    with TestClient(service.app) as c:
        # 1. 登録
        r = c.post("/tenants/acme/templates",
                   data={"template_id": "corp", "name": "コーポレート"},
                   files={"file": ("corp.pptx", data, "application/vnd.openxml"
                                   "formats-officedocument.presentationml.presentation")})
        ok = r.status_code == 200
        print("upload   ", r.status_code, r.json() if not ok else "")
        check(ok, "登録できる")
        if not ok:
            return 1
        rep = r.json()
        check(rep["version"] == 1, "初回はv1", str(rep["version"]))
        check(rep["status"] == "draft", "draftで保存される", rep["status"])
        check(rep["slides_removed"] > 0, "スライドが削除される",
              f"{rep['slides_removed']}枚")
        check(rep["size"]["base"] < rep["size"]["uploaded"], "ベースは軽くなる",
              f"{rep['size']['uploaded']} -> {rep['size']['base']}")
        check(rep["font"]["status"] in ("ok", "substituted", "approximated",
                                       "unavailable"), "フォント判定がある",
              f"{rep['font']['status']} ({rep['font']['declared']})")

        # 2. 一覧と取得（実体を返さないこと）
        r = c.get("/tenants/acme/templates")
        check(len(r.json()["templates"]) == 1, "一覧に出る")
        r = c.get("/tenants/acme/templates/corp")
        check("base" not in r.json(), "取得でpptx実体を返さない")

        # 3. draft のまま生成できる（発行前は latest が使われる）
        r = c.post("/render", json={"markdown": MD, "tenant_id": "acme",
                                    "template_id": "corp"})
        ok = r.status_code == 200 and len(r.content) > 10000
        print("render   ", r.status_code, f"{len(r.content)} bytes")
        check(ok, "テンプレート指定で生成できる")

        # 4. 発行してバージョン固定で生成
        r = c.post("/tenants/acme/templates/corp/publish", params={"version": 1})
        check(r.json()["status"] == "published", "発行できる")
        r = c.post("/render", json={"markdown": MD, "tenant_id": "acme",
                                    "template_id": "corp", "template_version": 1})
        check(r.status_code == 200, "バージョン指定で生成できる")

        # 5. 2回目の登録は v2 になり、v1 は書き換わらない
        r = c.post("/tenants/acme/templates",
                   data={"template_id": "corp"},
                   files={"file": ("corp.pptx", data, "application/octet-stream")})
        check(r.json()["version"] == 2, "再登録でv2になる", str(r.json()["version"]))
        r = c.get("/tenants/acme/templates/corp", params={"version": 1})
        check(r.json()["status"] == "published", "v1は発行済みのまま",
              r.json()["status"])

        # 6. 受入検査
        r = c.post("/tenants/acme/templates", data={"template_id": "bad"},
                   files={"file": ("x.pptx", b"not a zip", "application/octet-stream")})
        check(r.status_code == 400, "壊れたファイルを弾く", str(r.status_code))
        r = c.post("/tenants/acme/templates", data={"template_id": "../escape"},
                   files={"file": ("corp.pptx", data, "application/octet-stream")})
        check(r.status_code == 400, "不正なIDを弾く", str(r.status_code))
        r = c.post("/render", json={"markdown": MD, "tenant_id": "acme",
                                    "template_id": "nope"})
        check(r.status_code == 404, "未登録テンプレートは404", str(r.status_code))

    print("\n" + ("ALL PASSED" if not fails else f"FAILED: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
