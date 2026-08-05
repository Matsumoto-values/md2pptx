#!/usr/bin/env python3
"""theme.py -- デザイントークンの読み込みとマージ.

CSS(プレビュー)とpython-pptx(本番)の配色を単一の正に保つための層。
テナントごとのテーマは themes/<id>.json、またはリクエストで直接渡す。
"""
import copy
import json
import pathlib

THEMES_DIR = pathlib.Path(__file__).parent / "themes"

DEFAULTS = {
    "canvas": {"w": 1280, "h": 720, "slide_in": [13.333, 7.5]},
    "font": "Noto Sans JP",
    "palette": ["4C6FFF", "12B5A5", "F5A524", "E5484D", "8E4EC6"],
    "text": "333333",
    "heading": "1B1B1F",
    "grid": "E6E6E6",
    "axis": "CCCCCC",
    "margin": {"x": 80, "top": 64, "body_top": 170},
    "table": {"header_bg": "1B1B1F", "header_fg": "FFFFFF", "row_bg": "FFFFFF",
              "row_alt_bg": "F5F6F8", "border": "DDDDDD", "border_w": 1,
              "pad": [12, 16, 12, 16], "zebra": True},
    "chart": {"gap_width": 60, "data_labels": True, "legend": "bottom",
              "gridlines": True},
    "number_format": "#,##0",
}


def deep_merge(base, over):
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def load(theme_id="default", override=None):
    """themes/<id>.json を DEFAULTS に重ね、さらに override を重ねて返す."""
    path = THEMES_DIR / f"{theme_id}.json"
    if not path.is_file():
        raise FileNotFoundError(f"unknown theme: {theme_id}")
    return deep_merge(deep_merge(DEFAULTS, json.loads(path.read_text())), override)


def available():
    return sorted(p.stem for p in THEMES_DIR.glob("*.json"))
