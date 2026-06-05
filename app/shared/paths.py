from __future__ import annotations

import sys
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def word_export_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return project_root() / "dist"


def embedded_build_tag() -> str:
    if not getattr(sys, "frozen", False):
        return ""
    meipass = getattr(sys, "_MEIPASS", None)
    if not meipass:
        return ""
    try:
        return Path(meipass).joinpath("build_tag.txt").read_text(encoding="utf-8").strip()
    except OSError:
        return ""

