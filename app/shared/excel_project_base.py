"""Общие функции для сохранения проектов в Excel."""

from __future__ import annotations

import re
from pathlib import Path

from app.shared.user_errors import MSG_MISSING_OPENPYXL


def safe_stem(name: str) -> str:
    text = name.strip()
    text = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:120] or "proekt"


def user_settings_dir() -> Path:
    directory = Path.home() / ".electroproject"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def last_project_marker(filename: str) -> Path:
    return user_settings_dir() / filename


def remember_last_project(marker_path: Path, path: Path) -> None:
    marker_path.write_text(str(path.resolve()), encoding="utf-8")


def read_last_project_path(marker_path: Path) -> Path | None:
    if not marker_path.is_file():
        return None
    try:
        raw = marker_path.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    if not raw:
        return None
    path = Path(raw)
    return path if path.is_file() else None


def dialog_initial_dir(marker_path: Path) -> Path:
    """Каталог для диалога Open/Save: папка последнего файла или домашний."""
    last = read_last_project_path(marker_path)
    if last is not None:
        return last.parent
    return Path.home()


def require_openpyxl():
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError(MSG_MISSING_OPENPYXL) from exc
    return openpyxl


def cell_str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
