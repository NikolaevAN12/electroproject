"""Сохранение и загрузка модели сети ЩСН в Excel."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.shared.excel_project_base import (
    dialog_initial_dir as _dialog_initial_dir,
    last_project_marker,
    read_last_project_path as _read_last_project_path,
    remember_last_project as _remember_last_project,
    safe_stem,
)

from .excel_io import load_network_excel, save_blank_template, save_network_excel
from .models import DEFAULT_K_TEMP, NetworkElement

_MARKER = last_project_marker("shsn_last_project.txt")
TEMPLATE_NAME = "Модель сети 0,4 кВ.xlsx"


@dataclass(slots=True)
class ShsnProjectSnapshot:
    project_name: str
    k_temp: float
    elements: list[NetworkElement]


def default_excel_filename(project_name: str) -> str:
    return f"{safe_stem(project_name or 'ЩСН')}.xlsx"


def remember_last_project(path: Path) -> None:
    _remember_last_project(_MARKER, path)


def read_last_project_path() -> Path | None:
    return _read_last_project_path(_MARKER)


def dialog_initial_dir() -> Path:
    return _dialog_initial_dir(_MARKER)


def save_project(snapshot: ShsnProjectSnapshot, path: Path) -> Path:
    target = save_network_excel(snapshot.elements, path)
    remember_last_project(target)
    return target


def load_project(path: Path) -> ShsnProjectSnapshot:
    elements = load_network_excel(path)
    stem = path.stem
    return ShsnProjectSnapshot(
        project_name=stem,
        k_temp=DEFAULT_K_TEMP,
        elements=elements,
    )


def create_blank_template(path: Path) -> Path:
    return save_blank_template(path)
