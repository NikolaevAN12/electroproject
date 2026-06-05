"""Сохранение и загрузка введённых данных расчёта в Excel."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.shared.excel_project_base import (
    cell_str,
    dialog_initial_dir as _dialog_initial_dir,
    last_project_marker,
    read_last_project_path as _read_last_project_path,
    remember_last_project as _remember_last_project,
    require_openpyxl,
    safe_stem,
)
SHEET = "расчёт"
META_PROJECT = "Проект"
META_R1 = "R1 (мОм)"
META_X1 = "X1 (мОм)"
META_AUTO_IK = "Авто Iкз из R1/X1"
ROW_HEADERS = (
    "№",
    "Кабель",
    "Сечение, мм²",
    "Допустимый ток, А",
    "Рабочий ток, А",
    "Ток КЗ, кА",
    "Время отключения, с",
)

_MARKER = last_project_marker("fire_check_last_project.txt")


@dataclass(slots=True)
class FireCheckProjectSnapshot:
    project_name: str
    r1_mohm: str
    x1_mohm: str
    auto_ik1: bool
    rows: list[dict[str, str]]


def default_excel_filename(project_name: str) -> str:
    return f"{safe_stem(project_name)}.xlsx"


def remember_last_project(path: Path) -> None:
    _remember_last_project(_MARKER, path)


def read_last_project_path() -> Path | None:
    return _read_last_project_path(_MARKER)


def dialog_initial_dir() -> Path:
    return _dialog_initial_dir(_MARKER)


def save_project(snapshot: FireCheckProjectSnapshot, path: Path) -> Path:
    openpyxl = require_openpyxl()
    target = Path(path)
    if target.suffix.lower() != ".xlsx":
        target = target.with_suffix(".xlsx")
    target.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    ws["A1"] = "Поле"
    ws["B1"] = "Значение"
    ws["A2"] = META_PROJECT
    ws["B2"] = snapshot.project_name
    ws["A3"] = META_R1
    ws["B3"] = snapshot.r1_mohm
    ws["A4"] = META_X1
    ws["B4"] = snapshot.x1_mohm
    ws["A5"] = META_AUTO_IK
    ws["B5"] = "1" if snapshot.auto_ik1 else "0"
    ws["A6"] = "Сохранено"
    ws["B6"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    header_row = 8
    for col, title in enumerate(ROW_HEADERS, start=1):
        ws.cell(row=header_row, column=col, value=title)

    for offset, row in enumerate(snapshot.rows, start=1):
        r = header_row + offset
        ws.cell(row=r, column=1, value=row.get("num", str(offset)))
        ws.cell(row=r, column=2, value=row.get("cable", ""))
        ws.cell(row=r, column=3, value=row.get("sect", ""))
        ws.cell(row=r, column=4, value=row.get("allow", ""))
        ws.cell(row=r, column=5, value=row.get("work", ""))
        ws.cell(row=r, column=6, value=row.get("ikz", ""))
        ws.cell(row=r, column=7, value=row.get("tok", ""))

    wb.save(target)
    remember_last_project(target)
    return target


def load_project(path: Path) -> FireCheckProjectSnapshot:
    openpyxl = require_openpyxl()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active

    meta: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, max_row=6, min_col=1, max_col=2, values_only=True):
        key = cell_str(row[0])
        if key in (META_PROJECT, META_R1, META_X1, META_AUTO_IK):
            meta[key] = cell_str(row[1] if len(row) > 1 else "")

    rows: list[dict[str, str]] = []
    header_found = False
    for row in ws.iter_rows(min_row=8, values_only=True):
        cells = [cell_str(c) for c in row]
        if not any(cells):
            if header_found:
                break
            continue
        if not header_found:
            if cells[0] == ROW_HEADERS[0] and cells[: len(ROW_HEADERS)] == list(ROW_HEADERS):
                header_found = True
            continue
        padded = cells + [""] * (7 - len(cells))
        rows.append(
            {
                "num": padded[0],
                "cable": padded[1],
                "sect": padded[2],
                "allow": padded[3],
                "work": padded[4],
                "ikz": padded[5],
                "tok": padded[6],
            }
        )

    wb.close()

    auto_raw = meta.get(META_AUTO_IK, "1").strip().lower()
    auto_ik1 = auto_raw not in ("0", "нет", "false", "no")
    project_name = meta.get(META_PROJECT, "").strip() or path.stem

    snapshot = FireCheckProjectSnapshot(
        project_name=project_name,
        r1_mohm=meta.get(META_R1, ""),
        x1_mohm=meta.get(META_X1, ""),
        auto_ik1=auto_ik1,
        rows=rows,
    )
    remember_last_project(path)
    return snapshot
