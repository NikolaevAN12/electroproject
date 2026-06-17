"""Загрузка и сохранение модели сети 0,4 кВ в Excel (формат Colab-скрипта)."""

from __future__ import annotations

import math
from pathlib import Path

from app.shared.excel_project_base import cell_str, require_openpyxl

from .models import REQUIRED_COLUMNS, NetworkElement

_TEMPLATE_SHEET = "Модель сети"
_BOOL_TRUE = frozenset({"1", "true", "yes", "да", "x", "v", "+"})


def _parse_float(value: object, default: float = float("nan")) -> float:
    text = cell_str(value)
    if not text:
        return default
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return default


def _parse_bool(value: object) -> bool:
    if value is True:
        return True
    if value is False:
        return False
    if isinstance(value, (int, float)) and not (isinstance(value, float) and math.isnan(value)):
        return bool(value)
    text = cell_str(value).lower()
    return text in _BOOL_TRUE


def load_network_excel(path: Path) -> list[NetworkElement]:
    openpyxl = require_openpyxl()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[_TEMPLATE_SHEET] if _TEMPLATE_SHEET in wb.sheetnames else wb.active

    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []

    headers = [cell_str(c) for c in rows[0]]
    col_index = {name: idx for idx, name in enumerate(headers)}
    for col in REQUIRED_COLUMNS:
        if col not in col_index:
            col_index[col] = -1

    elements: list[NetworkElement] = []
    for row in rows[1:]:
        if not row or not any(cell_str(c) for c in row):
            continue

        def cell(col: str) -> object:
            idx = col_index[col]
            if idx < 0 or idx >= len(row):
                return ""
            return row[idx]

        name = cell_str(cell("Name"))
        if not name:
            continue

        parent = cell_str(cell("Parent_Name")) or "нет"
        phase = cell_str(cell("Phase_Type")).upper()
        curve = cell_str(cell("CB_Curve")).upper().replace("NAN", "").strip()

        elements.append(
            NetworkElement(
                name=name,
                type=cell_str(cell("Type")),
                parent_name=parent,
                phase_type=phase,
                r1=_parse_float(cell("R1")),
                x1=_parse_float(cell("X1")),
                r0=_parse_float(cell("R0")),
                x0=_parse_float(cell("X0")),
                u_nom=_parse_float(cell("U_nom"), 0.4),
                is_sc_point=_parse_bool(cell("Is_SC_Point")),
                cb_nominal=_parse_float(cell("CB_Nominal")),
                cb_multiplier=_parse_float(cell("CB_Multiplier")),
                cb_time=_parse_float(cell("CB_Time")),
                cb_curve=curve,
            )
        )
    return elements


def save_network_excel(elements: list[NetworkElement], path: Path) -> Path:
    openpyxl = require_openpyxl()
    target = Path(path)
    if target.suffix.lower() != ".xlsx":
        target = target.with_suffix(".xlsx")
    target.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _TEMPLATE_SHEET
    for col, header in enumerate(REQUIRED_COLUMNS, start=1):
        ws.cell(row=1, column=col, value=header)

    for row_idx, el in enumerate(elements, start=2):
        ws.cell(row=row_idx, column=1, value=el.name)
        ws.cell(row=row_idx, column=2, value=el.type)
        ws.cell(row=row_idx, column=3, value=el.parent_name)
        ws.cell(row=row_idx, column=4, value=el.phase_type)
        ws.cell(row=row_idx, column=5, value=el.r1)
        ws.cell(row=row_idx, column=6, value=el.x1)
        ws.cell(row=row_idx, column=7, value=el.r0)
        ws.cell(row=row_idx, column=8, value=el.x0)
        ws.cell(row=row_idx, column=9, value=el.u_nom)
        ws.cell(row=row_idx, column=10, value=el.is_sc_point)
        ws.cell(row=row_idx, column=11, value=el.cb_nominal)
        ws.cell(row=row_idx, column=12, value=el.cb_multiplier)
        ws.cell(row=row_idx, column=13, value=el.cb_time)
        ws.cell(row=row_idx, column=14, value=el.cb_curve)

    wb.save(target)
    return target


def save_blank_template(path: Path) -> Path:
    return save_network_excel([], path)
