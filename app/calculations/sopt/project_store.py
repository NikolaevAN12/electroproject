"""Сохранение и загрузка проекта СОПТ в Excel."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from app.calculations.sopt.models import (
    EQUIPMENT_KEY_BY_LABEL,
    EQUIPMENT_LABEL_BY_KEY,
    SoptEquipmentItem,
    SoptSection,
    SoptSubsection,
)
from app.shared.excel_project_base import (
    cell_str,
    dialog_initial_dir as _dialog_initial_dir,
    last_project_marker,
    read_last_project_path as _read_last_project_path,
    remember_last_project as _remember_last_project,
    require_openpyxl,
    safe_stem,
)
SHEET = "СОПТ"
META_PROJECT = "Проект"
ROW_HEADERS = (
    "Раздел",
    "Подраздел",
    "Тип оборудования",
    "Условное обозначение",
    "Iном, А",
    "Хар-ка",
    "Кратн.",
    "t, с",
    "R, Ом",
    "L, м",
    "Y",
    "S, мм²",
    "Проверка на селективность",
)
_LEGACY_ROW_HEADERS_10 = (
    "Раздел",
    "Подраздел",
    "Тип оборудования",
    "Условное обозначение",
    "Iном, А",
    "R, Ом",
    "L, м",
    "Y",
    "S, мм²",
    "Разд. 4",
)
_IN_HEADER_ALIASES = frozenset({"Iном, А", "Номинальный ток, А"})
_SEL_HEADER_ALIASES = frozenset({
    "Разд. 4",
    "Селективность",
    "ТКЗ в селективности",
    "ТКЗ в селект.",
    "Проверка на селективность",
})
_ROW_FIELD_COUNT = len(ROW_HEADERS)


def _parse_selectivity_cell(text: str) -> bool:
    value = text.strip().lower()
    return value in ("да", "+", "1", "x", "v", "yes", "true", "y", "✓")


def _headers_core_match(cells: list[str], count: int) -> bool:
    for idx in range(count):
        expected = ROW_HEADERS[idx]
        if idx == 4:
            if cells[idx] not in _IN_HEADER_ALIASES:
                return False
        elif idx == len(ROW_HEADERS) - 1:
            if cells[idx] not in _SEL_HEADER_ALIASES and cells[idx] != expected:
                return False
        elif cells[idx] != expected:
            return False
    return True


def _row_headers_match(cells: list[str]) -> bool:
    if not cells or cells[0] != ROW_HEADERS[0]:
        return False
    if len(cells) >= len(ROW_HEADERS) and _headers_core_match(cells, len(ROW_HEADERS)):
        return True
    if len(cells) >= len(ROW_HEADERS) - 1 and _headers_core_match(cells, len(ROW_HEADERS) - 1):
        return True
    if len(cells) >= len(_LEGACY_ROW_HEADERS_10) and tuple(cells[: len(_LEGACY_ROW_HEADERS_10)]) == _LEGACY_ROW_HEADERS_10:
        return True
    return False


def _legacy_10col_headers(cells: list[str]) -> bool:
    return len(cells) >= len(_LEGACY_ROW_HEADERS_10) and tuple(cells[: len(_LEGACY_ROW_HEADERS_10)]) == _LEGACY_ROW_HEADERS_10


def _legacy_row_to_full(
    row: tuple[str, str, str, str, str, str, str, str, str, str],
) -> tuple[str, str, str, str, str, str, str, str, str, str, str, str, str]:
    return (
        row[0],
        row[1],
        row[2],
        row[3],
        row[4],
        "",
        "",
        "",
        row[5],
        row[6],
        row[7],
        row[8],
        row[9],
    )


META_FIELDS: tuple[tuple[str, str], ...] = (
    ("Проект", "project_name"),
    ("АКБ: Название", "battery_name"),
    ("АКБ: Количество элементов", "battery_cells_count"),
    ("АКБ: Количество АКБ", "battery_count"),
    ("АКБ: Rэл", "r_el"),
    ("АКБ: ρ", "rho"),
    ("АКБ: S", "jumper_section"),
    ("Вводный предохранитель", "input_fuse_label"),
    ("Вводный предохранитель: Rпр", "input_fuse_resistance"),
    ("QN=1, А·ч", "qn1_ah"),
    ("Qрасч, А·ч", "q_calc_ah"),
)


@dataclass(slots=True)
class SoptProjectSnapshot:
    project_name: str
    battery_name: str = ""
    battery_cells_count: int = 0
    battery_count: int = 0
    r_el: float = 0.0
    rho: float = 0.0
    jumper_section: float = 0.0
    input_fuse_label: str = ""
    input_fuse_resistance: float = 0.0
    qn1_ah: float = 0.0
    q_calc_ah: float = 0.0
    sections: list[SoptSection] = field(default_factory=list)


_MARKER = last_project_marker("sopt_last_project.txt")


def default_excel_filename(project_name: str) -> str:
    return f"{safe_stem(project_name)}.xlsx"


def remember_last_project(path: Path) -> None:
    _remember_last_project(_MARKER, path)


def read_last_project_path() -> Path | None:
    return _read_last_project_path(_MARKER)


def dialog_initial_dir() -> Path:
    return _dialog_initial_dir(_MARKER)


def _type_key_from_label(label: str) -> str:
    text = label.strip()
    if text in EQUIPMENT_KEY_BY_LABEL:
        return EQUIPMENT_KEY_BY_LABEL[text]
    for key, ru in EQUIPMENT_LABEL_BY_KEY.items():
        if ru.lower() == text.lower():
            return key
    return "fuse"


def snapshot_to_sections(
    flat_rows: list[tuple[str, str, str, str, str, str, str, str, str, str, str, str, str]],
) -> list[SoptSection]:
    section_order: list[str] = []
    subsection_order: dict[str, list[str]] = {}
    buckets: dict[str, dict[str, list[SoptEquipmentItem]]] = {}
    for (
        section_name,
        subsection_name,
        type_label,
        designation,
        in_text,
        curve_text,
        mult_text,
        time_text,
        r_text,
        l_text,
        y_text,
        s_text,
        sel_text,
    ) in flat_rows:
        s_name = section_name.strip() or "Раздел"
        sub_name = subsection_name.strip() or "Подраздел 1"
        if s_name not in buckets:
            section_order.append(s_name)
            subsection_order[s_name] = []
            buckets[s_name] = {}
        if sub_name not in buckets[s_name]:
            subsection_order[s_name].append(sub_name)
            buckets[s_name][sub_name] = []
        try:
            rated_current_a = float(in_text.replace(",", ".") or "0")
        except ValueError:
            rated_current_a = 0.0
        try:
            resistance_ohm = float(r_text.replace(",", ".") or "0")
        except ValueError:
            resistance_ohm = 0.0
        try:
            cable_length_m = float(l_text.replace(",", ".") or "0")
        except ValueError:
            cable_length_m = 0.0
        try:
            cable_gamma = float(y_text.replace(",", ".") or "0")
        except ValueError:
            cable_gamma = 0.0
        try:
            cable_section_mm2 = float(s_text.replace(",", ".") or "0")
        except ValueError:
            cable_section_mm2 = 0.0
        try:
            cb_multiplier = float(mult_text.replace(",", ".") or "0")
        except ValueError:
            cb_multiplier = 0.0
        try:
            cb_trip_time_s = float(time_text.replace(",", ".") or "0")
        except ValueError:
            cb_trip_time_s = 0.0
        cb_curve = curve_text.strip().upper().replace("NAN", "")
        if cb_curve == "—":
            cb_curve = ""
        if type_label.strip() or designation.strip():
            buckets[s_name][sub_name].append(
                SoptEquipmentItem(
                    item_type=_type_key_from_label(type_label),
                    designation=designation.strip(),
                    rated_current_a=rated_current_a,
                    resistance_ohm=resistance_ohm,
                    cable_length_m=cable_length_m,
                    cable_gamma=cable_gamma,
                    cable_section_mm2=cable_section_mm2,
                    cb_curve=cb_curve,
                    cb_multiplier=cb_multiplier,
                    cb_trip_time_s=cb_trip_time_s,
                    include_in_selectivity=_parse_selectivity_cell(sel_text),
                )
            )
    sections: list[SoptSection] = []
    for section_name in section_order:
        subsections = [
            SoptSubsection(
                name=sub_name,
                items=buckets[section_name][sub_name],
            )
            for sub_name in subsection_order[section_name]
        ]
        sections.append(SoptSection(name=section_name, subsections=subsections))
    return sections


def save_project(snapshot: SoptProjectSnapshot, path: Path) -> Path:
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
    row_idx = 2
    for key, attr in META_FIELDS:
        ws.cell(row=row_idx, column=1, value=key)
        ws.cell(row=row_idx, column=2, value=getattr(snapshot, attr))
        row_idx += 1
    ws.cell(row=row_idx, column=1, value="Сохранено")
    ws.cell(row=row_idx, column=2, value=datetime.now().strftime("%Y-%m-%d %H:%M"))

    header_row = row_idx + 2
    for col, title in enumerate(ROW_HEADERS, start=1):
        ws.cell(row=header_row, column=col, value=title)

    row_idx = header_row
    for section in snapshot.sections:
        section_name = section.name.strip() or "Раздел"
        if not section.subsections:
            row_idx += 1
            ws.cell(row=row_idx, column=1, value=section_name)
            ws.cell(row=row_idx, column=2, value="")
            ws.cell(row=row_idx, column=3, value="")
            ws.cell(row=row_idx, column=4, value="")
            ws.cell(row=row_idx, column=5, value="")
            ws.cell(row=row_idx, column=6, value="")
            ws.cell(row=row_idx, column=7, value="")
            ws.cell(row=row_idx, column=8, value="")
            for col in range(9, _ROW_FIELD_COUNT + 1):
                ws.cell(row=row_idx, column=col, value="")
            continue
        for subsection in section.subsections:
            subsection_name = subsection.name.strip() or "Подраздел 1"
            if not subsection.items:
                row_idx += 1
                ws.cell(row=row_idx, column=1, value=section_name)
                ws.cell(row=row_idx, column=2, value=subsection_name)
                ws.cell(row=row_idx, column=3, value="")
                ws.cell(row=row_idx, column=4, value="")
                ws.cell(row=row_idx, column=5, value="")
                ws.cell(row=row_idx, column=6, value="")
                ws.cell(row=row_idx, column=7, value="")
                for col in range(8, _ROW_FIELD_COUNT + 1):
                    ws.cell(row=row_idx, column=col, value="")
                continue
            for item in subsection.items:
                row_idx += 1
                label = EQUIPMENT_LABEL_BY_KEY.get(item.item_type, item.item_type)
                ws.cell(row=row_idx, column=1, value=section_name)
                ws.cell(row=row_idx, column=2, value=subsection_name)
                ws.cell(row=row_idx, column=3, value=label)
                ws.cell(row=row_idx, column=4, value=item.designation)
                ws.cell(
                    row=row_idx,
                    column=5,
                    value=item.rated_current_a if item.rated_current_a > 0 else "",
                )
                curve_value = ""
                mult_value = ""
                time_value = ""
                if item.item_type == "breaker":
                    curve_value = item.cb_curve or ""
                    mult_value = item.cb_multiplier if item.cb_multiplier > 0 else ""
                    time_value = item.cb_trip_time_s if item.cb_trip_time_s > 0 else ""
                ws.cell(row=row_idx, column=6, value=curve_value)
                ws.cell(row=row_idx, column=7, value=mult_value)
                ws.cell(row=row_idx, column=8, value=time_value)
                ws.cell(row=row_idx, column=9, value=item.resistance_ohm if item.resistance_ohm > 0 else "")
                ws.cell(row=row_idx, column=10, value=item.cable_length_m if item.cable_length_m > 0 else "")
                ws.cell(row=row_idx, column=11, value=item.cable_gamma if item.cable_gamma > 0 else "")
                ws.cell(row=row_idx, column=12, value=item.cable_section_mm2 if item.cable_section_mm2 > 0 else "")
                sel_value = ""
                if item.item_type == "kz_point" and item.include_in_selectivity:
                    sel_value = "да"
                ws.cell(row=row_idx, column=13, value=sel_value)

    wb.save(target)
    remember_last_project(target)
    return target


def load_project(path: Path) -> SoptProjectSnapshot:
    openpyxl = require_openpyxl()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active

    meta: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
        key = cell_str(row[0])
        if not key:
            continue
        if key == ROW_HEADERS[0]:
            break
        meta[key] = cell_str(row[1] if len(row) > 1 else "")

    flat: list[tuple[str, str, str, str, str, str, str, str, str, str, str, str, str]] = []
    header_found = False
    old_layout_3col = False
    old_layout_4col = False
    old_layout_8col = False
    old_layout_10col = False
    for row in ws.iter_rows(min_row=2, values_only=True):
        cells = [cell_str(c) for c in row]
        if not any(cells):
            if header_found:
                break
            continue
        if not header_found:
            if _row_headers_match(cells):
                header_found = True
            elif cells[0] == "Раздел" and len(cells) > 1 and cells[1] == "Тип оборудования":
                header_found = True
                old_layout_3col = True
            elif cells[:4] == ["Раздел", "Подраздел", "Тип оборудования", "Условное обозначение"]:
                header_found = True
                old_layout_4col = True
            elif cells[:8] == list(ROW_HEADERS)[:8] and cells[0] == "Раздел":
                header_found = True
                old_layout_8col = True
            elif _legacy_10col_headers(cells):
                header_found = True
                old_layout_10col = True
            continue
        if old_layout_3col:
            padded = cells + [""] * (3 - len(cells))
            flat.append(
                _legacy_row_to_full((padded[0], "Подраздел 1", padded[1], padded[2], "", "", "", "", "", ""))
            )
        elif old_layout_4col:
            padded = cells + [""] * (4 - len(cells))
            flat.append(_legacy_row_to_full((padded[0], padded[1], padded[2], padded[3], "", "", "", "", "", "")))
        elif old_layout_8col:
            padded = cells + [""] * (8 - len(cells))
            flat.append(
                _legacy_row_to_full(
                    (padded[0], padded[1], padded[2], padded[3], "", padded[4], padded[5], padded[6], padded[7], "")
                )
            )
        elif old_layout_10col:
            padded = cells + [""] * (10 - len(cells))
            flat.append(_legacy_row_to_full(tuple(padded[:10])))  # type: ignore[arg-type]
        else:
            padded = cells + [""] * (_ROW_FIELD_COUNT - len(cells))
            flat.append(tuple(padded[:_ROW_FIELD_COUNT]))  # type: ignore[arg-type]

    wb.close()

    project_name = meta.get(META_PROJECT, "").strip() or path.stem
    battery_name = meta.get("АКБ: Название", "").strip()
    try:
        battery_cells_count = int(meta.get("АКБ: Количество элементов", "").strip() or "0")
    except ValueError:
        battery_cells_count = 0
    try:
        battery_count = int(meta.get("АКБ: Количество АКБ", "").strip() or "0")
    except ValueError:
        battery_count = 0
    try:
        raw_r_el = meta.get("АКБ: Rэл", "").strip() or meta.get("АКБ: r_эл", "").strip()
        r_el = float(raw_r_el.replace(",", ".") or "0")
    except ValueError:
        r_el = 0.0
    try:
        rho = float(meta.get("АКБ: ρ", "").strip().replace(",", ".") or "0")
    except ValueError:
        rho = 0.0
    try:
        jumper_section = float(meta.get("АКБ: S", "").strip().replace(",", ".") or "0")
    except ValueError:
        jumper_section = 0.0
    input_fuse_label = meta.get("Вводный предохранитель", "").strip()
    try:
        input_fuse_resistance = float(
            meta.get("Вводный предохранитель: Rпр", "").strip().replace(",", ".") or "0"
        )
    except ValueError:
        input_fuse_resistance = 0.0
    try:
        qn1_ah = float(meta.get("QN=1, А·ч", "").strip().replace(",", ".") or "0")
    except ValueError:
        qn1_ah = 0.0
    try:
        q_calc_ah = float(meta.get("Qрасч, А·ч", "").strip().replace(",", ".") or "0")
    except ValueError:
        q_calc_ah = 0.0
    sections = snapshot_to_sections(flat)
    remember_last_project(path)
    return SoptProjectSnapshot(
        project_name=project_name,
        battery_name=battery_name,
        battery_cells_count=battery_cells_count,
        battery_count=battery_count,
        r_el=r_el,
        rho=rho,
        jumper_section=jumper_section,
        input_fuse_label=input_fuse_label,
        input_fuse_resistance=input_fuse_resistance,
        qn1_ah=qn1_ah,
        q_calc_ah=q_calc_ah,
        sections=sections,
    )
