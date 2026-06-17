"""Сохранение и загрузка проекта СОПТ в Excel."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from app.calculations.sopt.models import (
    EQUIPMENT_KEY_BY_LABEL,
    EQUIPMENT_LABEL_BY_KEY,
    KZ_POINT_ITEM_TYPE,
    KZ_POINT_LABEL,
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
    "Обозначение точки КЗ",
    "Проверка на селективность",
)
_PREV_ROW_HEADERS_14_INOM_FIRST = (
    "Раздел",
    "Подраздел",
    "Тип оборудования",
    "Iном, А",
    "Хар-ка",
    "Кратн.",
    "t, с",
    "R, Ом",
    "Условное обозначение",
    "L, м",
    "Y",
    "S, мм²",
    "Обозначение точки КЗ",
    "Проверка на селективность",
)
_PREV_ROW_HEADERS_15 = (
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
    "Точка КЗ",
    "Обозначение точки КЗ",
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
_KZ_HEADER_ALIASES = frozenset({"Точка КЗ", "Точка кз", "ТКЗ"})
_ROW_FIELD_COUNT = len(ROW_HEADERS)
_EquipmentRow = tuple[str, str, str, str, str, str, str, str, str, str, str, str, str, str]


def _parse_yes_cell(text: str) -> bool:
    value = text.strip().lower()
    return value in ("да", "+", "1", "x", "v", "yes", "true", "y", "✓")


def _parse_selectivity_cell(text: str) -> bool:
    return _parse_yes_cell(text)


def _format_designation(text: str) -> str:
    return text.strip()


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


def _detect_equipment_layout(cells: list[str]) -> str:
    if (
        len(cells) >= len(ROW_HEADERS)
        and cells[3] == "Условное обозначение"
        and cells[4] in _IN_HEADER_ALIASES
    ):
        return "new14"
    if (
        len(cells) >= len(_PREV_ROW_HEADERS_14_INOM_FIRST)
        and cells[3] in _IN_HEADER_ALIASES
        and cells[8] == "Условное обозначение"
    ):
        return "prev14_inom_first"
    if len(cells) >= len(_PREV_ROW_HEADERS_15) and cells[12] in _KZ_HEADER_ALIASES:
        return "prev15"
    if len(cells) >= 13 and cells[3] == "Условное обозначение" and cells[12] in _SEL_HEADER_ALIASES:
        return "prev13"
    return "new14"


def _row_headers_match(cells: list[str]) -> bool:
    if not cells or cells[0] != ROW_HEADERS[0]:
        return False
    if len(cells) >= len(ROW_HEADERS) and _headers_core_match(cells, len(ROW_HEADERS)):
        return True
    if len(cells) >= len(ROW_HEADERS) - 1 and _headers_core_match(cells, len(ROW_HEADERS) - 1):
        return True
    if len(cells) >= len(_PREV_ROW_HEADERS_15) and tuple(cells[: len(_PREV_ROW_HEADERS_15)]) == _PREV_ROW_HEADERS_15:
        return True
    if (
        len(cells) >= len(_PREV_ROW_HEADERS_14_INOM_FIRST)
        and tuple(cells[: len(_PREV_ROW_HEADERS_14_INOM_FIRST)]) == _PREV_ROW_HEADERS_14_INOM_FIRST
    ):
        return True
    if (
        len(cells) >= 13
        and cells[3] == "Условное обозначение"
        and cells[12] in _SEL_HEADER_ALIASES
    ):
        return True
    if len(cells) >= len(_LEGACY_ROW_HEADERS_10) and tuple(cells[: len(_LEGACY_ROW_HEADERS_10)]) == _LEGACY_ROW_HEADERS_10:
        return True
    return False


def _legacy_10col_headers(cells: list[str]) -> bool:
    return len(cells) >= len(_LEGACY_ROW_HEADERS_10) and tuple(cells[: len(_LEGACY_ROW_HEADERS_10)]) == _LEGACY_ROW_HEADERS_10


def _legacy_row_to_full(
    row: tuple[str, str, str, str, str, str, str, str, str, str],
) -> _EquipmentRow:
    return (
        row[0],
        row[1],
        row[2],
        _format_designation(row[3]),
        row[4],
        "",
        "",
        "",
        row[5],
        row[6],
        row[7],
        row[8],
        "",
        row[9],
    )


def _from_prev15_row(cells: list[str]) -> _EquipmentRow:
    kz_name = cells[13].strip() if len(cells) > 13 else ""
    if not kz_name and len(cells) > 12 and _parse_yes_cell(cells[12]):
        kz_name = "k1"
    return (
        cells[0],
        cells[1],
        cells[2],
        _format_designation(cells[3]),
        cells[4],
        cells[5],
        cells[6],
        cells[7],
        cells[8],
        cells[9],
        cells[10],
        cells[11],
        kz_name,
        cells[14] if len(cells) > 14 else "",
    )


def _from_prev13_row(cells: list[str]) -> _EquipmentRow:
    return (
        cells[0],
        cells[1],
        cells[2],
        _format_designation(cells[3]),
        cells[4] if len(cells) > 4 else "",
        cells[5] if len(cells) > 5 else "",
        cells[6] if len(cells) > 6 else "",
        cells[7] if len(cells) > 7 else "",
        cells[8] if len(cells) > 8 else "",
        cells[9] if len(cells) > 9 else "",
        cells[10] if len(cells) > 10 else "",
        cells[11] if len(cells) > 11 else "",
        "",
        cells[12] if len(cells) > 12 else "",
    )


def _from_new14_row(cells: list[str]) -> _EquipmentRow:
    padded = cells + [""] * (_ROW_FIELD_COUNT - len(cells))
    return (
        padded[0],
        padded[1],
        padded[2],
        _format_designation(padded[3]),
        padded[4],
        padded[5],
        padded[6],
        padded[7],
        padded[8],
        padded[9],
        padded[10],
        padded[11],
        padded[12],
        padded[13],
    )


def _from_prev14_inom_first_row(cells: list[str]) -> _EquipmentRow:
    padded = cells + [""] * (_ROW_FIELD_COUNT - len(cells))
    return (
        padded[0],
        padded[1],
        padded[2],
        _format_designation(padded[8]),
        padded[3],
        padded[4],
        padded[5],
        padded[6],
        padded[7],
        padded[9],
        padded[10],
        padded[11],
        padded[12],
        padded[13],
    )


def _normalize_equipment_row(row: tuple[str, ...], layout: str = "new14") -> _EquipmentRow:
    cells = list(row)
    if layout == "prev15":
        return _from_prev15_row(cells)
    if layout == "prev13":
        return _from_prev13_row(cells)
    if layout == "prev14_inom_first":
        return _from_prev14_inom_first_row(cells)
    if len(cells) >= 14 and cells[3] in _IN_HEADER_ALIASES and cells[8] == "Условное обозначение":
        return _from_prev14_inom_first_row(cells)
    if len(cells) >= 15 and cells[3] == "Условное обозначение" and cells[12] in _KZ_HEADER_ALIASES:
        return _from_prev15_row(cells)
    if len(cells) >= 13 and cells[3] == "Условное обозначение" and (
        len(cells) < 14 or cells[8] != "Условное обозначение"
    ):
        return _from_prev13_row(cells)
    return _from_new14_row(cells)


META_FIELDS: tuple[tuple[str, str], ...] = (
    ("Проект", "project_name"),
    ("АКБ: Название", "battery_name"),
    ("АКБ: Количество элементов, шт", "battery_cells_count"),
    ("АКБ: Количество АКБ, шт", "battery_count"),
    ("АКБ: Внутреннее сопротивление блока Rэл, Ом", "r_el"),
    ("АКБ: Удельное сопротивление перемычки ρ", "rho"),
    ("АКБ: Сечение перемычки S, мм²", "jumper_section"),
    ("Вводный предохранитель: Iном, А", "input_fuse_label"),
    ("Вводный предохранитель: Rпр, Ом", "input_fuse_resistance"),
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


def _collapse_meta_keys(meta: dict[str, str]) -> dict[str, str]:
    return {" ".join(key.split()): value for key, value in meta.items()}


def _meta_get(meta: dict[str, str], key: str) -> str:
    return str(meta.get(key, "")).strip()


def _parse_input_fuse_nominal_a(label: str) -> int | float | str:
    import re

    text = label.strip()
    if not text:
        return ""
    match = re.search(r"(\d+(?:[.,]\d+)?)", text)
    if not match:
        return text
    raw = match.group(1).replace(",", ".")
    try:
        value = float(raw)
    except ValueError:
        return text
    return int(value) if value == int(value) else value


def _input_fuse_label_from_nominal(raw: str) -> str:
    text = raw.strip()
    if not text:
        return ""
    if "А" in text.upper():
        return text
    return f"{text} А"


def _meta_value_for_save(snapshot: SoptProjectSnapshot, key: str, attr: str) -> object:
    if key == "Вводный предохранитель: Iном, А":
        return _parse_input_fuse_nominal_a(snapshot.input_fuse_label)
    return getattr(snapshot, attr)


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


def snapshot_to_sections(flat_rows: list[_EquipmentRow]) -> list[SoptSection]:
    section_order: list[str] = []
    subsection_order: dict[str, list[str]] = {}
    buckets: dict[str, dict[str, list[SoptEquipmentItem]]] = {}
    current_section = ""
    current_subsection = ""
    for row in flat_rows:
        (
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
            kz_name_text,
            sel_text,
        ) = row
        if section_name.strip():
            current_section = section_name.strip()
        if subsection_name.strip():
            current_subsection = subsection_name.strip()
        s_name = current_section or "Раздел"
        sub_name = current_subsection or "Подраздел 1"
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
        type_key = _type_key_from_label(type_label)
        if type_key == KZ_POINT_ITEM_TYPE:
            if type_label.strip() or designation.strip():
                buckets[s_name][sub_name].append(
                    SoptEquipmentItem(
                        item_type=KZ_POINT_ITEM_TYPE,
                        designation=designation.strip() or "k1",
                        include_in_selectivity=_parse_selectivity_cell(sel_text),
                    )
                )
            continue
        if not (type_label.strip() or designation.strip()):
            continue
        buckets[s_name][sub_name].append(
            SoptEquipmentItem(
                item_type=type_key,
                designation=_format_designation(designation),
                rated_current_a=rated_current_a,
                resistance_ohm=resistance_ohm,
                cable_length_m=cable_length_m,
                cable_gamma=cable_gamma,
                cable_section_mm2=cable_section_mm2,
                cb_curve=cb_curve,
                cb_multiplier=cb_multiplier,
                cb_trip_time_s=cb_trip_time_s,
            )
        )
        if kz_name_text.strip():
            buckets[s_name][sub_name].append(
                SoptEquipmentItem(
                    item_type=KZ_POINT_ITEM_TYPE,
                    designation=kz_name_text.strip(),
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
        ws.cell(row=row_idx, column=2, value=_meta_value_for_save(snapshot, key, attr))
        row_idx += 1
    row_idx += 2

    header_row = row_idx
    for col, title in enumerate(ROW_HEADERS, start=1):
        ws.cell(row=header_row, column=col, value=title)

    row_idx = header_row
    last_section_name = ""
    last_subsection_name = ""
    for section in snapshot.sections:
        section_name = section.name.strip() or "Раздел"
        if not section.subsections:
            row_idx += 1
            ws.cell(row=row_idx, column=1, value=section_name)
            last_section_name = section_name
            last_subsection_name = ""
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
                section_changed = section_name != last_section_name
                subsection_changed = subsection_name != last_subsection_name
                row_idx += 1
                ws.cell(row=row_idx, column=1, value=section_name if section_changed else "")
                ws.cell(
                    row=row_idx,
                    column=2,
                    value=subsection_name if section_changed or subsection_changed else "",
                )
                last_section_name = section_name
                last_subsection_name = subsection_name
                ws.cell(row=row_idx, column=3, value="")
                ws.cell(row=row_idx, column=4, value="")
                ws.cell(row=row_idx, column=5, value="")
                ws.cell(row=row_idx, column=6, value="")
                ws.cell(row=row_idx, column=7, value="")
                for col in range(8, _ROW_FIELD_COUNT + 1):
                    ws.cell(row=row_idx, column=col, value="")
                continue
            item_idx = 0
            while item_idx < len(subsection.items):
                item = subsection.items[item_idx]
                if item.item_type == KZ_POINT_ITEM_TYPE:
                    item_idx += 1
                    continue
                kz_item: SoptEquipmentItem | None = None
                if (
                    item_idx + 1 < len(subsection.items)
                    and subsection.items[item_idx + 1].item_type == KZ_POINT_ITEM_TYPE
                ):
                    kz_item = subsection.items[item_idx + 1]
                    item_idx += 2
                else:
                    item_idx += 1
                section_changed = section_name != last_section_name
                subsection_changed = subsection_name != last_subsection_name
                row_idx += 1
                label = EQUIPMENT_LABEL_BY_KEY.get(item.item_type, item.item_type)
                ws.cell(row=row_idx, column=1, value=section_name if section_changed else "")
                ws.cell(
                    row=row_idx,
                    column=2,
                    value=subsection_name if section_changed or subsection_changed else "",
                )
                last_section_name = section_name
                last_subsection_name = subsection_name
                ws.cell(row=row_idx, column=3, value=label)
                ws.cell(row=row_idx, column=4, value=_format_designation(item.designation))
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
                ws.cell(row=row_idx, column=13, value=kz_item.designation if kz_item is not None else "")
                sel_value = ""
                if kz_item is not None and kz_item.include_in_selectivity:
                    sel_value = "да"
                ws.cell(row=row_idx, column=14, value=sel_value)

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

    raw_flat: list[tuple[str, ...]] = []
    equipment_layout = "new14"
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
                equipment_layout = _detect_equipment_layout(cells)
            elif cells[0] == "Раздел" and len(cells) > 1 and cells[1] == "Тип оборудования":
                header_found = True
                old_layout_3col = True
            elif cells[:4] == ["Раздел", "Подраздел", "Тип оборудования", "Условное обозначение"]:
                header_found = True
                old_layout_4col = True
            elif cells[:8] == [
                "Раздел",
                "Подраздел",
                "Тип оборудования",
                "Условное обозначение",
                "Iном, А",
                "Хар-ка",
                "Кратн.",
                "t, с",
            ]:
                header_found = True
                old_layout_8col = True
            elif _legacy_10col_headers(cells):
                header_found = True
                old_layout_10col = True
            continue
        if old_layout_3col:
            padded = cells + [""] * (3 - len(cells))
            raw_flat.append(
                (padded[0], "Подраздел 1", padded[1], padded[2], "", "", "", "", "", "")
            )
        elif old_layout_4col:
            padded = cells + [""] * (4 - len(cells))
            raw_flat.append((padded[0], padded[1], padded[2], padded[3], "", "", "", "", "", ""))
        elif old_layout_8col:
            padded = cells + [""] * (8 - len(cells))
            raw_flat.append(
                (padded[0], padded[1], padded[2], padded[3], "", padded[4], padded[5], padded[6], padded[7], "")
            )
        elif old_layout_10col:
            padded = cells + [""] * (10 - len(cells))
            raw_flat.append(tuple(padded[:10]))  # type: ignore[arg-type]
        else:
            padded = cells + [""] * (max(_ROW_FIELD_COUNT, 15) - len(cells))
            raw_flat.append(tuple(padded))

    wb.close()

    meta = _collapse_meta_keys(meta)
    project_name = _meta_get(meta, META_PROJECT) or path.stem
    battery_name = _meta_get(meta, "АКБ: Название")
    try:
        battery_cells_count = int(_meta_get(meta, "АКБ: Количество элементов, шт") or "0")
    except ValueError:
        battery_cells_count = 0
    try:
        battery_count = int(_meta_get(meta, "АКБ: Количество АКБ, шт") or "0")
    except ValueError:
        battery_count = 0
    try:
        raw_r_el = _meta_get(meta, "АКБ: Внутреннее сопротивление блока Rэл, Ом")
        r_el = float(raw_r_el.replace(",", ".") or "0")
    except ValueError:
        r_el = 0.0
    try:
        rho = float(_meta_get(meta, "АКБ: Удельное сопротивление перемычки ρ").replace(",", ".") or "0")
    except ValueError:
        rho = 0.0
    try:
        jumper_section = float(_meta_get(meta, "АКБ: Сечение перемычки S, мм²").replace(",", ".") or "0")
    except ValueError:
        jumper_section = 0.0
    input_fuse_label = _input_fuse_label_from_nominal(_meta_get(meta, "Вводный предохранитель: Iном, А"))
    try:
        input_fuse_resistance = float(
            _meta_get(meta, "Вводный предохранитель: Rпр, Ом").replace(",", ".") or "0"
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
    flat: list[_EquipmentRow] = []
    for row in raw_flat:
        if old_layout_3col or old_layout_4col or old_layout_8col or old_layout_10col:
            flat.append(_legacy_row_to_full(row))  # type: ignore[arg-type]
        else:
            flat.append(_normalize_equipment_row(row, equipment_layout))
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
