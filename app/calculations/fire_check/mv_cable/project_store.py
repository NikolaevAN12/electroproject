"""Загрузка Excel-шаблона расчёта КЛ свыше 1 кВ."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.shared.excel_project_base import cell_str, require_openpyxl
from app.shared.parsing import parse_number

from .excel_template import (
    META_IKZ3,
    META_PROJECT,
    META_Q0,
    META_Q_OKR,
    META_RU_TYPE,
    META_SUBSTATION,
    META_T_BREAKER,
    META_U_N,
    ROW_HEADERS,
    SHEET,
)
from .conductor_phrase import (
    build_cable_route_text,
    build_existing_cable_route_text,
    conductor_nominative,
    normalize_cable_mark,
    parse_conductor_material,
    validate_cores_count,
)
from .models import MvCableFireInput

_META_KEYS = frozenset(
    {
        META_PROJECT,
        META_U_N,
        META_RU_TYPE,
        META_SUBSTATION,
        META_IKZ3,
        META_T_BREAKER,
        META_Q0,
        META_Q_OKR,
    }
)

_STANDARD_LINE_TYPES = frozenset({"ТСН", "ДГР"})

_ROW_KEYS = (
    "cable_no",
    "line_type",
    "power_kva",
    "i_nom_a",
    "cable_installation",
    "cable_mark",
    "cores",
    "section_mm2",
    "conductor_material",
    "sheath_material_gen",
    "i_dd_a",
    "theta_dd_c",
    "t_main_s",
    "t_backup_s",
    "tae_s",
    "b_const",
    "theta_limit_heating_c",
    "theta_limit_fire_c",
)


@dataclass(slots=True)
class MvCableProjectSnapshot:
    project_name: str
    voltage_kv: float
    ru_type: str
    substation_name: str
    ikz3_ka: float
    t_breaker_s: float
    theta_0_c: float
    theta_okr_c: float
    rows: list[dict[str, str]]


def _require_float(value: object, *, field: str) -> float:
    parsed = parse_number(cell_str(value))
    if parsed is None:
        raise ValueError(f"Не заполнено или неверное число: {field}")
    return parsed


def _require_int(value: object, *, field: str) -> int:
    parsed = parse_number(cell_str(value))
    if parsed is None:
        raise ValueError(f"Не заполнено или неверное число: {field}")
    return int(parsed)


def _require_text(value: object, *, field: str) -> str:
    text = cell_str(value)
    if not text:
        raise ValueError(f"Не заполнено поле: {field}")
    return text


def _section_label(section_mm2: float) -> str:
    if abs(section_mm2 - round(section_mm2)) < 1e-9:
        return str(int(round(section_mm2)))
    return str(section_mm2).replace(".", ",")


def _cable_designation(mark: str, cores: int, section_mm2: float) -> str:
    return f"{mark} {cores}х{_section_label(section_mm2)}"


def _source_zru_label(ru_type: str, voltage_kv: float) -> str:
    ru = ru_type.strip()
    if not ru:
        raise ValueError("Не заполнено поле: Тип РУ")
    return f"{ru} {int(voltage_kv)} кВ"


def _read_meta(ws: object) -> dict[str, str]:
    meta: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, max_row=20, min_col=1, max_col=2, values_only=True):
        key = cell_str(row[0] if row else "")
        if key not in _META_KEYS:
            continue
        meta[key] = cell_str(row[1] if len(row) > 1 else "")
    return meta


def _find_header_row(ws: object) -> int:
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=40, values_only=True), start=1):
        cells = [cell_str(c) for c in row]
        if cells[: len(ROW_HEADERS)] == list(ROW_HEADERS):
            return row_idx
    raise ValueError("Не найдена строка заголовков таблицы кабелей в Excel")


def _read_rows(ws: object, header_row: int) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        cells = [cell_str(c) for c in row]
        if not any(cells):
            break
        padded = cells + [""] * (len(_ROW_KEYS) - len(cells))
        if not padded[1].strip():
            continue
        rows.append({key: padded[idx] for idx, key in enumerate(_ROW_KEYS)})
    return rows


def load_mv_cable_project(path: Path) -> MvCableProjectSnapshot:
    openpyxl = require_openpyxl()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active

    meta = _read_meta(ws)
    header_row = _find_header_row(ws)
    rows = _read_rows(ws, header_row)
    wb.close()

    if not rows:
        raise ValueError("В Excel нет строк с данными кабелей")

    project_name = meta.get(META_PROJECT, "").strip() or path.stem
    return MvCableProjectSnapshot(
        project_name=project_name,
        voltage_kv=_require_float(meta.get(META_U_N, ""), field=META_U_N),
        ru_type=_require_text(meta.get(META_RU_TYPE, ""), field=META_RU_TYPE),
        substation_name=_require_text(meta.get(META_SUBSTATION, ""), field=META_SUBSTATION),
        ikz3_ka=_require_float(meta.get(META_IKZ3, ""), field=META_IKZ3),
        t_breaker_s=_require_float(meta.get(META_T_BREAKER, ""), field=META_T_BREAKER),
        theta_0_c=_require_float(meta.get(META_Q0, ""), field=META_Q0),
        theta_okr_c=_require_float(meta.get(META_Q_OKR, ""), field=META_Q_OKR),
        rows=rows,
    )


def snapshot_to_inputs(snap: MvCableProjectSnapshot) -> list[MvCableFireInput]:
    return [snapshot_to_input(snap, row_index=i) for i in range(len(snap.rows))]


def snapshot_to_input(snap: MvCableProjectSnapshot, *, row_index: int = 0) -> MvCableFireInput:
    if row_index < 0 or row_index >= len(snap.rows):
        raise ValueError("Нет строки кабеля для расчёта")

    row = snap.rows[row_index]
    cable_label = row.get("cable_no", "").strip() or str(row_index + 1)
    try:
        return _row_to_input(snap, row)
    except ValueError as exc:
        raise ValueError(f"Кабель №{cable_label}: {exc}") from exc


def _parse_load_name(line_type: str) -> tuple[str, bool]:
    raw = line_type.strip()
    if not raw:
        raise ValueError("Не заполнено поле: Тип линии")
    upper = raw.upper()
    if upper in _STANDARD_LINE_TYPES:
        return upper, True
    return raw, False


def _resolve_i_nom_override(
    row: dict[str, str],
    *,
    is_standard_load: bool,
) -> float | None:
    override = parse_number(cell_str(row.get("i_nom_a", "")))
    if override is not None:
        return override
    if not is_standard_load:
        raise ValueError(
            "Для типа линии, отличного от ТСН и ДГР, необходимо заполнить Iн, А"
        )
    return None


def _parse_cable_installation(value: object) -> bool:
    text = cell_str(value).strip().upper().replace(".", "")
    if text in ("НОВ",):
        return False
    if text in ("СУЩ",):
        return True
    raise ValueError('Поле «Нов/Сущ.» должно быть «Нов» или «Сущ.»')


def _row_to_input(snap: MvCableProjectSnapshot, row: dict[str, str]) -> MvCableFireInput:
    load_name, is_standard_load = _parse_load_name(
        _require_text(row["line_type"], field="Тип линии")
    )
    voltage = snap.voltage_kv
    i_nom_override = _resolve_i_nom_override(row, is_standard_load=is_standard_load)
    power_kva = parse_number(cell_str(row.get("power_kva", "")))
    if i_nom_override is None and power_kva is None:
        raise ValueError("Укажите P, кВА или Iн, А")
    mark = normalize_cable_mark(_require_text(row["cable_mark"], field="Марка кабеля"))
    cores = _require_int(row["cores"], field="Количество жил")
    validate_cores_count(cores)
    kind = parse_conductor_material(
        _require_text(row["conductor_material"], field="Материал жил")
    )
    section = _require_float(row["section_mm2"], field="Сечение, мм²")
    designation = _cable_designation(mark, cores, section)
    is_existing_cable = _parse_cable_installation(
        _require_text(row["cable_installation"], field="Нов/Сущ.")
    )
    sheath_material_gen = _require_text(
        row["sheath_material_gen"], field="Материал оболочки, р.п."
    )
    cable_no = row.get("cable_no", "").strip() or "1"

    project_name = snap.project_name.strip()
    if not project_name:
        project_name = f"Выбор и проверка КЛ {int(voltage)} кВ до {load_name}"

    return MvCableFireInput(
        project_name=project_name,
        voltage_kv=voltage,
        tsn_power_kva=power_kva,
        i_nom_override_a=i_nom_override,
        is_standard_load=is_standard_load,
        ikz3_ka=snap.ikz3_ka,
        substation_name=snap.substation_name,
        cable_mark=mark,
        cores_count=cores,
        section_mm2=section,
        conductor_material=conductor_nominative(kind),
        sheath_material_gen=sheath_material_gen,
        i_dd_a=_require_float(row["i_dd_a"], field="Iдд, А"),
        theta_0_c=snap.theta_0_c,
        theta_dd_c=_require_float(row["theta_dd_c"], field="Qдд, °C"),
        theta_okr_c=snap.theta_okr_c,
        theta_limit_heating_c=_require_float(
            row["theta_limit_heating_c"], field="Tдоп нагрев, °C"
        ),
        theta_limit_fire_c=_require_float(
            row["theta_limit_fire_c"], field="Tдоп невозгорание, °C"
        ),
        b_const=_require_float(row["b_const"], field="b, мм2/кА2·с"),
        tae_s=_require_float(row["tae_s"], field="Та.эк., с"),
        t_main_protection_s=_require_float(
            row["t_main_s"], field="tоткл., с (основная защита)"
        ),
        t_backup_protection_s=_require_float(
            row["t_backup_s"], field="tоткл., с (резервная защита)"
        ),
        t_breaker_s=snap.t_breaker_s,
        source_zru=_source_zru_label(snap.ru_type, voltage),
        load_name=load_name,
        cable_no=cable_no,
        is_existing_cable=is_existing_cable,
        cable_route_text=(
            build_existing_cable_route_text(designation)
            if is_existing_cable
            else build_cable_route_text(
                load_name,
                voltage,
                designation,
                cores=cores,
                conductor_kind=kind,
                sheath_material_gen=sheath_material_gen,
                is_standard_load=is_standard_load,
            )
        ),
    )
