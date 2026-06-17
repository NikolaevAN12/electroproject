"""Материал и число жил кабеля — фразы для отчёта Word."""

from __future__ import annotations

from typing import Literal

ConductorKind = Literal["aluminum", "copper"]

_MAX_CORES = 5

_CORE_WITH_MATERIAL: dict[ConductorKind, dict[int, str]] = {
    "aluminum": {
        1: "с одной алюминиевой жилой",
        2: "с двумя алюминиевыми жилами",
        3: "с тремя алюминиевыми жилами",
        4: "с четырьмя алюминиевыми жилами",
        5: "с пятью алюминиевыми жилами",
    },
    "copper": {
        1: "с одной медной жилой",
        2: "с двумя медными жилами",
        3: "с тремя медными жилами",
        4: "с четырьмя медными жилами",
        5: "с пятью медными жилами",
    },
}

_INSTRUMENTAL_PLURAL: dict[ConductorKind, str] = {
    "aluminum": "алюминиевыми",
    "copper": "медными",
}

_GENITIVE: dict[ConductorKind, str] = {
    "aluminum": "алюминия",
    "copper": "меди",
}

_NOMINATIVE: dict[ConductorKind, str] = {
    "aluminum": "алюминий",
    "copper": "медь",
}


def validate_cores_count(cores: int) -> None:
    if cores < 1 or cores > _MAX_CORES:
        raise ValueError(
            f"Количество жил должно быть от 1 до {_MAX_CORES}, указано: {cores}"
        )


def normalize_cable_mark(cable_mark: str) -> str:
    """Нормализация марки кабеля: как в Excel (без изменения регистра)."""
    return cable_mark.strip()


def parse_conductor_material(value: str) -> ConductorKind:
    """Материал жил из Excel: «медь» или «алюминий»."""
    text = value.strip().lower()
    if text == "алюминий":
        return "aluminum"
    if text == "медь":
        return "copper"
    raise ValueError(
        f"Материал жил должен быть «медь» или «алюминий», указано: {value}"
    )


def conductor_kind_from_nominative(value: str) -> ConductorKind:
    return parse_conductor_material(value)


def cores_with_material_phrase(cores: int, kind: ConductorKind) -> str:
    validate_cores_count(cores)
    return _CORE_WITH_MATERIAL[kind][cores]


def conductor_instrumental_plural(kind: ConductorKind) -> str:
    return _INSTRUMENTAL_PLURAL[kind]


def conductor_genitive(kind: ConductorKind) -> str:
    return _GENITIVE[kind]


def conductor_nominative(kind: ConductorKind) -> str:
    return _NOMINATIVE[kind]


def build_existing_cable_route_text(designation: str) -> str:
    return f"Существующий кабель марки {designation}."


def build_cable_route_text(
    load_name: str,
    voltage_kv: float,
    designation: str,
    *,
    cores: int,
    conductor_kind: ConductorKind,
    sheath_material_gen: str,
    is_standard_load: bool = True,
) -> str:
    cores_phrase = cores_with_material_phrase(cores, conductor_kind)
    material = sheath_material_gen.strip()
    if not material:
        raise ValueError("Не заполнен материал оболочки (родительный падеж)")
    if is_standard_load:
        intro = f"Подключение {load_name} будет выполнено новыми кабелями {int(voltage_kv)} кВ "
    else:
        intro = f"Подключение будет выполнено новыми кабелями {int(voltage_kv)} кВ "
    return (
        f"{intro}"
        f"с оболочкой из {material} {cores_phrase} марки: "
        f"{designation}."
    )
