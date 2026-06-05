from __future__ import annotations

from dataclasses import dataclass, field

# (ключ, подпись в интерфейсе и Excel)
EQUIPMENT_TYPES: tuple[tuple[str, str], ...] = (
    ("fuse", "Предохранитель"),
    ("breaker", "Автоматический выключатель"),
    ("cable", "Кабель"),
    ("kz_point", "Точка КЗ"),
)

EQUIPMENT_LABEL_BY_KEY: dict[str, str] = dict(EQUIPMENT_TYPES)
EQUIPMENT_KEY_BY_LABEL: dict[str, str] = {label: key for key, label in EQUIPMENT_TYPES}

BREAKER_CURVE_CHOICES: tuple[str, ...] = ("B", "C", "D", "K", "Z")
BREAKER_CURVE_LABELS: dict[str, str] = {key: key for key in BREAKER_CURVE_CHOICES}
BREAKER_CURVE_BY_LABEL: dict[str, str] = dict(BREAKER_CURVE_LABELS)


@dataclass(slots=True)
class SoptEquipmentItem:
    item_type: str
    designation: str
    rated_current_a: float = 0.0
    resistance_ohm: float = 0.0
    cable_length_m: float = 0.0
    cable_gamma: float = 0.0
    cable_section_mm2: float = 0.0
    # Параметры автомата для карт селективности (как CB_Curve / CB_Multiplier / CB_Time).
    cb_curve: str = ""
    cb_multiplier: float = 0.0
    cb_trip_time_s: float = 0.0
    # Точка КЗ участвует в проверке селективности (раздел 4 Word); на расчёт §3 не влияет.
    include_in_selectivity: bool = False


@dataclass(slots=True)
class SoptSubsection:
    name: str
    items: list[SoptEquipmentItem] = field(default_factory=list)


@dataclass(slots=True)
class SoptSection:
    name: str
    subsections: list[SoptSubsection] = field(default_factory=list)


@dataclass(slots=True)
class SoptInput:
    project_name: str = ""
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


@dataclass(slots=True)
class SoptResult:
    message: str = ""
    r_ab: float = 0.0
    r_per: float = 0.0
    r_kz: float = 0.0
    r_gr: float = 0.0
    n_required: float = 0.0
    n_accepted: int = 0
    selective_ok: bool = False
