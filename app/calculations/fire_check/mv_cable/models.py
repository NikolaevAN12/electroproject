from __future__ import annotations

from dataclasses import dataclass, field

from .conductor_phrase import build_cable_route_text


def _default_cable_route_text() -> str:
    return build_cable_route_text(
        "ТСН",
        6.0,
        "АВВГнг(А)-LS 3х70",
        cores=3,
        conductor_kind="aluminum",
        sheath_material_gen="поливинилхлоридного пластиката",
    )


@dataclass(slots=True)
class MvCableFireInput:
    """Входные данные проверки КЛ свыше 1 кВ (из Excel)."""

    project_name: str = ""

    # Исходные данные
    voltage_kv: float = 6.0
    tsn_power_kva: float | None = 160.0
    i_nom_override_a: float | None = None
    is_standard_load: bool = True
    ikz3_ka: float = 10.176
    substation_name: str = "ПС 110 кВ Курумоч"

    cable_mark: str = "АВВГнг(А)-LS"
    cores_count: int = 3
    section_mm2: float = 70.0
    conductor_material: str = "алюминий"
    sheath_material_gen: str = "поливинилхлоридного пластиката"

    # Табличные / нормативные
    s_min_mm2: float = 70.0
    i_dd_a: float = 156.0

    # Температуры (°C)
    theta_0_c: float = 25.0
    theta_dd_c: float = 70.0
    theta_okr_c: float = 25.0
    theta_limit_heating_c: float = 160.0
    theta_limit_fire_c: float = 350.0

    # Материал жилы (циркуляр Ц-02-98(Э))
    b_const: float = 45.65
    a_const: float = 228.0

    # КЗ — нагрев (основная защита)
    tae_s: float = 0.04
    t_main_protection_s: float = 0.0
    t_breaker_s: float = 0.04

    # КЗ — невозгорание (резервная защита)
    t_backup_protection_s: float = 0.2

    # Текстовые подписи для отчёта
    source_zru: str = "ЗРУ 6 кВ"
    load_name: str = "ТСН"
    cable_no: str = "1"
    is_existing_cable: bool = False
    cable_route_text: str = field(default_factory=_default_cable_route_text)


@dataclass(slots=True)
class MvCableCheckVerdict:
    passed: bool
    left_text: str
    right_text: str
    limit_text: str


@dataclass(slots=True)
class MvCableFireResult:
    """Расчётные величины для подстановки в Word."""

    i_nom_a: float
    cable_designation: str

    # § минимальное сечение
    section_ok: bool

    # § длительный ток
    i_dd_ok: bool

    # § нагрев при КЗ
    theta_n_c: float
    b_heating_kas2: float
    k_heating: float
    theta_k_heating_c: float
    heating_ok: bool

    # § невозгорание
    t_off_fire_s: float
    b_fire_kas2: float
    k_fire: float
    theta_k_fire_c: float
    fire_ok: bool

    t_off_heating_s: float

    section_check: MvCableCheckVerdict
    current_check: MvCableCheckVerdict
    heating_check: MvCableCheckVerdict
    fire_check: MvCableCheckVerdict
