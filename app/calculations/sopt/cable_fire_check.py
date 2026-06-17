"""Проверка кабелей на невозгорание для расчёта СОПТ."""

from __future__ import annotations

from dataclasses import dataclass

from app.calculations.fire_check.calculator import FireCheckCalculator

from .models import SoptEquipmentItem, SoptInput, SoptResult
from .resistance import item_resistance_ohm

FIRE_PASS_MAX_TEMP_C = 350.0
_FIRE_FALLBACK_CHECK_DISTANCE_M = 20.0
_DEFAULT_BREAKER_TRIP_TIME_S = 0.1
_RESERVE_TRIP_MARGIN_S = 0.1
_THETA_O_C = 25.0
_THETA_DD_C = 70.0
_THETA_OKR_C = 25.0

# Длительно допустимый ток Iдд для ВВГЭнг(А)-LS 2х по сечению, А.
_IDD_BY_SECTION_MM2: dict[float, float] = {
    2.5: 27.0,
    4.0: 36.0,
    6.0: 46.0,
    10.0: 63.0,
    16.0: 84.0,
    25.0: 112.0,
}

_FIRE_CALC = FireCheckCalculator()


@dataclass(slots=True)
class _CableFireEval:
    i_kz: float
    i_nom: float
    i_dd: float
    t_off: float
    theta_before: int | None
    theta_after: float | None
    passes: bool


@dataclass(slots=True)
class CableFireCheckRow:
    number: int
    cable_mark: str
    i_nom_a: float
    i_dd_a: float
    theta_before_c: int | None
    i_kz_ka: float
    shutdown_time_s: float
    theta_after_c: float | None
    passes: bool


def _section_key(section_mm2: float) -> float:
    return round(section_mm2, 4)


def _idd_for_section(section_mm2: float) -> float:
    key = _section_key(section_mm2)
    if key in _IDD_BY_SECTION_MM2:
        return _IDD_BY_SECTION_MM2[key]
    for known, value in _IDD_BY_SECTION_MM2.items():
        if abs(known - section_mm2) < 1e-6:
            return value
    return 0.0


def _breaker_trip_time_s(item: SoptEquipmentItem) -> float:
    if item.cb_trip_time_s > 0:
        return item.cb_trip_time_s
    return _DEFAULT_BREAKER_TRIP_TIME_S


def _breakers_before(items: list[SoptEquipmentItem], item_idx: int) -> list[SoptEquipmentItem]:
    return [item for idx, item in enumerate(items) if idx < item_idx and item.item_type == "breaker"]


def _max_breaker_trip_time_s(input_model: SoptInput) -> float:
    max_t = 0.0
    for section in input_model.sections:
        for subsection in section.subsections:
            for item in subsection.items:
                if item.item_type == "breaker":
                    max_t = max(max_t, _breaker_trip_time_s(item))
    return max_t if max_t > 0 else _DEFAULT_BREAKER_TRIP_TIME_S


def _reserve_shutdown_time_s(
    items: list[SoptEquipmentItem],
    cable_idx: int,
    *,
    max_breaker_trip_s: float,
) -> float:
    breakers = _breakers_before(items, cable_idx)
    if len(breakers) >= 2:
        return _breaker_trip_time_s(breakers[-2])
    return max_breaker_trip_s + _RESERVE_TRIP_MARGIN_S


def _upstream_breaker_nominal_a(items: list[SoptEquipmentItem], cable_idx: int) -> float:
    """Iраб — по номиналу непосредственно вышестоящего автомата в цепочке."""
    breakers = _breakers_before(items, cable_idx)
    if breakers:
        return breakers[-1].rated_current_a
    return 0.0


def _cable_check_distance_m(cable: SoptEquipmentItem) -> float:
    """Расстояние от начала кабеля для повторной проверки: 20 м или конец кабеля."""
    length = cable.cable_length_m
    if length <= 0:
        return 0.0
    return min(_FIRE_FALLBACK_CHECK_DISTANCE_M, length)


def _cable_resistance_to_distance_ohm(cable: SoptEquipmentItem, distance_m: float) -> float:
    if distance_m <= 0:
        return 0.0
    full_r = item_resistance_ohm(cable)
    if full_r <= 0 or cable.cable_length_m <= 0:
        return 0.0
    return full_r * (distance_m / cable.cable_length_m)


def _ikz_at_cable_point(
    items: list[SoptEquipmentItem],
    cable_idx: int,
    distance_into_cable_m: float,
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
) -> float:
    """Ток КЗ в точке на расстоянии distance_into_cable_m от начала кабеля."""
    chain_r = 0.0
    for idx, item in enumerate(items):
        if idx > cable_idx:
            break
        if idx == cable_idx:
            chain_r += _cable_resistance_to_distance_ohm(item, distance_into_cable_m)
            continue
        if item.item_type != "kz_point":
            chain_r += item_resistance_ohm(item)
    r_kz = r_ab + r_per + 2 * chain_r
    if r_kz <= 0:
        return 0.0
    e_calc = 1.7 if r_kz < r_gr else 1.93
    return e_calc * n_total_elements / r_kz


def _parse_temperature_c(display: str | None) -> float | None:
    if not display:
        return None
    try:
        return float(display.replace(",", "."))
    except ValueError:
        return None


def _evaluate_cable_fire(
    *,
    section_mm2: float,
    i_nom_a: float,
    i_dd_a: float,
    i_kz_a: float,
    shutdown_time_s: float,
) -> tuple[int | None, float | None, bool]:
    if section_mm2 <= 0 or i_dd_a <= 0 or i_nom_a <= 0 or i_kz_a <= 0:
        return None, None, False
    theta_before = _FIRE_CALC._wire_temperature(
        i_dd_a,
        i_nom_a,
        _THETA_O_C,
        _THETA_DD_C,
        _THETA_OKR_C,
    )
    i_kz_ka = i_kz_a / 1000.0
    bk = _FIRE_CALC._thermal_impulse_k2s(i_kz_ka, shutdown_time_s, 0.0)
    theta_after_display = _FIRE_CALC._fire_formula_result(theta_before, bk, section_mm2)
    theta_after = _parse_temperature_c(theta_after_display)
    passes = theta_after is not None and theta_after < FIRE_PASS_MAX_TEMP_C
    return theta_before, theta_after, passes


def _evaluate_cable_fire_with_fallback(
    items: list[SoptEquipmentItem],
    cable_idx: int,
    cable: SoptEquipmentItem,
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
    i_nom: float,
    i_dd: float,
    t_off: float,
) -> _CableFireEval:
    """Проверка в начале кабеля; при «Нет» — повтор на 20 м или в конце кабеля."""
    sc_kw = dict(
        r_ab=r_ab,
        r_per=r_per,
        r_gr=r_gr,
        n_total_elements=n_total_elements,
    )
    section_mm2 = cable.cable_section_mm2

    i_kz_start = _ikz_at_cable_point(items, cable_idx, 0.0, **sc_kw)
    theta_before, theta_after, passes = _evaluate_cable_fire(
        section_mm2=section_mm2,
        i_nom_a=i_nom,
        i_dd_a=i_dd,
        i_kz_a=i_kz_start,
        shutdown_time_s=t_off,
    )
    if passes:
        return _CableFireEval(
            i_kz=i_kz_start,
            i_nom=i_nom,
            i_dd=i_dd,
            t_off=t_off,
            theta_before=theta_before,
            theta_after=theta_after,
            passes=True,
        )

    check_distance_m = _cable_check_distance_m(cable)
    if check_distance_m <= 0:
        return _CableFireEval(
            i_kz=i_kz_start,
            i_nom=i_nom,
            i_dd=i_dd,
            t_off=t_off,
            theta_before=theta_before,
            theta_after=theta_after,
            passes=False,
        )

    i_kz_fallback = _ikz_at_cable_point(items, cable_idx, check_distance_m, **sc_kw)
    theta_before_fb, theta_after_fb, passes_fb = _evaluate_cable_fire(
        section_mm2=section_mm2,
        i_nom_a=i_nom,
        i_dd_a=i_dd,
        i_kz_a=i_kz_fallback,
        shutdown_time_s=t_off,
    )
    return _CableFireEval(
        i_kz=i_kz_fallback,
        i_nom=i_nom,
        i_dd=i_dd,
        t_off=t_off,
        theta_before=theta_before_fb,
        theta_after=theta_after_fb,
        passes=passes_fb,
    )


def build_cable_fire_check_rows(
    input_model: SoptInput,
    result_model: SoptResult,
    *,
    n_total_elements: int,
) -> list[CableFireCheckRow]:
    by_section: dict[float, list[_CableFireEval]] = {}
    max_breaker_trip_s = _max_breaker_trip_time_s(input_model)

    for section in input_model.sections:
        for subsection in section.subsections:
            items = subsection.items
            for cable_idx, item in enumerate(items):
                if item.item_type != "cable":
                    continue
                section_mm2 = item.cable_section_mm2
                if section_mm2 <= 0:
                    continue
                key = _section_key(section_mm2)
                i_nom = _upstream_breaker_nominal_a(items, cable_idx)
                i_dd = _idd_for_section(section_mm2)
                t_off = _reserve_shutdown_time_s(items, cable_idx, max_breaker_trip_s=max_breaker_trip_s)
                evaluation = _evaluate_cable_fire_with_fallback(
                    items,
                    cable_idx,
                    item,
                    r_ab=result_model.r_ab,
                    r_per=result_model.r_per,
                    r_gr=result_model.r_gr,
                    n_total_elements=n_total_elements,
                    i_nom=i_nom,
                    i_dd=i_dd,
                    t_off=t_off,
                )
                by_section.setdefault(key, []).append(evaluation)

    sorted_sections = sorted(by_section.keys())
    rows: list[CableFireCheckRow] = []
    for number, section_mm2 in enumerate(sorted_sections, start=1):
        evaluations = by_section[section_mm2]
        passes = all(item.passes for item in evaluations)
        if passes:
            representative = max(evaluations, key=lambda item: item.i_kz)
        else:
            failed = [item for item in evaluations if not item.passes]
            representative = max(
                failed,
                key=lambda item: item.theta_after if item.theta_after is not None else 0.0,
            )
        i_kz = representative.i_kz
        i_nom = representative.i_nom
        i_dd = representative.i_dd
        t_off = representative.t_off
        theta_before = representative.theta_before
        theta_after = representative.theta_after
        section_text = (
            str(int(section_mm2))
            if abs(section_mm2 - round(section_mm2)) < 1e-9
            else f"{section_mm2:g}"
        )
        rows.append(
            CableFireCheckRow(
                number=number,
                cable_mark=f"ВВГЭнг(А)-LS 2х{section_text}",
                i_nom_a=i_nom,
                i_dd_a=i_dd,
                theta_before_c=theta_before,
                i_kz_ka=i_kz / 1000.0,
                shutdown_time_s=t_off,
                theta_after_c=theta_after,
                passes=passes,
            )
        )
    return rows
