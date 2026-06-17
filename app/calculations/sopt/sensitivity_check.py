"""Проверка автоматических выключателей на чувствительность для расчёта СОПТ."""

from __future__ import annotations

from dataclasses import dataclass, replace

from .models import SoptEquipmentItem, SoptInput, SoptResult
from .resistance import item_resistance_ohm

SENSITIVITY_THRESHOLD = 1.5
_SENSITIVITY_CURVE_MULTIPLIERS = {"B": 7, "C": 15, "D": 20, "K": 30, "Z": 8}


@dataclass(slots=True)
class BreakerSensitivityRow:
    number: int
    designation: str
    rated_current_a: float
    cb_curve: str
    cb_multiplier: float
    i_kz_a: float
    k_sensitivity: float
    passes: bool


def _curve_label(item: SoptEquipmentItem) -> str:
    curve = (item.cb_curve or "C").strip().upper().replace("NAN", "")
    return curve if curve else "C"


def _effective_multiplier(item: SoptEquipmentItem) -> float:
    if item.cb_multiplier > 0:
        return item.cb_multiplier
    curve = _curve_label(item)
    return _SENSITIVITY_CURVE_MULTIPLIERS.get(curve, 15.0)


def _merge_breaker_items(left: SoptEquipmentItem, right: SoptEquipmentItem) -> SoptEquipmentItem:
    """Объединяет параметры одного автомата из разных цепочек (приоритет заполненным полям)."""
    multiplier = left.cb_multiplier if left.cb_multiplier > 0 else right.cb_multiplier
    curve = left.cb_curve.strip() if left.cb_curve.strip() else right.cb_curve
    return replace(left, cb_multiplier=multiplier, cb_curve=curve)


def _ikz_from_chain_r(
    chain_r: float,
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
) -> float:
    r_kz = r_ab + r_per + 2 * chain_r
    if r_kz <= 0:
        return 0.0
    e_calc = 1.7 if r_kz < r_gr else 1.93
    return e_calc * n_total_elements / r_kz


def _chain_r_at_first_kz_after_breaker(
    items: list[SoptEquipmentItem],
    breaker_idx: int,
) -> float | None:
    """Сопротивление цепи до первой точки КЗ в зоне автомата (до следующего АВ)."""
    chain_r = 0.0
    for idx, item in enumerate(items):
        if item.item_type == "kz_point":
            if idx > breaker_idx:
                return chain_r
            continue
        if idx > breaker_idx and item.item_type == "breaker":
            return chain_r
        if item.item_type != "kz_point":
            chain_r += item_resistance_ohm(item)
    if breaker_idx >= len(items) - 1:
        return None
    return chain_r


def _ikz_for_breaker_in_subsection(
    items: list[SoptEquipmentItem],
    breaker_idx: int,
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
) -> float:
    chain_r = _chain_r_at_first_kz_after_breaker(items, breaker_idx)
    if chain_r is None:
        return 0.0
    return _ikz_from_chain_r(
        chain_r,
        r_ab=r_ab,
        r_per=r_per,
        r_gr=r_gr,
        n_total_elements=n_total_elements,
    )


def _sensitivity_coefficient(item: SoptEquipmentItem, i_kz_a: float) -> float:
    i_nom = item.rated_current_a
    k_mult = _effective_multiplier(item)
    if i_nom <= 0 or k_mult <= 0 or i_kz_a <= 0:
        return 0.0
    return i_kz_a / (k_mult * i_nom)


def _breaker_passes_sensitivity(item: SoptEquipmentItem, i_kz_a: float) -> bool:
    return _sensitivity_coefficient(item, i_kz_a) > SENSITIVITY_THRESHOLD


def _breaker_merge_key(
    item: SoptEquipmentItem,
    *,
    section_idx: int,
    subsection_idx: int,
    item_idx: int,
) -> str:
    designation = item.designation.strip()
    if designation:
        return designation
    return f"__{section_idx}.{subsection_idx}.{item_idx}"


def build_breaker_sensitivity_rows(
    input_model: SoptInput,
    result_model: SoptResult,
    *,
    n_total_elements: int,
) -> list[BreakerSensitivityRow]:
    """Сводная таблица чувствительности по уникальным автоматам (мин. Iкз по цепям)."""
    merged: dict[str, tuple[SoptEquipmentItem, float]] = {}

    for section_idx, section in enumerate(input_model.sections, start=1):
        for subsection_idx, subsection in enumerate(section.subsections, start=1):
            items = subsection.items
            for item_idx, item in enumerate(items):
                if item.item_type != "breaker":
                    continue
                i_kz = _ikz_for_breaker_in_subsection(
                    items,
                    item_idx,
                    r_ab=result_model.r_ab,
                    r_per=result_model.r_per,
                    r_gr=result_model.r_gr,
                    n_total_elements=n_total_elements,
                )
                key = _breaker_merge_key(
                    item,
                    section_idx=section_idx,
                    subsection_idx=subsection_idx,
                    item_idx=item_idx,
                )
                if key not in merged:
                    merged[key] = (item, i_kz)
                else:
                    prev_item, prev_i_kz = merged[key]
                    merged[key] = (
                        _merge_breaker_items(prev_item, item),
                        min(prev_i_kz, i_kz),
                    )

    sorted_keys = sorted(
        merged.keys(),
        key=lambda k: (
            merged[k][0].designation.strip() or k
        ).casefold(),
    )

    rows: list[BreakerSensitivityRow] = []
    for number, key in enumerate(sorted_keys, start=1):
        item, i_kz = merged[key]
        k_mult = _effective_multiplier(item)
        rows.append(
            BreakerSensitivityRow(
                number=number,
                designation=item.designation.strip() or f"{int(item.rated_current_a)} А",
                rated_current_a=item.rated_current_a,
                cb_curve=_curve_label(item),
                cb_multiplier=k_mult,
                i_kz_a=i_kz,
                k_sensitivity=_sensitivity_coefficient(item, i_kz),
                passes=_breaker_passes_sensitivity(item, i_kz),
            )
        )
    return rows
