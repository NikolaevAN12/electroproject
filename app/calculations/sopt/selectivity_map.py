"""Карты селективности автоматических выключателей для расчёта СОПТ."""

from __future__ import annotations

import io
from dataclasses import dataclass

import numpy as np

from app.shared.user_errors import MSG_MISSING_MATPLOTLIB

from .models import SoptEquipmentItem
from .resistance import item_resistance_ohm

KC_DEFAULT = 0.6
_CURVE_MULTIPLIERS = {"B": 5, "C": 10, "D": 20, "K": 14, "Z": 3}


@dataclass(slots=True)
class SelectivityMapResult:
    block_title: str
    upstream_label: str
    downstream_label: str
    i_max_a: float
    i_min_a: float
    overlap: bool
    image: io.BytesIO


def _kz_for_selectivity(item: SoptEquipmentItem) -> bool:
    return item.item_type == "kz_point" and item.include_in_selectivity


def _furthest_kz_point_in_zone(items: list[SoptEquipmentItem], breaker_index: int) -> str:
    """Крайняя (последняя по цепи) отмеченная точка КЗ в зоне нижнего автомата."""
    last_name = ""
    for idx in range(breaker_index + 1, len(items)):
        if items[idx].item_type == "breaker":
            break
        if _kz_for_selectivity(items[idx]):
            designation = items[idx].designation.strip()
            if designation:
                last_name = designation
    if last_name:
        return last_name
    for item in reversed(items):
        if _kz_for_selectivity(item):
            designation = item.designation.strip()
            if designation:
                return designation
    return "K1"


def _kz_point_sc_currents(
    items: list[SoptEquipmentItem],
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
    kc: float = KC_DEFAULT,
) -> tuple[float, float]:
    """Максимальный Iкз и минимальный Iкзд по отмеченным точкам КЗ подраздела."""
    running_components: list[float] = []
    i_kz_values: list[float] = []
    i_kzd_values: list[float] = []
    has_any_kz = any(item.item_type == "kz_point" for item in items)

    def append_point(chain_r: float) -> None:
        r_kz = r_ab + r_per + 2 * chain_r
        if r_kz <= 0:
            return
        e_calc = 1.7 if r_kz < r_gr else 1.93
        i_kz = e_calc * n_total_elements / r_kz
        i_kzd = kc * i_kz
        if i_kz > 0:
            i_kz_values.append(i_kz)
        if i_kzd > 0:
            i_kzd_values.append(i_kzd)

    for item in items:
        if item.item_type == "kz_point":
            if item.include_in_selectivity:
                append_point(sum(running_components))
            continue
        comp_r = item_resistance_ohm(item)
        if comp_r > 0:
            running_components.append(comp_r)

    if not i_kz_values and not has_any_kz:
        append_point(sum(running_components))

    if not i_kz_values or not i_kzd_values:
        return 0.0, 0.0
    return max(i_kz_values), min(i_kzd_values)


def _breaker_indices(items: list[SoptEquipmentItem]) -> list[int]:
    return [idx for idx, item in enumerate(items) if item.item_type == "breaker"]


def _breaker_label(item: SoptEquipmentItem) -> str:
    return item.designation.strip() or f"{int(item.rated_current_a)} А"


def _selectivity_breaker_pairs(items: list[SoptEquipmentItem]) -> list[tuple[int, int]]:
    breaker_idxs = _breaker_indices(items)
    if len(breaker_idxs) < 2:
        return []
    return list(zip(breaker_idxs, breaker_idxs[1:], strict=False))


def _curve_params_from_item(item: SoptEquipmentItem) -> tuple[str, float, float]:
    curve = (item.cb_curve or "C").strip().upper()
    return curve, item.cb_multiplier, item.cb_trip_time_s


def get_curve_points(
    item: SoptEquipmentItem,
    *,
    curve: str | None = None,
    multiplier: float | None = None,
    trip_time_s: float | None = None,
) -> tuple[np.ndarray, np.ndarray]:
    if curve is None or multiplier is None or trip_time_s is None:
        item_curve, item_mult, item_time = _curve_params_from_item(item)
        curve = item_curve if curve is None else curve
        multiplier = item_mult if multiplier is None else multiplier
        trip_time_s = item_time if trip_time_s is None else trip_time_s
    i_nom = item.rated_current_a if item.rated_current_a > 0 else 16.0
    curve_text = (curve or "C").strip().upper().replace("NAN", "")
    if curve_text in {"", "NAN"}:
        i_trip = i_nom * (multiplier if multiplier > 0 else 1.0)
        x = np.array([i_trip, i_trip * 10.0])
        y = np.array([trip_time_s if trip_time_s > 0 else 0.1] * 2)
        return x, y

    k_mult_val = multiplier
    t_set = trip_time_s
    if k_mult_val > 0:
        mag_mult = k_mult_val
    else:
        mag_mult = _CURVE_MULTIPLIERS.get(curve_text, 10.0)

    t100_start = 100.0
    if curve_text == "B":
        i_start, t_end = 1.6 * i_nom, 4.15
    elif curve_text == "C":
        i_start, t_end = 1.6 * i_nom, 1.0
    elif curve_text == "D":
        i_start, t_end = 1.13 * i_nom, 0.3
    else:
        i_start, t_end = 1.13 * i_nom, 0.5

    t_em = t_set if t_set > 0 else 0.015
    x = np.logspace(-1, 5, 400) * i_nom
    y: list[float] = []
    i_start_ratio = i_start / i_nom
    for val in x:
        i_rel = val / i_nom
        if val < i_start:
            t = 10000.0
        elif val < (i_nom * mag_mult):
            denom = mag_mult - i_start_ratio
            ratio = (i_rel - i_start_ratio) / denom if denom > 0 else 0.0
            t = t100_start * ((t_end / t100_start) ** ratio)
        else:
            t = t_em
        y.append(max(0.001, t))
    return x, np.array(y)


def create_selectivity_map(
    upstream: SoptEquipmentItem,
    downstream: SoptEquipmentItem,
    i_min_a: float,
    i_max_a: float,
) -> tuple[io.BytesIO, bool]:
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise RuntimeError(
            MSG_MISSING_MATPLOTLIB
        ) from exc

    up_x, up_y = get_curve_points(upstream)
    down_x, down_y = get_curve_points(downstream)
    up_label = _breaker_label(upstream)
    down_label = _breaker_label(downstream)

    fig, ax = plt.subplots(figsize=(7, 6))
    ax.loglog(down_x, down_y, "b-", label=f"Нижний: {down_label}")
    ax.loglog(up_x, up_y, "r-", label=f"Верхний: {up_label}")

    if i_min_a > 0:
        ax.axvline(i_min_a, color="orange", linestyle="--", label=f"Iкз.мин={i_min_a:.1f} А")
    if i_max_a > 0:
        ax.axvline(i_max_a, color="green", linestyle="--", label=f"Iкз.макс={i_max_a:.1f} А")

    ax.set_xlabel("Ток (А)")
    ax.set_ylabel("Время (с)")
    ax.grid(True, which="both", ls="--")
    ax.legend(loc="upper right")

    overlap = False
    if i_min_a > 0 and i_max_a > 0 and i_max_a >= i_min_a:
        lo, hi = i_min_a, i_max_a
        mask = (down_x >= lo) & (down_x <= hi)
        if np.any(mask):
            up_interp = np.interp(down_x[mask], up_x, up_y)
            if np.any(down_y[mask] >= up_interp):
                overlap = True

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf, overlap


def build_subsection_selectivity_maps(
    items: list[SoptEquipmentItem],
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
    kc: float = KC_DEFAULT,
) -> list[SelectivityMapResult]:
    pairs = _selectivity_breaker_pairs(items)
    if not pairs:
        return []

    i_max_a, i_min_a = _kz_point_sc_currents(
        items,
        r_ab=r_ab,
        r_per=r_per,
        r_gr=r_gr,
        n_total_elements=n_total_elements,
        kc=kc,
    )
    if i_max_a <= 0 or i_min_a <= 0:
        return []

    results: list[SelectivityMapResult] = []
    for up_idx, down_idx in pairs:
        upstream = items[up_idx]
        downstream = items[down_idx]
        image, overlap = create_selectivity_map(upstream, downstream, i_min_a, i_max_a)
        results.append(
            SelectivityMapResult(
                block_title=_furthest_kz_point_in_zone(items, down_idx),
                upstream_label=_breaker_label(upstream),
                downstream_label=_breaker_label(downstream),
                i_max_a=i_max_a,
                i_min_a=i_min_a,
                overlap=overlap,
                image=image,
            )
        )
    return results
