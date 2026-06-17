"""Карты селективности автоматических выключателей для расчёта СОПТ."""

from __future__ import annotations

import io
from dataclasses import dataclass

import numpy as np

from app.shared.user_errors import MSG_MISSING_MATPLOTLIB

from .kc import kc_from_r_kz
from .models import SoptEquipmentItem
from .resistance import item_resistance_ohm
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


def _breaker_label(item: SoptEquipmentItem) -> str:
    return item.designation.strip() or f"{int(item.rated_current_a)} А"


def _chain_r_at_index(items: list[SoptEquipmentItem], kz_idx: int) -> float:
    chain_r = 0.0
    for idx in range(kz_idx):
        item = items[idx]
        if item.item_type == "kz_point":
            continue
        comp_r = item_resistance_ohm(item)
        if comp_r > 0:
            chain_r += comp_r
    return chain_r


def _sc_currents_at_chain_r(
    chain_r: float,
    *,
    r_ab: float,
    r_per: float,
    r_gr: float,
    n_total_elements: int,
) -> tuple[float, float]:
    r_kz = r_ab + r_per + 2 * chain_r
    if r_kz <= 0:
        return 0.0, 0.0
    e_calc = 1.7 if r_kz < r_gr else 1.93
    i_kz = e_calc * n_total_elements / r_kz
    i_kzd = kc_from_r_kz(r_kz) * i_kz
    if i_kz <= 0 or i_kzd <= 0:
        return 0.0, 0.0
    return i_kz, i_kzd


def _breaker_pair_before_kz(
    items: list[SoptEquipmentItem],
    kz_idx: int,
    *,
    section_head_breaker: SoptEquipmentItem | None = None,
) -> tuple[SoptEquipmentItem, SoptEquipmentItem] | None:
    """Верхний и нижний автомат на пути к отмеченной точке КЗ (два ближайших к точке).

    Если в подразделе только один автомат перед точкой КЗ, верхним берётся
    вводной автомат раздела (2QF4, 1QF1 — не веточный SF*).
    """
    downstream_idx: int | None = None
    upstream_idx: int | None = None
    for idx in range(kz_idx - 1, -1, -1):
        if items[idx].item_type != "breaker":
            continue
        if downstream_idx is None:
            downstream_idx = idx
            continue
        upstream_idx = idx
        break

    if downstream_idx is None:
        return None

    downstream = items[downstream_idx]
    if upstream_idx is not None:
        return items[upstream_idx], downstream

    if section_head_breaker is not None:
        head_label = _breaker_label(section_head_breaker).casefold()
        down_label = _breaker_label(downstream).casefold()
        if head_label and head_label != down_label:
            return section_head_breaker, downstream

    return None


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
    ax.set_xlim(left=max(1, down_x[0]), right=max(i_max_a * 5, 10000))
    ax.set_ylim(0.001, 1000)
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
    section_head_breaker: SoptEquipmentItem | None = None,
) -> list[SelectivityMapResult]:
    """По одной карте селективности на каждую отмеченную точку КЗ подраздела."""
    results: list[SelectivityMapResult] = []
    for kz_idx, item in enumerate(items):
        if not _kz_for_selectivity(item):
            continue
        pair = _breaker_pair_before_kz(
            items,
            kz_idx,
            section_head_breaker=section_head_breaker,
        )
        if pair is None:
            continue
        upstream, downstream = pair
        i_max_a, i_min_a = _sc_currents_at_chain_r(
            _chain_r_at_index(items, kz_idx),
            r_ab=r_ab,
            r_per=r_per,
            r_gr=r_gr,
            n_total_elements=n_total_elements,
        )
        if i_max_a <= 0 or i_min_a <= 0:
            continue
        image, overlap = create_selectivity_map(upstream, downstream, i_min_a, i_max_a)
        block_title = item.designation.strip() or f"K{kz_idx + 1}"
        results.append(
            SelectivityMapResult(
                block_title=block_title,
                upstream_label=_breaker_label(upstream),
                downstream_label=_breaker_label(downstream),
                i_max_a=i_max_a,
                i_min_a=i_min_a,
                overlap=overlap,
                image=image,
            )
        )
    return results
