"""Схемы замещения и карты селективности (matplotlib)."""

from __future__ import annotations

import io
import math
from typing import Any

import numpy as np

from app.shared.user_errors import MSG_MISSING_MATPLOTLIB

from .models import CB_CURVE_MULTIPLIERS, ScPointResult


def _require_plt():
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as patches
    except ImportError as exc:
        raise RuntimeError(MSG_MISSING_MATPLOTLIB) from exc
    return plt, patches


def _is_valid_number(value: float) -> bool:
    return not (math.isnan(value) or math.isinf(value))


def get_curve_points(node: dict[str, Any]) -> tuple[np.ndarray, np.ndarray]:
    i_nom = node.get("CB_Nominal", 16.0)
    if not _is_valid_number(i_nom) or i_nom <= 0:
        i_nom = 16.0
    curve = str(node.get("CB_Curve", "")).upper().strip()
    k_mult_val = node.get("CB_Multiplier", float("nan"))
    t_set = node.get("CB_Time", 0.1)
    x = np.logspace(-1, 5, 400) * i_nom

    if _is_valid_number(k_mult_val) and k_mult_val > 0:
        mag_mult = k_mult_val
    else:
        mag_mult = CB_CURVE_MULTIPLIERS.get(curve, 10)

    t100_start = 100.0
    if curve == "B":
        i_start, t_end = 1.6 * i_nom, 4.15
    elif curve == "C":
        i_start, t_end = 1.6 * i_nom, 1.0
    elif curve == "D":
        i_start, t_end = 1.13 * i_nom, 0.3
    else:
        i_start, t_end = 1.13 * i_nom, 0.5

    t_em = t_set if (_is_valid_number(t_set) and t_set > 0) else 0.015
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


def create_replacement_diagram(
    name: str,
    result: ScPointResult,
    seq_type: str = "1",
) -> io.BytesIO:
    plt, patches = _require_plt()
    valid_nodes = [node for node in result.path if any(e.name == node for e in result.elements_data)]
    step_x = 2.5
    fig_width = max(8, (len(valid_nodes) + 2) * step_x * 0.7)
    fig, ax = plt.subplots(figsize=(fig_width, 3.5))
    ax.set_aspect("equal")
    ax.axis("off")
    curr_x = 0.5
    if seq_type == "1":
        ax.add_patch(patches.Circle((curr_x, 0), 0.3, fill=False, color="black", linewidth=1.5))
        ax.text(curr_x, 0, "~", ha="center", va="center", fontsize=18)
        ax.text(curr_x, -1.0, "Источник", ha="center", fontsize=9, weight="bold")
        start_conn_x = curr_x + 0.3
    else:
        ax.plot([curr_x - 0.4, curr_x + 0.4], [0, 0], "k-", linewidth=2.5)
        ax.text(curr_x, -1.0, "Нейтраль", ha="center", fontsize=9, weight="bold")
        start_conn_x = curr_x
    prev_conn_x = start_conn_x
    for node_name in valid_nodes:
        comp_center_x = prev_conn_x + 1.2
        ax.plot([prev_conn_x, comp_center_x - 0.5], [0, 0], "k-", linewidth=1.5)
        ax.add_patch(
            patches.Rectangle(
                (comp_center_x - 0.5, -0.25), 1.0, 0.5, fill=False, color="black", linewidth=1.5
            )
        )
        node = next(e for e in result.elements_data if e.name == node_name)
        r = getattr(node, f"r{seq_type}")
        x_val = getattr(node, f"x{seq_type}")
        ax.text(comp_center_x, 0.4, node_name, ha="center", fontsize=9, weight="bold", color="blue")
        ax.text(comp_center_x, -0.8, f"R={r:.4f} Ом", ha="center", fontsize=8)
        ax.text(comp_center_x, -1.3, f"X={x_val:.4f} Ом", ha="center", fontsize=8)
        prev_conn_x = comp_center_x + 0.5
    ax.plot([prev_conn_x, prev_conn_x + 0.8], [0, 0], "k-", linewidth=1.5)
    fault_x = prev_conn_x + 1.2
    ax.text(fault_x, 0, "⚡", fontsize=25, color="red", ha="center", va="center")
    ax.text(fault_x, -1.0, f"КЗ {name}", ha="center", fontsize=9, color="red", weight="bold")
    ax.set_xlim(-0.5, fault_x + 1.0)
    ax.set_ylim(-1.8, 1.8)
    mem_file = io.BytesIO()
    plt.savefig(mem_file, format="png", bbox_inches="tight", dpi=150)
    plt.close(fig)
    mem_file.seek(0)
    return mem_file


def create_selectivity_map(
    up_node: dict[str, Any],
    down_node: dict[str, Any],
    i_min_ka: float,
    i_max_ka: float,
) -> tuple[io.BytesIO, bool]:
    plt, _patches = _require_plt()
    up_i, up_t = get_curve_points(up_node)
    dw_i, dw_t = get_curve_points(down_node)
    up_label = f"Верхний: {up_node.get('Name', 'Unknown')}"
    dw_label = f"Нижний: {down_node.get('Name', 'Unknown')}"
    fig, ax = plt.subplots(figsize=(7, 6))
    ax.loglog(dw_i, dw_t, "b-", label=dw_label)
    ax.loglog(up_i, up_t, "r-", label=up_label)
    i_min, i_max = i_min_ka * 1000.0, i_max_ka * 1000.0
    ax.axvline(i_min, color="orange", linestyle="--", label=f"Iкз.мин={i_min:.0f}A")
    ax.axvline(i_max, color="green", linestyle="--", label=f"Iкз.макс={i_max:.0f}A")
    ax.set_xlabel("Ток (A)")
    ax.set_ylabel("Время (с)")
    ax.grid(True, which="both", ls="--")
    ax.set_xlim(left=max(1, dw_i[0]), right=max(i_max * 5, 10000))
    ax.set_ylim(0.001, 1000)
    ax.legend(loc="upper right")
    overlap = False
    mask = (dw_i >= i_min) & (dw_i <= i_max)
    if len(dw_i[mask]) > 0:
        if any(dw_t[mask] >= np.interp(dw_i[mask], up_i, up_t)):
            overlap = True
    mem_file = io.BytesIO()
    plt.savefig(mem_file, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    mem_file.seek(0)
    return mem_file, overlap
