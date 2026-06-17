"""Расчёт сопротивления элемента цепи для отчёта и схем."""

from __future__ import annotations

from .models import SoptEquipmentItem


def item_resistance_ohm(item: SoptEquipmentItem) -> float:
    if item.item_type in ("fuse", "breaker"):
        return max(0.0, item.resistance_ohm)
    if (
        item.item_type == "cable"
        and item.cable_gamma > 0
        and item.cable_section_mm2 > 0
        and item.cable_length_m > 0
    ):
        return item.cable_length_m / (item.cable_gamma * item.cable_section_mm2)
    return 0.0
