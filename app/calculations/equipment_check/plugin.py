from __future__ import annotations

from app.core.contracts import CalculationPlugin

from .ui import EquipmentCheckWidget


def build_plugin() -> CalculationPlugin:
    return CalculationPlugin(
        id="equipment_check",
        title="Проверка оборудования",
        widget_factory=EquipmentCheckWidget,
    )

