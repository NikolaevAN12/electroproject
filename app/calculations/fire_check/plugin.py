from __future__ import annotations

from app.core.contracts import CalculationPlugin

from .ui import FireCheckWidget


def build_plugin() -> CalculationPlugin:
    return CalculationPlugin(
        id="fire_check",
        title="Проверка на невозгорание",
        widget_factory=FireCheckWidget,
    )

