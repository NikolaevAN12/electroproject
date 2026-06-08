from __future__ import annotations

from app.core.contracts import CalculationPlugin

from .ui import ShsnWidget


def build_plugin() -> CalculationPlugin:
    return CalculationPlugin(
        id="shsn",
        title="Расчёт ТКЗ 0,4 кВ (ЩСН)",
        widget_factory=ShsnWidget,
    )

