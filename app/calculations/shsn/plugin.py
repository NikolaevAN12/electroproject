from __future__ import annotations

from app.core.contracts import CalculationPlugin

from .ui import ShsnWidget


def build_plugin() -> CalculationPlugin:
    return CalculationPlugin(id="shsn", title="Расчет ЩСН", widget_factory=ShsnWidget)

