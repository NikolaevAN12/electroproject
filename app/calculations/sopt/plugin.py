from __future__ import annotations

from app.core.contracts import CalculationPlugin

from .ui import SoptWidget


def build_plugin() -> CalculationPlugin:
    return CalculationPlugin(id="sopt", title="Расчет СОПТ", widget_factory=SoptWidget)

