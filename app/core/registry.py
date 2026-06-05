from __future__ import annotations

from app.calculations.equipment_check.plugin import build_plugin as build_equipment_check_plugin
from app.calculations.fire_check.plugin import build_plugin as build_fire_check_plugin
from app.calculations.shsn.plugin import build_plugin as build_shsn_plugin
from app.calculations.sopt.plugin import build_plugin as build_sopt_plugin
from app.core.contracts import CalculationPlugin


def get_calculation_plugins() -> list[CalculationPlugin]:
    return [
        build_fire_check_plugin(),
        build_sopt_plugin(),
        build_shsn_plugin(),
        build_equipment_check_plugin(),
    ]

