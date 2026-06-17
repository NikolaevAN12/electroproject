"""Демо-входные данные (без Excel) — по образцу «Выбор и проверка КЛ 6 кВ до ТСН»."""

from __future__ import annotations

from .models import MvCableFireInput


def demo_mv_cable_input() -> MvCableFireInput:
    return MvCableFireInput(
        project_name="Выбор и проверка КЛ 6 кВ до ТСН",
    )
