"""Проверка кабеля свыше 1 кВ (6–10 кВ) на невозгорание — отчёт Word."""

from .calculator import MvCableFireCalculator
from .defaults import demo_mv_cable_input
from .excel_template import TEMPLATE_FILENAME, build_mv_cable_excel_template_bytes
from .models import MvCableFireInput, MvCableFireResult
from .project_store import load_mv_cable_project, snapshot_to_input, snapshot_to_inputs
from .word_export import build_mv_cable_word_report, build_mv_cable_word_report_multi

__all__ = [
    "MvCableFireCalculator",
    "MvCableFireInput",
    "MvCableFireResult",
    "TEMPLATE_FILENAME",
    "build_mv_cable_excel_template_bytes",
    "build_mv_cable_word_report",
    "build_mv_cable_word_report_multi",
    "demo_mv_cable_input",
    "load_mv_cable_project",
    "snapshot_to_input",
    "snapshot_to_inputs",
]
