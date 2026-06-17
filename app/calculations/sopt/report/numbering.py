"""Нумерация разделов отчёта Word для СОПТ."""

REPORT_SECTION_DESCRIPTION = "1"
REPORT_SECTION_RESISTANCE = "2"
REPORT_SECTION_TKZ = "3"
REPORT_SECTION_TKZ_BUS = "3.1"
TKZ_BUS_KZ_POINT = "K1-1(2)"
REPORT_SECTION_SENSITIVITY = "4"
REPORT_SECTION_CABLE_FIRE = "5"
REPORT_SECTION_SELECTIVITY = "6"
REPORT_SECTION_CONCLUSION = "7"
REPORT_SECTION_REFERENCES = "8"


def tkz_user_section(section_idx: int) -> str:
    """Пользовательский раздел ТКЗ: 3.2, 3.3, … (section_idx с 1)."""
    return f"{REPORT_SECTION_TKZ}.{section_idx + 1}"


def tkz_user_subsection(section_idx: int, subsection_idx: int) -> str:
    return f"{tkz_user_section(section_idx)}.{subsection_idx}"


def cable_fire_table(table_idx: int) -> str:
    """Номер таблицы в разделе невозгорания: 5.1, 5.2, …"""
    return f"{REPORT_SECTION_CABLE_FIRE}.{table_idx}"


def selectivity_section(section_idx: int) -> str:
    """Раздел карт селективности: 6.1, 6.2, … (section_idx с 1)."""
    return f"{REPORT_SECTION_SELECTIVITY}.{section_idx}"


def selectivity_subsection(section_idx: int, map_idx: int) -> str:
    return f"{selectivity_section(section_idx)}.{map_idx}"


def sensitivity_table(table_idx: int) -> str:
    """Номер таблицы в разделе чувствительности: 4.1, 4.2, …"""
    return f"{REPORT_SECTION_SENSITIVITY}.{table_idx}"


def tkz_currents_table(table_idx: int = 1) -> str:
    """Номер сводной таблицы токов КЗ в разделе 3: 3.1, …"""
    return f"{REPORT_SECTION_TKZ}.{table_idx}"


def section3_last_figure_number(substitution_scheme_count: int) -> int:
    """Последний рисунок §3: Рис.1–2 в §3.1 и по одной схеме замещения на подраздел."""
    return 2 + substitution_scheme_count


def shpt_scheme_figure_number(substitution_scheme_count: int) -> int:
    """Первый рисунок общей схемы ЩПТ — следующий после последнего в §3."""
    return section3_last_figure_number(substitution_scheme_count) + 1


def shpt_scheme_figure_label(substitution_scheme_count: int) -> str:
    """Номер рисунка общей схемы ЩПТ (один номер, даже если листов CDW несколько)."""
    return str(shpt_scheme_figure_number(substitution_scheme_count))


# Резерв под рисунок между общей схемой ЩПТ (§3) и картами селективности (§6).
MANUAL_FIGURES_BEFORE_SELECTIVITY = 2


def selectivity_figure_start(fig_num_after_last_scheme: int) -> int:
    """Первый номер рисунка в §6 после резерва под ручные вставки.

    fig_num_after_last_scheme — следующий номер сразу после последней схемы замещения.
    При MANUAL_FIGURES_BEFORE_SELECTIVITY=2 резервируются два номера (N+1, N+2),
    карты селективности начинаются с N+3.
    """
    return fig_num_after_last_scheme + MANUAL_FIGURES_BEFORE_SELECTIVITY
