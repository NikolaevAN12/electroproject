"""Нумерация разделов отчёта Word для СОПТ."""

REPORT_SECTION_RESISTANCE = "2"
REPORT_SECTION_TKZ = "3"
REPORT_SECTION_TKZ_BUS = "3.1"
REPORT_SECTION_SELECTIVITY = "4"


def tkz_user_section(section_idx: int) -> str:
    """Пользовательский раздел ТКЗ: 3.2, 3.3, … (section_idx с 1)."""
    return f"{REPORT_SECTION_TKZ}.{section_idx + 1}"


def tkz_user_subsection(section_idx: int, subsection_idx: int) -> str:
    return f"{tkz_user_section(section_idx)}.{subsection_idx}"


def selectivity_section(section_idx: int) -> str:
    """Раздел карт селективности: 4.2, 4.3, … (section_idx с 1)."""
    return f"{REPORT_SECTION_SELECTIVITY}.{section_idx + 1}"


def selectivity_subsection(section_idx: int, map_idx: int) -> str:
    return f"{selectivity_section(section_idx)}.{map_idx}"
