from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

REQUIRED_COLUMNS: tuple[str, ...] = (
    "Name",
    "Type",
    "Parent_Name",
    "Phase_Type",
    "R1",
    "X1",
    "R0",
    "X0",
    "U_nom",
    "Is_SC_Point",
    "CB_Nominal",
    "CB_Multiplier",
    "CB_Time",
    "CB_Curve",
)

DISPLAY_COLUMN_RENAME: dict[str, str] = {
    "Name": "Наименование",
    "Type": "Тип",
    "Parent_Name": "Вышестоящий элемент",
    "Phase_Type": "Кол-во полюсов",
    "U_nom": "Uном",
    "R1": "R1 (мОм)",
    "X1": "X1 (мОм)",
    "R0": "R0 (мОм)",
    "X0": "X0 (мОм)",
}

EXCLUDE_FROM_SOURCE_TABLE: frozenset[str] = frozenset({
    "Is_SC_Point",
    "CB_Nominal",
    "CB_Multiplier",
    "CB_Time",
    "CB_Curve",
})

CB_CURVE_MULTIPLIERS: dict[str, int] = {"B": 5, "C": 10, "D": 20, "K": 14, "Z": 3}

DEFAULT_K_TEMP = 1.3


@dataclass(slots=True)
class NetworkElement:
    name: str
    type: str
    parent_name: str
    phase_type: str
    r1: float
    x1: float
    r0: float
    x0: float
    u_nom: float
    is_sc_point: bool
    cb_nominal: float
    cb_multiplier: float
    cb_time: float
    cb_curve: str

    def as_graph_node(self) -> dict[str, Any]:
        return {
            "Name": self.name,
            "Type": self.type,
            "Parent_Name": self.parent_name,
            "Phase_Type": self.phase_type,
            "R1": self.r1,
            "X1": self.x1,
            "R0": self.r0,
            "X0": self.x0,
            "U_nom": self.u_nom,
            "Is_SC_Point": self.is_sc_point,
            "CB_Nominal": self.cb_nominal,
            "CB_Multiplier": self.cb_multiplier,
            "CB_Time": self.cb_time,
            "CB_Curve": self.cb_curve,
        }


@dataclass(slots=True)
class ElementOnPath:
    name: str
    type: str
    r1: float
    x1: float
    r0: float
    x0: float


@dataclass(slots=True)
class ScPointResult:
    path: list[str]
    phase_type: str
    elements_data: list[ElementOnPath]
    r_sum_1: float
    x_sum_1: float
    r_sum_0: float
    x_sum_0: float
    i_sc_3: float
    i_sc_1_cold: float
    i_sc_1_hot: float
    u_nom: float
    r_total_1ph_cold: float
    r_total_1ph_hot: float
    x_total_1ph: float


@dataclass(slots=True)
class ShsnInput:
    elements: list[NetworkElement] = field(default_factory=list)
    k_temp: float = DEFAULT_K_TEMP
    project_name: str = ""


@dataclass(slots=True)
class ShsnResult:
    sc_results: dict[str, ScPointResult] = field(default_factory=dict)
    message: str = ""
