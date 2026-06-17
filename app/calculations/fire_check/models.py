from __future__ import annotations



from dataclasses import dataclass





@dataclass(slots=True)

class FireCheckTsnParams:

    """R₁, X₁ до точки КЗ для 1-й линии (мОм) — расчёт Iкз и подстановка в Word."""

    r1_mohm: float

    x1_mohm: float

    tae_s: float = 0.02

    u_line_v: float = 400.0

    theta_o_c: float = 25.0

    theta_dd_c: float = 70.0

    theta_okr_c: float = 25.0





@dataclass(slots=True)

class FireCheckRowInput:

    number: int

    cable: str

    section_mm2: float | None

    allowed_current_a: float | None

    working_current_a: float | None

    short_circuit_ka: float | None

    shutdown_time_s: float | None





@dataclass(slots=True)

class FireCheckInput:

    rows: list[FireCheckRowInput]

    tsn: FireCheckTsnParams | None = None

    project_name: str = ""





@dataclass(slots=True)

class FireCheckRowResult:

    number: int

    cable: str

    section_mm2_text: str

    allowed_current_a_text: str

    working_current_a_text: str

    short_circuit_ka_text: str

    shutdown_time_s_text: str

    wire_temperature_c: int | None

    bk_value: float | None

    bk_display: str

    final_temperature_display: str

    # Ток трёхфазного КЗ по R₁,X₁ ТСН (только для отображения / Word), кА

    sc_three_phase_ka: float | None = None





@dataclass(slots=True)

class FireCheckResult:

    rows: list[FireCheckRowResult]


