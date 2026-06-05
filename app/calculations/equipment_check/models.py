from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class EquipmentCheckInput:
    note: str = ""


@dataclass(slots=True)
class EquipmentCheckResult:
    message: str = "Скоро"

