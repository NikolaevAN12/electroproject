from __future__ import annotations

from .models import EquipmentCheckInput, EquipmentCheckResult


class EquipmentCheckCalculator:
    def calculate(self, input_model: EquipmentCheckInput) -> EquipmentCheckResult:
        del input_model
        return EquipmentCheckResult()

