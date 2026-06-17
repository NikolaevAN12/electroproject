from __future__ import annotations

from app.core.contracts import ValidationError

from .models import EquipmentCheckInput


class EquipmentCheckValidator:
    def validate(self, input_model: EquipmentCheckInput) -> list[ValidationError]:
        del input_model
        return []

