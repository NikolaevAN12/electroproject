from __future__ import annotations

from app.core.contracts import ValidationError

from .models import ShsnInput


class ShsnValidator:
    def validate(self, input_model: ShsnInput) -> list[ValidationError]:
        del input_model
        return []

