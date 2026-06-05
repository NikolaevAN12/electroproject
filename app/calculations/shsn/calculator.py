from __future__ import annotations

from .models import ShsnInput, ShsnResult


class ShsnCalculator:
    def calculate(self, input_model: ShsnInput) -> ShsnResult:
        del input_model
        return ShsnResult()

