from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class ValidationError:
    field: str
    message: str
