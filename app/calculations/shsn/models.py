from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class ShsnInput:
    note: str = ""


@dataclass(slots=True)
class ShsnResult:
    message: str = "Скоро"

