from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Generic, Protocol, TypeVar

import tkinter as tk

InputT = TypeVar("InputT")
ResultT = TypeVar("ResultT")


class InputModel(Protocol):
    """Marker protocol for calculation input models."""


class ResultModel(Protocol):
    """Marker protocol for calculation result models."""


@dataclass(slots=True)
class ValidationError:
    field: str
    message: str


class Validator(Protocol, Generic[InputT]):
    def validate(self, input_model: InputT) -> list[ValidationError]:
        ...


class Calculator(Protocol, Generic[InputT, ResultT]):
    def calculate(self, input_model: InputT) -> ResultT:
        ...


class WordExporter(Protocol, Generic[InputT, ResultT]):
    def export(self, parent: tk.Misc, input_model: InputT, result_model: ResultT) -> None:
        ...


class CalculationWidget(Protocol):
    def build(self, parent: tk.Misc) -> tk.Widget:
        ...


@dataclass(slots=True)
class CalculationPlugin:
    id: str
    title: str
    widget_factory: Callable[[], CalculationWidget]

