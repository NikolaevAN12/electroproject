from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from app.core.contracts import CalculationWidget

from .calculator import EquipmentCheckCalculator
from .exporter import EquipmentCheckWordExporter
from .models import EquipmentCheckInput
from .validator import EquipmentCheckValidator


class EquipmentCheckWidget(CalculationWidget):
    def __init__(self) -> None:
        self._validator = EquipmentCheckValidator()
        self._calculator = EquipmentCheckCalculator()
        self._exporter = EquipmentCheckWordExporter()
        self._root: tk.Widget | None = None
        self._result_lbl: tk.Label | None = None

    def build(self, parent: tk.Misc) -> tk.Widget:
        root = tk.Frame(parent, bg="#f0f0f0")
        self._root = root
        tk.Label(root, text="Проверка оборудования", font=("Segoe UI", 16, "bold"), bg="#f0f0f0").pack(pady=(24, 8))
        tk.Label(root, text="Скоро", font=("Segoe UI", 36, "bold"), bg="#f0f0f0").pack(pady=8)
        self._result_lbl = tk.Label(root, text="", font=("Segoe UI", 12), bg="#f0f0f0", fg="#2d3748")
        self._result_lbl.pack(pady=(0, 8))
        controls = tk.Frame(root, bg="#f0f0f0")
        controls.pack(pady=12)
        tk.Button(controls, text="Расчет", command=self._run).pack(side=tk.LEFT, padx=4)
        tk.Button(controls, text="В Word", command=lambda: self._export(root)).pack(side=tk.LEFT, padx=4)
        return root

    def _run(self) -> None:
        model = EquipmentCheckInput()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning("Проверка данных", "\n".join(err.message for err in errors[:10]), parent=self._root)
            return
        result = self._calculator.calculate(model)
        if self._result_lbl is not None:
            self._result_lbl.configure(text=result.message)

    def _export(self, parent: tk.Misc) -> None:
        model = EquipmentCheckInput()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning("Проверка данных", "\n".join(err.message for err in errors[:10]), parent=self._root)
            return
        result = self._calculator.calculate(model)
        if self._result_lbl is not None:
            self._result_lbl.configure(text=result.message)
        self._exporter.export(parent, model, result)

