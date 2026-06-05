from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from .models import EquipmentCheckInput, EquipmentCheckResult


class EquipmentCheckWordExporter:
    def export(
        self,
        parent: tk.Misc,
        input_model: EquipmentCheckInput,
        result_model: EquipmentCheckResult,
    ) -> None:
        del input_model, result_model
        messagebox.showinfo("Word", "Экспорт для проверки оборудования будет добавлен позже.", parent=parent)

