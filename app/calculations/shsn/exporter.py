from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from .models import ShsnInput, ShsnResult


class ShsnWordExporter:
    def export(self, parent: tk.Misc, input_model: ShsnInput, result_model: ShsnResult) -> None:
        del input_model, result_model
        messagebox.showinfo("Word", "Экспорт для расчета ЩСН будет добавлен позже.", parent=parent)

