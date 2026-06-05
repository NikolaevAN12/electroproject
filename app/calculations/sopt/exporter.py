from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from app.shared.user_errors import MSG_MISSING_PILLOW
from app.shared.word_export_base import export_word_document

from .models import SoptInput, SoptResult
from .report.word_export import populate_sopt_document


def _check_pillow(parent: tk.Misc) -> bool:
    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        messagebox.showerror("Word", MSG_MISSING_PILLOW, parent=parent)
        return False
    return True


class SoptWordExporter:
    def export(self, parent: tk.Misc, input_model: SoptInput, result_model: SoptResult) -> None:
        stem = input_model.project_name.strip() or "SOPT"
        export_word_document(
            parent,
            title="Сохранить расчёт СОПТ в Word",
            default_stem=stem,
            prefix="SOPT_",
            font_name="Arial",
            font_size_pt=12,
            populate=populate_sopt_document,
            input_model=input_model,
            result_model=result_model,
            extra_checks=_check_pillow,
        )
