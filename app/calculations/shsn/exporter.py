from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from app.shared.user_errors import MSG_MISSING_MATPLOTLIB
from app.shared.word_export_base import export_word_document

from .models import ShsnInput, ShsnResult
from .word_export import populate_shsn_document


def _check_matplotlib(parent: tk.Misc) -> bool:
    try:
        import matplotlib  # noqa: F401
    except ImportError:
        messagebox.showerror("Word", MSG_MISSING_MATPLOTLIB, parent=parent)
        return False
    return True


class ShsnWordExporter:
    def export(self, parent: tk.Misc, input_model: ShsnInput, result_model: ShsnResult) -> None:
        if not result_model.sc_results:
            messagebox.showwarning(
                "Word",
                "Сначала выполните расчёт — нет результатов для экспорта.",
                parent=parent,
            )
            return
        stem = input_model.project_name.strip() or "Raschet_TKZ_0_4kV"
        export_word_document(
            parent,
            title="Сохранить расчёт ЩСН в Word",
            default_stem=stem,
            prefix="Приложение А. ",
            font_name="Times New Roman",
            font_size_pt=14,
            populate=populate_shsn_document,
            input_model=input_model,
            result_model=result_model,
            extra_checks=_check_matplotlib,
        )
