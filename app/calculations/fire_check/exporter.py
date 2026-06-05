from __future__ import annotations

import tkinter as tk

from app.shared.word_export_base import export_word_document

from .models import FireCheckInput, FireCheckResult
from .raschet_sn_word_clone import populate_raschet_sn_clone


class FireCheckWordExporter:
    def export(self, parent: tk.Misc, input_model: FireCheckInput, result_model: FireCheckResult) -> None:
        stem = input_model.project_name.strip() or "Raschet_SN"
        export_word_document(
            parent,
            title="Сохранить в Word",
            default_stem=stem,
            prefix="",
            font_name="Times New Roman",
            font_size_pt=14,
            populate=populate_raschet_sn_clone,
            input_model=input_model,
            result_model=result_model,
        )
