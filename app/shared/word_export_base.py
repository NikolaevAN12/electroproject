"""Общая логика экспорта в Word: диалог сохранения и запись файла."""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox

from app.shared.paths import word_export_dir
from app.shared.user_errors import MSG_MISSING_DOCX, MSG_WORD_SAVE_FAILED


def require_docx(parent: tk.Misc) -> tuple[Any, Any] | None:
    """Возвращает (Document, Pt) или None, если python-docx недоступен."""
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        messagebox.showerror("Word", MSG_MISSING_DOCX, parent=parent)
        return None
    return Document, Pt


def ask_word_save_path(
    parent: tk.Misc,
    *,
    title: str,
    default_stem: str,
    prefix: str = "",
) -> Path | None:
    safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in default_stem.strip())[:60]
    safe = safe or "proekt"
    name_part = f"{prefix}{safe}" if prefix else safe
    default_name = f"{name_part}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    default_path = word_export_dir() / default_name
    output_file = filedialog.asksaveasfilename(
        parent=parent,
        title=title,
        defaultextension=".docx",
        filetypes=[("Word document", "*.docx"), ("All files", "*.*")],
        initialfile=default_path.name,
        initialdir=str(default_path.parent),
    )
    if not output_file:
        return None
    return Path(output_file)


def save_word_document(parent: tk.Misc, doc: Any, output_path: Path) -> bool:
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
    except OSError:
        messagebox.showerror("Word", MSG_WORD_SAVE_FAILED, parent=parent)
        return False
    return True


def export_word_document(
    parent: tk.Misc,
    *,
    title: str,
    default_stem: str,
    prefix: str,
    font_name: str,
    font_size_pt: int,
    populate: Callable[[Any, Any, Any], None],
    input_model: Any,
    result_model: Any,
    extra_checks: Callable[[tk.Misc], bool] | None = None,
) -> None:
    """Проверка зависимостей, диалог, заполнение и сохранение Word."""
    docx_types = require_docx(parent)
    if docx_types is None:
        return
    Document, Pt = docx_types

    if extra_checks is not None and not extra_checks(parent):
        return

    output_path = ask_word_save_path(
        parent,
        title=title,
        default_stem=default_stem,
        prefix=prefix,
    )
    if output_path is None:
        return

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(font_size_pt)
    populate(doc, input_model, result_model)

    if not save_word_document(parent, doc, output_path):
        return

    messagebox.showinfo("Word", f"Документ сохранён:\n{output_path.resolve()}", parent=parent)
