from __future__ import annotations

from pathlib import Path

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox

from app.core.contracts import CalculationWidget
from app.shared.parsing import parse_number
from app.shared.user_errors import MSG_EXCEL_LOAD_FAILED, MSG_EXCEL_SAVE_FAILED

from .calculator import ShsnCalculator
from .exporter import ShsnWordExporter
from .models import DEFAULT_K_TEMP, ShsnInput, ShsnResult
from .project_store import (
    TEMPLATE_NAME,
    ShsnProjectSnapshot,
    create_blank_template,
    default_excel_filename,
    dialog_initial_dir,
    load_project,
    remember_last_project,
    save_project,
)
from .validator import ShsnValidator


class ShsnWidget(CalculationWidget):
    def __init__(self) -> None:
        self._validator = ShsnValidator()
        self._calculator = ShsnCalculator()
        self._exporter = ShsnWordExporter()
        self._root: tk.Widget | None = None
        self._result_lbl: tk.Label | None = None
        self._file_lbl: tk.Label | None = None
        self._stats_lbl: tk.Label | None = None
        self._k_temp_var = tk.StringVar(value=str(DEFAULT_K_TEMP))
        self._project_file_var = tk.StringVar(value="Файл не выбран")
        self._current_project_path: Path | None = None
        self._elements: list = []
        self._last_result: ShsnResult | None = None
        self._button_font: tkfont.Font | None = None

    def build(self, parent: tk.Misc) -> tk.Widget:
        root = tk.Frame(parent, bg="#f0f0f0")
        self._root = root
        try:
            self._button_font = tkfont.Font(family="Segoe UI", size=12)
            title_font = tkfont.Font(family="Segoe UI", size=16, weight="bold")
            hint_font = tkfont.Font(family="Segoe UI", size=11)
        except tk.TclError:
            self._button_font = tkfont.Font(size=12)
            title_font = tkfont.Font(size=16, weight="bold")
            hint_font = tkfont.Font(size=11)

        tk.Label(
            root,
            text="Расчёт ТКЗ в сети 0,4 кВ и проверка чувствительности АВ",
            font=title_font,
            bg="#f0f0f0",
            wraplength=900,
            justify=tk.CENTER,
        ).pack(pady=(16, 8))

        tk.Label(
            root,
            text="Модель сети задаётся в Excel (столбцы Name, Type, Parent_Name, …). "
            "Отметьте точки КЗ в столбце Is_SC_Point.",
            font=hint_font,
            bg="#f0f0f0",
            fg="#4a5568",
            wraplength=880,
            justify=tk.CENTER,
        ).pack(pady=(0, 12))

        toolbar = tk.Frame(root, bg="#f0f0f0")
        toolbar.pack(fill=tk.X, padx=12, pady=4)
        for text, cmd in (
            ("Открыть Excel…", self._open_project),
            ("Сохранить", self._save_project),
            ("Сохранить как…", self._save_project_as),
            ("Скачать бланк", self._save_template),
        ):
            tk.Button(
                toolbar,
                text=text,
                font=self._button_font,
                command=cmd,
                padx=10,
                pady=4,
            ).pack(side=tk.LEFT, padx=4)

        meta = tk.Frame(root, bg="#f0f0f0")
        meta.pack(fill=tk.X, padx=16, pady=8)
        tk.Label(meta, text="k_temp (нагрев):", font=hint_font, bg="#f0f0f0").pack(side=tk.LEFT)
        tk.Entry(meta, textvariable=self._k_temp_var, width=8, font=hint_font).pack(side=tk.LEFT, padx=(6, 20))
        self._file_lbl = tk.Label(
            meta,
            textvariable=self._project_file_var,
            font=hint_font,
            bg="#f0f0f0",
            fg="#2d3748",
        )
        self._file_lbl.pack(side=tk.LEFT)

        self._stats_lbl = tk.Label(root, text="", font=hint_font, bg="#f0f0f0", fg="#2d3748")
        self._stats_lbl.pack(pady=4)

        controls = tk.Frame(root, bg="#f0f0f0")
        controls.pack(pady=12)
        tk.Button(
            controls,
            text="Расчёт",
            font=self._button_font,
            command=self._run,
            padx=16,
            pady=6,
        ).pack(side=tk.LEFT, padx=6)
        tk.Button(
            controls,
            text="В Word",
            font=self._button_font,
            command=self._export,
            padx=16,
            pady=6,
        ).pack(side=tk.LEFT, padx=6)

        self._result_lbl = tk.Label(root, text="", font=hint_font, bg="#f0f0f0", fg="#276749")
        self._result_lbl.pack(pady=(4, 16))
        return root

    def _build_input(self) -> ShsnInput:
        k_temp = parse_number(self._k_temp_var.get())
        if k_temp is None or k_temp <= 0:
            k_temp = DEFAULT_K_TEMP
        project_name = (
            self._current_project_path.stem if self._current_project_path else "ЩСН"
        )
        return ShsnInput(
            elements=list(self._elements),
            k_temp=float(k_temp),
            project_name=project_name,
        )

    def _update_stats(self) -> None:
        if self._stats_lbl is None:
            return
        n = len(self._elements)
        sc = sum(1 for el in self._elements if el.is_sc_point)
        self._stats_lbl.configure(text=f"Элементов в модели: {n}  |  Точек КЗ: {sc}")

    def _apply_snapshot(self, snapshot: ShsnProjectSnapshot, path: Path | None) -> None:
        self._elements = list(snapshot.elements)
        self._k_temp_var.set(str(snapshot.k_temp))
        self._current_project_path = path
        self._last_result = None
        if path is not None:
            self._project_file_var.set(path.name)
        else:
            self._project_file_var.set("Файл не выбран")
        self._update_stats()
        if self._result_lbl is not None:
            self._result_lbl.configure(text="")

    def _open_project(self) -> None:
        path_str = filedialog.askopenfilename(
            parent=self._root,
            title="Открыть модель сети",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialdir=str(dialog_initial_dir()),
        )
        if not path_str:
            return
        path = Path(path_str)
        try:
            snapshot = load_project(path)
        except (OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("Excel", f"{MSG_EXCEL_LOAD_FAILED}\n\n{exc}", parent=self._root)
            return
        remember_last_project(path)
        self._apply_snapshot(snapshot, path)

    def _save_snapshot_to(self, path: Path) -> None:
        k_temp = parse_number(self._k_temp_var.get())
        if k_temp is None or k_temp <= 0:
            k_temp = DEFAULT_K_TEMP
        snapshot = ShsnProjectSnapshot(
            project_name=path.stem,
            k_temp=float(k_temp),
            elements=list(self._elements),
        )
        try:
            save_project(snapshot, path)
        except (OSError, RuntimeError) as exc:
            messagebox.showerror("Excel", f"{MSG_EXCEL_SAVE_FAILED}\n\n{exc}", parent=self._root)
            return
        self._current_project_path = path
        self._project_file_var.set(path.name)
        messagebox.showinfo("Excel", f"Сохранено:\n{path.resolve()}", parent=self._root)

    def _save_project(self) -> None:
        if self._current_project_path is not None:
            self._save_snapshot_to(self._current_project_path)
            return
        self._save_project_as()

    def _save_project_as(self) -> None:
        default_name = default_excel_filename("ЩСН")
        path_str = filedialog.asksaveasfilename(
            parent=self._root,
            title="Сохранить модель сети",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_name,
            initialdir=str(dialog_initial_dir()),
        )
        if not path_str:
            return
        self._save_snapshot_to(Path(path_str))

    def _save_template(self) -> None:
        path_str = filedialog.asksaveasfilename(
            parent=self._root,
            title="Сохранить бланк Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=TEMPLATE_NAME,
            initialdir=str(dialog_initial_dir()),
        )
        if not path_str:
            return
        try:
            create_blank_template(Path(path_str))
        except (OSError, RuntimeError) as exc:
            messagebox.showerror("Excel", f"{MSG_EXCEL_SAVE_FAILED}\n\n{exc}", parent=self._root)
            return
        messagebox.showinfo("Excel", f"Бланк сохранён:\n{Path(path_str).resolve()}", parent=self._root)

    def _run(self) -> None:
        model = self._build_input()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning(
                "Проверка данных",
                "\n".join(err.message for err in errors[:10]),
                parent=self._root,
            )
            return
        result = self._calculator.calculate(model)
        self._last_result = result
        if self._result_lbl is not None:
            self._result_lbl.configure(text=result.message)

    def _export(self) -> None:
        model = self._build_input()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning(
                "Проверка данных",
                "\n".join(err.message for err in errors[:10]),
                parent=self._root,
            )
            return
        if self._last_result is None or not self._last_result.sc_results:
            self._last_result = self._calculator.calculate(model)
        assert self._root is not None
        self._exporter.export(self._root, model, self._last_result)
