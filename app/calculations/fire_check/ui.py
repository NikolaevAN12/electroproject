from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import cast

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from app.core.contracts import CalculationWidget
from app.shared.parsing import parse_number
from app.shared.user_errors import MSG_EXCEL_LOAD_FAILED, MSG_EXCEL_SAVE_FAILED

from .calculator import FireCheckCalculator
from .exporter import FireCheckWordExporter
from .models import FireCheckInput, FireCheckResult, FireCheckRowInput, FireCheckTsnParams
from .project_store import (
    FireCheckProjectSnapshot,
    default_excel_filename,
    dialog_initial_dir,
    load_project,
    save_project,
)
from .validator import FireCheckValidator



TABLE_HEADERS = (

    "№",

    "Кабель",

    "Сечение провода, мм²",

    "Допустимый ток, А.",

    "Рабочий ток, А",

    "Температура провода",

    "Ток КЗ, кА",

    "Время отключения, с",

    "Значение Bк",

    "Температура после КЗ",

)

TABLE_COLS = len(TABLE_HEADERS)

NUM_COL_WIDTH = 44

CALC_CELL_BG = "#dbeafe"

CALC_CELL_FG = "#1e3a8a"





@dataclass(slots=True)

class _RowWidgets:

    num: tk.Label

    cable: tk.Entry

    sect: tk.Entry

    allow: tk.Entry

    work: tk.Entry

    temp: tk.Label

    ikz: tk.Entry

    tok: tk.Entry

    vk: tk.Label

    itog: tk.Label





class FireCheckWidget(CalculationWidget):

    def __init__(self) -> None:

        self._rows: list[_RowWidgets] = []

        self._validator = FireCheckValidator()

        self._calculator = FireCheckCalculator()

        self._exporter = FireCheckWordExporter()

        self._root: tk.Widget | None = None

        self._canvas: tk.Canvas | None = None

        self._grid: tk.Frame | None = None

        self._plus_bar: tk.Frame | None = None

        self._cell_font: tkfont.Font | None = None

        self._button_font: tkfont.Font | None = None



        self._r1_var = tk.StringVar(value="")
        self._x1_var = tk.StringVar(value="")
        self._auto_ik1_var = tk.IntVar(value=1)
        self._project_var = tk.StringVar(value="")
        self._project_file_var = tk.StringVar(value="Файл не выбран")
        self._current_project_path: Path | None = None



    def build(self, parent: tk.Misc) -> tk.Widget:

        root = tk.Frame(parent, bg="#f0f0f0")

        self._root = root



        try:

            self._button_font = tkfont.Font(family="Segoe UI", size=13)

            self._cell_font = tkfont.Font(family="Segoe UI", size=11)

            header_font = tkfont.Font(family="Segoe UI", size=11, weight="bold")

        except tk.TclError:

            self._button_font = tkfont.Font(size=13)

            self._cell_font = tkfont.Font(size=11)

            header_font = tkfont.Font(size=11, weight="bold")



        params = tk.Frame(root, bg="#e8eef5", padx=12, pady=10)
        params.pack(fill=tk.X, padx=12, pady=(8, 0))

        tk.Label(
            params,
            text="1-я линия — подробный расчёт в Word: эквивалентные R₁эк, X₁эк до точки КЗ (мОм). "
            "Остальные линии: ток КЗ — только в столбце таблицы.",
            font=self._cell_font,
            bg="#e8eef5",
            fg="#1a202c",
            wraplength=720,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=7, sticky=tk.W, pady=(0, 4))

        tk.Label(params, text="Проект:", font=self._cell_font, bg="#e8eef5").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 6), pady=2
        )
        tk.Entry(params, textvariable=self._project_var, width=28, font=self._cell_font).grid(
            row=1, column=1, columnspan=2, sticky=tk.W, pady=2
        )

        btn_style = dict(
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            padx=10,
            pady=2,
        )
        tk.Button(
            params,
            text="Открыть…",
            bg="#2c5282",
            fg="white",
            activebackground="#2b6cb0",
            command=self._open_project_dialog,
            **btn_style,
        ).grid(row=1, column=3, padx=(12, 4))
        tk.Button(
            params,
            text="Сохранить…",
            bg="#38a169",
            fg="white",
            activebackground="#2f855a",
            command=self._save_project_dialog,
            **btn_style,
        ).grid(row=1, column=4, padx=(4, 0))
        tk.Label(params, text="Файл:", font=self._cell_font, bg="#e8eef5").grid(
            row=2, column=0, sticky=tk.W, pady=(6, 0)
        )
        tk.Label(
            params,
            textvariable=self._project_file_var,
            font=self._cell_font,
            bg="#e8eef5",
            fg="#4a5568",
            wraplength=820,
            justify=tk.LEFT,
        ).grid(row=2, column=1, columnspan=4, sticky=tk.W, pady=(6, 0))

        tk.Label(params, text="R1 (мОм):", font=self._cell_font, bg="#e8eef5").grid(
            row=3, column=0, sticky=tk.W, padx=(0, 6), pady=2
        )

        tk.Entry(params, textvariable=self._r1_var, width=12, font=self._cell_font).grid(
            row=3, column=1, sticky=tk.W, pady=2
        )
        tk.Label(params, text="X1 (мОм):", font=self._cell_font, bg="#e8eef5").grid(
            row=3, column=2, sticky=tk.W, padx=(12, 6), pady=2
        )
        tk.Entry(params, textvariable=self._x1_var, width=12, font=self._cell_font).grid(
            row=3, column=3, sticky=tk.W, pady=2
        )
        tk.Checkbutton(
            params,
            text="Если в 1-й строке «Ток КЗ» пуст — рассчитать Iкз из R1 и X1",
            variable=self._auto_ik1_var,
            font=self._cell_font,
            bg="#e8eef5",
            activebackground="#e8eef5",
        ).grid(row=4, column=0, columnspan=8, sticky=tk.W, pady=(6, 0))



        table_outer = tk.Frame(root, bg="#cbd5e0", padx=1, pady=1)

        table_outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)



        canvas = tk.Canvas(table_outer, bg="#f0f0f0", highlightthickness=0)

        self._canvas = canvas

        scroll_y = tk.Scrollbar(table_outer, orient=tk.VERTICAL, command=canvas.yview)

        canvas.configure(yscrollcommand=scroll_y.set)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)



        inner = tk.Frame(canvas, bg="#f0f0f0")

        canvas_window = canvas.create_window((0, 0), window=inner, anchor=tk.NW)



        def sync_scroll_region(_event: tk.Event | None = None) -> None:

            del _event

            canvas.update_idletasks()

            canvas.configure(scrollregion=canvas.bbox("all"))



        def stretch_inner(event: tk.Event) -> None:

            canvas.itemconfigure(canvas_window, width=event.width)



        inner.bind("<Configure>", sync_scroll_region)

        canvas.bind("<Configure>", stretch_inner)



        grid = tk.Frame(inner, bg="#f0f0f0")

        self._grid = grid

        grid.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        grid.grid_columnconfigure(0, weight=0, minsize=NUM_COL_WIDTH)

        for col in range(1, TABLE_COLS):

            grid.grid_columnconfigure(col, weight=1, uniform="data_cols")



        for idx, title in enumerate(TABLE_HEADERS):

            tk.Label(

                grid,

                text=title,

                font=header_font,

                bg="#e2e8f0",

                fg="#1a202c",

                anchor=tk.CENTER if idx == 0 else tk.W,

                padx=6 if idx == 0 else 8,

                pady=6,

                relief=tk.FLAT,

            ).grid(row=0, column=idx, sticky="nsew", padx=1, pady=1)



        plus_bar = tk.Frame(grid, bg="#f0f0f0")

        self._plus_bar = plus_bar

        tk.Button(

            plus_bar,

            text="+",

            font=self._button_font,

            width=4,

            cursor="hand2",

            relief=tk.FLAT,

            bg="#38a169",

            fg="white",

            activebackground="#2f855a",

            command=self._add_row,

        ).pack(side=tk.LEFT, padx=4)

        tk.Button(

            plus_bar,

            text="Расчет",

            font=self._button_font,

            cursor="hand2",

            relief=tk.FLAT,

            bg="#2c5282",

            fg="white",

            activebackground="#2b6cb0",

            padx=16,

            command=self._run_calculation,

        ).pack(side=tk.LEFT, padx=(12, 4))

        tk.Button(

            plus_bar,

            text="В Word",

            font=self._button_font,

            cursor="hand2",

            relief=tk.FLAT,

            bg="#553c9a",

            fg="white",

            activebackground="#6b46c1",

            padx=16,

            command=self._export_to_word,

        ).pack(side=tk.LEFT, padx=(8, 4))



        self._add_row()
        self._layout_plus_row()
        sync_scroll_region()
        return root



    def _set_entry(self, entry: tk.Entry, text: str) -> None:
        entry.delete(0, tk.END)
        if text:
            entry.insert(0, text)

    def _collect_snapshot(self) -> FireCheckProjectSnapshot:
        name = self._project_var.get().strip()
        row_data: list[dict[str, str]] = []
        for idx, row in enumerate(self._rows, start=1):
            row_data.append(
                {
                    "num": str(idx),
                    "cable": row.cable.get().strip(),
                    "sect": row.sect.get().strip(),
                    "allow": row.allow.get().strip(),
                    "work": row.work.get().strip(),
                    "ikz": row.ikz.get().strip(),
                    "tok": row.tok.get().strip(),
                }
            )
        return FireCheckProjectSnapshot(
            project_name=name,
            r1_mohm=self._r1_var.get().strip(),
            x1_mohm=self._x1_var.get().strip(),
            auto_ik1=bool(self._auto_ik1_var.get()),
            rows=row_data,
        )

    def _apply_snapshot(self, snap: FireCheckProjectSnapshot) -> None:
        self._project_var.set(snap.project_name)
        self._r1_var.set(snap.r1_mohm)
        self._x1_var.set(snap.x1_mohm)
        self._auto_ik1_var.set(1 if snap.auto_ik1 else 0)
        self._clear_table_rows()
        if not snap.rows:
            self._add_row()
        else:
            for row in snap.rows:
                self._add_row()
                widgets = self._rows[-1]
                self._set_entry(widgets.cable, row.get("cable", ""))
                self._set_entry(widgets.sect, row.get("sect", ""))
                self._set_entry(widgets.allow, row.get("allow", ""))
                self._set_entry(widgets.work, row.get("work", ""))
                self._set_entry(widgets.ikz, row.get("ikz", ""))
                self._set_entry(widgets.tok, row.get("tok", ""))
                widgets.temp.configure(text="")
                widgets.vk.configure(text="")
                widgets.itog.configure(text="")
        self._layout_plus_row()
        if self._canvas is not None:
            self._canvas.update_idletasks()
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _clear_table_rows(self) -> None:
        for row in self._rows:
            for widget in (
                row.num,
                row.cable,
                row.sect,
                row.allow,
                row.work,
                row.temp,
                row.ikz,
                row.tok,
                row.vk,
                row.itog,
            ):
                widget.destroy()
        self._rows.clear()

    def _show_validation_errors(self, errors: list) -> None:
        messagebox.showwarning(
            "Проверка данных",
            "\n".join(err.message for err in errors[:10]),
            parent=self._root,
        )

    def _set_current_project_path(self, path: Path | None) -> None:
        self._current_project_path = path.resolve() if path is not None else None
        if self._current_project_path is None:
            self._project_file_var.set("Файл не выбран")
        else:
            self._project_file_var.set(str(self._current_project_path))

    def _pick_open_path(self) -> Path | None:
        if self._root is None:
            return None
        selected = filedialog.askopenfilename(
            parent=self._root,
            title="Открыть проект",
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialdir=str(dialog_initial_dir()),
        )
        if not selected:
            return None
        return Path(selected)

    def _pick_save_path(self) -> Path | None:
        if self._root is None:
            return None
        initial = self._current_project_path
        initialdir = str(initial.parent) if initial is not None else str(dialog_initial_dir())
        initialfile = initial.name if initial is not None else default_excel_filename(self._project_var.get())
        selected = filedialog.asksaveasfilename(
            parent=self._root,
            title="Сохранить проект",
            defaultextension=".xlsx",
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if not selected:
            return None
        path = Path(selected)
        if path.suffix.lower() != ".xlsx":
            path = path.with_suffix(".xlsx")
        return path

    def _open_project_dialog(self) -> None:
        path = self._pick_open_path()
        if path is None:
            return
        self._load_from_path(path)

    def _save_project_dialog(self) -> None:
        if not self._project_var.get().strip():
            messagebox.showwarning(
                "Сохранение проекта",
                "Введите название проекта — оно попадёт в файл Excel и в отчёт.",
                parent=self._root,
            )
            return
        path = self._pick_save_path()
        if path is None:
            return
        try:
            saved = save_project(self._collect_snapshot(), path)
        except (OSError, RuntimeError):
            messagebox.showerror("Сохранение проекта", MSG_EXCEL_SAVE_FAILED, parent=self._root)
            return
        self._set_current_project_path(saved)
        messagebox.showinfo("Сохранение проекта", f"Данные сохранены:\n{saved}", parent=self._root)

    def _load_from_path(self, path: Path) -> None:
        try:
            snap = load_project(path)
        except (OSError, RuntimeError, ValueError):
            messagebox.showerror("Загрузка проекта", MSG_EXCEL_LOAD_FAILED, parent=self._root)
            return
        self._apply_snapshot(snap)
        self._set_current_project_path(path)

    def _collect_line1_params(self) -> FireCheckTsnParams | None:

        r1 = parse_number(self._r1_var.get())

        x1 = parse_number(self._x1_var.get())

        if r1 is None or x1 is None:

            return None

        return FireCheckTsnParams(r1_mohm=r1, x1_mohm=x1)



    def _maybe_autofill_ik_row1(self, model: FireCheckInput) -> FireCheckInput:

        if not self._auto_ik1_var.get() or not model.tsn or not self._rows:

            return model

        tsn = model.tsn

        if tsn.r1_mohm <= 0 or tsn.x1_mohm <= 0:

            return model

        ik3 = FireCheckCalculator.three_phase_sc_current_ka(tsn.r1_mohm, tsn.x1_mohm, tsn.u_line_v)

        if ik3 is None:

            return model

        if self._rows[0].ikz.get().strip():

            return model

        e = self._rows[0].ikz

        e.delete(0, tk.END)

        e.insert(0, f"{ik3:.3f}".replace(".", ","))

        return self._collect_input()



    def _add_row(self) -> None:

        if self._grid is None or self._cell_font is None:

            return

        row_index = len(self._rows) + 1

        grid = self._grid

        font = self._cell_font



        num_lbl = tk.Label(grid, text=str(row_index), font=font, bg="#edf2f7", fg="#1a202c", padx=4, pady=6)

        num_lbl.grid(row=row_index, column=0, sticky="nsew", padx=1, pady=1)

        cable_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        cable_e.grid(row=row_index, column=1, sticky="nsew", padx=1, pady=1, ipady=4)

        sect_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        sect_e.grid(row=row_index, column=2, sticky="nsew", padx=1, pady=1, ipady=4)

        allow_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        allow_e.grid(row=row_index, column=3, sticky="nsew", padx=1, pady=1, ipady=4)

        work_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        work_e.grid(row=row_index, column=4, sticky="nsew", padx=1, pady=1, ipady=4)

        temp_lbl = tk.Label(grid, text="", font=font, bg=CALC_CELL_BG, fg=CALC_CELL_FG, padx=6, pady=6)

        temp_lbl.grid(row=row_index, column=5, sticky="nsew", padx=1, pady=1)

        ikz_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        ikz_e.grid(row=row_index, column=6, sticky="nsew", padx=1, pady=1, ipady=4)

        tok_e = tk.Entry(grid, font=font, relief=tk.FLAT, bg="white")

        tok_e.grid(row=row_index, column=7, sticky="nsew", padx=1, pady=1, ipady=4)

        vk_lbl = tk.Label(grid, text="", font=font, bg=CALC_CELL_BG, fg=CALC_CELL_FG, padx=6, pady=6)

        vk_lbl.grid(row=row_index, column=8, sticky="nsew", padx=1, pady=1)

        itog_lbl = tk.Label(grid, text="", font=font, bg=CALC_CELL_BG, fg=CALC_CELL_FG, padx=6, pady=6)

        itog_lbl.grid(row=row_index, column=9, sticky="nsew", padx=1, pady=1)



        self._rows.append(

            _RowWidgets(

                num=num_lbl,

                cable=cable_e,

                sect=sect_e,

                allow=allow_e,

                work=work_e,

                temp=temp_lbl,

                ikz=ikz_e,

                tok=tok_e,

                vk=vk_lbl,

                itog=itog_lbl,

            )

        )

        self._layout_plus_row()

        if self._canvas is not None:

            self._canvas.update_idletasks()

            self._canvas.configure(scrollregion=self._canvas.bbox("all"))



    def _layout_plus_row(self) -> None:

        if self._plus_bar is None:

            return

        self._plus_bar.grid_forget()

        self._plus_bar.grid(

            row=len(self._rows) + 1,

            column=0,

            columnspan=TABLE_COLS,

            sticky="ew",

            padx=1,

            pady=(6, 4),

        )



    def _collect_input(self) -> FireCheckInput:

        rows: list[FireCheckRowInput] = []

        for idx, row in enumerate(self._rows, start=1):

            rows.append(

                FireCheckRowInput(

                    number=idx,

                    cable=row.cable.get().strip(),

                    section_mm2=parse_number(row.sect.get()),

                    allowed_current_a=parse_number(row.allow.get()),

                    working_current_a=parse_number(row.work.get()),

                    short_circuit_ka=parse_number(row.ikz.get()),

                    shutdown_time_s=parse_number(row.tok.get()),

                )

            )

        line1 = self._collect_line1_params()

        return FireCheckInput(
            rows=rows,
            tsn=line1,
            project_name=self._project_var.get().strip(),
        )



    def _render_result(self, result: FireCheckResult) -> None:

        for row_widgets, row_result in zip(self._rows, result.rows):

            row_widgets.temp.configure(

                text="" if row_result.wire_temperature_c is None else str(row_result.wire_temperature_c)

            )

            row_widgets.vk.configure(text=row_result.bk_display)

            row_widgets.itog.configure(text=row_result.final_temperature_display)



    def _run_calculation(self) -> FireCheckResult | None:

        model = self._collect_input()

        model = self._maybe_autofill_ik_row1(model)

        errors = self._validator.validate(model)

        if errors:

            messagebox.showwarning("Проверка данных", "\n".join(err.message for err in errors[:10]), parent=self._root)

            return None

        result = self._calculator.calculate(model)

        self._render_result(result)

        return result



    def _export_to_word(self) -> None:

        model = self._collect_input()

        model = self._maybe_autofill_ik_row1(model)

        errors = self._validator.validate(model)

        if errors:

            messagebox.showwarning("Проверка данных", "\n".join(err.message for err in errors[:10]), parent=self._root)

            return

        result = self._calculator.calculate(model)

        self._render_result(result)
        self._exporter.export(cast(tk.Misc, self._root), model, result)


