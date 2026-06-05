from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import cast

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from app.core.contracts import CalculationWidget
from app.shared.user_errors import MSG_EXCEL_LOAD_FAILED, MSG_EXCEL_SAVE_FAILED

from .calculator import SoptCalculator
from .exporter import SoptWordExporter
from .models import (
    BREAKER_CURVE_CHOICES,
    EQUIPMENT_LABEL_BY_KEY,
    EQUIPMENT_TYPES,
    SoptEquipmentItem,
    SoptInput,
    SoptSection,
    SoptSubsection,
)
from .project_store import (
    SoptProjectSnapshot,
    default_excel_filename,
    dialog_initial_dir,
    load_project,
    save_project,
)
from .validator import SoptValidator

_INPUT_FUSE_PRESETS: tuple[tuple[str, float], ...] = (
    ("16 А", 0.011),
    ("100 А", 0.001),
)

# Минимальная ширина колонок таблицы оборудования (пиксели), общая для заголовка и строк.
_EQUIPMENT_TABLE_COL_WIDTHS: tuple[int, ...] = (
    200,
    175,
    86,
    86,
    86,
    86,
    86,
    98,
    98,
    98,
    112,
    36,
)
_COL_TYPE = 0
_COL_DESIGNATION = 1
_COL_INOM = 2
_COL_CURVE = 3
_COL_MULT = 4
_COL_TIME = 5
_COL_R = 6
_COL_L = 7
_COL_Y = 8
_COL_S = 9
_COL_SEL = 10
_COL_DEL = 11


@dataclass(slots=True)
class _EquipmentRowWidgets:
    type_label: tk.Label
    designation: tk.Entry
    rated_current_a: tk.Entry
    cb_curve: ttk.Combobox | None
    cb_multiplier: tk.Entry | None
    cb_trip_time_s: tk.Entry | None
    resistance_ohm: tk.Entry
    cable_length_m: tk.Entry
    cable_gamma: tk.Entry
    cable_section_mm2: tk.Entry
    include_in_selectivity: tk.BooleanVar | None
    row_frame: tk.Frame
    item_type: str


@dataclass(slots=True)
class _SectionWidgets:
    frame: tk.LabelFrame
    name_entry: tk.Entry
    subsections_frame: tk.Frame
    subsections: list["_SubsectionWidgets"]


@dataclass(slots=True)
class _SubsectionWidgets:
    frame: tk.LabelFrame
    name_entry: tk.Entry
    equipment_table: tk.Frame
    next_equipment_row: int
    type_combo: ttk.Combobox
    equipment: list[_EquipmentRowWidgets]


class SoptWidget(CalculationWidget):
    def __init__(self) -> None:
        self._validator = SoptValidator()
        self._calculator = SoptCalculator()
        self._exporter = SoptWordExporter()
        self._root: tk.Widget | None = None
        self._sections_host: tk.Frame | None = None
        self._canvas: tk.Canvas | None = None
        self._sections: list[_SectionWidgets] = []
        self._cell_font: tkfont.Font | None = None
        self._button_font: tkfont.Font | None = None
        self._header_font: tkfont.Font | None = None
        self._project_var = tk.StringVar(value="")
        self._project_file_var = tk.StringVar(value="Файл не выбран")
        self._current_project_path: Path | None = None
        self._battery_name_var = tk.StringVar(value="")
        self._battery_cells_count_var = tk.StringVar(value="")
        self._battery_count_var = tk.StringVar(value="")
        self._r_el_var = tk.StringVar(value="")
        self._rho_var = tk.StringVar(value="")
        self._jumper_section_var = tk.StringVar(value="")
        self._input_fuse_pick_var = tk.StringVar(value="")
        self._input_fuse_r_var = tk.StringVar(value="")
        self._qn1_var = tk.StringVar(value="")
        self._q_calc_var = tk.StringVar(value="")
        self._result_lbl: tk.Label | None = None
        self._type_labels = [label for _, label in EQUIPMENT_TYPES]
        self._breaker_curve_labels = list(BREAKER_CURVE_CHOICES)
        self._disabled_field_bg = "#e2e8f0"
        self._row_height_px = 32

    def build(self, parent: tk.Misc) -> tk.Widget:
        root = tk.Frame(parent, bg="#f0f0f0")
        self._root = root

        try:
            self._button_font = tkfont.Font(family="Segoe UI", size=13)
            self._cell_font = tkfont.Font(family="Segoe UI", size=11)
            self._header_font = tkfont.Font(family="Segoe UI", size=12, weight="bold")
            self._table_header_small_font = tkfont.Font(family="Segoe UI", size=9)
        except tk.TclError:
            self._button_font = tkfont.Font(size=13)
            self._cell_font = tkfont.Font(size=11)
            self._header_font = tkfont.Font(size=12, weight="bold")
            self._table_header_small_font = tkfont.Font(size=9)

        combo_style = ttk.Style()
        combo_style.configure(
            "Sopt.TCombobox",
            padding=(2, 3),
            font=self._cell_font,
        )

        top = tk.Frame(root, bg="#e8eef5", padx=12, pady=10)
        top.pack(fill=tk.X, padx=12, pady=(8, 0))

        tk.Label(
            top,
            text="Разделы → подразделы → оборудование (предохранитель, автоматический выключатель, кабель). "
            "Для автоматов укажите Iном, характеристику (B/C/D…), кратность и время срабатывания, R (Ом); "
            "для кабеля — L (м), Y и S (мм²); для точки КЗ — обозначение; "
            "галочка «Проверка на селективность» — точка КЗ участвует в проверке селективности (§3 отчёта точка всегда в расчёте).",
            font=self._cell_font,
            bg="#e8eef5",
            fg="#1a202c",
            wraplength=900,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=7, sticky=tk.W, pady=(0, 6))

        tk.Label(top, text="Проект:", font=self._cell_font, bg="#e8eef5").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 6)
        )
        tk.Entry(top, textvariable=self._project_var, width=28, font=self._cell_font).grid(
            row=1, column=1, columnspan=2, sticky=tk.W
        )

        btn_kw = dict(font=self._cell_font, cursor="hand2", relief=tk.FLAT, padx=10, pady=2)
        tk.Button(
            top,
            text="Открыть…",
            bg="#2c5282",
            fg="white",
            activebackground="#2b6cb0",
            command=self._open_project_dialog,
            **btn_kw,
        ).grid(row=1, column=3, padx=(12, 4))
        tk.Button(
            top,
            text="Сохранить…",
            bg="#38a169",
            fg="white",
            activebackground="#2f855a",
            command=self._save_project_dialog,
            **btn_kw,
        ).grid(row=1, column=4, padx=(4, 0))
        tk.Label(top, text="Файл:", font=self._cell_font, bg="#e8eef5").grid(
            row=2, column=0, sticky=tk.W, pady=(6, 0)
        )
        tk.Label(
            top,
            textvariable=self._project_file_var,
            font=self._cell_font,
            bg="#e8eef5",
            fg="#4a5568",
            wraplength=820,
            justify=tk.LEFT,
        ).grid(row=2, column=1, columnspan=4, sticky=tk.W, pady=(6, 0))

        battery_frame = tk.LabelFrame(
            root,
            text="АКБ (заполняется перед выбором цепочки)",
            font=self._header_font,
            bg="#f7fafc",
            fg="#1a202c",
            padx=10,
            pady=8,
        )
        battery_frame.pack(fill=tk.X, padx=12, pady=(8, 4))

        fields = (
            ("Название", self._battery_name_var, 32),
            ("Количество элементов", self._battery_cells_count_var, 18),
            ("Количество АКБ", self._battery_count_var, 18),
            ("Rэл", self._r_el_var, 12),
            ("ρ (удельное сопротивление)", self._rho_var, 22),
            ("S (сечение перемычки)", self._jumper_section_var, 20),
        )
        for idx, (label, var, width) in enumerate(fields):
            row = idx // 3
            col = (idx % 3) * 2
            tk.Label(
                battery_frame,
                text=f"{label}:",
                font=self._cell_font,
                bg="#f7fafc",
            ).grid(row=row, column=col, sticky=tk.W, padx=(0, 6), pady=(2, 2))
            tk.Entry(
                battery_frame,
                textvariable=var,
                width=width,
                font=self._cell_font,
                relief=tk.FLAT,
            ).grid(row=row, column=col + 1, sticky=tk.W, padx=(0, 14), pady=(2, 2))

        fuse_frame = tk.LabelFrame(
            root,
            text="Вводной предохранитель и емкость",
            font=self._header_font,
            bg="#f7fafc",
            fg="#1a202c",
            padx=10,
            pady=8,
        )
        fuse_frame.pack(fill=tk.X, padx=12, pady=(4, 4))
        tk.Label(
            fuse_frame,
            text="Вводной предохранитель:",
            font=self._cell_font,
            bg="#f7fafc",
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 6), pady=(2, 2))
        fuse_values = [f"{label} ({value:g} Ом)" for label, value in _INPUT_FUSE_PRESETS] + ["Пользовательский"]
        self._input_fuse_combo = ttk.Combobox(
            fuse_frame,
            textvariable=self._input_fuse_pick_var,
            values=fuse_values,
            state="readonly",
            width=28,
            font=self._cell_font,
        )
        self._input_fuse_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 14), pady=(2, 2))
        self._input_fuse_combo.bind("<<ComboboxSelected>>", self._on_input_fuse_selected)
        if fuse_values:
            self._input_fuse_pick_var.set(fuse_values[0])

        tk.Label(
            fuse_frame,
            text="Rпр, Ом:",
            font=self._cell_font,
            bg="#f7fafc",
        ).grid(row=0, column=2, sticky=tk.W, padx=(0, 6), pady=(2, 2))
        tk.Entry(
            fuse_frame,
            textvariable=self._input_fuse_r_var,
            width=14,
            font=self._cell_font,
            relief=tk.FLAT,
        ).grid(row=0, column=3, sticky=tk.W, padx=(0, 14), pady=(2, 2))
        tk.Label(
            fuse_frame,
            text="QN=1, А·ч:",
            font=self._cell_font,
            bg="#f7fafc",
        ).grid(row=1, column=0, sticky=tk.W, padx=(0, 6), pady=(2, 2))
        tk.Entry(
            fuse_frame,
            textvariable=self._qn1_var,
            width=14,
            font=self._cell_font,
            relief=tk.FLAT,
        ).grid(row=1, column=1, sticky=tk.W, padx=(0, 14), pady=(2, 2))
        tk.Label(
            fuse_frame,
            text="Qрасч, А·ч:",
            font=self._cell_font,
            bg="#f7fafc",
        ).grid(row=1, column=2, sticky=tk.W, padx=(0, 6), pady=(2, 2))
        tk.Entry(
            fuse_frame,
            textvariable=self._q_calc_var,
            width=14,
            font=self._cell_font,
            relief=tk.FLAT,
        ).grid(row=1, column=3, sticky=tk.W, padx=(0, 14), pady=(2, 2))
        self._on_input_fuse_selected()

        toolbar = tk.Frame(root, bg="#f0f0f0")
        toolbar.pack(fill=tk.X, padx=12, pady=(8, 4))
        tk.Button(
            toolbar,
            text="+ Раздел",
            font=self._button_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#38a169",
            fg="white",
            activebackground="#2f855a",
            padx=14,
            command=lambda: self._add_section(),
        ).pack(side=tk.LEFT)

        scroll_outer = tk.Frame(root, bg="#cbd5e0", padx=1, pady=1)
        scroll_outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        canvas = tk.Canvas(scroll_outer, bg="#f0f0f0", highlightthickness=0)
        self._canvas = canvas
        scroll_y = tk.Scrollbar(scroll_outer, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = tk.Frame(canvas, bg="#f0f0f0")
        win_id = canvas.create_window((0, 0), window=inner, anchor=tk.NW)

        def sync_scroll(_event: tk.Event | None = None) -> None:
            del _event
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        def stretch(event: tk.Event) -> None:
            canvas.itemconfigure(win_id, width=event.width)

        inner.bind("<Configure>", sync_scroll)
        canvas.bind("<Configure>", stretch)

        self._sections_host = inner

        bottom = tk.Frame(root, bg="#f0f0f0")
        bottom.pack(fill=tk.X, padx=12, pady=(0, 12))
        tk.Button(
            bottom,
            text="Расчёт",
            font=self._button_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#2c5282",
            fg="white",
            activebackground="#2b6cb0",
            padx=16,
            command=self._run,
        ).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(
            bottom,
            text="В Word",
            font=self._button_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#553c9a",
            fg="white",
            activebackground="#6b46c1",
            padx=16,
            command=lambda: self._export(cast(tk.Misc, self._root)),
        ).pack(side=tk.LEFT)
        self._result_lbl = tk.Label(bottom, text="", font=self._cell_font, bg="#f0f0f0", fg="#2d3748")
        self._result_lbl.pack(side=tk.LEFT, padx=(16, 0))

        self._add_section("Раздел 1")
        sync_scroll()
        return root

    def _sync_scroll(self) -> None:
        if self._canvas is not None:
            self._canvas.update_idletasks()
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _set_entry(self, entry: tk.Entry, text: str) -> None:
        entry.delete(0, tk.END)
        if text:
            entry.insert(0, text)

    def _configure_equipment_table(self, table: tk.Frame) -> None:
        for col, width in enumerate(_EQUIPMENT_TABLE_COL_WIDTHS):
            table.grid_columnconfigure(col, minsize=width, weight=0)

    def _build_equipment_table_header(self, table: tk.Frame) -> None:
        if self._cell_font is None:
            return
        headers: tuple[tuple[int, str, str | None], ...] = (
            (_COL_TYPE, "Тип", None),
            (_COL_DESIGNATION, "Условное обозначение", None),
            (_COL_INOM, "Iном, А", None),
            (_COL_CURVE, "Хар-ка", None),
            (_COL_MULT, "Кратн.", None),
            (_COL_TIME, "t, с", None),
            (_COL_R, "R, Ом", None),
            (_COL_L, "L, м", None),
            (_COL_Y, "Y", None),
            (_COL_S, "S, мм²", None),
            (_COL_SEL, "Проверка на\nселективность", "small"),
            (_COL_DEL, "", None),
        )
        for col, text, kind in headers:
            font = self._table_header_small_font if kind == "small" else self._cell_font
            tk.Label(
                table,
                text=text,
                font=font,
                bg="#e2e8f0",
                anchor=tk.CENTER if col == _COL_SEL else tk.W,
                justify=tk.CENTER if col == _COL_SEL else tk.LEFT,
            ).grid(row=0, column=col, sticky="nsew", padx=2, pady=4)

    def _to_int(self, raw: str) -> int:
        text = raw.strip()
        if not text:
            return 0
        try:
            return int(text)
        except ValueError:
            return 0

    def _to_float(self, raw: str) -> float:
        text = raw.strip().replace(",", ".")
        if not text:
            return 0.0
        try:
            return float(text)
        except ValueError:
            return 0.0

    def _set_number_var(self, var: tk.StringVar, value: int | float) -> None:
        if isinstance(value, int):
            var.set("" if value == 0 else str(value))
            return
        var.set("" if value == 0 else f"{value:g}")

    def _collect_input(self) -> SoptInput:
        sections: list[SoptSection] = []
        for block in self._sections:
            subsections: list[SoptSubsection] = []
            for sub_block in block.subsections:
                items: list[SoptEquipmentItem] = []
                for row in sub_block.equipment:
                    type_text = row.type_label.cget("text")
                    key = next(
                        (k for k, label in EQUIPMENT_TYPES if label == type_text),
                        "fuse",
                    )
                    rated_current = self._to_float(row.rated_current_a.get())
                    resistance = self._to_float(row.resistance_ohm.get())
                    include_sel = (
                        row.include_in_selectivity is not None
                        and row.include_in_selectivity.get()
                    )
                    cb_curve = ""
                    cb_multiplier = 0.0
                    cb_trip_time_s = 0.0
                    if key == "breaker" and row.cb_curve is not None:
                        cb_curve = row.cb_curve.get().strip().upper()
                        if cb_curve not in self._breaker_curve_labels:
                            cb_curve = "C"
                        if row.cb_multiplier is not None:
                            cb_multiplier = self._to_float(row.cb_multiplier.get())
                        if row.cb_trip_time_s is not None:
                            cb_trip_time_s = self._to_float(row.cb_trip_time_s.get())
                    items.append(
                        SoptEquipmentItem(
                            item_type=key,
                            designation=row.designation.get().strip(),
                            rated_current_a=rated_current,
                            resistance_ohm=resistance,
                            cable_length_m=self._to_float(row.cable_length_m.get()),
                            cable_gamma=self._to_float(row.cable_gamma.get()),
                            cable_section_mm2=self._to_float(row.cable_section_mm2.get()),
                            cb_curve=cb_curve,
                            cb_multiplier=cb_multiplier,
                            cb_trip_time_s=cb_trip_time_s,
                            include_in_selectivity=include_sel if key == "kz_point" else False,
                        )
                    )
                subsections.append(
                    SoptSubsection(
                        name=sub_block.name_entry.get().strip(),
                        items=items,
                    )
                )
            sections.append(
                SoptSection(
                    name=block.name_entry.get().strip(),
                    subsections=subsections,
                )
            )
        return SoptInput(
            project_name=self._project_var.get().strip(),
            battery_name=self._battery_name_var.get().strip(),
            battery_cells_count=self._to_int(self._battery_cells_count_var.get()),
            battery_count=self._to_int(self._battery_count_var.get()),
            r_el=self._to_float(self._r_el_var.get()),
            rho=self._to_float(self._rho_var.get()),
            jumper_section=self._to_float(self._jumper_section_var.get()),
            input_fuse_label=self._input_fuse_pick_var.get().strip(),
            input_fuse_resistance=self._to_float(self._input_fuse_r_var.get()),
            qn1_ah=self._to_float(self._qn1_var.get()),
            q_calc_ah=self._to_float(self._q_calc_var.get()),
            sections=sections,
        )

    def _collect_snapshot(self) -> SoptProjectSnapshot:
        model = self._collect_input()
        return SoptProjectSnapshot(
            project_name=model.project_name,
            battery_name=model.battery_name,
            battery_cells_count=model.battery_cells_count,
            battery_count=model.battery_count,
            r_el=model.r_el,
            rho=model.rho,
            jumper_section=model.jumper_section,
            input_fuse_label=model.input_fuse_label,
            input_fuse_resistance=model.input_fuse_resistance,
            qn1_ah=model.qn1_ah,
            q_calc_ah=model.q_calc_ah,
            sections=model.sections,
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
            title="Открыть проект СОПТ",
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
            title="Сохранить проект СОПТ",
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

    def _clear_sections(self) -> None:
        for block in self._sections:
            block.frame.destroy()
        self._sections.clear()

    def _apply_snapshot(self, snap: SoptProjectSnapshot) -> None:
        self._project_var.set(snap.project_name)
        self._battery_name_var.set(snap.battery_name)
        self._set_number_var(self._battery_cells_count_var, snap.battery_cells_count)
        self._set_number_var(self._battery_count_var, snap.battery_count)
        self._set_number_var(self._r_el_var, snap.r_el)
        self._set_number_var(self._rho_var, snap.rho)
        self._set_number_var(self._jumper_section_var, snap.jumper_section)
        fuse_values = [f"{label} ({value:g} Ом)" for label, value in _INPUT_FUSE_PRESETS]
        if snap.input_fuse_label in fuse_values:
            self._input_fuse_pick_var.set(snap.input_fuse_label)
        else:
            self._input_fuse_pick_var.set(fuse_values[0] if fuse_values else snap.input_fuse_label)
        self._set_number_var(self._input_fuse_r_var, snap.input_fuse_resistance)
        self._set_number_var(self._qn1_var, snap.qn1_ah)
        self._set_number_var(self._q_calc_var, snap.q_calc_ah)
        if not self._input_fuse_r_var.get().strip():
            self._on_input_fuse_selected()
        self._clear_sections()
        if not snap.sections:
            self._add_section("Раздел 1")
        else:
            for section in snap.sections:
                block = self._add_section(section.name, create_default_subsection=False)
                if section.subsections:
                    for subsection in section.subsections:
                        sub_block = self._add_subsection(block, subsection.name)
                        for item in subsection.items:
                            self._add_equipment_row(
                                sub_block,
                                item.item_type,
                                item.designation,
                                item.rated_current_a,
                                item.resistance_ohm,
                                item.cable_length_m,
                                item.cable_gamma,
                                item.cable_section_mm2,
                                item.include_in_selectivity,
                                item.cb_curve,
                                item.cb_multiplier,
                                item.cb_trip_time_s,
                            )
                else:
                    self._add_subsection(block)
        self._sync_scroll()

    def _on_input_fuse_selected(self, _event: object | None = None) -> None:
        del _event
        selected = self._input_fuse_pick_var.get().strip()
        for label, value in _INPUT_FUSE_PRESETS:
            preset = f"{label} ({value:g} Ом)"
            if selected == preset:
                self._input_fuse_r_var.set(f"{value:g}")
                return

    def _add_section(self, name: str = "", *, create_default_subsection: bool = True) -> _SectionWidgets:
        if self._sections_host is None or self._cell_font is None:
            raise RuntimeError("UI not built")

        title = name.strip() or f"Раздел {len(self._sections) + 1}"
        lf = tk.LabelFrame(
            self._sections_host,
            text="",
            font=self._header_font,
            bg="#f7fafc",
            fg="#1a202c",
            padx=10,
            pady=8,
        )
        lf.pack(fill=tk.X, padx=4, pady=6)

        head = tk.Frame(lf, bg="#f7fafc")
        head.pack(fill=tk.X, pady=(0, 6))
        tk.Label(head, text="Название раздела:", font=self._cell_font, bg="#f7fafc").pack(side=tk.LEFT)
        name_entry = tk.Entry(head, font=self._cell_font, width=40, relief=tk.FLAT)
        name_entry.pack(side=tk.LEFT, padx=(8, 8), ipady=3)
        self._set_entry(name_entry, title)

        def refresh_title(_event: object | None = None) -> None:
            del _event
            text = name_entry.get().strip() or "Раздел"
            lf.configure(text=text)

        name_entry.bind("<KeyRelease>", refresh_title)
        refresh_title()

        delete_btn = tk.Button(
            head,
            text="Удалить раздел",
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#e53e3e",
            fg="white",
            activebackground="#c53030",
            padx=8,
            command=lambda: None,
        )
        delete_btn.pack(side=tk.RIGHT)

        sections_bar = tk.Frame(lf, bg="#f7fafc")
        sections_bar.pack(fill=tk.X, pady=(0, 6))

        subsections_frame = tk.Frame(lf, bg="#f7fafc")
        subsections_frame.pack(fill=tk.X)

        block = _SectionWidgets(
            frame=lf,
            name_entry=name_entry,
            subsections_frame=subsections_frame,
            subsections=[],
        )

        tk.Button(
            sections_bar,
            text="+ Подраздел",
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#2c5282",
            fg="white",
            activebackground="#2b6cb0",
            padx=10,
            command=lambda b=block: self._add_subsection(b),
        ).pack(side=tk.LEFT)

        def delete_section() -> None:
            if messagebox.askyesno(
                "Удалить раздел",
                f"Удалить раздел «{name_entry.get().strip() or 'Раздел'}»?",
                parent=self._root,
            ):
                block.frame.destroy()
                self._sections.remove(block)
                self._sync_scroll()

        delete_btn.configure(command=delete_section)

        self._sections.append(block)
        if create_default_subsection:
            self._add_subsection(block)
        self._sync_scroll()
        return block

    def _add_subsection(self, section_block: _SectionWidgets, name: str = "") -> _SubsectionWidgets:
        if self._cell_font is None:
            raise RuntimeError("UI not built")

        title = name.strip() or f"Подраздел {len(section_block.subsections) + 1}"
        lf = tk.LabelFrame(
            section_block.subsections_frame,
            text="",
            font=self._cell_font,
            bg="#edf2f7",
            fg="#1a202c",
            padx=8,
            pady=6,
        )
        lf.pack(fill=tk.X, pady=4)

        head = tk.Frame(lf, bg="#edf2f7")
        head.pack(fill=tk.X, pady=(0, 6))
        tk.Label(head, text="Название подраздела:", font=self._cell_font, bg="#edf2f7").pack(side=tk.LEFT)
        name_entry = tk.Entry(head, font=self._cell_font, width=36, relief=tk.FLAT)
        name_entry.pack(side=tk.LEFT, padx=(8, 8), ipady=3)
        self._set_entry(name_entry, title)

        def refresh_title(_event: object | None = None) -> None:
            del _event
            text = name_entry.get().strip() or "Подраздел"
            lf.configure(text=text)

        name_entry.bind("<KeyRelease>", refresh_title)
        refresh_title()

        equipment_table = tk.Frame(lf, bg="#edf2f7")
        equipment_table.pack(fill=tk.X)
        self._configure_equipment_table(equipment_table)
        self._build_equipment_table_header(equipment_table)

        add_bar = tk.Frame(lf, bg="#edf2f7")
        add_bar.pack(fill=tk.X, pady=(8, 0))
        type_combo = ttk.Combobox(
            add_bar,
            values=self._type_labels,
            state="readonly",
            width=32,
            font=self._cell_font,
        )
        type_combo.pack(side=tk.LEFT)
        if self._type_labels:
            type_combo.set(self._type_labels[0])

        sub_block = _SubsectionWidgets(
            frame=lf,
            name_entry=name_entry,
            equipment_table=equipment_table,
            next_equipment_row=1,
            type_combo=type_combo,
            equipment=[],
        )

        def delete_subsection() -> None:
            if messagebox.askyesno(
                "Удалить подраздел",
                f"Удалить подраздел «{name_entry.get().strip() or 'Подраздел'}»?",
                parent=self._root,
            ):
                sub_block.frame.destroy()
                section_block.subsections.remove(sub_block)
                self._sync_scroll()

        tk.Button(
            head,
            text="Удалить подраздел",
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#dd6b20",
            fg="white",
            activebackground="#c05621",
            padx=8,
            command=delete_subsection,
        ).pack(side=tk.RIGHT)

        tk.Button(
            add_bar,
            text="Добавить оборудование",
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#2c5282",
            fg="white",
            activebackground="#2b6cb0",
            padx=10,
            command=lambda: self._add_equipment_from_combo(sub_block),
        ).pack(side=tk.LEFT, padx=(8, 0))

        section_block.subsections.append(sub_block)
        self._sync_scroll()
        return sub_block

    def _add_equipment_from_combo(self, block: _SubsectionWidgets) -> None:
        label = block.type_combo.get().strip()
        if not label:
            return
        key = next((k for k, ru in EQUIPMENT_TYPES if ru == label), "fuse")
        self._add_equipment_row(block, key, "")
        self._sync_scroll()

    def _add_equipment_row(
        self,
        block: _SubsectionWidgets,
        item_type: str,
        designation: str,
        rated_current_a: float = 0.0,
        resistance_ohm: float = 0.0,
        cable_length_m: float = 0.0,
        cable_gamma: float = 0.0,
        cable_section_mm2: float = 0.0,
        include_in_selectivity: bool = False,
        cb_curve: str = "",
        cb_multiplier: float = 0.0,
        cb_trip_time_s: float = 0.0,
    ) -> None:
        if self._cell_font is None:
            return
        table = block.equipment_table
        row_idx = block.next_equipment_row
        block.next_equipment_row += 1
        table.grid_rowconfigure(row_idx, minsize=self._row_height_px)
        pad = dict(sticky="nsew", padx=2, pady=2)

        type_label = tk.Label(
            table,
            text=EQUIPMENT_LABEL_BY_KEY.get(item_type, item_type),
            font=self._cell_font,
            bg="#edf2f7",
            fg="#2d3748",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=_EQUIPMENT_TABLE_COL_WIDTHS[_COL_TYPE] - 12,
        )
        type_label.grid(row=row_idx, column=_COL_TYPE, sticky="nw", padx=4, pady=4)

        des_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        des_entry.grid(row=row_idx, column=_COL_DESIGNATION, **pad)
        self._set_entry(des_entry, designation)

        in_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        in_entry.grid(row=row_idx, column=_COL_INOM, **pad)
        self._set_entry(in_entry, "" if rated_current_a == 0 else f"{rated_current_a:g}")

        use_breaker = item_type == "breaker"
        curve_combo: ttk.Combobox | None = None
        mult_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        time_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        if use_breaker:
            curve_combo = ttk.Combobox(
                table,
                values=self._breaker_curve_labels,
                state="readonly",
                width=4,
                justify=tk.CENTER,
                style="Sopt.TCombobox",
            )
            curve_combo.grid(row=row_idx, column=_COL_CURVE, **pad)
            curve_key = (cb_curve or "C").strip().upper()
            if curve_key not in self._breaker_curve_labels:
                curve_key = "C"
            curve_combo.set(curve_key)
        else:
            curve_stub = tk.Entry(
                table,
                font=self._cell_font,
                relief=tk.FLAT,
                bg=self._disabled_field_bg,
                disabledbackground=self._disabled_field_bg,
                disabledforeground="#a0aec0",
                state=tk.DISABLED,
                justify=tk.CENTER,
            )
            curve_stub.grid(row=row_idx, column=_COL_CURVE, **pad)

        mult_entry.grid(row=row_idx, column=_COL_MULT, **pad)
        self._set_entry(mult_entry, "" if cb_multiplier == 0 else f"{cb_multiplier:g}")

        time_entry.grid(row=row_idx, column=_COL_TIME, **pad)
        self._set_entry(time_entry, "" if cb_trip_time_s == 0 else f"{cb_trip_time_s:g}")

        r_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        r_entry.grid(row=row_idx, column=_COL_R, **pad)
        self._set_entry(r_entry, "" if resistance_ohm == 0 else f"{resistance_ohm:g}")

        l_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        l_entry.grid(row=row_idx, column=_COL_L, **pad)
        self._set_entry(l_entry, "" if cable_length_m == 0 else f"{cable_length_m:g}")

        y_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        y_entry.grid(row=row_idx, column=_COL_Y, **pad)
        self._set_entry(y_entry, "" if cable_gamma == 0 else f"{cable_gamma:g}")

        s_entry = tk.Entry(table, font=self._cell_font, relief=tk.FLAT, bg="white")
        s_entry.grid(row=row_idx, column=_COL_S, **pad)
        self._set_entry(s_entry, "" if cable_section_mm2 == 0 else f"{cable_section_mm2:g}")

        use_device = item_type in ("fuse", "breaker")
        use_cable = item_type == "cable"

        if not use_device:
            in_entry.delete(0, tk.END)
        in_entry.configure(state=tk.NORMAL if use_device else tk.DISABLED)

        if mult_entry is not None:
            if not use_breaker:
                mult_entry.delete(0, tk.END)
            mult_entry.configure(
                state=tk.NORMAL if use_breaker else tk.DISABLED,
                bg="white" if use_breaker else self._disabled_field_bg,
                disabledbackground=self._disabled_field_bg,
                disabledforeground="#a0aec0",
            )
        if time_entry is not None:
            if not use_breaker:
                time_entry.delete(0, tk.END)
            time_entry.configure(
                state=tk.NORMAL if use_breaker else tk.DISABLED,
                bg="white" if use_breaker else self._disabled_field_bg,
                disabledbackground=self._disabled_field_bg,
                disabledforeground="#a0aec0",
            )

        if not use_device:
            r_entry.delete(0, tk.END)
        r_entry.configure(state=tk.NORMAL if use_device else tk.DISABLED)
        if not use_cable:
            l_entry.delete(0, tk.END)
            y_entry.delete(0, tk.END)
            s_entry.delete(0, tk.END)
        l_entry.configure(state=tk.NORMAL if use_cable else tk.DISABLED)
        y_entry.configure(state=tk.NORMAL if use_cable else tk.DISABLED)
        s_entry.configure(state=tk.NORMAL if use_cable else tk.DISABLED)

        sel_var: tk.BooleanVar | None = None
        if item_type == "kz_point":
            sel_var = tk.BooleanVar(value=include_in_selectivity)
            tk.Checkbutton(
                table,
                variable=sel_var,
                bg="#edf2f7",
                activebackground="#edf2f7",
                relief=tk.FLAT,
            ).grid(row=row_idx, column=_COL_SEL, sticky="n")
        else:
            tk.Label(table, text="", bg="#edf2f7").grid(row=row_idx, column=_COL_SEL, sticky="nsew")

        widgets = _EquipmentRowWidgets(
            type_label=type_label,
            designation=des_entry,
            rated_current_a=in_entry,
            cb_curve=curve_combo,
            cb_multiplier=mult_entry,
            cb_trip_time_s=time_entry,
            resistance_ohm=r_entry,
            cable_length_m=l_entry,
            cable_gamma=y_entry,
            cable_section_mm2=s_entry,
            include_in_selectivity=sel_var,
            row_frame=table,
            item_type=item_type,
        )

        def remove_row() -> None:
            for child in table.grid_slaves(row=row_idx):
                child.destroy()
            block.equipment.remove(widgets)
            self._sync_scroll()

        tk.Button(
            table,
            text="×",
            font=self._cell_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#fed7d7",
            fg="#9b2c2c",
            width=2,
            command=remove_row,
        ).grid(row=row_idx, column=_COL_DEL, sticky="nsew", padx=1, pady=2)

        block.equipment.append(widgets)

    def _run(self) -> None:
        model = self._collect_input()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning(
                "Проверка данных",
                "\n".join(err.message for err in errors[:10]),
                parent=self._root,
            )
            return
        result = self._calculator.calculate(model)
        if self._result_lbl is not None:
            self._result_lbl.configure(text=result.message)

    def _export(self, parent: tk.Misc) -> None:
        model = self._collect_input()
        errors = self._validator.validate(model)
        if errors:
            messagebox.showwarning(
                "Проверка данных",
                "\n".join(err.message for err in errors[:10]),
                parent=self._root,
            )
            return
        result = self._calculator.calculate(model)
        if self._result_lbl is not None:
            self._result_lbl.configure(text=result.message)
        self._exporter.export(parent, model, result)
