from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont

from app.core.contracts import CalculationPlugin
from app.core.registry import get_calculation_plugins


class MainWindow:
    def __init__(self, title: str) -> None:
        self._root = tk.Tk()
        self._root.title(title)
        self._root.geometry("1700x620")
        self._root.minsize(920, 440)
        self._root.configure(bg="#f0f0f0")

        self._plugins: list[CalculationPlugin] = get_calculation_plugins()
        self._cached_section_roots: dict[str, tk.Widget] = {}

        self._outer: tk.Frame | None = None
        self._home_view: tk.Frame | None = None
        self._section_view: tk.Frame | None = None
        self._section_inner: tk.Frame | None = None
        self._section_title_lbl: tk.Label | None = None

        try:
            self._title_font = tkfont.Font(family="Segoe UI", size=32, weight="bold")
            self._button_font = tkfont.Font(family="Segoe UI", size=13)
            self._section_header_font = tkfont.Font(family="Segoe UI", size=13)
        except tk.TclError:
            self._title_font = tkfont.Font(size=32, weight="bold")
            self._button_font = tkfont.Font(size=13)
            self._section_header_font = tkfont.Font(size=13)

    def build(self) -> None:
        outer = tk.Frame(self._root, bg="#f0f0f0")
        self._outer = outer
        outer.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        home = tk.Frame(outer, bg="#f0f0f0")
        self._home_view = home
        self._build_home(home)

        section = tk.Frame(outer, bg="#f0f0f0")
        self._section_view = section

        top_bar = tk.Frame(section, bg="#f0f0f0")
        top_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 8))
        tk.Button(
            top_bar,
            text="← На главную",
            font=self._button_font,
            cursor="hand2",
            relief=tk.FLAT,
            bg="#e2e8f0",
            fg="#1a202c",
            activebackground="#cbd5e0",
            padx=14,
            pady=6,
            command=self._show_home,
        ).pack(side=tk.LEFT)
        self._section_title_lbl = tk.Label(
            top_bar,
            text="",
            font=self._section_header_font,
            bg="#f0f0f0",
            fg="#1a1a1a",
        )
        self._section_title_lbl.pack(side=tk.LEFT, padx=(20, 0))

        inner = tk.Frame(section, bg="#f0f0f0")
        self._section_inner = inner
        inner.pack(fill=tk.BOTH, expand=True)

        home.pack(fill=tk.BOTH, expand=True)

    def _build_home(self, parent: tk.Frame) -> None:
        title = tk.Label(
            parent,
            text="Электропроект",
            font=self._title_font,
            bg="#f0f0f0",
            fg="#1a1a1a",
        )
        title.pack(side=tk.TOP, pady=(28, 12))

        middle = tk.Frame(parent, bg="#f0f0f0")
        middle.pack(fill=tk.BOTH, expand=True)
        btn_col = tk.Frame(middle, bg="#f0f0f0")
        btn_col.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        for plugin in self._plugins:
            b = tk.Button(
                btn_col,
                text=plugin.title,
                font=self._button_font,
                width=28,
                height=2,
                cursor="hand2",
                relief=tk.FLAT,
                bg="#2c5282",
                fg="white",
                activebackground="#2b6cb0",
                activeforeground="white",
                command=lambda p=plugin: self._show_section(p),
            )
            b.pack(pady=8)

    def _ensure_section_root(self, plugin: CalculationPlugin) -> tk.Widget:
        if plugin.id in self._cached_section_roots:
            return self._cached_section_roots[plugin.id]
        if self._section_inner is None:
            raise RuntimeError("MainWindow.build must run before opening a section")
        calc = plugin.widget_factory()
        root = calc.build(self._section_inner)
        self._cached_section_roots[plugin.id] = root
        return root

    def _show_section(self, plugin: CalculationPlugin) -> None:
        if self._home_view is None or self._section_view is None or self._section_inner is None:
            return
        root = self._ensure_section_root(plugin)
        for child in self._section_inner.winfo_children():
            child.pack_forget()
        root.pack(fill=tk.BOTH, expand=True)
        if self._section_title_lbl is not None:
            self._section_title_lbl.configure(text=plugin.title)
        self._home_view.pack_forget()
        self._section_view.pack(fill=tk.BOTH, expand=True)

    def _show_home(self) -> None:
        if self._home_view is None or self._section_view is None:
            return
        self._section_view.pack_forget()
        self._home_view.pack(fill=tk.BOTH, expand=True)

    def run(self) -> None:
        self._root.mainloop()
