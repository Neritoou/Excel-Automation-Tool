"""
Constructor dinámico de formularios para tareas.

Lee los TaskParam de cualquier BaseTask y genera automáticamente
los widgets de tkinter correspondientes (file chooser, combo box, etc.).
Esto permite que añadir una nueva tarea NO requiera tocar la GUI.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any

from core.base_task import BaseTask, TaskParam
from core.excel_handler import ExcelHandler


class TaskFrame(ttk.Frame):
    """Frame dinámico que renderiza los parámetros de una tarea."""

    def __init__(self, parent: tk.Widget, task: BaseTask, **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)
        self.task = task
        self.widgets: dict[str, Any] = {}
        self.vars: dict[str, tk.Variable] = {}
        self._file_paths: dict[str, str] = {}
        self._sheet_combos: dict[str, ttk.Combobox] = {}
        self._col_combos: dict[str, ttk.Combobox] = {}
        self._param_map: dict[str, TaskParam] = {}
        self._current_container: ttk.LabelFrame | ttk.Frame = self

        self._build()

    # ── Construcción del formulario ──────────────────────────────────
    def _build(self) -> None:
        params = self.task.get_params()
        current_group: str | None = None

        for p in params:
            self._param_map[p.name] = p

            if p.group != current_group:
                current_group = p.group
                grp_frame = ttk.LabelFrame(self, text=f"  {current_group}  ", padding=12)
                grp_frame.pack(fill="x", padx=8, pady=(10, 2))
                self._current_container = grp_frame

            row_frame = ttk.Frame(self._current_container)
            row_frame.pack(fill="x", pady=4)

            lbl = ttk.Label(row_frame, text=p.label, width=22, anchor="w")
            lbl.pack(side="left", padx=(0, 8))

            if p.param_type == "file":
                self._build_file_picker(row_frame, p)
            elif p.param_type == "sheet":
                self._build_sheet_combo(row_frame, p)
            elif p.param_type == "column":
                self._build_column_combo(row_frame, p)
            elif p.param_type == "text":
                self._build_text_entry(row_frame, p)
            elif p.param_type == "number":
                self._build_number_entry(row_frame, p)
            elif p.param_type == "bool":
                self._build_checkbox(row_frame, p)

    # ── Builders para cada tipo de widget ────────────────────────────
    def _build_file_picker(self, parent: ttk.Frame, p: TaskParam) -> None:
        var = tk.StringVar(value="Ningún archivo seleccionado")
        self.vars[p.name] = var

        entry = ttk.Entry(parent, textvariable=var, state="readonly", width=45)
        entry.pack(side="left", fill="x", expand=True)

        btn = ttk.Button(
            parent, text="📂 Buscar", width=10,
            command=lambda: self._on_file_select(p.name)
        )
        btn.pack(side="left", padx=(6, 0))

    def _build_sheet_combo(self, parent: ttk.Frame, p: TaskParam) -> None:
        var = tk.StringVar()
        self.vars[p.name] = var
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", width=42)
        combo.pack(side="left", fill="x", expand=True)
        self._sheet_combos[p.name] = combo

        if p.depends_on is not None:
            combo.bind("<<ComboboxSelected>>", lambda _e: self._on_sheet_select(p.name))

    def _build_column_combo(self, parent: ttk.Frame, p: TaskParam) -> None:
        var = tk.StringVar()
        self.vars[p.name] = var
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", width=42)
        combo.pack(side="left", fill="x", expand=True)
        self._col_combos[p.name] = combo

    def _build_text_entry(self, parent: ttk.Frame, p: TaskParam) -> None:
        default_val: str = str(p.default) if p.default is not None else ""
        var = tk.StringVar(value=default_val)
        self.vars[p.name] = var
        entry = ttk.Entry(parent, textvariable=var, width=45)
        entry.pack(side="left", fill="x", expand=True)

    def _build_number_entry(self, parent: ttk.Frame, p: TaskParam) -> None:
        default_val: int = int(p.default) if p.default is not None else 0
        var = tk.IntVar(value=default_val)
        self.vars[p.name] = var
        spin = ttk.Spinbox(parent, textvariable=var, from_=0, to=999999, width=20)
        spin.pack(side="left")

    def _build_checkbox(self, parent: ttk.Frame, p: TaskParam) -> None:
        default_val: bool = bool(p.default) if p.default is not None else False
        var = tk.BooleanVar(value=default_val)
        self.vars[p.name] = var
        chk = ttk.Checkbutton(parent, variable=var)
        chk.pack(side="left")

    # ── Callbacks de interacción ─────────────────────────────────────
    def _on_file_select(self, param_name: str) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")]
        )
        if not path:
            return

        # askopenfilename puede devolver str o tuple; forzar str
        resolved_path: str = str(path)
        self._file_paths[param_name] = resolved_path
        display = resolved_path if len(resolved_path) < 55 else "..." + resolved_path[-52:]
        self.vars[param_name].set(display)

        for name, p in self._param_map.items():
            if p.param_type == "sheet" and p.depends_on == param_name:
                try:
                    sheets = ExcelHandler.get_sheet_names(resolved_path)
                    combo = self._sheet_combos[name]
                    combo["values"] = sheets
                    if sheets:
                        combo.current(0)
                        self._on_sheet_select(name)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def _on_sheet_select(self, sheet_param_name: str) -> None:
        """Cuando se selecciona una hoja, actualizar combos de columna dependientes."""
        sheet_param = self._param_map[sheet_param_name]
        file_param_name = sheet_param.depends_on

        if file_param_name is None:
            return

        file_path = self._file_paths.get(file_param_name)
        if file_path is None:
            return

        sheet_name: str = str(self.vars[sheet_param_name].get())
        if not sheet_name:
            return

        for name, p in self._param_map.items():
            if p.param_type == "column" and p.depends_on == sheet_param_name:
                try:
                    headers = ExcelHandler.get_column_headers(file_path, sheet_name)
                    display_headers = [
                        f"{i+1}: {h}" if h else f"{i+1}: (vacía)"
                        for i, h in enumerate(headers)
                    ]
                    combo = self._col_combos[name]
                    combo["values"] = display_headers
                    if display_headers:
                        combo.current(0)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo leer las columnas:\n{e}")

    # ── Recolección de valores ───────────────────────────────────────
    def collect_params(self) -> dict[str, Any]:
        """Devuelve un dict con los valores listos para pasar a task.execute()."""
        result: dict[str, Any] = {}
        for name, p in self._param_map.items():
            if p.param_type == "file":
                result[name] = self._file_paths.get(name, "")
            elif p.param_type == "column":
                raw = str(self.vars[name].get())
                try:
                    result[name] = int(raw.split(":")[0])
                except (ValueError, IndexError):
                    result[name] = None
            elif p.param_type in ("bool", "number"):
                result[name] = self.vars[name].get()
            else:
                result[name] = str(self.vars[name].get())
        return result
