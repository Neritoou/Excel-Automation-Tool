import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, TYPE_CHECKING

from core.base_task import ParamType, TaskParam
from core.excel_handler import ExcelHandler

if TYPE_CHECKING:
    from core.base_task import BaseTask


class TaskFrame(ttk.Frame):
    """Frame dinámico que renderiza los parámetros de una tarea."""
    def __init__(self, parent: tk.Widget, task: "BaseTask", **kwargs: Any) -> None:
        super().__init__(parent, **kwargs)
        self.task = task
        self.vars: dict[str, tk.Variable] = {}
        self._file_paths: dict[str, str] = {}
        self._sheet_combos: dict[str, ttk.Combobox] = {}
        self._col_combos: dict[str, ttk.Combobox] = {}
        self._param_map: dict[str, TaskParam] = {}
        self._current_container: ttk.LabelFrame | ttk.Frame = self

        self._builders = {
            ParamType.FILE: self._build_file_picker,
            ParamType.SHEET: self._build_sheet_combo,
            ParamType.COLUMN: self._build_column_combo,
            ParamType.TEXT: self._build_text_entry,
            ParamType.NUMBER: self._build_number_entry,
            ParamType.BOOL: self._build_checkbox,
            ParamType.SELECT: self._build_select,

        }

        self._build()

    def _build(self) -> None:
        """Genera los widgets para cada parámetro declarado por la tarea."""
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

            ttk.Label(row_frame, text=p.label, width=22, anchor="w").pack(
                side="left", padx=(0, 8)
            )

            builder = self._builders.get(p.param_type)
            if builder:
                builder(row_frame, p)
        
    # --- BUILDERS ---

    def _build_file_picker(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un campo de texto readonly + botón para seleccionar archivo."""
        var = tk.StringVar(value="Ningún archivo seleccionado")
        self.vars[p.name] = var

        ttk.Entry(parent, textvariable=var, state="readonly", width=45).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(
            parent,
            text="📂 Buscar",
            width=10,
            command=lambda: self._on_file_select(p.name),
        ).pack(side="left", padx=(6, 0))

    def _build_sheet_combo(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un combobox para seleccionar hoja de un archivo Excel."""
        var = tk.StringVar()
        self.vars[p.name] = var
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", width=42)
        combo.pack(side="left", fill="x", expand=True)
        self._sheet_combos[p.name] = combo

        if p.depends_on is not None:
            combo.bind("<<ComboboxSelected>>", lambda _: self._on_sheet_select(p.name))

    def _build_column_combo(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un combobox para seleccionar columna de una hoja."""
        var = tk.StringVar()
        self.vars[p.name] = var
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", width=42)
        combo.pack(side="left", fill="x", expand=True)
        self._col_combos[p.name] = combo

    def _build_text_entry(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un campo de texto libre."""
        var = tk.StringVar(value=str(p.default) if p.default is not None else "")
        self.vars[p.name] = var
        ttk.Entry(parent, textvariable=var, width=45).pack(
            side="left", fill="x", expand=True
        )

    def _build_number_entry(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un spinbox numérico."""
        var = tk.IntVar(value=int(p.default) if p.default is not None else 0)
        self.vars[p.name] = var
        ttk.Spinbox(parent, textvariable=var, from_=0, to=999999, width=20).pack(
            side="left"
        )

    def _build_checkbox(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un checkbox booleano."""
        var = tk.BooleanVar(value=bool(p.default) if p.default is not None else False)
        self.vars[p.name] = var
        ttk.Checkbutton(parent, variable=var).pack(side="left")

    def _build_select(self, parent: ttk.Frame, p: TaskParam) -> None:
        """Crea un combobox con opciones fijas definidas en TaskParam.options."""
        default = str(p.default) if p.default is not None else ""
        var = tk.StringVar(value=default)
        self.vars[p.name] = var
        combo = ttk.Combobox(parent, textvariable=var, state="readonly",
                             values=list(p.options), width=42)
        combo.pack(side="left", fill="x", expand=True)
        if default and default in p.options:
            combo.set(default)
        elif p.options:
            combo.current(0)
    
    # --- CALLBACKS ---

    def _on_file_select(self, param_name: str) -> None:
        """Abre diálogo de archivo y actualiza combos de hojas dependientes."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")],
        )
        if not path:
            return

        resolved_path = str(path)
        self._file_paths[param_name] = resolved_path
        display = resolved_path if len(resolved_path) < 55 else f"...{resolved_path[-52:]}"
        self.vars[param_name].set(display)

        # Actualizar combos de hojas que dependen de este archivo
        for name, p in self._param_map.items():
            if p.param_type == ParamType.SHEET and p.depends_on == param_name:
                self._load_sheets(name, resolved_path)

    def _load_sheets(self, sheet_param_name: str, file_path: str) -> None:
        """Carga las hojas de un archivo Excel en el combobox correspondiente."""
        try:
            sheets = ExcelHandler.get_sheet_names(file_path)
            combo = self._sheet_combos[sheet_param_name]
            combo["values"] = sheets
            if sheets:
                combo.current(0)
                self._on_sheet_select(sheet_param_name)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def _on_sheet_select(self, sheet_param_name: str) -> None:
        """Cuando se selecciona una hoja, actualiza combos de columna dependientes."""
        sheet_param = self._param_map[sheet_param_name]
        file_param_name = sheet_param.depends_on
        if file_param_name is None:
            return

        file_path = self._file_paths.get(file_param_name)
        if not file_path:
            return

        sheet_name = str(self.vars[sheet_param_name].get())
        if not sheet_name:
            return

        for name, p in self._param_map.items():
            if p.param_type == ParamType.COLUMN and p.depends_on == sheet_param_name:
                self._load_columns(name, file_path, sheet_name)

    def _load_columns(self, col_param_name: str, file_path: str, sheet_name: str) -> None:
        """Carga los encabezados de columna en el combobox correspondiente."""
        try:
            headers = ExcelHandler.get_column_headers(file_path, sheet_name)
            
            display = [
                f"{i + 1}: {h}"
                for i, h in enumerate(headers)
                if h and str(h).strip()
            ]

            combo = self._col_combos[col_param_name]
            combo["values"] = display
            if display:
                combo.current(0)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer las columnas:\n{e}")

    def collect_params(self) -> dict[str, Any]:
        """Devuelve un dict con los valores listos para pasar a task.execute()."""
        result: dict[str, Any] = {}
        for name, p in self._param_map.items():
            if p.param_type == ParamType.FILE:
                result[name] = self._file_paths.get(name, "")
            elif p.param_type == ParamType.COLUMN:
                raw = str(self.vars[name].get())
                try:
                    result[name] = int(raw.split(":")[0])
                except (ValueError, IndexError):
                    result[name] = None
            elif p.param_type in (ParamType.BOOL, ParamType.NUMBER):
                result[name] = self.vars[name].get()
            else:
                result[name] = str(self.vars[name].get())
        return result