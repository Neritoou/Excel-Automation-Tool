import tkinter as tk
from tkinter import ttk
import threading
from typing import Any

from tasks import ALL_TASKS
from core.base_task import BaseTask, TaskResult
from core.exceptions import ValidationTaskError
from gui.task_frame import TaskFrame

class App(tk.Tk):
    """Ventana principal de la aplicación Excel Automator."""
    # --- PALETA DE COLORES ---
    BG_SIDEBAR:        str = "#1e293b"
    BG_SIDEBAR_HOVER:  str = "#334155"
    BG_SIDEBAR_ACTIVE: str = "#3b82f6"
    FG_SIDEBAR:        str = "#cbd5e1"
    FG_SIDEBAR_ACTIVE: str = "#ffffff"
    BG_MAIN:           str = "#f8fafc"
    BG_STATUSBAR:      str = "#e2e8f0"
    ACCENT:            str = "#3b82f6"
    SUCCESS:           str = "#16a34a"
    ERROR:             str = "#dc2626"

    VERSION: str = "v2.0"

    def __init__(self) -> None:
        super().__init__()
        self.title("TAREAS DE EXCEL")
        self.geometry("1000x680")
        self.minsize(820, 520)
        self.configure(bg=self.BG_MAIN)

        self._setup_styles()

        self.tasks = {t.task_id: t() for t in ALL_TASKS}
        self._active_task:      BaseTask | None  = None
        self._task_frame:       TaskFrame | None = None
        self._sidebar_buttons:  dict[str, tk.Button] = {}
        self._is_running:       bool             = False
        self._exec_btn:         ttk.Button | None = None
        self._log:              tk.Text | None   = None
        self._status_lbl:       ttk.Label | None = None

        self._build_layout()

    def _setup_styles(self) -> None:
        """Configura los estilos globales de ttk."""
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",           background=self.BG_MAIN)
        s.configure("TLabel",           background=self.BG_MAIN,      font=("Segoe UI", 10))
        s.configure("TLabelframe",      background=self.BG_MAIN,      font=("Segoe UI", 10, "bold"))
        s.configure("TLabelframe.Label",background=self.BG_MAIN,      foreground="#334155")
        s.configure("Header.TLabel",    font=("Segoe UI", 16, "bold"), foreground="#0f172a")
        s.configure("Desc.TLabel",      font=("Segoe UI", 10),         foreground="#64748b")
        s.configure("Status.TLabel",    background=self.BG_STATUSBAR,  font=("Segoe UI", 9), foreground="#475569")
        s.configure("Exec.TButton",     font=("Segoe UI", 11, "bold"), padding=(20, 10))
        s.configure("TCombobox",        padding=4)
        s.configure("TEntry",           padding=4)

    def _build_layout(self) -> None:
        """Construye la estructura principal: statusbar, sidebar y contenido."""
        self._build_statusbar()

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(main, bg=self.BG_SIDEBAR, width=230)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        self.content = ttk.Frame(main)
        self.content.pack(side="left", fill="both", expand=True)

        if self.tasks:
            self._select_task(next(iter(self.tasks.values())))

    def _build_statusbar(self) -> None:
        """Construye la barra de estado inferior."""
        sb = tk.Frame(self, bg=self.BG_STATUSBAR, height=28)
        sb.pack(side="bottom", fill="x")
        sb.pack_propagate(False)
        self._status_lbl = ttk.Label(sb, text="  Listo", style="Status.TLabel")
        self._status_lbl.pack(side="left", padx=(8, 0), fill="y")
        ttk.Label(sb, text=f"{self.VERSION}  ", style="Status.TLabel").pack(side="right", padx=(0, 12), fill="y")

    def _build_sidebar(self) -> None:
        """Construye el panel lateral con logo y botones de tareas."""
        lf = tk.Frame(self.sidebar, bg=self.BG_SIDEBAR)
        lf.pack(fill="x", padx=16, pady=(20, 8))

        tk.Label(
            lf, text="📊 TAREAS DE EXCEL", font=("Segoe UI", 13, "bold"),
            bg=self.BG_SIDEBAR, fg="#fff", anchor="w",
        ).pack(fill="x")

        tk.Label(
            lf, text="Herramientas de automatización",
            font=("Segoe UI", 8), bg=self.BG_SIDEBAR, fg="#94a3b8", anchor="w",
        ).pack(fill="x")

        tk.Frame(self.sidebar, bg="#334155", height=1).pack(fill="x", padx=16, pady=(12, 12))

        tk.Label(
            self.sidebar, text="TAREAS", font=("Segoe UI", 8, "bold"),
            bg=self.BG_SIDEBAR, fg="#64748b", anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 4))

        for tid, task in self.tasks.items():
            btn = tk.Button(
                self.sidebar, text=f"  {task.task_icon}  {task.task_name}",
                font=("Segoe UI", 10), bg=self.BG_SIDEBAR, fg=self.FG_SIDEBAR,
                activebackground=self.BG_SIDEBAR_HOVER, activeforeground=self.FG_SIDEBAR_ACTIVE,
                bd=0, anchor="w", padx=16, pady=8, cursor="hand2",
                command=lambda t=task: self._select_task(t),
            )
            btn.pack(fill="x", padx=8, pady=1)
            btn.bind("<Enter>", lambda e, b=btn: self._on_hover(b, True))
            btn.bind("<Leave>", lambda e, b=btn: self._on_hover(b, False))
            self._sidebar_buttons[tid] = btn

    def _select_task(self, task: BaseTask) -> None:
        """
        Carga la tarea seleccionada en el panel de contenido.

        Genera dinámicamente: encabezado, formulario de parámetros,
        barra de progreso, botón de ejecución y log.
        """
        if self._is_running:
            return
        self._active_task = task

        for tid, btn in self._sidebar_buttons.items():
            is_active = tid == task.task_id
            btn.configure(
                bg=self.BG_SIDEBAR_ACTIVE if is_active else self.BG_SIDEBAR,
                fg=self.FG_SIDEBAR_ACTIVE if is_active else self.FG_SIDEBAR,
            )

        for w in self.content.winfo_children():
            w.destroy()

        canvas, sf = self._build_scrollable_content()

        hdr = ttk.Frame(sf)
        hdr.pack(fill="x", padx=28, pady=(24, 4))
        ttk.Label(hdr, text=f"{task.task_icon}  {task.task_name}", style="Header.TLabel").pack(anchor="w")
        ttk.Label(hdr, text=task.task_description, style="Desc.TLabel", wraplength=620).pack(anchor="w", pady=(4, 0))
        ttk.Separator(sf).pack(fill="x", padx=28, pady=(12, 4))

        self._task_frame = TaskFrame(sf, task)
        self._task_frame.pack(fill="x", padx=28, pady=(4, 8))

        bf = ttk.Frame(sf)
        bf.pack(fill="x", padx=28, pady=(12, 8))
        self._exec_btn = ttk.Button(bf, text="▶  Ejecutar", style="Exec.TButton", command=self._execute)
        self._exec_btn.pack(side="right")
        ttk.Button(bf, text="Limpiar log", command=self._clear_log).pack(side="right", padx=(0, 8))

        self._build_log(sf)
        self._set_status(f"Tarea: {task.task_name}")

    def _build_scrollable_content(self) -> tuple[tk.Canvas, ttk.Frame]:
        """Crea un canvas con scroll que no interfiere con widgets internos."""
        canvas = tk.Canvas(self.content, bg=self.BG_MAIN, highlightthickness=0)
        vsb = ttk.Scrollbar(self.content, orient="vertical", command=canvas.yview)
        sf = ttk.Frame(canvas)

        sf.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sf, anchor="nw", tags="inner")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig("inner", width=e.width))

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        scroll_sequences = ("<MouseWheel>", "<Button-4>", "<Button-5>")

        def _wheel(event: "tk.Event[Any]") -> str:
            """Scroll vertical, retorna 'break' para evitar propagación."""
            direction = -1 if (event.num == 4 or event.delta > 0) else 1
            canvas.yview_scroll(direction * 3, "units")
            return "break"

        def _bind_scroll(_event: "tk.Event[Any]") -> None:
            """Activa scroll cuando el mouse entra al canvas."""
            for seq in scroll_sequences:
                canvas.bind_all(seq, _wheel)

        def _unbind_scroll(_event: "tk.Event[Any]") -> None:
            """Desactiva scroll cuando el mouse sale del canvas."""
            for seq in scroll_sequences:
                canvas.unbind_all(seq)

        canvas.bind("<Enter>", _bind_scroll)
        canvas.bind("<Leave>", _unbind_scroll)

        return canvas, sf

    def _build_log(self, parent: ttk.Frame) -> None:
        """Construye el widget de log de ejecuciones."""
        ttk.Label(parent, text="Historial de ejecuciones", style="Desc.TLabel").pack(
            anchor="w", padx=32, pady=(8, 2)
        )
        self._log = tk.Text(
            parent, height=10, wrap="word", font=("Consolas", 9),
            bg="#f1f5f9", fg="#334155", relief="flat", padx=12, pady=8,
            state="disabled", selectbackground=self.ACCENT, selectforeground="#fff",
        )
        self._log.pack(fill="x", padx=28, pady=(0, 24))
        self._log.tag_configure("ok",   foreground=self.SUCCESS)
        self._log.tag_configure("err",  foreground=self.ERROR)
        self._log.tag_configure("file", foreground=self.ACCENT)

    def _execute(self) -> None:
        """Valida parámetros y lanza la tarea en un hilo secundario."""
        if self._is_running or self._active_task is None or self._task_frame is None:
            return

        params = self._task_frame.collect_params()
        try:
            self._active_task.validate(params)
        except ValidationTaskError as e:
            self._log_msg(f"Validación fallida: {e}", err=True)
            return

        self._set_running(True)
        current_task = self._active_task

        def worker() -> None:
            result = current_task.execute(params)
            self.after(0, self._on_done, result)

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, result: TaskResult) -> None:
        """Callback al completarse la ejecución de la tarea."""
        self._set_running(False)
        if result.success:
            self._log_msg(result.message, files=result.output_files)
            self._set_status("Tarea completada")
        else:
            self._log_msg(result.message, err=True)
            self._set_status("Finalizó con errores")

    # --- HELPERS ---

    def _set_running(self, on: bool) -> None:
        """Bloquea/desbloquea controles durante la ejecución."""
        self._is_running = on
        state = "disabled" if on else "normal"
        if self._exec_btn is not None:
            self._exec_btn.config(state=state, text="⏳ Ejecutando..." if on else "▶  Ejecutar")
        for btn in self._sidebar_buttons.values():
            btn.config(state=state)

    def _log_msg(self, msg: str, err: bool = False, files: list[str] | None = None) -> None:
        """Escribe un mensaje en el log con formato."""
        if self._log is None:
            return
        self._log.config(state="normal")
        icon, tag = ("❌", "err") if err else ("✅", "ok")
        self._log.insert("end", f"{icon} {msg}\n", tag)
        if files:
            for f in files:
                self._log.insert("end", f"   📁 {f}\n", "file")
        self._log.insert("end", "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self) -> None:
        """Limpia el historial de ejecuciones."""
        if self._log is None:
            return
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _set_status(self, text: str) -> None:
        """Actualiza la barra de estado inferior."""
        if self._status_lbl is not None:
            self._status_lbl.config(text=f"  {text}")

    def _on_hover(self, btn: tk.Button, enter: bool) -> None:
        """Efecto hover en botones del sidebar."""
        if btn.cget("bg") == self.BG_SIDEBAR_ACTIVE or str(btn.cget("state")) == "disabled":
            return
        btn.configure(bg=self.BG_SIDEBAR_HOVER if enter else self.BG_SIDEBAR)