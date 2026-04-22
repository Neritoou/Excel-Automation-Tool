"""
Ventana principal de Excel Automator v2.

Mejoras v2:
  - Ejecución en hilo secundario (threading)
  - Bloqueo de controles durante ejecución
  - Barra de estado inferior con timestamp
  - Log acumulativo con historial de ejecuciones
  - Soporte de scroll multiplataforma (Linux/Mac/Windows)
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
from typing import Any

from core.task_registry import TaskRegistry
from core.base_task import BaseTask, TaskResult
from gui.task_frame import TaskFrame


class App(tk.Tk):

    BG_SIDEBAR: str = "#1e293b"
    BG_SIDEBAR_HOVER: str = "#334155"
    BG_SIDEBAR_ACTIVE: str = "#3b82f6"
    FG_SIDEBAR: str = "#cbd5e1"
    FG_SIDEBAR_ACTIVE: str = "#ffffff"
    BG_MAIN: str = "#f8fafc"
    BG_STATUSBAR: str = "#e2e8f0"
    ACCENT: str = "#3b82f6"
    SUCCESS: str = "#16a34a"
    ERROR: str = "#dc2626"

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Automator")
        self.geometry("1000x680")
        self.minsize(820, 520)
        self.configure(bg=self.BG_MAIN)

        self._setup_styles()

        self.registry: TaskRegistry = TaskRegistry()
        self.registry.discover("tasks")

        self._active_task: BaseTask | None = None
        self._task_frame: TaskFrame | None = None
        self._sidebar_buttons: dict[str, tk.Button] = {}
        self._is_running: bool = False
        self._exec_btn: ttk.Button | None = None
        self._prog_var: tk.DoubleVar = tk.DoubleVar(value=0)
        self._prog_lbl: ttk.Label | None = None
        self._log: tk.Text | None = None
        self._status_lbl: ttk.Label | None = None
        self._status_time: ttk.Label | None = None

        self._build_layout()

    def _setup_styles(self) -> None:
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame", background=self.BG_MAIN)
        s.configure("TLabel", background=self.BG_MAIN, font=("Segoe UI", 10))
        s.configure("TLabelframe", background=self.BG_MAIN, font=("Segoe UI", 10, "bold"))
        s.configure("TLabelframe.Label", background=self.BG_MAIN, foreground="#334155")
        s.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground="#0f172a")
        s.configure("Desc.TLabel", font=("Segoe UI", 10), foreground="#64748b")
        s.configure("Status.TLabel", background=self.BG_STATUSBAR, font=("Segoe UI", 9), foreground="#475569")
        s.configure("Exec.TButton", font=("Segoe UI", 11, "bold"), padding=(20, 10))
        s.configure("TCombobox", padding=4)
        s.configure("TEntry", padding=4)
        s.configure("green.Horizontal.TProgressbar", troughcolor="#e2e8f0", background=self.ACCENT, thickness=8)

    # ── Layout ───────────────────────────────────────────────────────
    def _build_layout(self) -> None:
        sb = tk.Frame(self, bg=self.BG_STATUSBAR, height=28)
        sb.pack(side="bottom", fill="x")
        sb.pack_propagate(False)
        self._status_lbl = ttk.Label(sb, text="  Listo", style="Status.TLabel")
        self._status_lbl.pack(side="left", padx=(8, 0), fill="y")
        self._status_time = ttk.Label(sb, text="", style="Status.TLabel")
        self._status_time.pack(side="right", padx=(0, 12), fill="y")

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(main, bg=self.BG_SIDEBAR, width=230)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        self.content = ttk.Frame(main)
        self.content.pack(side="left", fill="both", expand=True)

        tasks = self.registry.get_all()
        if tasks:
            self._select_task(list(tasks.values())[0])

    def _build_sidebar(self) -> None:
        lf = tk.Frame(self.sidebar, bg=self.BG_SIDEBAR)
        lf.pack(fill="x", padx=16, pady=(20, 8))
        tk.Label(lf, text="📊 Excel Automator", font=("Segoe UI", 13, "bold"),
                 bg=self.BG_SIDEBAR, fg="#fff", anchor="w").pack(fill="x")
        tk.Label(lf, text="Herramientas de automatización",
                 font=("Segoe UI", 8), bg=self.BG_SIDEBAR, fg="#94a3b8", anchor="w").pack(fill="x")

        tk.Frame(self.sidebar, bg="#334155", height=1).pack(fill="x", padx=16, pady=(12, 12))
        tk.Label(self.sidebar, text="TAREAS", font=("Segoe UI", 8, "bold"),
                 bg=self.BG_SIDEBAR, fg="#64748b", anchor="w").pack(fill="x", padx=20, pady=(0, 4))

        for tid, task in self.registry.get_all().items():
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

        tk.Label(self.sidebar, text="v2.0", font=("Segoe UI", 8),
                 bg=self.BG_SIDEBAR, fg="#475569").pack(side="bottom", pady=(0, 12))

    # ── Selección de tarea ───────────────────────────────────────────
    def _select_task(self, task: BaseTask) -> None:
        if self._is_running:
            return
        self._active_task = task

        for tid, btn in self._sidebar_buttons.items():
            bg = self.BG_SIDEBAR_ACTIVE if tid == task.task_id else self.BG_SIDEBAR
            fg = self.FG_SIDEBAR_ACTIVE if tid == task.task_id else self.FG_SIDEBAR
            btn.configure(bg=bg, fg=fg)

        for w in self.content.winfo_children():
            w.destroy()

        canvas = tk.Canvas(self.content, bg=self.BG_MAIN, highlightthickness=0)
        vsb = ttk.Scrollbar(self.content, orient="vertical", command=canvas.yview)
        sf = ttk.Frame(canvas)

        sf.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=sf, anchor="nw", tags="inner")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig("inner", width=e.width))

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _wheel(event: tk.Event[Any]) -> None:
            direction: int = -1 if (event.num == 4 or event.delta > 0) else 1
            canvas.yview_scroll(direction * 3, "units")

        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            canvas.bind_all(seq, _wheel)

        hdr = ttk.Frame(sf)
        hdr.pack(fill="x", padx=28, pady=(24, 4))
        ttk.Label(hdr, text=f"{task.task_icon}  {task.task_name}", style="Header.TLabel").pack(anchor="w")
        ttk.Label(hdr, text=task.task_description, style="Desc.TLabel", wraplength=620).pack(anchor="w", pady=(4, 0))

        ttk.Separator(sf).pack(fill="x", padx=28, pady=(12, 4))

        self._task_frame = TaskFrame(sf, task)
        self._task_frame.pack(fill="x", padx=28, pady=(4, 8))

        pf = ttk.Frame(sf)
        pf.pack(fill="x", padx=32, pady=(10, 0))
        self._prog_var = tk.DoubleVar(value=0)
        self._prog_lbl = ttk.Label(pf, text="", style="Desc.TLabel")
        self._prog_lbl.pack(anchor="w")
        ttk.Progressbar(pf, variable=self._prog_var, maximum=100,
                        style="green.Horizontal.TProgressbar").pack(fill="x", pady=(4, 0))

        bf = ttk.Frame(sf)
        bf.pack(fill="x", padx=28, pady=(12, 8))
        self._exec_btn = ttk.Button(bf, text="▶  Ejecutar", style="Exec.TButton", command=self._execute)
        self._exec_btn.pack(side="right")
        ttk.Button(bf, text="Limpiar log", command=self._clear_log).pack(side="right", padx=(0, 8))

        ttk.Label(sf, text="Historial de ejecuciones", style="Desc.TLabel").pack(anchor="w", padx=32, pady=(8, 2))
        self._log = tk.Text(
            sf, height=10, wrap="word", font=("Consolas", 9),
            bg="#f1f5f9", fg="#334155", relief="flat", padx=12, pady=8,
            state="disabled", selectbackground=self.ACCENT, selectforeground="#fff"
        )
        self._log.pack(fill="x", padx=28, pady=(0, 24))
        self._log.tag_configure("ok", foreground=self.SUCCESS)
        self._log.tag_configure("err", foreground=self.ERROR)
        self._log.tag_configure("ts", foreground="#94a3b8")
        self._log.tag_configure("file", foreground=self.ACCENT)

        self._set_status(f"Tarea: {task.task_name}")

    # ── Ejecución con threading ──────────────────────────────────────
    def _execute(self) -> None:
        if self._is_running or self._active_task is None or self._task_frame is None:
            return

        params: dict[str, Any] = self._task_frame.collect_params()
        ok, msg = self._active_task.validate(params)
        if not ok:
            self._log_msg(f"Validación fallida: {msg}", err=True)
            return

        self._set_running(True)
        self._prog_var.set(0)
        if self._prog_lbl is not None:
            self._prog_lbl.config(text="Procesando...")

        # Captura local para evitar que el hilo acceda a atributos mutables
        current_task: BaseTask = self._active_task

        def worker() -> None:
            def cb(pct: float, progress_msg: str) -> None:
                self.after(0, lambda: self._update_progress(pct, progress_msg))

            result: TaskResult = current_task.execute(params, progress_cb=cb)
            self.after(0, self._on_done, result)

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, pct: float, msg: str) -> None:
        self._prog_var.set(pct)
        if self._prog_lbl is not None:
            self._prog_lbl.config(text=msg)

    def _on_done(self, r: TaskResult) -> None:
        self._set_running(False)
        if r.success:
            self._prog_var.set(100)
            self._update_progress(100, "Completado")
            self._log_msg(r.message, files=r.output_files)
            self._set_status("Tarea completada")
        else:
            self._prog_var.set(0)
            self._update_progress(0, "Error")
            self._log_msg(r.message, err=True)
            self._set_status("Finalizó con errores")

    # ── Helpers ──────────────────────────────────────────────────────
    def _set_running(self, on: bool) -> None:
        self._is_running = on
        st: str = "disabled" if on else "normal"
        if self._exec_btn is not None:
            self._exec_btn.config(state=st, text="⏳ Ejecutando..." if on else "▶  Ejecutar")
        for b in self._sidebar_buttons.values():
            b.config(state=st)

    def _log_msg(self, msg: str, err: bool = False, files: list[str] | None = None) -> None:
        if self._log is None:
            return
        self._log.config(state="normal")
        ts: str = datetime.now().strftime("%H:%M:%S")
        icon: str
        tag: str
        icon, tag = ("❌", "err") if err else ("✅", "ok")
        self._log.insert("end", f"[{ts}] ", "ts")
        self._log.insert("end", f"{icon} {msg}\n", tag)
        if files is not None:
            for f in files:
                self._log.insert("end", f"   📁 {f}\n", "file")
        self._log.insert("end", "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self) -> None:
        if self._log is None:
            return
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._prog_var.set(0)
        if self._prog_lbl is not None:
            self._prog_lbl.config(text="")

    def _set_status(self, text: str) -> None:
        if self._status_lbl is not None:
            self._status_lbl.config(text=f"  {text}")
        if self._status_time is not None:
            self._status_time.config(text=datetime.now().strftime("%H:%M:%S"))

    def _on_hover(self, btn: tk.Button, enter: bool) -> None:
        if btn.cget("bg") == self.BG_SIDEBAR_ACTIVE or str(btn.cget("state")) == "disabled":
            return
        btn.configure(bg=self.BG_SIDEBAR_HOVER if enter else self.BG_SIDEBAR)
