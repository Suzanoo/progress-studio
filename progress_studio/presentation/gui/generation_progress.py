from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class GenerationProgressDialog(tk.Toplevel):
    """Small modal progress surface for workbook generation.

    The generator remains a headless service. This dialog is only a visual
    observer used by the mapping/export workspace.
    """

    STEPS = (
        ("read", "Read working schedule"),
        ("main", "Build main schedule"),
        ("timescale", "Build timescale"),
        ("mapping", "Apply amount mapping"),
        ("progress", "Build progress sheets + Dashboard"),
        ("distribution", "Generate plan distribution"),
        ("okd", "Build OKD sheets"),
        ("finalize", "Finalize workbook"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.title("Generating Workbook")
        self.transient(parent.winfo_toplevel())
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Generating Workbook", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="Progress Studio is rebuilding the workbook from the working tree.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        self.labels: dict[str, ttk.Label] = {}
        for key, text in self.STEPS:
            label = ttk.Label(body, text=f"○  {text}")
            label.pack(anchor="w", pady=2)
            self.labels[key] = label

        self.progress = ttk.Progressbar(
            body, mode="determinate", maximum=len(self.STEPS), length=410
        )
        self.progress.pack(fill="x", pady=(16, 6))
        self.status_var = tk.StringVar(value="Starting...")
        ttk.Label(body, textvariable=self.status_var, style="Muted.TLabel").pack(anchor="w")

        self.completed: set[str] = set()
        self.update_idletasks()
        top = parent.winfo_toplevel()
        x = top.winfo_rootx() + max(0, (top.winfo_width() - self.winfo_reqwidth()) // 2)
        y = top.winfo_rooty() + max(0, (top.winfo_height() - self.winfo_reqheight()) // 2)
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.lift()
        self.update_idletasks()

    def update_step(self, step: str, message: str) -> None:
        if step not in self.labels:
            return
        for key, label in self.labels.items():
            prefix = "✓" if key in self.completed else ("●" if key == step else "○")
            text = dict(self.STEPS)[key]
            label.configure(text=f"{prefix}  {text}")
        self.status_var.set(message)
        self.progress.configure(value=len(self.completed))
        self.update_idletasks()

    def complete_step(self, step: str, message: str = "") -> None:
        if step in self.labels:
            self.completed.add(step)
            self.labels[step].configure(text=f"✓  {dict(self.STEPS)[step]}")
            self.progress.configure(value=len(self.completed))
        if message:
            self.status_var.set(message)
        self.update_idletasks()

    def fail(self, message: str) -> None:
        self.status_var.set(message)
        self.update_idletasks()

    def close(self) -> None:
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
