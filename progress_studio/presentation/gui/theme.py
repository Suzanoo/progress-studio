from __future__ import annotations

from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk


@dataclass(frozen=True)
class Palette:
    canvas: str = "#f3f6fa"
    surface: str = "#ffffff"
    surface_alt: str = "#f8fafc"
    border: str = "#dce3ec"
    text: str = "#172033"
    muted: str = "#667085"
    primary: str = "#1677e8"
    primary_dark: str = "#0f5fbf"
    sidebar: str = "#17324d"
    sidebar_hover: str = "#214564"
    sidebar_text: str = "#eaf2f8"
    success: str = "#18864b"
    warning: str = "#b76e00"
    danger: str = "#b42318"


PALETTE = Palette()
FONT_UI = "Segoe UI"
FONT_MONO = "Consolas"


def configure_styles(root: tk.Misc) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    p = PALETTE
    style.configure("TFrame", background=p.canvas)
    style.configure("App.TFrame", background=p.canvas)
    style.configure("Surface.TFrame", background=p.surface)
    style.configure("Sidebar.TFrame", background=p.sidebar)
    style.configure("TLabel", background=p.canvas, foreground=p.text, font=(FONT_UI, 9))
    style.configure("Surface.TLabel", background=p.surface, foreground=p.text)
    style.configure("Muted.TLabel", background=p.surface, foreground=p.muted)
    style.configure("Title.TLabel", background=p.canvas, foreground=p.text, font=(FONT_UI, 17, "bold"))
    style.configure("Section.TLabel", background=p.surface, foreground=p.text, font=(FONT_UI, 10, "bold"))
    style.configure("Sidebar.TLabel", background=p.sidebar, foreground=p.sidebar_text, font=(FONT_UI, 9))
    style.configure("SidebarTitle.TLabel", background=p.sidebar, foreground="#ffffff", font=(FONT_UI, 12, "bold"))
    style.configure("Stage.TLabel", background=p.surface_alt, foreground=p.muted, padding=(12, 8))
    style.configure("StageActive.TLabel", background=p.primary, foreground="#ffffff", font=(FONT_UI, 9, "bold"), padding=(14, 8))
    style.configure("Card.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("Card.TLabel", background=p.surface, foreground=p.text)
    style.configure("Accent.TButton", font=(FONT_UI, 9, "bold"), padding=(15, 8))
    style.map("Accent.TButton", background=[("active", p.primary_dark), ("!disabled", p.primary)], foreground=[("!disabled", "#ffffff")])
    style.configure("Sidebar.TButton", anchor="w", padding=(12, 9), font=(FONT_UI, 9), background=p.sidebar, foreground=p.sidebar_text, borderwidth=0)
    style.map("Sidebar.TButton", background=[("active", p.sidebar_hover), ("pressed", p.primary)])
    style.configure("SidebarActive.TButton", anchor="w", padding=(12, 9), font=(FONT_UI, 9, "bold"), background=p.primary, foreground="#ffffff", borderwidth=0)
    style.map("SidebarActive.TButton", background=[("active", p.primary_dark)])
    style.configure("Treeview", rowheight=28, font=(FONT_UI, 9), background=p.surface, fieldbackground=p.surface, foreground=p.text, bordercolor=p.border)
    style.configure("Treeview.Heading", font=(FONT_UI, 9, "bold"), background=p.surface_alt, foreground=p.text, padding=(6, 7), relief="flat")
    style.map("Treeview", background=[("selected", "#dcecff")], foreground=[("selected", p.text)])
    style.configure("TNotebook", background=p.surface, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(14, 8), font=(FONT_UI, 9))
    style.configure("TProgressbar", thickness=12)
    style.configure("Status.TLabel", background=p.surface, foreground=p.muted, padding=(8, 4))
    style.configure("CommandBar.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("StatusBar.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("StatusReady.TLabel", background=p.surface, foreground=p.success, font=(FONT_UI, 9, "bold"), padding=(8, 4))
    style.configure("Empty.TLabel", background=p.canvas, foreground=p.muted, font=(FONT_UI, 9, "italic"), padding=(8, 4))
    style.configure("Loading.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("LoadingTitle.TLabel", background=p.surface, foreground=p.text, font=(FONT_UI, 11, "bold"))
