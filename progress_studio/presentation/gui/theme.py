from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk


@dataclass(frozen=True)
class Palette:
    canvas: str
    surface: str
    surface_alt: str
    border: str
    text: str
    muted: str
    primary: str
    primary_dark: str
    primary_soft: str
    sidebar: str
    sidebar_hover: str
    sidebar_text: str
    sidebar_muted: str
    success: str
    warning: str
    danger: str
    info: str
    selection: str
    console_bg: str
    console_fg: str


def _theme_path() -> Path:
    return Path(__file__).resolve().parents[2] / "config" / "theme.json"


def _load_theme() -> tuple[Palette, str, str]:
    data = json.loads(_theme_path().read_text(encoding="utf-8"))
    colors = data["colors"]
    palette = Palette(
        canvas=colors["canvas"], surface=colors["surface"],
        surface_alt=colors["surface_alt"], border=colors["border"],
        text=colors["text"], muted=colors["muted"],
        primary=colors["primary"], primary_dark=colors["primary_hover"],
        primary_soft=colors["primary_soft"], sidebar=colors["sidebar"],
        sidebar_hover=colors["sidebar_hover"], sidebar_text=colors["sidebar_text"],
        sidebar_muted=colors["sidebar_muted"], success=colors["success"],
        warning=colors["warning"], danger=colors["danger"], info=colors["info"],
        selection=colors["selection"], console_bg=colors["console_bg"],
        console_fg=colors["console_fg"],
    )
    fonts = data.get("fonts", {})
    return palette, fonts.get("ui", "Segoe UI"), fonts.get("mono", "Consolas")


PALETTE, FONT_UI, FONT_MONO = _load_theme()


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
    style.configure("WorkspaceTitle.TLabel", background=p.surface, foreground=p.text, font=(FONT_UI, 12, "bold"))
    style.configure("Section.TLabel", background=p.surface, foreground=p.text, font=(FONT_UI, 10, "bold"))
    style.configure("Sidebar.TLabel", background=p.sidebar, foreground=p.sidebar_text, font=(FONT_UI, 9))
    style.configure("SidebarMuted.TLabel", background=p.sidebar, foreground=p.sidebar_muted, font=(FONT_UI, 8))
    style.configure("SidebarTitle.TLabel", background=p.sidebar, foreground=p.sidebar_text, font=(FONT_UI, 12, "bold"))
    style.configure("Card.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("Card.TLabel", background=p.surface, foreground=p.text)
    style.configure("Accent.TButton", font=(FONT_UI, 9, "bold"), padding=(14, 7))
    style.map("Accent.TButton", background=[("active", p.primary_dark), ("!disabled", p.primary)], foreground=[("!disabled", "#FFFFFF")])
    style.configure("Sidebar.TButton", anchor="w", padding=(14, 10), font=(FONT_UI, 9), background=p.sidebar, foreground=p.sidebar_text, borderwidth=0)
    style.map("Sidebar.TButton", background=[("active", p.sidebar_hover)])
    style.configure("SidebarActive.TButton", anchor="w", padding=(14, 10), font=(FONT_UI, 9, "bold"), background=p.primary, foreground="#FFFFFF", borderwidth=0)
    style.map("SidebarActive.TButton", background=[("active", p.primary_dark)])
    style.configure("Treeview", rowheight=27, font=(FONT_UI, 9), background=p.surface, fieldbackground=p.surface, foreground=p.text, bordercolor=p.border)
    style.configure("Treeview.Heading", font=(FONT_UI, 9, "bold"), background=p.surface_alt, foreground=p.text, padding=(6, 7), relief="flat")
    style.map("Treeview", background=[("selected", p.selection)], foreground=[("selected", p.text)])
    style.configure("TNotebook", background=p.surface, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(14, 8), font=(FONT_UI, 9))
    style.configure("TProgressbar", thickness=10)
    style.configure("CommandBar.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("StatusBar.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("Status.TLabel", background=p.surface, foreground=p.muted, padding=(8, 4))
    style.configure("StatusReady.TLabel", background=p.surface, foreground=p.success, font=(FONT_UI, 9, "bold"), padding=(8, 4))
    style.configure("Empty.TLabel", background=p.surface, foreground=p.muted, font=(FONT_UI, 9, "italic"), padding=(8, 4))
    style.configure("Loading.TFrame", background=p.surface, relief="solid", borderwidth=1)
    style.configure("LoadingTitle.TLabel", background=p.surface, foreground=p.text, font=(FONT_UI, 11, "bold"))
