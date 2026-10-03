"""Small, scrollable Home and Help views. No workbook/service ownership."""
from tkinter import ttk
import tkinter as tk

from .strings import tr


def page(parent):
    canvas = tk.Canvas(parent, highlightthickness=0, background="#FAF9F6")
    scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    body = ttk.Frame(canvas, style="Surface.TFrame", padding=16)
    item = canvas.create_window((0, 0), window=body, anchor="nw")
    body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(item, width=e.width))
    return body


def label(parent, key, style="Muted.TLabel"):
    widget = ttk.Label(parent, text=tr(key), style=style, wraplength=400)
    widget.pack(anchor="w", fill="x", pady=(0, 8))
    widget.bind("<Configure>", lambda e: widget.configure(wraplength=max(100, e.width)))
    return widget


HOME_CARDS = (
    ("home.create", "home.create_help", "import"),
    ("home.rebuild", "home.rebuild_help", "rebuild"),
    ("home.mapping", "home.mapping_help", "mapping"),
    ("home.payment", "home.payment_help", "payment"),
)


def build_home(parent, navigate):
    body = page(parent)
    label(body, "home.title", "Title.TLabel")
    strip = ttk.Frame(body, style="Card.TFrame", padding=12)
    strip.pack(fill="x", pady=(8, 16))
    label(strip, "home.flow")
    ttk.Button(strip, text=tr("guide.view"), command=lambda: navigate("help")).pack(anchor="e")
    cards = ttk.Frame(body, style="Surface.TFrame")
    cards.pack(fill="both", expand=True)
    for col in (0, 1):
        cards.columnconfigure(col, weight=1, uniform="home")
    for index, (title, description, target) in enumerate(HOME_CARDS):
        card = ttk.Frame(cards, style="Card.TFrame", padding=18)
        card.grid(row=index // 2, column=index % 2, sticky="nsew", padx=5, pady=5)
        label(card, title, "WorkspaceTitle.TLabel")
        if index > 1:
            label(card, "home.optional")
        label(card, description)
        ttk.Button(card, text=tr(title if index < 2 else "Open"),
                   style="Accent.TButton" if index == 0 else "TButton",
                   command=lambda key=target: navigate(key)).pack(fill="x", pady=(8, 0))


def build_help(parent, navigate):
    body = page(parent)
    label(body, "guide.title", "Title.TLabel")
    for flow in ("first", "cycle"):
        section = ttk.Frame(body, style="Card.TFrame", padding=16)
        section.pack(fill="x", pady=(8, 12))
        label(section, f"guide.{flow}", "Section.TLabel")
        steps = ttk.Frame(section, style="Surface.TFrame")
        steps.pack(fill="x")
        for index in range(1, 4):
            steps.columnconfigure(index - 1, weight=1, uniform=flow)
            step = ttk.Frame(steps, style="Surface.TFrame", padding=10)
            step.grid(row=0, column=index - 1, sticky="nsew")
            ttk.Label(step, text=str(index) + ("  →" if index < 3 else ""), style="Section.TLabel").pack(anchor="w")
            label(step, f"guide.{flow}.{index}", "Section.TLabel")
            label(step, f"guide.{flow}.{index}.help")
    label(body, "guide.rebuild_note")
    label(body, "guide.optional", "Section.TLabel")
    label(body, "guide.ev_note")
    ttk.Button(body, text=tr("guide.back"), command=lambda: navigate("home")).pack(anchor="w", pady=8)
