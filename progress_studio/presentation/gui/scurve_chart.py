from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from progress_studio.services.scurve_service import SCurveData


class SCurveChart(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.figure = Figure(figsize=(6.4, 4.2), dpi=100, tight_layout=True)
        self.axes = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.show_empty("Create a progress workbook to preview the S-curve.")

    def show_empty(self, message: str) -> None:
        self.axes.clear()
        self.axes.text(0.5, 0.5, message, ha="center", va="center", transform=self.axes.transAxes)
        self.axes.set_axis_off()
        self.canvas.draw_idle()

    def render(self, data: SCurveData) -> None:
        self.axes.clear()
        self.axes.set_axis_on()
        self.axes.plot(data.dates, data.plan, label="Plan", linewidth=2.0)

        actual_dates = [date for date, value in zip(data.dates, data.actual) if value is not None]
        actual_values = [value for value in data.actual if value is not None]
        if actual_values:
            self.axes.plot(actual_dates, actual_values, label="Actual", linewidth=2.0)

        self.axes.set_title("Project S-Curve")
        self.axes.set_ylabel("Cumulative progress (%)")
        self.axes.set_ylim(bottom=0)
        self.axes.grid(True, alpha=0.25)
        self.axes.legend(loc="best")
        self.figure.autofmt_xdate()
        self.canvas.draw_idle()
