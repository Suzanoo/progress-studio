from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from progress_studio.config import Settings


@dataclass(frozen=True)
class CliOptions:
    input_file: Path | None
    cutoff_day: str | None
    amount: float


class CommandLineInterface:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(
            prog="progress-studio",
            description="Build a progress workbook from a Primavera P6 XML file.",
        )
        parser.add_argument(
            "--input",
            type=Path,
            help="Primavera P6 XML file. A file picker opens when omitted.",
        )
        parser.add_argument(
            "--cutoff-day",
            choices=["1", "2", "3", "4", "5", "6", "7"],
            help="Weekly cutoff day: 1=Monday through 7=Sunday. Default: Friday.",
        )
        parser.add_argument(
            "--amount",
            type=float,
            default=self._settings.default_activity_amount,
            help="Placeholder amount assigned to each activity.",
        )
        return parser

    def parse(self, argv: Sequence[str] | None = None) -> CliOptions:
        args = self.build_parser().parse_args(argv)
        return CliOptions(args.input, args.cutoff_day, args.amount)

    @staticmethod
    def select_xml_file() -> Path | None:
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilename(
                title="Select Primavera P6 XML",
                filetypes=[("Primavera XML", "*.xml"), ("All files", "*.*")],
            )
            root.destroy()
            return Path(selected).resolve() if selected else None
        except Exception:
            typed = input("Enter the Primavera XML path: ").strip().strip('"')
            return Path(typed).expanduser().resolve() if typed else None

    def select_cutoff_day(self) -> str:
        labels = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        print("\n" + "=" * 72)
        print("Select Weekly Cutoff Day")
        print("=" * 72)
        for number, label in enumerate(labels, start=1):
            default = "  [default]" if str(number) == self._settings.default_cutoff_day else ""
            print(f"{number}. {label}{default}")
        raw = input("Select [1-7] (Enter = 5 Friday): ").strip()
        return raw or self._settings.default_cutoff_day
