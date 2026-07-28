from __future__ import annotations

from pathlib import Path


class FixedDistributionPrompt:
    """Non-interactive distribution choice used by the desktop application."""

    def __init__(self, method: str = "auto") -> None:
        self.method = method

    def choose_method(self) -> str | None:
        return self.method

    def review(self, output_file: Path, method: str) -> str:
        return "accept"
