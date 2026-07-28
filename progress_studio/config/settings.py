from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    title: str = "Progress Studio"
    version: str = "2.0.4"
    output_root_name: str = "Progress_Studio_Output"
    default_activity_amount: float = 100_000.0
    default_cutoff_day: str = "5"
    project_root: Path = field(
        default_factory=lambda: Path(__file__).resolve().parents[2]
    )


SETTINGS = Settings()
