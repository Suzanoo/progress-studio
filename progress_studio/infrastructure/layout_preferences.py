from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass
class LayoutPreferences:
    mapping_inputs_collapsed: bool = False
    generator_collapsed: bool = False
    mapping_sash: int | None = None


class LayoutPreferencesRepository:
    """Small JSON-backed store for non-business UI preferences only."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (Path.home() / ".progress_studio" / "layout.json")

    def load(self) -> LayoutPreferences:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
            return LayoutPreferences(
                mapping_inputs_collapsed=bool(payload.get("mapping_inputs_collapsed", False)),
                generator_collapsed=bool(payload.get("generator_collapsed", False)),
                mapping_sash=self._optional_int(payload.get("mapping_sash")),
            )
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            return LayoutPreferences()

    def save(self, preferences: LayoutPreferences) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(json.dumps(asdict(preferences), indent=2), encoding="utf-8")
        temporary.replace(self.path)

    @staticmethod
    def _optional_int(value: object) -> int | None:
        if value is None:
            return None
        try:
            return int(value)
        except (TypeError, ValueError):
            return None
