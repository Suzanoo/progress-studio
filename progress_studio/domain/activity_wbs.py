from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ActivityWbsSequencer:
    """Generate stable child WBS codes under each parent WBS."""

    _counts: dict[str, int] = field(default_factory=dict)

    def next_code(self, parent_wbs: object, *, fallback: object = "") -> str:
        parent = str(parent_wbs or "").strip()
        if not parent:
            return str(fallback or "").strip()

        next_index = self._counts.get(parent, 0) + 1
        self._counts[parent] = next_index
        return f"{parent}.{next_index}"
