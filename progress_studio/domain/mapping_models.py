from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum


@dataclass(frozen=True, slots=True)
class ActivityRow:
    activity_id: str
    parent_wbs: str
    child_wbs: str
    description: str
    wbs_path: tuple[tuple[str, str], ...] = ()

    @property
    def search_text(self) -> str:
        hierarchy = " ".join(f"{code} {name}" for code, name in self.wbs_path)
        return (
            f"{self.activity_id} {self.parent_wbs} {self.child_wbs} "
            f"{hierarchy} {self.description}"
        ).lower()


@dataclass(frozen=True, slots=True)
class BOQRow:
    key: str
    source_sheet: str
    source_row: int
    wbs2: str
    wbs3: str
    wbs4: str
    description: str
    amount: float

    @property
    def search_text(self) -> str:
        return f"{self.wbs2} {self.wbs3} {self.wbs4} {self.description}".lower()


class MappingStatus(StrEnum):
    UNMAPPED = "Unmapped"
    PARTIAL = "Partial"
    FULL = "Full"


@dataclass(frozen=True, slots=True)
class AllocationRecord:
    """One BOQ-to-Activity allocation expressed as a percentage of BOQ amount."""

    boq_key: str
    activity_id: str
    share_percent: float


@dataclass(frozen=True, slots=True)
class MappingChange:
    """Identifiers affected by one mapping command."""

    boq_keys: tuple[str, ...]
    activity_ids: tuple[str, ...]
