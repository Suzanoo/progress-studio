from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord, SupplementalWBS
from progress_studio.domain.working_tree import WorkingTreeNode


SESSION_FORMAT = "progress-studio-mapping-session"
SESSION_VERSION = 6


@dataclass(frozen=True, slots=True)
class WorkbookFingerprint:
    path: str
    filename: str
    size: int
    modified_ns: int
    sha256: str

    @property
    def saved_path(self) -> Path:
        return Path(self.path)


@dataclass(frozen=True, slots=True)
class MappingSessionData:
    progress: WorkbookFingerprint
    boq: WorkbookFingerprint
    boq_sheet: str
    allocations: tuple[AllocationRecord, ...]
    saved_at: str
    supplemental_activities: tuple[ActivityRow, ...] = ()
    supplemental_wbs: tuple[SupplementalWBS, ...] = ()
    working_tree_nodes: tuple[WorkingTreeNode, ...] = ()
    format: str = SESSION_FORMAT
    version: int = SESSION_VERSION
