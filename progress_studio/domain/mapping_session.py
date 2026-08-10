from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord, SupplementalWBS
from progress_studio.domain.working_tree import WorkingTreeNode


SESSION_FORMAT = "progress-studio-mapping-session"
SESSION_VERSION = 8


@dataclass(frozen=True, slots=True)
class WorkbookFingerprint:
    path: str
    filename: str
    size: int
    modified_ns: int
    sha256: str
    # Version 7 adds a stable, workbook-aware identity. It hashes worksheet
    # names, cell coordinates, formulas and values, while deliberately ignoring
    # ZIP metadata, formatting and other binary details changed by Excel saves.
    semantic_sha256: str = ""
    identity_kind: str = "binary-sha256"

    @property
    def saved_path(self) -> Path:
        return Path(self.path)

    @property
    def has_semantic_identity(self) -> bool:
        return bool(self.semantic_sha256)


@dataclass(frozen=True, slots=True)
class WorkbookSnapshot:
    filename: str
    sha256: str
    content_b64: str

    @property
    def is_available(self) -> bool:
        return bool(self.content_b64)


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
    progress_snapshot: WorkbookSnapshot | None = None
    boq_snapshot: WorkbookSnapshot | None = None
    format: str = SESSION_FORMAT
    version: int = SESSION_VERSION
