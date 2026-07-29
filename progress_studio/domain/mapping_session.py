from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from progress_studio.domain.mapping_models import AllocationRecord


SESSION_FORMAT = "progress-studio-mapping-session"
SESSION_VERSION = 1


@dataclass(frozen=True, slots=True)
class WorkbookFingerprint:
    path: str
    size: int
    modified_ns: int
    sha256: str

    @property
    def filename(self) -> str:
        return Path(self.path).name


@dataclass(frozen=True, slots=True)
class MappingSessionData:
    progress: WorkbookFingerprint
    boq: WorkbookFingerprint
    boq_sheet: str
    allocations: tuple[AllocationRecord, ...]
    saved_at: str
    format: str = SESSION_FORMAT
    version: int = SESSION_VERSION
