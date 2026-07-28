from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class AmountSourceDecision:
    use_xml_amounts: bool
    source_label: str


@dataclass(frozen=True)
class AmountMappingResult:
    activity_count: int
    total_amount: float
    amount_source: str


@dataclass(frozen=True)
class AmountApplicationResult:
    mapped_activities: int
    unmapped_activity_ids: tuple[str, ...]
    wbs_rollups: int
    project_rollups: int
    source_rows_used: int
    source_rows_skipped: int
    duplicate_activity_ids: tuple[str, ...]
    source_label: str
