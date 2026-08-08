from dataclasses import dataclass

@dataclass(frozen=True, slots=True)
class EditedWorkbookMigrationResult:
    source_activity_count: int
    target_activity_count: int
    matched_activity_count: int
    matched_by_activity_id: int
    matched_by_signature: int
    unmatched_activity_ids: tuple[str, ...]
    ambiguous_activity_ids: tuple[str, ...]
    amount_cells_migrated: int
    plan_cells_migrated: int
    actual_cells_migrated: int

    @property
    def has_review_items(self) -> bool:
        return bool(self.unmatched_activity_ids or self.ambiguous_activity_ids)
