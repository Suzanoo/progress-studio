from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from openpyxl.workbook.workbook import Workbook

from progress_studio.infrastructure.excel.calculation_policy import (
    configure_user_driven_save_recalculation,
)
from progress_studio.infrastructure.excel.workbook_guide import build_workbook_guide
from progress_studio.infrastructure.excel.workbook_protection import apply_final_sheet_protection
from progress_studio.infrastructure.excel.workbook_visibility import apply_final_sheet_visibility
from progress_studio.infrastructure.excel.traditional_overlay_workbook import reassert_traditional_overlay_transparency


class FinalWorkbookMode(str, Enum):
    """Final workbook policy modes.

    Both modes share the same user-driven F9 / Save formula-calculation policy.
    The mode remains explicit because Snapshot and Live still differ in who owns
    generated caches/snapshots outside Excel formula calculation.
    """

    SNAPSHOT = "snapshot"
    LIVE = "live"


@dataclass(frozen=True, slots=True)
class FinalWorkbookPolicyResult:
    mode: FinalWorkbookMode
    guide_sheet: str | None
    visible_sheets: tuple[str, ...]
    hidden_sheets: tuple[str, ...]
    very_hidden_sheets: tuple[str, ...]
    protected_sheets: tuple[str, ...]


def finalize_workbook(
    workbook: Workbook,
    *,
    mode: FinalWorkbookMode | str = FinalWorkbookMode.SNAPSHOT,
    include_guide: bool = True,
) -> FinalWorkbookPolicyResult:
    """Apply the shared final portable-workbook contract.

    This function intentionally orchestrates existing proven helpers. It does
    not build/rebuild progress, monthly, dashboard, payment or overlay data.
    Those calculation/rendering engines remain owned by their current pipelines.

    Order matters:
      1. Guide is created first so visibility/protection include it.
      2. Visibility establishes the final sheet-state contract.
      3. Protection locks formulas/support data and re-opens intended inputs.
      4. Final files use one explicit F9 / Save formula-calculation contract.
    """
    resolved_mode = FinalWorkbookMode(mode)

    guide_sheet: str | None = None
    if include_guide:
        build_workbook_guide(workbook)
        guide_sheet = "README"

    # Preserve renderer-owned transparent overlay appearance across openpyxl
    # load/save round-trips. This changes presentation only; it does not rebuild
    # series, cutoff logic, or any business data.
    reassert_traditional_overlay_transparency(workbook)

    visible, hidden, very_hidden = apply_final_sheet_visibility(workbook)
    protected = apply_final_sheet_protection(workbook)

    # Final user-facing workbooks are intentionally Manual. This prevents
    # Excel from continuously recalculating the large progress workbook while
    # the user edits it. F9 recalculates formulas on demand and calcOnSave asks
    # Excel to calculate formulas before Save. Python-owned generated snapshots
    # remain the responsibility of Progress Studio Rebuild.
    configure_user_driven_save_recalculation(workbook)

    return FinalWorkbookPolicyResult(
        mode=resolved_mode,
        guide_sheet=guide_sheet,
        visible_sheets=visible,
        hidden_sheets=hidden,
        very_hidden_sheets=very_hidden,
        protected_sheets=protected,
    )
