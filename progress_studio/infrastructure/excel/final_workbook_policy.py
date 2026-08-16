from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from openpyxl.workbook.workbook import Workbook

from progress_studio.infrastructure.excel.calculation_policy import (
    configure_incremental_excel_recalculation,
    configure_live_save_recalculation,
)
from progress_studio.infrastructure.excel.workbook_guide import build_workbook_guide
from progress_studio.infrastructure.excel.workbook_protection import apply_final_sheet_protection
from progress_studio.infrastructure.excel.workbook_visibility import apply_final_sheet_visibility


class FinalWorkbookMode(str, Enum):
    """Final workbook policy modes.

    SNAPSHOT keeps the existing dependency-based automatic recalculation policy
    used by Create Progress, Mapping Export, Snapshot Rebuild and Payment.

    LIVE keeps the lightweight/manual-until-save policy used by Live Rebuild.
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
      4. Recalculation is selected by workbook mode, never XML source type.
    """
    resolved_mode = FinalWorkbookMode(mode)

    guide_sheet: str | None = None
    if include_guide:
        build_workbook_guide(workbook)
        guide_sheet = "README"

    visible, hidden, very_hidden = apply_final_sheet_visibility(workbook)
    protected = apply_final_sheet_protection(workbook)

    if resolved_mode is FinalWorkbookMode.LIVE:
        configure_live_save_recalculation(workbook)
    else:
        configure_incremental_excel_recalculation(workbook)

    return FinalWorkbookPolicyResult(
        mode=resolved_mode,
        guide_sheet=guide_sheet,
        visible_sheets=visible,
        hidden_sheets=hidden,
        very_hidden_sheets=very_hidden,
        protected_sheets=protected,
    )
