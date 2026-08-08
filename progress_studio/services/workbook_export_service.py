from __future__ import annotations

from pathlib import Path
from collections.abc import Callable

from progress_studio.domain.export_models import ExportResult, ExportValidation
from progress_studio.domain.mapping_models import AllocationRecord, BOQRow, MappingStatus
from progress_studio.domain.working_tree import WorkingTreeNode
from progress_studio.infrastructure.excel.mapped_workbook_exporter import MappedWorkbookExporter
from progress_studio.services.mapping_store import MappingStore


class WorkbookExportService:
    def __init__(self, exporter: MappedWorkbookExporter | None = None) -> None:
        self.exporter = exporter or MappedWorkbookExporter()

    @staticmethod
    def validate(store: MappingStore) -> ExportValidation:
        if not store.activities_by_id:
            raise ValueError('No Progress Activities are loaded.')
        if not store.boq_by_id:
            raise ValueError('No BOQ items are loaded.')

        statuses = {key: store.boq_status(key) for key in store.boq_order}
        mapped_activities = {
            activity_id for _, activity_id in store.allocations
            if store.activity_amount(activity_id) > store.EPSILON
        }
        return ExportValidation(
            activity_count=len(store.activities_by_id),
            boq_count=len(store.boq_by_id),
            allocation_count=len(store.allocations),
            mapped_activity_count=len(mapped_activities),
            mapped_boq_count=store.mapped_item_count,
            full_boq_count=sum(status is MappingStatus.FULL for status in statuses.values()),
            partial_boq_count=sum(status is MappingStatus.PARTIAL for status in statuses.values()),
            unmapped_boq_count=sum(status is MappingStatus.UNMAPPED for status in statuses.values()),
            total_boq_amount=store.total_amount,
            allocated_amount=store.mapped_amount,
            remaining_amount=store.remaining_amount,
        )

    def export(
        self,
        progress_file: Path,
        output_file: Path,
        store: MappingStore,
        *,
        overwrite: bool = False,
        progress_callback: Callable[[str, str, bool], None] | None = None,
        edited_workbook: Path | None = None,
    ) -> ExportResult:
        validation = self.validate(store)
        return self.exporter.export(
            progress_file,
            output_file,
            list(store.boq_by_id.values()),
            store.allocation_records(),
            validation,
            activities=list(store.activities_by_id.values()),
            supplemental_wbs=list(store.supplemental_wbs_nodes),
            working_tree_nodes=list(store.working_tree_nodes()),
            overwrite=overwrite,
            progress_callback=progress_callback,
            edited_workbook=edited_workbook,
        )
