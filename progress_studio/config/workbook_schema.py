from dataclasses import dataclass


@dataclass(frozen=True)
class WorkbookSchema:
    main_sheet: str = "main"
    main_monthly_sheet: str = "main_monthly"
    mapping_sheet: str = "Amount Mapping"
    info_sheet: str = "Info"
    progress_sheet: str = "progress"
    progress_table_sheet: str = "progress_table"


WORKBOOK_SCHEMA = WorkbookSchema()
