from dataclasses import dataclass


@dataclass(frozen=True)
class WorkbookSchema:
    main_sheet: str = "main"
    mapping_sheet: str = "Amount Mapping"
    info_sheet: str = "Info"
    progress_sheet: str = "progress"
    progress_table_sheet: str = "progress_table"


WORKBOOK_SCHEMA = WorkbookSchema()
