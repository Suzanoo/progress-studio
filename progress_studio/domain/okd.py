from dataclasses import dataclass


@dataclass(frozen=True)
class OkdExportResult:
    activities: int
    weeks: int
    table_rows: int
    checked_links: int
