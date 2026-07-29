from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ScheduleXmlIssue:
    source_order: int
    activity_label: str
    field: str
    message: str


class ScheduleXmlValidationError(ValueError):
    """Raised when an XML schedule cannot satisfy the import contract."""

    def __init__(self, issues: list[ScheduleXmlIssue]) -> None:
        self.issues = tuple(issues)
        super().__init__(self._build_message())

    def _build_message(self) -> str:
        counts: dict[str, int] = {}
        for issue in self.issues:
            counts[issue.field] = counts.get(issue.field, 0) + 1

        lines = [
            "Import Failed",
            "",
            "The XML schedule is missing or contains invalid required activity data.",
            "Required fields: Activity Name, Plan Start, Plan Finish.",
            "",
        ]
        for field in ("Activity Name", "Plan Start", "Plan Finish", "Schedule Window", "Activities"):
            if field in counts:
                lines.append(f"{field}: {counts[field]} issue(s)")

        lines.extend(["", "Examples:"])
        for issue in self.issues[:8]:
            lines.append(
                f'- Item {issue.source_order + 1} ({issue.activity_label or "unnamed"}): {issue.message}'
            )
        if len(self.issues) > 8:
            lines.append(f"- ...and {len(self.issues) - 8} more issue(s)")
        lines.extend(["", "No workbook was created."])
        return "\n".join(lines)
