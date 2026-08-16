from __future__ import annotations

from dataclasses import dataclass

from .normalized_schedule import NormalizedSchedule


@dataclass(frozen=True)
class NormalizedScheduleIssue:
    """One source-neutral schedule contract violation."""

    field: str
    message: str
    source_order: int | None = None
    item_label: str = ""


class NormalizedScheduleValidationError(ValueError):
    """Raised when a normalized schedule is unsafe to pass to Progress Engine."""

    def __init__(self, issues: list[NormalizedScheduleIssue]) -> None:
        self.issues = tuple(issues)
        super().__init__(self._build_message())

    def _build_message(self) -> str:
        lines = [
            "Normalized Schedule Validation Failed",
            "",
            "The normalized schedule is not safe to pass to Progress Engine.",
            "",
        ]
        for issue in self.issues[:12]:
            where = ""
            if issue.source_order is not None:
                where = f" at source order {issue.source_order}"
            label = f" ({issue.item_label})" if issue.item_label else ""
            lines.append(f"- {issue.field}{where}{label}: {issue.message}")
        if len(self.issues) > 12:
            lines.append(f"- ...and {len(self.issues) - 12} more issue(s)")
        return "\n".join(lines)


class NormalizedScheduleValidator:
    """Validate the source-neutral contract before workbook generation.

    N-6 intentionally knows nothing about XML format, Excel, timescale display
    margins, fake Amount, or workbook rendering.  Those remain later concerns.
    """

    def validate(self, schedule: NormalizedSchedule) -> NormalizedSchedule:
        issues: list[NormalizedScheduleIssue] = []

        self._validate_wbs(schedule, issues)
        self._validate_activities(schedule, issues)
        self._validate_project_window(schedule, issues)

        if issues:
            raise NormalizedScheduleValidationError(issues)
        return schedule

    @staticmethod
    def _validate_wbs(
        schedule: NormalizedSchedule, issues: list[NormalizedScheduleIssue]
    ) -> None:
        wbs_by_code: dict[str, object] = {}

        for row in schedule.wbs:
            code = row.wbs_code.strip()
            if not code:
                issues.append(
                    NormalizedScheduleIssue(
                        field="WBS Code",
                        message="WBS code is blank.",
                        source_order=row.source_order,
                        item_label=row.wbs_name,
                    )
                )
                continue
            if code in wbs_by_code:
                issues.append(
                    NormalizedScheduleIssue(
                        field="WBS Code",
                        message=f"Duplicate WBS code {code!r}.",
                        source_order=row.source_order,
                        item_label=row.wbs_name,
                    )
                )
            else:
                wbs_by_code[code] = row

        known_codes = set(wbs_by_code)
        parents: dict[str, str | None] = {}
        for row in schedule.wbs:
            code = row.wbs_code.strip()
            if not code:
                continue
            parent = row.parent_wbs_code.strip() if row.parent_wbs_code else None
            parents[code] = parent
            if parent and parent not in known_codes:
                issues.append(
                    NormalizedScheduleIssue(
                        field="WBS Parent",
                        message=f"Parent WBS {parent!r} does not exist.",
                        source_order=row.source_order,
                        item_label=code,
                    )
                )
            if parent == code:
                issues.append(
                    NormalizedScheduleIssue(
                        field="WBS Parent",
                        message="WBS cannot be its own parent.",
                        source_order=row.source_order,
                        item_label=code,
                    )
                )

        # Detect cycles that can still exist even when every parent reference resolves.
        for start in parents:
            seen: set[str] = set()
            current: str | None = start
            while current is not None and current in parents:
                if current in seen:
                    issues.append(
                        NormalizedScheduleIssue(
                            field="WBS Parent",
                            message=f"WBS hierarchy contains a cycle involving {current!r}.",
                            item_label=start,
                        )
                    )
                    break
                seen.add(current)
                current = parents[current]

    @staticmethod
    def _validate_activities(
        schedule: NormalizedSchedule, issues: list[NormalizedScheduleIssue]
    ) -> None:
        if not schedule.activities:
            issues.append(
                NormalizedScheduleIssue(
                    field="Activities",
                    message="Schedule contains no activities.",
                )
            )
            return

        known_wbs = {row.wbs_code.strip() for row in schedule.wbs if row.wbs_code.strip()}
        seen_ids: dict[str, int] = {}

        for row in schedule.activities:
            activity_id = row.activity_id.strip()
            label = activity_id or row.activity_name.strip() or "unnamed activity"

            if not activity_id:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Activity ID",
                        message="Activity ID is blank; Progress Studio will not fabricate one.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )
            elif activity_id in seen_ids:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Activity ID",
                        message=(
                            f"Duplicate Activity ID {activity_id!r}; first seen at source order "
                            f"{seen_ids[activity_id]}."
                        ),
                        source_order=row.source_order,
                        item_label=label,
                    )
                )
            else:
                seen_ids[activity_id] = row.source_order

            if not row.activity_name.strip():
                issues.append(
                    NormalizedScheduleIssue(
                        field="Activity Name",
                        message="Activity name is blank.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )

            wbs_code = row.wbs_code.strip()
            if not wbs_code:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Activity WBS",
                        message="Activity WBS code is blank.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )
            elif wbs_code not in known_wbs:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Activity WBS",
                        message=f"Activity references unknown WBS code {wbs_code!r}.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )

            if row.plan_start is None:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Plan Start",
                        message="Plan Start is missing or invalid.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )
            if row.plan_finish is None:
                issues.append(
                    NormalizedScheduleIssue(
                        field="Plan Finish",
                        message="Plan Finish is missing or invalid.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )
            if (
                row.plan_start is not None
                and row.plan_finish is not None
                and row.plan_start > row.plan_finish
            ):
                issues.append(
                    NormalizedScheduleIssue(
                        field="Plan Dates",
                        message="Plan Start is after Plan Finish.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )

            if (
                row.actual_start is not None
                and row.actual_finish is not None
                and row.actual_start > row.actual_finish
            ):
                issues.append(
                    NormalizedScheduleIssue(
                        field="Actual Dates",
                        message="Actual Start is after Actual Finish.",
                        source_order=row.source_order,
                        item_label=label,
                    )
                )

    @staticmethod
    def _validate_project_window(
        schedule: NormalizedSchedule, issues: list[NormalizedScheduleIssue]
    ) -> None:
        start = schedule.project.plan_start
        finish = schedule.project.plan_finish
        if start is not None and finish is not None and start > finish:
            issues.append(
                NormalizedScheduleIssue(
                    field="Project Window",
                    message="Project Plan Start is after Project Plan Finish.",
                    item_label=schedule.project.project_name,
                )
            )
