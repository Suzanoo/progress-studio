from __future__ import annotations

from progress_studio.domain import Activity


class ScheduleService:
    def roll_up_summary_dates(self, rows: list[Activity]) -> None:
        for index, row in enumerate(rows):
            if not row.is_summary:
                continue
            descendants: list[Activity] = []
            for child in rows[index + 1 :]:
                if child.outline_level <= row.outline_level:
                    break
                if not child.is_summary:
                    descendants.append(child)

            plan_starts = [child.plan_start for child in descendants if child.plan_start]
            plan_finishes = [child.plan_finish for child in descendants if child.plan_finish]
            actual_starts = [child.actual_start for child in descendants if child.actual_start]
            actual_finishes = [child.actual_finish for child in descendants if child.actual_finish]

            row.plan_start = min(plan_starts) if plan_starts else None
            row.plan_finish = max(plan_finishes) if plan_finishes else None
            row.actual_start = min(actual_starts) if actual_starts else None
            row.actual_finish = max(actual_finishes) if actual_finishes else None
