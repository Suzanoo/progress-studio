from __future__ import annotations

from dataclasses import dataclass

from progress_studio.domain.mapping_models import ActivityRow, BOQRow


@dataclass(frozen=True, slots=True)
class Page:
    ids: tuple[str, ...]
    number: int
    pages: int
    total: int
    start: int
    end: int


class MappingStore:
    def __init__(self, activity_page_size: int = 200, boq_page_size: int = 300) -> None:
        self.activity_page_size = activity_page_size
        self.boq_page_size = boq_page_size
        self.activities_by_id: dict[str, ActivityRow] = {}
        self.activity_order: list[str] = []
        self.boq_by_id: dict[str, BOQRow] = {}
        self.boq_order: list[str] = []
        self.assignments: dict[str, str] = {}
        self.selected_activity_ids: set[str] = set()
        self.selected_boq_ids: set[str] = set()
        self._undo_stack: list[dict[str, str | None]] = []
        self.activity_page = 1
        self.boq_page = 1
        self.activity_query = ""
        self.boq_query = ""
        self.boq_wbs2 = ""
        self.boq_wbs3 = ""
        self._boq_ids_by_wbs2: dict[str, list[str]] = {}
        self._boq_ids_by_wbs23: dict[tuple[str, str], list[str]] = {}

    def load_activities(self, rows: list[ActivityRow]) -> None:
        self.activities_by_id = {row.activity_id: row for row in rows}
        self.activity_order = [row.activity_id for row in rows]
        self.selected_activity_ids.clear()
        self.activity_page = 1

    def load_boq(self, rows: list[BOQRow]) -> None:
        self.boq_by_id = {row.key: row for row in rows}
        self.boq_order = [row.key for row in rows]
        self._boq_ids_by_wbs2 = {}
        self._boq_ids_by_wbs23 = {}
        for row in rows:
            self._boq_ids_by_wbs2.setdefault(row.wbs2, []).append(row.key)
            self._boq_ids_by_wbs23.setdefault((row.wbs2, row.wbs3), []).append(row.key)
        self.boq_wbs2 = ""
        self.boq_wbs3 = ""
        self.assignments.clear()
        self.selected_boq_ids.clear()
        self._undo_stack.clear()
        self.boq_page = 1

    def _filtered_activity_ids(self) -> list[str]:
        query = self.activity_query.strip().lower()
        if not query:
            return self.activity_order
        return [
            key
            for key in self.activity_order
            if query in self.activities_by_id[key].search_text
        ]

    def boq_wbs2_values(self) -> tuple[str, ...]:
        return tuple(sorted(value for value in self._boq_ids_by_wbs2 if value))

    def boq_wbs3_values(self, wbs2: str = "") -> tuple[str, ...]:
        if wbs2:
            values = {wbs3 for (item_wbs2, wbs3) in self._boq_ids_by_wbs23 if item_wbs2 == wbs2 and wbs3}
        else:
            values = {row.wbs3 for row in self.boq_by_id.values() if row.wbs3}
        return tuple(sorted(values))

    def _filtered_boq_ids(self) -> list[str]:
        query = self.boq_query.strip().lower()
        if self.boq_wbs2 and self.boq_wbs3:
            candidates = self._boq_ids_by_wbs23.get((self.boq_wbs2, self.boq_wbs3), [])
        elif self.boq_wbs2:
            candidates = self._boq_ids_by_wbs2.get(self.boq_wbs2, [])
        elif self.boq_wbs3:
            candidates = [key for key in self.boq_order if self.boq_by_id[key].wbs3 == self.boq_wbs3]
        else:
            candidates = self.boq_order
        if not query:
            return list(candidates)
        return [key for key in candidates if query in self.boq_by_id[key].search_text]

    @staticmethod
    def _page(ids: list[str], page: int, page_size: int) -> Page:
        total = len(ids)
        pages = max(1, (total + page_size - 1) // page_size)
        page = max(1, min(page, pages))
        start_index = (page - 1) * page_size
        end_index = min(start_index + page_size, total)
        return Page(
            ids=tuple(ids[start_index:end_index]),
            number=page,
            pages=pages,
            total=total,
            start=0 if total == 0 else start_index + 1,
            end=end_index,
        )

    def activity_page_data(self) -> Page:
        page = self._page(self._filtered_activity_ids(), self.activity_page, self.activity_page_size)
        self.activity_page = page.number
        return page

    def boq_page_data(self) -> Page:
        page = self._page(self._filtered_boq_ids(), self.boq_page, self.boq_page_size)
        self.boq_page = page.number
        return page

    def toggle_activity(self, activity_id: str) -> None:
        # Mapping V1 allows exactly one selected Activity.
        if activity_id in self.selected_activity_ids:
            self.selected_activity_ids.clear()
        else:
            self.selected_activity_ids = {activity_id}

    def toggle_boq(self, key: str) -> None:
        if key in self.selected_boq_ids:
            self.selected_boq_ids.remove(key)
        else:
            self.selected_boq_ids.add(key)

    def map_selected(self) -> tuple[str, tuple[str, ...]]:
        if len(self.selected_activity_ids) != 1:
            raise ValueError("Select exactly one Activity.")
        if not self.selected_boq_ids:
            raise ValueError("Select one or more BOQ items.")
        activity_id = next(iter(self.selected_activity_ids))
        keys = tuple(self.selected_boq_ids)
        previous = {key: self.assignments.get(key) for key in keys}
        self._undo_stack.append(previous)
        for key in keys:
            self.assignments[key] = activity_id
        self.selected_boq_ids.clear()
        return activity_id, keys

    def unmap_selected(self) -> tuple[str, ...]:
        keys = tuple(self.selected_boq_ids)
        if not keys:
            raise ValueError("Select one or more BOQ items.")
        previous = {key: self.assignments.get(key) for key in keys}
        self._undo_stack.append(previous)
        for key in keys:
            self.assignments.pop(key, None)
        self.selected_boq_ids.clear()
        return keys

    def undo(self) -> bool:
        if not self._undo_stack:
            return False
        previous = self._undo_stack.pop()
        for key, activity_id in previous.items():
            if activity_id is None:
                self.assignments.pop(key, None)
            else:
                self.assignments[key] = activity_id
        return True

    def clear_all(self) -> None:
        if self.assignments:
            self._undo_stack.append(dict(self.assignments))
        self.assignments.clear()
        self.selected_boq_ids.clear()

    def activity_amount(self, activity_id: str) -> float:
        return sum(
            self.boq_by_id[key].amount
            for key, assigned_id in self.assignments.items()
            if assigned_id == activity_id and key in self.boq_by_id
        )

    @property
    def total_amount(self) -> float:
        return sum(row.amount for row in self.boq_by_id.values())

    @property
    def mapped_amount(self) -> float:
        return sum(self.boq_by_id[key].amount for key in self.assignments if key in self.boq_by_id)
