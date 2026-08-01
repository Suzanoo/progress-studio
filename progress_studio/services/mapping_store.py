from __future__ import annotations

from dataclasses import dataclass

from progress_studio.domain.mapping_models import (
    ActivityRow,
    AllocationRecord,
    BOQRow,
    MappingChange,
    MappingStatus,
    SupplementalWBS,
)
from progress_studio.domain.working_tree import WorkingScheduleTree, WorkingTreeNode


@dataclass(frozen=True, slots=True)
class Page:
    ids: tuple[str, ...]
    number: int
    pages: int
    total: int
    start: int
    end: int


class MappingStore:
    """In-memory source of truth for the mapping screen."""

    EPSILON = 1e-9

    def __init__(self, activity_page_size: int = 200, boq_page_size: int = 300) -> None:
        self.activity_page_size = activity_page_size
        self.boq_page_size = boq_page_size
        self.activities_by_id: dict[str, ActivityRow] = {}
        self.activity_order: list[str] = []
        self.boq_by_id: dict[str, BOQRow] = {}
        self.boq_order: list[str] = []
        # (BOQ key, Activity ID) -> percentage of original BOQ amount.
        self.allocations: dict[tuple[str, str], float] = {}
        self.selected_activity_ids: set[str] = set()
        self.selected_boq_ids: set[str] = set()
        self._undo_stack: list[dict[tuple[str, str], float | None]] = []
        self.activity_page = 1
        self.boq_page = 1
        self.activity_query = ""
        self.boq_query = ""
        self.boq_wbs2 = ""
        self.boq_wbs3 = ""
        self._boq_ids_by_wbs2: dict[str, list[str]] = {}
        self._boq_ids_by_wbs23: dict[tuple[str, str], list[str]] = {}
        self.boq_selection_anchor: str | None = None
        self.collapsed_wbs_paths: set[tuple[str, ...]] = set()
        self.supplemental_wbs_nodes: list[SupplementalWBS] = []
        self.selected_wbs_path: tuple[tuple[str, str], ...] = ()
        self.selected_node_id: str = ""
        self.working_tree = WorkingScheduleTree()

    def load_activities(self, rows: list[ActivityRow]) -> None:
        self.activities_by_id = {row.activity_id: row for row in rows}
        self.activity_order = [row.activity_id for row in rows]
        self.allocations.clear()
        self.selected_activity_ids.clear()
        self.selected_boq_ids.clear()
        self._undo_stack.clear()
        self.activity_page = 1
        self.collapsed_wbs_paths.clear()
        self.selected_wbs_path = ()
        self.selected_node_id = ""
        self.supplemental_wbs_nodes = []
        self.working_tree = WorkingScheduleTree.build(list(self.activities_by_id.values()), [])

    def _rebuild_working_tree(self) -> None:
        self.working_tree = WorkingScheduleTree.build(
            [self.activities_by_id[key] for key in self.activity_order],
            self.supplemental_wbs_nodes,
        )
        if self.selected_wbs_path:
            selected = self.working_tree.find_wbs_by_path(self.selected_wbs_path)
            self.selected_node_id = selected.node_id if selected else ""
        elif self.selected_activity_ids:
            activity_id = next(iter(self.selected_activity_ids))
            selected = self.working_tree.find_activity(activity_id)
            self.selected_node_id = selected.node_id if selected else ""
        else:
            self.selected_node_id = ""

    def working_tree_nodes(self) -> tuple[WorkingTreeNode, ...]:
        return self.working_tree.nodes()

    def all_wbs_paths(self) -> list[tuple[tuple[str, str], ...]]:
        paths = {tuple(row.wbs_path[:level]) for row in self.activities_by_id.values() for level in range(1, len(row.wbs_path)+1)}
        paths.update(node.path for node in self.supplemental_wbs_nodes)
        return sorted(paths, key=lambda path: (len(path), tuple(code for code, _ in path)))

    def select_wbs(self, path: tuple[tuple[str, str], ...]) -> None:
        self.selected_wbs_path = tuple(path)
        node = self.working_tree.find_wbs_by_path(self.selected_wbs_path)
        self.selected_node_id = node.node_id if node else ""

    def add_supplemental_wbs(self, *, parent_path: tuple[tuple[str, str], ...], code: str, name: str) -> SupplementalWBS:
        code, name = code.strip(), name.strip()
        if not code or not name:
            raise ValueError("WBS code and WBS name are required.")
        if any(code == item_code for path in self.all_wbs_paths() for item_code, _ in path):
            raise ValueError(f"WBS code already exists: {code}")
        if parent_path and tuple(parent_path) not in self.all_wbs_paths():
            raise ValueError("The selected parent WBS no longer exists.")
        node = SupplementalWBS(
            code=code,
            name=name,
            parent_path=tuple(parent_path),
            node_id=WorkingScheduleTree.created_node_id(),
        )
        self.supplemental_wbs_nodes.append(node)
        self.selected_wbs_path = node.path
        self._rebuild_working_tree()
        return node

    def edit_supplemental_wbs(self, path: tuple[tuple[str, str], ...], *, code: str, name: str) -> SupplementalWBS:
        index = next((i for i, node in enumerate(self.supplemental_wbs_nodes) if node.path == tuple(path)), -1)
        if index < 0:
            raise ValueError("Only WBS nodes created in Progress Studio can be edited.")
        old = self.supplemental_wbs_nodes[index]
        code, name = code.strip(), name.strip()
        if not code or not name:
            raise ValueError("WBS code and WBS name are required.")
        if code != old.code and any(code == item_code for candidate in self.all_wbs_paths() for item_code, _ in candidate):
            raise ValueError(f"WBS code already exists: {code}")
        replacement = SupplementalWBS(
            code=code,
            name=name,
            parent_path=old.parent_path,
            origin=old.origin,
            node_id=old.node_id or WorkingScheduleTree.created_node_id(),
        )
        old_path, new_path = old.path, replacement.path
        self.supplemental_wbs_nodes[index] = replacement
        for i, node in enumerate(list(self.supplemental_wbs_nodes)):
            if node.parent_path[:len(old_path)] == old_path:
                self.supplemental_wbs_nodes[i] = SupplementalWBS(
                    node.code, node.name, new_path + node.parent_path[len(old_path):],
                    node.origin, node.node_id,
                )
        for activity_id, row in list(self.activities_by_id.items()):
            if row.wbs_path[:len(old_path)] == old_path:
                from dataclasses import replace
                self.activities_by_id[activity_id] = replace(row, wbs_path=new_path + row.wbs_path[len(old_path):], parent_wbs=(new_path + row.wbs_path[len(old_path):])[-1][0])
        self.selected_wbs_path = new_path
        self._rebuild_working_tree()
        return replacement

    def remove_supplemental_wbs(self, path: tuple[tuple[str, str], ...]) -> None:
        node = next((node for node in self.supplemental_wbs_nodes if node.path == tuple(path)), None)
        if node is None:
            raise ValueError("Only WBS nodes created in Progress Studio can be removed.")
        if any(other.parent_path[:len(path)] == tuple(path) for other in self.supplemental_wbs_nodes if other is not node):
            raise ValueError("Remove child WBS nodes first.")
        if any(row.wbs_path[:len(path)] == tuple(path) for row in self.activities_by_id.values()):
            raise ValueError("Remove or move Activities under this WBS first.")
        self.supplemental_wbs_nodes.remove(node)
        self.selected_wbs_path = node.parent_path
        self._rebuild_working_tree()

    def add_supplemental_activity(
        self, *, parent_path: tuple[tuple[str, str], ...], wbs_code: str,
        wbs_name: str, activity_id: str, description: str
    ) -> ActivityRow:
        activity_id = activity_id.strip().upper()
        wbs_code = wbs_code.strip()
        wbs_name = wbs_name.strip()
        description = description.strip()
        if not activity_id or not description or not wbs_code or not wbs_name:
            raise ValueError("WBS code, WBS name, Activity ID, and Activity name are required.")
        if activity_id in self.activities_by_id:
            raise ValueError(f"Activity ID already exists: {activity_id}")
        # Activities are created *under* an existing WBS, so the selected
        # parent WBS code is expected to already exist.  Uniqueness belongs
        # to Activity ID, not to the parent WBS code.
        if parent_path and tuple(parent_path) not in self.all_wbs_paths():
            raise ValueError("The selected parent WBS no longer exists.")
        # Backward-compatible behavior: callers may either pass the selected
        # WBS itself (create Activity directly under it), or pass a new WBS
        # code/name to create the Activity under that supplemental leaf.
        direct_parent = bool(parent_path) and wbs_code == parent_path[-1][0]
        activity_path = tuple(parent_path) if direct_parent else tuple(parent_path) + ((wbs_code, wbs_name),)
        row = ActivityRow(
            activity_id=activity_id,
            parent_wbs=activity_path[-1][0] if activity_path else "",
            child_wbs=activity_path[-1][0] if activity_path else wbs_code,
            description=description,
            wbs_path=activity_path,
            origin="user_created",
            supplemental_wbs_code=activity_path[-1][0] if activity_path else wbs_code,
            supplemental_wbs_name=activity_path[-1][1] if activity_path else wbs_name,
            node_id=WorkingScheduleTree.created_node_id(),
        )
        self.activities_by_id[activity_id] = row
        self.activity_order.append(activity_id)
        self.selected_activity_ids = {activity_id}
        self.activity_page = 1
        self._rebuild_working_tree()
        return row

    def supplemental_activities(self) -> list[ActivityRow]:
        return [self.activities_by_id[key] for key in self.activity_order if self.activities_by_id[key].is_supplemental]

    def remove_supplemental_activity(self, activity_id: str) -> None:
        row = self.activities_by_id.get(activity_id)
        if row is None or not row.is_supplemental:
            raise ValueError("Only activities created in Progress Studio can be removed.")
        if any(assigned == activity_id for _, assigned in self.allocations):
            raise ValueError("Unmap all BOQ items from this Activity before removing it.")
        self.activities_by_id.pop(activity_id)
        self.activity_order.remove(activity_id)
        self.selected_activity_ids.discard(activity_id)
        self._rebuild_working_tree()

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
        self.allocations.clear()
        self.selected_boq_ids.clear()
        self.boq_selection_anchor = None
        self._undo_stack.clear()
        self.boq_page = 1

    @staticmethod
    def wbs_path_key(path: tuple[tuple[str, str], ...], level: int | None = None) -> tuple[str, ...]:
        items = path if level is None else path[:level]
        return tuple(code for code, _name in items)

    def toggle_wbs(self, path_key: tuple[str, ...]) -> None:
        if path_key in self.collapsed_wbs_paths:
            self.collapsed_wbs_paths.remove(path_key)
        else:
            self.collapsed_wbs_paths.add(path_key)
        self.activity_page = 1

    def collapse_all_wbs(self) -> None:
        self.collapsed_wbs_paths = {
            self.wbs_path_key(row.wbs_path, level)
            for row in self.activities_by_id.values()
            for level in range(1, len(row.wbs_path) + 1)
        }
        self.activity_page = 1

    def expand_all_wbs(self) -> None:
        self.collapsed_wbs_paths.clear()
        self.activity_page = 1

    def is_activity_visible(self, activity_id: str) -> bool:
        row = self.activities_by_id[activity_id]
        return not any(
            self.wbs_path_key(row.wbs_path, level) in self.collapsed_wbs_paths
            for level in range(1, len(row.wbs_path) + 1)
        )

    def _filtered_activity_ids(self) -> list[str]:
        query = self.activity_query.strip().lower()
        candidates = self.activity_order if not query else [
            key for key in self.activity_order
            if query in self.activities_by_id[key].search_text
        ]
        # Search results always expand their parent context. In normal browsing,
        # retain one representative Activity per collapsed subtree so the UI
        # can still render the WBS header while hiding its descendants.
        if query or not self.collapsed_wbs_paths:
            return list(candidates)
        result: list[str] = []
        represented: set[tuple[str, ...]] = set()
        for key in candidates:
            row = self.activities_by_id[key]
            collapsed_prefix = next((
                self.wbs_path_key(row.wbs_path, level)
                for level in range(1, len(row.wbs_path) + 1)
                if self.wbs_path_key(row.wbs_path, level) in self.collapsed_wbs_paths
            ), None)
            if collapsed_prefix is None:
                result.append(key)
            elif collapsed_prefix not in represented:
                represented.add(collapsed_prefix)
                result.append(key)
        return result

    def boq_wbs2_values(self) -> tuple[str, ...]:
        return tuple(sorted(value for value in self._boq_ids_by_wbs2 if value))

    def boq_wbs3_values(self, wbs2: str = "") -> tuple[str, ...]:
        if wbs2:
            values = {
                wbs3 for (item_wbs2, wbs3) in self._boq_ids_by_wbs23
                if item_wbs2 == wbs2 and wbs3
            }
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
        return Page(tuple(ids[start_index:end_index]), page, pages, total, 0 if total == 0 else start_index + 1, end_index)

    def activity_page_data(self) -> Page:
        page = self._page(self._filtered_activity_ids(), self.activity_page, self.activity_page_size)
        self.activity_page = page.number
        return page

    def boq_page_data(self) -> Page:
        page = self._page(self._filtered_boq_ids(), self.boq_page, self.boq_page_size)
        self.boq_page = page.number
        return page

    def toggle_activity(self, activity_id: str) -> None:
        if activity_id in self.selected_activity_ids:
            self.selected_activity_ids.clear()
            self.selected_node_id = ""
        else:
            self.selected_activity_ids = {activity_id}
            self.selected_wbs_path = ()
            node = self.working_tree.find_activity(activity_id)
            self.selected_node_id = node.node_id if node else ""

    def toggle_boq(self, key: str, *, additive: bool = True) -> None:
        if not additive:
            self.selected_boq_ids.clear()
        if key in self.selected_boq_ids:
            self.selected_boq_ids.remove(key)
        else:
            self.selected_boq_ids.add(key)
        self.boq_selection_anchor = key

    def select_boq_range(self, key: str) -> tuple[str, ...]:
        filtered = self._filtered_boq_ids()
        if key not in filtered:
            return ()
        if self.boq_selection_anchor not in filtered:
            self.toggle_boq(key)
            return (key,)
        start = filtered.index(self.boq_selection_anchor)
        end = filtered.index(key)
        selected = tuple(filtered[min(start, end): max(start, end) + 1])
        self.selected_boq_ids.update(selected)
        return selected

    def select_boq_page(self) -> tuple[str, ...]:
        ids = self.boq_page_data().ids
        self.selected_boq_ids.update(ids)
        if ids:
            self.boq_selection_anchor = ids[-1]
        return ids

    def select_all_filtered_boq(self) -> tuple[str, ...]:
        ids = tuple(self._filtered_boq_ids())
        self.selected_boq_ids.update(ids)
        if ids:
            self.boq_selection_anchor = ids[-1]
        return ids

    def clear_boq_selection(self) -> None:
        self.selected_boq_ids.clear()
        self.boq_selection_anchor = None

    @property
    def selected_boq_amount(self) -> float:
        return sum(self.boq_by_id[key].amount for key in self.selected_boq_ids if key in self.boq_by_id)

    @staticmethod
    def _validate_share(share_percent: float) -> float:
        try:
            share = float(share_percent)
        except (TypeError, ValueError) as exc:
            raise ValueError("Share must be a number greater than 0 and not more than 100.") from exc
        if share <= 0 or share > 100:
            raise ValueError("Share must be greater than 0 and not more than 100.")
        return share

    def boq_share_percent(self, key: str) -> float:
        return sum(share for (boq_key, _), share in self.allocations.items() if boq_key == key)

    def allocation_share(self, key: str, activity_id: str) -> float:
        return self.allocations.get((key, activity_id), 0.0)

    def allocation_records(self) -> list[AllocationRecord]:
        return [
            AllocationRecord(key, activity_id, share)
            for (key, activity_id), share in sorted(self.allocations.items())
        ]

    def mapped_activities(self, key: str) -> tuple[str, ...]:
        return tuple(sorted(activity_id for (boq_key, activity_id) in self.allocations if boq_key == key))

    def mapped_to_text(self, key: str) -> str:
        parts = [
            f"{activity_id} ({self.allocations[(key, activity_id)]:g}%)"
            for activity_id in self.mapped_activities(key)
        ]
        return ", ".join(parts)

    def mapped_to_compact_text(self, key: str) -> str:
        """Return a table-friendly mapping summary while retaining full detail elsewhere."""
        activities = self.mapped_activities(key)
        if not activities:
            return "—"
        first = activities[0]
        first_share = self.allocations[(key, first)]
        suffix = f" +{len(activities) - 1}" if len(activities) > 1 else ""
        return f"{first} ({first_share:g}%){suffix}"

    def map_selected(self, share_percent: float = 100.0) -> MappingChange:
        if len(self.selected_activity_ids) != 1:
            raise ValueError("Select exactly one Activity.")
        if not self.selected_boq_ids:
            raise ValueError("Select one or more BOQ items.")
        share = self._validate_share(share_percent)
        activity_id = next(iter(self.selected_activity_ids))
        keys = tuple(sorted(self.selected_boq_ids))
        changed_pairs = {(key, activity_id) for key in keys}
        previous = {pair: self.allocations.get(pair) for pair in changed_pairs}

        for key in keys:
            existing_other = self.boq_share_percent(key) - self.allocation_share(key, activity_id)
            if existing_other + share > 100.0 + self.EPSILON:
                remaining = max(0.0, 100.0 - existing_other)
                raise ValueError(
                    f"{key} has only {remaining:g}% remaining. "
                    f"Requested share for {activity_id}: {share:g}%."
                )

        self._undo_stack.append(previous)
        for pair in changed_pairs:
            self.allocations[pair] = share
        self.selected_boq_ids.clear()
        return MappingChange(keys, (activity_id,))

    def unmap_selected(self) -> MappingChange:
        if len(self.selected_activity_ids) != 1:
            raise ValueError("Select exactly one Activity to unmap.")
        keys = tuple(sorted(self.selected_boq_ids))
        if not keys:
            raise ValueError("Select one or more BOQ items.")
        activity_id = next(iter(self.selected_activity_ids))
        pairs = {(key, activity_id) for key in keys}
        previous = {pair: self.allocations.get(pair) for pair in pairs}
        if not any(value is not None for value in previous.values()):
            raise ValueError("The selected BOQ items are not mapped to this Activity.")
        self._undo_stack.append(previous)
        for pair in pairs:
            self.allocations.pop(pair, None)
        self.selected_boq_ids.clear()
        return MappingChange(keys, (activity_id,))

    def undo(self) -> MappingChange | None:
        if not self._undo_stack:
            return None
        previous = self._undo_stack.pop()
        affected_boq = {key for key, _ in previous}
        affected_activities = {activity_id for _, activity_id in previous}
        for pair, share in previous.items():
            if share is None:
                self.allocations.pop(pair, None)
            else:
                self.allocations[pair] = share
        return MappingChange(tuple(sorted(affected_boq)), tuple(sorted(affected_activities)))


    def restore_allocations(self, records: list[AllocationRecord]) -> MappingChange:
        """Replace all allocations after validating workbook identifiers and shares."""
        restored: dict[tuple[str, str], float] = {}
        totals: dict[str, float] = {}
        for record in records:
            if record.boq_key not in self.boq_by_id:
                raise ValueError(f"Session BOQ item was not found: {record.boq_key}")
            if record.activity_id not in self.activities_by_id:
                raise ValueError(f"Session Activity was not found: {record.activity_id}")
            share = self._validate_share(record.share_percent)
            pair = (record.boq_key, record.activity_id)
            if pair in restored:
                raise ValueError(
                    f"Session contains a duplicate allocation: {record.boq_key} -> {record.activity_id}"
                )
            totals[record.boq_key] = totals.get(record.boq_key, 0.0) + share
            if totals[record.boq_key] > 100.0 + self.EPSILON:
                raise ValueError(f"Session allocation exceeds 100% for {record.boq_key}.")
            restored[pair] = share

        affected_boq = {key for key, _ in self.allocations} | {key for key, _ in restored}
        affected_activities = {activity for _, activity in self.allocations} | {activity for _, activity in restored}
        self.allocations = restored
        self.selected_activity_ids.clear()
        self.selected_boq_ids.clear()
        self._undo_stack.clear()
        return MappingChange(tuple(sorted(affected_boq)), tuple(sorted(affected_activities)))

    def clear_all(self) -> MappingChange:
        """Clear every allocation as one undoable command."""
        if not self.allocations:
            raise ValueError("There are no mappings to clear.")
        previous = {pair: share for pair, share in self.allocations.items()}
        self._undo_stack.append(previous)
        affected_boq = tuple(sorted({key for key, _ in self.allocations}))
        affected_activities = tuple(sorted({activity for _, activity in self.allocations}))
        self.allocations.clear()
        self.selected_boq_ids.clear()
        return MappingChange(affected_boq, affected_activities)

    def activity_amount(self, activity_id: str) -> float:
        return sum(
            self.boq_by_id[key].amount * share / 100.0
            for (key, assigned_id), share in self.allocations.items()
            if assigned_id == activity_id and key in self.boq_by_id
        )

    def boq_allocated_amount(self, key: str) -> float:
        row = self.boq_by_id.get(key)
        return 0.0 if row is None else row.amount * self.boq_share_percent(key) / 100.0

    def boq_remaining_amount(self, key: str) -> float:
        row = self.boq_by_id.get(key)
        return 0.0 if row is None else max(0.0, row.amount - self.boq_allocated_amount(key))

    def boq_remaining_percent(self, key: str) -> float:
        """Return the unallocated BOQ share as a percentage from 0 to 100."""
        if key not in self.boq_by_id:
            return 0.0
        return max(0.0, min(100.0, 100.0 - self.boq_share_percent(key)))

    def boq_status(self, key: str) -> MappingStatus:
        share = self.boq_share_percent(key)
        if share <= self.EPSILON:
            return MappingStatus.UNMAPPED
        if share < 100.0 - self.EPSILON:
            return MappingStatus.PARTIAL
        return MappingStatus.FULL

    @property
    def total_amount(self) -> float:
        return sum(row.amount for row in self.boq_by_id.values())

    @property
    def mapped_amount(self) -> float:
        return sum(self.boq_allocated_amount(key) for key in self.boq_order)

    @property
    def remaining_amount(self) -> float:
        return max(0.0, self.total_amount - self.mapped_amount)

    @property
    def mapped_item_count(self) -> int:
        return sum(1 for key in self.boq_order if self.boq_share_percent(key) > self.EPSILON)
