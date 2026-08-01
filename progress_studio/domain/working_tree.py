from __future__ import annotations

from dataclasses import dataclass, replace
from enum import StrEnum
from uuid import UUID, uuid4, uuid5

from progress_studio.domain.mapping_models import ActivityRow, SupplementalWBS


WORKING_TREE_NAMESPACE = UUID("c4e2ed4a-e6b6-4f46-8125-2f850f4fc78f")
WORKING_TREE_VERSION = 1


class WorkingNodeKind(StrEnum):
    WBS = "wbs"
    ACTIVITY = "activity"


class WorkingNodeOrigin(StrEnum):
    WORKBOOK = "workbook"
    USER_CREATED = "user_created"


@dataclass(frozen=True, slots=True)
class WorkingTreeNode:
    """One stable node in the editable working schedule tree.

    `node_id` is the identity used by the editor. Display codes, names, paths,
    and Activity IDs may change later without breaking mappings or history.
    """

    node_id: str
    kind: WorkingNodeKind
    parent_id: str | None
    code: str
    name: str
    origin: WorkingNodeOrigin
    order: int
    source_activity_id: str = ""
    source_path: tuple[tuple[str, str], ...] = ()
    deleted: bool = False

    @property
    def editable(self) -> bool:
        # MS11.1 establishes the model. MS11.2 will enable controlled edits
        # for workbook-origin nodes through commands and validation.
        return not self.deleted


class WorkingScheduleTree:
    """Unified projection of workbook and user-created schedule nodes.

    The tree deliberately separates stable identity from display values. It is
    rebuilt from workbook rows plus persisted user-created nodes, so the UI no
    longer needs two separate "original" and "supplemental" trees.
    """

    def __init__(self, nodes: list[WorkingTreeNode] | None = None) -> None:
        self._nodes: dict[str, WorkingTreeNode] = {}
        self._children: dict[str | None, list[str]] = {}
        for node in nodes or []:
            self._insert(node)

    @staticmethod
    def workbook_wbs_id(path: tuple[tuple[str, str], ...]) -> str:
        key = "/".join(f"{code}\x1f{name}" for code, name in path)
        return str(uuid5(WORKING_TREE_NAMESPACE, f"workbook:wbs:{key}"))

    @staticmethod
    def workbook_activity_id(activity_id: str) -> str:
        return str(uuid5(WORKING_TREE_NAMESPACE, f"workbook:activity:{activity_id.strip().upper()}"))

    @staticmethod
    def created_node_id(existing: str = "") -> str:
        return existing.strip() or str(uuid4())

    @staticmethod
    def legacy_created_node_id(kind: str, key: str) -> str:
        """Deterministic identity used while migrating pre-MS11 projects."""
        return str(uuid5(WORKING_TREE_NAMESPACE, f"legacy:{kind}:{key}"))

    @classmethod
    def build(
        cls,
        activities: list[ActivityRow],
        created_wbs: list[SupplementalWBS] | None = None,
    ) -> "WorkingScheduleTree":
        nodes: list[WorkingTreeNode] = []
        seen_paths: set[tuple[tuple[str, str], ...]] = set()
        order = 0

        for activity in activities:
            for level in range(1, len(activity.wbs_path) + 1):
                path = tuple(activity.wbs_path[:level])
                if path in seen_paths:
                    continue
                seen_paths.add(path)
                parent_path = path[:-1]
                code, name = path[-1]
                nodes.append(
                    WorkingTreeNode(
                        node_id=cls.workbook_wbs_id(path),
                        kind=WorkingNodeKind.WBS,
                        parent_id=cls.workbook_wbs_id(parent_path) if parent_path else None,
                        code=code,
                        name=name,
                        origin=WorkingNodeOrigin.WORKBOOK,
                        order=order,
                        source_path=path,
                    )
                )
                order += 1

        for wbs in created_wbs or []:
            parent_id = cls._resolve_parent_id(wbs.parent_path, nodes)
            nodes.append(
                WorkingTreeNode(
                    node_id=cls.created_node_id(wbs.node_id),
                    kind=WorkingNodeKind.WBS,
                    parent_id=parent_id,
                    code=wbs.code,
                    name=wbs.name,
                    origin=WorkingNodeOrigin.USER_CREATED,
                    order=order,
                    source_path=wbs.path,
                )
            )
            order += 1

        for activity in activities:
            parent_id = cls._resolve_parent_id(activity.wbs_path, nodes)
            origin = (
                WorkingNodeOrigin.USER_CREATED
                if activity.is_supplemental
                else WorkingNodeOrigin.WORKBOOK
            )
            node_id = (
                cls.created_node_id(activity.node_id)
                if activity.is_supplemental
                else cls.workbook_activity_id(activity.activity_id)
            )
            nodes.append(
                WorkingTreeNode(
                    node_id=node_id,
                    kind=WorkingNodeKind.ACTIVITY,
                    parent_id=parent_id,
                    code=activity.activity_id,
                    name=activity.description,
                    origin=origin,
                    order=order,
                    source_activity_id=activity.activity_id,
                    source_path=activity.wbs_path,
                )
            )
            order += 1

        return cls(nodes)

    @classmethod
    def _resolve_parent_id(
        cls,
        path: tuple[tuple[str, str], ...],
        nodes: list[WorkingTreeNode],
    ) -> str | None:
        if not path:
            return None
        # Prefer a created WBS with the exact path, otherwise use the stable
        # workbook WBS identity. This keeps descendants attached to the node
        # the user sees, regardless of node origin.
        for node in reversed(nodes):
            if node.kind is WorkingNodeKind.WBS and node.source_path == tuple(path):
                return node.node_id
        return cls.workbook_wbs_id(tuple(path))

    def _insert(self, node: WorkingTreeNode) -> None:
        if node.node_id in self._nodes:
            raise ValueError(f"Duplicate working-tree node ID: {node.node_id}")
        self._nodes[node.node_id] = node
        self._children.setdefault(node.parent_id, []).append(node.node_id)
        self._children.setdefault(node.node_id, [])

    def nodes(self, *, include_deleted: bool = False) -> tuple[WorkingTreeNode, ...]:
        values = sorted(self._nodes.values(), key=lambda item: item.order)
        if include_deleted:
            return tuple(values)
        return tuple(item for item in values if not item.deleted)

    def get(self, node_id: str) -> WorkingTreeNode | None:
        return self._nodes.get(node_id)

    def children(self, node_id: str | None) -> tuple[WorkingTreeNode, ...]:
        return tuple(
            self._nodes[child_id]
            for child_id in self._children.get(node_id, ())
            if not self._nodes[child_id].deleted
        )

    def find_wbs_by_path(
        self, path: tuple[tuple[str, str], ...]
    ) -> WorkingTreeNode | None:
        candidates = [
            node
            for node in self._nodes.values()
            if node.kind is WorkingNodeKind.WBS
            and node.source_path == tuple(path)
            and not node.deleted
        ]
        if not candidates:
            return None
        # A user-created node wins when both origins temporarily share a path.
        return sorted(
            candidates,
            key=lambda node: node.origin is WorkingNodeOrigin.USER_CREATED,
            reverse=True,
        )[0]

    def find_activity(self, activity_id: str) -> WorkingTreeNode | None:
        normalized = activity_id.strip().upper()
        return next(
            (
                node
                for node in self._nodes.values()
                if node.kind is WorkingNodeKind.ACTIVITY
                and node.source_activity_id.upper() == normalized
                and not node.deleted
            ),
            None,
        )

    def path_for(self, node_id: str) -> tuple[WorkingTreeNode, ...]:
        result: list[WorkingTreeNode] = []
        current = self.get(node_id)
        visited: set[str] = set()
        while current is not None:
            if current.node_id in visited:
                raise ValueError("Working tree contains a parent cycle.")
            visited.add(current.node_id)
            result.append(current)
            current = self.get(current.parent_id) if current.parent_id else None
        result.reverse()
        return tuple(result)

    def validate(self) -> tuple[str, ...]:
        errors: list[str] = []
        activity_codes: set[str] = set()
        for node in self._nodes.values():
            if node.parent_id is not None and node.parent_id not in self._nodes:
                errors.append(f"Orphan node: {node.node_id}")
            if node.kind is WorkingNodeKind.ACTIVITY:
                normalized = node.code.strip().upper()
                if normalized in activity_codes:
                    errors.append(f"Duplicate Activity ID: {node.code}")
                activity_codes.add(normalized)
            try:
                self.path_for(node.node_id)
            except ValueError as exc:
                errors.append(str(exc))
        return tuple(dict.fromkeys(errors))


    def snapshot(self) -> tuple[WorkingTreeNode, ...]:
        return tuple(replace(node) for node in self.nodes(include_deleted=True))

    def descendants(self, node_id: str, *, include_deleted: bool = False) -> tuple[WorkingTreeNode, ...]:
        result: list[WorkingTreeNode] = []
        pending = list(self._children.get(node_id, ()))
        while pending:
            child_id = pending.pop(0)
            child = self._nodes[child_id]
            if include_deleted or not child.deleted:
                result.append(child)
            pending[0:0] = list(self._children.get(child_id, ()))
        return tuple(result)

    def add_node(
        self,
        *,
        kind: WorkingNodeKind,
        parent_id: str | None,
        code: str,
        name: str,
        origin: WorkingNodeOrigin = WorkingNodeOrigin.USER_CREATED,
        source_activity_id: str = "",
    ) -> WorkingTreeNode:
        code, name = code.strip(), name.strip()
        if not code or not name:
            raise ValueError("Node code and name are required.")
        if parent_id is not None:
            parent = self.get(parent_id)
            if parent is None or parent.deleted or parent.kind is not WorkingNodeKind.WBS:
                raise ValueError("The selected parent WBS is not available.")
        if kind is WorkingNodeKind.ACTIVITY:
            normalized = code.upper()
            if any(
                node.kind is WorkingNodeKind.ACTIVITY
                and not node.deleted
                and node.code.upper() == normalized
                for node in self._nodes.values()
            ):
                raise ValueError(f"Activity ID already exists: {code}")
        elif any(
            node.kind is WorkingNodeKind.WBS
            and not node.deleted
            and node.code == code
            for node in self._nodes.values()
        ):
            raise ValueError(f"WBS code already exists: {code}")
        order = max((node.order for node in self._nodes.values()), default=-1) + 1
        node = WorkingTreeNode(
            node_id=self.created_node_id(),
            kind=kind,
            parent_id=parent_id,
            code=code,
            name=name,
            origin=origin,
            order=order,
            source_activity_id=source_activity_id or (code if kind is WorkingNodeKind.ACTIVITY else ""),
        )
        self._insert(node)
        return node

    def rename(self, node_id: str, *, code: str, name: str) -> WorkingTreeNode:
        node = self.get(node_id)
        if node is None or node.deleted:
            raise ValueError("The selected node is not available.")
        code, name = code.strip(), name.strip()
        if not code or not name:
            raise ValueError("Node code and name are required.")
        if node.kind is WorkingNodeKind.ACTIVITY:
            normalized = code.upper()
            if any(
                other.node_id != node_id
                and other.kind is WorkingNodeKind.ACTIVITY
                and not other.deleted
                and other.code.upper() == normalized
                for other in self._nodes.values()
            ):
                raise ValueError(f"Activity ID already exists: {code}")
        elif code != node.code and any(
            other.node_id != node_id
            and other.kind is WorkingNodeKind.WBS
            and not other.deleted
            and other.code == code
            for other in self._nodes.values()
        ):
            raise ValueError(f"WBS code already exists: {code}")
        updated = replace(node, code=code.upper() if node.kind is WorkingNodeKind.ACTIVITY else code, name=name)
        self.replace(updated)
        return updated

    def reparent(self, node_id: str, parent_id: str | None) -> WorkingTreeNode:
        node = self.get(node_id)
        if node is None or node.deleted:
            raise ValueError("The selected node is not available.")
        if parent_id == node_id:
            raise ValueError("A node cannot be its own parent.")
        if parent_id is not None:
            parent = self.get(parent_id)
            if parent is None or parent.deleted or parent.kind is not WorkingNodeKind.WBS:
                raise ValueError("The new parent must be an active WBS node.")
            if any(item.node_id == parent_id for item in self.descendants(node_id, include_deleted=True)):
                raise ValueError("A WBS cannot be moved below one of its descendants.")
        if node.kind is WorkingNodeKind.ACTIVITY and parent_id is None:
            raise ValueError("An Activity must belong to a WBS.")
        updated = replace(node, parent_id=parent_id)
        self.replace(updated)
        return updated

    def move_sibling(self, node_id: str, offset: int) -> bool:
        node = self.get(node_id)
        if node is None or node.deleted or offset == 0:
            return False
        siblings = [
            self._nodes[item_id]
            for item_id in self._children.get(node.parent_id, ())
            if not self._nodes[item_id].deleted
        ]
        siblings.sort(key=lambda item: item.order)
        index = next((i for i, item in enumerate(siblings) if item.node_id == node_id), -1)
        target = index + (-1 if offset < 0 else 1)
        if index < 0 or target < 0 or target >= len(siblings):
            return False
        other = siblings[target]
        self.replace(replace(node, order=other.order))
        self.replace(replace(other, order=node.order))
        return True

    def soft_delete(self, node_id: str) -> tuple[WorkingTreeNode, ...]:
        node = self.get(node_id)
        if node is None or node.deleted:
            raise ValueError("The selected node is not available.")
        affected = (node,) + self.descendants(node_id)
        for item in affected:
            self.replace(replace(item, deleted=True))
        return affected

    def replace(self, node: WorkingTreeNode) -> None:
        old = self._nodes.get(node.node_id)
        if old is None:
            raise KeyError(node.node_id)
        if old.parent_id != node.parent_id:
            self._children[old.parent_id].remove(node.node_id)
            self._children.setdefault(node.parent_id, []).append(node.node_id)
        self._nodes[node.node_id] = replace(node)
