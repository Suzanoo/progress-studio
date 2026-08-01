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

    def replace(self, node: WorkingTreeNode) -> None:
        old = self._nodes.get(node.node_id)
        if old is None:
            raise KeyError(node.node_id)
        if old.parent_id != node.parent_id:
            self._children[old.parent_id].remove(node.node_id)
            self._children.setdefault(node.parent_id, []).append(node.node_id)
        self._nodes[node.node_id] = replace(node)
