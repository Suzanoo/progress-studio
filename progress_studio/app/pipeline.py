from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import Literal, Protocol

from .context import PipelineContext


class PipelineStep(Protocol):
    name: str

    def execute(self, context: PipelineContext) -> PipelineContext:
        ...


@dataclass(frozen=True)
class PipelineEvent:
    status: Literal["started", "completed", "failed"]
    step_name: str
    step_index: int
    step_count: int
    context: PipelineContext
    error: Exception | None = None

    @property
    def progress_percent(self) -> float:
        if self.status == "started":
            completed = self.step_index - 1
        else:
            completed = self.step_index
        return max(0.0, min(100.0, completed / self.step_count * 100.0))


class Pipeline:
    def __init__(self, steps: Sequence[PipelineStep]) -> None:
        self._steps = tuple(steps)

    @property
    def steps(self) -> tuple[PipelineStep, ...]:
        return self._steps

    def run(
        self,
        context: PipelineContext,
        observer: Callable[[PipelineEvent], None] | None = None,
    ) -> PipelineContext:
        current = context
        count = len(self._steps)
        for index, step in enumerate(self._steps, start=1):
            if observer:
                observer(PipelineEvent("started", step.name, index, count, current))
            try:
                current = step.execute(current)
            except Exception as exc:
                if observer:
                    observer(PipelineEvent("failed", step.name, index, count, current, exc))
                raise
            if observer:
                observer(PipelineEvent("completed", step.name, index, count, current))
        return current
