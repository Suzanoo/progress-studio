from collections import Counter
from dataclasses import dataclass, field


@dataclass(frozen=True)
class DistributionResult:
    method: str
    generated: int
    skipped_no_dates: int
    skipped_outside_range: int
    method_counts: Counter[str] = field(default_factory=Counter)
