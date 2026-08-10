
from __future__ import annotations

from collections import OrderedDict

from progress_studio.domain.main_dataset import MainDataset, MainRow
from progress_studio.domain.monthly_cache import (
    MonthlyArchitectureDecision,
    MonthlyCache,
    MonthlyPeriod,
    MonthlyRow,
)


def _month_buckets(dataset: MainDataset) -> list[tuple[tuple[int, int], list[int]]]:
    grouped: OrderedDict[tuple[int, int], list[int]] = OrderedDict()
    for period in dataset.periods:
        if period.reporting_date is None:
            continue
        key = (period.reporting_date.year, period.reporting_date.month)
        grouped.setdefault(key, []).append(period.column)
    return list(grouped.items())


class MonthlyCacheDeriver:
    """Materialize the Live monthly view as values derived once from MainDataset."""

    def derive(self, dataset: MainDataset) -> MonthlyCache:
        buckets = _month_buckets(dataset)
        periods: list[MonthlyPeriod] = []
        for index, (_, columns) in enumerate(buckets, start=1):
            last = next(
                (p for p in reversed(dataset.periods) if p.column == columns[-1]),
                None,
            )
            periods.append(
                MonthlyPeriod(
                    key=f"M{index}",
                    reporting_date=last.reporting_date if last else None,
                    source_columns=tuple(columns),
                )
            )

        rows: list[MonthlyRow] = []
        for row in dataset.rows:
            values: list[float | None] = []
            for period in periods:
                source_values = [row.period_value(col) for col in period.source_columns]
                numeric = [float(value) for value in source_values if value is not None]
                # Cumulative S-curve rows use the month's last reporting value.
                if row.row_type.strip().lower() == "s-curve" and row.pa.strip().upper() in {"AP", "AA"}:
                    last_value = row.period_value(period.source_columns[-1])
                    values.append(float(last_value) if last_value is not None else None)
                else:
                    values.append(sum(numeric) if numeric else None)
            rows.append(
                MonthlyRow(
                    source_row=row.row_number,
                    row_type=row.row_type,
                    pa=row.pa,
                    wbs=row.wbs,
                    description=row.description,
                    activity_id=row.activity_id,
                    outline_level=row.outline_level,
                    values=tuple(values),
                )
            )
        return MonthlyCache(periods=tuple(periods), rows=tuple(rows))


class MonthlyArchitectureEvaluator:
    """LW-6 decision gate.

    Formula and cache have the same visible monthly cell count, but formulas add
    workbook dependency edges and Excel recalculation work. Direct rendering has
    the smallest workbook payload but cannot preserve the user-visible Monthly
    worksheet contract. Therefore Live selects materialized cache values.
    """

    def evaluate(self, dataset: MainDataset) -> MonthlyArchitectureDecision:
        cache = MonthlyCacheDeriver().derive(dataset)
        visible_cells = cache.value_cell_count
        # Direct render needs only the project S-curve month points for the chart,
        # but it would remove the Monthly worksheet the product currently exposes.
        direct_cells = len(cache.periods) * 2
        return MonthlyArchitectureDecision(
            winner="cache",
            formula_cells=visible_cells,
            cache_value_cells=visible_cells,
            direct_render_cells=direct_cells,
            rationale=(
                "Cache preserves the Monthly worksheet contract.",
                "Cache removes Excel formula dependency edges from the monthly timescale.",
                "Direct render is smallest but would remove the user-visible Monthly view.",
                "Monthly cache is regenerated from MainDataset at the rebuild/save boundary.",
            ),
        )
