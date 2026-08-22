
from __future__ import annotations

import argparse
from pathlib import Path
from time import perf_counter

from progress_studio.infrastructure.excel.rebuild_workbook_reader import RebuildWorkbookReader
from progress_studio.services.monthly_cache_deriver import MonthlyArchitectureEvaluator, MonthlyCacheDeriver


def main() -> int:
    parser = argparse.ArgumentParser(description="LW-6 monthly architecture benchmark")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    reader = RebuildWorkbookReader()
    t0 = perf_counter()
    dataset = reader.read_main_dataset(args.workbook)
    t1 = perf_counter()
    cache = MonthlyCacheDeriver().derive(dataset)
    t2 = perf_counter()
    decision = MonthlyArchitectureEvaluator().evaluate(dataset)

    print(f"main_parse_seconds={t1-t0:.6f}")
    print(f"monthly_derive_seconds={t2-t1:.6f}")
    print(f"rows={len(cache.rows)}")
    print(f"months={len(cache.periods)}")
    print(f"formula_cells={decision.formula_cells}")
    print(f"cache_value_cells={decision.cache_value_cells}")
    print(f"direct_render_cells={decision.direct_render_cells}")
    print(f"winner={decision.winner}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
