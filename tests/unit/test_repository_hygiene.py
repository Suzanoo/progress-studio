from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def test_removed_placeholder_and_orphan_modules_stay_absent() -> None:
    removed = [
        ROOT / "progress_studio/infrastructure/excel/workbook_writer.py",
        ROOT / "progress_studio/infrastructure/excel/workbook_reader.py",
        ROOT / "progress_studio/infrastructure/excel/formulas.py",
        ROOT / "progress_studio/services/dashboard_service.py",
    ]
    assert all(not path.exists() for path in removed)


def test_benchmarks_use_behavior_names_not_lw_milestone_names() -> None:
    scripts = ROOT / "scripts"
    assert not (scripts / "benchmark_lw6_monthly.py").exists()
    assert not (scripts / "benchmark_lw10_monthly.py").exists()
    assert (scripts / "benchmarks/monthly_parse.py").is_file()
    assert (scripts / "benchmarks/monthly_formula_metrics.py").is_file()


def test_example_golden_directory_contains_only_referenced_progress_fixture() -> None:
    golden = ROOT / "example/golden"
    assert (golden / "progress.xlsx").is_file()
    assert not (golden / "BOQ.xlsx").exists()


def test_packaging_boundary_and_package_check_exist() -> None:
    assert (ROOT / "packaging/windows/README.md").is_file()
    assert (ROOT / "packaging/macos/README.md").is_file()
    assert (ROOT / "scripts/check-package.py").is_file()
    assert (ROOT / "scripts/check-package.ps1").is_file()
