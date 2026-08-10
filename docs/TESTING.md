# Progress Studio Test Tiers

MS-TEST1 keeps every regression test, but routine development no longer runs the
entire historical suite for every small refactor.

## Tiers

### Smoke

Fast contract gate for local edits.

```powershell
python -m pytest -m smoke -q
```

Use before a quick commit or while iterating on one milestone.

### Active

Current workbook-generation, dashboard, payment, and rebuild regression suite.

```powershell
python -m pytest -m "smoke or active" -q
```

Use before handing off a milestone.

### Frozen

Stable legacy behavior: XML import, mapping core, working tree, persistent session,
older release architecture, and mature mapping UX.

Frozen does **not** mean deleted or ignored. These tests remain part of the release
gate. Do not rewrite a frozen test merely to make a new implementation pass. If a
frozen contract is intentionally retired, document the contract change first and
then reclassify/update the test deliberately.

Run frozen tests directly when touching a legacy subsystem:

```powershell
python -m pytest -m frozen -q
```

### Release

Every collected test is marked `release`.

```powershell
python -m pytest -m release
```

Equivalent full gate:

```powershell
python -m pytest
```

Run before merge to `main`, version tag, installer build, or production release.

## Classification rule

`tests/conftest.py` owns file-level classification. Any new `test_*.py` module that is
not explicitly classified as ACTIVE or FROZEN aborts collection. This prevents new
tests from silently falling outside the intended regression tiers.

A test may also be in `smoke`; smoke is an additional fast-gate marker, not a separate
ownership category.
