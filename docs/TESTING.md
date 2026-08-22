# Progress Studio Test Strategy

The active test suite is organized by product behavior rather than historical milestone names. Test location communicates what a test protects and when it should run.

## Test layout

```text
tests/
├─ unit/         isolated parser/domain/helper behavior
├─ integration/  Create / Mapping / Payment / Rebuild / Desktop boundaries
├─ regression/   product bugs and contracts that must not return
├─ acceptance/   release-level product acceptance contracts
└─ fixtures/     shared test data
```

Historical milestone names such as `N7`, `LW10`, `MS-RB3`, and `WP2` are no longer
used as the primary organization mechanism. Git history preserves when a contract
was introduced; the test tree describes what the product must do today.

## Profiles

### Unit

Fast isolated logic tests.

```powershell
.\scripts\test-unit.ps1 -q
# or
python -m pytest -m unit -q
```

Use continuously while changing a parser, model, formula helper, or small service.

### Smoke

Small high-value handoff gate. Smoke is a subset of the other profiles and answers:
"Is this build safe enough to hand to a user for a quick real-workbook check?"

```powershell
.\scripts\test-smoke.ps1 -q
```

Run before every ZIP handoff.

### Regression

Known product contracts and bugs that must never return: reporting margins, cutoff,
chart OOXML integrity, overlay geometry, workbook protection, recalculation, etc.

```powershell
.\scripts\test-regression.ps1 -q
```

Run after bug fixes and before merging a subsystem change.

### Integration

Checks component/workflow boundaries such as Create Progress, Mapping, Payment,
Rebuild, and Desktop routing.

```powershell
.\scripts\test-integration.ps1 -q
```

Run when changing an engine boundary or workflow.

### Acceptance

Release-level product contracts and representative end-to-end behavior.

```powershell
.\scripts\test-acceptance.ps1 -q
```

Acceptance can be slower. It is not required for every small local edit.

### Release / Full test

Every collected test is marked `release`.

```powershell
.\scripts\test-release.ps1
# equivalent
python -m pytest
```

Required before a release candidate, installer build, production tag, or merge that
changes architecture. It is intentionally not the inner development loop.

## Normal development loop

For a small bug:

```text
focused test → regression → smoke → handoff
```

For a workflow/engine change:

```text
focused test → integration → regression → smoke → handoff
```

For a production milestone:

```text
unit + integration + regression + acceptance → full/release → tag
```

## Adding tests

Place a new test in one primary directory. `tests/conftest.py` automatically assigns
its primary marker from that directory and marks every test as `release`. Collection
aborts if a test module is placed outside `unit`, `integration`, `regression`, or
`acceptance`.

Do not create new milestone-named modules such as `test_p3_2_hotfix.py`. Name the file
after the behavior, for example `test_dashboard_reporting_range.py`.
