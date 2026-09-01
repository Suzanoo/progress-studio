# Progress Studio — EV-1 Calculation Engine

EV-1 is intentionally calculation-only. It does not render an Excel sheet, modify Dashboard, or integrate with standalone Rebuild.

## Frozen EV-0 contract implemented

- BAC = mapped BOQ amount.
- `main` weekly period values are increments; EV-1 accumulates them before calculating value.
- PV = BAC × cumulative Plan.
- EV = BAC × cumulative Actual.
- EV stops after cutoff; PV can continue to project finish.
- Margin periods outside the active Plan reporting range are excluded.
- Every positive-amount BOQ row must be allocated to 100% within 0.01 percentage point; otherwise EV calculation hard-stops.
- BOQ EV reverse-aggregates through `AllocationRecord`, preserving BOQ → Activity provenance.
- No AC/CV/CPI/EAC/ETC/VAC/TCPI in EV-1.

## Added files

- `progress_studio/domain/earned_value.py`
- `progress_studio/services/earned_value_deriver.py`
- `tests/test_earned_value_deriver.py`

## Test coverage in this delivery

1. Weekly increments accumulate before PV/EV calculation.
2. EV stops after cutoff while PV continues.
3. BOQ reverse aggregation reproduces the agreed 64 PV / 48 EV / -16 SV / 0.75 SPI example.
4. BOQ mapping below 100% is a hard stop.
5. Allocation rounding tolerance is accepted.
6. Unknown Activity mappings are rejected.
7. Left/right margin periods are excluded from EV reporting output.

`apply_ev1.ps1` performs the feature branch, targeted test, milestone commit/tag, and push automatically when run from a clean `main` working tree.

Git contract for this delivery:

- Feature branch: `feat/earned-value`
- EV-1 commit: `feat: EV-1 calculation engine`
- EV-1 tag: `ev-1-stable`
- Passed milestone: push branch and tag to `origin`
- Do **not** merge `main`; EV-2 and later EV milestones continue on the same feature branch.
