# Dashboard V1 Freeze — LW-13

Status: **FROZEN / STABLE**  
Baseline tag: `dashboard-v1-stable`

LW-13 does not lock the Excel sheet and does not prevent future Dashboard work.
It freezes the accepted **behavior contract** so later Rebuild Lite / workbook
performance work cannot silently change the Dashboard.

## Frozen behavior

- Dashboard controls stay at `G5` (View) and `K5` (Cutoff Date).
- Traditional Weekly and Monthly overlays each own their **independent** cutoff
  selector in column `M`.
- Weekly cutoff display is `dd/mm/yyyy`.
- Monthly cutoff display is `mmm yyyy`.
- Plan curve renders the full project timeline.
- Actual curve is presentation-masked after the selected cutoff; historical
  Actual data is not rewritten by changing the cutoff.
- Weekly, Monthly, and Dashboard cutoff values do not drive each other.
- Traditional overlay cutoff is a red dashed vertical line with a readable
  10 pt label and no chart legend.
- Rebuild continues to regenerate the Dashboard and traditional overlays from
  the current calculation/data contracts.

## Change rule

A refactor may change implementation freely **only while the LW-13 regression
gate passes**.  If product behavior above must change intentionally:

1. create a new Dashboard contract version (for example `dashboard-v2`),
2. update/add regression tests for the new behavior,
3. document the migration/change in `CHANGELOG.md`, and
4. create a new stable Git tag after QA.

Do not weaken or delete the V1 tests merely to make an unrelated refactor pass.

## Release gate

Run:

```bash
pytest -q tests/test_lw13_dashboard_freeze.py
pytest -q
```

The first command is the fast Dashboard freeze gate.  The second is the full
Progress Studio regression suite.
