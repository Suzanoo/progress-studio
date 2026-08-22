# WP-2 — Shared Final Workbook Policy

Status: **Implemented**

Base: `wp-1-audit`

## Goal

Create Progress, Mapping Export, Snapshot Rebuild, Live Rebuild and Payment must finish a workbook through one policy boundary instead of independently remembering visibility, protection, guide and recalculation rules.

WP-2 does not merge calculation/rendering engines and does not change XML normalization, fake amount, S-Curve math, `progress_table` ownership or workspace UI.

## Shared boundary

`progress_studio/infrastructure/excel/final_workbook_policy.py`

```text
calculation/rendering pipeline
          |
          v
finalize_workbook(workbook, mode=...)
          |
          +-- Workbook README guide
          +-- final sheet visibility
          +-- final sheet/cell protection
          +-- mode-specific recalculation
```

### Snapshot mode

Used by:

- Create Progress from XML
- Mapping Export
- Snapshot Rebuild
- Payment preparation / snapshot workflows

Recalculation remains dependency-based automatic/incremental.

### Live mode

Used by Live Rebuild paths.

Recalculation remains manual during editing with calculate-on-save.

## Visibility policy retained

WP-2 intentionally centralizes the proven visibility helper without changing its current states:

- Visible: `README`, `main`, `main_monthly`, `Payment Input`, `Payment`, `Dashboard`
- Hidden: `progress`, `progress_table`, `Dashboard_Data`
- VeryHidden: other support/internal sheets

A future product-policy milestone can change Hidden vs VeryHidden without touching each generator again.

## Protection policy retained

WP-2 reuses the existing internal sheet password configuration and protection helper. There is no password input in Create Progress.

Key behavior remains:

- formula/support cells are locked,
- intended `main` Activity inputs remain unlocked,
- `main_monthly` local cutoff remains unlocked,
- Dashboard View / Cutoff / Status controls remain unlocked,
- Payment Input activity percentages remain unlocked,
- workbook structure remains unlocked.

## Create Progress change

The final XML-generated workbook now receives the same final workbook policy as Rebuild/Mapping/Payment. This closes the main WP-1 drift: a workbook no longer needs to pass through Rebuild before it receives guide, visibility and protection behavior.

## Deferred

- Create/Mapping workspace toolbar cleanup
- `progress_table` snapshot decision
- overlay ownership
- S-Curve calculation unification
- changing support sheets from VeryHidden to Hidden
