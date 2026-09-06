# Progress Studio — Earned Value Live Workbook Contract

## Scope

This document is the frozen contract for the live Earned Value workbook.
It supersedes the earlier reporting-cutoff interpretation used during the
EV reporting-context bug-fix iterations.

The design follows the established Live Progress pattern: Python owns
structural derivation at Rebuild; Excel owns live Plan/Actual reporting from
`main` after the workbook has been created.

## Ownership

### Structural / EV-Rebuild owned

A successful EV Rebuild freezes the structural EV topology until the next
EV Rebuild:

- BOQ identity and BOQ amount / BAC authority.
- BOQ → Activity mapping relationship.
- Mapping share and allocated amount.
- Activity identity and WBS binding used by EV.
- Project, Activity and BOQ BAC values.

The existing EV-1 mapping-completeness rule is unchanged. A positive-amount
BOQ that is not fully allocated still hard-stops EV Rebuild.

### Live / Excel owned

After EV has been built, the current `main` worksheet is the live authority
for:

- Activity Plan progress.
- Activity Actual progress.

The following therefore recalculate in Excel after user edits followed by
F9 / Save, without EV Rebuild:

- PV.
- EV.
- SV.
- SPI.
- Project KPI values.
- WBS performance.
- Top negative BOQ variance.
- EV Table values.
- Project PV / EV chart values.

User edits to BOQ Amount, mapping topology, mapping share, or allocated BAC
remain outside the live contract and require EV Rebuild.

## EV View Date

The Earned Value Status Date is presentation state only.

The workbook exposes one semantic defined name:

```text
EV_View_Date
```

In the current layout it is bound to:

```text
'Earned Value'!$M$3
```

Downstream EV formulas must use `EV_View_Date`, not the physical `$M$3`
coordinate. This keeps calculation independent from future dashboard-layout
changes.

The Status Date default is neutral UI state, not reporting logic:

1. EV refresh preserves the existing `EV_View_Date`.
2. First EV creation uses the latest canonical monthly reporting point from
   `Dashboard_Data` when available.
3. A standalone workbook falls back to the latest valid reporting date in `main`.

Rebuild never inspects Actual progress and never reads Dashboard/main cutoff to
choose the EV view. The legacy `cutoff_date` field carried by the EV derivation
result is retained only as the renderer's view-date seed; it is not an EV
calculation authority.

## Live calculation contract

For the selected `EV_View_Date`:

```text
PV = BAC × cumulative Plan through EV_View_Date
EV = BAC × cumulative Actual through EV_View_Date
SV = EV - PV
SPI = EV / PV
```

Plan and Actual both use the same view date.

There is no Earned Value calculation rule of the form:

```text
MIN(EV_View_Date, Dashboard cutoff)
```

and EV formulas must not use `Dashboard!K5` as an Actual boundary.

If later Actual periods are blank / zero, they add no cumulative progress, so
EV remains flat naturally. When the user later enters Actual into those
periods in `main`, F9 / Save immediately updates EV through the selected view
date without rebuilding EV.

## Dashboard independence

The Progress Dashboard cutoff is Dashboard presentation state. It is not an
Earned Value calculation input.

Changing Dashboard reporting controls without changing `EV_View_Date` must
not change Earned Value calculations.

## EV_Data contract

`EV_Data` is a hidden, thin live helper layer over `main` plus static EV
structure. It is not a Python numeric snapshot of mutable reporting state.

```text
main
  Plan / Actual
      |
      v
EV_Data live formula layer
  Activity BAC + live progress
      |
      +--> WBS aggregation
      |
      +--> mapped BOQ aggregation
      |       |
      |       +--> Top Negative ranking
      |       +--> EV Table
      |
      +--> Project PV / EV chart source

BOQ + Mapping --EV Rebuild--> static BAC / topology
EV_View_Date ---------------> all live EV views
```

The live Activity layer is intentionally O(Activity) and follows the same
direct-to-`main` pattern used by the Progress Dashboard Activity table. It
must not create an Activity × time Python snapshot cache.

## Visible-view contract

All visible Earned Value views consume the same live contract:

- Project KPI cards: BAC is structural; PV / EV / SV / SPI are live.
- Project chart: PV and EV are cumulative live curves from `main`; the selected
  Status Date only masks / marks the view.
- Active WBS Performance: aggregates the live Activity layer.
- Top 10 Negative Variance: ranks the live BOQ SV / SPI layer.
- EV Table: BAC is structural; PV / EV / SV / SPI are live.

No visible view may fall back to a Python PV / EV / SV / SPI snapshot after a
successful live EV build.

## View-date calendar

When `Dashboard_Data!K` exists, EV reuses its complete canonical monthly date
calendar as a selectable date list only.

The calendar is not a Dashboard-cutoff dependency. In standalone/unit
workbooks without `Dashboard_Data`, EV builds a full monthly view-date list
from the project reporting points and does not truncate it at the rebuild
cutoff.

## Rebuild boundary

No EV Rebuild is required for:

- Plan progress edits in `main`.
- Actual progress edits in `main`.
- Changing the Earned Value Status Date / `EV_View_Date`.

EV Rebuild is required for:

- BOQ / BAC structural changes.
- Mapping relationship changes.
- Mapping share / allocated amount changes.
- Other structural changes that invalidate embedded EV topology.

## Acceptance contract

The implementation is accepted only when all of the following remain true:

1. Create EV while Actual currently exists only through an earlier period.
2. Enter later Actual directly in `main` after EV creation.
3. F9 / Save updates Project EV, WBS, Top Negative, EV Table and the EV curve
   without EV Rebuild.
4. Editing Plan in `main` likewise updates all dependent PV / SV / SPI views.
5. Changing Dashboard cutoff controls does not change EV for an unchanged
   `EV_View_Date`.
6. EV formulas reference `EV_View_Date` semantically rather than depending on
   the physical `$M$3` coordinate.
7. Mapping incompleteness continues to hard-stop EV Rebuild.

## Out of scope

AC / CV / CPI / EAC / ETC / VAC / TCPI remain outside this contract.
