# Progress Studio — Earned Value Live Workbook Contract

## Scope

This document freezes the live-workbook contract agreed for the EV reporting-context fix and EV-L3 refactor.

EV rebuild remains the structural calculation boundary. The workbook then owns live reporting behavior for user edits to Plan and Actual progress.

## Ownership

### Structural / rebuild-owned

The following are frozen by a successful EV rebuild and remain unchanged until EV is rebuilt again:

- BOQ identity and BOQ amount / BAC authority.
- BOQ → Activity mapping relationship.
- Mapping share and allocated amount.
- Activity identity and WBS binding used by EV.
- Project, Activity and BOQ BAC values.

A mapping validation failure continues to hard-stop EV rebuild. The existing EV-1 allocation completeness contract is unchanged.

### Live / Excel-owned

The following are read from the current `main` worksheet and must respond to F9 / Save without an EV rebuild:

- Activity Plan progress.
- Activity Actual progress.

Derived EV reporting values are therefore live:

- PV.
- EV.
- SV.
- SPI.
- WBS performance.
- Top negative BOQ variance.
- EV Table values.
- Project EV chart values.

User edits to BOQ Amount or mapping topology are outside the live contract and require EV rebuild.

## Status Date and reporting cutoff

`Earned Value!M3` is dashboard view state only. It never becomes the project reporting cutoff.

The authoritative reporting cutoff remains the existing Progress Studio reporting context, normally `Dashboard!K5` with the established fallback used by the rebuild service.

For any selected EV Status Date:

```text
Plan Date             = EV!M3
Effective Actual Date = MIN(EV!M3, Reporting Cutoff)
```

Therefore, when Actual is available only through April and the user selects August:

- PV is evaluated through August.
- EV uses the latest permitted Actual through April.
- The EV graph carries that April EV value forward flat through August.
- WBS, Top Negative and EV Table use the same August view / April Actual boundary.

Selecting a historical Status Date before the reporting cutoff evaluates both Plan and Actual at that historical date.

## EV-L3 data-flow contract

EV-L3 follows the established Live Progress pattern: Python builds a compact structural/formula contract; Excel recalculates values from `main`.

```text
main
  Plan / Actual
      |
      v
EV_Data live formula layer
  Activity progress + BAC
      |
      +--> WBS live aggregation
      |
      +--> mapped BOQ live aggregation
      |       |
      |       +--> Top Negative live ranking
      |       +--> EV Table
      |
      +--> Project PV / EV chart source

Earned Value!M3 --------> selected view date
Dashboard!K5 -----------> Actual reporting boundary
```

EV-L3 must not store Python-computed PV / EV / SV / SPI snapshots that become stale after the user edits Plan or Actual in `main`.

The helper layer may retain static BAC, identity and mapping topology because those remain rebuild-owned.

## Rebuild boundary

No EV rebuild is required for:

- Plan progress edits in `main`.
- Actual progress edits in `main`.
- Changing `Earned Value!M3`.

EV rebuild is required for:

- BOQ / BAC structural changes.
- Mapping relationship changes.
- Mapping share / allocated amount changes.
- Other structural changes that invalidate embedded EV topology.

## Compatibility

- The existing full canonical monthly Status Date list is reused from `Dashboard_Data!K` when available.
- EV rebuild still hard-stops for incomplete mapping exactly as EV-1 requires.
- AC / CV / CPI / EAC / ETC / VAC / TCPI remain outside this scope.

## EV-L4 visible-view contract

All visible Earned Value views consume the same EV-L3 live formula layer. No
visible view may fall back to a Python PV / EV / SV / SPI snapshot after a
successful rebuild.

The following views are bound to `Earned Value!M3` and recalculate in Excel:

- Project KPI cards: BAC remains rebuild-owned; PV / EV / SV / SPI are live.
- Management curve: PV follows the selected view date; EV uses Actual through
  `MIN(M3, Reporting Cutoff)` and carries the last permitted EV value forward
  flat through the selected view date.
- Active WBS Performance: BAC is structural; PV / EV / SV / SPI aggregate the
  one-row-per-Activity live helper layer.
- Top 10 Negative Variance: rank, BOQ label, Activity ID, SV and SPI are taken
  from the selected-date live BOQ layer.
- EV Table: one row per BOQ; BAC is structural while PV / EV / SV / SPI are
  formula-driven from the same selected-date live BOQ layer.

`M3` is explicitly a **view selector**, not a reporting-cutoff control. Changing
it must not alter `Dashboard!K5`, mapping topology, BAC, or any other structural
state.
