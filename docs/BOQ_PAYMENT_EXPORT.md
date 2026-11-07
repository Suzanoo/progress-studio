# BP-2 BOQ Payment Workbook Export

Status: BP-2 accepted. BP-2.1/BP-3 implementation handoff; their acceptance remains pending.
Baseline: `cc6d1476882dc6899cb503aa085c6d6068c06692` (accepted BP-1).

In Payment, choose **Export BOQ Payment Workbook**. Select a saved Project Session,
current saved Payment Workbook, and optionally a compatible BOQ. Review the source
hashes/revision and diagnostics, acknowledge applicable revision differences,
then export a new `.xlsx` destination. Existing files and input aliases are refused.
Unsaved Mapping UI or Excel edits are not imported or saved implicitly.

The output contains Percentage, Amount, Trace, Payment Input, Payment, and
hidden/protected `_BP_Data` in that order after BP-3. The first two tabs retain the same Project hierarchy. Their native
selectors choose Cumulative or Period views over immutable BP-1 numerical data.
**Select → F9 → Read/Print/Save**. Formula caches may be empty until the spreadsheet
engine calculates. These are projected Requirements, not certified payment.
No VBA, ActiveX or external workbook relationship is created.

Report grids are intentionally unprotected so native row outlines work. Accidental
report edits are recovered by re-export. The hidden snapshot is protected with the
existing workbook password; protection is not cryptographic security. The new
workbook is a derived report, not a BOQ source for BU-1. No automatic import exists.

Group summaries count distinct descendant leaf items with a fixed valid BOQ Amount
denominator. Positive-Amount incomplete/unmapped descendants prevent final group
percentages and monetary totals. Zero-Amount items retain trace/status without
changing a monetary denominator. Unknown, invalid and zero states remain distinct.
BOQ subtotal rows are context only. Duplicate WBS codes, noncontiguous descendants
and descendants preceding parents are rejected as ambiguous hierarchy.

Exact numeric zero is visually blank in report grids. Consecutive exact 100% is
suppressed only in Cumulative %. Underlying values remain numeric; near-zero and
near-100 values are not suppressed based on rounding. Period negatives are retained.
Trace uses unique Item and Period selectors. Navigation only opens Trace; it does
not transfer a selected row. Trace retains every contributor, raw/derived state,
current/previous contribution, signed change, known Amount and source cell. The
source/reconciliation diagnostics follow the trace table. Raw blank is preserved.

Sources are hashed at Review and checked again around export. A changed source
invalidates the reviewed export. Temporary output is validated before atomic,
no-clobber publication. Cancellation during assembly discards work; cancellation
arriving after successful publication cannot retract a completed export. Sources
and existing output files are never overwritten. Unsupported filesystem hard-link
publication fails safely rather than using an unsafe overwrite fallback.

The BP-1 service retains its existing snapshot and compatibility restrictions.
BP-2 does not alter Payment Position, Rebuild, Mapping, BOQ Update, EV or XML.
Actual main palettes, existing recalculation policy and native validation are reused.
Printing uses landscape A3, repeated identifiers/headers and period page breaks;
Windows print-preview usability is an acceptance gate.

Engineer verification:
`python -m pytest -q tests/integration/payment/test_boq_payment_export.py`
followed by full `python -m pytest -q` and real NKC1 export/projection comparison.
Native Windows acceptance: open without repair; grouping/palette readability;
all view selectors + F9 + save/reopen; Trace selections; print-preview readability.


## BP-2.1 Known Projected Amount

Amount row 5 is a frozen/repeated-print monetary summary following B3's view.
It aggregates distinct leaf known Amounts, including partial mapped contributions,
without adding hierarchy/subtotal values or redistributing unknown money.
Period mode sums signed changes with a zero opening baseline; later changes
require known endpoints for the same leaf. No known contributions displays
Unavailable rather than zero. Valid zero displays 0.00 in this summary.
The visible coverage warning is independent of color and remains for incomplete
positive-Amount or invalid-Amount descendants. The existing Project Summary
continues to withhold complete totals when coverage is incomplete.
No Activity input, BP-1 mathematics or Payment eligibility is changed.


## BP-3 Payment informational snapshots

Export requires the selected source workbook to contain both `Payment Input` and
`Payment`. Both are protected read-only saved-value snapshots, never new editable
Payment inputs and never a second Payment engine. Source content stays at its
original coordinates. A snapshot banner and source/hash note replace the A1 title;
other labels, Thai text, styles, periods, numeric zeros, negative values, saved
blank strings, merged ranges, row/column dimensions and outline levels are retained.
Cached source error values are preserved and counted in Review diagnostics.

No cell formulas, data validations or hyperlinks are copied to these snapshots.
Local conditional formatting may use frozen same-sheet helper values exclusively;
this preserves source warning/100% presentation without live business calculations.
Dependent names/foreign-sheet formatting fail closed. Source pictures are retained.
Charts use literal cached series; missing chart caches or unsupported DrawingML
shapes fail closed instead of silently dropping content. No external links,
macros or source table controls are imported. Snapshot protection can restrict
native outline interaction; grouped rows are expanded so all saved Activity detail
remains readable. Hidden helper columns and manually hidden ungrouped rows retain
the source state. Frozen panes use actual split coordinates, not a scrolled
origin. For recognized Activity headers, an oversized frozen body is reduced to
the header so expanded static detail remains readable; initial view returns to
the first unfrozen cell. Source files and cell coordinates are unchanged.

Missing formula caches block review/export with a precise cell and instruction to
recalculate/save the source first. A genuine saved empty string is a valid blank,
not a missing cache or numeric zero. Numeric cache freshness cannot be established
from XLSX alone: Review, Trace and snapshot banners explicitly say it is unverified.
Even if a stored value is stale, export never silently recalculates Payment or
substitutes a derived projection. Recalculate/save the selected source before
export; edits to source after Review invalidate export via existing source hashes.
Snapshot values remain fixed after report dropdown/F9 changes. To update them,
edit/save/Rebuild the source Payment workbook as appropriate, then re-export.

BP-3 is an incremental package on top of the exact BP-2.1 package. Source integrity,
revision acknowledgement, cancellation and no-overwrite publication remain owned
by the existing export service. The two original report worksheets remain
unprotected for grouping; snapshot tabs and internal data are protected.
