# BP-4 — Manual BOQ planned requirements

## Workflow and ownership

Payment workspace: **Prepare BOQ Requirements → BOQ Requirement Input → Save → Export BOQ Payment Workbook**.
Select the saved Project Session, current Payment source workbook and optional compatible BOQ revision. Review diagnostics and acknowledge applicable revisions. Prepare publishes a new source `.xlsx`; it never overwrites an existing file or the inputs. Continue with this new source for subsequent edits and exports.

`BOQ Requirement Input` is the persistent editable source. The separate seven-sheet BOQ Payment report remains a derived snapshot; edits to its report cells are never imported. No Activity Payment Rebuild is necessary to use saved BOQ requirements. Activity Payment Input, eligibility, Payment Position and certified Payment snapshots retain their accepted semantics.

## Input and calculation contract

Only wholly unmapped BOQ leaves are eligible. Each editable cell is a signed **Period % of the full BOQ leaf Amount**, stored as a numeric fraction. Explicit zero is known zero; blank is unknown. Text, formulas, booleans, nonfinite values and values outside -100%..100% are rejected. A known cumulative prefix must stay between 0% and 100%; corrections may be negative. Excel validation helps entry, but export/Prepare validation is authoritative.

A period amount is known whenever that period has a valid numeric entry and the BOQ Amount is valid. Cumulative requirements are known only when **every period from opening through the selected period is entered**. A blank does not carry forward and is not replaced by zero. After a gap, later explicit Period % still contributes to period known totals, while cumulative values remain unavailable until the gap is resolved.

No Mapping Share is invented. Manual origin retains zero Activity Mapping coverage. Mapped BP-1 results are retained unchanged and composed with disjoint manual leaves. Row 5 combines known mapped and manual planned money; its percentage denominator remains the complete Total BOQ Amount. Unknown positive descendants withhold complete project/group results. Explicit period completeness is independent from cumulative history. Zero/invalid Amount handling and exact completion display preserve BP-3.2; an unknown or invalid value cannot create completion.

## Identity, reconciliation and safety

Schema `bp4-signed-period-v1` stores stable identity, catalog/allocation signatures, period identities/dates and source lineage in protected/hidden references. Prepare refreshes references and preserves existing entries by identity. Changed catalog/allocation revisions require Prepare before export. Period identity/date changes, missing/duplicate/changed identities, or previously manual rows receiving Mapping stop with a conflict, including blank existing rows; no automatic deletion or migration is performed. Preserve the old source and resolve explicitly. Partial-mapping residual entry is unsupported.

Prepare grafts only the owned input sheet and necessary workbook/style metadata into the original OOXML package. Other worksheets, caches and drawings remain byte-preserved. Mapping Save As carries the sheet forward; Snapshot/Live Payment Rebuild preserves its values and visibility. Subsequent export still validates source compatibility and Mapping conflicts. Prepare does not recalculate cached Payment snapshots or save unsaved Mapping UI changes.

Input references are locked; numeric period cells are unlocked. Identity columns and schema metadata are hidden, with freeze panes at I9 and repeated print identifiers. Protection is accidental-edit protection, not cryptographic tamper resistance. Existing objects/relationships on the owned input sheet are unsupported and stop reconciliation rather than being discarded. New paths only, source hash rechecks, temporary files and atomic no-clobber publication prevent source overwrite and partial output. Unsupported hard-link publication fails safely.

## Report and trace

The seven worksheets are Percentage, Amount, BOQ Mapping, Activity Mapping, Payment Input, Payment and _BP_Data. The Amount diagnostics appendix includes source/reconciliation diagnostics and mapped/manual subtotals across all periods. Manual requirements never create Activity Mapping relationships. The protected snapshot retains the underlying manual values, source cells and availability states.

Per-period manual coverage retains item counts, valid Amount coverage and invalid Amount counts. Activity reconciliation and unmapped BOQ amounts remain separate diagnostics.

Dropdown → F9 → Read/Print/Save remains the report workflow. Re-export restores report edits. Workbook formulas are presentation-only; the projection snapshot remains immutable and protected. BP-4 does not add a Rebuild target, persistence store, partial allocation engine, Finance, EV or XML behavior.

## Verification

Focused: `python -m pytest -q tests/integration/payment/test_boq_manual_requirements.py`

Full: `python -m pytest -q`

Real proof and engineering evidence accompany the implementation handoff. Native Windows Excel acceptance remains separate from automated and LibreOffice verification. Demonstration inputs in example workbooks are clearly identified as synthetic entries, not customer-approved requirements.
