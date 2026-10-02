# EV monetary sources

The PO-authorized EV monetary source extension supersedes the BOQ-only input,
project-progress shortcut and zero-SPI rules of the original EV Live contract.

In Rebuild, explicitly select **Activity Amount** or **BOQ Mapping** before
Generate / Refresh EV. There is no fallback. The selected source and allocated
Contract Value basis appear in Earned Value and EV Table.

## Persistent inputs

New Create workbooks contain **EV Monetary Inputs**. Its Allocated Contract
Value column is the current monetary authority for Activity Amount EV. It is
independent of main's Progress Weight/Amount. Amount Create seeds ordinary
values from the selected field and retains the milestone's original monetary
value separately from its zero Progress Weight. Equal/Duration create blank
monetary inputs; enter real allocated Contract Values before building EV.

For existing workbooks, **Prepare monetary inputs** creates a new output file.
For Amount-origin workbooks it copies current main Amount, not original XML or
historic metadata amounts. For dummy/unknown origins values remain blank.
Verify every value, explicitly mark milestones and enter their dates. Old MS-2
workbooks discarded milestone values: enter those values explicitly; they
cannot be recovered from a zero progress weight. No XML reread occurs.

Do not enter cost, cash flow or dummy units as allocated Contract Value. Edit
Contract Value here and explicitly refresh EV. The previous EV remains based
on its last successful refresh until then. Changing source also requires refresh.

Missing, nonnumeric, negative, non-finite, unsupported values, duplicate IDs
and a nonpositive project BAC block readiness. Individual zero is valid.
Excel float precision applies. Values below Excel's normal numeric range or
above its supported maximum are rejected. Formula-valued BAC inputs are not
accepted; store explicit numeric Contract Values.

## Milestones

Milestone flag is structural; refresh after changing it. Progress Weight must
be zero. Planned Milestone Date and Actual Completion Date are editable live
Excel dates. Blank Actual Completion Date means not completed. Dates are
compared by calendar day: before the date earns zero; on/after earns 100%.
The completion date is the explicit record of completion, not a payment signal.
Changing these dates recalculates EV historically without EV refresh.

Mapping regeneration preserves these independent inputs by Activity ID and
keeps known milestone Progress Weight zero. Newly introduced activities require
explicit monetary inputs. Both Progress rebuild modes preserve the input sheet.

## Calculation and BOQ integrity

Both modes normalize Activity BAC and use one calculation/renderer. Project
and WBS PV/EV sum underlying Activities exactly once. Project progress percent
is never multiplied by Project BAC. SPI is blank when PV is zero.

BOQ mode retains completeness/identity/share checks. Ordinary main amounts
must reconcile to allocated BOQ amounts, and each persisted Allocated Amount
must reconcile to BOQ Amount × Share. Tolerance is 1e-7 absolute / 1e-10 relative.
Milestone BAC uses BOQ allocation while its progress weight remains zero.
Independent Activity Amount totals need not equal BOQ totals.

Activity mode does not read BOQ sheets. EV Table and negative variance rank
Activities; BOQ mode retains BOQ detail. Refresh replaces all three EV-owned
sheets coherently. EV Status Date remains independent of Dashboard cutoff.
Ordinary Plan/Actual remain live from main. BAC is frozen at refresh.

Excel F9, first-open integrity and Save/Close/Reopen require Desktop acceptance;
automated formula/package tests do not establish those gates.
