# FF-3 Financial Forecast desktop integration

Implementation baseline: `main @ 997a815ca51aff08a2f1ab6fd4204e2a9d4fc8c2`.
Scope: FF-3 desktop integration and regression only.

The Product Owner has accepted FF-0, FF-1 and FF-2, including FF-2 first-open
integrity, live Excel recalculation and Save/Close/Reopen persistence. This
milestone does not reopen that investigation. The calculation source of truth
remains [FINANCIAL_FORECAST_CONTRACT.md](FINANCIAL_FORECAST_CONTRACT.md); the
accepted workbook/input contract remains [FF2_WORKBOOK.md](FF2_WORKBOOK.md).

## Normal application workflow

1. Create a Progress workbook in the application, or use an existing saved
   Progress Studio `.xlsx`. Mapping and Payment remain optional independent
   workflows. Save and close Excel before using the workbook in the app.
2. Open **Finance** in the sidebar, **Tools → Financial Forecast Workspace**, or
   **Open Financial Forecast** on Home.
3. **Browse...** selects and checks a workbook. If typing/pasting the path, click
   **Check Workbook**. A changed path invalidates the previous check.
4. First creation: enter **Opening date**, **Actuals through** (YYYY-MM-DD), and
   **Project currency**. These are finance dates; they are not inferred from
   schedule dates or EV cutoff. Prepared rows default to 100 per input table.
5. Click **Create Finance Workbook** and choose a new `.xlsx` path. The source
   and existing destinations cannot be overwritten. Cancelling creates nothing.
6. **Open Result** opens the new workbook in the OS-associated spreadsheet app.
   Fill in reconciled opening cash, actual cash, remaining obligations and any
   overrides in **Finance Input**. New workbooks start with zero opening cash,
   empty records and incomplete declarations. Defaults remain 30 calendar days
   and 5% retention. Review **Cash Flow** after F9 / Save.
7. Save/close Excel. Return to Finance and check the saved workbook again. The
   successful output is already selected in the workspace. An existing finance
   workbook offers **Refresh Finance Workbook**; dates/currency are displayed
   from the workbook and disabled in the app. Edit these settings in Excel.
8. Refresh uses saved Finance Input, including notes and explicit zero overrides.
   Increase **Prepared rows per table** when needed (maximum 2000); shrinking is
   rejected. Refresh always saves another new file. For ordinary value changes,
   F9 / Save is sufficient; Python refresh is not required for recalculation.

The result does not silently change the selected source in another workspace.
For Mapping, Payment or Rebuild, explicitly select the latest saved workbook.
Finance has no Snapshot/Live selector: its Cash Flow remains the FF-2 live view.
Progress/Payment Rebuild scopes and ownership are unchanged.

## Implementation and boundaries

- The existing desktop shell hosts `FinanceFrame`, using the established ttk
  card styles, scrollable body, file dialogs, worker-thread/queue pattern and
  OS workbook opener pattern. Controls are disabled during work; exceptions
  return to the UI thread and leave Check Workbook available for recovery.
  Application close waits for a running finance operation to finish.
- `FinancialForecastWorkbookService.analyze` provides read-only readiness:
  `.xlsx`, existing `main`, supported date system and accepted finance schema /
  input validation when present. It does not infer cash from `main`, require
  BOQ mapping, or treat incomplete finance as invalid. The worker rechecks the
  selected saved file immediately before generation.
- The same FF-2 `generate` method handles desktop and CLI writes. Its atomic
  temporary-file lifecycle, validation, exact calculation-property preservation
  and finance package merge are unchanged. No new writer or calculation engine
  is introduced. Coordinates remain in existing Excel adapters/renderers.
- The workbook layout, formulas, protection, visibility, Mapping preservation
  and FF-1 semantics are unchanged. Financing remains separate from operating
  Cash Out; dated events determine peak funding and monthly values only aggregate.
- `.xlsm` remains unsupported by Finance V1. Close Excel before refresh; the
  application reads saved disk contents, not unsaved changes in an Excel window.
  Fix invalid finance inputs in Excel before checking again. Existing adapter
  limitations, including unsupported finance-sheet comments/hyperlinks, remain.

## Product Owner desktop acceptance — separate FF-3 gate

Run using the normal desktop launcher with the FF-3 files applied to baseline.
Record OS, app baseline, Excel version, display scaling and each result. Use
copies of project files and unique output names. Automated headless tests do
not establish native window layout, OS file association or Desktop Excel behavior.

### A. Entry points and first creation

1. Launch `python desktop.py` (or the normal installed desktop entry point built
   from these files). Verify **Finance** in the sidebar, the Tools menu entry
   and Home button all open the same Financial Forecast workspace. Mapping's
   command bar must not appear there. At minimum window size / normal display
   scaling, use the vertical scrollbar to reach every control.
2. Create a Progress workbook through **Create Progress Bar** with a normal
   supported schedule XML. Save it as `project.xlsx`. An existing accepted
   project workbook can be used as a second run. Close it in Excel.
3. In Finance, Browse to `project.xlsx`. First-creation fields must enable;
   no dates/currency are inferred. Set Opening date `2026-01-01`, Actuals through
   `2026-01-31`, currency `THB`, prepared rows `4` for the synthetic check below.
4. Cancel the Save dialog once: no result should be created. Try selecting the
   source or an existing output: the app must refuse overwrite. Then create
   `project_finance.xlsx` at a new path. During work, operation controls must be
   disabled and closing the application must ask you to wait.
5. Click **Open Result**. Verify no repair dialog; Finance Input/Cash Flow visible,
   Finance_Data normally hidden; input cells editable. Record first-open values
   before F9 separately. Empty finance must be conditional/incomplete, not a
   declaration that the project needs no funding. Verify other existing sheets,
   charts and editable controls against `project.xlsx`.

### B. Enter a synthetic scenario through the generated workbook

This small dataset verifies the actual app-produced workbook. It is not a real
project forecast. Use date-valued Excel entries, and leave unspecified cells blank.

Set Finance Input **B6 = 10** (opening cash) and **B10:B13 = TRUE** for this fully
specified synthetic case. Keep credit days 30 and retention 5%.

In the forecast table, enter:

| Row | J: Item ID | K: Amount | L: Direction | M: Expected / certification date | N: Category | Q: Kind | R: Retention release |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 25 | pay | 100 | out | 2026-02-05 | operating | manual | blank |
| 26 | invoice | 120 | in | 2026-02-25 | operating | manual | blank |
| 27 | cert | 100 | in | 2026-01-31 | operating | certificate | 2026-06-01 |

In the actual cash table enter row 25: **A = receipt**, **B = 2026-01-20**,
**C = 40**, **D = in**, **E = operating**, **F = cert:net**, **H = Paid cash**.

1. Press F9. Expected: peak funding **50 on 5-Feb-2026**, minimum cash **-50**,
   remaining receivable **180**, remaining cash out **100**, retention **5**.
   Monthly closing cash January–June: **50, 70, 125, 125, 125, 130**.
2. At the last prepared override row enter **U28 = cert:net**, **V28 = 0**,
   **Y28 = Cancel remaining**, **Z28 = Keep this note**. F9: remaining receivable
   becomes **125**, final cash **75**, actual receipt stays **40**. Clear V28:
   receivable returns to **180**, final cash **130**. Restore V28 to zero.
3. Save/close/reopen. Confirm zero, reason, note, formulas and expected figures
   persist. Close Excel again.

### C. Refresh and normal workbook lifecycle

1. Back in Finance, Check Workbook on `project_finance.xlsx`. Expect Refresh
   Finance Workbook; workbook dates/currency displayed but disabled. Set prepared
   rows to `3`: shrinking must be refused. Set `6` and refresh to a new file
   `project_finance_refreshed.xlsx`.
2. Open Result. Confirm all input values, zero override, reason and note remain;
   new prepared rows are editable; F9 gives the same figures. No repair dialog.
   Save/close/reopen and check again.
3. On copies of this refreshed file, use normal application **Mapping Export**,
   **Payment → Prepare Payment Input**, **Payment → Build Payment Breakdown**,
   Rebuild **Snapshot + Progress**, **Live + Progress**, **Snapshot + Payment**,
   **Live + Payment**. Use valid BOQ/mapping and Payment requirements for each
   workflow. Where the real project is EV-ready, also refresh Earned Value.
   Each workflow must keep finance inputs, notes, zero override, visibility,
   unlocked inputs and live calculations, while its own existing views/charts
   retain their accepted behavior. Do not change unrelated owner inputs merely
   to make an unready workflow pass.
4. Browse one resulting saved workbook in Finance, Check Workbook and refresh
   to another new file. Repeat Open Result / F9 / Save / Close / Reopen. This
   closes the normal app → Excel → Mapping/Rebuild → Finance return path.

### D. Recovery

1. Type a different path after a successful check: generation must disable until
   Check Workbook runs. A missing/corrupt workbook, `.xlsm`, workbook without
   `main`, or invalid finance schema must report an error without creating output.
2. On a copy, enter an invalid cash settlement link and save/close. Check Workbook
   must report validation failure. Correct it in Excel, save/close, then check and
   refresh successfully. Incomplete declarations or undated positive obligations
   may remain conditional; they must not be silently assigned dates.
3. Check a file, then modify it in Excel before generation. Generation must read
   the latest saved data; a changed finance-sheet presence or larger capacity
   must require another check instead of silently changing the operation.

FF-3 is not Product Owner accepted until these application/Desktop Excel checks
are signed off. FF-2 acceptance is retained. No later milestone is included.
