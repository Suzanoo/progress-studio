# 009 Three-File Chart Regression

Evidence workbooks supplied on 2026-08-16:

- `009-MSP_final_progress(1).xlsx` — Create Progress output
- `009-MSP_test_payment.xlsx` — Payment output
- `009-MSP_test_progress_live.xlsx` — Live Progress rebuild output

## Root causes

1. **Create Dashboard Monthly curve squeezed left**
   - Dashboard chart used a text/category axis over a helper range sized for the larger Weekly view.
   - When Monthly was selected, trailing blank category formulas still occupied category slots.
   - Fix: use a real Excel DateAxis so actual dates own horizontal geometry.

2. **Monthly traditional overlay lost the final 100% reporting point**
   - Overlay bounds were recomputed from the calendar month of raw Project Finish.
   - The final reporting period can legally fall in the following month when it overlaps Project Finish.
   - Fix: Dashboard_Data nonblank Plan rows are authoritative for overlay source bounds; physical month columns are mapped from those reporting dates.

3. **Live Dashboard Plan fell from 100% to trailing zeroes**
   - Live selected Plan returned blank source cells directly. Excel evaluates referenced blank cells as zero inside the selector formula.
   - Display-margin dates therefore became plotted zero points.
   - Fix: blank Weekly/Monthly Plan sources explicitly return `#N/A`.

4. **Live Dashboard presentation drifted from normal Dashboard**
   - Live renderer had its own default axes/grid styling and 100% hard ceiling.
   - Fix: align DateAxis, 110% headroom, grid and axis styling with the normal Dashboard visual contract.

## Ownership preserved

- Normalize/data engines unchanged.
- Traditional overlay renderer owns only overlay source geometry and chart presentation.
- Dashboard renderer owns Dashboard chart geometry/presentation.
- Payment-only rebuild continues to preserve Progress-owned overlays.
- Rebuild 2×2 matrix remains the cross-workflow safety gate.
