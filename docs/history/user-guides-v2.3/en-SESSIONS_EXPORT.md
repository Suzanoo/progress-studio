# Sessions and Export

## Save a mapping session

A session is a `.json` file containing mapping allocations and fingerprints of the source workbooks.

1. Load the Progress workbook and BOQ worksheet.
2. Create at least one mapping if needed.
3. Click **Save Session...**.
4. Save the file near the project workbooks.

After the first manual save, later mapping changes are auto-saved to the same session file.

## Resume work

Use either:

- **Load Session...** to browse for a session file.
- **Recent...** to choose a recently used session.

Progress Studio verifies that the Progress and BOQ workbooks are identical to the files used when the session was saved.

If a workbook was moved or renamed, choose **Browse** when asked to relink it. Only a matching workbook is accepted.

## Export the mapped workbook

1. Click **Export...**.
2. Review:
   - Allocated amount and percentage
   - Full BOQ item count
   - Partial BOQ item count
   - Unmapped BOQ item count
3. If mapping is incomplete, choose whether to export a partial workbook.
4. Select the output filename.

The export process:

- copies the Progress workbook safely;
- updates Activity amounts in the required `main` worksheet;
- writes mapping records, BOQ ID, share percentage, and allocated amount;
- preserves formulas and Excel tables;
- requests full recalculation when the workbook opens.

## Final Excel step

After export:

1. Open the output workbook in Microsoft Excel.
2. Wait until calculation finishes.
3. Review `main`, `progress`, and `progress_table`.
4. Save the workbook.

The original Progress workbook is not overwritten unless the user explicitly chooses the same filename and confirms replacement.
