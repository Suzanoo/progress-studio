# Quick Start

## 1. Install and run

Open PowerShell in the project folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python desktop.py
```

The Progress Studio desktop window opens maximized.

## 2. Create a Progress workbook

1. In **Project input**, click **Browse...** beside **Schedule XML**.
2. Select the schedule XML file.
3. Choose the **Weekly cutoff day**.
4. Leave **Fallback amount / activity** unchanged unless the schedule contains no amount data and you intentionally need a placeholder amount.
5. Choose **Plan distribution**:
   - `Auto` — use the program's distribution rules.
   - `Flat` — equal distribution.
   - `Front` — more progress near the start.
   - `Back` — more progress near the finish.
   - `Bell` — more progress in the middle.
6. Click **Create Progress Workbook**.
7. Wait until the status shows completion.
8. Click **Open output workbook** or **Open output folder**.

The source XML is never modified.

## 3. Load the mapping workbooks

Open the **Amount Mapping** tab.

1. Click **Load Progress...** and select the generated Progress workbook.
2. Click **Load BOQ...** and select the BOQ workbook.
3. Select the correct **BOQ worksheet**.
4. Click **Load selected sheet**.

When both files are loaded, activities appear on the left and BOQ items appear on the right.

## 4. Map BOQ items

1. Tick one Activity on the left.
2. Tick one or more BOQ items on the right.
3. Enter **Share = 100** for full allocation, or a lower percentage for partial allocation.
4. Click **Map**.
5. Check the top summary:

```text
Mapped amount / Total amount | Remaining amount | Mapped items / Total items
```

Use **Undo** to reverse the latest action. Use **Unmap** to remove selected allocations.

## 5. Save and export

1. Click **Save Session...** and save the `.json` file.
2. Continue mapping until complete, or export a partial mapping when needed.
3. Click **Export...**.
4. Review the allocation summary.
5. Choose the output `.xlsx` filename.
6. Open the exported workbook in Microsoft Excel.
7. Wait for formula recalculation, then save the workbook.

Do not upload or distribute the workbook before Excel has finished recalculating and saving it.
