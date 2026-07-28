# Desktop Mapping Performance — Acceptance

- Progress and BOQ Excel data are streamed with `iter_rows(values_only=True)`.
- Workbooks are closed after data is loaded into the mapping store.
- Tkinter tables do not own the complete dataset.
- Activity table displays at most 200 rows per page.
- BOQ table displays at most 300 rows per page.
- Checkboxes are lightweight `☐/☑` Treeview values, not thousands of widgets.
- Search runs only when the user presses Search or Enter.
- Mapping state is held in `MappingStore` and exported safely to a new workbook.
- Undo restores the previous mapping operation.
