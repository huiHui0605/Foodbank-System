# Community Admin Desktop — Dynamic Excel Edition

## Install (Windows)
1. Install Python 3.12 or 3.13 and Microsoft PowerPoint (for PDF certificates).
2. In this folder run `py -m pip install -r requirements.txt`.
3. Run `py main.py`.
4. Click **Import Excel**, choose any `.xlsx` or `.xlsm` workbook, and select its worksheet.
5. All columns and nonempty/empty rows are displayed regardless of their headings. Edit cells, add rows, delete rows, then click **Save Changes** to write back to the same original workbook.
6. If you change the Excel file outside the app, the app reloads automatically when there are no unsaved edits.
7. Use **Certificate Settings** to identify Name, Type, Phone, and Date columns. Place your real PPTX files at `templates/donor_template.pptx` and `templates/volunteer_template.pptx`. Templates should contain `{{NAME}}` (and optionally `{{DATE}}`, `{{TYPE}}`). Click **Generate PDFs** to create PDFs using Microsoft PowerPoint.
8. The **Share** button opens WhatsApp with a prepared message. Attach the PDF manually.

## Important limitations
- The dynamic worksheet view supports any ordinary tabular Excel layout with the first row treated as column headings. It does not reproduce Excel charts, merged-cell layouts, formulas' calculated results, or macros in the desktop table.
- Editing a formula cell will replace that formula with text. Saving can affect formula references and Excel tables when rows are deleted. **Back up the workbook before editing.**
- The original workbook is updated only when you click Save Changes; external file changes reload automatically if there are no unsaved edits.
- Google Sheets synchronization for arbitrary column structures is **not implemented in this edition**; the button explains this rather than risking data loss. The earlier structured version supports Google Sheets sync.
- PDFs are generated when you click Generate PDFs (not automatically on every edit). This avoids generating certificates from incomplete records.
- PowerPoint must be installed on Windows for PPTX-to-PDF conversion.
- The included `records.xlsx` uploaded by the user has headers but no data records.


## Display contrast fix
This edition explicitly sets dark text on light backgrounds for the main window, labels, search box, worksheet selector, table cells, alternating rows and selection states. It also sets white text on dark table headers and blue buttons. This avoids white-on-white text under Windows dark mode or Qt palette inheritance. Restart the application after replacing the previous files.
