# Manual testing

Launch the Windows app and choose **Try sample project**. The sample contains three fictional employees; no confidential files are required.

## Two-way values and formulas

1. Double-click cell D2 and enter 650. The increment field should show 650.00 and final salary should become 4,850.00.
2. Edit the green employee-name field directly in the document. The value in A2 should match immediately.
3. The blue final-salary field comes from a formula and is read-only in the document. Change its spreadsheet formula or input cells instead.
4. Switch to the next record. Jordan's values should appear; Avery's edits should remain in row 2.

## Visual mapping and placement

1. Choose **Single document**.
2. Select a spreadsheet cell, place the cursor in the document, and choose **Link cell → document**.
3. Select existing document text and connect it to an empty spreadsheet cell. Choose **Use the selected document text**; both sides should match.
4. Cut and paste a mapped field to another paragraph. Its source link should remain.
5. Select a range using Shift-click and choose **Insert linked table**.
6. Open **Mappings**, select a field, choose a different spreadsheet cell, then **Use the selected spreadsheet cell**.
7. Open **Header / footer** and edit their text. The page preview should update.

## Batch documents

1. Switch to **Batch documents** and choose the source worksheet and header row.
2. Use the previous/next buttons to inspect all three employees.
3. Individual cell fields follow the current record row; multi-row linked tables stay fixed.
4. Export record rows **2-4** as one combined DOCX or PDF.
5. Export the same rows as **Separate files in a ZIP**.
6. Export rows **2,4** and verify the middle employee is omitted.

## Layout, printing, and exports

1. Open **Page preview**. Add enough paragraphs or table rows to make several pages.
2. Test A4/Letter, portrait/landscape, margins, and an explicit page break.
3. Click a preview page, then export **Current preview page**.
4. Test a custom range such as **1-2,4** and an open range such as **2-**.
5. Export DOCX, PDF, PNG, JPEG, HTML, and TXT. Multi-page image output should be a ZIP with an image for each page.
6. Export the workbook as XLSX and ODS, or the current worksheet as CSV/TSV. Reopen the results and verify values.
7. Print to an available printer or an operating-system PDF printer. Check page selection, paper size, and margins.

## File imports and project reuse

Use the fictional files in **samples/** to try DOCX, ODT, RTF, HTML, TXT, XLSX, ODS, and CSV import. Files can be imported independently; the other side can remain blank.

1. Click the project name in the top bar and rename it to **Payroll review**. Edit a cell, switch records, or change page settings; the name should remain editable and retain the rename.
2. Click **Save project** (or Ctrl+S). No file dialog or download should appear, and the unsaved-change dot should disappear.
3. Make another edit without saving, then close/relaunch the app (or reload the browser). The last saved checkpoint, including its name, document structure, mappings, spreadsheet values, header/footer, and layout, should return. The unsaved edit should not return.
4. Start a new project, give it a different name, and save it. **Open project** should list both checkpoints. Open the first one, rename it, and save; it should update the same checkpoint without creating a duplicate.
5. Open **Export / Print**. Change **File name** to **Salary notices**, then export a DOCX/PDF or workbook. The suggested output filename should match, with the correct extension. Separate batch files and multiple-page images should use **.zip**.
6. Select **Project** in Export / Print to create a **.paperlink** file. Import it using **Open project → Import .paperlink file**. Click Save project to retain it as a browser checkpoint.

A checkpoint includes the spreadsheet and document data, not just the mapping definitions. Browser checkpoints stay on this device; clearing the browser/app profile data removes them. Use Export / Print for a transferable backup.

Word can reflow editable DOCX exports differently from the app's page preview. Inspect imported corporate templates and their exported DOCX files before issuing them. The PDF and image exports use the app's paginated layout.
