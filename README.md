# PaperLink

Prepare letters and other documents using information from your spreadsheet. Open your files, connect the information once, then check and save the finished documents.

## Open the app

The latest build is in **release/v0.1.27/**:

- **PaperLink-portable.exe** runs without installing.
- **PaperLink-Setup.exe** installs the app.
- **win-unpacked/PaperLink.exe** is the unpacked build.

Microsoft Office is not required to use PaperLink. File conversion runs on this device. Sending email is optional: when you choose Send and confirm, the message and its document go to your configured SMTP server.

## Get started

Follow the three steps shown across the top:

1. **Open files** — choose your spreadsheet and document, use a saved template, or start blank.
2. **Add information** — click a spreadsheet cell, click where it belongs in the document, then choose **Add to document**. Green items stay connected and update on both sides.
3. **Check and save** — choose **Check pages**, check the information, then choose **Save file / Print**.

**One document** prepares a single document. **One per row** makes a document for each filled row, such as each employee. The arrows show the current person or item and the spreadsheet row being used. **Choose rows** contains the sheet and headings-row settings.

**More tools** contains less-used writing tools. **Connected information** shows where each value comes from; its connection settings are under **Change the connection**. Drag the middle divider to give either view more room. Double-click it to reset.

The workspace uses tighter surrounding spacing to leave more room for both files. Choose **Expand workspace** to hide the top actions and guidance, keeping the spreadsheet, document, editing tools, and row controls visible. **Restore layout** or **Escape** brings the top actions back. This changes the on-screen view; document page settings and text formatting are preserved.

Read [the quick-start guide](docs/QUICK-START.md), choose **Help → How to use PaperLink**, or use **Try an example** to practise with fictional payroll data.

The top actions use consistent outline icons with visible text labels. Hover over a button for a description and available desktop shortcut. Buttons have accessible descriptions, keyboard focus outlines, and large click areas; template actions keep their labels because their symbols alone can be ambiguous.

When dragging spreadsheet cells or moving a connected value in the document, a small label stays away from the pointer so the target text remains visible. A green line shows where the item will be inserted. The insertion line follows the scaled document view, and the ordinary typing caret is hidden during the drag. You can also click a cell, click the insertion point, and choose **Add to document**.

## What the save choices mean

| Choice | What it does |
| --- | --- |
| **Save work** | Keeps your spreadsheet, document, and connections in this app so you can continue later. |
| **Saved work** | Opens work you kept earlier. The last saved work returns when the app starts. |
| **Templates** | Remembers a document and its connections for use with new spreadsheet values. |
| **Save template** | Creates a new template from the document in the right pane, using a unique name. |
| **Save file / Print** | Creates a Word file, PDF, spreadsheet, picture, work backup, or template file; also opens printing settings. |

Saving work or a template inside the app does not create an output file. For another device or a backup, use **Save file / Print → Backup of my work** or **Template to share or reuse**. Saved work and templates inside the app are removed if the app's stored data is cleared.

In **Saved work**, each saved copy has **Open** and **Delete** actions. Deleting requires confirmation and keeps the currently open work, templates, and exported files. If the deleted copy is currently open, it is marked as needing saving; **Save work** creates a new saved copy. Deleting also clears its startup restore selection.

On Windows, saved work and templates are held in an IndexedDB database inside the app's data folder, normally **%APPDATA%\PaperLink**, including when using the portable app. When upgrading from Flowdesk, PaperLink reuses the existing **%APPDATA%\Flowdesk** profile if there is no newer PaperLink profile, so saved work and templates remain available. Choose **Help → Open saved data folder** to find the running app's actual data folder. In a normal browser, the data is held in that browser's site storage instead.

## Reuse a document

Open **Templates → Preview document** to recognise the saved layout. From that screen, choose **Open your spreadsheet**, check the sheet choices, then choose **Use this template**. The document returns with values from the new spreadsheet, without opening the Word file again or connecting every value again.

The document preview fills connected locations with current values as soon as you choose a spreadsheet and sheet. Hover over a connected value to see its cell location. Without a chosen spreadsheet, locations such as **[Employees!A2]** appear instead of old values. A read-only spreadsheet preview beside it shows the actual headings and values in the chosen sheet, with controls to browse more rows and columns. For **One per row**, the document preview uses the first filled row below the headings. The previews do not change the current document; text typed directly into the document stays as written.

Templates keep the document, connection positions, page settings, and row settings. They exclude old spreadsheet records, formulas, and stored copies of connected values. Sheet names can be reassigned, and the app checks that the saved headings remain in the expected columns. **Use with blank spreadsheet** creates empty sheets with the saved headings.

Use **Save template** on the main screen to create a new template. The dialog identifies the document and work being saved. Duplicate names are rejected, including differences only in capitalization or extra spaces.

Use **Templates → Edit** to change a saved template's name, document, layout, or connections. Editing uses a separate copy of your currently open spreadsheet, retaining its values and formulas so you can try changes and see the results in the document and **Check pages**. You can also open another spreadsheet for reference. **Save changes** and **Cancel editing** return you to your previous work. Preview records and formulas are not saved in the template. See [the template guide](docs/TEMPLATES.md).

Choose **Delete** beside a saved template to remove it after confirming. Your current document, saved work, and exported files are kept.

## Email documents

Choose **Email documents** beside **Save file / Print**. Enter the sender address and write a subject and message. Enter or edit **To** and **CC** directly in **3. Review each email**; each email keeps its own addresses. Optional **Fill addresses from spreadsheet columns** reads addresses from each row automatically. Choose PDF or Word attachments and their names, then review each email. **One per row** prepares a separate email with only that row's document attached.

Inside the email window, set **How many emails? → One email per spreadsheet row** for bulk sending. Choose the sheet, headings row, and included rows. This sets the workspace to **One per row** too, so document connections follow each person's row. The **Recipients and their attachments** list pairs each row's addresses with its document. Choose **Review** to edit that person's To and CC, then **Preview this person's document** to inspect the content. **Reset addresses** restores that email's source addresses. Address corrections reuse already prepared documents and update the draft; invalid addresses block sending the affected email and the complete batch. **Save work** keeps individual address edits, which are cleared when a new spreadsheet is imported.

Under **2. Attach the document → Password protection**, choose **No password**, **Use one password for all attachments**, or **Use each person's password from a spreadsheet column**. PDF and Word attachments require the password to open, including the copies in drafts, the draft ZIP, and Save attachment. A typed password is cleared on closing the email window and excluded from saved work; spreadsheet password cells remain part of the workbook and its saved copies. Share passwords separately with recipients. Changing protection settings requires preparing the attachments again. Protection failure stops preparation. The in-app preview remains readable for review.

Under **Email server settings**, enter the outgoing SMTP settings supplied by IT. **Test connection** checks connectivity and login without sending. **Send test email** sends a plain message to a chosen address with no payroll data, attachments, or CC, after confirmation. Passwords stay in the desktop process for the session; **Remember password securely on this computer** optionally stores them with OS encryption. They are excluded from saved work, work backups, and templates.

Choose **Prepare emails**, then **Send this email** or **Send all emails** and confirm the final recipient and attachment summary. Results distinguish server acceptance, rejection, partial acceptance, and uncertain delivery. Sending stops on a problem; **Stop remaining** stops after the current email finishes. Recent send results are stored separately in the app data folder, and previously accepted or unconfirmed copies are skipped unless you explicitly select resend. Server acceptance does not confirm inbox delivery. PaperLink does not write a separate Sent-folder copy; the server may do so.

Draft options remain available: **Open this draft**, **Save this draft**, and **Save drafts (.zip)**. The ZIP includes .eml files, separate attachments, and matching text details for webmail. Draft editing support varies by email app. See [the email guide](docs/EMAIL-DRAFTS.md) for a first sending trial and SMTP compatibility details.

## File formats

| Content | Open | Save as |
| --- | --- | --- |
| Spreadsheet | XLSX, XLS, XLSM, ODS, CSV, TSV | XLSX, ODS, CSV, TSV, PDF, PNG, JPEG |
| Document | DOCX, ODT, RTF, HTML, TXT | DOCX, PDF, PNG, JPEG, HTML, TXT |
| Work backup | PaperLink work file (.paperlink) | PaperLink work file (.paperlink) |
| Template | PaperLink template (.papertemplate) | PaperLink template (.papertemplate) |

Older **.flowdesk** work backups and **.flowtemplate** templates can still be opened. Newly exported backups and templates use **.paperlink** and **.papertemplate**. Existing saved entries are read using their original identifiers; the Windows installer also retains its upgrade identifier across the rename.

For several documents, choose one combined file or separate files in a ZIP folder. Pictures use one image for each page; several images are saved together in a ZIP file. Specific spreadsheet rows and page numbers are available when needed. Finished filenames default to the work name and can be changed before saving.

For **One per row → Separate files in a ZIP folder**, name the ZIP itself under **ZIP file name**. Under **Name the files inside the ZIP**, choose a spreadsheet column (such as Employee name), add optional text before each name, or type individual names in the preview. The preview uses the rows selected for saving. Blank source cells fall back to their row number; matching names receive numbered suffixes to keep every document. Unsupported characters in spreadsheet-based names are replaced automatically. Individual name edits are kept while the save window is open, and **Reset individual names** restores the column-based suggestions. For pictures, a document with several pages is a named ZIP inside the outer ZIP.

The input files are kept. **Save file / Print** creates the output files through a save dialog.


## Current boundaries

This is a working first desktop version for manual evaluation.

- Office imports preserve supported text, tables, embedded images, fonts, alignment, and basic formatting. Complex section layouts, charts, text boxes, comments, tracked changes, and embedded objects are not fully supported. Multiple document sections use one page layout. Word field instructions are not retained; check imported page-number text and use the app's automatic numbering where needed. Legacy DOC files should first be saved as DOCX.
- The document editor flows continuously; the separate live **Check pages** paginates it. Paragraphs split across pages, tables split between rows and repeat header rows, and explicit page breaks are supported. Oversized rows or blocks stop document export rather than silently clipping content.
- Headers and footers automatically reserve the space their content needs. Empty footers use no extra space, page numbers have their own area, and body text flows onto another page when needed. The effective margins are carried into Word export as well as PDF, print, and pictures. A message is shown only when the header and footer leave less than 24 mm for the document body.
- Editable DOCX output may reflow in Word depending on fonts and Word's layout engine. PDF and image exports use the app's page layout. Check corporate templates and generated Word files before issuing them.
- Spreadsheet formulas are retained and common formulas recalculate locally. Formula-derived document fields are read-only; their formulas and input cells are edited in the spreadsheet. Unsupported calculations are shown explicitly and block document export until corrected. Spreadsheet data can still be exported for a full spreadsheet application to recalculate.
- Spreadsheet formatting, charts, macros, external links, and advanced Excel features are not reproduced with full Office fidelity. Macros are not executed.
- The current Windows build is unsigned. Physical printing still needs a manual check with the intended printer.

Use [the manual-testing guide](docs/MANUAL-TESTING.md) and the fictional files in **samples/** to evaluate the workflow.

## Development

Requires Node.js 22 or later on Windows.

    npm ci
    node node_modules/electron/install.js
    npm start

    npm test
    npm run test:desktop
    npm run samples
    npm run package

The Electron runtime download is a setup step, not a runtime internet requirement. The lockfile fixes installed library versions.

Windows packaging uses a small preload to preserve the operating-system environment when the installer helper is launched. This works around the partial environment passed by electron-builder 26.15.3 and is not included in the distributed app.

The app uses Electron for native files, print dialogs, and PDF output; Tiptap for document editing; SheetJS for spreadsheet I/O; and local conversion/layout modules. Third-party license notices are included in **docs/THIRD-PARTY-NOTICES.txt**.

The earlier payroll-only browser prototype is retained in **legacy/payroll-prototype.html**.
