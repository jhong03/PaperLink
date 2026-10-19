# Use a document again with new information

A template remembers your document and where spreadsheet information belongs. For example, it can remember the positions for employee name, current salary, increment, and final salary in a letter.

## Save a template once

1. Prepare the document and connect the values you want from your spreadsheet.
2. Choose **One document** or **One per row**, and set the page layout as needed.
3. Choose **Save template** on the main screen.
4. Check the document name and work name shown in the dialog. These identify the document in the right pane that will be saved.
5. Give the new template a name, such as **Salary letters**, and choose **Save template**.

**Save template** always creates a new template. If the name is already in use, the app asks you to choose another name. Capitalization and extra spaces do not make a name different. Existing templates are never changed by this action.

The template keeps the document, images, page settings, and connected cell locations, such as **Employees!A2**. It does not keep the old connected values or value labels. Connected names and figures will come from the next spreadsheet. Text you typed directly into the document stays the same, so connect any information that should change.

## Use it next time

1. Open **Templates** and choose **Preview document**.
2. Look through the saved document to make sure it is the one you want. Before a spreadsheet is chosen, locations such as **[Employees!A2]** show where information will come from.
3. Choose **Open your spreadsheet** and select the new Excel file. You can also open it before choosing the template.
4. Choose the sheet and compare the two previews. The document immediately shows values from the current spreadsheet at the saved cell locations, including connected items in headers and footers. Changing the file or sheet refreshes those values. Hover over a connected value to see its cell location. Empty cells stay blank. If the template needs several sheets, **View this sheet** lets you check each choice.
5. Choose **Use this template** when the headings match.

Your document is filled from the new spreadsheet. There is no need to open the Word file again or connect each value again.

For **One per row**, this preview shows the first filled row below the headings, just as it will when you use the template. Fields set to stay on a particular cell keep that location. Existing saved templates also use current values; you do not need to recreate them.

The spreadsheet preview is read only. Scroll to see the cells, use **More rows** or **Columns →** to check another area, or enter a cell such as **A50** beside **Go to cell**. The preview shows a small section at a time so large spreadsheets stay manageable. When no sheet is chosen yet, a first look may be shown; choose the sheet above before using the template.

Use the arrows for **One per row** to check each person or item. Then choose **Check pages** and **Save file / Print**.

If a heading does not match, the screen shows which cell to check and what heading should be there. Choose the correct sheet or put the spreadsheet columns in their original positions.

**Use with blank spreadsheet** lets you enter new values yourself. Templates do not include old spreadsheet records or formulas.

## Edit a saved template

1. Open **Templates** and choose **Edit** beside the template you want to change.
2. An **Edit template** bar appears above the workspace, showing the template name.
3. The spreadsheet currently open in your work is copied into the editing session, including its values and formulas. The document and **Check pages** show those values. Try changing values, switch between rows, and adjust the document or connected cells while watching the result. These trials do not change your previous work.
4. If you need another reference spreadsheet, choose **Open file** above the spreadsheet. Without a reference spreadsheet, you can enter preview values yourself; empty connected cells show bracketed locations while editing. Sheets with the same name are used directly; a differently named sheet is matched when its headings identify exactly one match. Unmatched connections keep blank sheets so unrelated values are not filled in automatically.
5. Choose **Save changes** to update that saved template and return to your previous work. Only the document layout, page settings, headings, and cell connections are saved; preview records and formulas are not included in the template.

**Cancel editing** leaves the saved template as it was and returns to your previous work. If you have made changes, the app asks before discarding them. Your previous spreadsheet and document are kept separately while you edit. A renamed template must have a name that is not used by another template.

To keep a separate version, use a template in your work and choose **Save template** with a new name.

## Share a template

Choose **Save file / Print → Template to share or reuse** to create a template file.

On the other device, open **Templates → Open a template file**, then preview and use it. If that name is already saved, enter a different name when asked. Templates saved only inside the app stay on the current device.

## Delete a template

Open **Templates**, choose **Delete** beside the template, and confirm. This removes it from the saved template list. Your current document, saved work, and previously exported files stay in place.

## Where templates and saved work are kept

In the Windows desktop app, **Save work** and **Save template** use an IndexedDB database inside the app's data folder, normally **%APPDATA%\PaperLink**. Existing Flowdesk users keep their previous **%APPDATA%\Flowdesk** data folder unless a newer PaperLink profile already exists. Choose **Help → Open saved data folder** to open the actual data folder for the running app. These entries are database records, rather than separate Word or Excel files. The portable app also uses this Windows profile folder.

Use **Save file / Print → Backup of my work** or **Template to share or reuse** for a file you can keep elsewhere or move to another computer.

PaperLink opens older **.flowdesk** backups and **.flowtemplate** templates as well as its new **.paperlink** and **.papertemplate** files. Existing templates can be used, edited, and saved without recreating them.
