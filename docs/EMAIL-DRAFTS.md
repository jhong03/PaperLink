# Send document emails or save drafts in PaperLink

PaperLink can send reviewed emails through your company's outgoing SMTP server. It also prepares local draft files for review in another email app or company webmail. Preparing attachments or saving a ZIP does not send an email. Sending always requires an explicit Send action and final confirmation.

## First trial: send a test message to yourself

1. Choose **Email documents** and enter your own **Sender address**.
2. Expand **Email server settings**. Ask IT for the SMTP server, port, connection security, username, and password or app password.
3. Use **STARTTLS** for the setting supplied by IT, commonly port 587, or **SSL/TLS**, commonly port 465. A company relay can use **No login**. Unencrypted mode is available only for a relay without a login, when IT specifically requires it.
4. Choose **Test connection**. This checks the connection and configured login without sending a message. It does not verify that the chosen sender address is allowed to send.
5. Enter your own address under **Send a test email to**, then choose **Send test email** and confirm the sender and recipient.
6. Check your inbox and junk folder. The test email contains a short plain message without payroll data, attachments, or CC.

Your server must permit SMTP submission from this device and account. SMTP password/app-password authentication and company relays are supported. Automatic OAuth sign-in for Microsoft or Google accounts is not part of this version; an OAuth-only company account needs an IT-approved SMTP relay or other approved submission credentials. Your existing draft and webmail options remain available.

**Remember password securely on this computer** is optional. Without it, the password stays in the desktop app's memory until the app closes. With it, the password is encrypted using the operating system's protected storage. Passwords are never included in saved work, work backups, or templates. **Forget password** removes the retained password. Server settings and recent sending results are stored in the app data folder, available through **Help → Open saved data folder**.

## Prepare the document first

1. Open your spreadsheet and prepare the document as usual.
2. Choose **One document** for a single email, or **One per row** for a separate email and document for each included spreadsheet row.
3. Choose **Check pages** and review your document.
4. Choose **Email documents** beside **Save file / Print**, or from the desktop File menu.

## Bulk send different documents to different people

1. Put each person's details and email address on the same spreadsheet row, with column headings such as **Employee name**, **Email address**, and **Salary**.
2. Connect the person's cells to the document where their information belongs. Connections set to change with each row will fill each person's document; text typed directly into the document stays the same.
3. In **Email documents**, set **How many emails?** to **One email per spreadsheet row**. This also sets the workspace to **One per row**. Choose the sheet and row containing column headings.
4. In **3. Review each email**, enter To and optional CC directly for the selected person. To fill them automatically, expand **Fill addresses from spreadsheet columns (optional)** and choose the address columns. Each row reads its own addresses.
5. Leave **Spreadsheet rows to include** empty for all filled rows, or enter a selection such as `2-5, 8`.
6. Write the email subject and message.
7. Choose a name column under **Name files using**. Check **Recipients and their attachments**, then choose **Review** beside a row. Expand **Preview this person's document** to inspect that row's document content. The preview uses the PDF layout; Word may reflow when opened.
8. Choose **Prepare emails**, then **Send all … emails**. Each email includes only its own row's generated document. Confirm the final recipient and attachment summary to start sending.

Typing several addresses into the selected email's To box sends that row's document to those addresses. Changes to To or CC apply only to the selected email. Use the arrows or **Review** in the list to enter different addresses for the next person. **Reset addresses** restores the selected email's source values. You can also choose spreadsheet columns to fill all rows automatically, then correct individual emails in the review.

## Fill in the email details

- **Sender address:** Enter the email address this SMTP account is permitted to send from. If opening drafts in another email app, check the sending account there too.
- **Read spreadsheet information from:** One per row uses the same sheet as the documents. For One document, choose a sheet and row for addresses and personalized message information.
- **Spreadsheet rows to include:** Leave empty for all filled rows, or enter a range such as `2-5, 8`. This is available for One per row.
- **To:** Enter or edit addresses in the selected email under **3. Review each email**. These edits apply only to that email. Optional spreadsheet autofill reads the chosen column from each row.
- **CC:** Optional. Enter addresses in the selected email, or choose a column such as Manager email under the review's spreadsheet autofill options. Separate multiple addresses with semicolons. CC recipients receive that email's document too.
- **Subject and Message:** Write the text recipients should receive. Existing saved spreadsheet labels such as `{{Employee name}}` continue to fill from their corresponding rows.
- **Attachment format:** Choose PDF or Word. The attachment includes all pages of that row's document.
- **Name files using:** Choose a name column, optionally with text before each name. Blank cells use a row-based name; matching names receive numbered suffixes.

The **Try an example** screen includes fictional employee and manager addresses for practicing this flow.

## Password-protect attachments

In the Windows app, choose **Password protection** under **2. Attach the document**:

- **No password:** Creates ordinary PDF or Word attachments.
- **Use one password for all attachments:** Enter and confirm a password of up to 64 characters. It applies to every included attachment. The typed password is excluded from saved work and templates, and cleared when the email window closes. Enter it again next time.
- **Use each person's password from a spreadsheet column:** Choose the column containing each row's document password. The displayed cell value is used, including spaces and formatting. Every included row needs a nonblank password; use text cells when leading zeroes matter. Password cells remain part of the imported workbook and saved work.

Choose **Prepare emails** to generate the encrypted documents. Password protection applies to direct email attachments, saved attachment copies, embedded attachments in .eml drafts, and attachment copies in the draft ZIP. The ZIP itself is not encrypted. Passwords are not automatically placed in the message or draft notes; share them separately. The in-app document preview remains readable for review.

Changing password settings clears prepared copies and requires preparation again. Missing passwords, mismatched confirmation, or encryption failure stop preparation; PaperLink does not substitute an unprotected attachment. To and CC corrections retain the protected document. Very long multibyte passwords may need to be shortened for PDF compatibility. Test messages contain no attachments.

PDF protection uses [LibPDF's AES-256 open-password encryption](https://libpdf.documenso.com/docs/guides/encryption). Word protection uses [officecrypto-tool's Office encryption](https://github.com/zurmokeeper/officecrypto-tool), with its default Agile AES-256 / SHA-512 implementation. Processing runs locally in a background worker. Opening a protected document requires a reader that supports its encryption.

## Review and prepare

Use the arrows or email list on the right to review the sender, To, CC, subject, message, and attachment name for each row. Missing or invalid addresses and unresolved message fields must be corrected before preparing the batch.

Choose **Prepare emails**. Each email receives only the document generated for its own row. The progress message shows when the attachments are ready. Close the window to cancel preparation; preparing does not send an email.

Changing To or CC keeps prepared document attachments and updates the affected drafts immediately. Invalid addresses disable sending the affected email and sending the complete batch until corrected. Other changes to email details clear the prepared copies so you can prepare updated drafts. **Save work** keeps the email setup and individual address edits with the spreadsheet and document. Importing a new spreadsheet or clearing the spreadsheet removes individual address edits. The setup is not included in document-only templates.

## Send from PaperLink

Choose **Send this email** to send only the email currently shown, or **Send all emails** to send the prepared batch. The final confirmation shows the SMTP server, sender, recipients, CC, and attachment names. For a large batch, the confirmation summarizes the first emails; the complete list is available in the email preview before confirming.

The results distinguish **Accepted by server**, **Partly accepted**, **Not accepted**, **Check delivery**, and **Not attempted**. Server acceptance means the server took responsibility for the message; it does not confirm delivery to the recipient's inbox. PaperLink does not write a separate copy into the mailbox's Sent folder; the server may do so.

Sending stops on a failure, partial acceptance, or unconfirmed delivery. **Stop remaining** stops after the current email finishes. Closing the email window requests the same stop. Closing the main window during a send stops the remaining messages and keeps the app open until the current message finishes, so its result can be recorded.

Already accepted, partly accepted, and unconfirmed copies are skipped by default. The final confirmation offers an explicit resend checkbox when needed. A lost server response may mean the email was accepted even though the app could not confirm it: check delivery with IT or the recipient before selecting resend. If only some recipients were accepted, check those addresses before making another attempt. Recent delivery history covers the latest 10,000 message identities on this computer; it does not track emails sent later from another app.

## Open or save the drafts

- **Open this draft:** Opens the selected .eml in the Windows email app associated with this file type. An app that supports composing unsent .eml files can open it for editing; another app may show it as a saved message. Choose **Edit as new** or its equivalent if available. Check the sender account, recipients, and attachment before sending.
- **Save this draft:** Saves the selected .eml file with its attachment embedded.
- **Save attachment:** Saves the selected person's PDF or Word document separately.
- **Save drafts (.zip):** Saves all prepared emails together. Extract the ZIP before opening the files. It contains **Drafts**, **Attachments**, **Email details**, and **READ ME.txt**.

Opening a draft uses a temporary local file so the email app can read it. Saving a .eml file does not automatically place it in an account's Drafts folder. Draft copies can include emails already sent from PaperLink, so check the sending results before sending a draft again. PaperLink cannot determine whether you later sent it from another app.

## Company webmail and other email services

Use the preview's **Copy** buttons to copy the sender, To, CC, and subject. Choose **Copy message** and **Save attachment**. Compose a new email in your existing webmail account, paste the details, and attach the matching document. The ZIP includes the same details and separate attachments for working outside PaperLink.

Draft and webmail options work without configuring an SMTP connection in PaperLink. Directly opening an editable .eml depends on the email app. Microsoft documents opening .eml in new Outlook and Outlook on the web, but that does not guarantee editable-draft behavior: [Microsoft's file-opening guide](https://support.microsoft.com/en-us/outlook/mail/open-eml-msg-and-oft-files-in-new-outlook-and-outlook-on-the-web). Thunderbird documents **Edit Message as New**: [Thunderbird keyboard shortcuts](https://support.mozilla.org/en-US/kb/keyboard-shortcuts-thunderbird).

SMTP connection behavior follows [Nodemailer's SMTP transport documentation](https://nodemailer.com/smtp). Gmail app passwords require account eligibility and 2-Step Verification: [Google's app-password guide](https://support.google.com/mail/answer/185833?hl=en). Company policies can restrict or disable SMTP authentication: [Microsoft's SMTP AUTH guide](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission).

For large batches, prepare smaller groups of rows. PaperLink supports up to 1,000 drafts per batch and limits the combined prepared data to 160 MB. Your email service may impose its own attachment-size limits when sending.
