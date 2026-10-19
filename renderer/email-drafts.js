import './email-drafts.css';
import dialogHtml from './email-dialog.html?raw';
import JSZip from 'jszip';
import { draftPlan, emailColumns, emailDraftBytes, recipientOverrideKey } from '../core/email-drafts.mjs';
import { suggestedFileName, validateFileName } from '../core/export-names.mjs';
import { setupEmailSending } from './email-sending.js';

export function setupEmailDrafts(services) {
  document.body.insertAdjacentHTML('beforeend', dialogHtml);
  const $ = (selector) => document.querySelector(selector), dialog = $('#emailDialog');
  const button = document.createElement('button'); button.dataset.action = 'email-drafts'; button.className = 'action-button';
  button.setAttribute('aria-haspopup', 'dialog'); button.title = 'Review and send document emails, or save drafts';
  button.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m2 6 10 7L22 6"/></svg><span class="action-label">Email documents</span>';
  $('[data-action="export"]').before(button);
  let plan = [], prepared = new Map(), columns = [], working = false, generation = 0, timer, previewKey = '', previewRequest = 0;
  let recipientOverrides = Object.create(null), legacyAddresses = { to: '', cc: '' };
  const mailingRows = new Map(), attachmentCache = new Map();
  const settingIds = ['emailMode', 'emailHeaderRow', 'emailFrom', 'emailSheet', 'emailSingleRow', 'emailRows', 'emailToColumn', 'emailCcColumn', 'emailSubject', 'emailBody', 'emailFormat', 'emailNameColumn', 'emailNamePrefix', 'emailProtection', 'emailPasswordColumn', 'emailDocumentPassword', 'emailDocumentPasswordConfirm'];
  const config = () => ({ from: $('#emailFrom').value, sheet: $('#emailSheet').value, singleRow: $('#emailSingleRow').value, rows: $('#emailRows').value, toColumn: $('#emailToColumn').value, to: legacyAddresses.to, ccColumn: $('#emailCcColumn').value, cc: legacyAddresses.cc, recipientOverrides, subject: $('#emailSubject').value, body: $('#emailBody').value, format: $('#emailFormat').value, nameColumn: $('#emailNameColumn').value, prefix: $('#emailNamePrefix').value, archiveName: $('#emailZipName').value, protection: $('#emailProtection').value, passwordColumn: $('#emailPasswordColumn').value });
  const selected = () => plan.find((entry) => String(entry.row) === $('#emailPreviewRow').value);
  const sending = setupEmailSending({ services, dialog, perform, progress, failure, getFiles: () => prepared, getSelected: selected, getPlan: () => plan });
  function failure(error) { $('#emailError').textContent = error.message || String(error); $('#emailError').classList.remove('hidden'); }
  function progress(message) { $('#emailProgress').textContent = message; }
  function remember() { services.getProject().settings.emailDraft = config(); services.changed(); }
  function controls() {
    for (const id of [...settingIds, 'emailZipName']) $('#' + id).disabled = working;
    for (const option of $('#emailProtection').options) option.disabled = option.value !== 'none' && !services.native?.protectDocument;
    $('#emailFormat').querySelector('[value="pdf"]').disabled = !services.native?.renderPDF;
    const ready = prepared.size === plan.length && plan.length > 0;
    $('#emailPrepare').disabled = working || !plan.length || plan.some((entry) => entry.errors.length);
    $('#emailPrepare').textContent = ready ? 'Prepare again' : 'Prepare emails';
    $('#emailSaveAll').disabled = working || !ready;
    const entry = selected(), file = entry && prepared.get(entry.row);
    for (const id of ['emailPreviewTo', 'emailPreviewCc', 'emailResetAddresses']) $('#' + id).disabled = working || !entry;
    $('#emailOpenDraft').disabled = working || !file || !services.native?.openEmailDraft;
    $('#emailSaveOne').disabled = $('#emailSaveAttachment').disabled = working || !file;
    $('#emailPrevious').disabled = !entry || plan.indexOf(entry) <= 0;
    $('#emailNext').disabled = !entry || plan.indexOf(entry) >= plan.length - 1;
    sending.controls(working);
  }
  function renderMailingList() {
    if (mailingRows.size === plan.length && plan.every((entry) => mailingRows.has(entry.row))) {
      for (const entry of plan) {
        const item = mailingRows.get(entry.row); item.entry = entry;
        item.recipient.textContent = (entry.to.join('; ') || entry.toText || '(enter a recipient below)') + (entry.cc.length ? '\nCC: ' + entry.cc.join('; ') : entry.ccText ? '\nCC: ' + entry.ccText : '');
        item.name.textContent = entry.attachmentName;
      }
      $('#emailMailingList').classList.toggle('hidden', services.getProject().settings.mode !== 'batch');
      return;
    }
    mailingRows.clear();
    const fragment = document.createDocumentFragment();
    for (const entry of plan) {
      const tr = document.createElement('tr'), row = document.createElement('th'), recipient = document.createElement('td'), attachment = document.createElement('td'), name = document.createElement('span'), action = document.createElement('td'), button = document.createElement('button'), state = document.createElement('small');
      row.scope = 'row'; row.textContent = String(entry.row + 1);
      recipient.textContent = (entry.to.join('; ') || entry.toText || '(enter a recipient below)') + (entry.cc.length ? '\nCC: ' + entry.cc.join('; ') : entry.ccText ? '\nCC: ' + entry.ccText : '');
      name.textContent = entry.attachmentName; attachment.append(name, state);
      button.type = 'button'; button.textContent = 'Review'; button.setAttribute('aria-label', 'Review email and document for spreadsheet row ' + (entry.row + 1));
      button.onclick = () => { $('#emailPreviewRow').value = String(entry.row); review(); $('#emailPreview').scrollIntoView({ block: 'nearest', behavior: 'smooth' }); };
      action.append(button); tr.append(row, recipient, attachment, action); fragment.append(tr);
      mailingRows.set(entry.row, { tr, button, state, entry, recipient, name });
    }
    $('#emailMailingRows').replaceChildren(fragment);
    $('#emailMailingList').classList.toggle('hidden', services.getProject().settings.mode !== 'batch');
  }
  function fitDocumentPreview() {
    const host = $('#emailDocumentPages'), page = host.querySelector('.document-page');
    if (page) host.style.zoom = String(Math.min(1, Math.max(.2, ($('#emailDocumentViewport').clientWidth - 24) / parseFloat(page.style.width))));
  }
  new ResizeObserver(fitDocumentPreview).observe($('#emailDocumentViewport'));
  async function previewDocument() {
    const entry = selected(), details = $('#emailDocumentPreview');
    if (!details.open || !dialog.open || !entry || timer) return;
    const key = generation + ':' + entry.row;
    if (previewKey === key) { fitDocumentPreview(); return; }
    const request = ++previewRequest;
    $('#emailDocumentPages').replaceChildren(); $('#emailDocumentPreviewStatus').textContent = 'Preparing document preview...';
    try {
      await services.readyAssets();
      if (request !== previewRequest || !dialog.open || key !== generation + ':' + selected()?.row) return;
      const pages = services.documentPreview(entry);
      $('#emailDocumentPages').replaceChildren(...pages); previewKey = key; fitDocumentPreview();
      $('#emailDocumentPreviewStatus').textContent = entry.sheet + ' · Row ' + (entry.row + 1) + ' · ' + pages.length + ' page' + (pages.length === 1 ? '' : 's') + '. PDF uses this layout; Word may reflow when opened.';
    } catch (error) { if (request === previewRequest) $('#emailDocumentPreviewStatus').textContent = error.message || String(error); }
  }
  $('#emailDocumentPreview').addEventListener('toggle', previewDocument);
  function review() {
    const entry = selected(), file = entry && prepared.get(entry.row);
    $('#emailPreview').classList.toggle('hidden', !entry); $('#emailPreviewEmpty').classList.toggle('hidden', Boolean(entry));
    if (entry) {
      $('#emailPreviewFrom').textContent = entry.from || '(enter a sender address)';
      const key = recipientOverrideKey(entry.sheet, entry.row), override = recipientOverrides[key];
      for (const [id, field, prefix] of [['emailPreviewTo', 'toText', 'Recipient:'], ['emailPreviewCc', 'ccText', 'CC:']]) {
        const input = $('#' + id), errors = entry.errors.filter((message) => message.startsWith(prefix));
        if (input.dataset.entryKey !== key || document.activeElement !== input) input.value = entry[field] || '';
        input.dataset.entryKey = key; input.setAttribute('aria-invalid', String(Boolean(errors.length)));
        $('#' + id + 'Error').textContent = errors.join(' ');
      }
      $('#emailAddressEditHint').textContent = 'These addresses apply only to this email: ' + entry.sheet + ', row ' + (entry.row + 1) + '.' + (override ? ' Edited here.' : '');
      $('#emailResetAddresses').classList.toggle('hidden', !override);
      $('#emailPreviewSubject').textContent = entry.subject || '(enter a subject)';
      $('#emailPreviewBody').textContent = entry.body;
      $('#emailPreviewAttachment').textContent = entry.attachmentName;
      $('#emailAttachmentStatus').textContent = file ? 'Attached · ' + Math.max(1, Math.round(file.attachment.byteLength / 1024)) + ' KB · Ready to send or save as a draft' : 'Choose Prepare emails to create the attachment.';
      $('#emailAttachmentProtection').textContent = entry.passwordProtected ? (file ? 'Password protected. The recipient needs the document password to open this attachment.' : 'Password protection selected. It will be applied when the attachment is prepared.') : 'This attachment does not require a password.';
    }
    const invalid = plan.filter((item) => item.errors.length), first = entry?.errors.length ? entry : invalid[0];
    $('#emailValidation').classList.toggle('hidden', !first);
    if (first) $('#emailValidation').textContent = invalid.length + ' email' + (invalid.length === 1 ? ' needs' : 's need') + ' attention. Spreadsheet row ' + (first.row + 1) + ': ' + first.errors.join(' ');
    for (const [row, item] of mailingRows) {
      item.tr.classList.toggle('current-email', row === entry?.row);
      if (row === entry?.row) item.button.setAttribute('aria-current', 'true'); else item.button.removeAttribute('aria-current');
      item.state.textContent = (item.entry.errors.length ? 'Check email details' : prepared.has(row) ? 'Prepared' : 'Not prepared') + (item.entry.passwordProtected ? ' · Password required' : '');
    }
    controls();
    previewDocument();
  }
  function refresh() {
    clearTimeout(timer); timer = null;
    const previous = $('#emailPreviewRow').value;
    $('#emailError').classList.add('hidden');
    try {
      const header = Number($('#emailHeaderRow').value);
      if (services.getProject().settings.mode === 'batch' && (!Number.isInteger(header) || header < 1 || header > 1048575)) throw new Error('Enter the row containing column headings, from 1 to 1048575.');
      plan = draftPlan(services.getProject(), config(), columns);
      for (const entry of plan) {
        entry.passwordProtected = $('#emailProtection').value !== 'none';
        try { attachmentPassword(entry, $('#emailFormat').value); } catch (error) { entry.errors.push('Attachment password: ' + error.message); }
      }
      $('#emailPreviewRow').replaceChildren(...plan.map((entry) => new Option('Row ' + (entry.row + 1) + ' · ' + entry.name + (entry.errors.length ? ' · Check details' : ' · ' + entry.to.join('; ')), String(entry.row))));
      if (plan.some((entry) => String(entry.row) === previous)) $('#emailPreviewRow').value = previous;
      $('#emailCount').textContent = plan.length + ' separate email' + (plan.length === 1 ? '' : 's') + ' · one document attached to each email';
    } catch (error) { plan = []; $('#emailPreviewRow').replaceChildren(); $('#emailCount').textContent = 'Choose the rows and email details to continue.'; failure(error); }
    renderMailingList();
    const batch = services.getProject().settings.mode === 'batch';
    $('#emailRowsLabel').classList.toggle('hidden', !batch); $('#emailSingleRowLabel').classList.toggle('hidden', batch); $('#emailHeaderRowLabel').classList.toggle('hidden', !batch);
    $('#emailModeHint').textContent = batch ? 'Each included row creates a separate email and its own document. Leave the rows box empty for all filled rows below the headings.' : 'Prepares one email with the document currently set up in your workspace. The chosen row supplies email addresses and message information.';
    $('#emailScopeHint').textContent = batch ? 'Each person receives a separate email with only their own document. This also sets the workspace to One per row.' : 'Choose one email per row to send each person their own document.';
    $('#emailDocumentHint').textContent = services.documentHint();
    const protection = $('#emailProtection').value;
    $('#emailSharedPasswordFields').classList.toggle('hidden', protection !== 'shared');
    $('#emailPasswordColumnLabel').classList.toggle('hidden', protection !== 'column');
    $('#emailProtectionHint').textContent = !services.native?.protectDocument ? 'Password protection is available in the PaperLink Windows app.' : protection === 'none' ? 'Choose a password option to require a password when opening the attachment.' : protection === 'shared' ? 'Use up to 64 characters. This password is cleared when you close this window and is not included in saved work. Share it separately with the recipients.' : 'Each row must contain a nonblank password of up to 64 characters. Passwords in spreadsheet cells are part of your workbook and saved work. Share them separately with recipients.';
    review();
  }
  function attachmentPassword(entry, format) {
    const mode = $('#emailProtection').value;
    if (mode === 'none') return '';
    if (!services.native?.protectDocument) throw new Error('Open the PaperLink Windows app to password-protect attachments.');
    let password;
    if (mode === 'shared') {
      password = $('#emailDocumentPassword').value;
      if (password !== $('#emailDocumentPasswordConfirm').value) throw new Error('The document passwords do not match.');
    } else {
      const column = $('#emailPasswordColumn').value;
      if (!column || !columns.some((field) => field.column === column)) throw new Error('Choose the spreadsheet column containing document passwords.');
      const project = services.getProject(), address = column + (entry.row + 1), cell = project.cell(entry.sheet, address);
      if (cell?.t === 'e' || cell?.f && cell.v === undefined) throw new Error('Check the password cell in spreadsheet row ' + (entry.row + 1) + '.');
      password = project.display(entry.sheet, address);
    }
    if (!password?.trim()) throw new Error('Enter a nonblank document password.');
    if (password.length > 64 || format === 'pdf' && new TextEncoder().encode(password).length > 127) throw new Error('Use a shorter document password (up to 64 characters).');
    return password;
  }
  function invalidate() {
    generation++; prepared.clear(); attachmentCache.clear();
    previewKey = ''; previewRequest++; $('#emailDocumentPages').replaceChildren();
    progress('Email details changed. Prepare the updated emails and attachments before sending.');
    controls();
  }
  function columnChoices(values = {}, sourceColumns = null) {
    columns = sourceColumns || emailColumns(services.getProject(), $('#emailSheet').value);
    for (const [id, firstLabel] of [['emailToColumn', 'Enter To in each email below'], ['emailCcColumn', 'Enter CC in each email below'], ['emailNameColumn', 'Spreadsheet row number'], ['emailPasswordColumn', 'Choose a password column']]) {
      const current = values[id] ?? $('#' + id).value;
      const options = [new Option(firstLabel, ''), ...columns.map((field) => new Option(field.label + ' (column ' + field.column + ')', field.column))];
      if (current && !columns.some((field) => field.column === current)) options.push(new Option('Column ' + current + ' is missing — choose again', current));
      $('#' + id).replaceChildren(...options); $('#' + id).value = current;
    }
  }
  function changed(event) {
    if (working) return;
    if (['emailToColumn', 'emailCcColumn'].includes(event.target.id)) {
      legacyAddresses[event.target.id === 'emailToColumn' ? 'to' : 'cc'] = '';
      refreshRecipients(); return;
    }
    if (['emailMode', 'emailSheet', 'emailHeaderRow'].includes(event.target.id)) {
      const headerRow = Number($('#emailHeaderRow').value);
      if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 1048575) { invalidate(); refresh(); return; }
      services.configureDocuments({ mode: $('#emailMode').value, sheet: $('#emailSheet').value, headerRow: headerRow - 1 });
      columnChoices();
      if ($('#emailMode').value === 'batch' && !$('#emailToColumn').value) {
        const suggestion = recipientColumn(); if (suggestion) $('#emailToColumn').value = suggestion;
      }
      if ($('#emailMode').value === 'batch' && !$('#emailNameColumn').value) {
        $('#emailNameColumn').value = columns.find((field) => /^(?:employee(?:\s+name)?|full\s+name|name)$/i.test(field.label))?.column || '';
      }
    }
    invalidate(); remember(); clearTimeout(timer); timer = setTimeout(refresh, 180);
  }
  for (const id of settingIds) $('#' + id).addEventListener($('#' + id).tagName === 'SELECT' || id === 'emailHeaderRow' ? 'change' : 'input', changed);
  function refreshRecipients(row = null) {
    if (row === null) prepared.clear(); else prepared.delete(row);
    remember(); refresh();
    try {
      for (const entry of plan) {
        if (row !== null && entry.row !== row) continue;
        const cached = attachmentCache.get(entry.row);
        if (cached && !entry.errors.length) prepared.set(entry.row, { ...cached, entry, eml: null });
      }
      progress(attachmentCache.size ? 'Addresses updated. Prepared documents are kept. Correct any address errors before sending the affected emails.' : 'Addresses updated. Choose Prepare emails when the email details are ready.');
    } catch (error) { failure(error); }
    review();
  }
  for (const [id, field] of [['emailPreviewTo', 'to'], ['emailPreviewCc', 'cc']]) $('#' + id).addEventListener('input', () => {
    if (working) return;
    const entry = selected(); if (!entry) return;
    const key = recipientOverrideKey(entry.sheet, entry.row);
    recipientOverrides[key] = { ...recipientOverrides[key], [field]: $('#' + id).value };
    refreshRecipients(entry.row);
  });
  $('#emailResetAddresses').onclick = () => {
    if (working) return;
    const entry = selected(); if (!entry) return;
    delete recipientOverrides[recipientOverrideKey(entry.sheet, entry.row)];
    refreshRecipients(entry.row);
  };
  $('#emailZipName').addEventListener('input', remember);
  $('#emailPreviewRow').onchange = review;
  for (const [id, delta] of [['emailPrevious', -1], ['emailNext', 1]]) $('#' + id).onclick = () => { const select = $('#emailPreviewRow'); select.selectedIndex += delta; review(); };
  dialog.querySelector('form').onsubmit = (event) => event.preventDefault();
  for (const close of dialog.querySelectorAll('[data-email-close]')) close.onclick = () => dialog.close();
  dialog.addEventListener('close', () => { generation++; previewRequest++; previewKey = ''; clearTimeout(timer); timer = null; prepared.clear(); attachmentCache.clear(); plan = []; mailingRows.clear(); $('#emailMailingRows').replaceChildren(); $('#emailDocumentPages').replaceChildren(); $('#emailDocumentPassword').value = ''; $('#emailDocumentPasswordConfirm').value = ''; });
  async function copy(value) {
    try { if (services.native?.copyText) await services.native.copyText(value); else await navigator.clipboard.writeText(value); progress('Copied.'); }
    catch { failure(new Error('Select the text in the preview and press Ctrl+C to copy it.')); }
  }
  for (const copyButton of dialog.querySelectorAll('[data-email-copy]')) copyButton.onclick = () => {
    if (timer) refresh();
    const entry = selected(); if (!entry) return;
    const key = copyButton.dataset.emailCopy, value = ['to', 'cc'].includes(key) ? entry[key].join('; ') : entry[key]; copy(value || '');
  };
  async function prepare() {
    if (working) return;
    refresh();
    if (!plan.length || plan.some((entry) => entry.errors.length)) return;
    const entries = plan.map((entry) => ({ ...entry })), format = config().format, version = ++generation, files = new Map();
    working = true; prepared.clear(); attachmentCache.clear(); services.busy(true); controls();
    let totalBytes = 0;
    try {
      await services.readyAssets();
      for (const [index, entry] of entries.entries()) {
        if (version !== generation || !dialog.open) return;
        progress('Preparing email ' + (index + 1) + ' of ' + entries.length + ' · spreadsheet row ' + (entry.row + 1));
        await new Promise(requestAnimationFrame);
        const output = await services.attachment(entry, format);
        let attachment = output.bytes || output, contentKey = output.contentKey || '';
        if (version !== generation || !dialog.open) return;
        if (entry.passwordProtected) {
          progress('Password-protecting document ' + (index + 1) + ' of ' + entries.length + ' · spreadsheet row ' + (entry.row + 1));
          attachment = new Uint8Array(await services.native.protectDocument({ bytes: attachment, format, password: attachmentPassword(entry, format) }));
          const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(contentKey + '|password-protected-' + format));
          contentKey = [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
        }
        if (version !== generation || !dialog.open) return;
        const eml = emailDraftBytes(entry, attachment);
        totalBytes += attachment.byteLength + eml.byteLength;
        if (totalBytes > 160 * 1024 * 1024) throw new Error('These drafts are too large to prepare together. Choose fewer spreadsheet rows and prepare them in smaller groups.');
        files.set(entry.row, { entry, attachment, eml, contentKey });
      }
      prepared = files;
      for (const [row, file] of files) attachmentCache.set(row, { attachment: file.attachment, contentKey: file.contentKey });
      progress(entries.length + ' email' + (entries.length === 1 ? '' : 's') + ' prepared. Review, then send or save as drafts.');
    } catch (error) { if (dialog.open && version === generation) { failure(error); progress('Emails were not prepared. Check the message above.'); } }
    finally { working = false; services.busy(false); if (dialog.open) review(); }
  }
  $('#emailPrepare').onclick = prepare;
  async function perform(operation) {
    if (working) return;
    working = true; services.busy(true); controls(); $('#emailError').classList.add('hidden');
    try { await operation(); } catch (error) { failure(error); } finally { working = false; services.busy(false); if (dialog.open) controls(); }
  }
  function draftBytes(file) {
    // Address edits keep the attachment and discard its old email envelope.
    // Encode the updated draft only when it is opened or saved.
    return file.eml ||= emailDraftBytes(file.entry, file.attachment);
  }
  $('#emailOpenDraft').onclick = () => perform(async () => {
    const file = prepared.get(selected()?.row); if (!file) return;
    await services.native.openEmailDraft({ bytes: draftBytes(file), name: file.entry.name + '.eml' });
    progress('Draft opened in your email app. Check the sending results, sender, recipients, and attachment before sending.');
  });
  $('#emailSaveOne').onclick = () => perform(async () => {
    const file = prepared.get(selected()?.row); if (!file) return;
    const result = await services.saveBytes(draftBytes(file), file.entry.name + '.eml', 'eml'); if (result) progress('Draft saved: ' + result.name + '.');
  });
  $('#emailSaveAttachment').onclick = () => perform(async () => {
    const file = prepared.get(selected()?.row); if (!file) return;
    const result = await services.saveBytes(file.attachment, file.entry.attachmentName, file.entry.attachmentName.endsWith('.docx') ? 'docx' : 'pdf'); if (result) progress('Attachment saved: ' + result.name + '. Attach it when composing your email.');
  });
  $('#emailSaveAll').onclick = () => perform(async () => {
    if (!prepared.size || prepared.size !== plan.length) return;
    const archiveName = validateFileName($('#emailZipName').value), zip = new JSZip(), version = generation;
    progress('Preparing the draft ZIP file.');
    for (const file of prepared.values()) {
      const { entry, attachment } = file;
      zip.file('Drafts/' + entry.name + '.eml', draftBytes(file));
      zip.file('Attachments/' + entry.attachmentName, attachment);
      zip.file('Email details/' + entry.name + '.txt', 'Spreadsheet: ' + entry.sheet + ', row ' + (entry.row + 1) + '\r\nFrom: ' + entry.from + '\r\nTo: ' + entry.to.join('; ') + '\r\nCC: ' + entry.cc.join('; ') + '\r\nSubject: ' + entry.subject + '\r\nAttachment: ' + entry.attachmentName + '\r\n\r\n' + entry.body.replace(/\r?\n/g, '\r\n'));
    }
    zip.file('READ ME.txt', 'PaperLink email drafts\r\n\r\nSaving this ZIP does not send emails. Draft copies may include emails already sent from PaperLink. Check the sending results before sending a copy again.\r\n\r\n1. Extract the ZIP before opening files.\r\n2. Open a file in Drafts with an email app that supports .eml files.\r\n3. If it opens as a saved message, choose Edit as new or your app\'s equivalent.\r\n4. Check the sending account, To, CC, subject, message, and attachment. Send from your email app when ready.\r\n\r\nWebmail or an app that cannot compose from .eml:\r\nCreate a new email, copy the details from the matching file in Email details, and attach the matching document from Attachments.\r\n\r\n.eml editing support varies by app. These files are not synchronized into an account\'s Drafts folder. Your email app determines the actual sending account and its permissions.\r\n');
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 3 } });
    if (version !== generation || !dialog.open) return;
    const result = await services.saveBytes(bytes, archiveName + '.zip', 'zip'); if (result) progress('Drafts saved: ' + result.name + '.');
  });
  function recipientColumn() {
    return columns.find((field) => /^(?:e-?mail(?:\s+address)?|employee\s+e-?mail(?:\s+address)?|recipient(?:\s+e-?mail(?:\s+address)?)?)$/i.test(field.label))?.column || '';
  }
  function open() {
    if (working) return;
    const project = services.getProject(), saved = project.settings.emailDraft && typeof project.settings.emailDraft === 'object' ? project.settings.emailDraft : {}, batch = project.settings.mode === 'batch';
    const text = (key, fallback = '', length = 20000) => typeof saved[key] === 'string' ? saved[key].slice(0, length) : fallback;
    legacyAddresses = { to: text('to', '', 4000), cc: text('cc', '', 4000) };
    recipientOverrides = Object.create(null);
    if (saved.recipientOverrides && typeof saved.recipientOverrides === 'object' && !Array.isArray(saved.recipientOverrides)) {
      for (const [key, value] of Object.entries(saved.recipientOverrides).slice(0, 10000)) {
        if (!value || typeof value !== 'object') continue;
        try {
          const [sheet, row] = JSON.parse(key);
          if (!project.workbook.SheetNames.includes(sheet) || !Number.isInteger(row) || row < 0 || row > 1048575 || key !== recipientOverrideKey(sheet, row)) continue;
          const override = {};
          for (const field of ['to', 'cc']) if (typeof value[field] === 'string') override[field] = value[field].slice(0, 4000);
          if (Object.keys(override).length) recipientOverrides[key] = override;
        } catch { /* Ignore malformed saved address entries. */ }
      }
    }
    $('#emailSheet').replaceChildren(...project.workbook.SheetNames.map((sheet) => new Option(sheet, sheet)));
    $('#emailSheet').value = batch ? project.settings.batchSheet : project.workbook.SheetNames.includes(saved.sheet) ? saved.sheet : project.settings.sheet;
    $('#emailMode').value = project.settings.mode; $('#emailHeaderRow').value = project.settings.headerRow + 1;
    columns = emailColumns(project, $('#emailSheet').value);
    const toDefault = recipientColumn();
    const nameDefault = columns.find((field) => /^(?:employee(?:\s+name)?|full\s+name|name)$/i.test(field.label))?.column || '';
    $('#emailProtection').value = ['shared', 'column'].includes(saved.protection) ? saved.protection : 'none';
    $('#emailDocumentPassword').value = ''; $('#emailDocumentPasswordConfirm').value = '';
    columnChoices({ emailToColumn: text('toColumn', toDefault, 3), emailCcColumn: text('ccColumn', '', 3), emailNameColumn: text('nameColumn', nameDefault, 3), emailPasswordColumn: text('passwordColumn', '', 3) }, columns);
    for (const [id, key, fallback, max] of [['emailFrom', 'from', '', 254], ['emailSubject', 'subject', 'Your document', 600], ['emailBody', 'body', 'Hello,\n\nPlease find your document attached.\n\nThank you.', 20000], ['emailRows', 'rows', '', 200], ['emailNamePrefix', 'prefix', '', 80]]) $('#' + id).value = text(key, fallback, max);
    $('#emailSingleRow').value = Number(saved.singleRow) >= 1 && Number(saved.singleRow) <= 1048576 ? Number(saved.singleRow) : project.settings.activeRow + 1;
    $('#emailFormat').value = saved.format === 'docx' || !services.native?.renderPDF ? 'docx' : 'pdf';
    $('#emailZipName').value = text('archiveName', suggestedFileName(project.names.project + ' - email drafts', 'Email drafts'), 160);
    $('#emailRowsLabel').classList.toggle('hidden', !batch); $('#emailSingleRowLabel').classList.toggle('hidden', batch);
    $('#emailModeHint').textContent = batch ? 'One per row: every included row gets its own email and document. Leave the rows box empty to include all filled rows.' : 'One document: prepares one email. The selected row supplies addresses and personalized message information.';
    prepared.clear(); attachmentCache.clear(); generation++; previewKey = ''; progress('Prepare attachments before sending document emails. You can send a plain test message from Email server settings.'); refresh(); dialog.showModal(); previewDocument(); sending.load();
  }
  return { open };
}
