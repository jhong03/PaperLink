import { XLSX } from './model.mjs';
import { suggestedFileName, uniqueFileNames } from './export-names.mjs';

export function emailColumns(project, sheet) {
  const columns = new Set(), ws = project.sheet(sheet);
  for (const address in ws) {
    if (!Object.hasOwn(ws, address) || address.startsWith('!') || ws[address]?.v == null || String(ws[address].v).trim() === '') continue;
    const match = address.match(/^([A-Z]+)\d+$/i); if (match) columns.add(match[1].toUpperCase());
  }
  const used = new Set();
  return [...columns].sort((a, b) => XLSX.utils.decode_col(a) - XLSX.utils.decode_col(b)).map((column) => {
    let label = project.display(sheet, column + (project.settings.headerRow + 1)).trim().replace(/[{}\r\n]/g, ' ').slice(0, 100) || 'Column ' + column;
    if (used.has(label.toLowerCase())) label += ' (' + column + ')';
    while (used.has(label.toLowerCase())) label += ' (' + column + ')';
    used.add(label.toLowerCase());
    return { column, label, token: '{{' + label + '}}' };
  });
}

export function emailAddresses(value, required = false) {
  const parts = String(value || '').split(/[;,\r\n]+/).map((part) => part.trim()).filter(Boolean), result = [], used = new Set();
  if (required && !parts.length) throw new Error('Enter an email address or choose a column containing email addresses.');
  for (const address of parts) {
    const [local, domain] = address.split('@');
    if (!/^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+$/.test(address) || address.length > 254 || local.length > 64 || local.startsWith('.') || local.endsWith('.') || address.includes('..') || domain.split('.').some((part) => !part || part.startsWith('-') || part.endsWith('-') || part.length > 63)) throw new Error('Check this email address: ' + address);
    const key = address.toLowerCase(); if (!used.has(key)) { used.add(key); result.push(address); }
  }
  return result;
}

export const recipientOverrideKey = (sheet, row) => JSON.stringify([sheet, row]);

function cellValue(project, sheet, column, row) {
  if (!/^[A-Z]{1,3}$/.test(column) || XLSX.utils.decode_col(column) > 16383) throw new Error('Choose a spreadsheet column again.');
  const address = column + (row + 1), cell = project.cell(sheet, address);
  if (cell?.t === 'e' || cell?.f && cell.v === undefined) throw new Error('Check the calculation in ' + sheet + '!' + address + '.');
  return project.display(sheet, address);
}

export function draftPlan(project, config, sourceColumns = null) {
  const batch = project.settings.mode === 'batch', sheet = batch ? project.settings.batchSheet : config.sheet;
  project.sheet(sheet);
  const rows = batch ? project.recordRows(config.rows) : [Number(config.singleRow) - 1];
  if (!rows.length) throw new Error('There are no filled rows to prepare. Choose the spreadsheet rows to include.');
  if (rows.length > 1000) throw new Error('Prepare up to 1,000 drafts at a time. Choose a smaller range of spreadsheet rows.');
  if (rows.some((row) => !Number.isInteger(row) || row < 0 || row > 1048575)) throw new Error('Choose a spreadsheet row from 1 to 1048576.');
  const columns = sourceColumns || emailColumns(project, sheet), fields = new Map(columns.map((field) => [field.label, field.column]));
  for (const [label, column] of [['Recipient', config.toColumn], ['CC', config.ccColumn], ['File name', config.nameColumn]]) if (column && !columns.some((field) => field.column === column)) throw new Error(label + ': The selected column is missing from this sheet. Choose a column again.');
  const fill = (text, row) => String(text || '').replace(/\{\{([^{}]+)\}\}/g, (_, label) => {
    if (!fields.has(label)) throw new Error('The spreadsheet column "' + label + '" is missing. Update the corresponding {{' + label + '}} text or write a plain message.');
    return cellValue(project, sheet, fields.get(label), row);
  });
  let from = '', senderError = '';
  try { const addresses = emailAddresses(config.from, true); if (addresses.length !== 1) throw new Error('Enter one sender address.'); from = addresses[0]; }
  catch (failure) { senderError = 'Sender: ' + failure.message; }
  const entries = rows.map((row) => {
    const errors = senderError ? [senderError] : [];
    const read = (label, operation, fallback = '') => { try { return operation(); } catch (failure) { errors.push(label + ': ' + failure.message); return fallback; } };
    const key = recipientOverrideKey(sheet, row), overrides = config.recipientOverrides;
    const override = overrides && Object.hasOwn(overrides, key) ? overrides[key] : null;
    const toText = read('Recipient', () => typeof override?.to === 'string' ? override.to : config.toColumn ? cellValue(project, sheet, config.toColumn, row) : config.to);
    const ccText = read('CC', () => typeof override?.cc === 'string' ? override.cc : config.ccColumn ? cellValue(project, sheet, config.ccColumn, row) : config.cc);
    const to = read('Recipient', () => emailAddresses(toText, true), []), toSet = new Set(to.map((address) => address.toLowerCase()));
    const cc = read('CC', () => emailAddresses(ccText), []).filter((address) => !toSet.has(address.toLowerCase()));
    const subject = read('Subject', () => fill(config.subject, row)).replace(/[\r\n]+/g, ' ').trim();
    if (!subject) errors.push('Subject: Enter a subject for the email.');
    const body = read('Message', () => fill(config.body, row));
    const person = config.nameColumn ? read('File name', () => cellValue(project, sheet, config.nameColumn, row)).trim() : '';
    const fallback = 'Document - row ' + (row + 1);
    const name = suggestedFileName([config.prefix?.trim(), person || fallback].filter(Boolean).join(' - '), fallback);
    return { row, sheet, from, to, cc, toText, ccText, subject, body, name, errors };
  });
  return uniqueFileNames(entries).map((entry) => ({ ...entry, attachmentName: entry.name + '.' + (config.format === 'docx' ? 'docx' : 'pdf') }));
}

function base64(bytes) {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 32768) binary += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
  return btoa(binary);
}
const wrappedBase64 = (bytes) => base64(bytes).match(/.{1,76}/g)?.join('\r\n') || '';
function encodedSubject(text) {
  const chunks = []; let part = '';
  for (const character of text) {
    if (new TextEncoder().encode(part + character).length > 42) { chunks.push(part); part = ''; }
    part += character;
  }
  if (part) chunks.push(part);
  return chunks.map((chunk) => '=?UTF-8?B?' + base64(new TextEncoder().encode(chunk)) + '?=').join('\r\n ');
}
function attachmentHeader(name) {
  const fallback = name.replace(/[^A-Za-z0-9._ ()-]/g, '_').slice(0, 160);
  const encoded = encodeURIComponent(name).replace(/['()*]/g, (character) => '%' + character.charCodeAt(0).toString(16).toUpperCase());
  // RFC 2231 continuation segments keep UTF-8 names within header line limits.
  const tokens = encoded.match(/%[0-9A-F]{2}|[^%]/g) || [], segments = []; let segment = '';
  for (const token of tokens) { if (segment.length + token.length > 48) { segments.push(segment); segment = ''; } segment += token; }
  if (segment) segments.push(segment);
  return 'Content-Disposition: attachment;\r\n filename="' + fallback + '";\r\n ' + segments.map((part, index) => 'filename*' + index + '*=' + (index ? '' : "UTF-8''") + part).join(';\r\n ');
}

export function emailDraftBytes(entry, attachment) {
  if (entry.errors?.length) throw new Error('Fix the email details before preparing a draft.');
  const from = emailAddresses(entry.from, true), to = emailAddresses(entry.to.join(', '), true), cc = emailAddresses(entry.cc.join(', '));
  if (from.length !== 1) throw new Error('Enter one sender address.');
  if (!(attachment instanceof Uint8Array) || !attachment.byteLength) throw new Error('The attachment could not be prepared.');
  const boundary = 'PaperLink_' + crypto.randomUUID().replace(/-/g, '');
  const mime = entry.attachmentName.endsWith('.docx') ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf';
  const headers = ['From: ' + from[0], 'To: ' + to.join(',\r\n '), ...(cc.length ? ['Cc: ' + cc.join(',\r\n ')] : []), 'Subject: ' + encodedSubject(entry.subject.replace(/[\r\n]+/g, ' ')), 'Date: ' + new Date().toUTCString(), 'Message-ID: <' + crypto.randomUUID() + '@paperlink.invalid>', 'X-Unsent: 1', 'MIME-Version: 1.0', 'Content-Type: multipart/mixed; boundary="' + boundary + '"'];
  const message = headers.join('\r\n') + '\r\n\r\n--' + boundary + '\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n' + wrappedBase64(new TextEncoder().encode(entry.body.replace(/\r?\n/g, '\r\n'))) + '\r\n--' + boundary + '\r\nContent-Type: ' + mime + '\r\nContent-Transfer-Encoding: base64\r\n' + attachmentHeader(entry.attachmentName) + '\r\n\r\n' + wrappedBase64(attachment) + '\r\n--' + boundary + '--\r\n';
  return new TextEncoder().encode(message);
}
