const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const tls = require('node:tls');
const net = require('node:net');

const defaults = { host: '', port: 587, security: 'starttls', auth: 'password', username: '', rememberSecret: false };
const protectedStatuses = new Set(['accepted', 'partial', 'uncertain', 'submitting']);
function addresses(value, required = true) {
  if (!Array.isArray(value) || value.length > 100) throw new Error('Choose up to 100 addresses per email.');
  const result = [];
  for (const address of value) {
    if (typeof address !== 'string' || address.length > 254 || !/^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+$/.test(address) || address.includes('..') || address.split('@')[0].length > 64 || address.startsWith('.') || address.split('@')[0].endsWith('.') || address.split('@')[1].split('.').some((part) => !part || part.length > 63 || part.startsWith('-') || part.endsWith('-'))) throw new Error('Check the email addresses before sending.');
    if (!result.some((previous) => previous.toLowerCase() === address.toLowerCase())) result.push(address);
  }
  if (required && !result.length) throw new Error('Enter a sender and at least one recipient.');
  return result;
}
const hash = (value) => crypto.createHash('sha256').update(value).digest('hex');
function binding(profile) { return JSON.stringify([profile.host.toLowerCase(), profile.port, profile.security, profile.auth, profile.username]); }
function profileInput(value) {
  const host = String(value?.host || '').trim();
  if (!host || host.length > 253 || !net.isIP(host) && !/^[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?$/.test(host) || host.includes('..')) throw new Error('Enter the SMTP server name from your IT team, without https:// or a port number.');
  const port = Number(value.port);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('Enter an SMTP port from 1 to 65535.');
  if (!['starttls', 'tls', 'plain'].includes(value.security) || !['password', 'none'].includes(value.auth)) throw new Error('Choose the connection security and login method.');
  if (value.security === 'plain' && value.auth !== 'none') throw new Error('Use STARTTLS or SSL/TLS when signing in. Unencrypted mode is only available for a company relay without a login.');
  const username = value.auth === 'none' ? '' : String(value.username || '').trim();
  if (value.auth === 'password' && (!username || username.length > 512 || /[\r\n\0]/.test(username))) throw new Error('Enter the SMTP username supplied by your IT team.');
  return { host, port, security: value.security, auth: value.auth, username, rememberSecret: value.auth === 'password' && value.rememberSecret === true };
}
function messageInput(value) {
  const from = addresses([value?.from])[0], to = addresses(value?.to), cc = addresses(value?.cc || [], false).filter((address) => !to.some((recipient) => recipient.toLowerCase() === address.toLowerCase()));
  if (typeof value.subject !== 'string' || !value.subject.trim() || value.subject.length > 10000 || /[\r\n\0]/.test(value.subject)) throw new Error('Enter a valid email subject.');
  if (typeof value.body !== 'string' || value.body.length > 1000000) throw new Error('The email message is too large.');
  if (!(value.attachment instanceof Uint8Array) || !value.attachment.byteLength || value.attachment.byteLength > 100 * 1024 * 1024) throw new Error('Prepare an attachment of up to 100 MB before sending.');
  if (typeof value.attachmentName !== 'string' || value.attachmentName.length > 240 || /[<>:"/\\|?*\u0000-\u001f]/.test(value.attachmentName) || !/\.(?:pdf|docx)$/i.test(value.attachmentName)) throw new Error('Choose a PDF or Word attachment with a valid file name.');
  const attachment = Buffer.from(value.attachment);
  const contentKey = /^[a-f0-9]{64}$/i.test(value.contentKey || '') ? value.contentKey : hash(attachment);
  const fingerprint = hash(JSON.stringify([from.toLowerCase(), [...to].map((item) => item.toLowerCase()).sort(), [...cc].map((item) => item.toLowerCase()).sort(), value.subject, value.body, value.attachmentName, contentKey]));
  return { from, to, cc, subject: value.subject, text: value.body, attachment, attachmentName: value.attachmentName, row: Number.isInteger(value.row) ? value.row : null, fingerprint, passwordProtected: value.passwordProtected === true };
}

function registerEmailSending({ app, ipcMain, dialog, safeStorage, getWindow, getNodemailer }) {
  let profile = { ...defaults }, secret = '', initialized, exclusive = false, active = null;
  let journal = { deliveries: {}, lastReport: null }, journalError = '';
  const settingsFile = () => path.join(app.getPath('userData'), 'paperlink-email-settings.json');
  const historyFile = () => path.join(app.getPath('userData'), 'paperlink-email-history.json');
  const canRemember = () => safeStorage.isEncryptionAvailable() && (process.platform !== 'linux' || !safeStorage.getSelectedStorageBackend || safeStorage.getSelectedStorageBackend() !== 'basic_text');
  async function atomicWrite(filename, value) {
    const temporary = filename + '.' + crypto.randomUUID() + '.tmp';
    await fs.writeFile(temporary, JSON.stringify(value), { flag: 'wx' });
    await fs.rename(temporary, filename);
  }
  function initialize() {
    return initialized ||= (async () => {
      try {
        const saved = JSON.parse(await fs.readFile(settingsFile(), 'utf8'));
        profile = profileInput(saved.profile);
        if (profile.rememberSecret && typeof saved.encryptedSecret === 'string' && canRemember()) {
          try { secret = safeStorage.decryptString(Buffer.from(saved.encryptedSecret, 'base64')); } catch { secret = ''; }
        }
      } catch { profile = { ...defaults }; }
      try {
        const saved = JSON.parse(await fs.readFile(historyFile(), 'utf8'));
        if (!saved.deliveries || typeof saved.deliveries !== 'object' || Array.isArray(saved.deliveries)) throw new Error('Invalid history.');
        journal = { deliveries: {}, lastReport: saved.lastReport && Array.isArray(saved.lastReport.items) ? saved.lastReport : null };
        for (const [key, record] of Object.entries(saved.deliveries)) if (/^[a-f0-9]{64}$/.test(key) && record && typeof record === 'object') journal.deliveries[key] = { ...record, status: record.status === 'submitting' ? 'uncertain' : record.status };
        if (journal.lastReport) {
          for (const item of journal.lastReport.items) if (item.status === 'submitting') { item.status = 'uncertain'; item.message = 'The app closed before confirmation. Check delivery before resending.'; }
          journal.lastReport.finished = true;
        }
      } catch (error) { if (error.code !== 'ENOENT') journalError = 'The email sending history could not be opened. Restore the app data or contact support before sending again.'; }
    })();
  }
  async function saveProfile() {
    const encryptedSecret = profile.rememberSecret && secret ? safeStorage.encryptString(secret).toString('base64') : undefined;
    await atomicWrite(settingsFile(), { version: 1, profile, encryptedSecret });
  }
  async function saveJournal() {
    const entries = Object.entries(journal.deliveries);
    if (entries.length > 10000) journal.deliveries = Object.fromEntries(entries.sort((a, b) => b[1].updatedAt - a[1].updatedAt).slice(0, 10000));
    await atomicWrite(historyFile(), { version: 1, ...journal });
  }
  function configured() {
    profileInput(profile);
    if (profile.auth === 'password' && !secret) throw new Error('Enter your SMTP password or app password in Email server settings.');
  }
  function transport() {
    configured();
    const options = { host: profile.host, port: profile.port, secure: profile.security === 'tls', requireTLS: profile.security === 'starttls', ignoreTLS: profile.security === 'plain', connectionTimeout: 15000, greetingTimeout: 15000, socketTimeout: 45000, disableFileAccess: true, disableUrlAccess: true, logger: false, debug: false, tls: { minVersion: 'TLSv1.2', rejectUnauthorized: true } };
    // Include certificates trusted by the OS, including managed company roots.
    try { if (tls.getCACertificates) options.tls.ca = [...tls.getCACertificates('default'), ...tls.getCACertificates('system')]; } catch { /* Use Node's default trust store. */ }
    if (profile.auth === 'password') options.auth = { user: profile.username, pass: secret };
    return getNodemailer().createTransport(options);
  }
  function readable(error) {
    let message = String(error?.response || error?.message || 'The email server did not accept the request.').replace(/[\r\n]+/g, ' ').slice(0, 700);
    if (secret) message = message.split(secret).join('[hidden]');
    if (error?.code === 'EAUTH') return 'The server refused the login. Check the username and password or app password, and ask IT whether SMTP sending is enabled for this account.';
    if (['ENOTFOUND', 'EAI_AGAIN', 'EDNS'].includes(error?.code)) return 'The SMTP server could not be found. Check the server name and your company network or VPN.';
    if (/CERT|SELF_SIGNED|UNABLE_TO_VERIFY|TLS/.test(error?.code || '') || /certificate/i.test(message)) return 'The SMTP server certificate could not be verified. Check the server name and ask IT to configure a trusted certificate.';
    if (['ECONNREFUSED', 'ECONNECTION', 'ETIMEDOUT', 'ESOCKET'].includes(error?.code)) return 'The connection to the email server failed or timed out. Check the port, connection security, company network, and VPN. ' + message;
    return message;
  }
  function emit(payload) { const window = getWindow(); if (window && !window.isDestroyed()) window.webContents.send('email:sending-progress', payload); }
  const summary = (report) => ({ accepted: report.items.filter((item) => item.status === 'accepted').length, partial: report.items.filter((item) => item.status === 'partial').length, uncertain: report.items.filter((item) => ['uncertain', 'submitting'].includes(item.status)).length, failed: report.items.filter((item) => item.status === 'failed').length, remaining: report.items.filter((item) => item.status === 'pending').length });
  async function sendMessages(messages, jobId, kind) {
    configured(); if (journalError) throw new Error(journalError);
    const duplicates = messages.filter((message) => protectedStatuses.has(journal.deliveries[message.fingerprint]?.status));
    const window = getWindow();
    const list = messages.slice(0, 12).map((message) => 'To: ' + message.to.join('; ') + (message.cc.length ? '\nCC: ' + message.cc.join('; ') : '') + '\nAttachment: ' + (message.attachmentName || '(none — test message)') + (message.passwordProtected ? '\nPassword required to open attachment' : '')).join('\n\n');
    const confirmation = await dialog.showMessageBox(window, { type: 'question', title: kind === 'test' ? 'Send test email now?' : 'Send emails now?', message: kind === 'test' ? 'Send a plain test email without payroll data or attachments?' : 'Send ' + messages.length + ' email' + (messages.length === 1 ? '' : 's') + ' with the reviewed attachments?', detail: 'SMTP server: ' + profile.host + ':' + profile.port + '\nSender: ' + [...new Set(messages.map((message) => message.from))].join('; ') + '\n\n' + list + (messages.length > 12 ? '\n\nAnd ' + (messages.length - 12) + ' more emails from the reviewed list in PaperLink.' : '') + (duplicates.length ? '\n\n' + duplicates.length + ' email(s) were already accepted, partly accepted, or have unconfirmed delivery. They will be skipped unless you choose to resend them.' : ''), buttons: ['Cancel', kind === 'test' ? 'Send test email' : 'Send now'], defaultId: 0, cancelId: 0, ...(duplicates.length ? { checkboxLabel: 'Resend the ' + duplicates.length + ' previously accepted or unconfirmed email(s)', checkboxChecked: false } : {}) });
    if (confirmation.response !== 1) return { canceled: true };
    const report = { jobId, kind, startedAt: Date.now(), finished: false, items: messages.map((message) => ({ row: message.row, to: message.to, cc: message.cc, attachmentName: message.attachmentName || '', status: 'pending', message: '' })) };
    journal.lastReport = report; active = { jobId, cancel: false }; emit({ type: 'start', report });
    let transporter;
    try {
      transporter = transport();
      for (const [index, message] of messages.entries()) {
        if (active.cancel) break;
        const item = report.items[index], previous = journal.deliveries[message.fingerprint];
        if (protectedStatuses.has(previous?.status) && !confirmation.checkboxChecked) {
          Object.assign(item, { status: previous.status === 'submitting' ? 'uncertain' : previous.status, message: 'Skipped: previously attempted. ' + (previous.message || '') });
          emit({ type: 'item', jobId, index, item, summary: summary(report) }); continue;
        }
        Object.assign(item, { status: 'submitting', message: 'Waiting for the email server.' });
        journal.deliveries[message.fingerprint] = { status: 'submitting', updatedAt: Date.now(), to: message.to, cc: message.cc, row: message.row, attachmentName: message.attachmentName || '' };
        await saveJournal(); // A crash after this point must never cause an automatic retry.
        emit({ type: 'item', jobId, index, item, summary: summary(report) });
        try {
          const result = await transporter.sendMail({ from: message.from, to: message.to, cc: message.cc, subject: message.subject, text: message.text, ...(message.attachment ? { attachments: [{ filename: message.attachmentName, content: message.attachment, contentType: message.attachmentName.endsWith('.docx') ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 'application/pdf' }] } : {}), disableFileAccess: true, disableUrlAccess: true });
          item.accepted = result.accepted || []; item.rejected = result.rejected || []; item.messageId = result.messageId;
          item.status = item.rejected.length ? (item.accepted.length ? 'partial' : 'failed') : item.accepted.length ? 'accepted' : 'uncertain';
          item.message = item.status === 'accepted' ? 'Accepted by the email server. Inbox delivery is not confirmed.' : item.status === 'partial' ? 'Some addresses were rejected: ' + item.rejected.join('; ') + '. Do not resend to accepted recipients automatically.' : 'The server did not confirm acceptance. Check delivery before resending.';
        } catch (error) {
          const definitelyRejected = Number(error.responseCode) >= 400 || ['EAUTH', 'EDNS', 'ENOTFOUND', 'EAI_AGAIN', 'ECONNREFUSED', 'EENVELOPE'].includes(error.code);
          item.status = definitelyRejected ? 'failed' : 'uncertain';
          item.message = readable(error) + (definitelyRejected ? '' : ' Delivery is unconfirmed. Check with IT or the recipient before resending.');
        }
        journal.deliveries[message.fingerprint] = { ...item, updatedAt: Date.now() };
        await saveJournal(); emit({ type: 'item', jobId, index, item, summary: summary(report) });
        if (item.status !== 'accepted') break;
      }
    } catch (error) { report.error = 'Sending stopped. ' + readable(error) + ' Check the results before sending again.'; }
    finally {
      transporter?.close();
      for (const item of report.items) if (item.status === 'submitting') { item.status = 'uncertain'; item.message = 'Delivery is unconfirmed. Check before resending.'; }
      report.finished = true; report.finishedAt = Date.now(); report.stopped = active.cancel;
      try { await saveJournal(); } catch { report.error = 'The latest sending results could not be saved. Check delivery before sending again.'; }
      active = null; emit({ type: 'done', report, summary: summary(report) });
    }
    return report;
  }
  function handle(name, callback, lock = true) {
    ipcMain.handle(name, async (event, payload) => {
      if (event.sender !== getWindow()?.webContents || event.senderFrame !== event.sender.mainFrame) throw new Error('Email actions are only available in the PaperLink window.');
      await initialize();
      if (lock && exclusive) throw new Error('Wait for the current email operation to finish.');
      if (lock) exclusive = true;
      try { return await callback(payload); } finally { if (lock) exclusive = false; }
    });
  }
  handle('email:settings', () => ({ profile, hasSecret: Boolean(secret), canRemember: canRemember(), lastReport: journal.lastReport, historyError: journalError, busy: exclusive }), false);
  handle('email:configure', async (value) => {
    const next = profileInput(value), supplied = typeof value.secret === 'string' ? value.secret : '';
    if (supplied.length > 4096) throw new Error('The password is too long.');
    if (next.rememberSecret && !canRemember()) throw new Error('Encrypted password storage is unavailable. Leave Remember password unchecked.');
    const nextSecret = next.auth === 'none' ? '' : supplied || (binding(next) === binding(profile) ? secret : '');
    if (next.auth === 'password' && !nextSecret) throw new Error('Enter the SMTP password or app password for these server settings.');
    const oldProfile = profile, oldSecret = secret; profile = next; secret = nextSecret;
    try { await saveProfile(); } catch (error) { profile = oldProfile; secret = oldSecret; throw error; }
    return { profile, hasSecret: Boolean(secret), canRemember: canRemember() };
  });
  handle('email:forget-password', async () => { secret = ''; profile.rememberSecret = false; await saveProfile(); return { profile, hasSecret: false, canRemember: canRemember() }; });
  handle('email:verify', async () => { const client = transport(); try { await client.verify(); return { verified: true }; } catch (error) { throw new Error(readable(error)); } finally { client.close(); } });
  handle('email:send-test', async (value) => {
    const from = addresses([value?.from])[0], to = addresses([value?.to]);
    const message = { from, to, cc: [], subject: 'PaperLink SMTP test', text: 'This is a test email sent from PaperLink.\n\nThe SMTP server accepted this message for this recipient. No payroll data or document is attached.\n\nPrepared at ' + new Date().toISOString(), row: null };
    message.fingerprint = hash(JSON.stringify(message));
    return sendMessages([message], crypto.randomUUID(), 'test');
  });
  handle('email:send-batch', async (value) => {
    if (!Array.isArray(value?.messages) || !value.messages.length || value.messages.length > 1000 || !/^[a-f0-9-]{36}$/i.test(value.jobId || '')) throw new Error('Prepare between 1 and 1,000 emails before sending.');
    let total = 0;
    const messages = value.messages.map((message) => { const parsed = messageInput(message); total += parsed.attachment.byteLength; return parsed; });
    if (total > 160 * 1024 * 1024) throw new Error('Send a smaller group of emails at a time.');
    return sendMessages(messages, value.jobId, 'documents');
  });
  handle('email:stop', (jobId) => { if (active && (!jobId || active.jobId === jobId)) active.cancel = true; return { stopping: Boolean(active) }; }, false);
  return { isSending: () => Boolean(active), stop: () => { if (active) active.cancel = true; } };
}
module.exports = { registerEmailSending };
