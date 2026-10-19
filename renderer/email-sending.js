import './email-sending.css';
import { emailAddresses } from '../core/email-drafts.mjs';

const statusNames = { pending: 'Not attempted', submitting: 'Waiting for server', accepted: 'Accepted by server', partial: 'Partly accepted', failed: 'Not accepted', uncertain: 'Check delivery' };
export function setupEmailSending({ services, dialog, perform, progress, failure, getFiles, getSelected, getPlan }) {
  const $ = (selector) => dialog.querySelector(selector), native = services.native;
  const serverIds = ['emailSmtpHost', 'emailSmtpPort', 'emailSmtpSecurity', 'emailSmtpAuth', 'emailSmtpUser', 'emailSmtpPassword', 'emailSmtpRemember', 'emailSmtpSave', 'emailSmtpVerify', 'emailSmtpForget', 'emailTestTo', 'emailSendTest'];
  let stored = null, hasSecret = false, canRemember = false, working = false, loading = false, sending = false, stopRequested = false, activeJobId = null, report = null;
  const resultRows = new Map();
  const smtpStatus = (text) => { $('#emailSmtpStatus').textContent = text; };
  const values = () => ({ host: $('#emailSmtpHost').value.trim(), port: Number($('#emailSmtpPort').value), security: $('#emailSmtpSecurity').value, auth: $('#emailSmtpAuth').value, username: $('#emailSmtpUser').value.trim(), secret: $('#emailSmtpPassword').value, rememberSecret: $('#emailSmtpRemember').checked });
  const sameSettings = () => stored && ['host', 'port', 'security', 'auth', 'username'].every((key) => values()[key] === stored[key]);
  function controls(isWorking) {
    working = isWorking;
    const blocked = working || loading;
    for (const id of serverIds) $('#' + id).disabled = blocked || !native?.configureEmail;
    const noLogin = $('#emailSmtpAuth').value === 'none';
    for (const id of ['emailSmtpUser', 'emailSmtpPassword']) $('#' + id).disabled = blocked || !native?.configureEmail || noLogin;
    $('#emailSmtpRemember').disabled = blocked || !native?.configureEmail || noLogin || !canRemember;
    $('#emailSmtpForget').disabled = blocked || !hasSecret;
    $('#emailSmtpPassword').placeholder = hasSecret && sameSettings() ? 'Password is already available for these settings' : 'Enter the SMTP password';
    $('#emailSendOne').disabled = blocked || !native?.sendEmailBatch || !getFiles().has(getSelected()?.row);
    $('#emailSendAll').disabled = blocked || !native?.sendEmailBatch || !getPlan().length || getFiles().size !== getPlan().length || getPlan().some((entry) => entry.errors.length);
    $('#emailSendAll').classList.toggle('hidden', getPlan().length < 2);
    $('#emailSendAll').textContent = 'Send all ' + getPlan().length + ' emails';
    $('#emailStopSending').classList.toggle('hidden', !sending);
    $('#emailStopSending').disabled = !sending || stopRequested;
  }
  function resultRow(index, item) {
    let row = resultRows.get(index);
    if (!row) {
      const tr = document.createElement('tr'), recipient = document.createElement('td'), result = document.createElement('td'), label = document.createElement('strong'), detail = document.createElement('small');
      result.append(label, detail); tr.append(recipient, result); $('#emailSendResultsBody').append(tr);
      row = { recipient, result, label, detail }; resultRows.set(index, row);
    }
    row.recipient.textContent = (item.row == null ? 'Test email' : 'Row ' + (item.row + 1)) + ' · ' + item.to.join('; ') + (item.cc.length ? '\nCC: ' + item.cc.join('; ') : '') + (item.attachmentName ? '\n' + item.attachmentName : '');
    row.label.textContent = statusNames[item.status] || item.status; row.result.className = 'email-result-' + item.status; row.detail.textContent = item.message || '';
  }
  function summary() {
    if (!report) return '';
    const count = (status) => report.items.filter((item) => item.status === status).length;
    return count('accepted') + ' accepted · ' + count('partial') + ' partly accepted · ' + count('failed') + ' not accepted · ' + (count('uncertain') + count('submitting')) + ' unconfirmed · ' + count('pending') + ' not attempted' + (report.stopped ? ' · Remaining emails stopped' : '');
  }
  function showReport(next) {
    report = next; resultRows.clear(); $('#emailSendResultsBody').replaceChildren();
    $('#emailSendResults').classList.toggle('hidden', !report);
    if (!report) return;
    report.items.forEach((item, index) => resultRow(index, item)); $('#emailSendSummary').textContent = summary();
    if (report.error) failure(new Error(report.error));
  }
  native?.onEmailSendingProgress?.((event) => {
    if (event.type === 'stopping') { progress('Stopping remaining emails. Wait for the current email to finish, then close the window again.'); services.status?.('Stopping remaining emails. Wait for the current email to finish, then close the window again.'); }
    if (event.type === 'start') { activeJobId = event.report.jobId; showReport(event.report); }
    if (event.type === 'item' && report?.jobId === event.jobId) { report.items[event.index] = event.item; resultRow(event.index, event.item); $('#emailSendSummary').textContent = summary(); progress('Sending results: ' + summary()); }
    if (event.type === 'done') { activeJobId = null; showReport(event.report); progress(summary()); services.status?.('Email sending results: ' + summary()); }
  });
  async function configure() {
    try {
      const value = values();
      if (!value.host) { $('#emailSmtpSettings').open = true; $('#emailSmtpHost').focus(); throw new Error('Enter the SMTP server settings supplied by your IT team first.'); }
      if (value.auth === 'none') { value.username = ''; value.secret = ''; value.rememberSecret = false; }
      const result = await native.configureEmail(value);
      stored = result.profile; hasSecret = result.hasSecret; canRemember = result.canRemember;
      $('#emailSmtpPassword').value = ''; controls(working);
      smtpStatus('Settings saved. Use Test connection, then Send test email to check this account.');
    } catch (error) { $('#emailSmtpSettings').open = true; smtpStatus(error.message); throw error; }
  }
  $('#emailSmtpSave').onclick = () => perform(configure);
  $('#emailSmtpVerify').onclick = () => perform(async () => {
    await configure(); if (!dialog.open) return;
    smtpStatus('Checking the server and login. No email is being sent.');
    await native.verifyEmailConnection(); smtpStatus('Connection and login checked. Send a test email to check that the sender address can send.');
  });
  $('#emailSmtpForget').onclick = () => perform(async () => {
    const result = await native.forgetEmailPassword(); stored = result.profile; hasSecret = false;
    $('#emailSmtpPassword').value = ''; $('#emailSmtpRemember').checked = false; smtpStatus('Password forgotten. Enter it again when you want to send.');
  });
  async function send(operation, jobId = null) {
    sending = true; stopRequested = false; activeJobId = jobId; controls(working);
    try {
      const result = await operation();
      if (result.canceled) progress('Sending canceled.'); else { showReport(result); progress(summary()); }
    } finally { sending = false; activeJobId = null; controls(working); }
  }
  $('#emailSendTest').onclick = () => perform(async () => {
    const from = emailAddresses($('#emailFrom').value, true), to = emailAddresses($('#emailTestTo').value || $('#emailFrom').value, true);
    if (from.length !== 1 || to.length !== 1) throw new Error('Enter one sender address and one test recipient address.');
    await configure(); if (!dialog.open) return;
    await send(() => native.sendEmailTest({ from: from[0], to: to[0] }));
  });
  function sendFiles(files) {
    return perform(async () => {
      if (!files.length) throw new Error('Prepare the emails and their attachments first.');
      await configure(); if (!dialog.open) return;
      const messages = files.map(({ entry, attachment, contentKey }) => ({ from: entry.from, to: entry.to, cc: entry.cc, subject: entry.subject, body: entry.body, attachmentName: entry.attachmentName, attachment, row: entry.row, contentKey, passwordProtected: entry.passwordProtected === true }));
      const jobId = crypto.randomUUID();
      await send(() => native.sendEmailBatch({ jobId, messages }), jobId);
    });
  }
  $('#emailSendOne').onclick = () => sendFiles([getFiles().get(getSelected()?.row)].filter(Boolean));
  $('#emailSendAll').onclick = () => sendFiles([...getFiles().values()]);
  async function stop() {
    if (!sending) return;
    stopRequested = true; controls(working);
    try { await native.stopEmailSending(activeJobId); progress('Stopping the remaining emails after the current email finishes.'); services.status?.('Stopping email sending after the current email finishes.'); }
    catch (error) { failure(error); }
  }
  $('#emailStopSending').onclick = stop;
  dialog.addEventListener('close', () => { if (sending) stop(); $('#emailSmtpPassword').value = ''; });
  $('#emailSmtpSecurity').onchange = () => {
    const security = $('#emailSmtpSecurity').value, port = Number($('#emailSmtpPort').value);
    if ([25, 465, 587].includes(port)) $('#emailSmtpPort').value = security === 'tls' ? 465 : security === 'plain' ? 25 : 587;
    if (security === 'plain') { $('#emailSmtpAuth').value = 'none'; $('#emailSmtpRemember').checked = false; }
    controls(working);
  };
  for (const id of ['emailSmtpHost', 'emailSmtpPort', 'emailSmtpUser']) $('#' + id).addEventListener('input', () => controls(working));
  $('#emailSmtpAuth').onchange = () => { if ($('#emailSmtpAuth').value === 'none') $('#emailSmtpRemember').checked = false; controls(working); };
  function fillDefaults() {
    if (!$('#emailSmtpUser').value) $('#emailSmtpUser').value = $('#emailFrom').value.trim();
    if (!$('#emailTestTo').value) $('#emailTestTo').value = $('#emailFrom').value.trim();
  }
  $('#emailSmtpSettings').addEventListener('toggle', () => { if ($('#emailSmtpSettings').open) fillDefaults(); });
  async function load() {
    if (!native?.emailSettings) { smtpStatus('Direct sending is available in the desktop app. You can save drafts here.'); controls(false); return; }
    loading = true; controls(working);
    try {
      const result = await native.emailSettings(); stored = result.profile; hasSecret = result.hasSecret; canRemember = result.canRemember;
      for (const [id, key] of [['emailSmtpHost', 'host'], ['emailSmtpPort', 'port'], ['emailSmtpSecurity', 'security'], ['emailSmtpAuth', 'auth'], ['emailSmtpUser', 'username']]) $('#' + id).value = stored[key];
      $('#emailSmtpRemember').checked = stored.rememberSecret; $('#emailSmtpPassword').value = ''; $('#emailTestTo').value = $('#emailFrom').value.trim();
      $('#emailSmtpSettings').open = !stored.host || stored.auth === 'password' && !hasSecret;
      fillDefaults(); showReport(result.lastReport);
      smtpStatus(result.historyError || (stored.host ? 'Server settings loaded. Use Test connection or send a test email to your own address.' : 'Enter the outgoing SMTP settings from your IT team.'));
      controls(working);
    } catch (error) { failure(error); } finally { loading = false; controls(working); }
  }
  return { controls, load };
}
