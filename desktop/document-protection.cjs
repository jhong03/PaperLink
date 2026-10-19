const { Worker } = require('node:worker_threads');
const path = require('node:path');
const { randomUUID } = require('node:crypto');

function registerDocumentProtection({ app, ipcMain, getWindow }) {
  let worker, active;
  function fail() {
    if (active) { clearTimeout(active.timer); active.reject(new Error('Document protection could not finish. Prepare the emails again.')); active = null; }
  }
  function startWorker() {
    const filename = app.isPackaged ? path.join(process.resourcesPath, 'protection-worker.cjs') : path.join(__dirname, '..', '.cache', 'protection-worker.cjs');
    const instance = new Worker(filename); worker = instance;
    instance.on('message', (result) => {
      if (!active || active.id !== result.id) return;
      const task = active; active = null; clearTimeout(task.timer);
      if (result.error) task.reject(new Error(result.error)); else task.resolve(result.bytes);
    });
    instance.on('error', () => { if (worker === instance) { worker = null; fail(); } });
    instance.on('exit', () => { if (worker === instance) { worker = null; fail(); } });
    instance.unref();
  }
  ipcMain.handle('document:protect', async (event, value) => {
    const window = getWindow();
    if (!window || event.sender !== window.webContents || event.senderFrame !== window.webContents.mainFrame) throw new Error('Open document protection from the PaperLink window.');
    if (active) throw new Error('Wait for the current document to finish password protection.');
    if (!['pdf', 'docx'].includes(value?.format) || !(value.bytes instanceof Uint8Array) || !value.bytes.byteLength || value.bytes.byteLength > 100 * 1024 * 1024) throw new Error('Choose a PDF or Word attachment of up to 100 MB.');
    if (typeof value.password !== 'string' || !value.password.trim() || value.password.length > 64 || value.format === 'pdf' && Buffer.byteLength(value.password, 'utf8') > 127) throw new Error('Enter a nonblank document password of up to 64 characters. Use a shorter password if it contains many symbols or non-Latin characters.');
    const header = Buffer.from(value.bytes.buffer, value.bytes.byteOffset, Math.min(5, value.bytes.byteLength));
    if (value.format === 'pdf' ? header.toString() !== '%PDF-' : header[0] !== 0x50 || header[1] !== 0x4b) throw new Error('The attachment format could not be read. Prepare the document again.');
    if (!worker) startWorker();
    return new Promise((resolve, reject) => {
      const id = randomUUID(), timer = setTimeout(() => {
        const previous = worker; worker = null; fail(); previous?.terminate();
      }, 120000);
      active = { id, resolve, reject, timer };
      try { worker.postMessage({ id, bytes: value.bytes, format: value.format, password: value.password }); }
      catch { fail(); }
      finally { value.password = ''; }
    });
  });
  app.on('before-quit', () => { const previous = worker; worker = null; fail(); previous?.terminate(); });
}
module.exports = { registerDocumentProtection };
