const { app, BrowserWindow, Menu, dialog, ipcMain, session, shell, protocol, net, clipboard, safeStorage } = require('electron');
const fs = require('node:fs/promises');
const { existsSync, mkdirSync } = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
app.setName('PaperLink');
// Keep existing checkpoints in place. New installs use PaperLink's profile;
// upgrades reuse the previous profile without copying or moving a live database.
if (!app.commandLine.hasSwitch('user-data-dir')) {
  const currentProfile = path.join(app.getPath('appData'), 'PaperLink');
  const previousProfile = path.join(app.getPath('appData'), 'Flowdesk');
  const hasStorage = (directory) => ['IndexedDB', 'Local Storage'].some((name) => existsSync(path.join(directory, name)));
  const profile = !hasStorage(currentProfile) && hasStorage(previousProfile) ? previousProfile : currentProfile;
  mkdirSync(profile, { recursive: true }); app.setPath('userData', profile); app.setPath('sessionData', profile);
}
// This origin is a persistence identifier, not the displayed application name.
// Changing it would hide the user's existing IndexedDB records and preferences.
const STORAGE_ORIGIN = 'app://flowdesk';
// A stable browser origin keeps checkpoints available across app updates and portable launches.
protocol.registerSchemesAsPrivileged([{ scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true } }]);
let mainWindow, dirty = false;
require('./document-protection.cjs').registerDocumentProtection({ app, ipcMain, getWindow: () => mainWindow });
const { registerEmailSending } = require('./email-sending.cjs');
const emailSender = registerEmailSending({ app, ipcMain, dialog, safeStorage, getWindow: () => mainWindow, getNodemailer: () => require(app.isPackaged ? path.join(process.resourcesPath, 'mail', 'nodemailer') : 'nodemailer') });
const smoke = process.argv.includes('--smoke');
const filters = {
  spreadsheet: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xls', 'xlsm', 'ods', 'csv', 'tsv'] }],
  document: [{ name: 'Documents', extensions: ['docx', 'odt', 'rtf', 'html', 'htm', 'txt'] }],
  project: [{ name: 'PaperLink work backups', extensions: ['paperlink', 'flowdesk'] }],
  template: [{ name: 'PaperLink templates', extensions: ['papertemplate', 'flowtemplate'] }]
};
const sendAction = (action) => mainWindow?.webContents.send('menu:action', action);
function safeHTML(html) {
  if (typeof html !== 'string' || html.length > 60 * 1024 * 1024) throw new Error('The print document is too large.');
  return html;
}
async function outputWindow(html) {
  const win = new BrowserWindow({ show: false, width: 1300, height: 1600, webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, backgroundThrottling: false } });
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(safeHTML(html)));
  await win.webContents.executeJavaScript('Promise.all([document.fonts.ready, ...Array.from(document.images).map(i => i.complete ? Promise.resolve() : new Promise(r => { i.onload=r; i.onerror=r; }))])');
  return win;
}
app.whenReady().then(() => {
  const bundleRoot = path.join(__dirname, '..', 'dist');
  protocol.handle('app', (request) => {
    const url = new URL(request.url);
    if (url.host !== new URL(STORAGE_ORIGIN).host || !['GET', 'HEAD'].includes(request.method)) return new Response('Not found', { status: 404 });
    let pathname;
    try { pathname = decodeURIComponent(url.pathname); } catch { return new Response('Invalid path', { status: 400 }); }
    const filename = path.resolve(bundleRoot, '.' + (pathname === '/' ? '/index.html' : pathname));
    const relative = path.relative(bundleRoot, filename);
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) return new Response('Invalid path', { status: 400 });
    return net.fetch(pathToFileURL(filename).href);
  });
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    const url = details.url;
    callback({ cancel: !url.startsWith(STORAGE_ORIGIN + '/') && !url.startsWith('file:') && !url.startsWith('data:') && !url.startsWith('blob:') && !url.startsWith('http://127.0.0.1:5173/') });
  });
  mainWindow = new BrowserWindow({
    width: 1520, height: 980, minWidth: 1080, minHeight: 740, show: !smoke,
    title: 'PaperLink', backgroundColor: '#f4f6f7',
    webPreferences: { preload: path.join(__dirname, 'preload.cjs'), nodeIntegration: false, contextIsolation: true, sandbox: true, backgroundThrottling: false, offscreen: smoke, additionalArguments: smoke ? ['--paperlink-test'] : [] }
  });
  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  mainWindow.webContents.on('will-navigate', (event, url) => { if (!url.startsWith(STORAGE_ORIGIN + '/') && !url.startsWith('http://127.0.0.1:5173/')) event.preventDefault(); });
  mainWindow.on('close', (event) => {
    if (emailSender.isSending()) { event.preventDefault(); emailSender.stop(); mainWindow.webContents.send('email:sending-progress', { type: 'stopping' }); return; }
    if (!dirty || smoke) return;
    const choice = dialog.showMessageBoxSync(mainWindow, { type: 'question', title: 'Changes not saved', message: 'Your latest changes are not saved. Choose Keep working to save them, or close without these changes.', buttons: ['Keep working', 'Close without saving'], defaultId: 0, cancelId: 0 });
    if (choice === 0) event.preventDefault();
  });
  const menu = [
    { label: 'File', submenu: [
      { label: 'New work', accelerator: 'CmdOrCtrl+N', click: () => sendAction('new') },
      { label: 'Open Saved Work', accelerator: 'CmdOrCtrl+O', click: () => sendAction('open-project') },
      { label: 'Save work here', accelerator: 'CmdOrCtrl+S', click: () => sendAction('save-project') },
      { label: 'Save template…', click: () => sendAction('save-template') },
      { label: 'Templates…', click: () => sendAction('templates') },
      { type: 'separator' },
      { label: 'Open spreadsheet…', click: () => sendAction('import-sheet') },
      { label: 'Open document…', click: () => sendAction('import-doc') },
      { type: 'separator' },
      { label: 'Save file / Print…', accelerator: 'CmdOrCtrl+P', click: () => sendAction('export') },
      { label: 'Email documents…', click: () => sendAction('email-drafts') },
      { role: 'quit' }
    ] },
    { role: 'editMenu' },
    { label: 'View', submenu: [{ role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }] },
    { label: 'Help', submenu: [
      { label: 'How to use PaperLink', click: () => sendAction('help') },
      { label: 'Quick-start guide', click: () => shell.openPath(path.join(app.isPackaged ? process.resourcesPath : path.join(__dirname, '..'), 'docs', 'QUICK-START.md')) },
      { label: 'How to reuse a template', click: () => shell.openPath(path.join(app.isPackaged ? process.resourcesPath : path.join(__dirname, '..'), 'docs', 'TEMPLATES.md')) },
      { label: 'How to prepare email drafts', click: () => shell.openPath(path.join(app.isPackaged ? process.resourcesPath : path.join(__dirname, '..'), 'docs', 'EMAIL-DRAFTS.md')) },
      { label: 'Open saved data folder', click: async () => {
        const failure = await shell.openPath(session.defaultSession.getStoragePath() || app.getPath('userData'));
        if (failure) dialog.showMessageBox(mainWindow, { type: 'error', title: 'Could not open saved data folder', message: failure });
      } },
      { label: 'Example files', click: () => shell.openPath(path.join(app.isPackaged ? process.resourcesPath : path.join(__dirname, '..'), 'samples')) }
    ] }
  ];
  if (!app.isPackaged) menu.push({ label: 'Development', submenu: [{ role: 'toggleDevTools' }] });
  Menu.setApplicationMenu(Menu.buildFromTemplate(menu));
  if (process.env.PAPERLINK_DEV === '1' || process.env.FLOWDESK_DEV === '1') mainWindow.loadURL('http://127.0.0.1:5173/');
  else mainWindow.loadURL(STORAGE_ORIGIN + '/index.html');
});
ipcMain.on('project:dirty', (_, value) => { dirty = Boolean(value); });
ipcMain.handle('file:open', async (_, kind) => {
  if (!filters[kind]) throw new Error('Unsupported file category.');
  const result = await dialog.showOpenDialog(mainWindow, { properties: ['openFile'], filters: filters[kind] });
  if (result.canceled) return null;
  const selected = result.filePaths[0], stat = await fs.stat(selected);
  if (stat.size > 100 * 1024 * 1024) throw new Error('Choose a file smaller than 100 MB.');
  return { name: path.basename(selected), bytes: new Uint8Array(await fs.readFile(selected)) };
});
ipcMain.handle('file:save', async (_, payload) => {
  if (!payload || !payload.bytes || payload.bytes.byteLength > 200 * 1024 * 1024) throw new Error('Invalid or oversized export.');
  const name = path.basename(String(payload.name || 'export.bin'));
  const extensions = Array.isArray(payload.extensions) ? payload.extensions.filter((ext) => /^[a-z0-9]+$/i.test(ext)) : [];
  const result = await dialog.showSaveDialog(mainWindow, { defaultPath: name, filters: extensions.length ? [{ name: 'Your file', extensions }] : undefined });
  if (result.canceled) return null;
  await fs.writeFile(result.filePath, Buffer.from(payload.bytes));
  return { name: path.basename(result.filePath) };
});
ipcMain.handle('clipboard:copy-text', (_, text) => {
  if (typeof text !== 'string' || text.length > 1000000) throw new Error('This text is too large to copy.');
  clipboard.writeText(text);
});
ipcMain.handle('email:open-draft', async (_, payload) => {
  if (!(payload?.bytes instanceof Uint8Array) || !payload.bytes.byteLength || payload.bytes.byteLength > 180 * 1024 * 1024) throw new Error('This draft is too large to open. Save the draft file instead.');
  const bytes = Buffer.from(payload.bytes), headers = bytes.subarray(0, 65536).toString('ascii').split('\r\n\r\n')[0];
  if (!/^X-Unsent: 1\r?$/mi.test(headers) || !/^MIME-Version: 1\.0\r?$/mi.test(headers)) throw new Error('Choose a draft prepared by PaperLink.');
  const directory = await fs.mkdtemp(path.join(app.getPath('temp'), 'paperlink-draft-'));
  const stem = path.basename(String(payload.name || 'Email draft.eml')).replace(/\.eml$/i, '').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').slice(0, 140).replace(/[. ]+$/, '') || 'Email draft';
  const filename = path.join(directory, 'Draft - ' + stem + '.eml');
  await fs.writeFile(filename, bytes, { flag: 'wx' });
  const error = await shell.openPath(filename);
  if (error) throw new Error('Windows could not open this draft in an email app. Use Save this draft, then open it with an app that supports .eml files. For webmail, copy the email details and use Save attachment.');
  return { opened: true };
});
ipcMain.handle('output:pdf', async (_, html) => {
  const win = await outputWindow(html);
  try { return new Uint8Array(await win.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true, margins: { top: 0, bottom: 0, left: 0, right: 0 } })); }
  finally { win.destroy(); }
});
ipcMain.handle('output:print', async (_, html) => {
  const win = await outputWindow(html);
  try {
    const size = html.match(/@page\{size:([\d.]+)mm ([\d.]+)mm/), width = Number(size?.[1] || 210), height = Number(size?.[2] || 297);
    return await new Promise((resolve) => win.webContents.print({ silent: false, printBackground: true, margins: { marginType: 'none' }, pageSize: { width: Math.round(Math.min(width, height) * 1000), height: Math.round(Math.max(width, height) * 1000) }, landscape: width > height, usePrinterDefaultPageSize: false }, (success, reason) => resolve({ success, reason })));
  } finally { win.destroy(); }
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
