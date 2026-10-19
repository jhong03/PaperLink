const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('paperlink', {
  openFile: (kind) => ipcRenderer.invoke('file:open', kind),
  saveFile: (payload) => ipcRenderer.invoke('file:save', payload),
  renderPDF: (html) => ipcRenderer.invoke('output:pdf', html),
  protectDocument: (payload) => ipcRenderer.invoke('document:protect', payload),
  print: (html) => ipcRenderer.invoke('output:print', html),
  openEmailDraft: (payload) => ipcRenderer.invoke('email:open-draft', payload),
  copyText: (text) => ipcRenderer.invoke('clipboard:copy-text', text),
  emailSettings: () => ipcRenderer.invoke('email:settings'),
  configureEmail: (settings) => ipcRenderer.invoke('email:configure', settings),
  forgetEmailPassword: () => ipcRenderer.invoke('email:forget-password'),
  verifyEmailConnection: () => ipcRenderer.invoke('email:verify'),
  sendEmailTest: (message) => ipcRenderer.invoke('email:send-test', message),
  sendEmailBatch: (batch) => ipcRenderer.invoke('email:send-batch', batch),
  stopEmailSending: (jobId) => ipcRenderer.invoke('email:stop', jobId),
  onEmailSendingProgress: (callback) => { const listener = (_, update) => callback(update); ipcRenderer.on('email:sending-progress', listener); return () => ipcRenderer.removeListener('email:sending-progress', listener); },
  setDirty: (dirty) => ipcRenderer.send('project:dirty', Boolean(dirty)),
  onAction: (callback) => { const listener = (_, action) => callback(action); ipcRenderer.on('menu:action', listener); return () => ipcRenderer.removeListener('menu:action', listener); },
  platform: process.platform,
  isTest: process.argv.includes('--paperlink-test')
});
