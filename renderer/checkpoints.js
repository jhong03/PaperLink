// Browser storage holds working checkpoints; only Export / Print creates user files.
let connection;
function database() {
  if (!connection) connection = new Promise((resolve, reject) => {
    // Retain the existing database identifier across the product rename.
    const request = indexedDB.open('flowdesk-checkpoints', 2);
    request.onupgradeneeded = () => {
      const db = request.result;
      for (const name of ['projects', 'summaries', 'templates', 'templateSummaries']) if (!db.objectStoreNames.contains(name)) db.createObjectStore(name, { keyPath: 'id' });
      if (!db.objectStoreNames.contains('preferences')) db.createObjectStore('preferences');
    };
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error('Close other PaperLink windows, then try saving again.'));
    request.onsuccess = () => {
      const db = request.result;
      db.onversionchange = () => { db.close(); connection = undefined; };
      resolve(db);
    };
  }).catch((failure) => { connection = undefined; throw failure; });
  return connection;
}
async function read(store, key) {
  const db = await database();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(store, 'readonly');
    const request = key === undefined ? transaction.objectStore(store).getAll() : transaction.objectStore(store).get(key);
    transaction.oncomplete = () => resolve(request.result);
    transaction.onabort = () => reject(transaction.error || new Error('Unable to read browser checkpoints.'));
  });
}
async function write(stores, operation) {
  const db = await database();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(stores, 'readwrite', { durability: 'strict' });
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error || new Error('The browser checkpoint was not saved.'));
    operation(transaction);
  });
}
export async function saveCheckpoint(id, name, serialized) {
  const summary = { id, name, updatedAt: Date.now() };
  await write(['projects', 'summaries', 'preferences'], (transaction) => {
    transaction.objectStore('projects').put({ id, serialized });
    transaction.objectStore('summaries').put(summary);
    transaction.objectStore('preferences').put(id, 'activeCheckpoint');
  });
  return summary;
}
export async function listCheckpoints() {
  return (await read('summaries')).sort((a, b) => b.updatedAt - a.updatedAt);
}
export async function getCheckpoint(id) {
  const checkpoint = await read('projects', id);
  if (!checkpoint) throw new Error('This checkpoint is no longer available in browser storage.');
  return checkpoint;
}
export async function lastCheckpoint() {
  const id = await read('preferences', 'activeCheckpoint');
  return id ? getCheckpoint(id) : null;
}
export function selectCheckpoint(id) {
  return write(['preferences'], (transaction) => transaction.objectStore('preferences').put(id, 'activeCheckpoint'));
}
export function deleteCheckpoint(id) {
  return write(['projects', 'summaries', 'preferences'], (transaction) => {
    transaction.objectStore('projects').delete(id);
    transaction.objectStore('summaries').delete(id);
    const preferences = transaction.objectStore('preferences'), active = preferences.get('activeCheckpoint');
    active.onsuccess = () => { if (active.result === id) preferences.delete('activeCheckpoint'); };
  });
}
export const templateNameKey = (name) => String(name).normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase();
export async function saveTemplate(id, name, serialized, fieldCount, mode, operation = 'create') {
  if (!['create', 'edit'].includes(operation)) throw new Error('Choose whether to save a new template or edit a saved template.');
  if (operation === 'create' || operation === 'edit') {
    const db = await database();
    await new Promise((resolve, reject) => {
      const transaction = db.transaction(['templates', 'templateSummaries'], 'readwrite', { durability: 'strict' });
      let failure;
      transaction.oncomplete = () => resolve();
      transaction.onabort = () => reject(failure || transaction.error || new Error('The template was not saved.'));
      const request = transaction.objectStore('templateSummaries').getAll();
      request.onsuccess = () => {
        const summaries = request.result;
        if (operation === 'edit' && !summaries.some((item) => item.id === id)) failure = new Error('This template has been deleted. Cancel editing and use Save template to create a new one.');
        else if (operation === 'create' && summaries.some((item) => item.id === id)) failure = new Error('This template already exists. Choose another name for the new template.');
        else if (summaries.some((item) => item.id !== id && templateNameKey(item.name) === templateNameKey(name))) failure = new Error('A template named “' + name + '” already exists. Choose a different name. To change the saved template, use Templates → Edit.');
        if (failure) { transaction.abort(); return; }
        transaction.objectStore('templates').put({ id, serialized });
        transaction.objectStore('templateSummaries').put({ id, name, fieldCount, mode, updatedAt: Date.now() });
      };
    });
    return;
  }
}
export async function listTemplates() {
  return (await read('templateSummaries')).sort((a, b) => b.updatedAt - a.updatedAt);
}
export async function getTemplate(id) {
  const template = await read('templates', id);
  if (!template) throw new Error('This template is no longer available in browser storage.');
  return template;
}
export function deleteTemplate(id) {
  return write(['templates', 'templateSummaries'], (transaction) => {
    transaction.objectStore('templates').delete(id);
    transaction.objectStore('templateSummaries').delete(id);
  });
}
