import './style.css';
import JSZip from 'jszip';
import { generateHTML } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import { Project, XLSX, escapeHtml, clone, parseCell, parseRanges, colName } from '../core/model.mjs';
import { importDocument } from '../core/importers.js';
import { geometry, resolveHtml, paginate, printHtml, renumber, PAGE_EDGE_CSS } from '../core/layout.js';
import { documentBytes, spreadsheetBytes } from '../core/exports.js';
import { stripOutputExtension, validateFileName, suggestedFileName, uniqueFileNames } from '../core/export-names.mjs';
import { createEditor } from './editor.js';
import { setupPaneSplit } from './pane-split.js';
import { startDragFeedback } from './drag-feedback.js';
import { setupEmailDrafts } from './email-drafts.js';
import { saveCheckpoint, listCheckpoints, getCheckpoint, lastCheckpoint, selectCheckpoint, deleteCheckpoint, saveTemplate as storeTemplate, listTemplates, getTemplate, deleteTemplate, templateNameKey } from './checkpoints.js';
import { createTemplate, deserializeTemplate, templateFieldCount, templateCompatibility, suggestTemplateWorksheets, projectForTemplateEditing, applyTemplate as projectFromTemplate } from '../core/templates.mjs';

const edgeStyles = document.createElement('style'); edgeStyles.textContent = PAGE_EDGE_CSS; document.head.append(edgeStyles);
const $ = (s) => document.querySelector(s), native = window.paperlink;
let project = new Project(), activeEditor, mainEditor, headerEditor, footerEditor;
let syncing = false, dirty = false, previewTimer, activeField = '', currentPage = 1, importingNotes = '', busy = false;
let checkpointId = null, checkpointSaving = false, initializing = true, changeRevision = 0;
let templateSaving = false, selectedTemplate = null;
let templateEditSession = null;
let workspaceExpanded = false;
let templatePreviewGeneration = 0, templatePreviewReady = false, templatePreviewGeometry = null;
let templateSheetView = { sheet: '', source: '', row: 0, col: 0 };
const templateSheetRows = 30, templateSheetColumns = 12;
let connectedFields = [];
let selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }, view = { rowStart: 0, colStart: 0, rows: 50, cols: 12 };
let activeLayout = { pages: [], errors: [] }, dragging = false, history = [], redo = [], lastEditKey = '', lastEditAt = 0;
const allEditors = () => [mainEditor, headerEditor, footerEditor].filter(Boolean);
const rangeSelection = () => ({ s: { r: Math.min(selection.s.r, selection.e.r), c: Math.min(selection.s.c, selection.e.c) }, e: { r: Math.max(selection.s.r, selection.e.r), c: Math.max(selection.s.c, selection.e.c) } });
const selectedAddress = () => XLSX.utils.encode_cell(selection.s), selectedRange = () => XLSX.utils.encode_range(rangeSelection());
const status = (text) => { $('#globalStatus').textContent = text; };
function notice(message = '') { $('#notice').textContent = message; $('#notice').classList.toggle('hidden', !message); }
function readableError(failure) {
  const message = failure?.message || String(failure);
  if (failure?.name === 'QuotaExceededError') return 'There is not enough space to save your work here. Use Save file / Print to make a backup file.';
  if (/Formula calculation needs attention|A formula returned an error/i.test(message)) return 'PaperLink could not calculate a spreadsheet formula. Save your changes using Save file / Print → Spreadsheet, check the calculations in Excel, save that file, and open it here again.';
  if (/There is nothing to export|No non-empty records/i.test(message)) return 'There are no pages or filled rows to save yet. Add some information or open a file first.';
  if (/not a supported PaperLink template/i.test(message)) return 'This file is not a template made by PaperLink. Choose a template file saved using Save file / Print.';
  if (/not a supported PaperLink project/i.test(message)) return 'This file is not a work backup made by PaperLink. Choose a backup file saved using Save file / Print.';
  if (/no longer exists|missing source|missing worksheet/i.test(message)) return message.replace(/worksheet/gi, 'sheet') + ' Open the correct spreadsheet, or choose Connected information to select a different sheet.';
  return message.replace(/checkpoint/gi, 'saved work').replace(/workbook/gi, 'spreadsheet').replace(/worksheet/gi, 'sheet').replace(/record documents/gi, 'documents').replace(/document export/gi, 'saving the document');
}
function error(failure) { const message = readableError(failure); status(message); notice(message); }
function spreadsheetHasContent() { const bounds = project.bounds(), first = project.cell(project.settings.sheet, XLSX.utils.encode_cell(bounds.s)); return project.names.spreadsheet !== 'Blank spreadsheet' || XLSX.utils.encode_range(bounds) !== 'A1' || first?.v !== undefined; }
function documentHasContent() { return project.names.document !== 'Blank document' || !mainEditor.isEmpty || !headerEditor.isEmpty || !footerEditor.isEmpty || mainEditor.state.doc.childCount !== 1 || mainEditor.state.doc.firstChild?.type.name !== 'paragraph'; }
function updateSaveState() { $('#saveState').textContent = templateEditSession ? dirty ? 'Template changes not saved' : 'Editing saved template' : dirty ? 'Changes not saved here' : checkpointId ? 'Saved here' : 'Not saved here yet'; $('#saveState').classList.toggle('unsaved', dirty); }
function markDirty() { if (syncing) return; changeRevision++; dirty = true; $('#dirtyDot').classList.remove('hidden'); native?.setDirty(true); updateSaveState(); }
function saved() { dirty = false; $('#dirtyDot').classList.add('hidden'); native?.setDirty(Boolean(templateEditSession?.backup.dirty)); updateSaveState(); }
function commitProjectName() {
  const name = $('#projectName').value.trim() || 'Untitled work';
  if (project.names.project !== name) { project.names.project = name; markDirty(); }
  $('#projectName').value = name;
}
let nameBeforeEdit = project.names.project;
$('#projectName').onfocus = () => { nameBeforeEdit = project.names.project; };
$('#projectName').oninput = () => { project.names.project = $('#projectName').value; markDirty(); };
$('#projectName').onblur = commitProjectName;
$('#projectName').onkeydown = (event) => {
  if (event.key === 'Enter') { event.preventDefault(); event.target.blur(); }
  if (event.key === 'Escape') { event.preventDefault(); event.target.value = nameBeforeEdit; commitProjectName(); event.target.blur(); }
};
function undoPoint(key = '') {
  if (key && key === lastEditKey && Date.now() - lastEditAt < 1200) { lastEditAt = Date.now(); return; }
  history.push(clone(project.workbook)); if (history.length > 25) history.shift(); redo = []; lastEditKey = key; lastEditAt = Date.now();
}
function workbookUndo(forward = false) {
  const source = forward ? redo : history, target = forward ? history : redo; if (!source.length) return;
  target.push(clone(project.workbook)); project.workbook = source.pop(); project.recalculate(); markDirty(); refresh();
}
function findFields() {
  const result = [];
  for (const editor of allEditors()) editor.state.doc.descendants((node, pos) => { if (node.type.name === 'linkedField') result.push({ attrs: node.attrs, pos, editor, node }); });
  return result;
}
function syncFields() {
  syncing = true;
  try { for (const editor of allEditors()) {
    const tr = editor.state.tr; let changed = false;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name !== 'linkedField') return;
      const value = project.fieldValue(node.attrs); if (node.attrs.value !== value) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, value }); changed = true; }
    });
    if (changed) editor.view.dispatch(tr.setMeta('addToHistory', false));
  } } finally { syncing = false; }
}
function editSharedValue(field, value) {
  try { const source = project.source(field); undoPoint(source.sheet + '!' + source.address); project.editField(field, value); markDirty(); syncFields(); updateGridValues(); renderMappingList(); schedulePreview(); updateFormulaStatus(); status('Updated cell ' + source.address + ' and the document.'); }
  catch (failure) { error(failure); syncFields(); }
}
function editorUpdated() { if (syncing) return; markDirty(); renderMappingList(); updateGridValues(); schedulePreview(); }
function makeEditor(element, content) {
  let instance;
  instance = createEditor(element, {
    project: () => project, edit: editSharedValue, select: (id) => selectField(id, false), update: editorUpdated,
    templateEditing: () => Boolean(templateEditSession),
    focus: (editor) => { activeEditor = editor; updateConnectionHint(); }, selection: updateToolbar,
    drop: (data, pos) => { activeEditor = instance; if (pos != null) instance.commands.setTextSelection(pos); selection = data.range; project.settings.sheet = data.sheet; renderGrid(); insertLink(data.range, true).catch(error); }
  }, content); return instance;
}
mainEditor = makeEditor($('#documentEditor'), project.document);
headerEditor = makeEditor($('#headerEditor'), '<p></p>'); footerEditor = makeEditor($('#footerEditor'), '<p></p>'); activeEditor = mainEditor;
// A newly imported logo can finish loading after the first synchronous preview.
for (const editor of allEditors()) editor.view.dom.addEventListener('load', (event) => { if (event.target.tagName === 'IMG') schedulePreview(); }, true);
document.fonts.ready.then(schedulePreview);
function updateToolbar() { if (activeEditor) for (const name of ['bold', 'italic', 'underline', 'bulletList', 'orderedList']) $('[data-command="' + name + '"]')?.classList.toggle('active', activeEditor.isActive(name)); updateConnectionHint(); }
function updateFormulaStatus() { $('#formulaStatus').textContent = project.formulaError ? 'Check a calculation' : 'Ready'; $('#formulaStatus').style.color = project.formulaError ? '#b26c36' : ''; notice([importingNotes, project.formulaError ? readableError(new Error(project.formulaError)) : ''].filter(Boolean).join(' ')); }
function updateConnectionHint() {
  const address = selectedAddress(), editor = activeEditor;
  const selectedText = editor ? editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, ' ').trim() : '';
  const range = rangeSelection(); $('#insertRange').disabled = range.s.r === range.e.r && range.s.c === range.e.c;
  $('#insertRange').title = 'To select a table, hold Shift and click another cell, then choose this button.';
  $('#insertField').textContent = (selectedText ? 'Replace text with ' : 'Add ') + address + (selectedText ? '' : ' to document');
  $('#connectionHint').textContent = selectedText ? 'The highlighted text will be replaced by this cell. Both stay connected.' : 'Click its place in the document, then choose Add to document.';
}
function updateGuidance() {
  $('#editingPaper').classList.toggle('empty-document', mainEditor.isEmpty && mainEditor.state.doc.childCount === 1 && mainEditor.state.doc.firstChild?.type.name === 'paragraph');
  const hasDocument = documentHasContent();
  const hasSpreadsheet = spreadsheetHasContent();
  let step = 1, hint = 'Open a spreadsheet and a document, or choose a saved template.';
  if (templateEditSession) { step = 2; hint = 'Editing “' + templateEditSession.name + '”. Change its document and cell connections, then choose Save changes.'; }
  else if (project.settings.templateId && project.names.spreadsheet === 'Blank spreadsheet' && project.settings.mode === 'batch' && !recordRows().length) hint = 'Your template is ready. Open a spreadsheet or enter new values on the left.';
  else if (connectedFields.length || project.settings.templateId) { step = 3; hint = project.settings.mode === 'batch' ? 'Use the arrows to check each document, then choose Check pages and Save file / Print.' : 'Check the information and page layout, then choose Save file / Print.'; }
  else if (hasSpreadsheet || hasDocument) { step = 2; hint = hasSpreadsheet && !hasDocument ? 'Open a document, choose a template, or start writing on the right.' : 'To add a spreadsheet value, click its cell, click in the document, then choose Add to document.'; }
  for (const item of document.querySelectorAll('.work-steps li')) {
    item.classList.toggle('current', Number(item.dataset.step) === step);
    item.classList.toggle('earlier', Number(item.dataset.step) < step);
    if (Number(item.dataset.step) === step) item.querySelector('button').setAttribute('aria-current', 'step'); else item.querySelector('button').removeAttribute('aria-current');
  }
  if ($('#guidedHint').textContent !== hint) $('#guidedHint').textContent = hint;
  $('#modeHint').textContent = project.settings.mode === 'batch' ? 'One document for each filled row, such as each employee.' : 'Prepare one document using the cells you choose.';
}
function sheetControls() {
  $('#sheetFileName').textContent = project.names.spreadsheet; $('#docFileName').textContent = project.names.document;
  if (templateEditSession) $('#templateEditReference').textContent = project.names.spreadsheet === 'Blank spreadsheet' ? 'Enter preview values in the cells, or choose Open file on the left to try a spreadsheet.' : 'Preview spreadsheet: ' + project.names.spreadsheet + ' (working copy). Values used for preview are not saved in the template.';
  $('#sheetFileName').title = project.names.spreadsheet; $('#docFileName').title = project.names.document;
  if (document.activeElement !== $('#projectName')) $('#projectName').value = project.names.project;
  $('#sheetTabs').innerHTML = project.workbook.SheetNames.map((name) => '<button data-sheet="' + escapeHtml(name) + '" class="' + (name === project.settings.sheet ? 'active' : '') + '">' + escapeHtml(name) + '</button>').join('');
  $('#sheetTabs').querySelectorAll('button').forEach((button) => button.onclick = () => { project.settings.sheet = button.dataset.sheet; selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; view.rowStart = 0; view.colStart = 0; refresh(); markDirty(); });
  const options = project.workbook.SheetNames.map((name) => '<option>' + escapeHtml(name) + '</option>').join('');
  $('#batchSheet').innerHTML = options; $('#batchSheet').value = project.settings.batchSheet; $('#fieldSheet').innerHTML = options;
  $('#singleMode').classList.toggle('active', project.settings.mode === 'single'); $('#batchMode').classList.toggle('active', project.settings.mode === 'batch'); $('#batchControls').classList.toggle('hidden', project.settings.mode !== 'batch');
  $('#singleMode').setAttribute('aria-pressed', String(project.settings.mode === 'single')); $('#batchMode').setAttribute('aria-pressed', String(project.settings.mode === 'batch'));
  $('#headerRow').value = project.settings.headerRow + 1; $('#paper').value = project.settings.paper; $('#orientation').value = project.settings.landscape ? 'landscape' : 'portrait';
  if (![...$('#margin').options].some((o) => Number(o.value) === Number(project.settings.margin))) $('#margin').add(new Option(Math.round(project.settings.margin * 10) / 10 + ' mm margins', String(project.settings.margin)));
  $('#margin').value = String(project.settings.margin); $('#pageNumbers').checked = project.settings.pageNumbers; updateRecord(); updateGuidance();
}
function recordRows() { try { return project.recordRows(); } catch { return []; } }
function updateRecord() {
  if (project.settings.mode !== 'batch') return;
  const rows = recordRows(), index = rows.indexOf(project.settings.activeRow);
  const nameField = connectedFields.find((field) => field.attrs.sheet === project.settings.batchSheet && field.attrs.followRecord !== false && /name/i.test(field.attrs.label));
  let description = '';
  try { description = nameField ? project.fieldValue(nameField.attrs) : project.display(project.settings.batchSheet, XLSX.utils.encode_cell({ r: project.settings.activeRow, c: project.bounds(project.settings.batchSheet).s.c })); } catch {}
  $('#recordLabel').textContent = !rows.length ? 'No documents yet' : index < 0 ? 'This row is empty' : (index + 1) + ' of ' + rows.length + (description ? ' · ' + description : '');
  $('#recordLabel').title = description; $('#recordRowHint').textContent = rows.length ? 'Spreadsheet row ' + (project.settings.activeRow + 1) : 'Add values below the headings';
  $('#prevRecord').disabled = index >= 0 ? index === 0 : !rows.length || rows[0] >= project.settings.activeRow;
  $('#nextRecord').disabled = index >= 0 ? index === rows.length - 1 : !rows.length || rows.at(-1) <= project.settings.activeRow;
}
function setRecord(row) {
  project.settings.activeRow = row;
  const field = findFields().find((item) => item.attrs.id === activeField);
  if (field) {
    try { const source = project.source(field.attrs); if (source.sheet === project.settings.sheet) selection = { s: parseCell(source.address), e: parseCell(source.address) }; } catch {}
  } else if (project.settings.sheet === project.settings.batchSheet && selection.s.r > project.settings.headerRow) selection = { s: { r: row, c: selection.s.c }, e: { r: row, c: selection.s.c } };
  syncFields(); ensureVisible(selection.s); renderGrid(); renderMappingList(); updateRecord(); schedulePreview(); markDirty();
}
function setMode(mode) {
  project.settings.mode = mode; if (mode === 'batch') project.settings.activeRow = recordRows()[0] ?? project.settings.headerRow + 1;
  sheetControls(); syncFields(); renderGrid(); renderMappingList(); schedulePreview(); markDirty();
  status(mode === 'batch' ? 'Each filled row makes a document. Use the arrows to check each one.' : 'You are preparing one document. Choose the cells you want to include.');
}
function sourceCellSelected(coordinate, extend = false) { if (extend) selection.e = coordinate; else selection = { s: coordinate, e: coordinate }; ensureVisible(coordinate); paintSelection(); updateFormulaBar(); }
function ensureVisible(cell) {
  let changed = false;
  if (cell.r < view.rowStart || cell.r >= view.rowStart + view.rows) { view.rowStart = Math.max(0, cell.r - 5); changed = true; }
  if (cell.c < view.colStart || cell.c >= view.colStart + view.cols) { view.colStart = Math.max(0, cell.c - 3); changed = true; }
  if (changed) renderGrid();
}
function fieldSources() { const result = new Set(); for (const field of findFields()) try { const source = project.source(field.attrs); if (source.sheet === project.settings.sheet) result.add(source.address); } catch {} return result; }
function renderGrid() {
  const sheet = project.sheet(), linked = fieldSources(), range = rangeSelection(), scroll = $('#gridScroll'), top = scroll.scrollTop, left = scroll.scrollLeft;
  const merges = (sheet['!merges'] || []).filter((m) => m.e.r >= view.rowStart && m.s.r < view.rowStart + view.rows && m.e.c >= view.colStart && m.s.c < view.colStart + view.cols);
  let html = '<thead><tr><th class="corner"></th>';
  for (let c = view.colStart; c < Math.min(16384, view.colStart + view.cols); c++) html += '<th data-col="' + c + '">' + colName(c) + '</th>';
  html += '</tr></thead><tbody>';
  for (let r = view.rowStart; r < Math.min(1048576, view.rowStart + view.rows); r++) {
    html += '<tr><th data-row="' + r + '">' + (r + 1) + '</th>';
    for (let c = view.colStart; c < Math.min(16384, view.colStart + view.cols); c++) {
      const merge = merges.find((m) => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
      if (merge && (r !== Math.max(merge.s.r, view.rowStart) || c !== Math.max(merge.s.c, view.colStart))) continue;
      const actual = merge ? merge.s : { r, c }, address = XLSX.utils.encode_cell(actual), cell = sheet[address], classes = [];
      if (cell?.t === 'n') classes.push('numeric'); if (cell?.f) classes.push('formula-cell'); if (linked.has(address)) classes.push('linked-cell');
      if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) classes.push('selected');
      if (actual.r === selection.s.r && actual.c === selection.s.c) classes.push('active-cell');
      const spans = merge ? ' colspan="' + (Math.min(merge.e.c + 1, view.colStart + view.cols) - c) + '" rowspan="' + (Math.min(merge.e.r + 1, view.rowStart + view.rows) - r) + '"' : '';
      html += '<td draggable="true" data-address="' + address + '" data-row="' + actual.r + '" data-col="' + actual.c + '" class="' + classes.join(' ') + '"' + spans + ' title="' + escapeHtml(address + (cell?.f ? ' · =' + cell.f : '')) + '"><span class="cell-label">' + escapeHtml(project.display(project.settings.sheet, address)) + '</span></td>';
    }
    html += '</tr>';
  }
  $('#sheetGrid').innerHTML = html; scroll.scrollTop = top; scroll.scrollLeft = left; updateFormulaBar();
}
function paintSelection() {
  const range = rangeSelection();
  $('#sheetGrid').querySelectorAll('td').forEach((td) => { const r = Number(td.dataset.row), c = Number(td.dataset.col); td.classList.toggle('selected', r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c); td.classList.toggle('active-cell', td.dataset.address === selectedAddress()); });
  $('#selectionLabel').textContent = selectedRange() + ' selected'; updateConnectionHint();
}
function updateFormulaBar() { $('#cellAddress').value = selectedAddress(); if (document.activeElement !== $('#formulaInput')) $('#formulaInput').value = project.raw(project.settings.sheet, selectedAddress()); paintSelection(); }
function updateGridValues() {
  const linked = fieldSources();
  $('#sheetGrid').querySelectorAll('td').forEach((td) => { const cell = project.cell(project.settings.sheet, td.dataset.address); td.querySelector('.cell-label').textContent = project.display(project.settings.sheet, td.dataset.address); td.classList.toggle('numeric', cell?.t === 'n'); td.classList.toggle('formula-cell', Boolean(cell?.f)); td.classList.toggle('linked-cell', linked.has(td.dataset.address)); });
  updateFormulaBar(); updateRecord();
}
function applyCell(address, value, historyKey = '') { undoPoint(historyKey); project.setCell(project.settings.sheet, address, value); markDirty(); syncFields(); updateGridValues(); renderMappingList(); schedulePreview(); updateFormulaStatus(); }
function editGridCell(td, initial) {
  if (!td || td.querySelector('input')) return;
  const address = td.dataset.address, input = document.createElement('input'), original = project.raw(project.settings.sheet, address);
  input.value = initial ?? original; input.setAttribute('aria-label', address + ' value'); td.append(input); input.focus(); if (initial == null) input.select();
  input.oninput = () => { try { applyCell(address, input.value, project.settings.sheet + '!' + address); } catch (failure) { error(failure); } };
  input.onkeydown = (event) => {
    event.stopPropagation();
    if (event.key === 'Escape') { try { applyCell(address, original); } catch (failure) { error(failure); } input.blur(); }
    if (['Enter', 'Tab'].includes(event.key)) { event.preventDefault(); input.blur(); const cell = parseCell(address); sourceCellSelected({ r: Math.max(0, cell.r + (event.key === 'Enter' ? event.shiftKey ? -1 : 1 : 0)), c: Math.max(0, cell.c + (event.key === 'Tab' ? event.shiftKey ? -1 : 1 : 0)) }); $('#gridScroll').focus(); }
  };
  input.onblur = () => { input.remove(); updateGridValues(); };
}
$('#sheetGrid').addEventListener('pointerdown', (event) => {
  if (event.target.closest('input')) return; const td = event.target.closest('td');
  if (td) { sourceCellSelected({ r: Number(td.dataset.row), c: Number(td.dataset.col) }, event.shiftKey); dragging = true; $('#gridScroll').focus(); }
  const row = event.target.closest('th[data-row]'), col = event.target.closest('th[data-col]');
  if (row) { const r = Number(row.dataset.row); selection = { s: { r, c: 0 }, e: { r, c: project.bounds().e.c } }; updateFormulaBar(); }
  if (col) { const c = Number(col.dataset.col); selection = { s: { r: 0, c }, e: { r: project.bounds().e.r, c } }; updateFormulaBar(); }
});
$('#sheetGrid').addEventListener('pointerover', (event) => { const td = event.target.closest('td'); if (dragging && td && event.buttons === 1) { selection.e = { r: Number(td.dataset.row), c: Number(td.dataset.col) }; paintSelection(); } });
document.addEventListener('pointerup', () => { dragging = false; });
$('#sheetGrid').addEventListener('dblclick', (event) => editGridCell(event.target.closest('td')));
$('#sheetGrid').addEventListener('dragstart', (event) => {
  if (event.target.closest('input')) { event.preventDefault(); return; }
  dragging = false;
  if (!event.dataTransfer) return;
  const range = rangeSelection(), address = XLSX.utils.encode_cell(range.s), multiple = range.s.r !== range.e.r || range.s.c !== range.e.c;
  event.dataTransfer.setData('application/x-paperlink-cells', JSON.stringify({ sheet: project.settings.sheet, range })); event.dataTransfer.effectAllowed = 'copy';
  startDragFeedback(event, multiple ? 'Add cells ' + XLSX.utils.encode_range(range) : 'Add ' + address + ': ' + (project.display(project.settings.sheet, address) || '(empty cell)'));
});
function bulkValues(lines, clear = false) {
  undoPoint(); const range = rangeSelection();
  if (clear && (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) > 10000) throw new Error('Select at most 10,000 cells for a bulk edit.');
  if (clear) for (let r = range.s.r; r <= range.e.r; r++) for (let c = range.s.c; c <= range.e.c; c++) project.setCell(project.settings.sheet, XLSX.utils.encode_cell({ r, c }), '', { recalculate: false });
  else lines.forEach((line, r) => line.split('\t').forEach((value, c) => project.setCell(project.settings.sheet, XLSX.utils.encode_cell({ r: selection.s.r + r, c: selection.s.c + c }), value, { recalculate: false })));
  project.recalculate(); markDirty(); refresh();
}
$('#gridScroll').addEventListener('keydown', (event) => {
  if (event.target.tagName === 'INPUT') return;
  if ((event.ctrlKey || event.metaKey) && ['z', 'y'].includes(event.key.toLowerCase())) { event.preventDefault(); workbookUndo(event.key.toLowerCase() === 'y' || event.shiftKey); return; }
  const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, event.shiftKey ? -1 : 1] };
  if (moves[event.key]) { event.preventDefault(); const [r, c] = moves[event.key]; sourceCellSelected({ r: Math.max(0, Math.min(1048575, selection.s.r + r)), c: Math.max(0, Math.min(16383, selection.s.c + c)) }, event.shiftKey && event.key !== 'Tab'); return; }
  if (['Delete', 'Backspace'].includes(event.key)) { event.preventDefault(); try { bulkValues([], true); } catch (failure) { error(failure); } return; }
  if (event.key === 'F2' || event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) { event.preventDefault(); editGridCell($('#sheetGrid td[data-address="' + selectedAddress() + '"]'), event.key === 'F2' ? undefined : event.key); }
});
$('#gridScroll').addEventListener('copy', (event) => {
  if (event.target.tagName === 'INPUT') return;
  const range = rangeSelection(), rows = [];
  if ((range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) > 100000) { error(new Error('Select fewer cells to copy.')); return; }
  for (let r = range.s.r; r <= range.e.r; r++) { const cells = []; for (let c = range.s.c; c <= range.e.c; c++) cells.push(project.display(project.settings.sheet, XLSX.utils.encode_cell({ r, c }))); rows.push(cells.join('\t')); }
  event.preventDefault(); event.clipboardData.setData('text/plain', rows.join('\n'));
});
$('#gridScroll').addEventListener('paste', (event) => {
  if (event.target.tagName === 'INPUT') return;
  event.preventDefault(); const lines = event.clipboardData.getData('text/plain').replace(/\r/g, '').replace(/\n$/, '').split('\n');
  if (lines.length * Math.max(...lines.map((line) => line.split('\t').length)) > 10000) { error(new Error('Paste at most 10,000 cells at a time.')); return; }
  try { bulkValues(lines); } catch (failure) { error(failure); }
});
$('#formulaInput').onkeydown = (event) => { if (event.key === 'Enter') try { applyCell(selectedAddress(), event.target.value); } catch (failure) { error(failure); } };
$('#applyFormula').onclick = () => { try { applyCell(selectedAddress(), $('#formulaInput').value); } catch (failure) { error(failure); } };
$('#cellAddress').onkeydown = (event) => { if (event.key === 'Enter') try { sourceCellSelected(parseCell(event.target.value)); $('#gridScroll').focus(); } catch (failure) { error(failure); } };
$('#moreRows').onclick = () => { if (view.rows < 500) view.rows += 50; else view.rowStart = Math.min(1048076, view.rowStart + 500); renderGrid(); status('Type a cell name, such as A100, in the box above the spreadsheet to find it.'); };
$('#addSheet').onclick = () => { const name = window.prompt('What would you like to call the new sheet?', 'Sheet' + (project.workbook.SheetNames.length + 1)); if (name) try { project.addSheet(name); markDirty(); refresh(); } catch (failure) { error(failure); } };

function newField(sheet, address, follows = true) {
  const cell = parseCell(address), label = project.display(sheet, XLSX.utils.encode_cell({ r: project.settings.headerRow, c: cell.c })) || sheet + '!' + address;
  const attrs = { id: crypto.randomUUID(), sheet, address, label, followRecord: follows, value: '' }; attrs.value = project.fieldValue(attrs); return { type: 'linkedField', attrs };
}
async function chooseAnchor(sheetValue, text) {
  const dialog = $('#mappingDialog'); $('#anchorSheetValue').textContent = sheetValue || '(empty cell)'; $('#anchorDocumentValue').textContent = text; dialog.querySelector('[value="sheet"]').checked = true; dialog.showModal();
  return new Promise((resolve) => dialog.addEventListener('close', () => resolve(dialog.returnValue === 'connect' ? dialog.querySelector('[name="anchor"]:checked').value : null), { once: true }));
}
async function insertLink(range = rangeSelection(), forceTable = false) {
  if (busy) return;
  const editor = activeEditor || mainEditor, sheet = project.settings.sheet, rows = range.e.r - range.s.r + 1, cols = range.e.c - range.s.c + 1;
  if (rows * cols > 1000) throw new Error('Choose a smaller table. You can add up to 1,000 cells at once.');
  if (forceTable && rows * cols > 1) {
    const table = { type: 'table', content: [] };
    for (let r = range.s.r; r <= range.e.r; r++) { const row = { type: 'tableRow', content: [] }; for (let c = range.s.c; c <= range.e.c; c++) row.content.push({ type: 'tableCell', content: [{ type: 'paragraph', content: [newField(sheet, XLSX.utils.encode_cell({ r, c }), rows === 1)] }] }); table.content.push(row); }
    editor.chain().focus().insertContent([table, { type: 'paragraph' }]).run(); status(rows > 1 ? 'Table added. Each document will contain this same table.' : 'Table added. Its values stay connected to the spreadsheet.');
  } else {
    const field = newField(sheet, XLSX.utils.encode_cell(range.s)), { from, to } = editor.state.selection, text = editor.state.doc.textBetween(from, to, ' '), source = project.source(field.attrs), cell = project.cell(source.sheet, source.address), value = project.display(source.sheet, source.address);
    if (text.trim() && text !== value) {
      const anchor = await chooseAnchor(value, text); if (!anchor) return;
      if (anchor === 'document') { if (cell?.f) throw new Error('This cell contains a formula. Use the spreadsheet value or choose an input cell.'); undoPoint(); project.setCell(source.sheet, source.address, text); field.attrs.value = project.fieldValue(field.attrs); }
      editor.commands.setTextSelection({ from, to });
    }
    editor.chain().focus().insertContent(field).run(); selectField(field.attrs.id, false); status('Added ' + source.address + ' to the document. Change the green value here or in the spreadsheet to update both.');
  }
  markDirty(); syncFields(); updateGridValues(); renderMappingList(); schedulePreview();
}
$('#insertField').onclick = () => insertLink().catch(error); $('#insertRange').onclick = () => insertLink(rangeSelection(), true).catch(error);
function selectField(id, open = true) {
  activeField = id; if (open) $('#mappingDrawer').classList.remove('hidden');
  const field = findFields().find((f) => f.attrs.id === id); if (!field) return; activeEditor = field.editor;
  try { const source = project.source(field.attrs); if (project.settings.sheet !== source.sheet) { project.settings.sheet = source.sheet; sheetControls(); renderGrid(); } sourceCellSelected(parseCell(source.address)); } catch {}
  renderMappingList();
}
function renderMappingList() {
  const fields = [...new Map(findFields().map((field) => [field.attrs.id, field])).values()]; connectedFields = fields; $('#mappingCount').textContent = fields.length; updateRecord(); updateGuidance();
  $('#mappingList').innerHTML = fields.length ? fields.map(({ attrs }) => {
    let source; try { source = project.source(attrs); } catch { source = attrs; }
    const changing = project.settings.mode === 'batch' && attrs.followRecord !== false && attrs.sheet === project.settings.batchSheet;
    return '<div class="mapping-card ' + (activeField === attrs.id ? 'active' : '') + '" data-id="' + escapeHtml(attrs.id) + '"><b>' + escapeHtml(attrs.label) + '</b><small>From ' + escapeHtml(source.sheet + ', cell ' + source.address) + (changing ? ' · changes for each document' : ' · uses this cell') + '</small></div>';
  }).join('') : '<p class="small-note">Nothing is connected yet. Click a spreadsheet cell, click in the document, then choose Add to document.</p>';
  $('#mappingList').querySelectorAll('[data-id]').forEach((card) => card.onclick = () => selectField(card.dataset.id));
  const field = fields.find((f) => f.attrs.id === activeField); $('#fieldInspector').classList.toggle('hidden', !field);
  document.querySelectorAll('.mapped-field').forEach((el) => el.classList.toggle('is-selected', el.dataset.fieldId === activeField));
  if (!field) return; const attrs = field.attrs;
  for (const [id, value] of [['#fieldLabel', attrs.label], ['#fieldAddress', attrs.address]]) if (document.activeElement !== $(id)) $(id).value = value;
  $('#fieldSheet').value = attrs.sheet; $('#followRecord').checked = attrs.followRecord !== false;
  try { const source = project.source(attrs), cell = project.cell(source.sheet, source.address); if (document.activeElement !== $('#fieldValue')) $('#fieldValue').value = String(cell?.v ?? ''); $('#fieldValue').readOnly = Boolean(cell?.f); $('#fieldFormulaNote').textContent = cell?.f ? 'This value is calculated in the spreadsheet. Change the numbers used in its formula.' : 'Changes here also update ' + source.sheet + ', cell ' + source.address + '.'; }
  catch { $('#fieldValue').value = ''; $('#fieldFormulaNote').textContent = 'This sheet is not in your spreadsheet. Open the correct file or choose another sheet under Change the connection.'; }
}
function updateFieldAttributes(patch) {
  for (const field of findFields().filter((f) => f.attrs.id === activeField)) field.editor.view.dispatch(field.editor.state.tr.setNodeMarkup(field.pos, undefined, { ...field.attrs, ...patch }));
  markDirty(); syncFields(); renderMappingList(); updateGridValues(); schedulePreview();
}
$('#fieldLabel').oninput = () => updateFieldAttributes({ label: $('#fieldLabel').value });
$('#fieldSheet').onchange = () => updateFieldAttributes({ sheet: $('#fieldSheet').value });
$('#fieldAddress').onchange = () => { try { parseCell($('#fieldAddress').value); updateFieldAttributes({ address: $('#fieldAddress').value.toUpperCase() }); } catch (failure) { error(failure); renderMappingList(); } };
$('#followRecord').onchange = () => updateFieldAttributes({ followRecord: $('#followRecord').checked });
$('#fieldValue').oninput = () => { const field = findFields().find((f) => f.attrs.id === activeField); if (field) editSharedValue(field.attrs, $('#fieldValue').value); };
$('#rebindField').onclick = () => updateFieldAttributes({ sheet: project.settings.sheet, address: selectedAddress() });
$('#unlinkField').onclick = () => {
  for (const editor of allEditors()) {
    const matches = findFields().filter((f) => f.editor === editor && f.attrs.id === activeField).sort((a, b) => b.pos - a.pos), tr = editor.state.tr;
    for (const field of matches) { const value = project.fieldValue(field.attrs); if (value) tr.replaceWith(field.pos, field.pos + field.node.nodeSize, editor.schema.text(value)); else tr.delete(field.pos, field.pos + field.node.nodeSize); }
    if (matches.length) editor.view.dispatch(tr);
  }
  activeField = ''; markDirty(); renderMappingList(); updateGridValues(); schedulePreview();
};
$('#mappingDrawerButton').onclick = () => { $('#mappingDrawer').classList.toggle('hidden'); renderMappingList(); }; $('#closeDrawer').onclick = () => $('#mappingDrawer').classList.add('hidden');
function layoutForRow(row) {
  if (templateEditSession) {
    const map = Object.fromEntries(project.workbook.SheetNames.map((name) => [name, name]));
    const parts = allEditors().map((editor) => templateDocumentHtml(editor.getJSON(), project, map, true, true));
    return paginate(parts[0], project.settings, $('#measurementHost'), parts[1], parts[2]);
  }
  return paginate(resolveHtml(mainEditor.getHTML(), project, row), project.settings, $('#measurementHost'), resolveHtml(headerEditor.getHTML(), project, row), resolveHtml(footerEditor.getHTML(), project, row));
}
function schedulePreview() { clearTimeout(previewTimer); previewTimer = setTimeout(renderPreview, 180); }
function fitDocumentWorkspace() {
  const scroll = $('#documentScroll'), padding = getComputedStyle(scroll);
  const innerWidth = Math.max(1, scroll.clientWidth - (parseFloat(padding.paddingLeft) || 0) - (parseFloat(padding.paddingRight) || 0));
  const g = geometry(project.settings), zoom = Math.min(1, innerWidth / g.widthPx);
  Object.assign($('#editingPaper').style, { width: g.widthPx + 'px', minHeight: g.heightPx + 'px', fontFamily: project.settings.fontFamily, fontSize: project.settings.fontSize + 'pt', zoom: String(Math.max(.2, zoom)) });
  $('#pagePreview').style.zoom = String(Math.max(.2, zoom));
  Object.assign(mainEditor.view.dom.style, { padding: g.marginPx + 'px', minHeight: g.heightPx + 'px' });
}
function renderPreview() {
  try {
    activeLayout = layoutForRow(project.settings.activeRow); $('#pagePreview').replaceChildren(...activeLayout.pages); $('#pageCount').textContent = activeLayout.pages.length; currentPage = Math.min(currentPage, activeLayout.pages.length);
    $('#docStatus').textContent = activeLayout.errors.length ? readableError(new Error(activeLayout.errors[0])) : activeLayout.pages.length + ' page' + (activeLayout.pages.length === 1 ? '' : 's') + ' · ' + project.settings.paper;
    $('#docStatus').title = $('#docStatus').textContent;
    $('#docStatus').style.color = activeLayout.errors.length ? '#a56b3c' : '';
    fitDocumentWorkspace(); updatePreviewPage();
  } catch (failure) { error(failure); }
}
function updatePreviewPage() { $('#pagePreview').querySelectorAll('.document-page').forEach((page) => page.classList.toggle('current-page', Number(page.dataset.page) === currentPage)); }
$('#pagePreview').onclick = (event) => { const page = event.target.closest('.document-page'); if (page) { currentPage = Number(page.dataset.page); updatePreviewPage(); } const field = event.target.closest('[data-field-id]'); if (field) selectField(field.dataset.fieldId); };
$('#editorTab').onclick = () => { $('#editingPaper').classList.remove('hidden'); $('#pagePreview').classList.add('hidden'); $('#editorToolbar').classList.remove('hidden'); $('#editorTab').classList.add('active'); $('#previewTab').classList.remove('active'); };
$('#previewTab').onclick = () => { renderPreview(); $('#editingPaper').classList.add('hidden'); $('#pagePreview').classList.remove('hidden'); $('#editorToolbar').classList.add('hidden'); $('#editorTab').classList.remove('active'); $('#previewTab').classList.add('active'); };
$('#headerFooterButton').onclick = () => { if ($('#headerFooterPanel').classList.toggle('hidden')) { activeEditor = mainEditor; mainEditor.commands.focus(); } };
$('#moreTools').onclick = () => { const hidden = $('#advancedEditorTools').classList.toggle('hidden'); $('#moreTools').setAttribute('aria-expanded', String(!hidden)); $('#moreTools').textContent = hidden ? 'More tools' : 'Fewer tools'; };
$('#editorToolbar').onmousedown = (event) => { if (event.target.closest('button')) event.preventDefault(); };
$('#editorToolbar').querySelectorAll('[data-command]').forEach((button) => button.onclick = () => {
  const chain = activeEditor.chain().focus(), name = button.dataset.command;
  if (name === 'table') chain.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
  else if (name === 'row') chain.addRowAfter().run(); else if (name === 'column') chain.addColumnAfter().run();
  else if (name === 'pageBreak') chain.insertContent({ type: 'pageBreak' }).run();
  else chain['toggle' + name[0].toUpperCase() + name.slice(1)]().run(); updateToolbar();
});
$('#textStyle').onchange = () => { const level = Number($('#textStyle').value), chain = activeEditor.chain().focus(); if (level) chain.setHeading({ level }).run(); else chain.setParagraph().run(); };
$('#fontFamily').onchange = () => activeEditor.chain().focus().setFontFamily($('#fontFamily').value).run();
$('#fontSize').onchange = () => activeEditor.chain().focus().setFontSize($('#fontSize').value + 'pt').run();
$('#alignment').onchange = () => activeEditor.chain().focus().setTextAlign($('#alignment').value).run();
for (const selector of ['#paper', '#orientation', '#margin', '#pageNumbers']) $(selector).onchange = () => { project.settings.paper = $('#paper').value; project.settings.landscape = $('#orientation').value === 'landscape'; project.settings.margin = Number($('#margin').value); project.settings.pageNumbers = $('#pageNumbers').checked; markDirty(); schedulePreview(); };
$('#singleMode').onclick = () => setMode('single'); $('#batchMode').onclick = () => setMode('batch');
$('#batchSheet').onchange = () => { project.settings.batchSheet = $('#batchSheet').value; setMode('batch'); };
$('#headerRow').onchange = () => { const row = Number($('#headerRow').value); if (!Number.isInteger(row) || row < 1 || row > 1048575) { $('#headerRow').value = project.settings.headerRow + 1; error(new Error('Enter the row number containing your headings, such as 1.')); return; } project.settings.headerRow = row - 1; setMode(project.settings.mode); };
$('#prevRecord').onclick = () => { const row = recordRows().findLast((r) => r < project.settings.activeRow); if (row !== undefined) setRecord(row); };
$('#nextRecord').onclick = () => { const row = recordRows().find((r) => r > project.settings.activeRow); if (row !== undefined) setRecord(row); };
setupPaneSplit($('.workspace'), $('#workspaceDivider'), fitDocumentWorkspace);
function setWorkspaceExpanded(expanded) {
  workspaceExpanded = expanded; document.body.classList.toggle('workspace-expanded', expanded);
  $('#workspaceExpand').setAttribute('aria-pressed', String(expanded));
  $('#workspaceExpandLabel').textContent = expanded ? 'Restore layout' : 'Expand workspace';
  $('#workspaceExpandIcon').setAttribute('d', expanded ? 'M3 8h5V3M21 8h-5V3M16 21v-5h5M8 21v-5H3' : 'M8 3H3v5M16 3h5v5M21 16v5h-5M8 21H3v-5');
  $('#workspaceExpand').title = expanded ? 'Restore the top actions and guidance for saving, templates, and help. You can also press Escape.' : 'Give both panes more room by hiding the top actions and guidance';
  requestAnimationFrame(fitDocumentWorkspace);
  status(expanded ? 'Workspace expanded. Choose Restore layout to bring back the top actions, or press Escape.' : 'Full layout restored.');
}
$('#workspaceExpand').onclick = () => setWorkspaceExpanded(!workspaceExpanded);
document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape' || !workspaceExpanded || event.defaultPrevented || event.isComposing || document.querySelector('dialog[open]')) return;
  const menus = document.querySelectorAll('.batch-setup[open], .page-settings[open]');
  if (menus.length) { for (const menu of menus) menu.open = false; event.preventDefault(); return; }
  event.preventDefault(); setWorkspaceExpanded(false); $('#workspaceExpand').focus({ preventScroll: true });
});
function openStartChoices() {
  if (templateEditSession) { status('Use Open file above either pane while editing. Choose Save changes or Cancel editing to return to your work.'); return; }
  const hasContent = documentHasContent() || spreadsheetHasContent() || connectedFields.length > 0;
  $('#startHeading').textContent = hasContent ? 'Open your files' : 'How would you like to start?'; $('#startBlank').textContent = hasContent ? 'Keep working' : 'Start with blank pages';
  if (!$('#startDialog').open) { $('#startDialog').showModal(); $('#startDialog [data-start="import-sheet"]').focus(); }
}
document.querySelectorAll('[data-start]').forEach((button) => button.onclick = () => { $('#startDialog').close(); action(button.dataset.start); });
$('#guideFiles').onclick = openStartChoices;
$('#guideConnect').onclick = () => { $('#editorTab').click(); activeEditor = mainEditor; mainEditor.commands.focus(); status('Click a spreadsheet value, click where it belongs in the document, then choose Add to document.'); };
$('#guideReview').onclick = () => { $('#previewTab').click(); status('Check the pages and information. When you are ready, choose Save file / Print.'); };
document.addEventListener('pointerdown', (event) => { for (const menu of document.querySelectorAll('.batch-setup[open], .page-settings[open]')) if (!menu.contains(event.target)) menu.open = false; });

async function openFile(kind) {
  if (native) return native.openFile(kind);
  const input = $('#browserFile'); input.accept = { spreadsheet: '.xlsx,.xls,.xlsm,.ods,.csv,.tsv', document: '.docx,.odt,.rtf,.html,.htm,.txt', project: '.paperlink,.flowdesk', template: '.papertemplate,.flowtemplate' }[kind]; input.value = '';
  return new Promise((resolve) => { input.onchange = async () => { const file = input.files[0]; resolve(file ? { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) } : null); }; input.oncancel = () => resolve(null); input.click(); });
}
async function saveBytes(bytes, name, extension) {
  if (native) return native.saveFile({ bytes, name, extensions: [extension] });
  const blob = new Blob([bytes]), url = URL.createObjectURL(blob), link = document.createElement('a'); link.href = url; link.download = name; link.click(); setTimeout(() => URL.revokeObjectURL(url), 2000); return { name };
}
function documentState() { project.settings.headerDocument = headerEditor.getJSON(); project.settings.footerDocument = footerEditor.getJSON(); return mainEditor.getJSON(); }
function setEditorContent(content) {
  syncing = true; try { mainEditor.commands.setContent(content); headerEditor.commands.setContent(project.settings.headerDocument || '<p></p>'); footerEditor.commands.setContent(project.settings.footerDocument || '<p></p>'); } finally { syncing = false; } activeEditor = mainEditor;
}
function resumeProject(serialized, id) {
  const loaded = Project.deserialize(serialized);
  project = loaded; checkpointId = id; importingNotes = ''; activeField = ''; history = []; redo = [];
  selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; view.rowStart = 0; view.colStart = 0; currentPage = 1;
  setEditorContent(project.document); saved(); refresh();
}
const canReplaceProject = () => !dirty || window.confirm('Your latest changes are not saved here. Continue without keeping these changes?');
async function checkpointProject() {
  commitProjectName();
  const id = checkpointId || crypto.randomUUID(), revision = changeRevision;
  const serialized = project.serialize(documentState()), name = project.names.project;
  checkpointSaving = true; const button = $('[data-action="save-project"]'); button.disabled = true;
  try {
    await saveCheckpoint(id, name, serialized); checkpointId = id;
    if (revision === changeRevision) { saved(); status('Your work is saved here. Choose Saved work to continue later.'); }
    else status('Your work was saved, but the changes you just made still need saving. Choose Save work again.');
  } finally { checkpointSaving = false; button.disabled = false; }
}
async function openCheckpoints() {
  const dialog = $('#checkpointDialog'); checkpointSaving = true;
  $('#checkpointError').classList.add('hidden'); $('#checkpointStatus').textContent = ''; $('#checkpointList').textContent = 'Opening your saved work…'; dialog.showModal();
  try { await renderCheckpoints(); }
  catch (failure) { $('#checkpointList').textContent = ''; checkpointError(failure); }
  finally { checkpointSaving = false; }
}
function checkpointError(failure) { $('#checkpointError').textContent = readableError(failure); $('#checkpointError').classList.remove('hidden'); }
async function renderCheckpoints() {
  const checkpoints = await listCheckpoints(), list = $('#checkpointList'); list.replaceChildren();
  if (!checkpoints.length) list.textContent = 'No saved work here. Choose Save work on the main screen to keep your progress.';
  for (const checkpoint of checkpoints) {
    const card = document.createElement('div'), description = document.createElement('div'), actions = document.createElement('div');
    card.className = 'checkpoint-card'; actions.className = 'checkpoint-card-actions';
    const title = document.createElement('b'), detail = document.createElement('span');
    title.textContent = checkpoint.name; detail.textContent = 'Saved ' + new Date(checkpoint.updatedAt).toLocaleString() + (checkpoint.id === checkpointId ? ' · Open now' : '');
    const open = document.createElement('button'), remove = document.createElement('button'); open.type = remove.type = 'button';
    open.textContent = 'Open'; open.setAttribute('aria-label', 'Open saved work: ' + checkpoint.name);
    remove.textContent = 'Delete'; remove.className = 'delete-saved-work'; remove.setAttribute('aria-label', 'Delete saved work: ' + checkpoint.name);
    description.append(title, detail); actions.append(open, remove); card.append(description, actions); list.append(card);
    open.onclick = async () => {
      if (checkpointSaving || !canReplaceProject()) return; checkpointSaving = true; open.disabled = remove.disabled = true;
      try {
        const loaded = await getCheckpoint(checkpoint.id);
        if (!$('#checkpointDialog').open) return;
        // Validate before selecting it as the work to restore next launch.
        Project.deserialize(loaded.serialized); await selectCheckpoint(loaded.id);
        if (!$('#checkpointDialog').open) return;
        resumeProject(loaded.serialized, loaded.id); $('#checkpointDialog').close(); status('Your saved work is open. Your document and spreadsheet are ready to use.');
      } catch (failure) { checkpointError(failure); }
      finally { checkpointSaving = false; open.disabled = remove.disabled = false; }
    };
    remove.onclick = () => deleteSavedWork(checkpoint);
  }
}
async function deleteSavedWork(checkpoint) {
  if (checkpointSaving) return;
  if (!window.confirm('Delete the saved copy of “' + checkpoint.name + '”? It will be removed from Saved work. Your currently open spreadsheet and document, templates, and exported files will stay.')) return;
  checkpointSaving = true; $('#checkpointError').classList.add('hidden'); $('#checkpointStatus').textContent = 'Deleting saved copy…';
  try {
    await deleteCheckpoint(checkpoint.id);
    const openCopy = checkpointId === checkpoint.id;
    if (openCopy) { checkpointId = null; markDirty(); }
    await renderCheckpoints();
    const message = 'Saved copy deleted: “' + checkpoint.name + '”.' + (openCopy ? ' Your work is still open. Choose Save work to keep a new saved copy.' : ' Your currently open work is kept.');
    $('#checkpointStatus').textContent = message; status(message);
    ($('#checkpointList').querySelector('button') || $('#importProject')).focus();
  } catch (failure) { $('#checkpointStatus').textContent = ''; checkpointError(failure); }
  finally { checkpointSaving = false; }
}
$('#importProject').onclick = async () => {
  if (checkpointSaving || !canReplaceProject()) return; checkpointSaving = true;
  try {
    const file = await openFile('project'); if (!file) return;
    resumeProject(new TextDecoder().decode(file.bytes), null); markDirty(); $('#checkpointDialog').close();
    status('Your work backup is open. Choose Save work to keep it here in the app.');
  } catch (failure) { checkpointError(failure); }
  finally { checkpointSaving = false; }
};
function templateError(failure, applying = false) {
  const element = $(applying ? '#templateApplyError' : '#templateError'); element.textContent = readableError(failure); element.classList.remove('hidden');
}
async function renderTemplateLibrary() {
  const templates = await listTemplates(), list = $('#templateList'); list.replaceChildren();
  if (!templates.length) list.textContent = 'No saved templates yet. Prepare a document in the right pane, then choose Save template on the main screen.';
  for (const summary of templates) {
    const card = document.createElement('div'); card.className = 'template-card';
    const description = document.createElement('div'), title = document.createElement('b'), detail = document.createElement('span'), button = document.createElement('button');
    const actions = document.createElement('div'), edit = document.createElement('button'), remove = document.createElement('button'); actions.className = 'template-card-actions';
    title.textContent = summary.name; detail.textContent = summary.fieldCount + ' connected items · ' + (summary.mode === 'batch' ? 'One per row' : 'One document') + ' · ' + new Date(summary.updatedAt).toLocaleDateString();
    button.type = 'button'; button.textContent = 'Preview document'; button.setAttribute('aria-label', 'Preview document: ' + summary.name);
    button.onclick = async () => {
      if (templateSaving) return; templateSaving = true; button.disabled = true;
      try { const record = await getTemplate(summary.id); openTemplateApplication(record.id, deserializeTemplate(record.serialized)); }
      catch (failure) { templateError(failure); } finally { templateSaving = false; button.disabled = false; }
    };
    remove.type = 'button'; remove.textContent = 'Delete'; remove.className = 'delete-template'; remove.setAttribute('aria-label', 'Delete template: ' + summary.name);
    remove.onclick = () => deleteSavedTemplate(summary);
    edit.type = 'button'; edit.textContent = 'Edit'; edit.setAttribute('aria-label', 'Edit template: ' + summary.name); edit.onclick = () => editSavedTemplate(summary.id);
    actions.append(button, edit, remove); description.append(title, detail); card.append(description, actions); list.append(card);
  }
}
async function deleteSavedTemplate(summary) {
  if (templateSaving) return;
  if (!window.confirm('Delete “' + summary.name + '” from your saved templates? Your current document and saved work will stay.')) return;
  templateSaving = true; $('#templateError').classList.add('hidden'); $('#templateLibraryStatus').textContent = '';
  try {
    await deleteTemplate(summary.id);
    if (project.settings.templateId === summary.id) { delete project.settings.templateId; delete project.names.template; markDirty(); }
    await renderTemplateLibrary();
    $('#templateLibraryStatus').textContent = 'Template deleted: “' + summary.name + '”.'; status('Template deleted: ' + summary.name + '. Your document and saved work are kept.');
    ($('#templateList').querySelector('button') || $('#importTemplate')).focus();
  } catch (failure) { templateError(failure); }
  finally { templateSaving = false; }
}
async function openTemplates() {
  commitProjectName();
  $('#templateError').classList.add('hidden'); $('#templateLibraryStatus').textContent = ''; $('#templateList').textContent = 'Opening your templates…'; $('#templatesDialog').showModal();
  try { await renderTemplateLibrary(); } catch (failure) { $('#templateList').textContent = ''; templateError(failure); }
}
function openSaveTemplate() {
  commitProjectName();
  $('#templateName').value = '';
  $('#templateSourceDocument').textContent = 'Document in the right pane: ' + project.names.document;
  $('#templateSourceWork').textContent = 'Work: ' + project.names.project;
  $('#templateSourceDetails').textContent = findFields().length + ' connected items · ' + (project.settings.mode === 'batch' ? 'One per row' : 'One document') + ' · ' + project.settings.paper;
  $('#templateSaveError').classList.add('hidden'); $('#saveTemplateDialog').showModal(); $('#templateName').focus();
}
async function saveWorkingTemplate() {
  if (templateSaving || templateEditSession) return; templateSaving = true; $('#saveTemplate').disabled = true; $('#templateSaveError').classList.add('hidden');
  try {
    const template = createTemplate(project, documentState(), $('#templateName').value.trim());
    const id = crypto.randomUUID();
    await storeTemplate(id, template.name, JSON.stringify(template), templateFieldCount(template), template.settings.mode, 'create');
    project.settings.templateId = id; project.names.template = template.name; markDirty();
    $('#saveTemplateDialog').close(); status('New template saved: “' + template.name + '”. Find it in Templates.');
  } catch (failure) { $('#templateSaveError').textContent = readableError(failure); $('#templateSaveError').classList.remove('hidden'); $('#templateName').focus(); $('#templateName').select(); }
  finally { templateSaving = false; $('#saveTemplate').disabled = false; }
}
$('#saveTemplate').onclick = saveWorkingTemplate;
$('#templateName').onkeydown = (event) => { if (event.key === 'Enter') { event.preventDefault(); saveWorkingTemplate(); } };
$('#saveTemplateForm').onsubmit = (event) => { if (templateSaving) event.preventDefault(); };
$('#saveTemplateDialog').addEventListener('cancel', (event) => { if (templateSaving) event.preventDefault(); });
$('#templatesForm').onsubmit = (event) => { if (templateSaving) event.preventDefault(); };
$('#importTemplate').onclick = async () => {
  if (templateSaving) return; templateSaving = true;
  try {
    const file = await openFile('template'); if (!file) return;
    let template = deserializeTemplate(new TextDecoder().decode(file.bytes));
    const id = crypto.randomUUID(), existing = await listTemplates();
    while (existing.some((item) => templateNameKey(item.name) === templateNameKey(template.name))) {
      const name = window.prompt('A template named “' + template.name + '” already exists. Enter a different name for the template you are opening:', template.name + ' (copy)');
      if (name === null) return;
      if (!name.trim() || name.trim().length > 120) throw new Error('Give the template a name using 1 to 120 characters, then open the file again.');
      template = { ...template, name: name.trim() };
    }
    await storeTemplate(id, template.name, JSON.stringify(template), templateFieldCount(template), template.settings.mode, 'create'); await renderTemplateLibrary();
    $('#templateLibraryStatus').textContent = 'Template opened: “' + template.name + '”. Choose Preview document to use it, or Edit to change it.';
  } catch (failure) { templateError(failure); } finally { templateSaving = false; }
};
async function editSavedTemplate(id) {
  if (templateSaving || templateEditSession) return; templateSaving = true;
  try {
    const record = await getTemplate(id), template = deserializeTemplate(record.serialized);
    const { project: loaded, missingSheets } = projectForTemplateEditing(template, spreadsheetHasContent() ? project : null);
    // The reference workbook is copied once for this session. Edits to its
    // values, formulas, or cell connections cannot change the parked work.
    project.document = documentState();
    const backup = { project, dirty, checkpointId, importingNotes, activeField, history, redo, lastEditKey, lastEditAt, changeRevision, selection: clone(selection), view: { ...view }, currentPage, states: allEditors().map((editor) => editor.state), activeEditor, editorVisible: !$('#editingPaper').classList.contains('hidden'), headerVisible: !$('#headerFooterPanel').classList.contains('hidden'), drawerVisible: !$('#mappingDrawer').classList.contains('hidden'), documentScroll: { top: $('#documentScroll').scrollTop, left: $('#documentScroll').scrollLeft }, gridScroll: { top: $('#gridScroll').scrollTop, left: $('#gridScroll').scrollLeft } };
    templateEditSession = { id, name: template.name, backup };
    project = loaded; project.names.project = template.name; project.names.template = template.name; project.settings.templateId = id;
    checkpointId = null; importingNotes = ''; activeField = ''; history = []; redo = []; lastEditKey = ''; lastEditAt = 0;
    selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; view = { rowStart: 0, colStart: 0, rows: 50, cols: 12 }; currentPage = 1;
    $('#editingTemplateName').value = template.name; $('#projectName').disabled = true;
    $('#templateEditError').classList.add('hidden'); $('#templateEditBar').classList.remove('hidden'); document.body.classList.add('editing-template');
    $('#headerFooterPanel').classList.add('hidden'); $('#mappingDrawer').classList.add('hidden');
    setEditorContent(project.document);
    // New editor states keep Undo in this session from reaching the parked work.
    syncing = true; try { allEditors().forEach((editor) => editor.view.updateState(EditorState.create({ schema: editor.schema, doc: editor.state.doc, plugins: editor.state.plugins }))); } finally { syncing = false; }
    saved(); refresh(); $('#editorTab').click(); $('#templatesDialog').close();
    status(missingSheets.length ? 'Your spreadsheet is available on the left. No matching sheet was found for “' + missingSheets.join('”, “') + '”. Reconnect those items using Connected information, or open the matching spreadsheet.' : 'Editing template “' + template.name + '” using a copy of the spreadsheet shown on the left. Save changes keeps the layout and cell connections.');
  } catch (failure) { if (templateEditSession) finishTemplateEditing(); templateError(failure); }
  finally { templateSaving = false; }
}
function finishTemplateEditing() {
  const backup = templateEditSession.backup; clearTimeout(previewTimer); templateEditSession = null;
  project = backup.project; checkpointId = backup.checkpointId; importingNotes = backup.importingNotes; activeField = backup.activeField;
  history = backup.history; redo = backup.redo; lastEditKey = backup.lastEditKey; lastEditAt = backup.lastEditAt; changeRevision = backup.changeRevision;
  selection = backup.selection; view = backup.view; currentPage = backup.currentPage;
  document.body.classList.remove('editing-template'); $('#templateEditBar').classList.add('hidden'); $('#projectName').disabled = false;
  syncing = true; try { allEditors().forEach((editor, index) => editor.view.updateState(backup.states[index])); } finally { syncing = false; }
  activeEditor = backup.activeEditor;
  if (backup.dirty) { dirty = true; $('#dirtyDot').classList.remove('hidden'); native?.setDirty(true); updateSaveState(); } else saved();
  $('#headerFooterPanel').classList.toggle('hidden', !backup.headerVisible); $('#mappingDrawer').classList.toggle('hidden', !backup.drawerVisible);
  refresh(); $(backup.editorVisible ? '#editorTab' : '#previewTab').click(); updateToolbar();
  $('#documentScroll').scrollTo(backup.documentScroll.left, backup.documentScroll.top); $('#gridScroll').scrollTo(backup.gridScroll.left, backup.gridScroll.top);
}
async function saveEditedTemplate() {
  if (!templateEditSession || templateSaving) return; templateSaving = true; $('#saveTemplateChanges').disabled = true; $('#cancelTemplateEdit').disabled = true; $('#templateEditError').classList.add('hidden');
  try {
    const revision = changeRevision, template = createTemplate(project, documentState(), $('#editingTemplateName').value.trim());
    await storeTemplate(templateEditSession.id, template.name, JSON.stringify(template), templateFieldCount(template), template.settings.mode, 'edit');
    if (revision !== changeRevision) { templateEditSession.name = template.name; updateGuidance(); status('Template saved. You made more changes while saving; choose Save changes again to keep those too.'); return; }
    finishTemplateEditing(); status('Changes saved to template “' + template.name + '”. Your previous work is back.');
  } catch (failure) { $('#templateEditError').textContent = readableError(failure); $('#templateEditError').classList.remove('hidden'); $('#editingTemplateName').focus(); }
  finally { templateSaving = false; $('#saveTemplateChanges').disabled = false; $('#cancelTemplateEdit').disabled = false; }
}
$('#saveTemplateChanges').onclick = saveEditedTemplate;
$('#cancelTemplateEdit').onclick = () => { if (!templateEditSession || templateSaving) return; if (dirty && !window.confirm('Discard the changes to this template and return to your previous work?')) return; finishTemplateEditing(); status('Template editing canceled. Your previous work is back.'); };
$('#editingTemplateName').oninput = () => { project.names.project = $('#editingTemplateName').value; project.names.template = $('#editingTemplateName').value; markDirty(); };
$('#editingTemplateName').onkeydown = (event) => { if (event.key === 'Enter') { event.preventDefault(); saveEditedTemplate(); } };
function templateWorksheetMap() {
  return Object.fromEntries([...$('#templateWorksheetMap').querySelectorAll('select')].map((select) => [select.dataset.source, select.value]));
}
function showTemplateSpreadsheet(source, sheet) {
  templateSheetView = { source, sheet, row: 0, col: 0 };
  if (sheet) {
    const bounds = project.bounds(sheet);
    templateSheetView.row = bounds.s.r; templateSheetView.col = bounds.s.c;
  }
  $('#templateSpreadsheetCell').value = '';
  renderTemplateSpreadsheet();
}
function renderTemplateSpreadsheet() {
  const { sheet: name, source } = templateSheetView, grid = $('#templateSpreadsheetGrid');
  $('#templateSpreadsheetHint').classList.remove('invalid-preview-cell');
  const available = Boolean(name && project.workbook.SheetNames.includes(name));
  $('#templateSpreadsheetName').textContent = project.names.spreadsheet;
  $('#templateSpreadsheetName').title = project.names.spreadsheet;
  grid.classList.toggle('hidden', !available); $('#templateSpreadsheetEmpty').classList.toggle('hidden', available);
  for (const id of ['templateEarlierRows', 'templateMoreRows', 'templateEarlierColumns', 'templateMoreColumns', 'templateSpreadsheetCell', 'templateSpreadsheetGo']) $('#' + id).disabled = !available;
  if (!available) {
    grid.replaceChildren(); $('#templateSpreadsheetStatus').textContent = '';
    $('#templateSpreadsheetEmpty').textContent = project.names.spreadsheet === 'Blank spreadsheet' ? 'Open a spreadsheet to see its names and numbers here.' : 'Choose a sheet above to see its names and numbers here.';
    $('#templateSpreadsheetHint').textContent = 'The preview does not change your spreadsheet.'; return;
  }
  const sheet = project.sheet(name), bounds = project.bounds(name), view = templateSheetView;
  view.row = Math.max(bounds.s.r, Math.min(view.row, bounds.e.r)); view.col = Math.max(bounds.s.c, Math.min(view.col, bounds.e.c));
  const endRow = Math.min(bounds.e.r, view.row + templateSheetRows - 1), endCol = Math.min(bounds.e.c, view.col + templateSheetColumns - 1);
  const merges = (sheet['!merges'] || []).filter((merge) => merge.e.r >= view.row && merge.s.r <= endRow && merge.e.c >= view.col && merge.s.c <= endCol);
  const headings = new Set(selectedTemplate?.template.sheets.find((item) => item.name === source)?.headers.map((header) => header.address) || []);
  let html = '<caption class="sr-only">' + escapeHtml(name) + ' in ' + escapeHtml(project.names.spreadsheet) + '</caption><thead><tr><th scope="col"></th>';
  for (let c = view.col; c <= endCol; c++) html += '<th scope="col">' + colName(c) + '</th>';
  html += '</tr></thead><tbody>';
  for (let r = view.row; r <= endRow; r++) {
    html += '<tr><th scope="row">' + (r + 1) + '</th>';
    for (let c = view.col; c <= endCol; c++) {
      const merge = merges.find((item) => r >= item.s.r && r <= item.e.r && c >= item.s.c && c <= item.e.c);
      if (merge && (r !== Math.max(merge.s.r, view.row) || c !== Math.max(merge.s.c, view.col))) continue;
      const address = XLSX.utils.encode_cell(merge ? merge.s : { r, c }), cell = sheet[address];
      const value = project.display(name, address);
      const spans = merge ? ' rowspan="' + (Math.min(merge.e.r, endRow) - r + 1) + '" colspan="' + (Math.min(merge.e.c, endCol) - c + 1) + '"' : '';
      html += '<td' + spans + ' class="' + (cell?.t === 'n' ? 'numeric ' : '') + (headings.has(address) ? 'template-expected-heading' : '') + '" title="' + escapeHtml(address + ': ' + value) + '">' + escapeHtml(value) + '</td>';
    }
    html += '</tr>';
  }
  grid.innerHTML = html + '</tbody>';
  $('#templateSpreadsheetStatus').textContent = name + ' · Rows ' + (view.row + 1) + '–' + (endRow + 1) + ' · Columns ' + colName(view.col) + '–' + colName(endCol);
  $('#templateSpreadsheetHint').textContent = templateWorksheetMap()[source] === name ? 'Showing the sheet chosen for “' + source + '”. Shaded headings are checked against this template.' : 'Showing “' + name + '” for a first look. Choose a sheet above to use it.';
  $('#templateEarlierRows').disabled = view.row <= bounds.s.r; $('#templateMoreRows').disabled = endRow >= bounds.e.r;
  $('#templateEarlierColumns').disabled = view.col <= bounds.s.c; $('#templateMoreColumns').disabled = endCol >= bounds.e.c;
  $('#templateSpreadsheetViewport').scrollTo(0, 0);
}
function moveTemplateSpreadsheet(rows, columns) {
  if (!templateSheetView.sheet) return;
  templateSheetView.row += rows; templateSheetView.col += columns; renderTemplateSpreadsheet();
}
$('#templateEarlierRows').onclick = () => moveTemplateSpreadsheet(-templateSheetRows, 0);
$('#templateMoreRows').onclick = () => moveTemplateSpreadsheet(templateSheetRows, 0);
$('#templateEarlierColumns').onclick = () => moveTemplateSpreadsheet(0, -templateSheetColumns);
$('#templateMoreColumns').onclick = () => moveTemplateSpreadsheet(0, templateSheetColumns);
function jumpTemplateSpreadsheet() {
  const input = $('#templateSpreadsheetCell');
  try {
    const cell = parseCell(input.value.trim()), bounds = project.bounds(templateSheetView.sheet);
    if (cell.r < bounds.s.r || cell.r > bounds.e.r || cell.c < bounds.s.c || cell.c > bounds.e.c) throw new Error('Choose a cell in ' + XLSX.utils.encode_range(bounds) + ', the area used by this sheet.');
    templateSheetView.row = cell.r; templateSheetView.col = cell.c; renderTemplateSpreadsheet();
  } catch (failure) { $('#templateSpreadsheetHint').textContent = readableError(failure); $('#templateSpreadsheetHint').classList.add('invalid-preview-cell'); input.focus(); }
}
$('#templateSpreadsheetGo').onclick = jumpTemplateSpreadsheet;
$('#templateSpreadsheetCell').onkeydown = (event) => { if (event.key === 'Enter') { event.preventDefault(); jumpTemplateSpreadsheet(); } };
function updateTemplateCompatibility() {
  if (!selectedTemplate) return;
  const issues = templateCompatibility(selectedTemplate.template, project, templateWorksheetMap());
  $('#applyTemplate').disabled = !templatePreviewReady || Boolean(issues.length); $('#templateStartBlank').disabled = !templatePreviewReady;
  $('#templateCompatibility').textContent = issues.length ? 'Check these headings before using the template:\n' + issues.slice(0, 5).join('\n') + (issues.length > 5 ? '\nAnd ' + (issues.length - 5) + ' more differences.' : '') : 'The headings match. This document will use ' + templateFieldCount(selectedTemplate.template) + ' connected items from your spreadsheet.';
}
function populateTemplateWorksheets(template) {
  const container = $('#templateWorksheetMap'); container.replaceChildren();
  const suggested = suggestTemplateWorksheets(template, project);
  for (const sheet of template.sheets) {
    const label = document.createElement('label'), select = document.createElement('select');
    label.append('Read “' + sheet.name + '” from:'); select.dataset.source = sheet.name; select.add(new Option('Choose a sheet in your spreadsheet', ''));
    for (const name of project.workbook.SheetNames) select.add(new Option(name, name));
    select.value = suggested[sheet.name];
    select.onchange = () => { showTemplateSpreadsheet(sheet.name, select.value); renderTemplateDocument(template); };
    label.append(select); container.append(label);
    if (template.sheets.length > 1) {
      const preview = document.createElement('button'); preview.type = 'button'; preview.className = 'template-see-sheet'; preview.textContent = 'View this sheet';
      preview.onclick = () => showTemplateSpreadsheet(sheet.name, select.value); label.append(preview);
    }
  }
  const choices = [...container.querySelectorAll('select')], selected = choices.find((select) => select.value);
  showTemplateSpreadsheet(selected?.dataset.source || template.settings.sheet, selected?.value || (project.names.spreadsheet !== 'Blank spreadsheet' ? project.workbook.SheetNames[0] : ''));
}
function openTemplateApplication(id, template) {
  selectedTemplate = { id, template }; templatePreviewReady = false; templatePreviewGeometry = null;
  $('#templatesDialog').close(); $('#templateApplyName').textContent = template.name; $('#templateApplyError').classList.add('hidden');
  $('#templateDocumentPreview').replaceChildren(); $('#templatePreviewStatus').textContent = 'Opening document…'; $('#templatePreviewZoom').value = 'fit';
  populateTemplateWorksheets(template);
  updateTemplateCompatibility(); $('#templateApplyDialog').showModal(); renderTemplateDocument(template, true);
}
$('#templateChooseSpreadsheet').onclick = async () => {
  if (!selectedTemplate) return;
  $('#templateChooseSpreadsheet').disabled = true;
  try {
    const previousWorkbook = project.workbook;
    const result = await action('import-sheet');
    if (!selectedTemplate || !$('#templateApplyDialog').open) return;
    if (project.workbook !== previousWorkbook) { populateTemplateWorksheets(selectedTemplate.template); renderTemplateDocument(selectedTemplate.template); }
    if (result === false) templateError(new Error($('#globalStatus').textContent), true);
    else $('#templateApplyError').classList.add('hidden');
  } finally { $('#templateChooseSpreadsheet').disabled = false; }
};
function templateDocumentHtml(documentContent, preview, worksheetMap, useValues, showEmptyLocations = false) {
  const content = clone(documentContent);
  const resolveFields = (node) => {
    if (node.type === 'linkedField') {
      const field = node.attrs, target = worksheetMap[field.sheet];
      const location = field.sheet + '!' + field.address;
      // A missing choice must never fall back to the old sheet or saved value.
      if (useValues && target && preview.workbook.SheetNames.includes(target)) {
        const source = preview.source({ ...field, sheet: target });
        field.sheet = source.sheet; field.address = source.address; field.followRecord = false;
        field.label = source.sheet + '!' + source.address;
        field.value = preview.display(source.sheet, source.address);
        if (showEmptyLocations && field.value === '') field.value = '[' + field.label + ']';
      } else { field.label = location; field.value = '[' + location + ']'; }
    }
    for (const child of node.content || []) resolveFields(child);
  };
  resolveFields(content);
  const root = document.createElement('div'); root.innerHTML = generateHTML(content, mainEditor.extensionManager.extensions);
  for (const field of root.querySelectorAll('[data-field-id]')) field.title = field.dataset.label;
  return root.innerHTML;
}
function fitTemplatePreview() {
  if (!templatePreviewGeometry || !$('#templateApplyDialog').open) return;
  const setting = $('#templatePreviewZoom').value;
  const zoom = setting === 'fit' ? Math.min(1, Math.max(.2, ($('#templatePreviewViewport').clientWidth - 36) / templatePreviewGeometry.widthPx)) : Number(setting);
  $('#templateDocumentPreview').style.zoom = String(zoom);
}
async function renderTemplateDocument(template, resetScroll = false) {
  const generation = ++templatePreviewGeneration, host = $('#templateMeasurementHost');
  const viewport = $('#templatePreviewViewport'), scroll = { top: viewport.scrollTop, left: viewport.scrollLeft };
  templatePreviewReady = false; templatePreviewGeometry = null;
  $('#templateDocumentPreview').replaceChildren(); $('#templatePreviewStatus').textContent = 'Updating document…';
  $('#templateApplyError').classList.add('hidden'); updateTemplateCompatibility();
  try {
    const worksheetMap = templateWorksheetMap(), preview = new Project(), useValues = spreadsheetHasContent();
    // Share the read-only workbook and row index. No records are copied or recalculated.
    preview.workbook = project.workbook; preview.recordIndex = project.recordIndex;
    preview.settings = { ...preview.settings, ...template.settings, sheet: worksheetMap[template.settings.sheet], batchSheet: worksheetMap[template.settings.batchSheet] };
    if (template.settings.mode === 'batch' && preview.workbook.SheetNames.includes(preview.settings.batchSheet)) preview.settings.activeRow = preview.recordRows()[0] ?? template.settings.headerRow + 1;
    else preview.settings.activeRow = template.settings.headerRow + 1;
    const missingSheets = template.sheets.some((sheet) => !worksheetMap[sheet.name]);
    $('#templateDocumentHint').textContent = !useValues ? 'Choose a spreadsheet and sheet to fill the connected locations shown in brackets.' : 'Values come from ' + project.names.spreadsheet + (template.settings.mode === 'batch' && preview.settings.batchSheet ? ' · First document: spreadsheet row ' + (preview.settings.activeRow + 1) : '') + '. Hover over a connected value to see its cell.' + (missingSheets ? ' Choose the remaining sheets to fill every location.' : ' Empty cells stay blank.');
    const parts = [template.document, template.settings.headerDocument, template.settings.footerDocument].map((documentContent) => templateDocumentHtml(documentContent, preview, worksheetMap, useValues));
    // Decode embedded images before measuring the saved layout. This host is separate
    // from the working document's preview and does not create an editable editor.
    const source = document.createElement('div'); source.className = 'tiptap';
    const sections = parts.map((html) => { const section = document.createElement('div'); section.innerHTML = html; source.append(section); return section; });
    host.replaceChildren(source);
    await Promise.all([document.fonts.ready, ...[...source.querySelectorAll('img')].map((image) => image.decode().catch(() => {}))]);
    if (generation !== templatePreviewGeneration || !$('#templateApplyDialog').open) return;
    for (const image of source.querySelectorAll('img')) if (image.naturalWidth && image.naturalHeight && !image.hasAttribute('width') && !image.hasAttribute('height') && !image.style.width && !image.style.height) {
      image.setAttribute('width', String(image.naturalWidth)); image.setAttribute('height', String(image.naturalHeight));
    }
    const rendered = sections.map((section) => section.innerHTML);
    const layout = paginate(rendered[0], template.settings, host, rendered[1], rendered[2]);
    $('#templateDocumentPreview').replaceChildren(...layout.pages); templatePreviewGeometry = layout.geometry;
    $('#templatePreviewStatus').textContent = layout.pages.length + ' page' + (layout.pages.length === 1 ? '' : 's') + ' · ' + template.settings.paper + (template.settings.landscape ? ' landscape' : ' portrait');
    if (layout.errors.length) { $('#templateApplyError').textContent = layout.errors.join(' '); $('#templateApplyError').classList.remove('hidden'); }
    templatePreviewReady = true; fitTemplatePreview(); viewport.scrollTo(resetScroll ? 0 : scroll.left, resetScroll ? 0 : scroll.top); updateTemplateCompatibility();
  } catch (failure) {
    if (generation !== templatePreviewGeneration || !$('#templateApplyDialog').open) return;
    $('#templatePreviewStatus').textContent = 'Preview unavailable'; templateError(failure, true);
  } finally { if (generation === templatePreviewGeneration) host.replaceChildren(); }
}
$('#templatePreviewZoom').onchange = fitTemplatePreview; window.addEventListener('resize', fitTemplatePreview);
$('#templateDocumentPreview').onclick = (event) => { if (event.target.closest('a')) event.preventDefault(); };
$('#templateApplyDialog').addEventListener('close', () => {
  templatePreviewGeneration++; templatePreviewReady = false; templatePreviewGeometry = null; selectedTemplate = null;
  templateSheetView = { sheet: '', source: '', row: 0, col: 0 }; $('#templateSpreadsheetGrid').replaceChildren();
  $('#templateDocumentPreview').replaceChildren(); $('#templateMeasurementHost').replaceChildren();
});
function useSelectedTemplate(blank = false) {
  if (!selectedTemplate || templateSaving || !templatePreviewReady) return;
  if (blank && !canReplaceProject()) return;
  try {
    const loaded = projectFromTemplate(selectedTemplate.template, blank ? null : project, blank ? null : templateWorksheetMap());
    loaded.settings.templateId = selectedTemplate.id;
    project = loaded; checkpointId = null; importingNotes = ''; activeField = ''; history = []; redo = [];
    selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; view.rowStart = 0; view.colStart = 0; currentPage = 1;
    setEditorContent(project.document); markDirty(); refresh(); $('#templateApplyDialog').close();
    status(blank ? 'Your document is ready. Enter values on the left or open your spreadsheet.' : 'Your document now uses the new spreadsheet values. Check it, then choose Save file / Print.');
  } catch (failure) { templateError(failure, true); }
}
$('#applyTemplate').onclick = () => useSelectedTemplate(); $('#templateStartBlank').onclick = () => useSelectedTemplate(true);
$('#templateBack').onclick = () => { $('#templateApplyDialog').close(); openTemplates(); };
function emailDocumentLayout(entry) {
  if (project.formulaError) throw new Error(readableError(new Error(project.formulaError)));
  validateFields(entry.row);
  const layout = layoutForRow(entry.row);
  if (layout.errors.length) throw new Error('Spreadsheet row ' + (entry.row + 1) + ': ' + layout.errors.join(' '));
  return layout;
}
const emailDrafts = setupEmailDrafts({
  getProject: () => project, native, changed: markDirty, saveBytes, status,
  configureDocuments: ({ mode, sheet, headerRow }) => {
    project.sheet(sheet);
    if (!['single', 'batch'].includes(mode) || !Number.isInteger(headerRow) || headerRow < 0 || headerRow > 1048574) throw new Error('Check the document mode and headings row.');
    if (mode === 'batch') project.settings.batchSheet = sheet;
    project.settings.headerRow = headerRow;
    setMode(mode);
  },
  documentHint: () => {
    if (project.settings.mode !== 'batch') return 'The attachment uses your current document and connected cell locations.';
    const count = findFields().filter((field) => field.attrs.sheet === project.settings.batchSheet && field.attrs.followRecord !== false).length;
    return count ? count + ' connected item' + (count === 1 ? '' : 's') + ' will use each person\'s row in ' + project.settings.batchSheet + '. Text typed directly into the document stays the same.' : 'This document has no connections that change with each row of ' + project.settings.batchSheet + '. Close this window and connect the person\'s spreadsheet cells to the document to personalize attachments.';
  },
  documentPreview: (entry) => emailDocumentLayout(entry).pages.map((page) => page.cloneNode(true)),
  busy: (value) => { busy = value; document.body.classList.toggle('busy', value); },
  readyAssets: async () => { await document.fonts.ready; await Promise.all(allEditors().flatMap((editor) => [...editor.view.dom.querySelectorAll('img')]).map((image) => image.decode().catch(() => {}))); },
  attachment: async (entry, format) => {
    const layout = emailDocumentLayout(entry);
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(printHtml(layout.pages, project.settings)));
    const contentKey = [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
    return { bytes: await documentBytes(layout.pages, project.settings, format, native), contentKey };
  }
});
async function action(name) {
  try {
    if (busy || checkpointSaving || templateSaving || initializing) return;
    if ($('#emailDialog').open && !['save-project', 'help'].includes(name)) { status('Close Email documents before changing your files or document.'); return; }
    if (templateEditSession && name === 'save-project') { await saveEditedTemplate(); return; }
    if (templateEditSession && !['help', 'import-sheet', 'import-doc', 'blank-doc', 'blank-sheet', 'image'].includes(name)) { status('You are editing a template. Choose Save changes or Cancel editing to return to your work.'); return; }
    if (name === 'help') { $('#helpDialog').showModal(); return; }
    if (name === 'export') { openExport(); return; }
    if (name === 'email-drafts') { commitProjectName(); emailDrafts.open(); return; }
    if (name === 'templates') { await openTemplates(); return; }
    if (name === 'save-template') { openSaveTemplate(); return; }
    if (name === 'sample') { if (dirty && !window.confirm('Open the example instead of your current work? Choose Save work first if you want to keep your changes.')) return; loadSample(); return; }
    if (name === 'image') { $('#imageFile').click(); return; }
    if (name === 'new' && !canReplaceProject()) return;
    if (name === 'new') { for (const dialog of document.querySelectorAll('dialog[open]')) dialog.close(); project = new Project(); checkpointId = null; importingNotes = ''; activeField = ''; history = []; redo = []; selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; view.rowStart = 0; view.colStart = 0; currentPage = 1; setEditorContent(project.document); saved(); refresh(); openStartChoices(); return; }
    if (name === 'save-project') { await checkpointProject(); return; }
    if (name === 'open-project') { await openCheckpoints(); return; }
    if (name === 'import-sheet') {
      const file = await openFile('spreadsheet'); if (!file) return; const loaded = new Project(); loaded.importWorkbook(file.bytes, file.name);
      const previousLayout = { sheet: project.settings.sheet, batchSheet: project.settings.batchSheet, headerRow: project.settings.headerRow };
      history.push(project.workbook); redo = []; project.workbook = loaded.workbook; project.settings.sheet = loaded.settings.sheet; project.settings.batchSheet = loaded.settings.batchSheet; project.settings.headerRow = 0; project.settings.activeRow = 1; project.names.spreadsheet = file.name; project.formulaError = '';
      if (project.settings.emailDraft) delete project.settings.emailDraft.recipientOverrides;
      if (['Untitled work', 'Untitled project'].includes(project.names.project)) project.names.project = file.name.replace(/\.[^.]+$/, '').slice(0, 120) || 'Untitled work';
      if (project.settings.templateId && project.workbook.SheetNames.includes(previousLayout.sheet) && project.workbook.SheetNames.includes(previousLayout.batchSheet)) {
        Object.assign(project.settings, previousLayout); project.settings.activeRow = project.settings.mode === 'batch' ? recordRows()[0] ?? project.settings.headerRow + 1 : project.settings.headerRow + 1;
      }
      view = { rowStart: 0, colStart: 0, rows: 50, cols: 12 }; selection = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }; markDirty(); refresh(); status('Spreadsheet opened. Choose a template or add values to your document.'); return true;
    }
    if (name === 'import-doc') {
      const file = await openFile('document'); if (!file) return; status('Opening your document…'); const imported = await importDocument(file.bytes, file.name);
      project.settings = { ...project.settings, ...imported.layout }; syncing = true; mainEditor.commands.setContent(imported.html || '<p></p>'); headerEditor.commands.setContent(imported.header || '<p></p>'); footerEditor.commands.setContent(imported.footer || '<p></p>'); syncing = false;
      project.names.document = file.name; importingNotes = imported.warnings.join(' '); activeField = ''; activeEditor = mainEditor; markDirty(); refresh(); status('Document opened. You can write in it or add spreadsheet values. Choose Check pages to see the layout.'); return;
    }
    if (name === 'blank-doc') { if (!window.confirm('Clear this document and start with a blank page? Your spreadsheet will be kept.')) return; syncing = true; allEditors().forEach((editor) => editor.commands.setContent('<p></p>')); syncing = false; project.names.document = 'Blank document'; importingNotes = ''; activeField = ''; markDirty(); refresh(); return; }
    if (name === 'blank-sheet') { if (!window.confirm('Clear the spreadsheet and start blank? Connected values in your document will also become empty.')) return; project.workbook = new Project().workbook; project.settings.sheet = 'Sheet1'; project.settings.batchSheet = 'Sheet1'; project.names.spreadsheet = 'Blank spreadsheet'; project.formulaError = ''; if (project.settings.emailDraft) delete project.settings.emailDraft.recipientOverrides; markDirty(); refresh(); }
  } catch (failure) { syncing = false; error(failure); return false; }
}
function loadSample() {
  project = new Project(); checkpointId = null; const data = [
    ['Employee name', 'Department', 'Current salary', 'Increment', 'Final salary', 'Effective date', 'Employee email', 'Manager email'],
    ['Avery Tan', 'Operations', 4200, 300, null, '1 January 2027', 'avery.tan@example.com', 'manager@example.com'],
    ['Jordan Lee', 'Finance', 5100, 400, null, '1 January 2027', 'jordan.lee@example.com', 'manager@example.com'],
    ['Morgan Lim', 'People', 3900, 250, null, '1 January 2027', 'morgan.lim@example.com', 'manager@example.com']
  ], ws = XLSX.utils.aoa_to_sheet(data);
  for (let r = 2; r <= 4; r++) { ws['E' + r] = { t: 'n', f: 'C' + r + '+D' + r, v: data[r - 1][2] + data[r - 1][3], z: '#,##0.00' }; ws['C' + r].z = '#,##0.00'; ws['D' + r].z = '#,##0.00'; }
  project.workbook = { SheetNames: ['Employees'], Sheets: { Employees: ws } }; Object.assign(project.settings, { sheet: 'Employees', batchSheet: 'Employees', mode: 'batch', activeRow: 1 });
  project.names = { spreadsheet: 'Fictional payroll sample.xlsx', document: 'Salary adjustment template', project: 'Sample project' };
  const field = (address) => newField('Employees', address), p = (text) => ({ type: 'paragraph', content: [{ type: 'text', text }] });
  const valueRow = (label, address) => ({ type: 'tableRow', content: [{ type: 'tableCell', content: [p(label)] }, { type: 'tableCell', content: [{ type: 'paragraph', content: [field(address)] }] }] });
  const doc = { type: 'doc', content: [
    { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Salary adjustment notice' }] },
    { type: 'paragraph', content: [{ type: 'text', text: 'Dear ' }, field('A2'), { type: 'text', text: ',' }] },
    p('We are pleased to confirm the following adjustment to your monthly salary.'),
    { type: 'table', content: [valueRow('Department', 'B2'), valueRow('Current salary', 'C2'), valueRow('Increment amount', 'D2'), valueRow('Final salary', 'E2')] },
    { type: 'paragraph', content: [{ type: 'text', text: 'Effective date: ' }, field('F2')] }, p('Thank you for your continued contribution.'), p('Authorized signatory: __________________________')
  ] };
  importingNotes = ''; history = []; redo = []; activeField = ''; setEditorContent(doc); headerEditor.commands.setContent('<p><strong>EXAMPLE COMPANY</strong></p>'); footerEditor.commands.setContent('<p>Fictional sample · for manual testing</p>'); markDirty(); refresh(); status('Example opened. Double-click D2 and enter 650. Both sides will update. Use the arrows to see the next person.');
}
$('#imageFile').onchange = (event) => { const file = event.target.files[0]; if (!file) return; if (file.size > 10 * 1024 * 1024) { error(new Error('Choose an image smaller than 10 MB.')); return; } const reader = new FileReader(); reader.onload = () => activeEditor.chain().focus().setImage({ src: reader.result, alt: file.name }).run(); reader.readAsDataURL(file); event.target.value = ''; };
document.querySelectorAll('[data-action]').forEach((button) => button.onclick = () => action(button.dataset.action)); native?.onAction(action);
function refresh() { sheetControls(); syncFields(); renderGrid(); renderMappingList(); updateFormulaStatus(); renderPreview(); }

const DOCUMENT_FORMATS = ['docx', 'pdf', 'png', 'jpeg', 'html', 'txt'], SHEET_FORMATS = ['xlsx', 'ods', 'csv', 'tsv', 'pdf', 'png', 'jpeg'];
const FORMAT_NAMES = { docx: 'Word document (.docx)', pdf: 'PDF document (.pdf)', xlsx: 'Excel spreadsheet (.xlsx)', png: 'PNG picture (.png)', jpeg: 'JPEG picture (.jpg)', html: 'Web page (.html)', txt: 'Plain text (.txt)', ods: 'OpenDocument spreadsheet (.ods)', csv: 'CSV spreadsheet (.csv)', tsv: 'TSV spreadsheet (.tsv)', paperlink: 'Work backup for PaperLink', papertemplate: 'Template file for PaperLink' };
const batchNameOverrides = new Map(), batchNameLabels = new Map(), batchNameInputs = new Map();
let batchPreviewRows = [];
function exportBaseName() { return validateFileName($('#exportFileName').value); }
function separateBatchFiles() { return $('#exportTarget').value === 'document' && project.settings.mode === 'batch' && $('#batchPackaging').value === 'separate'; }
function exportRecordRows() {
  const rows = project.settings.mode === 'batch' && $('#exportPages').value !== 'current' ? project.recordRows($('#exportRecords').value) : [project.settings.activeRow];
  if (!rows.length) throw new Error('There are no filled rows below the headings. Open your spreadsheet or add the information first.');
  if (rows.length > 1000) throw new Error('Save up to 1,000 documents at a time. Use Choose specific spreadsheet rows to save a smaller group.');
  return rows;
}
function setupBatchFileNames() {
  batchNameOverrides.clear();
  $('#batchNamePrefix').value = '';
  if (project.settings.mode !== 'batch') { $('#batchNameColumn').replaceChildren(new Option('Spreadsheet row number', '')); return; }
  const columns = new Set(), sheet = project.settings.batchSheet, ws = project.sheet(sheet);
  for (const address in ws) {
    if (!Object.hasOwn(ws, address) || address.startsWith('!') || ws[address]?.v == null || String(ws[address].v).trim() === '') continue;
    const match = address.match(/^([A-Z]+)\d+$/i); if (match) columns.add(match[1].toUpperCase());
  }
  const options = [new Option('Spreadsheet row number', '')];
  let nameColumn = '';
  for (const column of [...columns].sort((a, b) => XLSX.utils.decode_col(a) - XLSX.utils.decode_col(b))) {
    const heading = project.display(sheet, column + (project.settings.headerRow + 1)).trim();
    options.push(new Option((heading || 'Column ' + column) + (heading ? ' (column ' + column + ')' : ''), column));
    if (!nameColumn && /^(?:employee(?:\s+name)?|name|employee\s+full\s+name|full\s+name|recipient(?:\s+name)?)$/i.test(heading)) nameColumn = column;
  }
  $('#batchNameColumn').replaceChildren(...options); $('#batchNameColumn').value = nameColumn;
}
function batchRawFileName(row) {
  if (batchNameOverrides.has(row)) return batchNameOverrides.get(row);
  const column = $('#batchNameColumn').value, fallback = 'Document - row ' + (row + 1);
  const value = column ? project.display(project.settings.batchSheet, column + (row + 1)).trim() : '';
  return suggestedFileName([$('#batchNamePrefix').value.trim(), value || (column ? fallback : 'row ' + (row + 1))].filter(Boolean).join(' - '), fallback);
}
function batchFileNames(rows) {
  return uniqueFileNames(rows.map((row) => {
    try { return { row, name: validateFileName(batchRawFileName(row)) }; }
    catch (failure) { throw new Error('Spreadsheet row ' + (row + 1) + ': ' + failure.message); }
  }));
}
function updateBatchNameLabels() {
  try {
    const format = $('#exportFormat').value, ending = format === 'jpeg' ? 'jpg' : format;
    for (const entry of batchFileNames(batchPreviewRows)) batchNameLabels.get(entry.row).textContent = 'Will save as: ' + entry.name + '.' + ending + (['png', 'jpeg'].includes(format) ? ' (or .zip for several pages)' : '');
    $('#batchNameError').classList.add('hidden');
  } catch (failure) {
    for (const label of batchNameLabels.values()) label.textContent = 'Check the file name below.';
    $('#batchNameError').textContent = failure.message; $('#batchNameError').classList.remove('hidden');
  }
}
function renderBatchFileNames() {
  const separate = separateBatchFiles(); $('#batchFileNaming').classList.toggle('hidden', !separate);
  $('#exportFileNameLabel').textContent = separate ? 'ZIP file name' : 'File name';
  if (!separate) return;
  $('#batchNamesBody').replaceChildren(); batchNameLabels.clear(); batchNameInputs.clear(); batchPreviewRows = [];
  $('#batchNameSource').textContent = 'Uses values from the "' + project.settings.batchSheet + '" sheet. Individual name edits stay until you reset them or close this window.';
  try {
    batchPreviewRows = exportRecordRows();
    $('#batchNameCount').textContent = batchPreviewRows.length + ' file name' + (batchPreviewRows.length === 1 ? '' : 's');
    const fragment = document.createDocumentFragment();
    for (const row of batchPreviewRows) {
      const tr = document.createElement('tr'), rowCell = document.createElement('td'), nameCell = document.createElement('td'), input = document.createElement('input'), label = document.createElement('div');
      rowCell.textContent = String(row + 1); input.value = batchRawFileName(row); input.maxLength = 160; input.spellcheck = false;
      input.setAttribute('aria-label', 'File name for spreadsheet row ' + (row + 1)); label.id = 'batch-file-name-' + row; label.className = 'small-note'; input.setAttribute('aria-describedby', label.id);
      batchNameLabels.set(row, label); batchNameInputs.set(row, input);
      input.oninput = () => { batchNameOverrides.set(row, input.value); updateBatchNameLabels(); };
      nameCell.append(input, label); tr.append(rowCell, nameCell); fragment.append(tr);
    }
    $('#batchNamesBody').append(fragment); updateBatchNameLabels();
  } catch (failure) {
    $('#batchNameCount').textContent = 'Choose rows to preview their file names';
    $('#batchNameError').textContent = failure.message; $('#batchNameError').classList.remove('hidden');
  }
}
function updateSuggestedBatchNames() {
  // Changing a prefix must not rebuild hundreds of editable rows on each keypress.
  for (const [row, input] of batchNameInputs) if (!batchNameOverrides.has(row)) input.value = batchRawFileName(row);
  updateBatchNameLabels();
}
function openExport() {
  commitProjectName(); renderPreview();
  $('#exportFileName').value = stripOutputExtension(project.names.project).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/, '') || 'Untitled work';
  setupBatchFileNames();
  $('#exportRecords').value = ''; $('#sheetExportRange').value = XLSX.utils.encode_range(project.bounds()); $('#exportPageRange').value = ''; $('#exportPages').value = 'all'; $('#exportError').classList.add('hidden'); exportOptions(); $('#exportDialog').showModal();
}
function exportOptions() {
  const target = $('#exportTarget').value, old = $('#exportFormat').value, formats = target === 'template' ? ['papertemplate'] : target === 'project' ? ['paperlink'] : target === 'document' ? DOCUMENT_FORMATS : SHEET_FORMATS;
  $('#exportFormat').innerHTML = formats.map((format) => '<option value="' + format + '">' + FORMAT_NAMES[format] + '</option>').join('');
  if (formats.includes(old)) $('#exportFormat').value = old; updateExportUI();
}
function updateExportUI() {
  const target = $('#exportTarget').value, doc = target === 'document', projectFile = ['project', 'template'].includes(target), format = $('#exportFormat').value, pageFormat = doc || ['pdf', 'png', 'jpeg'].includes(format);
  $('#spreadsheetExportOptions').classList.toggle('hidden', target !== 'spreadsheet'); $('#batchExportOptions').classList.toggle('hidden', !doc || project.settings.mode !== 'batch'); $('#pageExportOptions').classList.toggle('hidden', !pageFormat); $('#imageQualityLabel').classList.toggle('hidden', !['png', 'jpeg'].includes(format));
  const workbookOption = $('#sheetExportScope').options[0];
  workbookOption.disabled = !doc && pageFormat;
  workbookOption.textContent = ['csv', 'tsv'].includes(format) ? 'The sheet you are viewing' : 'Whole spreadsheet (all sheets)';
  if (!doc && pageFormat) $('#sheetExportScope').value = 'range';
  $('#sheetExportRange').disabled = $('#sheetExportScope').value === 'workbook' && !pageFormat; $('#exportPageRange').disabled = $('#exportPages').value !== 'custom'; $('#nativePrint').disabled = !native || projectFile;
  $('#customPageLabel').classList.toggle('hidden', $('#exportPages').value !== 'custom');
  $('#sheetRangeLabel').classList.toggle('hidden', $('#sheetExportRange').disabled);
  $('#sheetRangeHint').classList.toggle('hidden', $('#sheetExportScope').value !== 'range' || pageFormat);
  $('#exportPageHint').textContent = $('#exportPages').value === 'current' ? 'Uses the page you selected under Check pages. Only the document currently shown is included.' : $('#exportPages').value === 'custom' ? (project.settings.mode === 'batch' && $('#batchPackaging').value === 'separate' && doc ? 'These page numbers are used for each document. ' : 'These page numbers are used for the finished file. ') + 'For example, 2- means page 2 to the end.' : 'Includes every page in the finished document.';
  let summary;
  if (target === 'template') summary = 'Keeps this document and its connections for reuse in PaperLink. Spreadsheet records are not included. Text you typed directly stays as written.';
  else if (projectFile) summary = 'Keeps your spreadsheet, document, and connections together. Open this backup in PaperLink to continue on another device.';
  else if (!doc) summary = 'Saves your spreadsheet changes. Your original file is kept.';
  else if ($('#exportPages').value === 'current') summary = 'Saves the selected preview page of the document you are looking at.';
  else if (project.settings.mode === 'batch') {
    try { const count = project.recordRows($('#exportRecords').value).length; summary = 'Creates ' + count + ' document' + (count === 1 ? '' : 's') + ', one for each filled row you include.'; }
    catch { summary = 'Choose the spreadsheet rows to include, such as 2-5, or leave the box empty for all filled rows.'; }
    summary += $('#batchPackaging').value === 'separate' ? ' A ZIP file keeps the separate documents together; open it to find each file.' : ' All documents will be in one file.';
  } else summary = 'Saves the document you have prepared. Your original file is kept.';
  if (['png', 'jpeg'].includes(format)) summary += ' Each page becomes a picture. More than one page is saved in a ZIP file.';
  $('#exportSummary').textContent = summary;
  const extension = doc && project.settings.mode === 'batch' && $('#batchPackaging').value === 'separate' ? 'zip' : format === 'jpeg' ? 'jpg' : format;
  $('#exportFileHint').textContent = 'We will add .' + extension + ' to the file name.' + (['png', 'jpeg'].includes(format) ? ' For several pages, the file ends in .zip.' : '');
  renderBatchFileNames();
}
$('#exportTarget').onchange = exportOptions; for (const id of ['#exportFormat', '#exportPages', '#batchPackaging', '#sheetExportScope']) $(id).onchange = updateExportUI;
$('#exportRecords').oninput = updateExportUI;
$('#batchNameColumn').onchange = updateSuggestedBatchNames; $('#batchNamePrefix').oninput = updateSuggestedBatchNames;
$('#batchResetNames').onclick = () => { batchNameOverrides.clear(); renderBatchFileNames(); };
function spreadsheetHtml(rangeText) {
  if (!/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(rangeText)) throw new Error('Use a cell range such as A1:F20.');
  const range = XLSX.utils.decode_range(rangeText.toUpperCase()); if (range.e.r < range.s.r || range.e.c < range.s.c || (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) > 20000) throw new Error('Choose a valid print range of at most 20,000 cells.');
  let html = '<table><tbody>'; for (let r = range.s.r; r <= range.e.r; r++) { html += '<tr>'; for (let c = range.s.c; c <= range.e.c; c++) { const tag = r === range.s.r ? 'th' : 'td'; html += '<' + tag + '><p>' + escapeHtml(project.display(project.settings.sheet, XLSX.utils.encode_cell({ r, c }))) + '</p></' + tag + '>'; } html += '</tr>'; } return html + '</tbody></table>';
}
function selectPages(pages) {
  const mode = $('#exportPages').value, indexes = mode === 'all' ? parseRanges('', pages.length) : mode === 'current' ? parseRanges(String(currentPage), pages.length) : parseRanges($('#exportPageRange').value, pages.length);
  return renumber(indexes.map((page) => pages[page - 1].cloneNode(true)), project.settings.pageNumbers);
}
function validateFields(row) {
  for (const field of findFields()) { const source = project.source(field.attrs, row), cell = project.cell(source.sheet, source.address); if (cell?.t === 'e') throw new Error('A linked field contains a spreadsheet error: ' + source.sheet + '!' + source.address + '.'); if (cell?.f && cell.v === undefined) throw new Error('A linked formula has no saved result: ' + source.sheet + '!' + source.address + '.'); }
}
function outputGroups({ nameFiles = true } = {}) {
  if (project.formulaError) throw new Error(project.formulaError);
  if ($('#exportTarget').value === 'spreadsheet') { const layout = paginate(spreadsheetHtml($('#sheetExportRange').value), project.settings, $('#measurementHost')); if (layout.errors.length) throw new Error(layout.errors.join(' ')); return [{ name: 'spreadsheet', pages: selectPages(layout.pages) }]; }
  const rows = exportRecordRows(), names = nameFiles && separateBatchFiles() ? new Map(batchFileNames(rows).map((entry) => [entry.row, entry.name])) : new Map();
  const groups = [];
  for (const row of rows) { validateFields(row); const layout = layoutForRow(row); if (layout.errors.length) throw new Error('Record row ' + (row + 1) + ': ' + layout.errors.join(' ')); groups.push({ name: names.get(row) || (project.settings.mode === 'batch' ? 'record-row-' + (row + 1) : 'document'), pages: layout.pages.map((page) => page.cloneNode(true)) }); }
  if ($('#batchPackaging').value === 'separate' && project.settings.mode === 'batch') return groups.map((group) => ({ ...group, pages: selectPages(group.pages) }));
  return [{ name: 'document', pages: selectPages(renumber(groups.flatMap((group) => group.pages), project.settings.pageNumbers)) }];
}
async function runOutput(print = false) {
  if (busy) return; busy = true; document.body.classList.add('busy'); $('#runExport').disabled = true; $('#nativePrint').disabled = true; $('#exportError').classList.add('hidden');
  try {
    await document.fonts.ready;
    await Promise.all(allEditors().flatMap((editor) => [...editor.view.dom.querySelectorAll('img')]).map((image) => image.decode().catch(() => {})));
    const baseName = print ? '' : exportBaseName(), target = $('#exportTarget').value, format = $('#exportFormat').value, sheetData = target === 'spreadsheet' && ['xlsx', 'ods', 'csv', 'tsv'].includes(format) && !print;
    let bytes, extension;
    if (target === 'template') { if (print) throw new Error('Choose Finished document or Spreadsheet to print.'); const template = createTemplate(project, documentState(), project.names.template || project.names.project); bytes = new TextEncoder().encode(JSON.stringify(template)); extension = 'papertemplate'; }
    else if (target === 'project') { if (print) throw new Error('Choose Finished document or Spreadsheet to print.'); bytes = new TextEncoder().encode(project.serialize(documentState())); extension = 'paperlink'; }
    else if (sheetData) { bytes = spreadsheetBytes(project, format, $('#sheetExportScope').value === 'range' ? $('#sheetExportRange').value : null); extension = format; }
    else {
      const groups = outputGroups({ nameFiles: !print }), scale = Number($('#imageQuality').value);
      if (print) { if (!native) throw new Error('Printing is available in the desktop app.'); const result = await native.print(printHtml(renumber(groups.flatMap((group) => group.pages), project.settings.pageNumbers), project.settings)); if (!result.success) { status(result.reason || 'Print canceled.'); return; } status('Document sent to the printer.'); $('#exportDialog').close(); return; }
      status('Preparing your ' + format.toUpperCase() + ' file…');
      if (groups.length > 1 || project.settings.mode === 'batch' && $('#batchPackaging').value === 'separate' && $('#exportTarget').value === 'document') {
        const zip = new JSZip();
        for (const group of groups) { const imageZip = ['png', 'jpeg'].includes(format) && group.pages.length > 1, ext = imageZip ? 'zip' : format === 'jpeg' ? 'jpg' : format; zip.file(group.name + '.' + ext, await documentBytes(group.pages, project.settings, format, native, scale)); }
        bytes = await zip.generateAsync({ type: 'uint8array' }); extension = 'zip';
      } else { const pages = groups[0].pages; extension = ['png', 'jpeg'].includes(format) && pages.length > 1 ? 'zip' : format === 'jpeg' ? 'jpg' : format; bytes = await documentBytes(pages, project.settings, format, native, scale); }
    }
    const result = await saveBytes(bytes, baseName + '.' + extension, extension); if (result) { status('File saved: ' + result.name); $('#exportDialog').close(); }
  } catch (failure) { $('#exportError').textContent = readableError(failure); $('#exportError').classList.remove('hidden'); status('Please check the message below before saving.'); }
  finally { busy = false; document.body.classList.remove('busy'); $('#runExport').disabled = false; $('#nativePrint').disabled = !native || ['project', 'template'].includes($('#exportTarget').value); renderPreview(); }
}
$('#runExport').onclick = () => runOutput(); $('#nativePrint').onclick = () => runOutput(true); $('#exportForm').onsubmit = (event) => { if (busy) event.preventDefault(); };
refresh(); saved();
async function restoreCheckpoint() {
  let restored = false;
  try { const checkpoint = await lastCheckpoint(); if (checkpoint) { resumeProject(checkpoint.serialized, checkpoint.id); restored = true; status('Your saved work is back. You can continue where you left off.'); } }
  catch (failure) { error(new Error('We could not open your saved work: ' + failure.message)); }
  finally { initializing = false; if (!restored) openStartChoices(); }
}
restoreCheckpoint().then(() => { if (native?.isTest) window.__paperlinkTest = {
  project: () => project, editors: allEditors, sample: loadSample, importDocument,
  importWorkbook: (bytes, name) => { project.importWorkbook(bytes, name); refresh(); },
  editCell: applyCell, select: (address) => sourceCellSelected(parseCell(address)), insertLink, setRecord, refresh,
  preview: () => { renderPreview(); return activeLayout; },
  serialize: () => project.serialize(documentState()),
  loadProject: (text) => { project = Project.deserialize(text); setEditorContent(project.document); refresh(); },
  outputGroups, documentBytes, spreadsheetBytes, printHtml
}; });
