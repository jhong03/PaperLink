import * as XLSX from 'xlsx';
import XLSX_CALC from 'xlsx-calc';
import * as functions from '@formulajs/formulajs';

XLSX_CALC.import_functions(functions);
export { XLSX };
export const clone = (value) => structuredClone(value);
export const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const colName = (col) => XLSX.utils.encode_col(col);
export const parseCell = (address) => {
  if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/i.test(address || '')) throw new Error('Use a cell address such as B2.');
  const cell = XLSX.utils.decode_cell(address.toUpperCase());
  if (cell.c > 16383 || cell.r > 1048575) throw new Error('The cell address is outside spreadsheet limits.');
  return cell;
};
function rangeIntervals(input, maximum, minimum = 1) {
  if (!Number.isInteger(maximum) || maximum < minimum) throw new Error('There is nothing to export.');
  const text = String(input || '').trim().toLowerCase();
  if (!text || text === 'all') return [[minimum, maximum]];
  const intervals = [];
  for (const token of text.split(',')) {
    const match = token.trim().match(/^(\d+)(?:\s*-\s*(\d*))?$/);
    if (!match) throw new Error('Use ranges such as 1-3, 5, 8- or leave blank for all.');
    const start = Number(match[1]), end = match[2] === undefined ? start : match[2] === '' ? maximum : Number(match[2]);
    if (start < minimum || end > maximum || start > end) throw new Error('The range must be between ' + minimum + ' and ' + maximum + '.');
    intervals.push([start, end]);
  }
  intervals.sort((a, b) => a[0] - b[0]);
  const merged = [];
  for (const interval of intervals) {
    const previous = merged.at(-1);
    if (previous && interval[0] <= previous[1] + 1) previous[1] = Math.max(previous[1], interval[1]);
    else merged.push(interval);
  }
  return merged;
}
export function parseRanges(input, maximum, minimum = 1) {
  const result = [];
  for (const [start, end] of rangeIntervals(input, maximum, minimum)) for (let row = start; row <= end; row++) result.push(row);
  return result;
}
export class Project {
  constructor() {
    this.workbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1' } } };
    this.settings = { mode: 'single', sheet: 'Sheet1', batchSheet: 'Sheet1', headerRow: 0, activeRow: 1, paper: 'A4', landscape: false, margin: 20, fontFamily: 'Arial', fontSize: 11, header: '', footer: '', pageNumbers: true };
    this.names = { spreadsheet: 'Blank spreadsheet', document: 'Blank document', project: 'Untitled work' };
    this.document = { type: 'doc', content: [{ type: 'paragraph' }] };
    this.formulaError = '';
    this.recordIndex = new WeakMap();
  }
  importWorkbook(bytes, name) {
    const workbook = XLSX.read(bytes, { type: 'array', cellDates: false, cellStyles: true, cellNF: true, sheetStubs: false });
    if (!workbook.SheetNames?.length) throw new Error('This file has no worksheets.');
    this.workbook = workbook;
    if (this.settings.emailDraft) delete this.settings.emailDraft.recipientOverrides;
    this.settings.sheet = workbook.SheetNames[0];
    this.settings.batchSheet = workbook.SheetNames[0];
    this.settings.headerRow = 0;
    this.settings.activeRow = 1;
    this.names.spreadsheet = name;
    this.formulaError = '';
  }
  sheet(name = this.settings.sheet) {
    const sheet = Object.hasOwn(this.workbook.Sheets, name) ? this.workbook.Sheets[name] : null;
    if (!sheet) throw new Error('Worksheet "' + name + '" no longer exists.');
    return sheet;
  }
  bounds(name = this.settings.sheet) { return XLSX.utils.decode_range(this.sheet(name)['!ref'] || 'A1'); }
  cell(sheet, address) { parseCell(address); return this.sheet(sheet)[address.toUpperCase()] || null; }
  raw(sheet, address) {
    const cell = this.cell(sheet, address);
    return cell?.f ? '=' + cell.f : String(cell?.v ?? '');
  }
  display(sheet, address) {
    const cell = this.cell(sheet, address);
    if (!cell) return '';
    try { return XLSX.utils.format_cell({ ...cell, w: undefined }); } catch { return String(cell.v ?? ''); }
  }
  setCell(sheetName, address, input, options = {}) {
    address = address.toUpperCase();
    const coordinate = parseCell(address), sheet = this.sheet(sheetName), previous = sheet[address];
    if (options.fromDocument && previous?.f) throw new Error('This value comes from a formula. Edit the formula or its input cells in the spreadsheet.');
    const value = String(input ?? '');
    const cell = { ...(previous || {}) };
    const originalType = ['n', 's'].includes(previous?.flowValueType) ? previous.flowValueType : previous?.t;
    cell.flowValueType = originalType;
    delete cell.w; delete cell.f;
    if (value.startsWith('=')) {
      if (value.length === 1) throw new Error('Enter an expression after =.');
      cell.f = value.slice(1); cell.t = 'n'; cell.v = previous?.v ?? 0;
    } else if (value.trim() === '') {
      delete cell.v; cell.t = 's'; cell.v = '';
    } else {
      const numeric = value.replace(/,/g, '');
      const isNumber = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(numeric) && !(options.fromDocument && originalType === 's');
      cell.v = isNumber ? Number(numeric) : value;
      cell.t = isNumber ? 'n' : 's';
      if (!options.fromDocument || !cell.flowValueType) cell.flowValueType = cell.t;
    }
    sheet[address] = cell;
    this.recordIndex.delete(sheet);
    const bounds = this.bounds(sheetName);
    sheet['!ref'] = XLSX.utils.encode_range({ s: { r: Math.min(bounds.s.r, coordinate.r), c: Math.min(bounds.s.c, coordinate.c) }, e: { r: Math.max(bounds.e.r, coordinate.r), c: Math.max(bounds.e.c, coordinate.c) } });
    if (options.recalculate !== false) this.recalculate();
    return cell;
  }
  recalculate() {
    // Formula results can add or remove populated rows on any worksheet.
    this.recordIndex = new WeakMap();
    this.formulaError = '';
    try { XLSX_CALC(this.workbook, { continue_after_error: false, log_error: false }); }
    catch (error) { this.formulaError = 'Formula calculation needs attention: ' + (error.message || String(error)) + '. Update the formula before document export.'; }
    for (const name of this.workbook.SheetNames) for (const [address, cell] of Object.entries(this.sheet(name))) {
      if (!address.startsWith('!') && cell?.f) {
        delete cell.w;
        if (cell.t === 'e' || typeof cell.v === 'number' && !Number.isFinite(cell.v)) this.formulaError ||= 'A formula returned an error. Correct it before document export.';
      }
    }
  }
  source(field, recordRow = this.settings.activeRow) {
    const coordinate = parseCell(field.address);
    const follows = this.settings.mode === 'batch' && field.followRecord !== false && field.sheet === this.settings.batchSheet;
    return { sheet: field.sheet, address: XLSX.utils.encode_cell({ c: coordinate.c, r: follows ? recordRow : coordinate.r }) };
  }
  fieldValue(field, recordRow) {
    try { const source = this.source(field, recordRow); return this.display(source.sheet, source.address); } catch { return '[Missing source]'; }
  }
  editField(field, value) {
    const source = this.source(field);
    return this.setCell(source.sheet, source.address, value, { fromDocument: true });
  }
  recordRows(input = '') {
    const sheet = this.settings.batchSheet, ws = this.sheet(sheet), end = this.bounds(sheet).e.r + 1, first = this.settings.headerRow + 2;
    const all = !String(input || '').trim() || String(input).trim().toLowerCase() === 'all';
    if (end < first && all) return [];
    const intervals = rangeIntervals(input, end, first);
    let index = this.recordIndex.get(ws);
    if (!index) {
      // Inspect each stored cell once, regardless of empty or formatted rows in !ref.
      // Avoid Object.entries: it allocates an address/value pair for every cell.
      const populated = new Set();
      for (const address in ws) {
        if (!Object.hasOwn(ws, address) || address.startsWith('!')) continue;
        const cell = ws[address];
        if (cell?.v === undefined || String(cell.v).trim() === '') continue;
        const match = address.match(/^[A-Z]{1,3}([1-9]\d{0,6})$/i);
        if (match) populated.add(Number(match[1]) - 1);
      }
      index = { rows: [...populated].sort((a, b) => a - b) };
      this.recordIndex.set(ws, index);
    }
    if (all) {
      if (index.first !== first || index.end !== end) {
        index.selected = index.rows.filter((row) => row + 1 >= first && row + 1 <= end);
        index.first = first; index.end = end;
      }
      return index.selected;
    }
    // Match actual populated rows to ranges without expanding a million empty rows.
    const selected = []; let interval = 0;
    for (const row of index.rows) {
      while (interval < intervals.length && row + 1 > intervals[interval][1]) interval++;
      if (interval === intervals.length) break;
      if (row + 1 >= intervals[interval][0]) selected.push(row);
    }
    return selected;
  }
  addSheet(name) {
    name = String(name || '').trim();
    if (!name || /[\\/*?:[\]]/.test(name) || name.length > 31) throw new Error('Use a worksheet name of 1–31 characters without / \\ * ? : [ ].');
    if (this.workbook.SheetNames.some((item) => item.toLowerCase() === name.toLowerCase())) throw new Error('A worksheet with this name already exists.');
    this.workbook.SheetNames.push(name); this.workbook.Sheets[name] = { '!ref': 'A1' }; this.settings.sheet = name;
  }
  serialize(document) {
    return JSON.stringify({ format: 'paperlink-project', version: 1, workbook: this.workbook, settings: this.settings, names: this.names, document: document || this.document, formulaError: this.formulaError });
  }
  static deserialize(text) {
    const data = JSON.parse(text);
    if (!['paperlink-project', 'flowdesk-project'].includes(data.format) || data.version !== 1 || !Array.isArray(data.workbook?.SheetNames) || !data.workbook?.Sheets || data.document?.type !== 'doc') throw new Error('This is not a supported PaperLink project.');
    if (!data.workbook.SheetNames.length || data.workbook.SheetNames.length > 500) throw new Error('Invalid worksheet count.');
    for (const name of data.workbook.SheetNames) {
      if (!Object.hasOwn(data.workbook.Sheets, name)) throw new Error('The project is missing a worksheet.');
      const ws = data.workbook.Sheets[name];
      for (const [address, cell] of Object.entries(ws)) {
        if (address.startsWith('!')) continue;
        parseCell(address);
        if (!cell || typeof cell !== 'object' || !['string', 'number', 'boolean', 'undefined'].includes(typeof cell.v)) throw new Error('Unsupported cell value in project.');
      }
    }
    const validateDoc = (node) => {
      if (node.type === 'image' && !/^data:image\/(png|jpeg|gif|webp|bmp);base64,/i.test(node.attrs?.src || '')) throw new Error('Project images must be embedded.');
      if (node.type === 'linkedField') {
        parseCell(node.attrs?.address);
        if (typeof node.attrs?.sheet !== 'string' || !node.attrs.sheet || node.attrs.sheet.length > 128) throw new Error('Invalid linked worksheet name.');
      }
      for (const child of node.content || []) validateDoc(child);
    };
    validateDoc(data.document);
    if (data.settings?.headerDocument) validateDoc(data.settings.headerDocument);
    if (data.settings?.footerDocument) validateDoc(data.settings.footerDocument);
    const project = new Project();
    project.workbook = data.workbook;
    project.settings = { ...project.settings, ...data.settings };
    if (!['single', 'batch'].includes(project.settings.mode)) throw new Error('Invalid project mode.');
    if (!Number.isInteger(project.settings.headerRow) || project.settings.headerRow < 0 || project.settings.headerRow > 1048574 || !Number.isInteger(project.settings.activeRow) || project.settings.activeRow < 0 || project.settings.activeRow > 1048575) throw new Error('Invalid record row in project.');
    if (!project.workbook.SheetNames.includes(project.settings.sheet)) project.settings.sheet = project.workbook.SheetNames[0];
    if (!project.workbook.SheetNames.includes(project.settings.batchSheet)) project.settings.batchSheet = project.workbook.SheetNames[0];
    project.names = { ...project.names, ...data.names };
    project.document = data.document;
    project.formulaError = String(data.formulaError || '');
    return project;
  }
}
