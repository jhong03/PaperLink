import { Project, XLSX, clone, parseCell } from './model.mjs';

const settingKeys = ['mode', 'sheet', 'batchSheet', 'headerRow', 'paper', 'landscape', 'margin', 'fontFamily', 'fontSize', 'pageNumbers', 'headerDocument', 'footerDocument'];
const emptyDocument = () => ({ type: 'doc', content: [{ type: 'paragraph' }] });
function cleanDocument(document, fields, depth = 0) {
  if (!document || typeof document.type !== 'string' || depth > 100) throw new Error('Invalid template document structure.');
  const { content, ...properties } = document;
  const result = clone(properties);
  if (result.type === 'linkedField') {
    const attrs = result.attrs || {};
    parseCell(attrs.address);
    if (typeof attrs.sheet !== 'string' || !attrs.sheet || typeof attrs.id !== 'string' || !/^[a-z0-9_-]{1,128}$/i.test(attrs.id)) throw new Error('Invalid linked field in template.');
    // Labels in older templates may contain a value copied from a data row.
    // Store and restore the cell location instead of keeping that old text.
    const address = attrs.address.toUpperCase();
    result.attrs = { id: attrs.id, sheet: attrs.sheet, address, label: attrs.sheet + '!' + address, followRecord: attrs.followRecord !== false };
    delete result.content;
    fields.push(result.attrs);
  } else if (content) {
    if (!Array.isArray(content)) throw new Error('Invalid template document content.');
    result.content = content.map((node) => cleanDocument(node, fields, depth + 1));
  }
  return result;
}
function templateSettings(settings, fields) {
  const defaults = new Project().settings, result = {};
  for (const key of settingKeys) if (Object.hasOwn(settings, key)) result[key] = clone(settings[key]);
  for (const key of ['headerDocument', 'footerDocument']) result[key] = cleanDocument(result[key] || emptyDocument(), fields);
  return { ...defaults, ...result, activeRow: (result.headerRow ?? defaults.headerRow) + 1 };
}
function validateSheetNames(sheets) {
  if (!Array.isArray(sheets) || !sheets.length || sheets.length > 500) throw new Error('Invalid template worksheets.');
  const names = new Set();
  for (const sheet of sheets) {
    if (typeof sheet.name !== 'string' || !sheet.name.trim() || sheet.name.length > 31 || /[\\/*?:[\]]/.test(sheet.name) || names.has(sheet.name.toLowerCase())) throw new Error('Invalid or duplicate template worksheet name.');
    names.add(sheet.name.toLowerCase());
    if (!Array.isArray(sheet.headers) || sheet.headers.length > 16384) throw new Error('Invalid template column headings.');
    const addresses = new Set();
    for (const header of sheet.headers) {
      parseCell(header.address);
      if (typeof header.value !== 'string' || addresses.has(header.address.toUpperCase())) throw new Error('Invalid or duplicate template column heading.');
      addresses.add(header.address.toUpperCase());
    }
  }
}
export function blankTemplateWorkbook(template) {
  const workbook = { SheetNames: [], Sheets: Object.create(null) };
  for (const sheet of template.sheets) {
    const ws = {};
    for (const header of sheet.headers) if (header.value) XLSX.utils.sheet_add_aoa(ws, [[header.value]], { origin: header.address });
    ws['!ref'] ||= 'A1'; workbook.SheetNames.push(sheet.name); workbook.Sheets[sheet.name] = ws;
  }
  return workbook;
}
export function deserializeTemplate(text) {
  const data = JSON.parse(text);
  if (!['paperlink-template', 'flowdesk-template'].includes(data.format) || data.version !== 1 || typeof data.name !== 'string' || !data.name.trim() || data.name.length > 120 || data.document?.type !== 'doc') throw new Error('This is not a supported PaperLink template.');
  validateSheetNames(data.sheets);
  const fields = [], document = cleanDocument(data.document, fields), settings = templateSettings(data.settings || {}, fields);
  const sheets = data.sheets.map((sheet) => ({ name: sheet.name, headers: sheet.headers.map((header) => ({ address: header.address.toUpperCase(), value: header.value })) }));
  const names = new Set(sheets.map((sheet) => sheet.name));
  if (!names.has(settings.sheet) || !names.has(settings.batchSheet) || fields.some((field) => !names.has(field.sheet))) throw new Error('A template refers to a worksheet that is not defined.');
  if (sheets.some((sheet) => sheet.headers.some((header) => parseCell(header.address).r !== settings.headerRow))) throw new Error('Template column headings must be on the configured header row.');
  const template = { format: 'paperlink-template', version: 1, name: data.name.trim(), document, settings, sheets };
  // Reuse project validation for embedded images, addresses, mode, and record settings.
  Project.deserialize(JSON.stringify({ format: 'paperlink-project', version: 1, workbook: blankTemplateWorkbook(template), document, settings }));
  return template;
}
export function createTemplate(project, document, name) {
  if (typeof name !== 'string' || !name.trim() || name.length > 120) throw new Error('Give your template a name, using 120 characters or fewer.');
  const fields = [], cleaned = cleanDocument(document, fields), settings = templateSettings(project.settings, fields);
  const required = new Set([settings.sheet, settings.batchSheet, ...fields.map((field) => field.sheet)]);
  const sheets = [...required].map((sheet) => {
    project.sheet(sheet);
    const columns = new Set(fields.filter((field) => field.sheet === sheet && parseCell(field.address).r > settings.headerRow && (settings.mode !== 'batch' || sheet === settings.batchSheet)).map((field) => parseCell(field.address).c));
    return { name: sheet, headers: [...columns].sort((a, b) => a - b).map((c) => {
      const address = XLSX.utils.encode_cell({ c, r: settings.headerRow });
      return { address, value: project.display(sheet, address) };
    }) };
  });
  return deserializeTemplate(JSON.stringify({ format: 'paperlink-template', version: 1, name, document: cleaned, settings, sheets }));
}
export function templateFieldCount(template) {
  const fields = [];
  for (const doc of [template.document, template.settings.headerDocument, template.settings.footerDocument]) cleanDocument(doc, fields);
  return new Set(fields.map((field) => field.id)).size;
}
const normalizeHeading = (value) => String(value).trim().replace(/\s+/g, ' ').toLowerCase();
export function templateCompatibility(template, project, worksheetMap) {
  const issues = [];
  for (const sheet of template.sheets) {
    const target = worksheetMap[sheet.name];
    if (!target || !project.workbook.SheetNames.includes(target)) { issues.push('Choose which sheet to use for “' + sheet.name + '”.'); continue; }
    for (const header of sheet.headers) {
      if (!header.value) continue;
      const actual = project.display(target, header.address);
      if (normalizeHeading(actual) !== normalizeHeading(header.value)) issues.push('In “' + target + '”, ' + header.address + ' should say “' + header.value + '”. It currently says “' + (actual || '(blank)') + '”.');
    }
  }
  return issues;
}
export function suggestTemplateWorksheets(template, project) {
  return Object.fromEntries(template.sheets.map((sheet) => {
    if (project.workbook.SheetNames.includes(sheet.name)) return [sheet.name, sheet.name];
    const matches = project.workbook.SheetNames.filter((name) => !templateCompatibility({ sheets: [sheet] }, project, { [sheet.name]: name }).length);
    return [sheet.name, matches.length === 1 ? matches[0] : ''];
  }));
}
function remapDocument(document, worksheetMap) {
  const result = clone(document);
  const visit = (node) => {
    if (node.type === 'linkedField') { node.attrs.sheet = worksheetMap[node.attrs.sheet]; node.attrs.label = node.attrs.sheet + '!' + node.attrs.address; node.attrs.value = ''; }
    for (const child of node.content || []) visit(child);
  };
  visit(result); return result;
}
export function applyTemplate(template, currentProject = null, worksheetMap = null) {
  template = deserializeTemplate(JSON.stringify(template));
  const map = worksheetMap || Object.fromEntries(template.sheets.map((sheet) => [sheet.name, sheet.name]));
  if (currentProject) {
    const issues = templateCompatibility(template, currentProject, map);
    if (issues.length) throw new Error('The workbook does not match this template. ' + issues[0]);
  }
  const project = new Project();
  project.workbook = currentProject ? currentProject.workbook : blankTemplateWorkbook(template);
  if (currentProject) project.recordIndex = currentProject.recordIndex;
  project.settings = { ...project.settings, ...clone(template.settings), sheet: map[template.settings.sheet], batchSheet: map[template.settings.batchSheet] };
  project.document = remapDocument(template.document, map);
  project.settings.headerDocument = remapDocument(template.settings.headerDocument, map);
  project.settings.footerDocument = remapDocument(template.settings.footerDocument, map);
  project.names = { ...project.names, ...(currentProject?.names || {}), document: template.name + ' (template)', template: template.name };
  if (!currentProject || ['Untitled work', 'Untitled project'].includes(project.names.project)) project.names.project = template.name;
  project.formulaError = currentProject?.formulaError || '';
  if (project.settings.mode === 'batch') project.settings.activeRow = project.recordRows()[0] ?? project.settings.headerRow + 1;
  return project;
}
export function projectForTemplateEditing(template, reference = null) {
  const draft = applyTemplate(template), missingSheets = [];
  if (!reference) return { project: draft, missingSheets };
  const map = suggestTemplateWorksheets(template, reference), workbook = clone(reference.workbook);
  for (const sheet of template.sheets) {
    if (map[sheet.name]) continue;
    // Keep unresolved connections separate instead of assigning an arbitrary
    // sheet's values. The imported sheets remain available to reconnect cells.
    missingSheets.push(sheet.name); map[sheet.name] = sheet.name;
    workbook.SheetNames.push(sheet.name); workbook.Sheets[sheet.name] = draft.workbook.Sheets[sheet.name];
  }
  draft.workbook = workbook;
  draft.settings.sheet = map[template.settings.sheet]; draft.settings.batchSheet = map[template.settings.batchSheet];
  draft.document = remapDocument(template.document, map);
  draft.settings.headerDocument = remapDocument(template.settings.headerDocument, map);
  draft.settings.footerDocument = remapDocument(template.settings.footerDocument, map);
  draft.names.spreadsheet = reference.names.spreadsheet; draft.formulaError = reference.formulaError;
  if (missingSheets.includes(template.settings.sheet)) draft.settings.sheet = reference.settings.sheet;
  if (draft.settings.mode === 'batch') {
    const rows = draft.recordRows();
    draft.settings.activeRow = reference.settings.batchSheet === draft.settings.batchSheet && rows.includes(reference.settings.activeRow) ? reference.settings.activeRow : rows[0] ?? draft.settings.headerRow + 1;
  }
  return { project: draft, missingSheets };
}
