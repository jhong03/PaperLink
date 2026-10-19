import test from 'node:test';
import assert from 'node:assert/strict';
import { Project, XLSX, parseRanges, parseCell } from '../core/model.mjs';

function fixture() {
  const project = new Project();
  const ws = XLSX.utils.aoa_to_sheet([['Name', 'Salary', 'Raise', 'Final'], ['Avery', 4200, 300, null], ['Jordan', 5100, 400, null]]);
  ws.D2 = { t: 'n', f: 'ROUND(B2+C2,2)', v: 4500, z: '#,##0.00' };
  ws.D3 = { t: 'n', f: 'B3+C3', v: 5500 };
  project.workbook = { SheetNames: ['People'], Sheets: { People: ws } };
  Object.assign(project.settings, { sheet: 'People', batchSheet: 'People', activeRow: 1 });
  return project;
}
test('linked edits change the same cell and dependent salary formula', () => {
  const project = fixture(), field = { sheet: 'People', address: 'C2', followRecord: true };
  project.editField(field, '625');
  assert.equal(project.cell('People', 'C2').v, 625);
  assert.equal(project.cell('People', 'D2').v, 4825);
  assert.equal(project.fieldValue({ sheet: 'People', address: 'D2' }), '4,825.00');
  assert.equal(project.formulaError, '');
  project.setCell('People', 'C2', '750');
  assert.equal(project.fieldValue(field), '750');
});
test('batch previews change their resolved row without modifying the other record', () => {
  const project = fixture(), field = { sheet: 'People', address: 'A2', followRecord: true };
  project.settings.mode = 'batch'; project.settings.activeRow = 2;
  assert.equal(project.fieldValue(field), 'Jordan');
  project.editField(field, 'Jordan revised');
  assert.equal(project.cell('People', 'A2').v, 'Avery');
  assert.equal(project.cell('People', 'A3').v, 'Jordan revised');
  assert.equal(project.fieldValue({ ...field, followRecord: false }), 'Avery');
  assert.deepEqual(project.recordRows('2-3'), [1, 2]);
});
test('formula cells are protected from direct document replacement', () => {
  const project = fixture();
  assert.throws(() => project.editField({ sheet: 'People', address: 'D2' }, '9999'), /formula/);
  assert.equal(project.cell('People', 'D2').f, 'ROUND(B2+C2,2)');
  project.setCell('People', 'C2', '-100');
  assert.equal(project.cell('People', 'D2').v, 4100);
});
test('clearing and retyping a linked numeric field keeps its numeric type', () => {
  const project = fixture(), field = { sheet: 'People', address: 'C2' };
  project.editField(field, '');
  project.editField(field, '6');
  project.editField(field, '650');
  assert.equal(project.cell('People', 'C2').t, 'n');
  assert.equal(project.cell('People', 'C2').v, 650);
  assert.equal(project.cell('People', 'D2').v, 4850);
});
test('unsupported formulas produce an explicit calculation error', () => {
  const project = fixture(); project.setCell('People', 'D2', '=UNSUPPORTED_FUNCTION(B2)');
  assert.match(project.formulaError, /calculation needs attention/);
});
test('workbook import/export retains formulas, numeric values and leading-zero identifiers', () => {
  const project = fixture(); project.setCell('People', 'E2', '00124');
  const bytes = XLSX.write(project.workbook, { type: 'array', bookType: 'xlsx' });
  const imported = new Project(); imported.importWorkbook(bytes, 'sample.xlsx');
  assert.equal(imported.cell('People', 'D2').f, 'ROUND(B2+C2,2)');
  assert.equal(imported.cell('People', 'E2').v, '00124');
  assert.equal(imported.cell('People', 'B2').v, 4200);
});
test('projects retain mappings and preserve missing sources so users can repair them', () => {
  const project = fixture(), document = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'linkedField', attrs: { id: 'test', sheet: 'People', address: 'A2', followRecord: true } }] }] };
  const restored = Project.deserialize(project.serialize(document));
  assert.equal(restored.document.content[0].content[0].attrs.address, 'A2');
  assert.equal(restored.fieldValue(restored.document.content[0].content[0].attrs), 'Avery');
  document.content[0].content[0].attrs.sheet = 'Missing';
  assert.equal(Project.deserialize(project.serialize(document)).fieldValue(document.content[0].content[0].attrs), '[Missing source]');
  document.content[0].content[0].attrs.address = 'A0';
  assert.throws(() => Project.deserialize(project.serialize(document)), /cell address/);
});
test('page and record ranges support open endings and reject partial invalid input', () => {
  assert.deepEqual(parseRanges('1-3, 5, 7-', 9), [1, 2, 3, 5, 7, 8, 9]);
  assert.deepEqual(parseRanges('all', 4, 2), [2, 3, 4]);
  for (const input of ['0', '2-1', '2,x', '1-11', '-3']) assert.throws(() => parseRanges(input, 10));
  assert.throws(() => parseCell('XFE1'));
  assert.throws(() => parseCell('A0'));
});
