import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { fileURLToPath } from 'node:url';
import { Project, XLSX } from '../core/model.mjs';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const resultDir = path.join(root, 'test-results'); await fs.mkdir(resultDir, { recursive: true });
const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE;
const app = await electron.launch({ executablePath: path.join(root, 'node_modules/electron/dist/electron.exe'), args: [root, '--smoke'], cwd: root, env, timeout: 30000 });
let page;
const failures = [];
async function capture(name) {
  const png = await app.evaluate(async ({ BrowserWindow }) => {
    const image = await BrowserWindow.getAllWindows()[0].webContents.capturePage(undefined, { stayHidden: true, stayAwake: true });
    return [...image.toPNG()];
  });
  await fs.writeFile(path.join(resultDir, name), new Uint8Array(png));
}
try {
  page = await app.firstWindow();
  page.on('pageerror', (error) => failures.push(error.message));
  page.on('dialog', (dialog) => dialog.accept());
  await page.waitForFunction(() => Boolean(window.__paperlinkTest), { timeout: 20000 });
  await page.getByRole('button', { name: 'Try sample project' }).click();
  await page.waitForFunction(() => document.querySelector('#mappingCount').textContent === '6');
  await page.locator('td[data-address="D2"]').dblclick();
  await page.getByLabel('D2 value', { exact: true }).fill('650');
  await page.getByLabel('D2 value', { exact: true }).press('Enter');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'E2').v), 4850);
  await page.getByLabel('Increment linked value', { exact: true }).fill('');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'D2').v), '');
  await page.getByLabel('Increment linked value', { exact: true }).pressSequentially('675', { delay: 20 });
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'D2').t), 'n');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'E2').v), 4875);
  await page.locator('#documentEditor .mapped-field input').first().fill('Avery revised');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'A2').v), 'Avery revised');
  await page.locator('#closeDrawer').click();
  await page.locator('#nextRecord').click();
  assert.equal(await page.locator('#documentEditor .mapped-field input').first().inputValue(), 'Jordan Lee');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'A2').v), 'Avery revised');
  await page.locator('#previewTab').click();
  await page.waitForFunction(() => document.querySelector('#pagePreview').innerText.includes('Jordan Lee'));
  await capture('workspace.png');
  const serialized = await page.evaluate(() => window.__paperlinkTest.serialize());
  await fs.writeFile(path.join(resultDir, 'sample.paperlink'), serialized);
  await page.evaluate((text) => window.__paperlinkTest.loadProject(text), serialized);
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Employees', 'A2').v), 'Avery revised');

  await page.locator('[data-action="export"]').click();
  await page.locator('#exportRecords').fill('2-4');
  await page.locator('#exportFormat').selectOption('docx');
  const docxPath = path.join(resultDir, 'batch.docx');
  await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, docxPath);
  await page.locator('#runExport').click();
  await page.waitForFunction(() => !document.querySelector('#exportDialog').open || !document.querySelector('#exportError').classList.contains('hidden'));
  assert.equal(await page.locator('#exportError').isVisible(), false, await page.locator('#exportError').textContent());
  const zip = await JSZip.loadAsync(await fs.readFile(docxPath)), xml = await zip.file('word/document.xml').async('string');
  for (const name of ['Avery revised', 'Jordan Lee', 'Morgan Lim']) assert.ok(xml.includes(name));
  assert.equal((xml.match(/<w:sectPr/g) || []).length, 3);

  await page.locator('[data-action="export"]').click();
  await page.locator('#exportFormat').selectOption('pdf');
  const pdfPath = path.join(resultDir, 'batch.pdf');
  await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, pdfPath);
  await page.locator('#runExport').click();
  await page.waitForFunction(() => !document.querySelector('#exportDialog').open || !document.querySelector('#exportError').classList.contains('hidden'));
  assert.equal(await page.locator('#exportError').isVisible(), false, await page.locator('#exportError').textContent());
  const pdf = await fs.readFile(pdfPath);
  assert.equal(pdf.subarray(0, 4).toString(), '%PDF');
  assert.equal((pdf.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length, 3);

  const imports = await page.evaluate(async (data) => {
    const t = window.__paperlinkTest, imported = await t.importDocument(new Uint8Array(data), 'batch.docx');
    return { html: imported.html, header: imported.header, footer: imported.footer };
  }, [...await fs.readFile(docxPath)]);
  assert.ok(imports.html.includes('Avery revised')); assert.ok(imports.header.includes('EXAMPLE COMPANY'));
  assert.ok(imports.footer.includes('Fictional sample'));

  await page.locator('[data-action="export"]').click();
  await page.locator('#exportFormat').selectOption('docx');
  await page.locator('#batchPackaging').selectOption('separate');
  const separatePath = path.join(resultDir, 'separate.zip');
  await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, separatePath);
  await page.locator('#runExport').click();
  await page.waitForFunction(() => !document.querySelector('#exportDialog').open || !document.querySelector('#exportError').classList.contains('hidden'));
  assert.equal(await page.locator('#exportError').isVisible(), false, await page.locator('#exportError').textContent());
  const separate = await JSZip.loadAsync(await fs.readFile(separatePath));
  assert.deepEqual(Object.keys(separate.files), ['record-row-2.docx', 'record-row-3.docx', 'record-row-4.docx']);

  await page.locator('[data-action="export"]').click();
  await page.locator('#batchPackaging').selectOption('combined');
  await page.locator('#exportPages').selectOption('current');
  const current = await page.evaluate(() => window.__paperlinkTest.outputGroups().map((group) => group.pages.map((p) => p.querySelector('.page-body').textContent)));
  assert.equal(current.length, 1); assert.equal(current[0].length, 1);
  assert.ok(current[0][0].includes('Jordan Lee')); assert.ok(!current[0][0].includes('Avery revised'));
  await page.locator('#exportDialog').getByRole('button', { name: 'Cancel', exact: true }).click();

  await page.locator('[data-action="export"]').click(); await page.locator('#exportFormat').selectOption('png');
  await page.locator('#exportRecords').fill('2'); await page.locator('#imageQuality').selectOption('1');
  const pngPath = path.join(resultDir, 'page.png'); await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, pngPath);
  await page.locator('#runExport').click();
  await page.waitForFunction(() => !document.querySelector('#exportDialog').open || !document.querySelector('#exportError').classList.contains('hidden'), { timeout: 30000 });
  assert.equal(await page.locator('#exportError').isVisible(), false, await page.locator('#exportError').textContent());
  const png = await fs.readFile(pngPath); assert.equal(png.subarray(1, 4).toString(), 'PNG');

  await page.locator('[data-action="export"]').click(); await page.locator('#exportFormat').selectOption('jpeg');
  await page.locator('#exportRecords').fill('3'); await page.locator('#imageQuality').selectOption('1');
  const jpgPath = path.join(resultDir, 'page.jpg'); await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, jpgPath);
  await page.locator('#runExport').click();
  await page.waitForFunction(() => !document.querySelector('#exportDialog').open || !document.querySelector('#exportError').classList.contains('hidden'));
  assert.equal(await page.locator('#exportError').isVisible(), false, await page.locator('#exportError').textContent());
  const jpg = await fs.readFile(jpgPath); assert.deepEqual([...jpg.subarray(0, 2)], [255, 216]);

  for (const format of ['odt', 'rtf', 'html', 'txt']) {
    const data = [...await fs.readFile(path.join(root, 'samples', 'salary-template.' + format))];
    const imported = await page.evaluate(async ({ data, format }) => (await window.__paperlinkTest.importDocument(new Uint8Array(data), 'template.' + format)).html, { data, format });
    assert.ok(imported.includes('Salary adjustment notice'), format);
    assert.ok(imported.includes('[Employee name]'), format);
  }
  for (const format of ['xlsx', 'ods', 'csv']) {
    const data = [...await fs.readFile(path.join(root, 'samples', 'payroll-fictional.' + format))];
    const imported = await page.evaluate(({ data, format }) => { const t = window.__paperlinkTest; t.importWorkbook(new Uint8Array(data), 'data.' + format); const p = t.project(); return p.cell(p.settings.sheet, 'A2').v; }, { data, format });
    assert.equal(imported, 'Avery Tan');
  }
  const sheetOutput = await page.evaluate(() => [...window.__paperlinkTest.spreadsheetBytes(window.__paperlinkTest.project(), 'ods')]);
  const ods = XLSX.read(new Uint8Array(sheetOutput), { type: 'array' }); assert.equal(ods.Sheets[ods.SheetNames[0]].A2.v, 'Avery Tan');
  const safeHtml = await page.evaluate(async () => (await window.__paperlinkTest.importDocument(new TextEncoder().encode('<p>Safe text</p><script>window.bad=true</script><img src="https://example.com/a.png">'), 'input.html')).html);
  assert.ok(safeHtml.includes('Safe text')); assert.ok(!safeHtml.includes('<script')); assert.ok(!safeHtml.includes('example.com'));

  // Exercise the document-to-blank-spreadsheet direction through the actual mapping dialog.
  await page.evaluate((text) => { const t = window.__paperlinkTest; t.loadProject(text); const editor = t.editors()[0]; editor.commands.setContent('<p>Corporate contact</p>'); editor.commands.setTextSelection({ from: 1, to: 18 }); t.select('B2'); }, new Project().serialize());
  await page.locator('#editorTab').click(); await page.locator('#insertField').click();
  await page.locator('#mappingDialog [name="anchor"][value="document"]').check();
  await page.locator('#mappingDialog button[value="connect"]').click();
  await page.waitForFunction(() => window.__paperlinkTest.project().cell('Sheet1', 'B2')?.v === 'Corporate contact');
  await page.locator('#documentEditor .mapped-field input').fill('Corporate contact updated');
  assert.equal(await page.evaluate(() => window.__paperlinkTest.project().cell('Sheet1', 'B2').v), 'Corporate contact updated');
  await page.locator('#closeDrawer').click();

  const pagination = await page.evaluate(() => {
    const t = window.__paperlinkTest, editor = t.editors()[0];
    const text = Array.from({ length: 1800 }, (_, i) => 'word' + i).join(' ');
    editor.commands.setContent('<p><strong>' + text + '</strong></p>');
    const layout = t.preview();
    return { pages: layout.pages.length, errors: layout.errors, actual: layout.pages.map((p) => p.querySelector('.page-body').textContent).join(''), expected: text };
  });
  assert.ok(pagination.pages > 1); assert.deepEqual(pagination.errors, []); assert.equal(pagination.actual, pagination.expected);
  const listImport = await page.evaluate(async () => {
    const t = window.__paperlinkTest; t.editors()[0].commands.setContent('<ol><li><p>First corporate item</p></li><li><p>Second corporate item</p></li></ol>');
    const layout = t.preview(), bytes = await t.documentBytes(layout.pages, t.project().settings, 'docx', window.paperlink);
    return (await t.importDocument(bytes, 'list.docx')).html;
  });
  assert.ok(listImport.includes('<ol')); assert.ok(listImport.includes('Second corporate item'));
  assert.deepEqual(failures, []);
  console.log('Desktop smoke passed: two-way edits, blank-side mapping, formulas, batch preview/ZIP/current-page output, project reload, DOCX/PDF/PNG/JPEG/ODS exports, supported document and spreadsheet imports, HTML sanitizing, and paragraph pagination.');
} catch (error) {
  if (page) { await capture('failure.png').catch(() => {}); console.error(await page.locator('#exportError').textContent().catch(() => '')); }
  console.error('Renderer errors:', failures); throw error;
} finally { await app.close(); }
