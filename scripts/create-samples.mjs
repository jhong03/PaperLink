import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import * as D from 'docx';
import { XLSX } from '../core/model.mjs';
const dir = path.resolve('samples'); await fs.mkdir(dir, { recursive: true });
const rows = [
  ['Employee name', 'Department', 'Current salary', 'Increment', 'Final salary', 'Effective date'],
  ['Avery Tan', 'Operations', 4200, 300, null, '1 January 2027'],
  ['Jordan Lee', 'Finance', 5100, 400, null, '1 January 2027'],
  ['Morgan Lim', 'People', 3900, 250, null, '1 January 2027']
];
const worksheet = XLSX.utils.aoa_to_sheet(rows);
worksheet['!cols'] = [{ wch: 20 }, { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 16 }, { wch: 22 }];
for (let r = 2; r <= 4; r++) {
  worksheet['E' + r] = { t: 'n', f: 'C' + r + '+D' + r, v: rows[r - 1][2] + rows[r - 1][3], z: '#,##0.00' };
  worksheet['C' + r].z = '#,##0.00'; worksheet['D' + r].z = '#,##0.00';
}
const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, worksheet, 'Employees');
for (const type of ['xlsx', 'ods']) await fs.writeFile(path.join(dir, 'payroll-fictional.' + type), new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: type })));
await fs.writeFile(path.join(dir, 'payroll-fictional.csv'), XLSX.utils.sheet_to_csv(worksheet));
const tableRows = [['Department', '[Department]'], ['Current salary', '[Current salary]'], ['Increment amount', '[Increment]'], ['Final salary', '[Final salary]']];
const doc = new D.Document({
  styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    headers: { default: new D.Header({ children: [new D.Paragraph({ children: [new D.TextRun({ text: 'EXAMPLE COMPANY', bold: true, size: 18 })] })] }) },
    footers: { default: new D.Footer({ children: [new D.Paragraph('Fictional sample · for manual testing')] }) },
    children: [
      new D.Paragraph({ text: 'Salary adjustment notice', heading: D.HeadingLevel.HEADING_1 }),
      new D.Paragraph('Dear [Employee name],'),
      new D.Paragraph('We are pleased to confirm the following adjustment to your monthly salary.'),
      new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: tableRows.map((row) => new D.TableRow({ children: row.map((value) => new D.TableCell({ children: [new D.Paragraph(value)] })) })) }),
      new D.Paragraph('Effective date: [Effective date]'),
      new D.Paragraph('Thank you for your continued contribution.'),
      new D.Paragraph('Authorized signatory: __________________________')
    ]
  }]
});
await fs.writeFile(path.join(dir, 'salary-template.docx'), await D.Packer.toBuffer(doc));
const html = '<h1>Salary adjustment notice</h1><p>Dear [Employee name],</p><p>We are pleased to confirm the following adjustment to your monthly salary.</p><table><tbody>' + tableRows.map((row) => '<tr>' + row.map((value) => '<td><p>' + value + '</p></td>').join('') + '</tr>').join('') + '</tbody></table><p>Effective date: [Effective date]</p><p>Authorized signatory: __________________________</p>';
await fs.writeFile(path.join(dir, 'salary-template.html'), '<!doctype html><html><head><meta charset="utf-8"></head><body>' + html + '</body></html>');
await fs.writeFile(path.join(dir, 'salary-template.txt'), 'Salary adjustment notice\n\nDear [Employee name],\n\nCurrent salary: [Current salary]\nFinal salary: [Final salary]\nEffective date: [Effective date]\n');
await fs.writeFile(path.join(dir, 'salary-template.rtf'), '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\f0\\fs22\\b Salary adjustment notice\\b0\\par Dear [Employee name],\\par Current salary: [Current salary]\\par Final salary: [Final salary]\\par Effective date: [Effective date]\\par}');
const zip = new JSZip();
zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
zip.file('content.xml', '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2"><office:body><office:text><text:h text:outline-level="1">Salary adjustment notice</text:h><text:p>Dear [Employee name],</text:p><table:table>' + tableRows.map((row) => '<table:table-row>' + row.map((value) => '<table:table-cell><text:p>' + value + '</text:p></table:table-cell>').join('') + '</table:table-row>').join('') + '</table:table><text:p>Effective date: [Effective date]</text:p></office:text></office:body></office:document-content>');
zip.file('styles.xml', '<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:master-styles><style:master-page style:name="Standard"><style:header><text:p>EXAMPLE COMPANY</text:p></style:header><style:footer><text:p>Fictional sample</text:p></style:footer></style:master-page></office:master-styles></office:document-styles>');
zip.file('META-INF/manifest.xml', '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>');
await fs.writeFile(path.join(dir, 'salary-template.odt'), await zip.generateAsync({ type: 'nodebuffer' }));
console.log('Fictional spreadsheet and document samples written to samples/.');
