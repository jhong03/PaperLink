import * as D from 'docx';
import JSZip from 'jszip';
import { toBlob, toJpeg } from 'html-to-image';
import { XLSX } from './model.mjs';
import { geometry, printHtml } from './layout.js';
export async function blobBytes(blob) { return new Uint8Array(await blob.arrayBuffer()); }
const parseNumber = (value, fallback) => Number.parseFloat(value) || fallback;
const halfPoints = (value, fallback = 11) => Math.round(parseNumber(value, fallback) * (/px$/.test(value || '') ? 1.5 : 2));
const hexColor = (value) => {
  if (!value) return undefined;
  if (/^#[0-9a-f]{6}$/i.test(value)) return value.slice(1);
  if (/^#[0-9a-f]{3}$/i.test(value)) return value.slice(1).split('').map((c) => c + c).join('');
  const rgb = value.match(/rgba?\((\d+)[ ,]+(\d+)[ ,]+(\d+)/);
  return rgb ? rgb.slice(1, 4).map((part) => Number(part).toString(16).padStart(2, '0')).join('') : undefined;
};
export async function docxBytes(pages, settings) {
  const g = geometry(settings);
  let listInstance = 0;
  function runs(node, inherited = {}) {
    if (node.nodeType === 3) return [new D.TextRun({ text: node.textContent, ...inherited })];
    if (node.nodeType !== 1) return [];
    const style = node.style, tag = node.tagName, options = { ...inherited };
    if (['B', 'STRONG', 'TH'].includes(tag) || style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) options.bold = true;
    if (['I', 'EM'].includes(tag) || style.fontStyle === 'italic') options.italics = true;
    if (tag === 'U' || style.textDecoration.includes('underline')) options.underline = {};
    if (['S', 'DEL'].includes(tag) || style.textDecoration.includes('line-through')) options.strike = true;
    if (style.fontFamily) options.font = style.fontFamily.replace(/["']/g, '').split(',')[0];
    if (style.fontSize) options.size = halfPoints(style.fontSize);
    if (hexColor(style.color)) options.color = hexColor(style.color);
    if (tag === 'BR') return [new D.TextRun({ break: 1, ...options })];
    if (tag === 'IMG') {
      const match = node.src.match(/^data:image\/(png|jpeg|jpg|gif|bmp);base64,(.+)$/s);
      if (!match) return [];
      const bytes = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0));
      let width = parseNumber(node.getAttribute('width') || style.width, 180), height = parseNumber(node.getAttribute('height') || style.height, 100);
      const ratio = Math.min(1, g.bodyWidth / width, g.bodyHeight / height); width *= ratio; height *= ratio;
      return [new D.ImageRun({ type: match[1] === 'jpeg' ? 'jpg' : match[1], data: bytes, transformation: { width, height } })];
    }
    const children = [...node.childNodes].flatMap((child) => runs(child, options));
    if (tag === 'A' && /^(https?:|mailto:)/.test(node.href)) return [new D.ExternalHyperlink({ link: node.href, children })];
    return children;
  }
  function paragraph(node, extra = {}, baseSize = Number(settings.fontSize) || 11, inCell = false, edge = false) {
    const tag = node.tagName, size = { H1: 23, H2: 17, H3: 14, H4: 12, H5: 11, H6: 11 }[tag] || baseSize;
    const alignment = { center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED, left: D.AlignmentType.LEFT }[node.style.textAlign] || D.AlignmentType.LEFT;
    const length = (value, defaultValue = 0) => value ? Math.round(parseNumber(value, 0) * (/px$/.test(value) ? 15 : /mm$/.test(value) ? 56.69 : 20)) : defaultValue;
    return new D.Paragraph({
      alignment,
      spacing: { before: length(node.style.marginTop, !edge && tag.startsWith('H') ? 120 : 0), after: length(node.style.marginBottom, inCell || edge ? 0 : tag.startsWith('H') ? 180 : 200), line: /(?:pt|px|mm)$/.test(node.style.lineHeight) ? length(node.style.lineHeight) : Math.round(parseNumber(node.style.lineHeight, edge || baseSize === 9 ? 1.2 : 1.4) * 240), lineRule: /(?:pt|px|mm)$/.test(node.style.lineHeight) ? D.LineRuleType.EXACT : D.LineRuleType.AUTO },
      indent: { left: length(node.style.marginLeft), firstLine: length(node.style.textIndent) },
      children: runs(node, { font: settings.fontFamily || 'Arial', size: size * 2, bold: tag.startsWith('H') }),
      ...extra
    });
  }
  function blocks(parent, baseSize = Number(settings.fontSize) || 11, edge = false) {
    const result = [];
    for (const node of parent.children) {
      if (node.tagName === 'TABLE') {
        result.push(new D.Table({
          width: { size: 100, type: D.WidthType.PERCENTAGE },
          rows: [...node.rows].map((row) => new D.TableRow({
            tableHeader: Boolean(row.querySelector('th')), cantSplit: true,
            children: [...row.cells].map((cell) => {
              const content = blocks(cell, baseSize, edge);
              return new D.TableCell({
              columnSpan: cell.colSpan || 1, rowSpan: cell.rowSpan || 1,
              margins: { top: 90, bottom: 90, left: 120, right: 120 },
              shading: { fill: hexColor(cell.style.backgroundColor) || (cell.tagName === 'TH' ? 'F0F4F2' : 'FFFFFF') },
              borders: Object.fromEntries(['top', 'bottom', 'left', 'right'].map((side) => [side, { style: D.BorderStyle.SINGLE, size: 4, color: 'CAD4CE' }])),
              children: content.length ? content : [paragraph(cell, {}, baseSize, true, edge)]
            }); })
          }))
        }));
      } else if (node.matches('ul,ol')) {
        const instance = ++listInstance;
        for (const item of node.children) {
          const p = item.querySelector('p') || item;
          result.push(paragraph(p, node.tagName === 'UL' ? { bullet: { level: 0 } } : { numbering: { reference: 'paperlink-numbered', level: 0, instance } }, baseSize, false, edge));
        }
      } else if (!node.matches('[data-page-break]')) result.push(paragraph(node, {}, baseSize, parent.matches('td,th'), edge));
    }
    return result;
  }
  const sections = pages.map((page) => {
    const body = page.querySelector('.page-body'), header = page.querySelector('.page-header'), footer = page.querySelector('.page-footer');
    const footerBlocks = blocks(footer, 9, true);
    if (settings.pageNumbers) footerBlocks.push(new D.Paragraph({ alignment: D.AlignmentType.RIGHT, spacing: { before: footerBlocks.length ? Math.round(2 / 25.4 * 1440) : 0, after: 0, line: 192, lineRule: D.LineRuleType.EXACT }, children: [new D.TextRun({ text: 'Page ', size: 16 }), new D.TextRun({ children: [D.PageNumber.CURRENT], size: 16 }), new D.TextRun({ text: ' of ', size: 16 }), new D.TextRun({ children: [D.PageNumber.TOTAL_PAGES], size: 16 })] }));
    const content = blocks(body);
    const pixelsToTwips = (value, fallback) => Math.round(parseNumber(value, fallback) * 15);
    return {
      properties: { type: D.SectionType.NEXT_PAGE, page: { size: { width: Math.round(g.width / 25.4 * 1440), height: Math.round(g.height / 25.4 * 1440) }, margin: { top: pixelsToTwips(page.dataset.bodyTop, g.marginPx), bottom: pixelsToTwips(page.dataset.bodyBottom, g.marginPx), left: Math.round(g.margin / 25.4 * 1440), right: Math.round(g.margin / 25.4 * 1440), header: pixelsToTwips(page.dataset.headerTop, 280 / 15), footer: pixelsToTwips(page.dataset.footerBottom, 280 / 15) } } },
      headers: { default: new D.Header({ children: blocks(header, 9, true) }) },
      footers: { default: new D.Footer({ children: footerBlocks }) },
      children: content.length ? content : [new D.Paragraph('')]
    };
  });
  const doc = new D.Document({ creator: 'PaperLink', title: 'Linked document export', styles: { default: { document: { run: { font: settings.fontFamily || 'Arial', size: (Number(settings.fontSize) || 11) * 2 } } } }, numbering: { config: [{ reference: 'paperlink-numbered', levels: [{ level: 0, format: D.LevelFormat.DECIMAL, text: '%1.', alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 200 } } } }] }] }, sections });
  return blobBytes(await D.Packer.toBlob(doc));
}
export function rangeWorkbook(project, rangeText) {
  const sheetName = project.settings.sheet, sheet = project.sheet(sheetName);
  if (!/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(rangeText || '')) throw new Error('Use a cell range such as A1:F20.');
  const range = XLSX.utils.decode_range(rangeText.toUpperCase()), data = [];
  if (range.e.r < range.s.r || range.e.c < range.s.c || (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) > 100000) throw new Error('Choose a valid range of at most 100,000 cells.');
  for (let row = range.s.r; row <= range.e.r; row++) {
    const cells = [];
    for (let col = range.s.c; col <= range.e.c; col++) cells.push(sheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v ?? '');
    data.push(cells);
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(data), sheetName);
  return workbook;
}
export function spreadsheetBytes(project, format, rangeText = null) {
  const workbook = rangeText ? rangeWorkbook(project, rangeText) : project.workbook;
  if (format === 'tsv') {
    const sheet = workbook.Sheets[rangeText ? workbook.SheetNames[0] : project.settings.sheet];
    return new TextEncoder().encode(XLSX.utils.sheet_to_csv(sheet, { FS: '\t' }));
  }
  return new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: format, sheet: project.settings.sheet }));
}
export async function imageBytes(page, format, scale = 2) {
  const clone = page.cloneNode(true); clone.classList.remove('current-page', 'overflow-page'); clone.style.boxShadow = 'none'; clone.style.margin = '0';
  const host = document.createElement('div'); Object.assign(host.style, { position: 'fixed', left: '-20000px', top: '0', width: page.style.width }); host.append(clone); document.body.append(host);
  try {
    if (format === 'jpeg') {
      const url = await toJpeg(clone, { pixelRatio: scale, quality: .95, backgroundColor: '#ffffff', cacheBust: false, skipFonts: true });
      return Uint8Array.from(atob(url.split(',')[1]), (c) => c.charCodeAt(0));
    }
    return blobBytes(await toBlob(clone, { pixelRatio: scale, backgroundColor: '#ffffff', cacheBust: false, skipFonts: true }));
  } finally { host.remove(); }
}
export async function documentBytes(pages, settings, format, native, scale = 2) {
  if (format === 'docx') return docxBytes(pages, settings);
  if (format === 'pdf') {
    if (!native?.renderPDF) throw new Error('PDF generation is available in the desktop app.');
    return new Uint8Array(await native.renderPDF(printHtml(pages, settings)));
  }
  if (format === 'html') return new TextEncoder().encode(printHtml(pages, settings));
  if (format === 'txt') return new TextEncoder().encode(pages.map((page) => page.querySelector('.page-body').innerText).join('\n\n'));
  if (['png', 'jpeg'].includes(format)) {
    if (pages.length === 1) return imageBytes(pages[0], format, scale);
    const zip = new JSZip();
    for (let i = 0; i < pages.length; i++) zip.file('page-' + String(i + 1).padStart(3, '0') + (format === 'jpeg' ? '.jpg' : '.png'), await imageBytes(pages[i], format, scale));
    return zip.generateAsync({ type: 'uint8array' });
  }
  throw new Error('Unsupported export format.');
}
