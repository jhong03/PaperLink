import JSZip from 'jszip';
import DOMPurify from 'dompurify';
import { RTFJS, WMFJS, EMFJS } from 'rtf.js';
import { escapeHtml } from './model.mjs';

const xml = (text) => {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('The document contains invalid XML.');
  return doc;
};
const kids = (node, tag) => [...(node?.children || [])].filter((child) => !tag || child.localName === tag);
const child = (node, tag) => kids(node, tag)[0];
const all = (node, tag) => [...(node?.getElementsByTagNameNS('*', tag) || [])];
const att = (node, name) => [...(node?.attributes || [])].find((attr) => attr.localName === name)?.value || '';
const css = (styles) => Object.entries(styles).filter(([, value]) => value !== '' && value !== undefined).map(([key, value]) => key + ':' + value).join(';');
const utf8 = (bytes) => new TextDecoder().decode(bytes);
function zipPath(base, target) {
  const parts = (base.slice(0, base.lastIndexOf('/') + 1) + target).split('/'), stack = [];
  for (const part of parts) { if (part === '..') stack.pop(); else if (part && part !== '.') stack.push(part); }
  return stack.join('/');
}
export function cleanHtml(html) {
  const safe = DOMPurify.sanitize(html, { ADD_ATTR: ['data-page-break', 'data-field-id', 'data-sheet', 'data-address', 'data-label', 'data-follow-record'], FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'svg', 'math'], FORBID_ATTR: ['srcset'] });
  const root = document.createElement('div'); root.innerHTML = safe;
  root.querySelectorAll('[style]').forEach((el) => { if (/url\s*\(|@import|expression\s*\(/i.test(el.getAttribute('style'))) el.removeAttribute('style'); });
  root.querySelectorAll('img').forEach((el) => { if (!/^data:image\/(png|jpeg|gif|webp|bmp);base64,/i.test(el.getAttribute('src') || '')) el.remove(); });
  root.querySelectorAll('a').forEach((el) => { if (!/^(https?:|mailto:|#)/i.test(el.getAttribute('href') || '')) el.removeAttribute('href'); });
  return root.innerHTML;
}
async function openZip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const files = Object.values(zip.files);
  if (files.length > 6000 || files.reduce((sum, file) => sum + (file._data?.uncompressedSize || 0), 0) > 120 * 1024 * 1024) throw new Error('This compressed document is too large to edit.');
  return zip;
}
async function readXml(zip, name) { const file = zip.file(name); return file ? xml(await file.async('string')) : null; }
async function relationships(zip, base) {
  const location = base.slice(0, base.lastIndexOf('/') + 1) + '_rels/' + base.slice(base.lastIndexOf('/') + 1) + '.rels';
  const doc = await readXml(zip, location), result = {};
  for (const rel of all(doc, 'Relationship')) result[att(rel, 'Id')] = { target: att(rel, 'Target'), external: att(rel, 'TargetMode') === 'External' };
  return result;
}
async function inlineImage(zip, file) {
  const mime = { png: 'png', jpg: 'jpeg', jpeg: 'jpeg', gif: 'gif', bmp: 'bmp', webp: 'webp' }[file.split('.').pop().toLowerCase()];
  if (!mime || !zip.file(file)) return '';
  return 'data:image/' + mime + ';base64,' + await zip.file(file).async('base64');
}
function runStyle(node) {
  const result = {}, fonts = child(node, 'rFonts'), size = att(child(node, 'sz'), 'val'), color = att(child(node, 'color'), 'val');
  const enabled = (tag) => { const value = att(child(node, tag), 'val'); return Boolean(child(node, tag)) && !['0', 'false', 'off', 'none'].includes(value); };
  if (fonts) result['font-family'] = att(fonts, 'ascii') || att(fonts, 'hAnsi');
  if (result['font-family']?.startsWith('+')) delete result['font-family'];
  if (size) result['font-size'] = Number(size) / 2 + 'pt';
  if (color && color !== 'auto') result.color = '#' + color;
  if (enabled('b')) result['font-weight'] = 'bold';
  if (enabled('i')) result['font-style'] = 'italic';
  if (enabled('u')) result['text-decoration'] = 'underline';
  if (enabled('strike')) result['text-decoration'] = (result['text-decoration'] || '') + ' line-through';
  const highlight = att(child(node, 'highlight'), 'val');
  if (highlight && highlight !== 'none') result['background-color'] = highlight;
  return result;
}
function paragraphStyle(node) {
  const result = {}, spacing = child(node, 'spacing'), indent = child(node, 'ind');
  const align = att(child(node, 'jc'), 'val');
  if (align) result['text-align'] = { both: 'justify', start: 'left', end: 'right' }[align] || align;
  for (const name of ['before', 'after']) if (att(spacing, name)) result['margin-' + (name === 'before' ? 'top' : 'bottom')] = Number(att(spacing, name)) / 20 + 'pt';
  if (att(spacing, 'line')) result['line-height'] = att(spacing, 'lineRule') === 'exact' ? Number(att(spacing, 'line')) / 20 + 'pt' : Number(att(spacing, 'line')) / 240;
  if (att(indent, 'left') || att(indent, 'start')) result['margin-left'] = Number(att(indent, 'left') || att(indent, 'start')) / 20 + 'pt';
  if (att(indent, 'firstLine')) result['text-indent'] = Number(att(indent, 'firstLine')) / 20 + 'pt';
  return result;
}
async function importDocx(bytes) {
  const zip = await openZip(bytes), doc = await readXml(zip, 'word/document.xml');
  if (!doc) throw new Error('This file does not contain a Word document.');
  const stylesDoc = await readXml(zip, 'word/styles.xml'), styleNodes = new Map(all(stylesDoc, 'style').map((style) => [att(style, 'styleId'), style]));
  const numberingDoc = await readXml(zip, 'word/numbering.xml');
  const abstractNums = new Map(all(numberingDoc, 'abstractNum').map((node) => [att(node, 'abstractNumId'), node]));
  const numDefinitions = new Map(all(numberingDoc, 'num').map((node) => [att(node, 'numId'), node]));
  const defaultRun = runStyle(all(stylesDoc, 'rPrDefault')[0] && child(all(stylesDoc, 'rPrDefault')[0], 'rPr'));
  function styleFor(id, seen = new Set()) {
    if (!id || seen.has(id) || !styleNodes.has(id)) return { p: {}, r: {}, name: '' };
    seen.add(id); const style = styleNodes.get(id), base = styleFor(att(child(style, 'basedOn'), 'val'), seen);
    return { p: { ...base.p, ...paragraphStyle(child(style, 'pPr')) }, r: { ...base.r, ...runStyle(child(style, 'rPr')) }, name: att(child(style, 'name'), 'val') || base.name };
  }
  async function convert(base, body) {
    const rels = await relationships(zip, base);
    const counts = new Map();
    async function run(node, defaults) {
      let text = '';
      for (const item of kids(node)) {
        if (item.localName === 't' || item.localName === 'delText') text += escapeHtml(item.textContent);
        if (item.localName === 'tab') text += '    ';
        if (item.localName === 'br' || item.localName === 'cr') text += att(item, 'type') === 'page' ? '\uFFF0' : '<br>';
        if (item.localName === 'drawing' || item.localName === 'pict') {
          const blip = all(item, 'blip')[0] || all(item, 'imagedata')[0], id = att(blip, 'embed') || att(blip, 'id'), rel = rels[id];
          if (rel && !rel.external) {
            const src = await inlineImage(zip, zipPath(base, rel.target)), extent = all(item, 'extent')[0];
            if (src) text += '<img src="' + src + '" width="' + Math.round(Number(att(extent, 'cx') || 1905000) / 9525) + '" height="' + Math.round(Number(att(extent, 'cy') || 952500) / 9525) + '">';
          }
        }
      }
      const styles = { ...defaults, ...runStyle(child(node, 'rPr')) };
      return Object.keys(styles).length ? '<span style="' + escapeHtml(css(styles)) + '">' + text + '</span>' : text;
    }
    async function paragraph(node) {
      const ppr = child(node, 'pPr'), id = att(child(ppr, 'pStyle'), 'val'), style = styleFor(id);
      let text = '';
      for (const item of kids(node)) {
        if (item.localName === 'r') text += await run(item, { ...defaultRun, ...style.r });
        if (item.localName === 'hyperlink') {
          const rel = rels[att(item, 'id')], href = rel?.external ? rel.target : '#';
          text += '<a href="' + escapeHtml(href) + '">' + (await Promise.all(kids(item, 'r').map((r) => run(r, { ...defaultRun, ...style.r })))).join('') + '</a>';
        }
        if (item.localName === 'sdt') for (const r of all(item, 'r')) text += await run(r, { ...defaultRun, ...style.r });
      }
      const heading = (style.name || id).match(/heading\s*([1-6])/i), tag = heading ? 'h' + heading[1] : 'p';
      const attributes = css({ ...style.p, ...paragraphStyle(ppr) });
      const list = child(ppr, 'numPr'), listId = att(child(list, 'numId'), 'val'), level = Number(att(child(list, 'ilvl'), 'val') || 0);
      let listAttributes = '';
      if (list && listId !== '0') {
        const definition = numDefinitions.get(listId), abstract = abstractNums.get(att(child(definition, 'abstractNumId'), 'val')), levelNode = kids(abstract, 'lvl').find((node) => Number(att(node, 'ilvl')) === level);
        const format = att(child(levelNode, 'numFmt'), 'val') === 'bullet' ? 'ul' : 'ol', key = listId + ':' + level, count = (counts.get(key) ?? Number(att(child(levelNode, 'start'), 'val') || 1) - 1) + 1;
        counts.set(key, count);
        for (const countKey of counts.keys()) if (countKey.startsWith(listId + ':') && Number(countKey.split(':')[1]) > level) counts.delete(countKey);
        listAttributes = ' data-list-id="' + escapeHtml(listId) + '" data-list-level="' + level + '" data-list-type="' + format + '" data-list-start="' + count + '"';
      }
      const pieces = text.split('\uFFF0').map((part) => '<' + tag + listAttributes + ' style="' + escapeHtml(attributes) + '">' + (part || '<br>') + '</' + tag + '>').join('<div data-page-break="true"></div>');
      const section = child(ppr, 'sectPr'), sectionBreak = section && att(child(section, 'type'), 'val') !== 'continuous';
      return (child(ppr, 'pageBreakBefore') ? '<div data-page-break="true"></div>' : '') + pieces + (sectionBreak ? '<div data-page-break="true"></div>' : '');
    }
    async function table(node) {
      const rows = [], active = new Map();
      for (const row of kids(node, 'tr')) {
        const cells = []; let col = 0;
        for (const cell of kids(row, 'tc')) {
          const props = child(cell, 'tcPr'), span = Number(att(child(props, 'gridSpan'), 'val') || 1), merge = child(props, 'vMerge');
          if (merge && att(merge, 'val') !== 'restart' && active.has(col)) { active.get(col).rowspan++; col += span; continue; }
          const descriptor = { cell, colspan: span, rowspan: 1, header: Boolean(child(child(row, 'trPr'), 'tblHeader')) };
          cells.push(descriptor);
          for (let i = 0; i < span; i++) { if (merge) active.set(col + i, descriptor); else active.delete(col + i); }
          col += span;
        }
        rows.push(cells);
      }
      const htmlRows = [];
      for (const row of rows) {
        const cells = [];
        for (const cell of row) {
          const props = child(cell.cell, 'tcPr'), fill = att(child(props, 'shd'), 'fill'), width = att(child(props, 'tcW'), 'w');
          const style = css({ 'background-color': fill && fill !== 'auto' ? '#' + fill : '', width: width ? Number(width) / 20 + 'pt' : '' });
          const tag = cell.header ? 'th' : 'td';
          cells.push('<' + tag + ' colspan="' + cell.colspan + '" rowspan="' + cell.rowspan + '" style="' + escapeHtml(style) + '">' + await blocks(cell.cell) + '</' + tag + '>');
        }
        htmlRows.push('<tr>' + cells.join('') + '</tr>');
      }
      return '<table><tbody>' + htmlRows.join('') + '</tbody></table>';
    }
    async function blocks(node) {
      const output = [];
      for (const item of kids(node)) {
        if (item.localName === 'p') output.push(await paragraph(item));
        if (item.localName === 'tbl') output.push(await table(item));
        if (item.localName === 'sdt') output.push(await blocks(child(item, 'sdtContent')));
      }
      const root = document.createElement('div'); root.innerHTML = output.join('');
      let stack = [];
      for (const item of [...root.children]) {
        if (!item.dataset.listId) { stack = []; continue; }
        const id = item.dataset.listId, level = Math.min(4, Number(item.dataset.listLevel) || 0), tag = item.dataset.listType || 'ul';
        if (stack[0]?.id !== id) stack = [];
        const depth = Math.min(level, stack.length);
        stack.length = Math.min(stack.length, depth + 1);
        if (!stack[depth] || stack[depth].id !== id || stack[depth].list.tagName.toLowerCase() !== tag) {
          const list = document.createElement(tag);
          if (tag === 'ol') list.setAttribute('start', item.dataset.listStart || '1');
          if (depth && stack[depth - 1]?.list.lastElementChild) stack[depth - 1].list.lastElementChild.append(list); else root.insertBefore(list, item);
          stack[depth] = { id, list };
        }
        const li = document.createElement('li'); stack[depth].list.append(li); li.append(item);
        for (const key of ['listId', 'listLevel', 'listType', 'listStart']) delete item.dataset[key];
      }
      return root.innerHTML;
    }
    return blocks(body);
  }
  const body = all(doc, 'body')[0], section = all(doc, 'sectPr').at(-1), rels = await relationships(zip, 'word/document.xml');
  let header = '', footer = '';
  for (const type of ['header', 'footer']) {
    const refs = kids(section, type + 'Reference'), ref = refs.find((item) => att(item, 'type') === 'default') || refs[0], rel = rels[att(ref, 'id')];
    if (rel && !rel.external) {
      const location = zipPath('word/document.xml', rel.target), part = await readXml(zip, location);
      if (part) { const value = await convert(location, part.documentElement); if (type === 'header') header = value; else footer = value; }
    }
  }
  const size = child(section, 'pgSz'), margin = child(section, 'pgMar');
  const width = Number(att(size, 'w')) / 1440 * 25.4, height = Number(att(size, 'h')) / 1440 * 25.4;
  const layout = { landscape: width > height };
  const short = Math.min(width, height), long = Math.max(width, height);
  if (Math.abs(short - 215.9) < 2 && Math.abs(long - 279.4) < 2) layout.paper = 'Letter';
  else if (Math.abs(short - 215.9) < 2 && Math.abs(long - 355.6) < 2) layout.paper = 'Legal';
  else layout.paper = 'A4';
  if (att(margin, 'left')) layout.margin = Math.max(8, Math.min(40, Number(att(margin, 'left')) / 1440 * 25.4));
  const warnings = [];
  if (all(doc, 'txbxContent').length || all(doc, 'chart').length || all(doc, 'object').length) warnings.push('Text boxes, charts, or embedded objects need a layout check; they are not fully editable in this version.');
  if (all(doc, 'sectPr').length > 1) warnings.push('Multiple section layouts are normalized to one page layout.');
  if (all(doc, 'fldSimple').length || all(doc, 'fldChar').length) warnings.push('Word fields are imported as their displayed text; field instructions are not retained.');
  return { html: cleanHtml(await convert('word/document.xml', body)), header: cleanHtml(header), footer: cleanHtml(footer), layout, warnings };
}
async function importOdt(bytes) {
  const zip = await openZip(bytes), doc = await readXml(zip, 'content.xml'), stylesDoc = await readXml(zip, 'styles.xml');
  if (!doc) throw new Error('This file does not contain an OpenDocument text document.');
  const styles = new Map();
  for (const style of [...all(stylesDoc, 'style'), ...all(doc, 'style')]) {
    const text = child(style, 'text-properties'), p = child(style, 'paragraph-properties');
    styles.set(att(style, 'name'), { 'font-family': att(text, 'font-family') || att(text, 'font-name'), 'font-size': att(text, 'font-size'), 'font-weight': att(text, 'font-weight'), 'font-style': att(text, 'font-style'), color: att(text, 'color'), 'text-align': att(p, 'text-align'), 'margin-top': att(p, 'margin-top'), 'margin-bottom': att(p, 'margin-bottom'), 'line-height': att(p, 'line-height') });
  }
  async function convert(node) {
    if (node.nodeType === 3) return escapeHtml(node.textContent);
    if (node.nodeType !== 1) return '';
    const tag = node.localName, style = css(styles.get(att(node, 'style-name')) || {}), content = async () => (await Promise.all([...node.childNodes].map(convert))).join('');
    if (tag === 'image') { const src = await inlineImage(zip, att(node, 'href').replace(/^\.\//, '')); return src ? '<img src="' + src + '">' : ''; }
    if (tag === 's') return '&nbsp;'.repeat(Math.min(200, Number(att(node, 'c') || 1)));
    if (tag === 'tab') return '    ';
    if (tag === 'line-break') return '<br>';
    if (tag === 'p' || tag === 'h') {
      const htmlTag = tag === 'h' ? 'h' + Math.min(6, Number(att(node, 'outline-level') || 1)) : 'p';
      return '<' + htmlTag + ' style="' + escapeHtml(style) + '"><span style="' + escapeHtml(style) + '">' + await content() + '</span></' + htmlTag + '>';
    }
    if (tag === 'span') return '<span style="' + escapeHtml(style) + '">' + await content() + '</span>';
    if (tag === 'a') return '<a href="' + escapeHtml(att(node, 'href')) + '">' + await content() + '</a>';
    if (tag === 'list') return '<ul>' + await content() + '</ul>';
    if (tag === 'list-item') return '<li>' + await content() + '</li>';
    if (tag === 'table') return '<table><tbody>' + await content() + '</tbody></table>';
    if (tag === 'table-row') return '<tr>' + await content() + '</tr>';
    if (tag === 'covered-table-cell') return '';
    if (tag === 'table-cell') {
      const repeat = Math.min(256, Number(att(node, 'number-columns-repeated') || 1));
      return ('<td colspan="' + (att(node, 'number-columns-spanned') || 1) + '" rowspan="' + (att(node, 'number-rows-spanned') || 1) + '">' + await content() + '</td>').repeat(repeat);
    }
    return content();
  }
  const content = all(doc, 'text')[0], master = all(stylesDoc, 'master-page')[0];
  return { html: cleanHtml(await convert(content)), header: cleanHtml(await convert(child(master, 'header') || document.createElement('div'))), footer: cleanHtml(await convert(child(master, 'footer') || document.createElement('div'))), layout: {}, warnings: ['OpenDocument content and basic formatting imported. Check page layout, lists, and table widths before issuing.'] };
}
export async function importDocument(bytes, name) {
  const extension = name.split('.').pop().toLowerCase();
  if (extension === 'docx') return importDocx(bytes);
  if (extension === 'odt') return importOdt(bytes);
  if (extension === 'rtf') {
    RTFJS.loggingEnabled(false); WMFJS.loggingEnabled(false); EMFJS.loggingEnabled(false);
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const rtf = new RTFJS.Document(buffer);
    const nodes = await rtf.render();
    return { html: cleanHtml(nodes.map((node) => node.outerHTML).join('')), header: '', footer: '', layout: {}, warnings: ['RTF content imported. Check table structure and page layout before export.'] };
  }
  if (['html', 'htm'].includes(extension)) return { html: cleanHtml(utf8(bytes)), header: '', footer: '', layout: {}, warnings: [] };
  if (extension === 'txt') return { html: utf8(bytes).split(/\r?\n/).map((line) => '<p>' + escapeHtml(line || ' ') + '</p>').join(''), header: '', footer: '', layout: {}, warnings: [] };
  throw new Error('Choose DOCX, ODT, RTF, HTML, or TXT. Legacy .doc files should be saved as .docx first.');
}
