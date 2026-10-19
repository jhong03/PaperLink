import { escapeHtml } from './model.mjs';
export const PX_PER_MM = 96 / 25.4;
export function geometry(settings) {
  let [width, height] = { A4: [210, 297], Letter: [215.9, 279.4], Legal: [215.9, 355.6] }[settings.paper] || [210, 297];
  if (settings.landscape) [width, height] = [height, width];
  const margin = Math.max(8, Math.min(40, Number(settings.margin) || 20));
  return { width, height, widthPx: width * PX_PER_MM, heightPx: height * PX_PER_MM, margin, marginPx: margin * PX_PER_MM, bodyWidth: (width - margin * 2) * PX_PER_MM, bodyHeight: (height - margin * 2) * PX_PER_MM };
}
export function resolveHtml(html, project, row) {
  const root = document.createElement('div'); root.innerHTML = html;
  root.querySelectorAll('[data-field-id]').forEach((el) => {
    const field = { sheet: el.dataset.sheet, address: el.dataset.address, followRecord: el.dataset.followRecord !== 'false' };
    el.textContent = project.fieldValue(field, row);
    el.removeAttribute('contenteditable'); el.className = '';
  });
  root.querySelectorAll('[data-page-break]').forEach((el) => { el.textContent = ''; });
  return root.innerHTML;
}
export function renumber(pages, enabled = true) {
  pages.forEach((page, index) => {
    const number = page.querySelector('.page-number');
    if (number) number.textContent = enabled ? 'Page ' + (index + 1) + ' of ' + pages.length : '';
    page.dataset.page = String(index + 1);
  });
  return pages;
}
export function paginate(html, settings, host, headerHtml = '', footerHtml = '') {
  const g = geometry(settings), pages = [], errors = [];
  host.replaceChildren();
  const edgeGap = 3 * PX_PER_MM, headerTop = 5 * PX_PER_MM;
  const numberBottom = 8 * PX_PER_MM, numberHeight = 8 * 96 / 72 * 1.2;
  const footerBottom = settings.pageNumbers ? numberBottom + numberHeight + 2 * PX_PER_MM : 5 * PX_PER_MM;
  let bodyTop = g.marginPx, bodyBottom = g.marginPx;
  function edge(className, html) {
    const element = document.createElement('div'); element.className = className; element.innerHTML = html;
    // Editors keep an empty paragraph even when the user has not added a footer.
    // An empty linked cell is also blank; images and tables still need space.
    if (!element.textContent.replace(/[\s\u200b\ufeff]/g, '') && !element.querySelector('img,table,hr')) element.replaceChildren();
    Object.assign(element.style, { left: g.marginPx + 'px', right: g.marginPx + 'px' });
    return element;
  }
  function edgeHeight(element) {
    return element.childNodes.length ? Math.ceil(Math.max(element.scrollHeight, element.getBoundingClientRect().height)) : 0;
  }
  let body;
  function newPage() {
    const page = document.createElement('div'); page.className = 'document-page';
    Object.assign(page.style, { width: g.widthPx + 'px', height: g.heightPx + 'px', fontFamily: settings.fontFamily || 'Arial', fontSize: (Number(settings.fontSize) || 11) + 'pt' });
    const header = edge('page-header', headerHtml); header.style.top = headerTop + 'px';
    body = document.createElement('div'); body.className = 'page-body';
    const footer = edge('page-footer', footerHtml); footer.style.bottom = footerBottom + 'px';
    const number = document.createElement('div'); number.className = 'page-number';
    Object.assign(number.style, { bottom: numberBottom + 'px', right: g.marginPx + 'px', lineHeight: '1.2' });
    page.append(header, body, footer, number); host.append(page); pages.push(page);
    if (pages.length === 1) {
      // Measure at the final page width before flowing any body text. Margins are
      // minimum clearances, so a multiline footer can grow without being clipped.
      const headerHeight = edgeHeight(header), footerHeight = edgeHeight(footer);
      bodyTop = Math.max(g.marginPx, headerHeight ? headerTop + headerHeight + edgeGap : 0);
      bodyBottom = Math.max(g.marginPx, footerHeight ? footerBottom + footerHeight + edgeGap : 0, settings.pageNumbers ? numberBottom + numberHeight + edgeGap : 0);
      g.bodyHeight = g.heightPx - bodyTop - bodyBottom;
      g.topMarginPx = bodyTop; g.bottomMarginPx = bodyBottom;
      if (g.bodyHeight < 24 * PX_PER_MM) errors.push('The header and footer leave too little room for the document. Shorten their content or use a taller page in Page setup.');
    }
    Object.assign(body.style, { position: 'absolute', left: g.marginPx + 'px', top: bodyTop + 'px', width: g.bodyWidth + 'px', height: Math.max(0, g.bodyHeight) + 'px' });
    // Word export uses the same reserved space as the preview, PDF and print.
    Object.assign(page.dataset, { bodyTop: String(bodyTop), bodyBottom: String(bodyBottom), headerTop: String(headerTop), footerBottom: String(settings.pageNumbers ? numberBottom : footerBottom) });
    for (const image of page.querySelectorAll('.page-header img,.page-footer img')) {
      const rect = image.getBoundingClientRect();
      if (rect.width && rect.height) { image.setAttribute('width', String(rect.width)); image.setAttribute('height', String(rect.height)); }
    }
    return page;
  }
  const fits = () => body.scrollHeight <= Math.ceil(g.bodyHeight) + 1;
  const root = document.createElement('div'); root.innerHTML = html;
  newPage();
  if (errors.length) { renumber(pages, settings.pageNumbers); return { pages, errors, geometry: g }; }
  function oversized(message) { pages.at(-1).classList.add('overflow-page'); errors.push(message); }
  function appendSimple(block) {
    const copy = block.cloneNode(true);
    copy.querySelectorAll('img').forEach((image) => { image.style.maxHeight = g.bodyHeight + 'px'; image.style.maxWidth = '100%'; image.style.objectFit = 'contain'; });
    body.append(copy);
    if (fits()) return;
    copy.remove();
    if (body.childNodes.length) newPage();
    body.append(copy);
    if (!fits()) {
      copy.remove();
      if (['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(copy.tagName)) splitParagraph(copy);
      else { body.append(copy); oversized('A content block is taller than one page. Split it or reduce its size before export.'); }
    }
  }
  function splitParagraph(block) {
    const points = [{ node: block, offset: 0 }];
    function scan(node) {
      if (node.nodeType === 3) {
        const pattern = /\S+\s*|\s+/g; let match;
        while ((match = pattern.exec(node.textContent))) points.push({ node, offset: match.index + match[0].length });
      } else if (node.nodeType === 1 && (node.matches('[data-field-id],img,br'))) {
        points.push({ node: node.parentNode, offset: [...node.parentNode.childNodes].indexOf(node) + 1 });
      } else for (const item of node.childNodes) scan(item);
    }
    [...block.childNodes].forEach(scan);
    const last = { node: block, offset: block.childNodes.length };
    points.push(last);
    let start = 0;
    const slice = (from, to) => {
      const range = document.createRange(); range.setStart(points[from].node, points[from].offset); range.setEnd(points[to].node, points[to].offset);
      return range.cloneContents();
    };
    while (start < points.length - 1) {
      const part = block.cloneNode(false); if (start) { part.style.textIndent = '0'; part.style.marginTop = '0'; }
      body.append(part);
      let low = start + 1, high = points.length - 1, best = start;
      while (low <= high) {
        const middle = Math.floor((low + high) / 2); part.replaceChildren(slice(start, middle));
        if (fits()) { best = middle; low = middle + 1; } else high = middle - 1;
      }
      if (best === start) {
        part.remove();
        if (body.childNodes.length) { newPage(); continue; }
        body.append(block.cloneNode(true)); oversized('A paragraph cannot fit on a page. Reduce its font size or margins.'); return;
      }
      part.replaceChildren(slice(start, best)); start = best;
      if (start < points.length - 1) newPage();
    }
  }
  function appendTable(table) {
    const rows = [...table.rows], header = rows[0]?.querySelector('th') ? rows[0] : null;
    const dataRows = header ? rows.slice(1) : rows;
    if (!dataRows.length) { appendSimple(table); return; }
    const groups = [];
    for (let index = 0; index < dataRows.length;) {
      let end = index;
      for (let rowIndex = index; rowIndex <= end && rowIndex < dataRows.length; rowIndex++) for (const cell of dataRows[rowIndex].cells) end = Math.max(end, rowIndex + (cell.rowSpan || 1) - 1);
      groups.push(dataRows.slice(index, Math.min(end + 1, dataRows.length))); index = end + 1;
    }
    let current, tbody, dataCount;
    function begin() {
      current = table.cloneNode(false); current.style.width = '100%';
      const columns = table.querySelector(':scope > colgroup');
      if (columns) {
        const copy = columns.cloneNode(true), cols = [...copy.children], total = cols.reduce((sum, col) => sum + (Number.parseFloat(col.style.width) || 0), 0);
        if (total > g.bodyWidth) cols.forEach((col) => { col.style.width = (Number.parseFloat(col.style.width) || 0) / total * g.bodyWidth + 'px'; });
        current.append(copy);
      }
      tbody = document.createElement('tbody'); current.append(tbody);
      if (header) tbody.append(header.cloneNode(true));
      body.append(current); dataCount = 0;
    }
    begin();
    for (const group of groups) {
      let copies = group.map((row) => row.cloneNode(true)); copies.forEach((row) => tbody.append(row));
      if (!fits()) {
        copies.forEach((row) => row.remove());
        if (!dataCount) current.remove();
        if (body.childNodes.length || dataCount) newPage();
        begin(); copies = group.map((row) => row.cloneNode(true)); copies.forEach((row) => tbody.append(row));
        if (!fits()) oversized('A table row or merged-row group is taller than one page. Split the row or reduce its content.');
      }
      dataCount += group.length;
    }
  }
  function appendList(list) {
    let current = list.cloneNode(false), count = Number(list.getAttribute('start') || 1);
    body.append(current);
    for (const li of [...list.children]) {
      const copy = li.cloneNode(true); current.append(copy);
      if (!fits()) {
        copy.remove(); if (!current.children.length) current.remove();
        if (body.childNodes.length) newPage();
        current = list.cloneNode(false); if (list.tagName === 'OL') current.setAttribute('start', count); body.append(current); current.append(copy);
        if (!fits()) oversized('A list item is taller than one page. Split the item or reduce its size.');
      }
      count++;
    }
  }
  for (const block of [...root.children]) {
    if (block.matches('[data-page-break]')) { if (body.childNodes.length) newPage(); continue; }
    if (block.tagName === 'TABLE') appendTable(block);
    else if (block.matches('ul,ol')) appendList(block);
    else appendSimple(block);
  }
  if (pages.length > 1 && !pages.at(-1).querySelector('.page-body').childNodes.length) pages.pop().remove();
  renumber(pages, settings.pageNumbers);
  return { pages, errors: [...new Set(errors)], geometry: g };
}
// One set of edge styles for measurement, the live preview and printed HTML.
// flow-root includes paragraph margins and floats in the measured height.
export const PAGE_EDGE_CSS = '.page-header,.page-footer{position:absolute;display:flow-root;overflow:visible;font-size:9pt;line-height:1.2;overflow-wrap:anywhere}.page-header:empty,.page-footer:empty{display:none}.page-header p,.page-footer p{margin:0}.page-header :is(h1,h2,h3,h4,h5,h6),.page-footer :is(h1,h2,h3,h4,h5,h6){margin:0;line-height:1.2}.page-header h1,.page-footer h1{font-size:23pt}.page-header h2,.page-footer h2{font-size:17pt}.page-header h3,.page-footer h3{font-size:14pt}.page-header h4,.page-footer h4{font-size:12pt}.page-header :is(h5,h6),.page-footer :is(h5,h6){font-size:11pt}.page-header :is(ul,ol),.page-footer :is(ul,ol){margin:0;padding-left:24px}.page-header table,.page-footer table{border-collapse:collapse;table-layout:fixed;width:100%;margin:0}.page-header :is(td,th),.page-footer :is(td,th){border:1px solid #cad4ce;padding:6px 8px;vertical-align:top;overflow-wrap:anywhere}.page-header img,.page-footer img{max-width:100%;height:auto}';
export const PAGE_CSS = 'html,body{margin:0;padding:0;background:#fff;color:#202d28}*{box-sizing:border-box}.document-page{position:relative;overflow:hidden;break-after:page;margin:0;box-shadow:none}.document-page:last-child{break-after:auto}.page-body{line-height:1.4;overflow:hidden;overflow-wrap:anywhere}.page-body p{margin:0 0 10pt}.page-body h1{font-size:23pt;line-height:1.15;margin:0 0 14pt}.page-body h2{font-size:17pt;line-height:1.2;margin:12pt 0 9pt}.page-body h3{font-size:14pt;margin:10pt 0 7pt}.page-body ul,.page-body ol{padding-left:24px}.page-body li p{margin-bottom:4pt}.page-body table{border-collapse:collapse;table-layout:fixed;width:100%;margin:12px 0}.page-body th,.page-body td{border:1px solid #cad4ce;padding:6px 8px;vertical-align:top;overflow-wrap:anywhere}.page-body th{background:#f0f4f2}.page-body td p,.page-body th p{margin:0}img{max-width:100%;height:auto}.page-number{position:absolute;bottom:8mm;right:20mm;font:8pt Arial;color:#8b918e}a{color:#26699f}' + PAGE_EDGE_CSS;
export function printHtml(pages, settings) {
  const g = geometry(settings);
  return '<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src data:; style-src \'unsafe-inline\'; font-src data:;"><title>PaperLink output</title><style>@page{size:' + g.width + 'mm ' + g.height + 'mm;margin:0}' + PAGE_CSS + '</style></head><body>' + pages.map((page) => page.outerHTML).join('') + '</body></html>';
}
