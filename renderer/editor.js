import { Editor, Node, Extension, mergeAttributes } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { Table, TableRow, TableCell, TableHeader } from '@tiptap/extension-table';
import { TextStyleKit } from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import Image from '@tiptap/extension-image';
import { startDragFeedback } from './drag-feedback.js';

const isCellDrag = (event) => [...(event.dataTransfer?.types || [])].some((type) => ['application/x-paperlink-cells', 'application/x-flowdesk-cells'].includes(type));

const Format = Extension.create({
  name: 'flowFormatting',
  addGlobalAttributes() {
    return [{ types: ['paragraph', 'heading', 'tableCell', 'tableHeader'], attributes: Object.fromEntries([
      ['marginTop', 'margin-top'], ['marginBottom', 'margin-bottom'], ['marginLeft', 'margin-left'], ['textIndent', 'text-indent'], ['lineHeight', 'line-height'], ['backgroundColor', 'background-color']
    ].map(([key, property]) => [key, { default: null, parseHTML: (el) => el.style.getPropertyValue(property) || null, renderHTML: (attrs) => attrs[key] ? { style: property + ':' + attrs[key] } : {} }])) }];
  }
});
const cellAttributes = (parent) => ({
  ...parent,
  colwidth: {
    default: null,
    parseHTML: (el) => {
      if (el.dataset.colwidth) return el.dataset.colwidth.split(',').map(Number);
      const width = Number.parseFloat(el.style.width || el.getAttribute('width') || '');
      return width ? [Math.round(width * (el.style.width.endsWith('pt') ? 96 / 72 : 1))] : null;
    },
    renderHTML: (attrs) => attrs.colwidth ? { 'data-colwidth': attrs.colwidth.join(',') } : {}
  }
});
const Cell = TableCell.extend({ addAttributes() { return cellAttributes(this.parent()); } });
const Header = TableHeader.extend({ addAttributes() { return cellAttributes(this.parent()); } });
const EmbeddedImage = Image.extend({
  addAttributes() { return { ...this.parent(), width: { default: null }, height: { default: null } }; }
});
const PageBreak = Node.create({
  name: 'pageBreak', group: 'block', atom: true,
  parseHTML: () => [{ tag: 'div[data-page-break]' }],
  renderHTML: () => ['div', { 'data-page-break': 'true', class: 'page-break' }]
});
export function createEditor(element, callbacks, content = '<p></p>') {
  const Field = Node.create({
    name: 'linkedField', inline: true, group: 'inline', atom: true, draggable: true, selectable: true,
    addAttributes() {
      return {
        id: { default: '', parseHTML: (el) => el.dataset.fieldId },
        sheet: { default: '', parseHTML: (el) => el.dataset.sheet },
        address: { default: 'A1', parseHTML: (el) => el.dataset.address },
        label: { default: 'Cell', parseHTML: (el) => el.dataset.label },
        followRecord: { default: true, parseHTML: (el) => el.dataset.followRecord !== 'false' },
        value: { default: '', parseHTML: (el) => el.textContent }
      };
    },
    parseHTML: () => [{ tag: 'span[data-field-id]' }],
    renderHTML({ node }) {
      return ['span', { 'data-field-id': node.attrs.id, 'data-sheet': node.attrs.sheet, 'data-address': node.attrs.address, 'data-label': node.attrs.label, 'data-follow-record': String(node.attrs.followRecord), class: 'mapped-output' }, node.attrs.value];
    },
    addNodeView() {
      return ({ node }) => {
        let current = node;
        const dom = document.createElement('span'), input = document.createElement('input');
        dom.className = 'mapped-field'; dom.contentEditable = 'false'; dom.draggable = true;
        input.type = 'text'; input.spellcheck = false; dom.append(input);
        function update() {
          dom.dataset.fieldId = current.attrs.id;
          input.setAttribute('aria-label', current.attrs.label + ' value from spreadsheet');
          const project = callbacks.project(), source = project.source(current.attrs);
          const cell = project.cell(source.sheet, source.address);
          input.title = current.attrs.label + ' from ' + source.sheet + ', cell ' + source.address + (cell?.f ? '. Calculated in the spreadsheet.' : '. Edit here or in the spreadsheet to update both.');
          input.readOnly = Boolean(cell?.f);
          dom.classList.toggle('formula-derived', Boolean(cell?.f));
          if (document.activeElement !== input) input.value = cell?.f ? current.attrs.value || '' : String(cell?.v ?? '');
          input.placeholder = callbacks.templateEditing?.() ? '[' + source.sheet + '!' + source.address + ']' : '';
          input.style.width = Math.min(55, Math.max(2, (input.value || input.placeholder).length + .8)) + 'ch';
        }
        input.addEventListener('focus', () => {
          callbacks.select(current.attrs.id);
          // Values stay in raw edit form here; formatting is applied in the live
          // page preview. Changing the input value on focus would destroy its selection.
          try { update(); } catch { input.readOnly = true; }
        });
        input.addEventListener('input', () => { callbacks.edit(current.attrs, input.value); update(); });
        input.addEventListener('blur', () => { try { update(); } catch { input.readOnly = true; } });
        // Embedded value controls own their typing and deletion events; the surrounding
        // document/table editor must not interpret those events as document edits.
        input.addEventListener('beforeinput', (event) => event.stopPropagation());
        input.addEventListener('keydown', (event) => { event.stopPropagation(); if (event.key === 'Enter') { event.preventDefault(); input.blur(); } });
        dom.addEventListener('click', () => callbacks.select(current.attrs.id));
        try { update(); } catch { input.value = '[Missing source]'; input.readOnly = true; }
        return {
          dom,
          update(updated) { if (updated.type !== current.type) return false; current = updated; try { update(); } catch { input.value = '[Missing source]'; input.readOnly = true; } return true; },
          stopEvent(event) { return event.target === input && !['dragstart', 'dragenter', 'dragover', 'dragleave', 'drop', 'dragend'].includes(event.type); },
          ignoreMutation: () => true
        };
      };
    }
  });
  return new Editor({
    element,
    extensions: [StarterKit.configure({ link: { openOnClick: false }, dropcursor: { color: '#087f5b', width: 3, class: 'document-drop-indicator' } }), TextStyleKit, TextAlign.configure({ types: ['heading', 'paragraph'] }), Table.configure({ resizable: true }), TableRow, Cell, Header, EmbeddedImage.configure({ inline: true, allowBase64: true }), Format, PageBreak, Field],
    content,
    onUpdate: callbacks.update,
    onFocus: ({ editor }) => callbacks.focus(editor),
    onSelectionUpdate: callbacks.selection,
    editorProps: {
      attributes: { spellcheck: 'true' },
      handleDOMEvents: {
        dragstart(view, event) {
          const field = event.target instanceof Element ? event.target.closest('.mapped-field') : null;
          if (field) {
            const value = field.querySelector('input');
            startDragFeedback(event, 'Move: ' + (value?.value || value?.placeholder || 'connected value'));
          }
          // ProseMirror still prepares and moves the original linked node.
          return false;
        },
        dragover(view, event) {
          if (!isCellDrag(event)) return false;
          event.preventDefault();
          const position = view.posAtCoords({ left: event.clientX, top: event.clientY });
          event.dataTransfer.dropEffect = position ? 'copy' : 'none';
          return true;
        },
        drop(view, event) {
          const data = event.dataTransfer?.getData('application/x-paperlink-cells') || event.dataTransfer?.getData('application/x-flowdesk-cells');
          if (!data) return false;
          event.preventDefault();
          const position = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos;
          if (position == null) return true;
          callbacks.drop(JSON.parse(data), position);
          return true;
        }
      }
    }
  });
}
