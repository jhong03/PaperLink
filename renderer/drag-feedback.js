let activeDrag = null, frame = 0;

function moveFeedback(event) {
  if (!activeDrag) return;
  activeDrag.x = event.clientX; activeDrag.y = event.clientY;
  activeDrag.inEditor = event.target instanceof Element && Boolean(event.target.closest('.tiptap[contenteditable="true"]'));
  if (frame) return;
  frame = requestAnimationFrame(() => {
    frame = 0;
    if (!activeDrag) return;
    const { badge, hint, x, y, inEditor } = activeDrag;
    badge.style.visibility = 'visible';
    hint.textContent = inEditor ? 'Drop at the green line' : 'Move into the document';
    // Keep the pointer and the text immediately beneath it unobstructed. Move
    // the badge to the other side at the window edges, without moving the target.
    const width = badge.offsetWidth, height = badge.offsetHeight;
    const left = x + 18 + width <= window.innerWidth - 8 ? x + 18 : x - width - 18;
    const top = y + 24 + height <= window.innerHeight - 8 ? y + 24 : y - height - 24;
    badge.style.left = Math.max(8, left) + 'px'; badge.style.top = Math.max(8, top) + 'px';
  });
}

export function endDragFeedback() {
  cancelAnimationFrame(frame); frame = 0;
  activeDrag?.badge.remove(); activeDrag?.image.remove(); activeDrag = null;
  document.body.classList.remove('placing-document-item');
}

export function startDragFeedback(event, label) {
  if (!event.dataTransfer) return;
  endDragFeedback();
  const image = document.createElement('canvas'); image.width = image.height = 1;
  image.className = 'transparent-drag-image'; image.setAttribute('aria-hidden', 'true');
  const badge = document.createElement('div'), title = document.createElement('div'), hint = document.createElement('div');
  badge.className = 'document-drag-badge'; badge.setAttribute('aria-hidden', 'true');
  title.className = 'document-drag-title'; title.textContent = label;
  hint.className = 'document-drag-hint'; badge.append(title, hint);
  document.body.append(image, badge);
  // Suppress the browser's cell/node snapshot, which sits over the drop point.
  // The separate badge cannot intercept events or change document coordinates.
  event.dataTransfer.setDragImage(image, 0, 0);
  document.body.classList.add('placing-document-item');
  activeDrag = { badge, hint, image }; moveFeedback(event);
}

document.addEventListener('dragover', moveFeedback, true);
document.addEventListener('dragleave', (event) => { if (activeDrag && !event.relatedTarget) activeDrag.badge.style.visibility = 'hidden'; }, true);
document.addEventListener('drop', endDragFeedback, true);
document.addEventListener('dragend', endDragFeedback, true);
document.addEventListener('keydown', (event) => { if (event.key === 'Escape') endDragFeedback(); }, true);
window.addEventListener('blur', endDragFeedback);
