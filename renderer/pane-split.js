const STORAGE_KEY = 'paperlink-workspace-split';
const DEFAULT_RATIO = 3 / 7;

export function setupPaneSplit(workspace, divider, onResize) {
  let preferredRatio = DEFAULT_RATIO, drag = null, pendingRatio = null, frame = 0;
  try {
    const value = Number(localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem('flowdesk-workspace-split'));
    if (Number.isFinite(value) && value > 0 && value < 1) preferredRatio = value;
  } catch { /* Resizing remains available if browser preference storage is disabled. */ }

  function dimensions() {
    const bounds = workspace.getBoundingClientRect(), style = getComputedStyle(workspace);
    const leftPadding = Number.parseFloat(style.paddingLeft) || 0, rightPadding = Number.parseFloat(style.paddingRight) || 0;
    const available = Math.max(1, workspace.clientWidth - leftPadding - rightPadding - divider.offsetWidth);
    return {
      left: bounds.left + leftPadding, available,
      minimum: Math.min(280, available * .35) / available,
      maximum: 1 - Math.min(340, available * .35) / available
    };
  }
  function apply(nextRatio = preferredRatio, remember = false) {
    const size = dimensions(), ratio = Math.max(size.minimum, Math.min(size.maximum, nextRatio));
    if (remember) preferredRatio = ratio;
    workspace.style.setProperty('--sheet-size', ratio + 'fr');
    workspace.style.setProperty('--document-size', (1 - ratio) + 'fr');
    divider.setAttribute('aria-valuemin', String(Math.round(size.minimum * 100)));
    divider.setAttribute('aria-valuemax', String(Math.round(size.maximum * 100)));
    divider.setAttribute('aria-valuenow', String(Math.round(ratio * 100)));
    divider.setAttribute('aria-valuetext', 'Spreadsheet ' + Math.round(ratio * 100) + '%, document ' + Math.round((1 - ratio) * 100) + '%');
    onResize();
  }
  function savePreference() {
    try { localStorage.setItem(STORAGE_KEY, String(preferredRatio)); } catch { /* Keep the current layout for this session. */ }
  }
  function schedule(nextRatio = null) {
    if (nextRatio !== null) pendingRatio = nextRatio;
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      if (pendingRatio !== null) { const next = pendingRatio; pendingRatio = null; apply(next, true); }
      else apply();
    });
  }
  function finish(cancel = false) {
    if (!drag) return;
    const previous = drag; drag = null;
    if (frame) { cancelAnimationFrame(frame); frame = 0; }
    if (cancel) { preferredRatio = previous.startRatio; apply(); }
    else if (pendingRatio !== null) apply(pendingRatio, true);
    pendingRatio = null;
    document.body.classList.remove('resizing-panes');
    if (divider.hasPointerCapture(previous.pointerId)) divider.releasePointerCapture(previous.pointerId);
    savePreference();
  }
  divider.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || drag) return;
    event.preventDefault(); divider.focus({ preventScroll: true });
    drag = { pointerId: event.pointerId, offset: event.clientX - divider.getBoundingClientRect().left, startRatio: preferredRatio };
    divider.setPointerCapture(event.pointerId); document.body.classList.add('resizing-panes');
  });
  divider.addEventListener('pointermove', (event) => {
    if (!drag || event.pointerId !== drag.pointerId) return;
    const size = dimensions(); schedule((event.clientX - size.left - drag.offset) / size.available);
  });
  divider.addEventListener('pointerup', (event) => {
    if (drag?.pointerId !== event.pointerId) return;
    const size = dimensions(); pendingRatio = (event.clientX - size.left - drag.offset) / size.available; finish();
  });
  divider.addEventListener('pointercancel', (event) => { if (drag?.pointerId === event.pointerId) finish(true); });
  divider.addEventListener('lostpointercapture', () => finish());
  window.addEventListener('blur', () => finish());
  function reset() {
    if (drag) finish(true);
    if (frame) { cancelAnimationFrame(frame); frame = 0; } pendingRatio = null;
    preferredRatio = DEFAULT_RATIO; apply(); savePreference();
  }
  divider.addEventListener('dblclick', (event) => { event.preventDefault(); reset(); });
  divider.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && drag) { event.preventDefault(); finish(true); return; }
    if (event.key === 'Enter') { event.preventDefault(); reset(); return; }
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const size = dimensions(), step = (event.shiftKey ? 80 : 20) / size.available;
    const current = Math.max(size.minimum, Math.min(size.maximum, preferredRatio));
    const next = event.key === 'Home' ? size.minimum : event.key === 'End' ? size.maximum : current + (event.key === 'ArrowLeft' ? -step : step);
    apply(next, true); savePreference();
  });
  // Keep the preferred ratio when a narrower window temporarily limits either pane.
  const observer = new ResizeObserver(() => schedule()); observer.observe(workspace);
  apply();
}
