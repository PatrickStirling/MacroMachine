export function createDiagnosticsController({ element, defaultEnabled = false, maxLines = 400 } = {}) {
  let enabled = !!defaultEnabled;
  if (element) element.hidden = !enabled;
  const maxCount = Number.isFinite(Number(maxLines)) ? Math.max(1, Math.floor(Number(maxLines))) : 400;

  function trimLines() {
    if (!element) return;
    while (element.childElementCount > maxCount) {
      try { element.removeChild(element.firstChild); } catch (_) { break; }
    }
  }

  function appendLine(text) {
    if (!element) return;
    element.hidden = false;
    const div = document.createElement('div');
    div.textContent = text;
    element.appendChild(div);
    trimLines();
  }

  if (enabled) {
    const ts = new Date().toLocaleTimeString();
    appendLine(`[${ts}] Diagnostics enabled (startup)`);
  }

  function log(line) {
    if (!enabled) return;
    const ts = new Date().toLocaleTimeString();
    const msg = `[${ts}] ${line}`;
    try {
      if (typeof console !== 'undefined' && console.log) console.log(msg);
    } catch (_) {
      /* no-op */
    }
    appendLine(msg);
  }

  function setEnabled(on) {
    enabled = !!on;
    if (element) element.hidden = !enabled;
    if (enabled) log('Diagnostics enabled');
  }

  function isEnabled() {
    return enabled;
  }

  function logTag(tag, message) {
    log(`[${tag}] ${message}`);
  }

  function toggle() {
    setEnabled(!enabled);
  }

  function clear() {
    if (!element) return;
    element.innerHTML = '';
  }

  return {
    log,
    logTag,
    setEnabled,
    isEnabled,
    toggle,
    clear,
  };
}
