export function createDiagnosticsController({ element, defaultEnabled = false } = {}) {
  let enabled = !!defaultEnabled;
  if (element) element.hidden = !enabled;
  function appendLine(text) {
    if (!element) return;
    element.hidden = false;
    const div = document.createElement('div');
    div.textContent = text;
    element.appendChild(div);
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

  return {
    log,
    logTag,
    setEnabled,
    isEnabled,
    toggle,
  };
}
