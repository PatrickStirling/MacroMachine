export function createExportMenuController(options = {}) {
  const { button, toggleButton, menu, getItems } = options;
  const trigger = toggleButton || button;
  let cleanup = null;

  const close = () => {
    if (cleanup) {
      cleanup();
      cleanup = null;
    }
    try { trigger?.setAttribute?.('aria-expanded', 'false'); } catch (_) {}
    if (!menu) return;
    menu.hidden = true;
    menu.innerHTML = '';
  };

  const open = () => {
    if (!menu || !trigger) return;
    close();
    const items = typeof getItems === 'function' ? getItems() : [];
    try { trigger?.setAttribute?.('aria-expanded', 'true'); } catch (_) {}
    menu.hidden = false;
    menu.innerHTML = '';
    items.forEach((item) => {
      if (item.type === 'sep') {
        const sep = document.createElement('div');
        sep.className = 'menu-sep';
        menu.appendChild(sep);
        return;
      }
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = item.label;
      btn.disabled = !item.enabled;
      btn.addEventListener('click', () => {
        close();
        item.action?.();
      });
      menu.appendChild(btn);
    });
    const onMouseDown = (ev) => {
      if (menu && menu.contains(ev.target)) return;
      if (button && button.contains(ev.target)) return;
      if (trigger && trigger.contains(ev.target)) return;
      close();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') close();
    };
    const onViewportChange = () => close();
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('scroll', onViewportChange, true);
    cleanup = () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onViewportChange);
      window.removeEventListener('scroll', onViewportChange, true);
    };
  };

  const setEnabled = (enabled) => {
    if (button) button.disabled = !enabled;
    if (trigger && trigger !== button) trigger.disabled = !enabled;
    if (!enabled) close();
  };

  const onToggle = (ev) => {
    if (!menu || !trigger) return;
    ev.preventDefault();
    ev.stopPropagation();
    if (!menu.hidden) {
      close();
      return;
    }
    open();
  };

  trigger?.addEventListener('click', onToggle);

  return {
    open,
    close,
    setEnabled,
  };
}
