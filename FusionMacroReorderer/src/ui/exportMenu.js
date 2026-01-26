export function createExportMenuController(options = {}) {
  const { button, menu, getItems } = options;
  let cleanup = null;

  const close = () => {
    if (cleanup) {
      cleanup();
      cleanup = null;
    }
    if (!menu) return;
    menu.hidden = true;
    menu.innerHTML = '';
  };

  const open = () => {
    if (!menu || !button) return;
    close();
    const items = typeof getItems === 'function' ? getItems() : [];
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
    if (!button) return;
    button.disabled = !enabled;
    if (!enabled) close();
  };

  const onToggle = (ev) => {
    if (!menu || !button) return;
    ev.preventDefault();
    ev.stopPropagation();
    if (!menu.hidden) {
      close();
      return;
    }
    open();
  };

  button?.addEventListener('click', onToggle);

  return {
    open,
    close,
    setEnabled,
  };
}
