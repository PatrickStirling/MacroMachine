export function createDocTabContextMenu(options = {}) {
  const {
    getDocument,
    onCreateDoc,
    onRenameDoc,
    onToggleDocSelection,
    onSelectAll,
    onClearSelection,
    onCloseDoc,
    onCloseOthers,
    onCloseAll,
  } = options;

  let menu = null;
  let cleanup = null;

  const close = () => {
    if (cleanup) {
      cleanup();
      cleanup = null;
    }
    if (menu) {
      try { menu.remove(); } catch (_) {}
      menu = null;
    }
  };

  const open = (docId, x, y) => {
    close();
    const doc = typeof getDocument === 'function' ? getDocument(docId) : null;
    if (!doc) return;
    const root = document.createElement('div');
    root.className = 'doc-tab-menu';
    root.setAttribute('role', 'menu');
    const addItem = (label, onClick, opts = {}) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'doc-tab-menu-item' + (opts.danger ? ' danger' : '');
      btn.textContent = label;
      btn.addEventListener('click', () => {
        close();
        onClick?.();
      });
      root.appendChild(btn);
    };
    const addSeparator = () => {
      const sep = document.createElement('div');
      sep.className = 'doc-tab-menu-sep';
      root.appendChild(sep);
    };
    addItem('New tab', () => onCreateDoc?.());
    if (!doc.isCsvBatch) {
      addItem('Rename...', () => onRenameDoc?.(docId));
    }
    addItem(doc.selected ? 'Unselect for Export' : 'Select for Export', () => onToggleDocSelection?.(docId));
    addItem('Select All for Export', () => onSelectAll?.());
    addItem('Clear Selection', () => onClearSelection?.());
    addSeparator();
    addItem('Close tab', () => onCloseDoc?.(docId), { danger: true });
    addItem('Close others', () => onCloseOthers?.(docId), { danger: true });
    addItem('Close all', () => onCloseAll?.(), { danger: true });

    document.body.appendChild(root);
    menu = root;
    const pad = 8;
    const placeMenu = () => {
      const rect = root.getBoundingClientRect();
      let left = Number.isFinite(x) ? x : pad;
      let top = Number.isFinite(y) ? y : pad;
      if (left + rect.width > window.innerWidth - pad) left = window.innerWidth - rect.width - pad;
      if (top + rect.height > window.innerHeight - pad) top = window.innerHeight - rect.height - pad;
      if (left < pad) left = pad;
      if (top < pad) top = pad;
      root.style.left = `${left}px`;
      root.style.top = `${top}px`;
    };
    placeMenu();
    requestAnimationFrame(placeMenu);

    const onMouseDown = (ev) => {
      if (!root.contains(ev.target)) close();
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

  return { open, close };
}
