export function createGlobalShortcuts(options = {}) {
  const {
    isElectron,
    getDocuments,
    getActiveDocId,
    isInteractiveTarget,
    onSwitchByIndex,
    onSwitchByOffset,
    onCloseDoc,
    onCreateDoc,
  } = options;

  try {
    window.addEventListener('beforeunload', (ev) => {
      try {
        if (isElectron) return;
        const docs = typeof getDocuments === 'function' ? getDocuments() : [];
        if (!docs.length) return;
        const dirty = docs.some(doc => doc && doc.isDirty);
        if (!dirty) return;
        // Use the native unload confirmation dialog instead of window.confirm.
        ev.preventDefault();
        ev.returnValue = '';
      } catch (_) {}
    });
  } catch (_) {}

  try {
    const keyHandler = (ev) => {
      try {
        const docs = typeof getDocuments === 'function' ? getDocuments() : [];
        if (!docs.length) return;
        if (!(ev.ctrlKey || ev.metaKey)) return;
        if (typeof isInteractiveTarget === 'function' && isInteractiveTarget(ev.target)) return;
        const key = ev.key ? ev.key.toLowerCase() : '';
        if (key >= '1' && key <= '9') {
          ev.preventDefault();
          onSwitchByIndex?.(parseInt(key, 10));
          return;
        }
        if (key === 'tab') {
          ev.preventDefault();
          onSwitchByOffset?.(ev.shiftKey ? -1 : 1);
          return;
        }
        if (key === 'w') {
          ev.preventDefault();
          const activeId = typeof getActiveDocId === 'function' ? getActiveDocId() : null;
          if (activeId) onCloseDoc?.(activeId);
          return;
        }
        if (key === 't') {
          ev.preventDefault();
          onCreateDoc?.();
        }
      } catch (_) {}
    };
    window.addEventListener('keydown', keyHandler);
  } catch (_) {}
}
