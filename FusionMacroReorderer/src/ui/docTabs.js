export function createDocTabsController(options = {}) {
  const {
    docTabsEl,
    docTabsWrap,
    docTabsPrev,
    docTabsNext,
    getDocuments,
    getActiveDocId,
    getDraggingDocId,
    setDraggingDocId,
    onReorderDocuments,
    onToggleDocSelection,
    onClearDocSelections,
    onSwitchDocument,
    onPromptRename,
    onOpenContextMenu,
    onCloseDocument,
    onCreateBlankDocument,
    onUpdateExportPathDisplay,
  } = options;

  const render = () => {
    if (!docTabsEl) return;
    const documents = getDocuments();
    if (documents.length === 0) {
      if (docTabsWrap) docTabsWrap.hidden = true;
      docTabsEl.innerHTML = '';
      onUpdateExportPathDisplay?.();
      return;
    }
    if (docTabsWrap) docTabsWrap.hidden = false;
    docTabsEl.innerHTML = '';
    if (!docTabsEl.dataset.docDnD) {
      docTabsEl.dataset.docDnD = '1';
      docTabsEl.addEventListener('dragover', (ev) => {
        if (!getDraggingDocId()) return;
        const target = ev.target;
        if (target && target.closest && target.closest('.doc-tab')) return;
        ev.preventDefault();
      });
      docTabsEl.addEventListener('drop', (ev) => {
        const dragging = getDraggingDocId();
        if (!dragging) return;
        const target = ev.target;
        if (target && target.closest && target.closest('.doc-tab')) return;
        ev.preventDefault();
        onReorderDocuments?.(dragging, null, true);
        setDraggingDocId(null);
      });
    }
    if (!docTabsEl.dataset.docScroll) {
      docTabsEl.dataset.docScroll = '1';
      docTabsEl.addEventListener('scroll', () => updateOverflow());
      window.addEventListener('resize', () => updateOverflow());
      docTabsPrev?.addEventListener('click', () => scroll(-1));
      docTabsNext?.addEventListener('click', () => scroll(1));
    }
    const activeDocId = getActiveDocId();
    documents.forEach((doc) => {
      const wrap = document.createElement('div');
      wrap.className = `doc-tab${doc.id === activeDocId ? ' active' : ''}${doc.selected ? ' selected' : ''}`;
      wrap.dataset.docId = doc.id;
      wrap.draggable = true;
      wrap.addEventListener('dragstart', (ev) => {
        setDraggingDocId(doc.id);
        wrap.classList.add('dragging');
        ev.dataTransfer?.setData('text/plain', doc.id);
      });
      wrap.addEventListener('dragend', () => {
        setDraggingDocId(null);
        wrap.classList.remove('dragging');
        docTabsEl.querySelectorAll('.doc-tab.drag-over').forEach(el => el.classList.remove('drag-over'));
      });
      wrap.addEventListener('dragover', (ev) => {
        if (!getDraggingDocId() || getDraggingDocId() === doc.id) return;
        ev.preventDefault();
        wrap.classList.add('drag-over');
      });
      wrap.addEventListener('dragleave', () => wrap.classList.remove('drag-over'));
      wrap.addEventListener('drop', (ev) => {
        const dragging = getDraggingDocId();
        if (!dragging || dragging === doc.id) return;
        ev.preventDefault();
        wrap.classList.remove('drag-over');
        onReorderDocuments?.(dragging, doc.id);
        setDraggingDocId(null);
      });
      const labelBtn = document.createElement('button');
      labelBtn.type = 'button';
      labelBtn.className = 'doc-tab-label';
      const label = doc.name || doc.fileName || 'Untitled';
      const lastExport = doc.snapshot && doc.snapshot.lastExportPath ? doc.snapshot.lastExportPath : '';
      const exportPath = doc.snapshot && doc.snapshot.exportFolder ? doc.snapshot.exportFolder : '';
      const exportLabel = lastExport || exportPath || '';
      labelBtn.textContent = doc.isDirty ? `${label} *` : label;
      labelBtn.title = `${doc.fileName || doc.name || 'Untitled'} • Export: ${exportLabel || 'Default (Fusion Templates)'}`;
      labelBtn.addEventListener('click', (ev) => {
        if (ev.ctrlKey || ev.metaKey) {
          ev.preventDefault();
          ev.stopPropagation();
          onToggleDocSelection?.(doc.id);
          return;
        }
        onClearDocSelections?.();
        onSwitchDocument?.(doc.id);
      });
      labelBtn.addEventListener('dblclick', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        onPromptRename?.(doc.id);
      });
      labelBtn.addEventListener('contextmenu', (ev) => {
        ev.preventDefault();
        onOpenContextMenu?.(doc.id, ev.clientX, ev.clientY);
      });
      labelBtn.draggable = false;
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'doc-tab-close';
      closeBtn.textContent = 'x';
      closeBtn.title = 'Close';
      closeBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        onCloseDocument?.(doc.id);
      });
      closeBtn.draggable = false;
      wrap.appendChild(labelBtn);
      wrap.appendChild(closeBtn);
      docTabsEl.appendChild(wrap);
    });
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'doc-tab doc-tab-add';
    addBtn.textContent = '+';
    addBtn.title = 'Open .setting';
    addBtn.addEventListener('click', () => onCreateBlankDocument?.());
    addBtn.draggable = false;
    docTabsEl.appendChild(addBtn);
    onUpdateExportPathDisplay?.();
    requestAnimationFrame(() => updateOverflow());
  };

  const updateOverflow = () => {
    if (!docTabsEl || !docTabsPrev || !docTabsNext) return;
    const overflow = docTabsEl.scrollWidth > docTabsEl.clientWidth + 4;
    docTabsPrev.hidden = !overflow;
    docTabsNext.hidden = !overflow;
    if (!overflow) return;
    const left = docTabsEl.scrollLeft;
    const maxLeft = docTabsEl.scrollWidth - docTabsEl.clientWidth - 2;
    docTabsPrev.disabled = left <= 2;
    docTabsNext.disabled = left >= maxLeft;
  };

  const scroll = (direction) => {
    if (!docTabsEl) return;
    const amount = Math.max(120, Math.floor(docTabsEl.clientWidth * 0.6));
    docTabsEl.scrollBy({ left: direction * amount, behavior: 'smooth' });
  };

  return {
    render,
    updateOverflow,
  };
}
