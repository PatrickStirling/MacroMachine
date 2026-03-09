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
    onCreateFromFile,
    onCreateFromClipboard,
    onUpdateExportPathDisplay,
    onSelectCsvBatch,
    getDocDisplayName,
  } = options;

  let addMenu = null;
  let addMenuCleanup = null;
  let overflowRaf = 0;

  const scheduleOverflowUpdate = () => {
    if (overflowRaf) return;
    overflowRaf = requestAnimationFrame(() => {
      overflowRaf = 0;
      try { updateOverflow(); } catch (_) {}
    });
  };

  const closeAddMenu = () => {
    if (addMenuCleanup) {
      addMenuCleanup();
      addMenuCleanup = null;
    }
    if (addMenu) {
      try { addMenu.remove(); } catch (_) {}
      addMenu = null;
    }
  };

  const openAddMenu = (x, y) => {
    closeAddMenu();
    const root = document.createElement('div');
    root.className = 'doc-tab-menu';
    root.setAttribute('role', 'menu');
    const addItem = (label, onClick) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'doc-tab-menu-item';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeAddMenu();
        onClick?.();
      });
      root.appendChild(btn);
    };
    addItem('Import from file', () => (onCreateFromFile || onCreateBlankDocument)?.());
    addItem('Import from clipboard', () => onCreateFromClipboard?.());
    document.body.appendChild(root);
    addMenu = root;
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
      if (!root.contains(ev.target)) closeAddMenu();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') closeAddMenu();
    };
    const onViewportChange = () => closeAddMenu();
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('scroll', onViewportChange, true);
    addMenuCleanup = () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onViewportChange);
      window.removeEventListener('scroll', onViewportChange, true);
    };
  };

  const render = () => {
    if (!docTabsEl) return;
    closeAddMenu();
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
      docTabsEl.addEventListener('scroll', scheduleOverflowUpdate);
      window.addEventListener('resize', scheduleOverflowUpdate);
      docTabsPrev?.addEventListener('click', () => scroll(-1));
      docTabsNext?.addEventListener('click', () => scroll(1));
    }
    const activeDocId = getActiveDocId();
    documents.forEach((doc) => {
      const isBatch = !!doc.isCsvBatch;
      const wrap = document.createElement('div');
      wrap.className = `doc-tab${doc.id === activeDocId ? ' active' : ''}${doc.selected ? ' selected' : ''}${isBatch ? ' csv-batch' : ''}`;
      wrap.dataset.docId = doc.id;
      if (!isBatch) {
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
      } else {
        wrap.draggable = false;
      }
      const labelBtn = document.createElement('button');
      labelBtn.type = 'button';
      labelBtn.className = 'doc-tab-label';
      const label = getDocDisplayName?.(doc, documents) || doc.name || doc.fileName || 'Untitled';
      const lastExport = doc.snapshot && doc.snapshot.lastExportPath ? doc.snapshot.lastExportPath : '';
      const exportPath = doc.snapshot && doc.snapshot.exportFolder ? doc.snapshot.exportFolder : '';
      const exportLabel = lastExport || exportPath || '';
      labelBtn.textContent = doc.isDirty ? `${label} *` : label;
      if (isBatch) {
        const count = doc.csvBatch && Number.isFinite(doc.csvBatch.count) ? doc.csvBatch.count : 0;
        const folder = doc.csvBatch && doc.csvBatch.folderPath ? doc.csvBatch.folderPath : '';
        labelBtn.title = `CSV batch: ${count} files${folder ? ` • ${folder}` : ''}`;
        labelBtn.addEventListener('click', (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          if (ev.ctrlKey || ev.metaKey) {
            onToggleDocSelection?.(doc.id);
            return;
          }
          onSelectCsvBatch?.(doc);
        });
        labelBtn.addEventListener('contextmenu', (ev) => {
          ev.preventDefault();
          onOpenContextMenu?.(doc.id, ev.clientX, ev.clientY);
        });
      } else {
        labelBtn.title = `${label} • Export: ${exportLabel || 'Default (Fusion Templates)'}`;
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
      }
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
      if (!isBatch) {
        wrap.appendChild(closeBtn);
      } else {
        closeBtn.title = 'Remove batch';
        wrap.appendChild(closeBtn);
      }
      docTabsEl.appendChild(wrap);
    });
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'doc-tab doc-tab-add';
    addBtn.textContent = '+';
    addBtn.title = 'Import from file (right-click for options)';
    addBtn.addEventListener('click', () => (onCreateFromFile || onCreateBlankDocument)?.());
    addBtn.addEventListener('contextmenu', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      openAddMenu(ev.clientX, ev.clientY);
    });
    addBtn.draggable = false;
    docTabsEl.appendChild(addBtn);
    onUpdateExportPathDisplay?.();
    scheduleOverflowUpdate();
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
    closeAddMenu,
  };
}

