import { createDiagnosticsController } from './src/diagnostics.js';
import { detectElectron, setupNativeBridge } from './src/nativeBridge.js';
import { appState } from './src/state.js';
import {
  parseSetting,
  findEnclosingGroupForIndex,
  extractQuotedProp,
  extractNumericProp,
  findAllInputsBlocksWithInstanceInputs,
  findInputsBlockAnywhere,
  findInputsBlockByBacktrack,
} from './src/parser.js';
import { getLineIndent, findMatchingBrace, isSpace, isIdentStart, isIdentPart, humanizeName, escapeQuotes, isQuoteEscaped } from './src/textUtils.js';
import { createPublishedControls } from './src/publishedControls.js';
import { createHistoryController } from './src/history.js';
import { createNodesPane } from './src/nodesPane.js';
import { createCatalogEditor } from './src/catalogEditor.js';

(() => {
  try {
    const stamp = new Date().toISOString();
    console.log('[FMR Build]', stamp);
  } catch (_) {}

  const state = appState;

  const fileInput = document.getElementById('fileInput');
  const docTabsWrap = document.getElementById('docTabsWrap');
  const docTabsEl = document.getElementById('docTabs');
  const docTabsPrev = document.getElementById('docTabsPrev');
  const docTabsNext = document.getElementById('docTabsNext');
  const docExportPathEl = document.getElementById('docExportPath');
  const introPanel = document.getElementById('introPanel');
  const introToggleBar = document.getElementById('introToggleBar');
  const introToggleBtn = document.getElementById('introToggleBtn');
  const openExplorerBtn = document.getElementById('openExplorerBtn');
  const macroExplorerPanel = document.getElementById('macroExplorerPanel');
  const macroExplorerMini = document.getElementById('macroExplorerMini');
  const closeExplorerPanelBtn = document.getElementById('closeExplorerPanelBtn');
  const exportMenuBtn = document.getElementById('exportMenuBtn');
  const exportMenuEl = document.getElementById('exportMenu');

  const dropZone = document.getElementById('dropZone');

  const fileInfo = document.getElementById('fileInfo');
  const dropHint = document.getElementById('dropHint');
  // Set app title / nameplate
  try {
    if (typeof document !== 'undefined') {
      document.title = 'Macro Machine';
      const title = document.getElementById('appTitle');
      if (title) {
        title.textContent = 'Macro Machine';
      }
      const subtitle = document.getElementById('appSubtitle');
      if (subtitle) {
        subtitle.innerHTML = '<span class="app-title-by"><a href="https://stirlingsupply.co" target="_blank" rel="noopener noreferrer">by Stirling Supply Co</a></span>';
      }
    }
  } catch (_) {}
  // Diagnostics flag and runtime toggle
  const LOG_DIAGNOSTICS = false;

  // Detect Electron (renderer with Node integration) and expose a thin native bridge.
  // This replaces the previous preload-based bridge, which was failing to load on some setups.
  const IS_ELECTRON = detectElectron();
  if (!IS_ELECTRON && macroExplorerMini) {
    macroExplorerMini.hidden = true;
  }

  const documents = [];
  let activeDocId = null;
  let docCounter = 1;
  let suppressDocDirty = false;
  let draggingDocId = null;
  let docContextMenu = null;
  let docContextMenuCleanup = null;
  let exportMenuCleanup = null;

  function createDocumentMeta({ name, fileName } = {}) {
    return {
      id: `doc_${docCounter++}`,
      name: name || fileName || 'Untitled',
      fileName: fileName || '',
      isDirty: false,
      selected: false,
      createdAt: Date.now(),
      snapshot: null,
    };
  }

  function addDocument(doc, makeActive = true) {
    if (!doc) return null;
    documents.push(doc);
    if (makeActive || !activeDocId) activeDocId = doc.id;
    renderDocTabs();
    return doc;
  }

  function getActiveDocument() {
    if (!activeDocId) return null;
    return documents.find((doc) => doc.id === activeDocId) || null;
  }

  function updateActiveDocumentMeta(meta = {}) {
    const doc = getActiveDocument();
    if (!doc) return null;
    if (meta.name) doc.name = meta.name;
    if (meta.fileName !== undefined) doc.fileName = meta.fileName || '';
    renderDocTabs();
    return doc;
  }

  function markActiveDocumentDirty() {
    const doc = getActiveDocument();
    if (!doc || doc.isDirty) return;
    doc.isDirty = true;
    renderDocTabs();
  }

  function markActiveDocumentClean() {
    const doc = getActiveDocument();
    if (!doc || !doc.isDirty) return;
    doc.isDirty = false;
    renderDocTabs();
  }

  function markDocumentClean(doc) {
    if (!doc || !doc.isDirty) return;
    doc.isDirty = false;
  }

  function clearDocSelections() {
    let changed = false;
    documents.forEach((doc) => {
      if (doc.selected) {
        doc.selected = false;
        changed = true;
      }
    });
    if (changed) renderDocTabs();
  }

  function toggleDocSelection(docId) {
    const doc = documents.find(d => d.id === docId);
    if (!doc) return;
    doc.selected = !doc.selected;
    renderDocTabs();
  }

  function resetStateToEmpty() {
    state.parseResult = null;
    state.originalText = '';
    state.originalFileName = '';
    state.originalFilePath = '';
    state.generatedText = '';
    state.lastDiffRange = null;
    state.newline = '\n';
    state.exportFolder = '';
    state.exportFolderSelected = false;
    state.lastExportPath = '';
    state.drfxLink = null;
  }

  function createBlankDocument() {
    const current = getActiveDocument();
    if (current) storeDocumentSnapshot(current);
    const name = `Untitled ${docCounter}`;
    const doc = createDocumentMeta({ name });
    doc.snapshot = null;
    doc.isDirty = false;
    addDocument(doc, true);
    resetStateToEmpty();
    clearUI();
    updateExportButtonLabelFromPath('');
    renderDocTabs();
    return doc;
  }

  function closeDocument(docId) {
    if (!docId) return;
    const index = documents.findIndex((doc) => doc.id === docId);
    if (index < 0) return;
    const doc = documents[index];
    const label = doc.name || doc.fileName || 'Untitled';
    if (doc.isDirty) {
      const ok = window.confirm(`Close "${label}" without exporting? Your changes will be lost.`);
      if (!ok) return;
    }
    const closingActive = docId === activeDocId;
    documents.splice(index, 1);
    if (closingActive) {
      let nextDoc = null;
      if (documents.length) {
        const nextIndex = index > 0 ? index - 1 : 0;
        nextDoc = documents[nextIndex] || documents[0] || null;
      }
      activeDocId = nextDoc ? nextDoc.id : null;
      if (nextDoc) {
        applyDocumentSnapshot(nextDoc);
      } else {
        resetStateToEmpty();
        clearUI();
        updateExportButtonLabelFromPath('');
      }
    }
    renderDocTabs();
  }

  function closeDocumentsByFilter(filterFn, message) {
    if (!documents.length) return;
    const remaining = [];
    const toClose = [];
    documents.forEach((doc) => {
      if (filterFn(doc)) toClose.push(doc);
      else remaining.push(doc);
    });
    if (!toClose.length) return;
    const dirtyCount = toClose.filter(doc => doc.isDirty).length;
    if (dirtyCount) {
      const prompt = message || `Close ${toClose.length} tab(s) without exporting?`;
      const ok = window.confirm(prompt);
      if (!ok) return;
    }
    documents.length = 0;
    remaining.forEach(doc => documents.push(doc));
    if (remaining.length) {
      if (!remaining.find(doc => doc.id === activeDocId)) {
        activeDocId = remaining[0].id;
        applyDocumentSnapshot(remaining[0]);
      }
    } else {
      activeDocId = null;
      resetStateToEmpty();
      clearUI();
      updateExportButtonLabelFromPath('');
    }
    renderDocTabs();
  }

  function buildDocumentSnapshot() {
    if (!state.parseResult) return null;
    const originalText = state.originalText || '';
    return {
      parseResult: state.parseResult,
      originalText,
      originalFileName: state.originalFileName || '',
      originalFilePath: state.originalFilePath || '',
      newline: state.newline || detectNewline(originalText),
      generatedText: state.generatedText || originalText,
      lastDiffRange: state.lastDiffRange || null,
      exportFolder: state.exportFolder || '',
      exportFolderSelected: !!state.exportFolderSelected,
      lastExportPath: state.lastExportPath || '',
      drfxLink: state.drfxLink || null,
    };
  }

  function storeDocumentSnapshot(doc) {
    if (!doc) return null;
    const snap = buildDocumentSnapshot();
    if (!snap) return doc;
    doc.snapshot = snap;
    return doc;
  }

  function applyDocumentSnapshot(doc) {
    const snap = doc && doc.snapshot ? doc.snapshot : null;
    const prevSuppress = suppressDocDirty;
    suppressDocDirty = true;
    try {
      if (!snap || !snap.parseResult) {
        resetStateToEmpty();
        clearUI();
        return;
      }
      state.parseResult = snap.parseResult;
      state.originalText = snap.originalText || '';
      state.originalFileName = snap.originalFileName || 'Imported.setting';
      state.originalFilePath = snap.originalFilePath || '';
      state.newline = snap.newline || detectNewline(state.originalText || '');
      state.lastDiffRange = snap.lastDiffRange || null;
      state.exportFolder = snap.exportFolder || '';
      state.exportFolderSelected = !!snap.exportFolderSelected;
      state.lastExportPath = snap.lastExportPath || '';
      state.drfxLink = snap.drfxLink || null;

      const codeText = snap.generatedText != null ? snap.generatedText : (state.originalText || '');
      updateCodeView(codeText);
      clearCodeHighlight();
      if (fileInfo) {
        fileInfo.textContent = `${state.originalFileName} (${(state.originalText || '').length.toLocaleString()} chars)`;
      }
      macroNameEl.textContent = state.parseResult.macroName || 'Unknown';
      ctrlCountEl.textContent = String(state.parseResult.entries ? state.parseResult.entries.length : 0);
      controlsSection.hidden = false;
      resetBtn.disabled = false;
      exportBtn.disabled = false;
      exportTemplatesBtn && (exportTemplatesBtn.disabled = false);
      exportClipboardBtn && (exportClipboardBtn.disabled = false);
      updateExportMenuButtonState();
      removeCommonPageBtn && (removeCommonPageBtn.disabled = false);
      publishSelectedBtn && (publishSelectedBtn.disabled = true);
      clearNodeSelectionBtn && (clearNodeSelectionBtn.disabled = true);
      pendingControlMeta.clear();
      pendingOpenNodes.clear();
      publishedControls?.resetPageOptions?.();
      renderList(state.parseResult.entries, state.parseResult.order);
      refreshPageTabs?.();
      updateRemoveSelectedState();
      hideDetailDrawer();
      clearHighlights?.();
      nodesPane?.clearNodeSelection?.();
      nodesPane?.parseAndRenderNodes?.();
      updateUndoRedoState();
      updateUtilityActionsState();
      updateExportButtonLabelFromPath(state.originalFilePath);
    } finally {
      suppressDocDirty = prevSuppress;
    }
  }

  function switchToDocument(docId) {
    if (!docId || docId === activeDocId) return;
    const current = getActiveDocument();
    if (current) storeDocumentSnapshot(current);
    const next = documents.find((doc) => doc.id === docId);
    if (!next) return;
    activeDocId = next.id;
    applyDocumentSnapshot(next);
    renderDocTabs();
  }

  function registerLoadedDocument(sourceName, options = {}) {
    const createDoc = options.createDoc !== false;
    const fileName = sourceName || '';
    const name = state.parseResult?.macroName || fileName || 'Untitled';
    if (!createDoc) {
      if (!getActiveDocument()) {
        const doc = addDocument(createDocumentMeta({ name, fileName }), true);
        storeDocumentSnapshot(doc);
        return doc;
      }
      updateActiveDocumentMeta({ name, fileName });
      const doc = getActiveDocument();
      storeDocumentSnapshot(doc);
      return doc;
    }
    const active = getActiveDocument();
    if (active && !active.snapshot) {
      updateActiveDocumentMeta({ name, fileName });
      storeDocumentSnapshot(active);
      return active;
    }
    const doc = addDocument(createDocumentMeta({ name, fileName }), true);
    storeDocumentSnapshot(doc);
    return doc;
  }

  function renderDocTabs() {
    if (!docTabsEl) return;
    if (documents.length === 0) {
      if (docTabsWrap) docTabsWrap.hidden = true;
      docTabsEl.innerHTML = '';
      updateDocExportPathDisplay();
      return;
    }
    if (docTabsWrap) docTabsWrap.hidden = false;
    docTabsEl.innerHTML = '';
    if (!docTabsEl.dataset.docDnD) {
      docTabsEl.dataset.docDnD = '1';
      docTabsEl.addEventListener('dragover', (ev) => {
        if (!draggingDocId) return;
        const target = ev.target;
        if (target && target.closest && target.closest('.doc-tab')) return;
        ev.preventDefault();
      });
      docTabsEl.addEventListener('drop', (ev) => {
        if (!draggingDocId) return;
        const target = ev.target;
        if (target && target.closest && target.closest('.doc-tab')) return;
        ev.preventDefault();
        reorderDocuments(draggingDocId, null, true);
        draggingDocId = null;
      });
    }
    if (!docTabsEl.dataset.docScroll) {
      docTabsEl.dataset.docScroll = '1';
      docTabsEl.addEventListener('scroll', () => updateDocTabsOverflow());
      window.addEventListener('resize', () => updateDocTabsOverflow());
      docTabsPrev?.addEventListener('click', () => scrollDocTabs(-1));
      docTabsNext?.addEventListener('click', () => scrollDocTabs(1));
    }
    documents.forEach((doc) => {
      const wrap = document.createElement('div');
      wrap.className = `doc-tab${doc.id === activeDocId ? ' active' : ''}${doc.selected ? ' selected' : ''}`;
      wrap.dataset.docId = doc.id;
      wrap.draggable = true;
      wrap.addEventListener('dragstart', (ev) => {
        draggingDocId = doc.id;
        wrap.classList.add('dragging');
        ev.dataTransfer?.setData('text/plain', doc.id);
      });
      wrap.addEventListener('dragend', () => {
        draggingDocId = null;
        wrap.classList.remove('dragging');
        docTabsEl.querySelectorAll('.doc-tab.drag-over').forEach(el => el.classList.remove('drag-over'));
      });
      wrap.addEventListener('dragover', (ev) => {
        if (!draggingDocId || draggingDocId === doc.id) return;
        ev.preventDefault();
        wrap.classList.add('drag-over');
      });
      wrap.addEventListener('dragleave', () => wrap.classList.remove('drag-over'));
      wrap.addEventListener('drop', (ev) => {
        if (!draggingDocId || draggingDocId === doc.id) return;
        ev.preventDefault();
        wrap.classList.remove('drag-over');
        reorderDocuments(draggingDocId, doc.id);
        draggingDocId = null;
      });
      const labelBtn = document.createElement('button');
      labelBtn.type = 'button';
      labelBtn.className = 'doc-tab-label';
      const label = doc.name || doc.fileName || 'Untitled';
      const lastExport = doc.snapshot && doc.snapshot.lastExportPath ? doc.snapshot.lastExportPath : '';
      const exportPath = doc.snapshot && doc.snapshot.exportFolder ? doc.snapshot.exportFolder : '';
      const exportLabel = lastExport || exportPath || '';
      labelBtn.textContent = doc.isDirty ? `${label} *` : label;
      labelBtn.title = `${doc.fileName || doc.name || 'Untitled'} — Export: ${exportLabel || 'Default (Fusion Templates)'}`;
      labelBtn.addEventListener('click', (ev) => {
        if (ev.ctrlKey || ev.metaKey) {
          ev.preventDefault();
          ev.stopPropagation();
          toggleDocSelection(doc.id);
          return;
        }
        clearDocSelections();
        switchToDocument(doc.id);
      });
      labelBtn.addEventListener('dblclick', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        promptRenameDocument(doc.id);
      });
      labelBtn.addEventListener('contextmenu', (ev) => {
        ev.preventDefault();
        openDocContextMenu(doc.id, ev.clientX, ev.clientY);
      });
      labelBtn.draggable = false;
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'doc-tab-close';
      closeBtn.textContent = 'x';
      closeBtn.title = 'Close';
      closeBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        closeDocument(doc.id);
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
    addBtn.title = 'New macro tab';
    addBtn.addEventListener('click', () => createBlankDocument());
    addBtn.draggable = false;
    docTabsEl.appendChild(addBtn);
    updateDocExportPathDisplay();
    requestAnimationFrame(() => updateDocTabsOverflow());
  }

  function isInteractiveTarget(node) {
    if (!node) return false;
    const tag = node.tagName ? node.tagName.toUpperCase() : '';
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
    if (node.isContentEditable) return true;
    return false;
  }

  function switchDocumentByOffset(offset) {
    if (!documents.length || documents.length === 1) return;
    const currentIndex = documents.findIndex(doc => doc.id === activeDocId);
    const base = currentIndex >= 0 ? currentIndex : 0;
    const nextIndex = (base + offset + documents.length) % documents.length;
    const next = documents[nextIndex];
    if (next) switchToDocument(next.id);
  }

  function switchDocumentByIndex(index) {
    if (!documents.length) return;
    const idx = Math.max(0, Math.min(documents.length - 1, index - 1));
    const doc = documents[idx];
    if (doc) switchToDocument(doc.id);
  }

  function reorderDocuments(fromId, targetId, appendToEnd = false) {
    const fromIndex = documents.findIndex(doc => doc.id === fromId);
    if (fromIndex < 0) return;
    const [moved] = documents.splice(fromIndex, 1);
    if (!moved) return;
    let insertIndex = documents.length;
    if (!appendToEnd && targetId) {
      const targetIndex = documents.findIndex(doc => doc.id === targetId);
      if (targetIndex >= 0) insertIndex = targetIndex;
    }
    documents.splice(insertIndex, 0, moved);
    renderDocTabs();
  }

  function updateDocTabsOverflow() {
    if (!docTabsEl || !docTabsPrev || !docTabsNext) return;
    const overflow = docTabsEl.scrollWidth > docTabsEl.clientWidth + 4;
    docTabsPrev.hidden = !overflow;
    docTabsNext.hidden = !overflow;
    if (!overflow) return;
    const left = docTabsEl.scrollLeft;
    const maxLeft = docTabsEl.scrollWidth - docTabsEl.clientWidth - 2;
    docTabsPrev.disabled = left <= 2;
    docTabsNext.disabled = left >= maxLeft;
  }

  function scrollDocTabs(direction) {
    if (!docTabsEl) return;
    const amount = Math.max(120, Math.floor(docTabsEl.clientWidth * 0.6));
    docTabsEl.scrollBy({ left: direction * amount, behavior: 'smooth' });
  }

  function updateDocExportPathDisplay() {
    if (!docExportPathEl) return;
    if (!documents.length || !activeDocId) {
      docExportPathEl.hidden = true;
      docExportPathEl.textContent = '';
      return;
    }
    const doc = getActiveDocument();
    const lastPath = (doc && doc.snapshot && doc.snapshot.lastExportPath) ? doc.snapshot.lastExportPath : (state.lastExportPath || '');
    const folderPath = (doc && doc.snapshot && doc.snapshot.exportFolder) ? doc.snapshot.exportFolder : (state.exportFolder || '');
    const path = lastPath || folderPath;
    const label = path ? `Export path: ${path}` : 'Export path: Default (Fusion Templates)';
    docExportPathEl.textContent = label;
    docExportPathEl.hidden = false;
  }

  function promptRenameDocument(docId) {
    const doc = documents.find(d => d.id === docId);
    if (!doc) return;
    const current = doc.name || doc.fileName || 'Untitled';
    const nextRaw = window.prompt('Rename macro tab:', current);
    if (nextRaw == null) return;
    const next = String(nextRaw).trim();
    if (!next || next === current) return;
    if (docId === activeDocId && state.parseResult) {
      state.parseResult.macroName = next;
      macroNameEl.textContent = next;
      updateActiveDocumentMeta({ name: next });
      markActiveDocumentDirty();
    } else {
      doc.name = next;
      if (doc.snapshot && doc.snapshot.parseResult) {
        doc.snapshot.parseResult.macroName = next;
      }
      doc.isDirty = true;
    }
    renderDocTabs();
  }

  function closeDocContextMenu() {
    if (docContextMenuCleanup) {
      docContextMenuCleanup();
      docContextMenuCleanup = null;
    }
    if (docContextMenu) {
      try { docContextMenu.remove(); } catch (_) {}
      docContextMenu = null;
    }
  }

  function openDocContextMenu(docId, x, y) {
    closeDocContextMenu();
    const doc = documents.find(d => d.id === docId);
    if (!doc) return;
    const menu = document.createElement('div');
    menu.className = 'doc-tab-menu';
    menu.setAttribute('role', 'menu');
    const addItem = (label, onClick, opts = {}) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'doc-tab-menu-item' + (opts.danger ? ' danger' : '');
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeDocContextMenu();
        onClick?.();
      });
      menu.appendChild(btn);
    };
    const addSeparator = () => {
      const sep = document.createElement('div');
      sep.className = 'doc-tab-menu-sep';
      menu.appendChild(sep);
    };
    addItem('New tab', () => createBlankDocument());
    addItem('Rename...', () => promptRenameDocument(docId));
    addItem(doc.selected ? 'Unselect for Export' : 'Select for Export', () => toggleDocSelection(docId));
    addItem('Clear Selection', () => clearDocSelections());
    addSeparator();
    addItem('Close tab', () => closeDocument(docId), { danger: true });
    addItem('Close others', () => closeDocumentsByFilter(d => d.id !== docId, 'Close other tabs without exporting?'), { danger: true });
    addItem('Close all', () => closeDocumentsByFilter(() => true, 'Close all tabs without exporting?'), { danger: true });

    document.body.appendChild(menu);
    docContextMenu = menu;
    const pad = 8;
    const placeMenu = () => {
      const rect = menu.getBoundingClientRect();
      let left = Number.isFinite(x) ? x : pad;
      let top = Number.isFinite(y) ? y : pad;
      if (left + rect.width > window.innerWidth - pad) left = window.innerWidth - rect.width - pad;
      if (top + rect.height > window.innerHeight - pad) top = window.innerHeight - rect.height - pad;
      if (left < pad) left = pad;
      if (top < pad) top = pad;
      menu.style.left = `${left}px`;
      menu.style.top = `${top}px`;
    };
    placeMenu();
    requestAnimationFrame(placeMenu);

    const onMouseDown = (ev) => {
      if (!menu.contains(ev.target)) closeDocContextMenu();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') closeDocContextMenu();
    };
    const onViewportChange = () => closeDocContextMenu();
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('scroll', onViewportChange, true);
    docContextMenuCleanup = () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onViewportChange);
      window.removeEventListener('scroll', onViewportChange, true);
    };
  }

  function updateExportMenuButtonState() {
    if (!exportMenuBtn) return;
    exportMenuBtn.disabled = !state.parseResult;
    if (exportMenuBtn.disabled) closeExportMenu();
  }

  function closeExportMenu() {
    if (exportMenuCleanup) {
      exportMenuCleanup();
      exportMenuCleanup = null;
    }
    if (!exportMenuEl) return;
    exportMenuEl.hidden = true;
    exportMenuEl.innerHTML = '';
  }

  function getLooseMacroCount() {
    return documents.filter((doc) => {
      if (!doc || !doc.snapshot || !doc.snapshot.parseResult) return false;
      const linked = doc.snapshot.drfxLink && doc.snapshot.drfxLink.linked;
      return !linked;
    }).length;
  }

  function isDocDrfxLinked(doc) {
    if (!doc) return false;
    const snap = doc.snapshot;
    if (snap && snap.drfxLink && snap.drfxLink.linked && snap.originalFilePath) return true;
    if (doc.id === activeDocId) {
      return !!(state.drfxLink && state.drfxLink.linked && state.originalFilePath);
    }
    return false;
  }

  function getDrfxDocumentsFrom(list) {
    return list.filter((doc) => {
      if (!doc) return false;
      if (doc.snapshot && doc.snapshot.parseResult) {
        return isDocDrfxLinked(doc);
      }
      if (doc.id === activeDocId && state.parseResult) {
        return isDocDrfxLinked(doc);
      }
      return false;
    });
  }

  function buildExportMenuItems() {
    const hasMacro = !!state.parseResult;
    const activeDrfxLinked = !!(state.drfxLink && state.drfxLink.linked);
    const selectedDocs = documents.filter(doc => doc && doc.selected);
    const selection = selectedDocs.length ? selectedDocs : documents;
    const drfxDocs = getDrfxDocumentsFrom(selection);
    const looseCount = getLooseMacroCount();
    const showDrfxMulti = looseCount >= 2;
    const showSourceBulk = drfxDocs.length >= 2;
    const sourceBulkLabel = selectedDocs.length ? 'To Source DRFX (Selected Tabs)' : 'To Source DRFX (All Tabs)';
    const items = [
      { id: 'clipboard', label: 'To Clipboard', enabled: hasMacro, action: exportToClipboard },
      { id: 'file', label: 'To File', enabled: hasMacro, action: exportToFile },
      { id: 'edit', label: 'To Edit Page', enabled: hasMacro, action: exportToEditPage },
    ];
    if (activeDrfxLinked || showDrfxMulti || showSourceBulk) items.push({ type: 'sep' });
    if (activeDrfxLinked) {
      items.push({ id: 'source', label: 'To Source DRFX', enabled: hasMacro, action: exportToSourceDrfx });
    }
    if (showSourceBulk) {
      items.push({ id: 'source-bulk', label: sourceBulkLabel, enabled: hasMacro, action: exportToSourceDrfxMulti });
    }
    if (showDrfxMulti) {
      items.push({ id: 'drfx', label: 'To DRFX', enabled: hasMacro, action: exportToDrfxMulti });
    }
    return items;
  }

  function openExportMenu() {
    if (!exportMenuEl || !exportMenuBtn) return;
    closeExportMenu();
    const items = buildExportMenuItems();
    exportMenuEl.hidden = false;
    exportMenuEl.innerHTML = '';
    items.forEach((item) => {
      if (item.type === 'sep') {
        const sep = document.createElement('div');
        sep.className = 'menu-sep';
        exportMenuEl.appendChild(sep);
        return;
      }
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = item.label;
      btn.disabled = !item.enabled;
      btn.addEventListener('click', () => {
        closeExportMenu();
        item.action?.();
      });
      exportMenuEl.appendChild(btn);
    });
    const onMouseDown = (ev) => {
      if (exportMenuEl && exportMenuEl.contains(ev.target)) return;
      if (exportMenuBtn && exportMenuBtn.contains(ev.target)) return;
      closeExportMenu();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') closeExportMenu();
    };
    const onViewportChange = () => closeExportMenu();
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('scroll', onViewportChange, true);
    exportMenuCleanup = () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onViewportChange);
      window.removeEventListener('scroll', onViewportChange, true);
    };
  }

  exportMenuBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    if (exportMenuEl && !exportMenuEl.hidden) {
      closeExportMenu();
      return;
    }
    openExportMenu();
  });

  function setIntroCollapsed(collapsed) {
    if (!introPanel) return;
    introPanel.hidden = collapsed;
    if (introToggleBtn) {
      introToggleBtn.setAttribute('aria-pressed', collapsed ? 'true' : 'false');
    }
    if (!collapsed) {
      if (typeof clearSelections === 'function') clearSelections();
      hideDetailDrawer();
    }
  }

  introToggleBtn?.addEventListener('click', () => {
    if (!introPanel) return;
    setIntroCollapsed(!introPanel.hidden);
  });

  try {
    window.addEventListener('beforeunload', (ev) => {
      try {
        if (!documents.length) return;
        const dirty = documents.some(doc => doc && doc.isDirty);
        if (!dirty) return;
        const ok = window.confirm('Close without exporting? Unsaved changes will be lost.');
        if (!ok) {
          ev.preventDefault();
          ev.returnValue = '';
        }
      } catch (_) {}
    });
  } catch (_) {}

  try {
    const keyHandler = (ev) => {
      try {
        if (!documents.length) return;
        if (!(ev.ctrlKey || ev.metaKey)) return;
        if (isInteractiveTarget(ev.target)) return;
        const key = ev.key ? ev.key.toLowerCase() : '';
        if (key >= '1' && key <= '9') {
          ev.preventDefault();
          switchDocumentByIndex(parseInt(key, 10));
          return;
        }
        if (key === 'tab') {
          ev.preventDefault();
          switchDocumentByOffset(ev.shiftKey ? -1 : 1);
          return;
        }
        if (key === 'w') {
          ev.preventDefault();
          closeDocument(activeDocId);
          return;
        }
        if (key === 't') {
          ev.preventDefault();
          createBlankDocument();
        }
      } catch (_) {}
    };
    window.addEventListener('keydown', keyHandler);
  } catch (_) {}

  const controlsSection = document.getElementById('controlsSection');

  const controlsList = document.getElementById('controlsList');

  const publishedSearch = document.getElementById('publishedSearch');
  const pageTabsEl = document.getElementById('pageTabs');

  // Nodes pane elements

  const nodesList = document.getElementById('nodesList');

  const nodesSearch = document.getElementById('nodesSearch');
  const addUtilityNodeBtn = document.getElementById('addUtilityNodeBtn');
  const autoUtilityNodeToggle = document.getElementById('autoUtilityNode');
  const detailDrawer = document.getElementById('detailDrawer');
  const detailDrawerTitle = document.getElementById('detailDrawerTitle');
  const detailDrawerSubtitle = document.getElementById('detailDrawerSubtitle');
  const detailDrawerBody = document.getElementById('detailDrawerBody');
  const detailDrawerToggleBtn = document.getElementById('detailDrawerToggleBtn');
  const codePane = document.getElementById('codePane');
  const codePaneToggleInput = document.getElementById('codePaneToggle');
  const codePaneBody = document.getElementById('codePaneBody');
  const codeView = document.getElementById('codeView');
  const catalogEditorRoot = document.getElementById('catalogEditor');
  const catalogEditorOpenBtn = document.getElementById('openCatalogEditorBtn');
  const catalogEditorCloseBtn = document.getElementById('closeCatalogEditorBtn');
  const catalogDatasetSelect = document.getElementById('catalogDatasetSelect');
  const catalogTypeSearch = document.getElementById('catalogTypeSearch');
  const catalogTypeList = document.getElementById('catalogTypeList');
  const catalogControlList = document.getElementById('catalogControlList');
  const catalogEmptyState = document.getElementById('catalogEmptyState');
  const catalogDownloadBtn = document.getElementById('catalogDownloadBtn');
  const addControlModal = document.getElementById('addControlModal');
  const addControlForm = document.getElementById('addControlForm');
  const addControlTitle = document.getElementById('addControlTitle');
  const addControlNameInput = document.getElementById('addControlName');
  const addControlTypeSelect = document.getElementById('addControlType');
  const addLabelOptions = document.getElementById('addLabelOptions');
  const addControlLabelCountInput = document.getElementById('addControlLabelCount');
  const addControlLabelDefaultSelect = document.getElementById('addControlLabelDefault');
  const addControlPageInput = document.getElementById('addControlPage');
  const addControlPageOptions = document.getElementById('addControlPageOptions');
  const addControlError = document.getElementById('addControlError');
  const addControlCancelBtn = document.getElementById('addControlCancelBtn');
  const addControlCloseBtn = document.getElementById('addControlCloseBtn');
  const addControlSubmitBtn = document.getElementById('addControlSubmitBtn');
  let nodesPane = null;
  let historyController = null;
  let detailDrawerCollapsed = false;
  let codePaneCollapsed = true;
  let codePaneRefreshTimer = null;
  let codeHighlightActive = false;
  let activeDetailEntryIndex = null;
  let renderIconHtml = () => '';
  const pushHistory = (...args) => {
    historyController?.pushHistory?.(...args);
    if (!suppressDocDirty) markActiveDocumentDirty();
  };
  const updateUndoRedoState = () => historyController?.updateUndoRedoState?.();
  const ENABLE_NODE_DRAG = true; // enable cross-pane node dragging
  const pendingControlMeta = new Map();
  const pendingOpenNodes = new Set();

  const AUTO_UTILITY_KEY = 'fmr.autoUtilityNode';
  const UTILITY_TEMPLATE_PATH = 'templates/utility-node.setting';
  let autoUtilityEnabled = false;
  let utilityTemplateText = null;

  const removeBtn = document.getElementById('removeSelected');

  const getSelectedSet = () => (state.parseResult?.selected instanceof Set ? state.parseResult.selected : new Set());
  const pendingMetaKey = (op, id) => `${op || ''}::${id || ''}`;
  const rememberPendingControlMeta = (op, id, meta) => {
    if (!op || !id || !meta) return;
    pendingControlMeta.set(pendingMetaKey(op, id), meta);
  };
  const getPendingControlMeta = (op, id) => {
    if (!op || !id) return null;
    return pendingControlMeta.get(pendingMetaKey(op, id)) || null;
  };
  const consumePendingControlMeta = (op, id) => {
    if (!op || !id) return;
    pendingControlMeta.delete(pendingMetaKey(op, id));
  };

  function captureHistoryExtras() {
    try {
      const extras = {
        originalText: state.originalText || '',
        newline: state.newline || '\n',
      };
      if (state.parseResult) {
        extras.pageOrder = Array.from(state.parseResult.pageOrder || []);
        extras.activePage = state.parseResult.activePage || null;
        extras.macroName = state.parseResult.macroName || null;
        extras.macroNameOriginal = state.parseResult.macroNameOriginal || null;
        extras.operatorType = state.parseResult.operatorType || null;
        extras.operatorTypeOriginal = state.parseResult.operatorTypeOriginal || null;
      }
      return extras;
    } catch (_) {
      return { originalText: state.originalText || '', newline: state.newline || '\n' };
    }
  }

  function restoreHistoryExtras(extra = {}, context) {
    try {
      if (typeof extra.originalText === 'string') {
        state.originalText = extra.originalText;
        state.newline = extra.newline || detectNewline(extra.originalText);
        updateCodeView(state.originalText || '');
      }
      if (state.parseResult) {
        if (Array.isArray(extra.pageOrder)) state.parseResult.pageOrder = [...extra.pageOrder];
        if (extra.activePage !== undefined) state.parseResult.activePage = extra.activePage;
        if (extra.macroName !== undefined) state.parseResult.macroName = extra.macroName;
        if (extra.macroNameOriginal !== undefined) state.parseResult.macroNameOriginal = extra.macroNameOriginal;
        if (extra.operatorType !== undefined) state.parseResult.operatorType = extra.operatorType;
        if (extra.operatorTypeOriginal !== undefined) state.parseResult.operatorTypeOriginal = extra.operatorTypeOriginal;
      }
      macroNameEl.textContent = state.parseResult?.macroName || 'Unknown';
      syncOperatorSelect();
      refreshPageTabs?.();
      runValidation(context || 'history');
    } catch (_) {}
  }

  try {
    if (typeof localStorage !== 'undefined') {
      autoUtilityEnabled = localStorage.getItem(AUTO_UTILITY_KEY) === '1';
    }
  } catch (_) {
    autoUtilityEnabled = false;
  }
  if (autoUtilityNodeToggle) {
    autoUtilityNodeToggle.checked = !!autoUtilityEnabled;
    autoUtilityNodeToggle.addEventListener('change', () => {
      setAutoUtilityEnabled(!!autoUtilityNodeToggle.checked);
    });
  }


  // Insert a cross-platform URL launcher into a ButtonControl block if it has none
  function insertLaunchCodeInRange(text, openBraceIndex, closeBraceIndex, url, eol) {
    try {
      const escLua = (s) => String(s).replace(/\\/g, "\\\\").replace(/\"/g, '\\"');
      const lua =
        'local url = "' + escLua(url) + '"\\n' +
        'local sep = package.config:sub(1,1)\\n' +
        'if sep == "\\\\" then\\n' +
        '  os.execute(\'start "" "\'..url..\'"\')\\n' +
        'else\\n' +
        '  local uname = io.popen("uname"):read("*l")\\n' +
        '  if uname == "Darwin" then\\n' +
        '    os.execute(\'open "\'..url..\'"\')\\n' +
        '  else\\n' +
        '    os.execute(\'xdg-open "\'..url..\'"\')\\n' +
        '  end\\n' +
        'end\\n';
      const indent = (function(){ try { return (getLineIndent(text, openBraceIndex) || '') + '\\t'; } catch(_) { return '\\t'; } })();
      const insertLine = indent + 'BTNCS_Execute = "' + lua + '",' + eol;
      return text.slice(0, closeBraceIndex) + eol + insertLine + text.slice(closeBraceIndex);
    } catch(_) { return text; }
  }

  const importCatalogBtn = document.getElementById('importCatalogBtn');

  const catalogInput = document.getElementById('catalogInput');

  const importModifierCatalogBtn = document.getElementById('importModifierCatalogBtn');

  const modifierCatalogInput = document.getElementById('modifierCatalogInput');

  const publishSelectedBtn = document.getElementById('publishSelectedBtn');

  const clearNodeSelectionBtn = document.getElementById('clearNodeSelectionBtn');

  const hideReplacedEl = document.getElementById('hideReplaced');

  const macroNameEl = document.getElementById('macroName');

  const ctrlCountEl = document.getElementById('ctrlCount');
  const operatorTypeSelect = document.getElementById('operatorType');

  const exportBtn = document.getElementById('exportBtn');
  const exportTemplatesBtn = document.getElementById('exportTemplatesBtn');

  const exportClipboardBtn = document.getElementById('exportClipboardBtn');
  const importClipboardBtn = document.getElementById('importClipboardBtn');
  let openNativeBtn = document.getElementById('openNativeBtn');

  const removeCommonPageBtn = document.getElementById('removeCommonPageBtn');
  const deselectAllBtn = document.getElementById('deselectAllBtn');

  const undoBtn = document.getElementById('undoBtn');

  const redoBtn = document.getElementById('redoBtn');

  const resetBtn = document.getElementById('resetOrder');

  const DEFAULT_EXPORT_LABEL = 'Export reordered .setting';
  const DRFX_EXPORT_LABEL = 'Export to Source';

  const messages = document.getElementById('messages');

  const diagnostics = document.getElementById('diagnostics');
  if (diagnostics && LOG_DIAGNOSTICS) diagnostics.hidden = false;
  const diagnosticsController = createDiagnosticsController({
    element: diagnostics,
    defaultEnabled: LOG_DIAGNOSTICS,
  });
  const toggleDiagBtn = document.getElementById('toggleDiagnosticsBtn');
  const logDiag = (...args) => {
    try { diagnosticsController.log(...args); } catch (_) {}
  };
  const logTag = (...args) => {
    try { diagnosticsController.logTag(...args); } catch (_) {}
  };
  function setDiagnosticsEnabled(on) {
    try {
      diagnosticsController.setEnabled(on);
      updateDiagnosticsButton();
    } catch (_) {}
  }
  function updateDiagnosticsButton() {
    try {
      if (!toggleDiagBtn) return;
      const enabled = diagnosticsController.isEnabled();
      toggleDiagBtn.checked = enabled;
      toggleDiagBtn.setAttribute('aria-checked', enabled ? 'true' : 'false');
    } catch (_) {}
  }
  try {
    if (toggleDiagBtn) {
      toggleDiagBtn.addEventListener('change', () => {
        try { setDiagnosticsEnabled(!!toggleDiagBtn.checked); } catch (_) {}
      });
      toggleDiagBtn.setAttribute('role', 'switch');
      updateDiagnosticsButton();
    }
  } catch (_) {}

  let pendingAddControlNode = null;

  function getKnownPageNames() {
    const set = new Set();
    set.add('Controls');
    try {
      if (Array.isArray(state.parseResult?.pageOrder)) {
        state.parseResult.pageOrder.forEach((page) => {
          const val = (page && String(page).trim()) ? String(page).trim() : '';
          if (val) set.add(val);
        });
      }
      if (Array.isArray(state.parseResult?.entries)) {
        state.parseResult.entries.forEach((entry) => {
          const page = entry && entry.page ? String(entry.page).trim() : '';
          if (page) set.add(page);
        });
      }
    } catch (_) {}
    return Array.from(set);
  }

  function updateAddControlPageOptionsList() {
    if (!addControlPageOptions) return;
    addControlPageOptions.innerHTML = '';
    const pages = getKnownPageNames();
    pages.forEach((page) => {
      const opt = document.createElement('option');
      opt.value = page;
      addControlPageOptions.appendChild(opt);
    });
  }

  function getSuggestedAddControlPage() {
    const active = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
    if (active) return active;
    if (Array.isArray(state.parseResult?.pageOrder) && state.parseResult.pageOrder.length) {
      return state.parseResult.pageOrder[0];
    }
    return 'Controls';
  }

  function updateAddControlTypeVisibility() {
    if (!addLabelOptions) return;
    const typeVal = (addControlTypeSelect?.value || 'label').toLowerCase();
    addLabelOptions.hidden = typeVal !== 'label';
  }

  function resetAddControlFormFields(nodeName) {
    if (addControlTitle) addControlTitle.textContent = nodeName || 'Node';
    if (addControlNameInput) addControlNameInput.value = '';
    if (addControlTypeSelect) addControlTypeSelect.value = 'label';
    if (addControlLabelCountInput) addControlLabelCountInput.value = '0';
    if (addControlLabelDefaultSelect) addControlLabelDefaultSelect.value = 'closed';
    if (addControlPageInput) {
      const suggested = getSuggestedAddControlPage();
      addControlPageInput.value = suggested && suggested !== 'Controls' ? suggested : '';
    }
    if (addControlError) addControlError.textContent = '';
    updateAddControlTypeVisibility();
  }

  function openAddControlDialog(nodeName) {
    if (!addControlModal) return;
    pendingAddControlNode = nodeName;
    updateAddControlPageOptionsList();
    resetAddControlFormFields(nodeName);
    addControlModal.hidden = false;
    setTimeout(() => {
      try { addControlNameInput?.focus(); } catch (_) {}
    }, 0);
  }

  function closeAddControlDialog() {
    if (!addControlModal) return;
    pendingAddControlNode = null;
    addControlModal.hidden = true;
  }

  addControlTypeSelect?.addEventListener('change', () => updateAddControlTypeVisibility());
  addControlCancelBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    closeAddControlDialog();
  });
  addControlCloseBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    closeAddControlDialog();
  });
  addControlModal?.addEventListener('click', (ev) => {
    if (ev.target === addControlModal) closeAddControlDialog();
  });
  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && addControlModal && !addControlModal.hidden) {
      closeAddControlDialog();
    }
  });
  if (addControlForm) {
    addControlForm.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      try {
        if (!pendingAddControlNode) {
          if (addControlError) addControlError.textContent = 'Select a node before adding controls.';
          return;
        }
        const typeRaw = (addControlTypeSelect?.value || 'label').toLowerCase();
        const allowedTypes = new Set(['label', 'separator', 'button', 'slider', 'screw']);
        const type = allowedTypes.has(typeRaw) ? typeRaw : 'label';
        let name = (addControlNameInput?.value || '').trim();
        if (!name) {
          if (type === 'separator') {
            name = 'Separator';
          } else {
            if (addControlError) addControlError.textContent = 'Control name is required.';
            addControlNameInput?.focus();
            return;
          }
        }
        const pageValue = (addControlPageInput?.value || '').trim();
        const config = {
          name,
          type,
          page: pageValue,
        };
        if (type === 'label') {
          const countVal = parseInt(addControlLabelCountInput?.value || '0', 10);
          config.labelCount = Number.isFinite(countVal) ? countVal : 0;
          config.labelDefault = (addControlLabelDefaultSelect?.value === 'open') ? 'open' : 'closed';
        }
        if (addControlError) addControlError.textContent = '';
        if (addControlSubmitBtn) addControlSubmitBtn.disabled = true;
        await addControlToNode(pendingAddControlNode, config);
        closeAddControlDialog();
      } catch (err) {
        if (addControlError) addControlError.textContent = err?.message || 'Unable to add control.';
      } finally {
        if (addControlSubmitBtn) addControlSubmitBtn.disabled = false;
      }
    });
  }

  function hideDetailDrawer() {
    activeDetailEntryIndex = null;
    if (!detailDrawer) return;
    detailDrawer.hidden = true;
    detailDrawer.classList.remove('open');
    detailDrawer.classList.remove('collapsed');
    detailDrawerCollapsed = false;
  }

  function renderDetailDrawer(index) {
    if (!detailDrawer || !detailDrawerTitle || !detailDrawerSubtitle || !detailDrawerBody) return;
    if (!state.parseResult || !Array.isArray(state.parseResult.entries) || !state.parseResult.entries[index]) {
      hideDetailDrawer();
      return;
    }
    const wasHidden = detailDrawer.hidden;
    activeDetailEntryIndex = index;
    const entry = state.parseResult.entries[index];
    if (wasHidden) detailDrawerCollapsed = true;
    detailDrawer.hidden = false;
    detailDrawer.classList.add('open');
    detailDrawer.classList.toggle('collapsed', !!detailDrawerCollapsed);
    updateDrawerToggleLabel();
    detailDrawerTitle.textContent = entry.displayName || entry.name || entry.source || 'Control';
    detailDrawerTitle.contentEditable = 'true';
    detailDrawerTitle.spellcheck = false;
    const commitDrawerName = () => {
      if (typeof setEntryDisplayNameApi === 'function') {
        setEntryDisplayNameApi(index, (detailDrawerTitle.textContent || '').trim());
      }
    };
    detailDrawerTitle.onblur = commitDrawerName;
    detailDrawerTitle.onkeydown = (ev) => {
      if (ev.key === 'Enter') { ev.preventDefault(); commitDrawerName(); detailDrawerTitle.blur(); }
      if (ev.key === 'Escape') {
        ev.preventDefault();
        detailDrawerTitle.textContent = entry.displayName || entry.name || entry.source || 'Control';
        detailDrawerTitle.blur();
      }
    };
    detailDrawerSubtitle.textContent = `${entry.sourceOp || 'Unknown'}${entry.source ? '.' + entry.source : ''}`;
    detailDrawerBody.innerHTML = '';
    const buildCollapsibleField = (title, defaultOpen = false) => {
      const field = document.createElement('div');
      field.className = 'detail-field detail-collapsible';
      const header = document.createElement('button');
      header.type = 'button';
      header.className = 'detail-collapsible-header';
      const iconSpan = document.createElement('span');
      iconSpan.className = 'icon';
      if (typeof createIcon === 'function') iconSpan.innerHTML = createIcon('chevron-right', 12);
      else iconSpan.textContent = '▶';
      const titleSpan = document.createElement('span');
      titleSpan.textContent = title;
      header.appendChild(iconSpan);
      header.appendChild(titleSpan);
      const body = document.createElement('div');
      body.className = 'detail-collapsible-body';
      let open = !!defaultOpen;
      const sync = () => {
        body.hidden = !open;
        field.classList.toggle('open', open);
        if (typeof createIcon === 'function') {
          iconSpan.innerHTML = createIcon(open ? 'chevron-down' : 'chevron-right', 12);
        } else {
          iconSpan.textContent = open ? '▼' : '▶';
        }
      };
      header.addEventListener('click', () => {
        open = !open;
        sync();
      });
      sync();
      field.appendChild(header);
      field.appendChild(body);
      return { field, body };
    };
    const meta = entry.controlMeta || {};
    const currentType = normalizeInputControlValue(meta.inputControl);
    const infoRow = document.createElement('div');
    infoRow.className = 'detail-info-row';
    const infoColumn = (label, value) => {
      const wrap = document.createElement('div');
      wrap.className = 'detail-info';
      const lbl = document.createElement('span');
      lbl.className = 'detail-info-label';
      lbl.textContent = label;
      const val = document.createElement('span');
      val.className = 'detail-info-value';
      val.textContent = value || '—';
      wrap.appendChild(lbl);
      wrap.appendChild(val);
      return wrap;
    };
    const pageDisplay = infoColumn('Page', entry.page || 'Controls');
    infoRow.appendChild(pageDisplay);
    const effectiveDataType = entry.isButton ? '"Number"' : (meta.dataType || null);
    const dataDisplay = infoColumn('Data Type', effectiveDataType ? String(effectiveDataType).replace(/"/g, '') : 'Unknown');
    infoRow.appendChild(dataDisplay);
    const controlWrap = document.createElement('div');
    controlWrap.className = 'detail-info detail-info-control';
    const controlLabel = document.createElement('span');
    controlLabel.className = 'detail-info-label';
    controlLabel.textContent = 'Input Control';
    controlWrap.appendChild(controlLabel);
    const select = document.createElement('select');
    select.className = 'detail-type-select';
    const allowed = getAllowedInputControls(entry);
    const knownTypes = getKnownInputControls();
    const hasKnownType = !!currentType;
    const addOption = (value) => {
      if (!value) return;
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = value;
      select.appendChild(opt);
    };
    if (!hasKnownType) {
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = 'Change Control Type';
      placeholder.disabled = true;
      placeholder.selected = true;
      select.appendChild(placeholder);
    }
    let types = Array.isArray(allowed) && allowed.length ? [...allowed] : [...knownTypes];
    if (currentType && !types.includes(currentType)) {
      types.unshift(currentType);
    }
    types.forEach(addOption);
    if (currentType) select.value = currentType;
    else select.value = '';
    select.addEventListener('change', () => {
      const val = select.value;
      if (!val) return;
      setEntryInputControl(index, val);
    });
    controlWrap.appendChild(select);
    infoRow.appendChild(controlWrap);
    detailDrawerBody.appendChild(infoRow);
    const defaultsField = document.createElement('div');
    defaultsField.className = 'detail-field';
    const defaultsLabel = document.createElement('label');
    defaultsLabel.textContent = entry.isLabel ? 'Default state' : 'Defaults';
    defaultsField.appendChild(defaultsLabel);
    if (entry.isLabel) {
      const toggleWrap = document.createElement('div');
      toggleWrap.className = 'detail-label-toggle';
      const toggleLabel = document.createElement('label');
      toggleLabel.className = 'detail-switch';
      const toggleInput = document.createElement('input');
      toggleInput.type = 'checkbox';
      const currentDefault = normalizeMetaValue(meta.defaultValue);
      toggleInput.checked = currentDefault === '0' ? false : true;
      const slider = document.createElement('span');
      slider.className = 'slider';
      toggleLabel.appendChild(toggleInput);
      toggleLabel.appendChild(slider);
      const stateText = document.createElement('span');
      stateText.className = 'detail-label-state';
      const applyState = () => {
        stateText.textContent = toggleInput.checked ? 'Open' : 'Closed';
      };
      toggleInput.addEventListener('change', () => {
        applyState();
        setEntryDefaultValue(index, toggleInput.checked ? '1' : '0');
      });
      applyState();
      toggleWrap.appendChild(toggleLabel);
      toggleWrap.appendChild(stateText);
      defaultsField.appendChild(toggleWrap);
    } else {
      const defaultRow = document.createElement('div');
      defaultRow.className = 'detail-default-row';
      const rangeRow = document.createElement('div');
      rangeRow.className = 'detail-default-row';
      const buildDefaultInput = (title, key, handler) => {
        const group = document.createElement('div');
        group.className = 'detail-default-item';
        const lbl = document.createElement('span');
        lbl.textContent = title;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = normalizeMetaValue(meta[key]) || '';
        const originalMeta = entry.controlMetaOriginal || {};
        if (originalMeta[key] != null) input.placeholder = `Original: ${originalMeta[key]}`;
        const commit = () => handler(input.value);
        input.addEventListener('keydown', (ev) => {
          if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
        });
        input.addEventListener('blur', commit);
        group.appendChild(lbl);
        group.appendChild(input);
        return group;
      };
      defaultRow.appendChild(buildDefaultInput('Default', 'defaultValue', (val) => setEntryDefaultValue(index, val)));
      rangeRow.appendChild(buildDefaultInput('Min', 'minScale', (val) => setEntryRangeValue(index, 'minScale', val)));
      rangeRow.appendChild(buildDefaultInput('Max', 'maxScale', (val) => setEntryRangeValue(index, 'maxScale', val)));
      defaultsField.appendChild(defaultRow);
      defaultsField.appendChild(rangeRow);
    }
    detailDrawerBody.appendChild(defaultsField);
    const onChangeHasValue = !!(entry.onChange && String(entry.onChange).trim());
    const onChangeSection = buildCollapsibleField('On-Change Script', onChangeHasValue);
    const onChangeBody = onChangeSection.body;
    const onChangeArea = document.createElement('textarea');
    onChangeArea.value = entry.onChange || '';
    onChangeArea.placeholder = 'Lua script to execute when this control changes.';
    onChangeBody.appendChild(onChangeArea);
    onChangeArea.addEventListener('blur', () => {
      updateEntryOnChange(index, onChangeArea.value);
    });
    const actions = document.createElement('div');
    actions.className = 'detail-actions';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.textContent = 'Save';
    saveBtn.addEventListener('click', () => {
      updateEntryOnChange(index, onChangeArea.value);
    });
    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', () => {
      onChangeArea.value = '';
      updateEntryOnChange(index, '');
    });
    actions.appendChild(saveBtn);
    actions.appendChild(clearBtn);
    onChangeBody.appendChild(actions);
    detailDrawerBody.appendChild(onChangeSection.field);
    if (entry.isButton) {
      const hasExecuteScript = !!(entry.buttonExecute && String(entry.buttonExecute).trim());
      const executeSection = buildCollapsibleField('Button Execute Script', hasExecuteScript);
      const execBody = executeSection.body;
      const execArea = document.createElement('textarea');
      execArea.value = entry.buttonExecute || '';
      execArea.placeholder = 'Lua script executed when this button fires.';
      const applyLauncherScript = (script) => {
        const val = script || '';
        execArea.value = val;
        updateEntryButtonExecute(index, val, { silent: true, skipHistory: true });
      };
      if (typeof appendLauncherUiApi === 'function') {
        const launcherWrap = document.createElement('div');
        launcherWrap.className = 'detail-launcher-column';
        execBody.appendChild(launcherWrap);
        try {
          appendLauncherUiApi(launcherWrap, entry, {
            onScriptChange: (script) => applyLauncherScript(script || ''),
          });
        } catch (_) {
          const placeholder = document.createElement('p');
          placeholder.className = 'detail-placeholder';
          placeholder.textContent = 'Launcher configuration unavailable.';
          launcherWrap.appendChild(placeholder);
        }
      }
      execBody.appendChild(execArea);
      execArea.addEventListener('blur', () => {
        updateEntryButtonExecute(index, execArea.value);
      });
      const execActions = document.createElement('div');
      execActions.className = 'detail-actions';
      const execSave = document.createElement('button');
      execSave.type = 'button';
      execSave.textContent = 'Save';
      execSave.addEventListener('click', () => {
        updateEntryButtonExecute(index, execArea.value);
      });
      const execClear = document.createElement('button');
      execClear.type = 'button';
      execClear.textContent = 'Clear';
      execClear.addEventListener('click', () => {
        execArea.value = '';
        updateEntryButtonExecute(index, '');
      });
      execActions.appendChild(execSave);
      execActions.appendChild(execClear);
      execBody.appendChild(execActions);
      detailDrawerBody.appendChild(executeSection.field);
    }
    if (entry.isLabel) {
      const visField = document.createElement('div');
      visField.className = 'detail-field';
      visField.innerHTML = `<label>Label Visibility</label><div>${entry.labelHidden ? 'Hidden' : 'Visible'}</div>`;
      detailDrawerBody.appendChild(visField);
    }
  }

  function updateEntryOnChange(index, script) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    const newVal = script != null ? String(script) : '';
    if (entry.onChange === newVal) return;
    pushHistory('edit on-change');
    entry.onChange = newVal;
    try { renderList(state.parseResult.entries, state.parseResult.order); } catch (_) {}
    renderDetailDrawer(index);
    markContentDirty();
  }

  function updateEntryButtonExecute(index, script, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    const newVal = script != null ? String(script) : '';
    if (entry.buttonExecute === newVal) return;
    const silent = !!opts.silent;
    const skipHistory = !!opts.skipHistory;
    if (!silent && !skipHistory) pushHistory('edit button execute');
    entry.buttonExecute = newVal;
    if (!silent) {
      try { renderList(state.parseResult.entries, state.parseResult.order); } catch (_) {}
      renderDetailDrawer(index);
    }
    markContentDirty();
  }

  function normalizeMetaValue(value) {
    if (value == null) return null;
    const trimmed = String(value).trim();
    return trimmed.length ? trimmed : null;
  }

  function setEntryDefaultValue(index, value) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    const next = normalizeMetaValue(value);
    const current = entry.controlMeta.defaultValue != null ? String(entry.controlMeta.defaultValue) : null;
    if ((current || null) === (next || null)) return;
    pushHistory('edit default');
    entry.controlMeta.defaultValue = next;
    entry.controlMetaDirty = true;
    if (typeof entry.raw === 'string') {
      const open = entry.raw.indexOf('{');
      const close = entry.raw.lastIndexOf('}');
      if (open >= 0 && close > open) {
        const indent = (getLineIndent(entry.raw, open) || '') + '\t';
        let body = entry.raw.slice(open + 1, close);
        if (next == null) {
          body = removeInstanceInputProp(body, 'Default');
        } else {
          body = setInstanceInputProp(body, 'Default', next, indent, state.newline || '\n');
        }
        entry.raw = entry.raw.slice(0, open + 1) + body + entry.raw.slice(close);
      }
    }
    renderDetailDrawer(index);
    markContentDirty();
  }

  async function reloadMacroFromCurrentText(options = {}) {
    try {
      const sourceName = state.originalFileName || 'Imported.setting';
      const prevHistory = state.parseResult ? state.parseResult.history : null;
      const prevFuture = state.parseResult ? state.parseResult.future : null;
      await loadMacroFromText(sourceName, state.originalText || '', {
        allowAutoUtility: false,
        preserveFileInfo: true,
        preserveFilePath: true,
        skipClear: !!options.skipClear,
        silentAuto: true,
        createDoc: false,
      });
      if (state.parseResult) {
        if (prevHistory) state.parseResult.history = prevHistory;
        if (prevFuture) state.parseResult.future = prevFuture;
        updateUndoRedoState();
      }
    } catch (_) {
      return Promise.reject(_);
    }
  }

  function buildControlDefinitionLines(type, config = {}) {
    const name = config.name ? escapeQuotes(config.name) : 'Custom';
    const page = config.page ? `ICS_ControlPage = "${escapeQuotes(config.page)}",` : null;
    if (type === 'slider' || type === 'screw') {
      const control = type === 'slider' ? 'SliderControl' : 'ScrewControl';
      const lines = [
        `LINKS_Name = "${name}",`,
        'LINKID_DataType = "Number",',
        'INP_Integer = false,',
        'INP_Default = 0,',
        'INP_MinScale = 0,',
        'INP_MaxScale = 1,',
        'INP_MinAllowed = -1000000,',
        'INP_MaxAllowed = 1000000,',
        'INP_SplineType = "Default",',
        `INPID_InputControl = "${control}",`,
      ];
      if (page) lines.push(page);
      return lines;
    }
    if (type === 'button') {
      const lines = [
        `LINKS_Name = "${name}",`,
        'INPID_InputControl = "ButtonControl",',
      ];
      if (page) lines.push(page);
      return lines;
    }
    if (type === 'separator') {
      const lines = [
        `LINKS_Name = "${name}",`,
        'LINKID_DataType = "Number",',
        'INPID_InputControl = "SeparatorControl",',
      ];
      if (page) lines.push(page);
      return lines;
    }
    // default to label
    const lines = [
      `LINKS_Name = "${name}",`,
      'INP_Integer = false,',
      'LBLC_DropDownButton = true,',
      'LINKID_DataType = "Number",',
      `LBLC_NumInputs = ${Math.max(0, Number.isFinite(config.labelCount) ? config.labelCount : 0)},`,
      'INPID_InputControl = "LabelControl",',
      'INP_SplineType = "Default",',
      `INP_Default = ${config.labelDefault === 'closed' ? 0 : 1},`,
    ];
    if (page) lines.push(page);
    return lines;
  }

  function insertUserControlBlock(text, ucBlock, controlId, lines, eol) {
    const blockIndent = (getLineIndent(text, ucBlock.open) || '') + '\t';
    const innerIndent = blockIndent + '\t';
    let block = `${blockIndent}${controlId} = {${eol}`;
    for (const line of lines) {
      block += `${innerIndent}${line}${eol}`;
    }
    block += `${blockIndent}},${eol}`;
    const insertPos = ucBlock.open + 1;
    const before = text.slice(0, insertPos);
    const inner = text.slice(insertPos, ucBlock.close);
    const after = text.slice(ucBlock.close);
    const injection = `${eol}${block}${inner}`;
    return before + injection + after;
  }

  function doesControlIdExist(text, range, controlId) {
    if (!range) return false;
    const segment = text.slice(range.open, range.close);
    const re = new RegExp(`(^|\\n|\\r)\\s*${escapeRegex(controlId)}\\s*=\\s*\\{`);
    return re.test(segment);
  }

  function generateUniqueControlId(text, range, desiredName) {
    const base = sanitizeIdent(desiredName || 'Control') || 'Control';
    let candidate = base;
    let counter = 2;
    while (doesControlIdExist(text, range, candidate)) {
      candidate = `${base}_${counter++}`;
    }
    return candidate;
  }

  async function addControlToNode(nodeName, config) {
    if (!nodeName || !state.originalText || !state.parseResult) throw new Error('Load a macro before adding controls.');
    pushHistory('add control');
    const newline = state.newline || detectNewline(state.originalText);
    const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
    if (!bounds) throw new Error('Unable to locate the macro group block.');
    const trimmedName = (config?.name || '').trim();
    if (!trimmedName) throw new Error('Control name is required.');
    const normalizedPage = (config?.page && String(config.page).trim()) ? String(config.page).trim() : 'Controls';
    const typeRaw = (config?.type || 'label').toLowerCase();
    const supportedTypes = new Set(['label', 'separator', 'button', 'slider', 'screw']);
    const type = supportedTypes.has(typeRaw) ? typeRaw : 'label';
    let workingText = state.originalText;
    const ensured = ensureToolUserControlsBlock(workingText, bounds, nodeName, newline);
    if (!ensured || !ensured.ucBlock || !ensured.toolBlock) throw new Error('Unable to locate UserControls for the selected node.');
    workingText = ensured.text;
    const controlId = generateUniqueControlId(workingText, ensured.toolBlock, trimmedName);
    const labelDefault = config?.labelDefault === 'open' ? 'open' : 'closed';
    const lines = buildControlDefinitionLines(type, {
      name: trimmedName,
      page: normalizedPage,
      labelCount: config?.labelCount,
      labelDefault,
    });
    const pendingMeta = {
      kind: type === 'label' ? 'label' : type === 'button' ? 'button' : null,
      labelCount: type === 'label' && Number.isFinite(config?.labelCount) ? Number(config.labelCount) : null,
      defaultValue: type === 'label' ? (labelDefault === 'closed' ? '0' : '1') : null,
      inputControl:
        type === 'slider' ? 'SliderControl'
        : type === 'screw' ? 'ScrewControl'
        : type === 'button' ? 'ButtonControl'
        : type === 'separator' ? 'SeparatorControl'
        : 'LabelControl',
    };
    workingText = insertUserControlBlock(workingText, ensured.ucBlock, controlId, lines, newline);
    rememberPendingControlMeta(nodeName, controlId, pendingMeta);
    autoPublishCreatedControl(nodeName, controlId, trimmedName, normalizedPage, pendingMeta);
    pendingOpenNodes.add(nodeName);
    const persisted = rebuildContentWithNewOrder(workingText, state.parseResult, newline);
    state.originalText = persisted;
    await reloadMacroFromCurrentText({ skipClear: true });
    runValidation('add-control');
    const typeLabel = type === 'button' ? 'button' : type === 'slider' ? 'slider' : type === 'screw' ? 'screw control' : type === 'separator' ? 'separator control' : 'label';
    info(`Added ${typeLabel} "${trimmedName}" to ${nodeName}.`);
  }

  function autoPublishCreatedControl(nodeName, controlId, displayName, pageName, meta) {
    try {
      if (!state.parseResult) return;
      const metaWithPage = { ...(meta || {}), page: pageName };
      const res = ensurePublished(nodeName, controlId, displayName, metaWithPage, { skipHistory: true, skipInsert: true });
      if (!res || typeof res.index !== 'number') return;
      const currentOrder = state.parseResult.order || [];
      let pos = currentOrder.length;
      let anchored = false;
      if (typeof activeDetailEntryIndex === 'number') {
        const anchorPos = currentOrder.indexOf(activeDetailEntryIndex);
        if (anchorPos >= 0) {
          pos = anchorPos + 1;
          anchored = true;
        }
      }
      if (!anchored && typeof getInsertionPosUnderSelection === 'function') {
        pos = getInsertionPosUnderSelection();
      }
      state.parseResult.order = insertIndicesAt(currentOrder, [res.index], pos);
      try { logDiag(`autoPublishCreatedControl: idx=${res.index} pos=${pos}`); } catch (_) {}
      if (pageName && pageName !== 'Controls') {
        state.parseResult.pageOrder = state.parseResult.pageOrder || [];
        if (!state.parseResult.pageOrder.includes(pageName)) state.parseResult.pageOrder.push(pageName);
      }
      if (state.parseResult) state.parseResult.activePage = pageName;
    } catch (_) {}
  }

  function setEntryRangeValue(index, key, value) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    if (key !== 'minScale' && key !== 'maxScale') return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    const next = normalizeMetaValue(value);
    const current = entry.controlMeta[key] != null ? String(entry.controlMeta[key]) : null;
    if ((current || null) === (next || null)) return;
    pushHistory('edit control range');
    entry.controlMeta[key] = next;
    entry.controlMetaDirty = true;
    renderDetailDrawer(index);
    markContentDirty();
  }

  function normalizeInputControlValue(value) {
    if (value == null) return null;
    let str = String(value).trim();
    if (!str) return null;
    if (str.startsWith('"') && str.endsWith('"')) {
      str = str.slice(1, -1);
    }
    return str || null;
  }

  function formatInputControlValue(value) {
    const norm = normalizeInputControlValue(value);
    if (!norm) return null;
    return `"${escapeQuotes(norm)}"`;
  }

  const NUMBER_INPUT_CONTROLS = ['SliderControl','ScrewControl','CheckboxControl','ButtonControl','ComboControl','LabelControl','SeparatorControl'];
  const POINT_INPUT_CONTROLS = ['OffsetControl'];

  function getAllowedInputControls(entry) {
    try {
      const typeRaw = (entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || '').replace(/"/g, '').trim().toLowerCase();
      if (!typeRaw) return NUMBER_INPUT_CONTROLS;
      if (typeRaw.includes('number')) return NUMBER_INPUT_CONTROLS;
      if (typeRaw.includes('point')) return POINT_INPUT_CONTROLS;
      return NUMBER_INPUT_CONTROLS;
    } catch (_) {
      return NUMBER_INPUT_CONTROLS;
    }
  }

  function getKnownInputControls() {
    const set = new Set();
    try {
      const add = (val) => {
        const norm = normalizeInputControlValue(val);
        if (norm) set.add(norm);
      };
      (state.parseResult?.entries || []).forEach(entry => {
        add(entry?.controlMeta?.inputControl);
        add(entry?.controlMetaOriginal?.inputControl);
      });
    } catch (_) {}
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }

  function setEntryInputControl(index, value, options = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    const norm = normalizeInputControlValue(value);
    const current = normalizeInputControlValue(entry.controlMeta?.inputControl);
    if ((current || '') === (norm || '')) return;
    const silent = !!options.silent;
    const skipHistory = !!options.skipHistory;
    if (!skipHistory) pushHistory('change control type');
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.inputControl = norm ? formatInputControlValue(norm) : null;
    entry.controlMetaDirty = true;
    entry.isButton = norm ? /buttoncontrol/i.test(norm) : false;
    entry.controlTypeEdited = true;
    if (entry.source && String(entry.source).toLowerCase() === 'blend') {
      const set = ensureBlendToggleSet(state.parseResult);
      const key = getEntryKey(entry);
      if (norm && /checkboxcontrol/i.test(norm)) {
        entry.isBlendToggle = true;
        if (key) set.add(key);
      } else if (norm && /slidercontrol/i.test(norm)) {
        entry.isBlendToggle = false;
        if (key) set.delete(key);
      }
    }
    if (typeof entry.raw === 'string') {
      const open = entry.raw.indexOf('{');
      const close = entry.raw.lastIndexOf('}');
      if (open >= 0 && close > open) {
        let body = entry.raw.slice(open + 1, close);
        const props = [
          'INPID_InputControl',
          'INP_Integer',
          'INP_Default',
          'INP_MinScale',
          'INP_MaxScale',
          'INP_MinAllowed',
          'INP_MaxAllowed',
          'CBC_TriState',
          'LINKID_DataType',
        ];
        props.forEach(prop => {
          body = removeInstanceInputProp(body, prop);
        });
        entry.raw = entry.raw.slice(0, open + 1) + body + entry.raw.slice(close);
      }
    }
    if (!silent) {
      try { renderList(state.parseResult.entries, state.parseResult.order); } catch (_) {}
      renderDetailDrawer(index);
    }
    markContentDirty();
  }

  function handleDetailTargetChange(index) {
    if (index == null || index < 0) {
      hideDetailDrawer();
      return;
    }
    renderDetailDrawer(index);
  }

  function updateDrawerToggleLabel() {
    if (!detailDrawerToggleBtn) return;
    const isCollapsed = detailDrawerCollapsed;
    detailDrawerToggleBtn.innerHTML = '';
    const icon = document.createElement('span');
    icon.className = 'icon';
    const html = renderIconHtml ? renderIconHtml(isCollapsed ? 'chevron-left' : 'chevron-right', 14) : '';
    icon.innerHTML = html || '';
    detailDrawerToggleBtn.appendChild(icon);
  }

  detailDrawerToggleBtn?.addEventListener('click', () => {
    if (!detailDrawer || detailDrawer.hidden) return;
    detailDrawerCollapsed = !detailDrawerCollapsed;
    detailDrawer.classList.toggle('collapsed', detailDrawerCollapsed);
    updateDrawerToggleLabel();
  });

  hideDetailDrawer();
  updateDrawerToggleLabel();

  function updateCodePaneToggle() {
    if (codePane) {
      codePane.classList.toggle('collapsed', !!codePaneCollapsed);
    }
    if (codePaneToggleInput) {
      codePaneToggleInput.checked = !codePaneCollapsed;
    }
  }

  codePaneToggleInput?.addEventListener('change', () => {
    if (!codePaneToggleInput) return;
    codePaneCollapsed = !codePaneToggleInput.checked;
    updateCodePaneToggle();
  });

  updateCodePaneToggle();

  const publishedControls = createPublishedControls({
    controlsList,
    removeBtn,
    deselectAllBtn,
    ctrlCountEl,
    state,
    logDiag,
    info,
    highlightNode,
    clearHighlights,
    pushHistory,
    getSelectedSet,
    refreshNodesChecks: () => {
      try { nodesPane?.refreshNodesChecks?.(); } catch (_) {}
    },
    shouldOfferInsertForControl,
    findEnclosingGroupForIndex,
    findToolBlockInGroup,
    findUserControlsInTool,
    findUserControlsInGroup,
    findControlBlockInUc,
    unescapeSettingString,
    makeUniqueKey,
    getOrAssignControlGroup,
    buildInstanceInputRaw,
    pageTabsEl,
    onDetailTargetChange: (index) => {
      try { handleDetailTargetChange(typeof index === 'number' ? index : null); } catch (_) {}
    },
    applyNodeControlMeta: (entry, meta) => applyNodeControlMeta(entry, meta),
    onRenderList: () => markContentDirty(),
    onEntryMutated: () => markContentDirty(),
  });

  const {
    renderList: renderListInternal,
    setFilter: setPublishedFilter,
    updateRemoveSelectedState,
    cleanupDropIndicator,
    updateDropIndicatorPosition,
    getInsertionPosUnderSelection,
    addOrMovePublishedItemsAt,
    createIcon,
    getCurrentDropIndex,
    refreshPageTabs,
    setEntryDisplayName: setEntryDisplayNameApi,
    appendLauncherUi: appendLauncherUiApi,
  } = publishedControls;
  const renderList = (entries, order) => {
    renderListInternal(entries, order);
  };
  renderIconHtml = typeof createIcon === 'function' ? createIcon : (() => '');
  updateDrawerToggleLabel();

  nodesPane = createNodesPane({
    state,
    nodesList,
    nodesSearch,
    hideReplacedEl,
    publishSelectedBtn,
    clearNodeSelectionBtn,
    importCatalogBtn,
    catalogInput,
    importModifierCatalogBtn,
    modifierCatalogInput,
    logDiag,
    logTag,
    error,
    highlightNode,
    clearHighlights,
    renderList: (entries, order) => renderList(entries, order),
    getInsertionPosUnderSelection,
    insertIndicesAt,
    isPublished,
    ensurePublished,
    ensureEntryExists,
    removePublished,
    createIcon,
    sanitizeIdent,
    normalizeId,
    enableNodeDrag: ENABLE_NODE_DRAG,
    requestAddControl: (nodeName) => openAddControlDialog(nodeName),
    getPendingControlMeta: (op, id) => getPendingControlMeta(op, id),
    consumePendingControlMeta: (op, id) => consumePendingControlMeta(op, id),
  });

  historyController = createHistoryController({
    state,
    renderList: (entries, order) => renderList(entries, order),
    ctrlCountEl,
    getNodesPane: () => nodesPane,
    undoBtn,
    redoBtn,
    logDiag,
    doc: (typeof document !== 'undefined') ? document : null,
    captureExtraState: () => captureHistoryExtras(),
    restoreExtraState: (extra, context) => restoreHistoryExtras(extra, context),
  });

  const nativeBridge = setupNativeBridge({
    isElectron: IS_ELECTRON,
    exportBtn,
    importClipboardBtn,
    exportClipboardBtn,
    setDiagnosticsEnabled: () => {
      try { setDiagnosticsEnabled(!diagnosticsController.isEnabled()); } catch (_) {}
    },
    handleNativeOpen,
  });

  function setExportButtonLabel(label) {
    if (!exportBtn) return;
    exportBtn.textContent = label || DEFAULT_EXPORT_LABEL;
  }

  async function updateExportButtonLabelFromPath(filePath) {
    if (!filePath || !nativeBridge || !nativeBridge.ipcRenderer) {
      if (exportBtn) setExportButtonLabel(DEFAULT_EXPORT_LABEL);
      state.drfxLink = null;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      return;
    }
    try {
      const res = await nativeBridge.ipcRenderer.invoke('drfx-get-link', { path: filePath });
      const link = (res && res.ok) ? {
        linked: !!res.linked,
        drfxPath: res.drfxPath || '',
        presetName: res.presetName || '',
      } : { linked: false };
      state.drfxLink = link;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      if (exportBtn) {
        setExportButtonLabel(link.linked ? DRFX_EXPORT_LABEL : DEFAULT_EXPORT_LABEL);
      }
    } catch (_) {
      if (exportBtn) setExportButtonLabel(DEFAULT_EXPORT_LABEL);
      state.drfxLink = null;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
    }
  }
  if (nativeBridge && nativeBridge.ipcRenderer) {
    nativeBridge.ipcRenderer.on('fmr-open-path', async (_event, payload = {}) => {
      const filePath = payload.path;
      if (!filePath) return;
      try {
        await handleDroppedPath(filePath);
      } catch (_) {
        /* ignore open-path errors */
      }
    });
    nativeBridge.ipcRenderer.on('fmr-set-export-folder', (_event, payload = {}) => {
      const folderPath = payload.path;
      if (!folderPath) return;
      setExportFolder(folderPath, { selected: true });
    });
  }

  openExplorerBtn?.addEventListener('click', () => {
    if (!IS_ELECTRON) {
      info('Macro Explorer is available in the desktop app.');
      return;
    }
    if (macroExplorerPanel) {
      macroExplorerPanel.hidden = false;
    }
    if (macroExplorerMini) {
      macroExplorerMini.hidden = true;
    }
  });
  closeExplorerPanelBtn?.addEventListener('click', () => {
    if (macroExplorerPanel) {
      macroExplorerPanel.hidden = true;
    }
    if (macroExplorerMini) {
      macroExplorerMini.hidden = false;
    }
  });

  const dropHintDefaultText = dropHint ? dropHint.textContent : '';
  function updateDropHint(text) {
    try {
      if (!dropHint) return;
      dropHint.textContent = text || dropHintDefaultText;
    } catch (_) {}
  }

  // Make macro name editable in-place
  try {
    if (macroNameEl) {
      macroNameEl.contentEditable = 'true';
      macroNameEl.spellcheck = false;
      macroNameEl.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter') { ev.preventDefault(); macroNameEl.blur(); }
      });
      macroNameEl.addEventListener('blur', () => {
        try {
          const newName = (macroNameEl.textContent || '').trim();
          if (!state.parseResult) return;
          const prevName = state.parseResult.macroName || '';
          if (newName) {
            state.parseResult.macroName = newName;
            updateActiveDocumentMeta({ name: newName });
            if (newName !== prevName) markActiveDocumentDirty();
          } else {
            // Revert to original if cleared
            const fallback = state.parseResult.macroNameOriginal || state.parseResult.macroName || 'Unknown';
            macroNameEl.textContent = fallback;
            updateActiveDocumentMeta({ name: fallback });
            if (fallback !== prevName) markActiveDocumentDirty();
          }
        } catch (_) {}
      });
    }
  } catch (_) {}

  function syncOperatorSelect() {
    try {
      if (!operatorTypeSelect) return;
      const val = (state.parseResult && state.parseResult.operatorType) ? state.parseResult.operatorType : 'GroupOperator';
      operatorTypeSelect.value = (val === 'MacroOperator') ? 'MacroOperator' : 'GroupOperator';
    } catch (_) {}
  }
  operatorTypeSelect?.addEventListener('change', () => {
    if (!state.parseResult) return;
    const val = operatorTypeSelect.value === 'MacroOperator' ? 'MacroOperator' : 'GroupOperator';
    const prev = state.parseResult.operatorType || '';
    state.parseResult.operatorType = val;
    if (val !== prev) markActiveDocumentDirty();
  });

  // If running under Electron and the native open API is available,
  // inject an "Open .setting…" button into the file loader section and
  // hide the older browser file picker to avoid redundant controls.
  try {
    const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
    const canNativeOpen = !!(native && native.isElectron && typeof native.openSettingFile === 'function');
    if (canNativeOpen && !openNativeBtn && dropZone) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.id = 'openNativeBtn';
      btn.textContent = 'Open .setting…';
      // Insert before the clipboard button if present, otherwise at end
      const ref = importClipboardBtn || fileInfo || null;
      if (ref && ref.parentNode === dropZone) {
        dropZone.insertBefore(btn, ref);
      } else {
        dropZone.appendChild(btn);
      }
      openNativeBtn = btn;
      // Hide legacy browser file picker controls (label + input)
      try {
        const legacyLabel = document.querySelector('label[for="fileInput"]');
        if (legacyLabel) legacyLabel.style.display = 'none';
        if (fileInput) fileInput.style.display = 'none';
      } catch (_) {}
    }
  } catch(_) {}

  // Nodes state handled in nodesPane.js
 
  // Published search/filter

  publishedSearch?.addEventListener('input', (e) => {
    setPublishedFilter(e.target.value || '');
    if (state.parseResult) renderList(state.parseResult.entries, state.parseResult.order);
  });



  // File picker

  fileInput?.addEventListener('change', async (e) => {

    const files = Array.from(e.target.files || []);

    if (files.length) {

      logDiag(`fileInput change - files: ${files.length}`);

      await handleFiles(files);

    }

  });

  // Native "Open .setting…" (Electron)
  if (openNativeBtn) {
    openNativeBtn.addEventListener('click', async () => {
      try {
        const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
        if (!native || !native.isElectron || typeof native.openSettingFile !== 'function') {
          // Fallback to browser file picker if native open is not available
          try { fileInput && fileInput.click(); } catch(_) {}
          return;
        }
        const res = await native.openSettingFile();
        if (!res || res.canceled) return;
        if (res.error) {
          error('Open failed: ' + res.error);
          return;
        }
        const name = res.baseName || res.filePath || 'Imported.setting';
        const text = res.content || '';
        if (!text) {
          error('Selected file was empty or could not be read.');
          return;
        }
        await loadSettingText(name, text);
        setExportFolderFromFilePath(res.filePath, { silent: true });
        info('Loaded macro from file: ' + name);
      } catch (err) {
        error('Native open failed: ' + (err?.message || err));
      }
    });
  }



  // Shared handler for native open (Electron)
  async function handleNativeOpen() {
    try {
      const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
      if (!native || !native.isElectron || typeof native.openSettingFile !== 'function') {
        try { fileInput && fileInput.click(); } catch(_) {}
        return;
      }
      const res = await native.openSettingFile();
      if (!res || res.canceled) return;
      if (res.error) {
        error('Open failed: ' + res.error);
        return;
      }
      const name = res.baseName || res.filePath || 'Imported.setting';
      const text = res.content || '';
      if (!text) {
        error('Selected file was empty or could not be read.');
        return;
      }
      await loadSettingText(name, text);
      setExportFolderFromFilePath(res.filePath, { silent: true });
      info('Loaded macro from file: ' + name);
    } catch (err) {
      error('Native open failed: ' + (err?.message || err));
    }
  }

  // Drag & drop file loading

  if (dropZone) {

    ['dragenter','dragover'].forEach(type => {

      dropZone.addEventListener(type, (e) => {

        e.preventDefault();

        e.stopPropagation();

        dropZone.classList.add('dragover');
        updateDropHint('Drop .setting file here… (over panel)');
        try { logDiag(`[Drop] ${type} on dropZone`); } catch(_) {}

      });

    });

    ['dragleave','drop'].forEach(type => {

      dropZone.addEventListener(type, (e) => {

        e.preventDefault();

        e.stopPropagation();

        dropZone.classList.remove('dragover');
        updateDropHint('Drop .setting file here…');
        try { logDiag(`[Drop] ${type} on dropZone`); } catch(_) {}

      });

    });

    dropZone.addEventListener('drop', async (e) => {

      const files = e.dataTransfer?.files;

      if (files && files.length) {

        try { logDiag('[Drop] dropZone drop (files)'); } catch(_) {}

        await handleFiles(Array.from(files));

        return;
      }

      // Electron: sometimes only a file path is provided via text/URI list
      try {
        const dt = e.dataTransfer;
        if (dt && IS_ELECTRON) {
          const uriList = dt.getData('text/uri-list') || dt.getData('text/plain');
          const filePath = extractFilePathFromData(uriList);
          if (filePath) {
            try { logDiag('[Drop] dropZone drop (path) ' + filePath); } catch(_) {}
            if (isDrfxPath(filePath)) {
              await handleDrfxFiles([filePath]);
            } else {
              await handleDroppedPath(filePath);
            }
          }
        }
      } catch (_) {}

    });

  }



  // Make drag & drop resilient across the whole page

  let dragDepth = 0;

  function hasFiles(e) {

    const dt = e.dataTransfer; if (!dt) return false;

    return Array.from(dt.types || []).includes('Files');

  }

  window.addEventListener('dragenter', (e) => {

    if (!hasFiles(e)) return;

    dragDepth++;

    dropZone && dropZone.classList.add('dragover');
    updateDropHint('Drop .setting file here… (window dragenter)');
    try { logDiag('[Drop] window dragenter'); } catch(_) {}

  });

  window.addEventListener('dragover', (e) => {
    if (hasFiles(e)) {
      e.preventDefault();
      updateDropHint('Drop .setting file here… (window dragover)');
      try { logDiag('[Drop] window dragover'); } catch(_) {}
    }
  });

  window.addEventListener('dragleave', (e) => {

    if (!hasFiles(e)) return;

    dragDepth = Math.max(0, dragDepth - 1);

    if (dragDepth === 0 && dropZone) {
      dropZone.classList.remove('dragover');
      updateDropHint('Drop .setting file here…');
      try { logDiag('[Drop] window dragleave'); } catch(_) {}
    }

  });

  window.addEventListener('drop', (e) => {

    if (hasFiles(e)) {
      e.preventDefault();
      try { logDiag('[Drop] window drop'); } catch(_) {}
    }

    dragDepth = 0;

    dropZone && dropZone.classList.remove('dragover');
    updateDropHint('Drop .setting file here…');

  });

  document.addEventListener('drop', async (e) => {

    const dt = e.dataTransfer;

    if (!dt) return;

    const files = dt.files;
    if (files && files.length > 0) {
      e.preventDefault();
      e.stopPropagation();
      try { logDiag('[Drop] document drop (files)'); } catch(_) {}
      await handleFiles(Array.from(files));
      dropZone && dropZone.classList.remove('dragover');
      updateDropHint('Drop .setting file here…');
      return;
    }

    // Electron path-based drop
    try {
      if (IS_ELECTRON) {
        const uriList = dt.getData('text/uri-list') || dt.getData('text/plain');
        const filePath = extractFilePathFromData(uriList);
        if (filePath) {
          e.preventDefault();
          e.stopPropagation();
          try { logDiag('[Drop] document drop (path) ' + filePath); } catch(_) {}
          if (isDrfxPath(filePath)) {
            await handleDrfxFiles([filePath]);
          } else {
            await handleDroppedPath(filePath);
          }
          dropZone && dropZone.classList.remove('dragover');
          updateDropHint('Drop .setting file here…');
        }
      }
    } catch (_) {}

  });

  async function handleDroppedPath(filePath) {
    try {
      if (!IS_ELECTRON) return;
      // Lazy-require to avoid errors in browser
      // eslint-disable-next-line global-require
      const fs = require('fs');
      // eslint-disable-next-line global-require
      const path = require('path');
      const content = fs.readFileSync(filePath, 'utf8');
      const name = path.basename(filePath);
      await loadSettingText(name, content);
      setExportFolderFromFilePath(filePath, { silent: true });
      updateExportButtonLabelFromPath(filePath);
      info('Loaded macro from dropped file: ' + name);
    } catch (err) {
      error('Failed to load dropped file: ' + (err?.message || err));
    }
  }

  function extractFilePathFromData(data) {
    try {
      if (!data) return null;
      const firstLine = String(data).split(/\r?\n/)[0].trim();
      if (!firstLine) return null;
      // Handle file:///C:/... URIs
      if (firstLine.startsWith('file:///')) {
        const decoded = decodeURI(firstLine.replace('file:///', ''));
        return decoded.replace(/\//g, '\\');
      }
      // Direct Windows-style path
      if (/^[a-zA-Z]:\\/.test(firstLine)) return firstLine;
      return null;
    } catch (_) { return null; }
  }

  function setExportFolder(folderPath, options = {}) {
    const { silent = false, selected = false } = options || {};
    if (!folderPath) return;
    state.exportFolder = folderPath;
    if (selected) state.exportFolderSelected = true;
    if (!silent) info(`Export folder set to ${folderPath}`);
    const doc = getActiveDocument();
    if (doc) storeDocumentSnapshot(doc);
    updateDocExportPathDisplay();
  }

  function setExportFolderFromFilePath(filePath, options = {}) {
    if (!filePath || !IS_ELECTRON) return;
    state.originalFilePath = filePath;
    updateExportButtonLabelFromPath(filePath);
    const doc = getActiveDocument();
    if (doc) storeDocumentSnapshot(doc);
    updateDocExportPathDisplay();
    if (state.exportFolderSelected) return;
    try {
      // eslint-disable-next-line global-require
      const path = require('path');
      const dir = path.dirname(filePath);
      if (dir) setExportFolder(dir, options);
    } catch (_) {}
  }

  function getFusionBasePath() {
    if (!IS_ELECTRON) return '';
    try {
      // eslint-disable-next-line global-require
      const os = require('os');
      // eslint-disable-next-line global-require
      const path = require('path');
      // eslint-disable-next-line global-require
      const fs = require('fs');
      if (process.platform === 'win32') {
        const appdata = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
        const support = path.join(appdata, 'Blackmagic Design', 'DaVinci Resolve', 'Support', 'Fusion');
        const legacy = path.join(appdata, 'Blackmagic Design', 'DaVinci Resolve', 'Fusion');
        return fs.existsSync(support) ? support : legacy;
      }
      if (process.platform === 'darwin') {
        return path.join(
          os.homedir(),
          'Library',
          'Application Support',
          'Blackmagic Design',
          'DaVinci Resolve',
          'Fusion'
        );
      }
      return path.join(
        os.homedir(),
        '.local',
        'share',
        'Blackmagic Design',
        'DaVinci Resolve',
        'Fusion'
      );
    } catch (_) {
      return '';
    }
  }

  function getFusionTemplatesPath() {
    const base = getFusionBasePath();
    if (!base) return '';
    try {
      // eslint-disable-next-line global-require
      const path = require('path');
      return path.join(base, 'Templates');
    } catch (_) {
      return '';
    }
  }

  function resolveExportFolder() {
    return state.exportFolder || getFusionTemplatesPath();
  }

  function buildExportDefaultPath(fileName) {
    const folder = resolveExportFolder();
    if (!folder) return fileName;
    try {
      // eslint-disable-next-line global-require
      const path = require('path');
      return path.join(folder, fileName);
    } catch (_) {
      return fileName;
    }
  }

  function writeSettingToFolder(folderPath, fileName, content) {
    if (!folderPath) throw new Error('No export folder set.');
    // eslint-disable-next-line global-require
    const fs = require('fs');
    // eslint-disable-next-line global-require
    const path = require('path');
    fs.mkdirSync(folderPath, { recursive: true });
    const destPath = path.join(folderPath, fileName);
    fs.writeFileSync(destPath, content, 'utf8');
    return destPath;
  }



  async function loadMacroFromText(sourceName, rawText, options = {}) {

    const {
      allowAutoUtility = true,
      skipClear = false,
      preserveFileInfo = false,
      silentAuto = false,
      preserveFilePath = false,
      createDoc = true,
    } = options || {};

    if (createDoc) {
      const currentDoc = getActiveDocument();
      if (currentDoc) storeDocumentSnapshot(currentDoc);
    }

    const prevSuppress = suppressDocDirty;
    suppressDocDirty = true;

    if (!preserveFilePath) {
      state.originalFilePath = '';
      state.drfxLink = null;
    }

    if (!skipClear) clearUI();

    state.originalFileName = sourceName || 'Imported.setting';

    if (!preserveFileInfo && fileInfo) {
      fileInfo.textContent = `${state.originalFileName} (${(rawText || '').length.toLocaleString()} chars)`;
    }

    try {
      let text = rawText || '';
      const wrapRes = maybeWrapUngroupedNodes(text, { sourceName });
      text = wrapRes.text;
      const wrapped = wrapRes.changed;
      text = stripDanglingRootUserControls(text);
      text = stripDanglingRootInputs(text);
      if (allowAutoUtility) {
        text = await maybeAutoAddUtility(text, { silent: silentAuto });
      }
      cancelPendingCodeRefresh();
      state.originalText = text;
      state.lastDiffRange = null;
      updateCodeView(text);
      state.newline = detectNewline(text);
      if (diagnostics) { diagnostics.hidden = false; diagnostics.innerHTML = ''; }
      try {
        const prim = findAllInputsBlocksWithInstanceInputs(text) || [];
        const any = findInputsBlockAnywhere(text);
        const back = findInputsBlockByBacktrack ? findInputsBlockByBacktrack(text) : null;
        logDiag(`Inputs blocks - primary: ${prim.length}, any: ${any ? 1 : 0}, backtrack: ${back ? 1 : 0}`);
      } catch (e) {
        logDiag(`Diagnostics error while scanning Inputs: ${e.message || e}`);
      }
      state.parseResult = parseSetting(text);
      state.parseResult.pageOrder = derivePageOrderFromEntries(state.parseResult.entries);
      state.parseResult.activePage = null;
      hydrateEntryPagesFromUserControls(state.originalText, state.parseResult);
      hydrateControlPageIcons(state.originalText, state.parseResult);
      runValidation('load');
      hydrateBlendToggleState(state.originalText, state.parseResult);
      hydrateOnChangeScripts(state.originalText, state.parseResult);
      hydrateLabelVisibility(state.originalText, state.parseResult);
      hydrateControlMetadata(state.originalText, state.parseResult);
      if (!state.parseResult.operatorType) state.parseResult.operatorType = state.parseResult.operatorTypeOriginal || 'GroupOperator';
      syncOperatorSelect();
      publishedControls.resetPageOptions?.();
      state.parseResult.buttonExactInsert = new Set();
      state.parseResult.insertUrlMap = new Map();
      state.parseResult.insertClickedKeys = new Set();
      state.parseResult.nodesCollapsed = new Set();
      state.parseResult.nodesPublishedOnly = new Set();
      state.parseResult.nodesViewInitialized = false;
      state.parseResult.nodeSelection = new Set();
      const existingCG = (state.parseResult.entries || []).map(e => e && e.controlGroup).filter(v => Number.isFinite(v));
      state.parseResult.cgNext = existingCG.length ? Math.max(...existingCG) + 1 : 1;
      state.parseResult.cgMap = new Map();
      state.parseResult.collapsed = new Set();
      state.parseResult.collapsedCG = new Set();
      state.parseResult.selected = new Set();
      if (!state.parseResult.entries.length) {
        info('No published controls found. Use the Nodes pane to add controls.');
      }
      macroNameEl.textContent = state.parseResult.macroName || 'Unknown';
      ctrlCountEl.textContent = String(state.parseResult.entries.length);
      registerLoadedDocument(state.originalFileName || sourceName, { createDoc });
      logDiag(`Parsed ok - macro: ${state.parseResult.macroName || 'Unknown'}, controls: ${state.parseResult.entries.length}`);
      controlsSection.hidden = false;
      resetBtn.disabled = false;
      exportBtn.disabled = false;
      exportTemplatesBtn && (exportTemplatesBtn.disabled = false);
      exportClipboardBtn && (exportClipboardBtn.disabled = false);
      updateExportMenuButtonState();
      removeCommonPageBtn && (removeCommonPageBtn.disabled = false);
      publishSelectedBtn && (publishSelectedBtn.disabled = true);
      clearNodeSelectionBtn && (clearNodeSelectionBtn.disabled = true);
      state.parseResult.history = [];
      state.parseResult.future = [];
      updateUndoRedoState();
      renderList(state.parseResult.entries, state.parseResult.order);
      refreshPageTabs?.();
      updateRemoveSelectedState();
      if (wrapped) info('Detected loose nodes and wrapped them into a macro automatically.');
      info(`Parsed ${state.parseResult.entries.length} controls.`);
      nodesPane && nodesPane.parseAndRenderNodes();
      if (pendingOpenNodes.size && nodesPane) {
        pendingOpenNodes.forEach((name) => {
          try { nodesPane.expandNode?.(name, 'open'); } catch (_) {}
        });
        pendingOpenNodes.clear();
        nodesPane.parseAndRenderNodes();
      }
      updateUtilityActionsState();
      if (IS_ELECTRON) setIntroCollapsed(true);
    } catch (err) {
      const msg = err.message || String(err);
      error(msg);
      logDiag(`Parse error: ${msg}`);
      controlsSection.hidden = true;
      updateUtilityActionsState();
    } finally {
      suppressDocDirty = prevSuppress;
    }
  }

async function handleFile(file) {

    if (!file) return;

    if (fileInfo) fileInfo.textContent = `${file.name} (${file.size.toLocaleString()} bytes)`;

    try {

      const text = await file.text();

      await loadMacroFromText(file.name, text, { preserveFileInfo: true, createDoc: true });
      setExportFolderFromFilePath(file.path, { silent: true });

    } catch (err) {

      error(err?.message || err || 'Failed to read file.');

    }

  }

  function isDrfxPath(value) {
    const lower = String(value || '').toLowerCase();
    return lower.endsWith('.drfx') || lower.endsWith('.drfx.disabled');
  }

  async function handleDrfxFiles(files) {
    if (!files || !files.length) return;
    if (!IS_ELECTRON || !nativeBridge || !nativeBridge.ipcRenderer) {
      error('DRFX import is available in the desktop app.');
      return;
    }
    const paths = files
      .map(file => (typeof file === 'string' ? file : file.path))
      .filter(filePath => filePath && isDrfxPath(filePath));
    if (!paths.length) return;
    // eslint-disable-next-line global-require
    const path = require('path');
    let opened = 0;
    let failed = 0;
    for (const drfxPath of paths) {
      try {
        const res = await nativeBridge.ipcRenderer.invoke('drfx-list', { path: drfxPath });
        if (!res || !res.ok) {
          failed += 1;
          error(res?.error || 'Failed to read DRFX.');
          continue;
        }
        const presetNames = (res.presets || [])
          .map(preset => preset.entryName || preset.name)
          .filter(Boolean);
        if (!presetNames.length) {
          info(`No presets found in ${path.basename(drfxPath)}.`);
          continue;
        }
        if (presetNames.length > 5) {
          const ok = window.confirm(`Import ${presetNames.length} presets from ${path.basename(drfxPath)}? This will open a tab for each.`);
          if (!ok) continue;
        }
        for (const presetName of presetNames) {
          const extractRes = await nativeBridge.ipcRenderer.invoke('drfx-extract-preset', {
            path: drfxPath,
            presetName,
          });
          if (!extractRes || !extractRes.ok || !extractRes.filePath) {
            failed += 1;
            continue;
          }
          // eslint-disable-next-line no-await-in-loop
          await handleDroppedPath(extractRes.filePath);
          opened += 1;
        }
      } catch (err) {
        failed += 1;
        error(err?.message || err || 'Failed to import DRFX.');
      }
    }
    if (opened && !failed) {
      info(`Imported ${opened} preset(s) from DRFX.`);
    } else if (opened || failed) {
      info(`Imported ${opened} preset(s), ${failed} failed.`);
    }
  }

  async function handleFiles(files) {

    if (!files || !files.length) return;

    for (const file of files) {
      const filePath = (typeof file === 'string') ? file : (file?.path || file?.name || '');
      if (isDrfxPath(filePath)) {
        // eslint-disable-next-line no-await-in-loop
        await handleDrfxFiles([file]);
        continue;
      }

      // Load sequentially to preserve tab order.
      // eslint-disable-next-line no-await-in-loop
      await handleFile(file);

    }

  }

  // Load a .setting from text (clipboard/paste)
  async function loadSettingText(sourceName, text) {
    setExportButtonLabel(DEFAULT_EXPORT_LABEL);
    await loadMacroFromText(sourceName || 'Clipboard.setting', text || '', { createDoc: true });
  }
  // Auto-load catalogs when running under Electron via fetch from local files
  (async () => {
    try {
      const isElectron = (typeof window !== 'undefined') && !!(window.FusionMacroReordererNative && window.FusionMacroReordererNative.isElectron);
      if (!isElectron) return;

      const pending = [];

      if (!nodesPane?.getNodeCatalog?.()) {
        pending.push((async () => {
          try {
            const resp = await fetch('FusionNodeCatalog.cleaned.json');
            if (!resp.ok) return;
            const json = await resp.json();
            nodesPane?.setNodeCatalog?.(json || null);
            logTag('Catalog', 'Loaded node catalog via fetch: ' + Object.keys(json || {}).length + ' tool types');
          } catch (err) {
            logTag('Catalog', 'Node catalog fetch failed: ' + (err?.message || err));
          }
        })());
      }

      if (!nodesPane?.getModifierCatalog?.()) {
        pending.push((async () => {
          try {
            const resp = await fetch('FusionModifierCatalog.json');
            if (!resp.ok) return;
            const json = await resp.json();
            nodesPane?.setModifierCatalog?.(json || null);
            logTag('Catalog', 'Loaded modifier catalog via fetch: ' + Object.keys(json || {}).length + ' modifier types');
          } catch (err) {
            logTag('Catalog', 'Modifier catalog fetch failed: ' + (err?.message || err));
          }
        })());
      }

      if (pending.length) {
        await Promise.all(pending);
        // If a macro is already loaded, refresh the Nodes pane with the new catalogs
        if (state.parseResult && nodesPane) nodesPane.parseAndRenderNodes();
      }
    } catch (err) {
      logTag('Catalog', 'Auto-load failed: ' + (err?.message || err));
    }
  })();

  resetBtn?.addEventListener('click', () => {

    if (!state.parseResult) return;

    pushHistory('reset order');

    state.parseResult.order = [...state.parseResult.originalOrder];

    renderList(state.parseResult.entries, state.parseResult.order);

    info('Order reset.');

  });

  addUtilityNodeBtn?.addEventListener('click', async () => {
    if (!state.originalText) {
      error('Load a macro before adding the Utility node.');
      return;
    }
    if (macroHasUtilityNode(state.originalText)) {
      info('Utility node already present.');
      return;
    }
    const tpl = await ensureUtilityTemplate();
    if (!tpl || !tpl.trim()) {
      error('Utility template unavailable.');
      return;
    }
    const updated = injectUtilityNodeText(state.originalText, tpl, state.newline || detectNewline(state.originalText));
    if (updated === state.originalText) {
      info('Utility node already present.');
      return;
    }
    await loadMacroFromText(state.originalFileName || 'Imported.setting', updated, {
      allowAutoUtility: false,
      skipClear: true,
      preserveFileInfo: true,
      silentAuto: true,
      createDoc: false,
      preserveFilePath: true,
    });
    info('Utility node added.');
  });


  // Remove selected published controls

  removeBtn?.addEventListener('click', () => {

    if (!state.parseResult) return;

    pushHistory('remove selected');

    const sel = getSelectedSet();

    if (!sel || sel.size === 0) return;

    removePublishedByIndices(Array.from(sel));

  });



  async function exportToSourceDrfx() {
    if (!state.parseResult) return;
    const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
    if (!native || typeof native.saveDrfxPreset !== 'function' || !state.originalFilePath) {
      error('Active macro is not linked to a DRFX preset.');
      return;
    }
    try {
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline);
      const drfxRes = await native.saveDrfxPreset({
        sourcePath: state.originalFilePath,
        content: newContent,
      });
      if (drfxRes && drfxRes.ok) {
        info('Saved back to DRFX.');
        if (state.originalFilePath) state.lastExportPath = state.originalFilePath;
        const doc = getActiveDocument();
        if (doc) storeDocumentSnapshot(doc);
        updateDocExportPathDisplay();
        markActiveDocumentClean();
        return;
      }
      if (drfxRes && drfxRes.code === 'not-linked') {
        error('Active macro is not linked to a DRFX preset.');
        return;
      }
      error('Save back to DRFX failed: ' + (drfxRes?.error || 'Unknown error.'));
    } catch (errDrfx) {
      error('Save back to DRFX failed: ' + (errDrfx?.message || errDrfx));
    }
  }

  async function exportToSourceDrfxMulti() {
    if (!state.parseResult) return;
    const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
    if (!native || typeof native.saveDrfxPreset !== 'function') {
      error('Export to Source DRFX is available in the desktop app.');
      return;
    }
    const active = getActiveDocument();
    if (active) storeDocumentSnapshot(active);
    const selected = documents.filter(doc => doc && doc.selected);
    const candidates = selected.length ? selected : documents;
    const linkedDocs = getDrfxDocumentsFrom(candidates);
    if (linkedDocs.length < 2) {
      error('Select at least two tabs linked to a DRFX preset.');
      return;
    }
    const drfxPaths = new Set();
    linkedDocs.forEach((doc) => {
      const link = doc.snapshot?.drfxLink;
      if (link && link.drfxPath) drfxPaths.add(link.drfxPath);
    });
    const confirmMsg = drfxPaths.size > 1
      ? `Export ${linkedDocs.length} preset(s) back to ${drfxPaths.size} DRFX packs? This will overwrite the source packs.`
      : `Export ${linkedDocs.length} preset(s) back to the source DRFX? This will overwrite the source pack.`;
    const ok = window.confirm(confirmMsg);
    if (!ok) return;
    let saved = 0;
    let failed = 0;
    let activeSaved = false;
    for (const doc of linkedDocs) {
      const snap = doc.snapshot;
      if (!snap || !snap.parseResult || !snap.originalFilePath) {
        failed += 1;
        continue;
      }
      try {
        const content = rebuildContentWithNewOrder(snap.originalText || '', snap.parseResult, snap.newline);
        // eslint-disable-next-line no-await-in-loop
        const res = await native.saveDrfxPreset({
          sourcePath: snap.originalFilePath,
          content,
        });
        if (res && res.ok) {
          saved += 1;
          if (doc.snapshot) {
            doc.snapshot.lastExportPath = snap.originalFilePath;
          }
          doc.isDirty = false;
          if (doc.id === activeDocId) {
            state.lastExportPath = snap.originalFilePath;
            updateDocExportPathDisplay();
            activeSaved = true;
          }
        } else {
          failed += 1;
        }
      } catch (_) {
        failed += 1;
      }
    }
    if (saved) {
      renderDocTabs();
      if (activeSaved) markActiveDocumentClean();
    }
    if (saved && !failed) {
      info(`Saved ${saved} preset(s) back to DRFX.`);
      return;
    }
    if (saved || failed) {
      info(`Saved ${saved} preset(s), ${failed} failed.`);
    }
  }

  async function exportToFile() {
    if (!state.parseResult) return;
    try {
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline);
      const outName = suggestOutputName(state.originalFileName);
      const defaultPath = buildExportDefaultPath(outName);
      const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
      if (native && typeof native.saveSettingFile === 'function') {
        try {
          const res = await native.saveSettingFile({ defaultPath, content: newContent });
          if (res && res.filePath) {
            state.lastExportPath = res.filePath;
            const doc = getActiveDocument();
            if (doc) storeDocumentSnapshot(doc);
            updateDocExportPathDisplay();
          }
          info('Saved reordered .setting via native dialog');
          markActiveDocumentClean();
        } catch (err2) {
          error('Native save failed: ' + (err2?.message || err2));
        }
      } else {
        triggerDownload(outName, newContent);
        info('Exported reordered .setting');
        markActiveDocumentClean();
      }
    } catch (err) {
      error(err.message || String(err));
    }
  }

  async function exportToEditPage() {
    if (!state.parseResult) return;
    try {
      if (!state.exportFolderSelected) {
        const msg = 'Choose a destination folder in Macro Explorer before exporting to the Edit Page.';
        try { window.alert(msg); } catch (_) {}
        error(msg);
        return;
      }
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline);
      const outName = suggestOutputName(state.originalFileName);
      if (!IS_ELECTRON) {
        triggerDownload(outName, newContent);
        info('Exported reordered .setting');
        markActiveDocumentClean();
        return;
      }
      const targetFolder = resolveExportFolder();
      const destPath = writeSettingToFolder(targetFolder, outName, newContent);
      info(`Exported to ${destPath}`);
      state.lastExportPath = destPath;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      updateDocExportPathDisplay();
      markActiveDocumentClean();
    } catch (err) {
      error(err.message || String(err));
    }
  }

  async function exportToClipboard() {
    if (!state.parseResult) return;
    try {
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline);
      const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;
      if (native && typeof native.writeClipboard === 'function') {
        try {
          native.writeClipboard(newContent);
          info('Copied reordered .setting to clipboard (native).');
        } catch (err2) {
          error('Native clipboard copy failed: ' + (err2?.message || err2));
        }
      } else {
        await writeToClipboard(newContent);
        info('Copied reordered .setting to clipboard.');
      }
    } catch (err) {
      error('Clipboard copy failed: ' + (err.message || err));
    }
  }

  function normalizePresetName(value) {
    const name = String(value || '').replace(/\.setting$/i, '').trim();
    return name || 'Preset';
  }

  function buildUniquePresetNames(entries) {
    const used = new Set();
    const results = [];
    entries.forEach((item) => {
      const base = normalizePresetName(item.name);
      let name = base;
      let suffix = 2;
      while (used.has(name.toLowerCase())) {
        name = `${base} ${suffix}`;
        suffix += 1;
      }
      used.add(name.toLowerCase());
      results.push({ ...item, name });
    });
    return results;
  }

  function getLooseDocumentsFrom(list) {
    return list.filter((doc) => {
      if (!doc || !doc.snapshot || !doc.snapshot.parseResult) return false;
      const linked = doc.snapshot.drfxLink && doc.snapshot.drfxLink.linked;
      return !linked;
    });
  }

  async function exportToDrfxMulti() {
    if (!IS_ELECTRON || !nativeBridge || !nativeBridge.ipcRenderer) {
      error('Export to DRFX is available in the desktop app.');
      return;
    }
    const active = getActiveDocument();
    if (active) storeDocumentSnapshot(active);
    const selected = documents.filter(doc => doc && doc.selected);
    const candidates = selected.length ? selected : documents;
    const looseDocs = getLooseDocumentsFrom(candidates);
    if (looseDocs.length < 2) {
      error('Open at least two loose macros to export as a DRFX pack.');
      return;
    }
    let defaultName = 'Macro Machine Export';
    let defaultCategory = 'Edit/Effects/Stirling Supply Co';
    try {
      const storedName = localStorage.getItem('fmr.drfxExportName');
      const storedCategory = localStorage.getItem('fmr.drfxExportCategory');
      if (storedName) defaultName = storedName;
      if (storedCategory) defaultCategory = storedCategory;
    } catch (_) {}
    const drfxName = (window.prompt('DRFX name:', defaultName) || '').trim();
    if (!drfxName) return;
    const categoryRaw = window.prompt('Category path inside DRFX (e.g., Edit/Effects/Your Brand):', defaultCategory);
    if (categoryRaw == null) return;
    const categoryPath = String(categoryRaw || '').trim();
    if (!categoryPath) {
      error('Category path is required.');
      return;
    }
    try {
      localStorage.setItem('fmr.drfxExportName', drfxName);
      localStorage.setItem('fmr.drfxExportCategory', categoryPath);
    } catch (_) {}

    const presetInputs = looseDocs.map((doc) => {
      const snap = doc.snapshot;
      const name = snap.parseResult?.macroName || doc.name || doc.fileName || 'Preset';
      const content = rebuildContentWithNewOrder(snap.originalText || '', snap.parseResult, snap.newline);
      return { doc, name, content };
    });
    const presets = buildUniquePresetNames(presetInputs.map(({ name, content }) => ({ name, content })));
    const payload = {
      drfxName,
      categoryPath,
      presets,
    };
    try {
      const res = await nativeBridge.ipcRenderer.invoke('drfx-export-pack', payload);
      if (!res || !res.ok || !res.filePath) {
        error(res?.error || 'Failed to export DRFX pack.');
        return;
      }
      const exportedPath = res.filePath;
      looseDocs.forEach((doc) => {
        if (doc.snapshot) {
          doc.snapshot.lastExportPath = exportedPath;
          doc.isDirty = false;
        }
      });
      state.lastExportPath = exportedPath;
      const activeDoc = getActiveDocument();
      if (activeDoc) storeDocumentSnapshot(activeDoc);
      updateDocExportPathDisplay();
      renderDocTabs();
      info(`Exported DRFX pack to ${exportedPath}`);
    } catch (err) {
      error(err?.message || err || 'Failed to export DRFX pack.');
    }
  }

  exportBtn?.addEventListener('click', async () => {
    if (state.drfxLink && state.drfxLink.linked) {
      exportToSourceDrfx();
    } else {
      exportToFile();
    }
  });

  exportTemplatesBtn?.addEventListener('click', async () => {
    exportToEditPage();
  });



  // Import .setting content from clipboard (with native/browse fallback)
  importClipboardBtn?.addEventListener('click', async () => {
    try {
      let text = '';
      const native = (typeof window !== 'undefined') ? window.FusionMacroReordererNative : null;

      // 1) Try native clipboard (Electron)
      if (native && typeof native.readClipboard === 'function') {
        try { text = native.readClipboard() || ''; } catch (_) { text = ''; }
      }

      // 2) If still empty, try browser clipboard
      if ((!text || !text.trim()) && navigator && navigator.clipboard && navigator.clipboard.readText) {
        try { text = await navigator.clipboard.readText(); } catch (_) { /* ignore */ }
      }

      // 3) If still empty: browser can prompt, Electron just errors
      if (!text || !text.trim()) {
        const isElectron = !!(native && native.isElectron);
        if (!isElectron) {
          const pasted = window.prompt('Paste .setting content here:', '');
          if (!pasted || !pasted.trim()) {
            error('Clipboard empty or read denied.');
            return;
          }
          text = pasted;
        } else {
          error('Clipboard empty or read denied.');
          return;
        }
      }

      await loadSettingText('Clipboard.setting', text);
      info('Loaded macro from clipboard');
    } catch (err) {
      error('Clipboard import failed: ' + (err?.message || err));
    }
  });



  // Accept drags from Nodes pane onto Published list
  if (ENABLE_NODE_DRAG && controlsList) {
    controlsList.addEventListener('dragover', (e) => {
      if (!state.parseResult || !nodesPane || !nodesPane.isNodeDragEvent?.(e)) return;
      e.preventDefault();
      try { e.dataTransfer && (e.dataTransfer.dropEffect = 'copy'); } catch (_) {}
      updateDropIndicatorPosition(e, state.parseResult.order);
    });

    controlsList.addEventListener('drop', (e) => {
      if (!state.parseResult || !nodesPane || !nodesPane.isNodeDragEvent?.(e)) return;
      e.preventDefault();
      const payload = nodesPane.parseNodeDragData?.(e);
      if (!payload) { cleanupDropIndicator(); return; }
      const items = nodesPane.buildItemsFromPayload?.(payload) || [];
      if (items.length) addOrMovePublishedItemsAt(items, getCurrentDropIndex());
      cleanupDropIndicator();
    });
  }



  exportClipboardBtn?.addEventListener('click', async () => {
    exportToClipboard();
  });



    // Deep-link and cross-pane highlight

  function linkKey(sourceOp, source) {

    if (!sourceOp) return null; const s = source ? ('.' + source) : ''; return `${sourceOp}${s}`;

  }

  function setDeepLink(op, src) {

    const key = linkKey(op, src); if (!key) return; location.hash = encodeURIComponent(key);

  }

  function parseDeepLink() {

    const h = decodeURIComponent((location.hash || '').replace(/^#/, ''));

    if (!h) return null; const dot = h.indexOf('.');

    if (dot < 0) return { op: h, src: null }; return { op: h.slice(0, dot), src: h.slice(dot+1) };

  }

  function clearHighlights() {

    try {

      document.querySelectorAll('.hl-pub').forEach(e => e.classList.remove('hl-pub'));

      document.querySelectorAll('.hl-node').forEach(e => e.classList.remove('hl-node'));

    } catch (_) {}

  }

  // Safe CSS attribute value escape for selectors

  function cssEsc(v) {

    try {

      return (window.CSS && CSS.escape) ? CSS.escape(String(v)) : String(v).replace(/["\\\]]/g, '\\$&');

    } catch (_) {

      return String(v);

    }

  }

  function highlightPublished(op, src, pulse=true) {

  try {

    if (!state.parseResult) return false;

    const idx = (state.parseResult.entries||[]).findIndex(e => e && e.sourceOp===op && e.source===src);

    if (idx < 0) return false;

    const row = controlsList.querySelector('li[data-index="'+idx+'"]');

    if (!row) return false;

    row.classList.add('hl-pub');

    if (pulse) row.classList.add('pulse');

    scrollRowIntoView(row, controlsList);

    setTimeout(()=>row.classList.remove('pulse'), 800);

    return true;

  } catch (_) { return false; }

  }

  function highlightNode(op, src, pulse = true) {

    try {

      if (nodesPane) {
        nodesPane.expandNode?.(op, 'open');
        nodesPane.parseAndRenderNodes?.();
      }

      const findEl = () => {

        if (!nodesList) return null;

        if (src) {

          let el = nodesList.querySelector('input.node-ctrl[data-source-op="' + op + '"][data-source="' + src + '"]');

          if (el) return el;

          let base = null;

          try { base = (typeof deriveColorBaseFromId === 'function') ? deriveColorBaseFromId(src) : null; } catch(_) { base = null; }

          if (!base) {

            const suf = ['Red','Green','Blue','Alpha'].find(s => String(src||'').endsWith(s));

            if (suf) base = String(src).slice(0, String(src).length - suf.length);

          }

          if (base) {

            el = nodesList.querySelector('input.node-ctrl.group[data-source-op="' + op + '"][data-group-base="' + base + '"]');

            if (el) return el;

          }

          const inputs = nodesList.querySelectorAll('input.node-ctrl');

          for (const input of inputs) {

            if (!input.classList.contains('group')) {

              if ((input.dataset.sourceOp || '') === op && (input.dataset.source || '') === src) return input;

            } else {

              const dsBase = input.dataset.groupBase || '';

              if (dsBase && (input.dataset.sourceOp || '') === op) {

                if (!base) {

                  const suf = ['Red','Green','Blue','Alpha'].find(s => String(src||'').endsWith(s));

                  if (suf) base = String(src).slice(0, String(src).length - suf.length);

                }

                if (base && dsBase === base) return input;

              }

            }

          }

          return null;

        } else {

          const el = nodesList.querySelector('input.node-ctrl.group[data-source-op="' + op + '"]');

          if (el) return el;

          const inputs = nodesList.querySelectorAll('input.node-ctrl.group');

          for (const input of inputs) {

            if ((input.dataset.sourceOp || '') === op) return input;

          }

          return null;

        }

      };

      let target = findEl();

      if (!target && nodesPane?.getNodeFilter?.()) {
        nodesPane?.clearFilter?.();
        target = findEl();
      }

      if (!target) {

        // Fallback: highlight the node header by data attribute

        try {

          const wrap = nodesList && nodesList.querySelector('.node[data-op="' + op + '"] .node-header');

          if (wrap) {

            wrap.classList.add('hl-node');

            if (pulse) wrap.classList.add('pulse');

            scrollRowIntoView(wrap, nodesList);

            setTimeout(() => wrap.classList.remove('pulse'), 800);

            return true;

          }

        } catch(_) {}

        return false;

      }

      const row = target.closest('.node-row') || target;

      row.classList.add('hl-node');

      if (pulse) row.classList.add('pulse');

      scrollRowIntoView(row, nodesList);

      setTimeout(() => row.classList.remove('pulse'), 800);

      return true;

    } catch(_) { return false; }

  }

  function scrollRowIntoView(row, container) {
    try {
      if (!row) return;
      const host = container || row.closest('.nodes-list') || row.closest('.list');
      if (host) {
        const hostRect = host.getBoundingClientRect();
        const rowRect = row.getBoundingClientRect();
        const fullyVisible = rowRect.top >= hostRect.top && rowRect.bottom <= hostRect.bottom;
        if (fullyVisible) return;
        const offset = (rowRect.top - hostRect.top) + host.scrollTop;
        const targetTop = Math.max(0, offset - host.clientHeight / 2);
        host.scrollTo({ top: targetTop, behavior: 'smooth' });
        return;
      }
      row.scrollIntoView({ block: 'center', behavior: 'smooth' });
    } catch (_) {
      try { row.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch (_) {}
    }
  }

  try {
    nodesPane?.setHighlightHandler?.(highlightNode);
    publishedControls?.setHighlightHandler?.(highlightNode);
  } catch (_) {}

  function applyDeepLinkFromHash() {

    try { const p = parseDeepLink(); if (!p || !state.parseResult) return; if (!highlightPublished(p.op, p.src)) { highlightNode(p.op, p.src); } } catch (_) {}

}

  function clearUI() {

    messages.textContent = '';

    controlsList.innerHTML = '';

    controlsSection.hidden = true;
    if (pageTabsEl) {
      pageTabsEl.hidden = true;
      pageTabsEl.innerHTML = '';
    }

    resetBtn.disabled = true;

    exportBtn.disabled = true;
    setExportButtonLabel(DEFAULT_EXPORT_LABEL);
    exportTemplatesBtn && (exportTemplatesBtn.disabled = true);
    if (operatorTypeSelect) operatorTypeSelect.value = 'GroupOperator';
    removeCommonPageBtn && (removeCommonPageBtn.disabled = true);
    deselectAllBtn && (deselectAllBtn.disabled = true);
    hideDetailDrawer();
    updateUtilityActionsState();
    updateExportMenuButtonState();
    cancelPendingCodeRefresh();
    updateCodeView('');
    state.lastDiffRange = null;
    pendingControlMeta.clear();
    pendingOpenNodes.clear();

  }

  function updateUtilityActionsState() {
    if (addUtilityNodeBtn) {
      addUtilityNodeBtn.disabled = !state.parseResult;
    }
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function updateCodeView(text) {
    if (!codeView) return;
    const next = text == null ? '' : String(text);
    state.generatedText = next;
    codeHighlightActive = false;
    codeView.textContent = next;
  }

  function clearCodeHighlight() {
    if (!codeView) return;
    codeHighlightActive = false;
    codeView.textContent = state.generatedText || '';
  }

  function activateCodeHighlight(range, index) {
    if (!codeView) return;
    const text = state.generatedText || '';
    const length = text.length;
    const start = Math.max(0, Math.min(length, Number(range?.start) || 0));
    const end = Math.max(start, Math.min(length, Number(range?.end) || start));
    if (end <= start) return;
    const before = escapeHtml(text.slice(0, start));
    const target = escapeHtml(text.slice(start, end)) || '&nbsp;';
    const after = escapeHtml(text.slice(end));
    codeView.innerHTML = `${before}<mark class="code-highlight">${target}</mark>${after}`;
    codeHighlightActive = true;
    requestAnimationFrame(() => {
      const mark = codeView.querySelector('mark.code-highlight');
      if (mark && typeof mark.offsetTop === 'number') {
        const container = codeView;
        const containerRect = container.getBoundingClientRect();
        const markRect = mark.getBoundingClientRect();
        const currentScroll = container.scrollTop || 0;
        const delta = markRect.top - containerRect.top;
        const targetTop = Math.max(0, currentScroll + delta - container.clientHeight / 2);
        if (typeof container.scrollTo === 'function') {
          container.scrollTo({ top: targetTop, behavior: 'smooth' });
        } else {
          container.scrollTop = targetTop;
        }
      }
    });
  }

  function applyPendingHighlight() {
    const range = state.lastDiffRange;
    state.lastDiffRange = null;
    if (range) {
      activateCodeHighlight(range, null);
    } else if (codeHighlightActive) {
      clearCodeHighlight();
    }
  }

  function computeDiffRange(oldText, newText) {
    try {
      const nextText = newText == null ? '' : String(newText);
      const prevText = oldText == null ? '' : String(oldText);
      if (nextText === prevText) return null;
      const minLen = Math.min(prevText.length, nextText.length);
      let prefix = 0;
      while (prefix < minLen && prevText[prefix] === nextText[prefix]) prefix++;
      if (prefix >= nextText.length && nextText.length === prevText.length) return null;
      const start = alignRangeStart(nextText, prefix);
      let rawEnd = findRealignPoint(prevText, nextText, prefix);
      if (rawEnd == null) {
        let endPrev = prevText.length - 1;
        let endNext = nextText.length - 1;
        while (endPrev >= prefix && endNext >= prefix && prevText[endPrev] === nextText[endNext]) {
          endPrev--;
          endNext--;
        }
        rawEnd = endNext + 1;
      }
      if (rawEnd <= prefix) rawEnd = Math.min(nextText.length, prefix + 240);
      rawEnd = Math.min(nextText.length, rawEnd);
      let end = alignRangeEnd(nextText, rawEnd);
      end = limitRangeSpan(nextText, start, end);
      return { start, end };
    } catch (_) {
      return null;
    }
  }

  function alignRangeStart(text, index) {
    if (!text) return 0;
    let i = Math.max(0, Math.min(text.length, index || 0));
    while (i > 0 && text[i - 1] !== '\n') i--;
    return i;
  }

  function alignRangeEnd(text, index) {
    if (!text) return 0;
    let i = Math.max(0, Math.min(text.length, index || 0));
    while (i < text.length && text[i] !== '\n') i++;
    if (i < text.length) i++;
    return Math.max(i, index || 0);
  }

  function findRealignPoint(prevText, nextText, start) {
    const WINDOW = 160;
    const fromPrev = prevText.slice(start, Math.min(prevText.length, start + WINDOW));
    if (hasMeaningfulContent(fromPrev)) {
      const idx = nextText.indexOf(fromPrev, start + 1);
      if (idx >= 0) return idx;
    }
    const fromNext = nextText.slice(start, Math.min(nextText.length, start + WINDOW));
    if (hasMeaningfulContent(fromNext)) {
      const idxPrev = prevText.indexOf(fromNext, start + 1);
      if (idxPrev >= 0) {
        const delta = idxPrev - start;
        if (delta > 0) return Math.min(nextText.length, start + delta);
      }
    }
    return null;
  }

  function hasMeaningfulContent(str) {
    return !!str && /[^\s]/.test(str);
  }

  function limitRangeSpan(text, start, desiredEnd) {
    if (!text) return start;
    const MAX_CHAR_SPAN = 4000;
    const MAX_LINE_SPAN = 80;
    const MIN_SPAN = 32;
    let end = Math.min(text.length, Math.max(start + 1, desiredEnd || start));
    if (end - start > MAX_CHAR_SPAN) {
      end = start + MAX_CHAR_SPAN;
    }
    let lines = 0;
    for (let i = start; i < end; i++) {
      if (text.charCodeAt(i) === 10) {
        lines++;
        if (lines >= MAX_LINE_SPAN) {
          end = i + 1;
          break;
        }
      }
    }
    if (end <= start) {
      end = Math.min(text.length, start + MIN_SPAN);
    }
    return end;
  }

  function cancelPendingCodeRefresh() {
    if (codePaneRefreshTimer) {
      clearTimeout(codePaneRefreshTimer);
      codePaneRefreshTimer = null;
    }
  }

  function scheduleCodePaneRefresh() {
    if (!state.parseResult || !state.originalText) return;
    cancelPendingCodeRefresh();
    codePaneRefreshTimer = setTimeout(() => {
      codePaneRefreshTimer = null;
      try {
        const previousText = state.generatedText || '';
        const rebuilt = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline);
        state.lastDiffRange = computeDiffRange(previousText, rebuilt);
        updateCodeView(rebuilt);
        applyPendingHighlight();
      } catch (err) {
        try { logDiag(`[CodePane] rebuild failed: ${err?.message || err}`); } catch (_) {}
      }
    }, 80);
  }

  function markContentDirty() {
    scheduleCodePaneRefresh();
    if (suppressDocDirty) return;
    markActiveDocumentDirty();
    const active = getActiveDocument();
    if (active) storeDocumentSnapshot(active);
  }



  function info(msg) {

    const div = document.createElement('div');

    div.textContent = msg;

    messages.appendChild(div);

  }

  function error(msg) {

    const div = document.createElement('div');

    div.style.color = '#ff9a9a';

    div.textContent = `Error: ${msg}`;

    messages.appendChild(div);

  }



  function detectNewline(text) {

    const idx = text.indexOf('\r\n');

    if (idx >= 0) return '\r\n';

    return '\n';

  }



  function setAutoUtilityEnabled(on) {

    autoUtilityEnabled = !!on;

    try {

      if (typeof localStorage !== 'undefined') {

        localStorage.setItem(AUTO_UTILITY_KEY, autoUtilityEnabled ? '1' : '0');

      }

    } catch (_) {}

  }



  function isAutoUtilityEnabled() {

    return !!autoUtilityEnabled;

  }



  async function ensureUtilityTemplate() {

    if (utilityTemplateText != null) return utilityTemplateText;

    try {

      const resp = await fetch(UTILITY_TEMPLATE_PATH);

      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);

      utilityTemplateText = await resp.text();

    } catch (err) {

      utilityTemplateText = '';

      try { logDiag(`Utility template fetch failed: ${err?.message || err}`); } catch (_) {}

    }

    return utilityTemplateText;

  }



  function maybeWrapUngroupedNodes(text, options = {}) {
    try {
      if (!text) return { text, changed: false };
      if (/\=\s*(GroupOperator|MacroOperator)\s*\{/.test(text)) return { text, changed: false };
      const toolsPos = text.indexOf('Tools = ordered()');
      if (toolsPos < 0) return { text, changed: false };
      const toolsOpen = text.indexOf('{', toolsPos);
      if (toolsOpen < 0) return { text, changed: false };
      const toolsClose = findMatchingBrace(text, toolsOpen);
      if (toolsClose < 0) return { text, changed: false };
      const toolsInner = text.slice(toolsOpen + 1, toolsClose);
      if (!toolsInner.trim()) return { text, changed: false };
      const toolNames = extractTopLevelToolNames(text, toolsOpen, toolsClose);
      if (!toolNames.length) return { text, changed: false };
      const newline = detectNewline(text);
      const macroName = deriveAutoMacroName(options?.sourceName || '');
      const blockIndent = getLineIndent(text, toolsPos) || '\t';
      const operatorIndent = blockIndent + '\t';
      const toolIndent = operatorIndent + '\t';
      const bounds = computeToolBounds(text, toolsOpen + 1, toolsClose);
      let groupTools = toolsInner;
      if (bounds) {
        const centerX = (bounds.minX + bounds.maxX) / 2;
        const centerY = (bounds.minY + bounds.maxY) / 2;
        if (Number.isFinite(centerX) && Number.isFinite(centerY)) {
          groupTools = shiftToolPositions(toolsInner, centerX, centerY);
        }
      }
      const reindentedTools = reindent(groupTools, toolIndent, newline);
      const outputNode = guessOutputNodeName(text, toolNames);
      const macroParts = [];
      macroParts.push(`${newline}${operatorIndent}${macroName} = GroupOperator {`);
      macroParts.push(`${operatorIndent}\tCtrlWZoom = false,`);
      macroParts.push(`${operatorIndent}\tNameSet = true,`);
      macroParts.push(`${operatorIndent}\tInputs = ordered() {`);
      macroParts.push(`${operatorIndent}\t},`);
      macroParts.push(`${operatorIndent}\tOutputs = {`);
      if (outputNode) {
        macroParts.push(`${operatorIndent}\t\tMainOutput1 = InstanceOutput {`);
        macroParts.push(`${operatorIndent}\t\t\tSourceOp = "${outputNode}",`);
        macroParts.push(`${operatorIndent}\t\t\tSource = "Output",`);
        macroParts.push(`${operatorIndent}\t\t},`);
      }
      macroParts.push(`${operatorIndent}\t},`);
      macroParts.push(...buildGroupViewInfo(bounds, operatorIndent));
      macroParts.push(`${operatorIndent}\tTools = ordered() {`);
      macroParts.push(reindentedTools);
      macroParts.push(`${operatorIndent}\t},`);
      macroParts.push(`${operatorIndent}},`);
      const wrappedGroup = macroParts.join(newline);
      const before = text.slice(0, toolsOpen + 1);
      const originalInner = text.slice(toolsOpen + 1, toolsClose);
      const after = text.slice(toolsClose);
      const firstChar = originalInner ? originalInner[0] : '';
      const needsSeparator = firstChar !== '\n' && firstChar !== '\r';
      const newInner = `${wrappedGroup}${needsSeparator ? (newline || '\n') : ''}${originalInner}`;
      let updated = `${before}${newInner}${after}`;
      updated = ensureActiveToolName(updated, macroName, newline);
      return { text: updated, changed: true };
    } catch (_) {
      return { text, changed: false };
    }
  }

  function deriveAutoMacroName(sourceName) {
    try {
      const raw = sourceName ? String(sourceName).replace(/\.[^.]+$/, '') : 'ImportedMacro';
      const cleaned = sanitizeIdent(raw || 'ImportedMacro');
      return cleaned || 'ImportedMacro';
    } catch (_) {
      return 'ImportedMacro';
    }
  }

  function extractTopLevelToolNames(text, toolsOpen, toolsClose) {
    try {
      const names = [];
      let cursor = toolsOpen + 1;
      while (cursor < toolsClose) {
        while (cursor < toolsClose && (text[cursor] === ',' || isSpace(text[cursor]))) cursor++;
        if (cursor >= toolsClose) break;
        if (!isIdentStart(text[cursor])) { cursor++; continue; }
        const start = cursor;
        cursor++;
        while (cursor < toolsClose && isIdentPart(text[cursor])) cursor++;
        const name = text.slice(start, cursor);
        let pos = cursor;
        while (pos < toolsClose && isSpace(text[pos])) pos++;
        if (text[pos] !== '=') { cursor = pos + 1; continue; }
        pos++;
        while (pos < toolsClose && isSpace(text[pos])) pos++;
        if (!isIdentStart(text[pos])) { cursor = pos + 1; continue; }
        while (pos < toolsClose && isIdentPart(text[pos])) pos++;
        while (pos < toolsClose && isSpace(text[pos])) pos++;
        if (text[pos] !== '{') { cursor = pos + 1; continue; }
        const blockOpen = pos;
        const blockClose = findMatchingBrace(text, blockOpen);
        if (blockClose < 0 || blockClose > toolsClose) break;
        names.push(name);
        cursor = blockClose + 1;
      }
      return names;
    } catch (_) {
      return [];
    }
  }

  function guessOutputNodeName(text, toolNames) {
    try {
      if (!toolNames || !toolNames.length) return null;
      const counts = new Map();
      toolNames.forEach((name) => {
        const re = new RegExp(`SourceOp\\s*=\\s*"${escapeRegex(name)}"`, 'g');
        let match;
        let total = 0;
        while ((match = re.exec(text)) !== null) {
          total++;
          if (re.lastIndex === match.index) re.lastIndex++;
        }
        counts.set(name, total);
      });
      const zeroRefs = toolNames.filter(name => (counts.get(name) || 0) === 0);
      if (zeroRefs.length) return zeroRefs[zeroRefs.length - 1];
      return toolNames[toolNames.length - 1];
    } catch (_) {
      return toolNames && toolNames.length ? toolNames[toolNames.length - 1] : null;
    }
  }

  function computeToolBounds(text, startIndex, endIndex) {
    try {
      if (startIndex == null || endIndex == null || endIndex <= startIndex) return null;
      const slice = text.slice(startIndex, endIndex);
      const re = /ViewInfo\s*=\s*OperatorInfo\s*\{\s*Pos\s*=\s*\{\s*([-0-9.+]+)\s*,\s*([-0-9.+]+)/g;
      let match;
      let minX = Infinity;
      let maxX = -Infinity;
      let minY = Infinity;
      let maxY = -Infinity;
      while ((match = re.exec(slice)) !== null) {
        const x = parseFloat(match[1]);
        const y = parseFloat(match[2]);
        if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
        if (re.lastIndex === match.index) re.lastIndex++;
      }
      if (!Number.isFinite(minX) || !Number.isFinite(minY) || !Number.isFinite(maxX) || !Number.isFinite(maxY)) {
        return null;
      }
      return { minX, minY, maxX, maxY };
    } catch (_) {
      return null;
    }
  }

  function shiftToolPositions(text, offsetX, offsetY) {
    try {
      if (!text || !Number.isFinite(offsetX) || !Number.isFinite(offsetY)) return text;
      const fmt = (num) => Number(num.toFixed(3));
      let result = '';
      let cursor = 0;
      const regex = /ViewInfo\s*=\s*OperatorInfo\s*\{/g;
      let match;
      while ((match = regex.exec(text)) !== null) {
        const braceOpen = text.indexOf('{', match.index);
        if (braceOpen < 0) break;
        const braceClose = findMatchingBrace(text, braceOpen);
        if (braceClose < 0) break;
        const inner = text.slice(braceOpen + 1, braceClose);
        const posRegex = /Pos\s*=\s*\{\s*([-0-9.+]+)\s*,\s*([-0-9.+]+)\s*\}/;
        let newInner = inner;
        const posMatch = posRegex.exec(inner);
        if (posMatch) {
          const x = parseFloat(posMatch[1]);
          const y = parseFloat(posMatch[2]);
          if (Number.isFinite(x) && Number.isFinite(y)) {
            const newX = fmt(x - offsetX);
            const newY = fmt(y - offsetY);
            newInner = inner.replace(posRegex, `Pos = { ${newX}, ${newY} }`);
          }
        }
        result += text.slice(cursor, braceOpen + 1) + newInner;
        cursor = braceClose;
        regex.lastIndex = braceClose;
      }
      return result + text.slice(cursor);
    } catch (_) {
      return text;
    }
  }

  function buildGroupViewInfo(bounds, indent) {
    const spanX = bounds ? Math.max(1, bounds.maxX - bounds.minX) : 1;
    const spanY = bounds ? Math.max(1, bounds.maxY - bounds.minY) : 1;
    const paddedWidth = spanX + 400;
    const paddedHeight = spanY + 400;
    const fitWidth = 900;
    const fitHeight = 500;
    const scaleX = fitWidth / paddedWidth;
    const scaleY = fitHeight / paddedHeight;
    const scale = bounds ? Math.min(1, scaleX, scaleY) : 1;
    const lines = [];
    lines.push(`${indent}\tViewInfo = GroupInfo {`);
    lines.push(`${indent}\t\tPos = { 0, 0 },`);
    lines.push(`${indent}\t\tFlags = {`);
    lines.push(`${indent}\t\t\tExpanded = true,`);
    lines.push(`${indent}\t\t\tAllowPan = false,`);
    lines.push(`${indent}\t\t\tAutoSnap = true,`);
    lines.push(`${indent}\t\t\tRemoveRouters = true`);
    lines.push(`${indent}\t\t},`);
    lines.push(`${indent}\t\tSize = { 300, 150, 200, 40 },`);
    lines.push(`${indent}\t\tDirection = "Horizontal",`);
    lines.push(`${indent}\t\tPipeStyle = "Direct",`);
    lines.push(`${indent}\t\tScale = ${Number(scale.toFixed(3))},`);
    lines.push(`${indent}\t\tOffset = { 0, 0 }`);
    lines.push(`${indent}\t},`);
    return lines;
  }

  function ensureActiveToolName(text, macroName, newline) {
    try {
      if (!text || !macroName) return text;
      const re = /ActiveTool\s*=\s*"([^"]*)"/;
      if (re.test(text)) return text.replace(re, `ActiveTool = "${macroName}"`);
      const nl = newline || detectNewline(text) || '\n';
      const trimmed = text.replace(/\s+$/, '');
      const lastBrace = trimmed.lastIndexOf('}');
      if (lastBrace >= 0) {
        const indent = getLineIndent(trimmed, lastBrace) || '\t';
        let before = trimmed.slice(0, lastBrace);
        let needsComma = true;
        for (let i = before.length - 1; i >= 0; i--) {
          const ch = before[i];
          if (isSpace(ch)) continue;
          if (ch === ',') needsComma = false;
          break;
        }
        const insertion = `${needsComma ? ',' : ''}${nl}${indent}ActiveTool = "${macroName}"${nl}`;
        return `${before}${insertion}${trimmed.slice(lastBrace)}`;
      }
      return `${trimmed}${nl}ActiveTool = "${macroName}"${nl}`;
    } catch (_) {
      return text;
    }
  }

  function macroHasUtilityNode(text) {

    if (!text) return false;

    return /\bUTILITY\s*=/.test(text);

  }



  function injectUtilityNodeText(text, template, newline) {

    try {

      if (!text || !template || !template.trim()) return text;

      if (macroHasUtilityNode(text)) return text;

      const toolsPos = text.indexOf('Tools = ordered()');

      if (toolsPos < 0) return text;

      const braceOpen = text.indexOf('{', toolsPos);

      if (braceOpen < 0) return text;

      const braceClose = findMatchingBrace(text, braceOpen);

      if (braceClose < 0) return text;

      const blockIndent = getLineIndent(text, toolsPos) || '';

      const entryIndent = blockIndent + '\t';

      const normalized = reindent(template.trim(), entryIndent, newline || '\n');

      const insert = `${newline || '\n'}${normalized}${newline || '\n'}`;

      return text.slice(0, braceClose) + insert + text.slice(braceClose);

    } catch (_) {

      return text;

    }

  }



  async function maybeAutoAddUtility(text, options = {}) {

    const { silent = false } = options || {};

    if (!isAutoUtilityEnabled() || !text || macroHasUtilityNode(text)) return text;

    const tpl = await ensureUtilityTemplate();

    if (!tpl || !tpl.trim()) return text;

    const updated = injectUtilityNodeText(text, tpl, detectNewline(text));

    if (updated !== text && !silent) info('Utility node added automatically.');

    return updated;

  }



  function suggestOutputName(name) {

    const dot = name.lastIndexOf('.');

    if (dot > 0) return name.slice(0, dot) + '.reordered' + name.slice(dot);

    return name + '.reordered.setting';

  }



  function triggerDownload(fileName, content) {

    const blob = new Blob([content], { type: 'text/plain' });

    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');

    a.href = url;

    a.download = fileName;

    document.body.appendChild(a);

    a.click();

    URL.revokeObjectURL(url);

    a.remove();

  }



  function rebuildContentWithNewOrder(original, result, eol) {
    const { entries } = result;
    const exportOrder = getOrderedEntryIndices(result);
    try {
      if (typeof logDiag === 'function') {
        const orderedKeys = exportOrder.map(i => {
          const e = entries[i];
          return e ? `${e.sourceOp || ''}::${e.source || ''}` : `?${i}`;
        });
        logDiag(`[Export] order=${orderedKeys.join(', ')}`);
      }
    } catch (_) {}
    syncBlendToggleFlags(result);
    // Prefer rebuilding into the GroupOperator-level Inputs block (or inserting one)
    let bounds = locateMacroGroupBounds(original, result);
    let updated = original;
    updated = applyLabelCountEdits(updated, result, eol);
    updated = applyLabelVisibilityEdits(updated, result, eol);
    updated = applyLabelDefaultStateEdits(updated, result, eol);
    updated = applyUserControlPages(updated, result, eol);
    // Safety: strip BTNCS_Execute from managed buttons unless explicitly clicked
    updated = stripUnclickedBtncs(updated, result, eol);
    // Insert exact launcher only for explicitly clicked controls
    updated = ensureExactLauncherInserted(updated, result, eol);
    updated = stripLegacyLauncherArtifacts(updated);
    updated = applyMacroNameRename(updated, result);
    updated = ensureGroupInputsBlock(updated, result, eol);
    updated = rewritePrimaryInputsBlock(updated, result, eol);
    return updated;
  }

  function ensureGroupInputsBlock(text, result, eol) {
    try {
      if (!text) return text;
      if (text.includes('Inputs = ordered()')) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds || bounds.groupOpenIndex == null || bounds.groupCloseIndex == null) return text;
      let insertPos = text.indexOf('Outputs =', bounds.groupOpenIndex);
      if (insertPos < 0 || insertPos > bounds.groupCloseIndex) {
        insertPos = text.indexOf('ViewInfo =', bounds.groupOpenIndex);
      }
      if (insertPos < 0 || insertPos > bounds.groupCloseIndex) {
        insertPos = text.indexOf('Tools = ordered()', bounds.groupOpenIndex);
      }
      if (insertPos < 0 || insertPos > bounds.groupCloseIndex) {
        insertPos = bounds.groupOpenIndex + 1;
      }
      const newline = eol || detectNewline(text) || '\n';
      const baseIndent = getLineIndent(text, bounds.groupOpenIndex + 1) || '\t';
      const innerIndent = baseIndent + '\t';
      const needsLeadingNewline = insertPos > bounds.groupOpenIndex + 1 && text[insertPos - 1] !== '\n';
      const snippetParts = [];
      if (needsLeadingNewline) snippetParts.push(newline);
      snippetParts.push(`${innerIndent}Inputs = ordered() {`);
      snippetParts.push(`${newline}${innerIndent}\t`);
      snippetParts.push(`${newline}${innerIndent}},${newline}`);
      const snippet = snippetParts.join('');
      return text.slice(0, insertPos) + snippet + text.slice(insertPos);
    } catch (_) {
      return text;
    }
  }

  function rewritePrimaryInputsBlock(text, result, eol) {
    try {
      if (!text || !result || !Array.isArray(result.entries) || !result.entries.length) return text;
      const exportOrder = getOrderedEntryIndices(result);
      if (!exportOrder.length) return text;
      const idx = text.indexOf('Inputs = ordered()');
      if (idx < 0) return text;
      const braceOpen = text.indexOf('{', idx);
      if (braceOpen < 0) return text;
      const braceClose = findMatchingBrace(text, braceOpen);
      if (braceClose < 0) return text;
      let after = braceClose + 1;
      let trailingWhitespace = '';
      while (after < text.length && /\s/.test(text[after])) {
        trailingWhitespace += text[after];
        after++;
      }
      let hasComma = false;
      if (text[after] === ',') {
        hasComma = true;
        after++;
        // preserve any whitespace that originally followed the comma
        while (after < text.length && /\s/.test(text[after])) {
          trailingWhitespace += text[after];
          after++;
        }
      }
      const blockIndent = getLineIndent(text, idx) || '';
      const innerIndent = blockIndent + '\t';
      const entryIndent = innerIndent + '\t';
      const newline = eol || detectNewline(text) || '\n';
      const segments = [];
      exportOrder.forEach(i => {
        const entry = result.entries[i];
        if (!entry) return;
        let raw = applyEntryOverrides(entry, applyNameIfEdited(entry, eol), eol);
        raw = ensureInstanceInputKey(raw, entry);
        let chunk = reindent(raw, entryIndent, eol);
        if (!chunk || !chunk.trim().length) return;
        chunk = ensureTrailingComma(chunk);
        segments.push({ index: i, chunk });
      });
      const header = `${blockIndent}Inputs = ordered() {`;
      const inner = segments.length ? newline + newline + segments.map(seg => seg.chunk).join(newline) + newline + innerIndent : '';
      let replacement = `${header}${inner}${blockIndent}}`;
      if (hasComma) replacement += ',';
      replacement += trailingWhitespace || newline;
      if (typeof logDiag === 'function' && diagnosticsController?.isEnabled?.()) {
        try {
          const orderedKeys = exportOrder.map(i => {
            const e = result.entries[i];
            return e ? `${e.sourceOp || ''}::${e.source || ''}` : `?${i}`;
          });
          logDiag(`[Inputs] rewritten order=${orderedKeys.join(', ')}`);
        } catch (_) {}
      }
      return text.slice(0, idx) + replacement + text.slice(after);
    } catch (_) {
      return text;
    }
  }

  function ensureTrailingComma(raw) {
    try {
      const s = String(raw);
      // If already has a trailing comma after the closing brace, keep as is
      if (/}\s*,\s*$/.test(s)) return s;
      // Only add if it terminates with a closing brace
      if (/}\s*$/.test(s)) return s.replace(/}(\s*)$/, '},$1');
      return s;
    } catch(_) { return raw; }
  }

  // Locate the GroupOperator-level Inputs block (top-level inside the group)
  function findGroupInputsBlocks(text, groupOpen, groupClose) {
    try {
      const blocks = [];
      let i = groupOpen + 1;
      let depth = 1;
      let inStr = false;
      while (i < groupClose) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && text[i - 1] !== '\\') inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; continue; }
        if (depth === 1 && text.slice(i, i + 6) === 'Inputs') {
          let j = i + 6;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < groupClose && isSpace(text[j])) j++;
          const maybe = text.slice(j, j + 8);
          if (maybe === 'ordered(') {
            j += 8;
            while (j < groupClose && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < groupClose && isSpace(text[j])) j++;
          }
          if (text[j] !== '{') { i++; continue; }
          const openIndex = j;
          const closeIndex = findMatchingBrace(text, openIndex);
          if (closeIndex > openIndex && closeIndex <= groupClose) {
            blocks.push({ inputsHeaderStart: i, openIndex, closeIndex });
          }
        }
        i++;
      }
      return blocks;
    } catch (_) {
      return [];
    }
  }

  function findGroupUserControlsBlock(text, groupOpen, groupClose) {
    try {
      const toolsPos = text.indexOf('Tools = ordered()', groupOpen);
      const limit = (toolsPos >= 0 && toolsPos < groupClose) ? toolsPos : groupClose;
      let i = groupOpen + 1;
      let depth = 1;
      let inStr = false;
      while (i < limit) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && text[i - 1] !== '\\') inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; if (depth <= 0) break; i++; continue; }
        if (depth === 1 && text.slice(i, i + 12) === 'UserControls') {
          let j = i + 12;
          while (j < limit && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < limit && isSpace(text[j])) j++;
          if (text.slice(j, j + 8) === 'ordered(') {
            j += 8;
            while (j < limit && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < limit && isSpace(text[j])) j++;
          }
          if (text[j] !== '{') { i++; continue; }
          const openIndex = j;
          const closeIndex = findMatchingBrace(text, openIndex);
          if (closeIndex > openIndex && closeIndex <= groupClose) {
            return { openIndex, closeIndex };
          }
        }
        i++;
      }
      return null;
    } catch (_) { return null; }
  }

  function removeCommonPageFromText(text, result, eol) {
    try {
      if (!text || !result) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const configs = new Map();
      configs.set('Common', { visible: false, priority: 0 });
      return ensureControlPagesDeclared(text, bounds, configs, eol || '\n', result);
    } catch (_) { return text; }
  }

  function removeInputBlockByName(text, name) {
    try {
      let updated = text;
      const re = new RegExp(`(^|\\n)(\\s*)${name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}\\s*=\\s*Input\\s*\\{`, 'g');
      let match;
      while ((match = re.exec(updated)) !== null) {
        const start = match.index + match[1].length;
        const braceStart = updated.indexOf('{', start);
        if (braceStart < 0) continue;
        const braceEnd = findMatchingBrace(updated, braceStart);
        if (braceEnd < 0) break;
        let delStart = match.index + match[1].length;
        while (delStart > 0 && /\s/.test(updated[delStart - 1])) delStart--;
        let delEnd = braceEnd + 1;
        while (delEnd < updated.length && /\s/.test(updated[delEnd])) {
          const ch = updated[delEnd];
          delEnd++;
          if (ch === '\n') break;
        }
        updated = updated.slice(0, delStart) + updated.slice(delEnd);
        re.lastIndex = Math.max(0, delStart - 1);
      }
      return updated;
    } catch (_) { return text; }
  }

  function removeInstanceInputsInRange(text, openIndex, closeIndex) {
    try {
      const start = openIndex + 1;
      const end = closeIndex;
      let inner = text.slice(start, end);
      // remove key = InstanceInput { ... } blocks repeatedly
      let changed = false;
      while (true) {
        const m = inner.match(/([A-Za-z_][A-Za-z0-9_]*)\s*=\s*InstanceInput\s*\{/);
        if (!m) break;
        const relStart = m.index;
        const absStart = start + relStart;
        const braceOpen = absStart + m[0].lastIndexOf('{');
        const braceClose = findMatchingBrace(text, braceOpen);
        if (braceClose < 0) break;
        // include optional trailing comma
        let delEnd = braceClose + 1;
        if (text[delEnd] === ',') delEnd++;
        text = text.slice(0, absStart) + text.slice(delEnd);
        // recompute inner
        const newEnd = (closeIndex - (delEnd - absStart));
        inner = text.slice(start, newEnd);
        closeIndex = newEnd + 0; // update close for subsequent operations
        changed = true;
      }
      return text;
    } catch(_) { return text; }
  }

  function removeBtncsExecuteInRange(text, openBraceIndex, closeBraceIndex) {
    try {
      const start = openBraceIndex + 1;
      const end = closeBraceIndex;
      const cleaned = stripBtncsEntries(text.slice(start, end));
      return text.slice(0, start) + cleaned + text.slice(end);
    } catch(_) { return text; }
  }

  function stripBtncsEntries(str) {
    try {
      if (!str || str.indexOf('BTNCS_Execute') < 0) return str;
      const key = 'BTNCS_Execute';
      let out = '';
      let i = 0;
      while (i < str.length) {
        const idx = str.indexOf(key, i);
        if (idx < 0) {
          out += str.slice(i);
          break;
        }
        out += str.slice(i, idx);
        let cursor = idx + key.length;
        while (cursor < str.length && /\s/.test(str[cursor])) cursor++;
        if (cursor >= str.length || str[cursor] !== '=') {
          out += key;
          i = idx + key.length;
          continue;
        }
        cursor++;
        while (cursor < str.length && /\s/.test(str[cursor])) cursor++;
        if (cursor >= str.length) {
          i = cursor;
          break;
        }
        if (str[cursor] === '"') {
          cursor++;
          let escaped = false;
          while (cursor < str.length) {
            const ch = str[cursor];
            if (escaped) {
              escaped = false;
              cursor++;
              continue;
            }
            if (ch === '\\') {
              escaped = true;
              cursor++;
              continue;
            }
            if (ch === '"') {
              cursor++;
              break;
            }
            cursor++;
          }
        } else {
          while (cursor < str.length && !/[,\n\r}]/.test(str[cursor])) cursor++;
        }
        while (cursor < str.length && /\s/.test(str[cursor])) cursor++;
        if (cursor < str.length && str[cursor] === ',') {
          cursor++;
          while (cursor < str.length && /\s/.test(str[cursor])) cursor++;
        }
        i = cursor;
      }
      return out;
    } catch (_) {
      return str;
    }
  }

  function stripUnclickedBtncs(text, result, eol) {
    try {
      if (!result) return text;
      const grp = locateMacroGroupBounds(text, result);
      if (!grp) return text;
      const keep = (result.insertClickedKeys instanceof Set) ? result.insertClickedKeys : new Set();
      let out = text;
      for (const e of (result.entries || [])) {
        if (!e || !e.sourceOp || !e.source) continue;
        // Only manage published ButtonControls and skip YouTubeButton template
        try {
          if (String(e.source) === 'YouTubeButton') continue;
          if (!isButtonControl(out, grp.groupOpenIndex, grp.groupCloseIndex, e.sourceOp, e.source)) continue;
          const key = `${e.sourceOp}.${e.source}`;
          if (keep.has(key)) continue; // preserve for explicit inserts
          // tool-level
          const tb = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, e.sourceOp);
          let uc = tb ? findUserControlsInTool(out, tb.open, tb.close) : null;
          if (!uc) uc = findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
          if (uc) {
            const cb = findControlBlockInUc(out, uc.open, uc.close, e.source);
            if (cb) out = removeBtncsExecuteInRange(out, cb.open, cb.close);
          }
          // group-level
          try {
            const ucg = findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
            if (ucg) {
              const cb2 = findControlBlockInUc(out, ucg.open, ucg.close, e.source);
              if (cb2) out = removeBtncsExecuteInRange(out, cb2.open, cb2.close);
            }
          } catch(_) {}
        } catch(_) {}
      }
      return out;
    } catch(_) { return text; }
  }

  function stripLegacyLauncherArtifacts(text) {
    try {
      const pattern = /",1\)\s*if\s+sep[\s\S]*?\n(\s*",)/g;
      return text.replace(pattern, '\n$1');
    } catch (_) {
      return text;
    }
  }


  // --- Button launcher (YouTube-style) minimal helpers ---
  // Built-in canonical YouTube launcher line (exact property line), used if no template exists in file
  const BUILTIN_YT_LINE = `BTNCS_Execute = "                                local url = \"https://www.youtube.com/@PatrickStirling\"\n                                local sep = package.config:sub(1,1)\n                                if sep == \"\\\\\" then\n                                    os.execute('start \"\" \"'..url..'\"')\n                                else\n                                    local uname = io.popen(\"uname\"):read(\"*l\")\n                                    if uname == \"Darwin\" then\n                                        os.execute('open \"'..url..'\"')\n                                    else\n                                        os.execute('xdg-open \"'..url..'\"')\n                                    end\n                                end\n                            ",`;
  function escapeSettingString(s) {
    try {
      // Escape backslashes first, then quotes, then state.newlines
      return String(s)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r?\n/g, "\\n");
    } catch(_) { return s; }
  }
  // Inverse of escapeSettingString for BTNCS_Execute bodies
  function unescapeSettingString(s) {
    try {
      return String(s)
        .replace(/\\n/g, '\n')
        .replace(/\\"/g, '"')
        .replace(/\\\\/g, '\\');
    } catch(_) { return s; }
  }
  // Build canonical YouTube launcher with explicit parts to ensure escapes
  function buildBuiltinYouTubeLine2() {
    try {
      const parts = [
        'local url = "https://www.youtube.com/@PatrickStirling"',
        'local sep = package.config:sub(1,1)',
        'if sep == "\\\\" then',
        '  os.execute(\'start "" "\'..url..\'"\')',
        'else',
        '  local uname = io.popen("uname"):read("*l")',
        '  if uname == "Darwin" then',
        '    os.execute(\'open "\'..url..\'"\')',
        '  else',
        '    os.execute(\'xdg-open "\'..url..\'"\')',
        '  end',
        'end'
      ];
      const lua = parts.join('\n') + '\n';
      const esc = escapeSettingString(lua);
      return 'BTNCS_Execute = "' + esc + '",';
    } catch(_) { return null; }
  }
  // Build a canonical YouTube launcher property line by escaping a raw Lua snippet
  function buildBuiltinYouTubeLine() {
    try {
      const lua = `local url = "https://www.youtube.com/@PatrickStirling"\nlocal sep = package.config:sub(1,1)\nif sep == "\\" then\n  os.execute('start "" "'..url..'"')\nelse\n  local uname = io.popen("uname"):read("*l")\n  if uname == "Darwin" then\n    os.execute('open "'..url..'"')\n  else\n    os.execute('xdg-open "'..url..'"')\n  end\nend\n`;
      const esc = escapeSettingString(lua);
      return 'BTNCS_Execute = "' + esc + '",';
    } catch(_) { return null; }
  }
  function extractBtncsExecuteString(text, ucOpen, ucClose, controlId) {
    try {
      // Find the control block first
      const cb = findControlBlockInUc(text, ucOpen, ucClose, controlId);
      if (!cb) return null;
      const slice = text.slice(cb.open + 1, cb.close);
      const key = 'BTNCS_Execute';
      const kpos = slice.indexOf(key);
      if (kpos < 0) return null;
      let i = kpos + key.length;
      // skip spaces and =
      while (i < slice.length && (/\s/.test(slice[i]) || slice[i] === '=')) i++;
      if (slice[i] !== '"') return null;
      // capture "..." with escapes intact
      const startQ = i;
      i++;
      let acc = 'BTNCS_Execute = "';
      let esc = false;
      while (i < slice.length) {
        const ch = slice[i++];
        acc += ch;
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === '"') break;
      }
      // optional trailing comma
      let j = i; while (j < slice.length && /\s/.test(slice[j])) j++;
      if (slice[j] === ',') acc += ',';
      return acc;
    } catch(_) { return null; }
  }
  function insertExactYouTubeLauncher(text, openBraceIndex, closeBraceIndex, eol, templateLine, pageName) {
    try {
      const start = openBraceIndex + 1;
      const end = closeBraceIndex;
      let body = stripBtncsEntries(text.slice(start, end));
      if (pageName) {
        const safePage = (String(pageName).trim()) || 'Controls';
        const replacement = `ICS_ControlPage = "${escapeQuotes(safePage)}",`;
        if (/ICS_ControlPage\s*=/.test(body)) {
          body = body.replace(/ICS_ControlPage\s*=\s*"(?:[^"\\]|\\.)*"\s*,?/g, replacement);
        } else {
          const propIndent = (getLineIndent(text, openBraceIndex) || '') + '\t';
          const newline = eol || '\n';
          body = `${newline}${propIndent}${replacement}${body}`;
        }
      }
      let indent = '\t';
      try { indent = (getLineIndent(text, openBraceIndex) || '') + '\t'; } catch(_) {}
      if (!templateLine) return text; // nothing to insert
      const line = indent + templateLine + (eol || '\n');
      return text.slice(0, start) + (eol || '\n') + line + body + text.slice(end);
    } catch(_) { return text; }
  }
  function isButtonControl(original, groupOpen, groupClose, toolName, controlId) {
    try {
      const tb = findToolBlockInGroup(original, groupOpen, groupClose, toolName);
      let uc = tb ? findUserControlsInTool(original, tb.open, tb.close) : null;
      if (!uc) uc = findUserControlsInGroup(original, groupOpen, groupClose);
      if (!uc) return false;
      const cb = findControlBlockInUc(original, uc.open, uc.close, controlId);
      if (!cb) return false;
      const slice = original.slice(cb.open, Math.min(cb.close + 1, original.length));
      return /INPID_InputControl\s*=\s*"ButtonControl"/.test(slice);
    } catch(_) { return false; }
  }
  function shouldOfferInsertForControl(original, groupOpen, groupClose, toolName, controlId) {
    try {
      if (!isButtonControl(original, groupOpen, groupClose, toolName, controlId)) return false;
      const key = (toolName && controlId) ? `${toolName}.${controlId}` : null;
      // If this control already has a URL tracked for it in this session,
      // keep the URL/launcher UI visible so the user can adjust it.
      try {
        if (key && state.parseResult) {
          if (state.parseResult.insertUrlMap instanceof Map && state.parseResult.insertUrlMap.has(key)) return true;
          if (state.parseResult.insertClickedKeys instanceof Set && state.parseResult.insertClickedKeys.has(key)) return true;
        }
      } catch(_) {}
      const tb = findToolBlockInGroup(original, groupOpen, groupClose, toolName);
      let uc = tb ? findUserControlsInTool(original, tb.open, tb.close) : findUserControlsInGroup(original, groupOpen, groupClose);
      if (!uc) return false;
      const cb = findControlBlockInUc(original, uc.open, uc.close, controlId);
      if (!cb) return false;
      const propLine = extractBtncsExecuteString(original, uc.open, uc.close, controlId);
      if (!propLine) return true;
      const firstQ = propLine.indexOf('"');
      const lastQ = propLine.lastIndexOf('"');
      const luaEsc = (firstQ >= 0 && lastQ > firstQ) ? propLine.slice(firstQ + 1, lastQ) : '';
      if (!luaEsc.length) return true; // explicit empty string
      // If this BTNCS_Execute looks like one of our generated launchers,
      // still offer the UI so the user can tweak the URL later.
      try {
        const lua = unescapeSettingString(luaEsc);
        if (lua && /local\s+url\s*=/.test(lua) && /local\s+sep\s*=\s*package\.config:sub\(1,1\)/.test(lua)) {
          return true;
        }
      } catch(_) {}
      // Otherwise, leave existing custom BTNCS_Execute alone and do not show insert UI.
      return false;
    } catch(_) { return false; }
  }
  function ensureExactLauncherInserted(resultText, result, eol) {
    try {
      // Only honor explicit UI clicks stored in insertClickedKeys
      const req = (result && result.insertClickedKeys instanceof Set) ? result.insertClickedKeys : null;
      if (!req || req.size === 0) return resultText;
      const grp = locateMacroGroupBounds(resultText, result);
      if (!grp) return resultText;
      let out = resultText;
      try { logDiag(`[Launcher] queue size: ${req.size}`); } catch(_) {}
      // Attempt to source template from an existing YouTubeButton control under same tool or any tool
      const entries = Array.isArray(result?.entries) ? result.entries : [];
      req.forEach((key) => {
        try {
          const dot = String(key).indexOf('.'); if (dot < 0) return;
          const tool = key.slice(0, dot); const ctrl = key.slice(dot + 1);
          const entry = entries.find(e => e && e.sourceOp === tool && e.source === ctrl) || null;
          const pageName = entry && entry.page ? entry.page : 'Controls';
          try { logDiag(`[Launcher] processing ${tool}.${ctrl}`); } catch(_) {}
          // find target control block
          const tbTarget = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, tool);
          let ucTarget = tbTarget ? findUserControlsInTool(out, tbTarget.open, tbTarget.close) : null;
          if (!ucTarget) ucTarget = findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
          if (!ucTarget) return;
          const cbTarget = findControlBlockInUc(out, ucTarget.open, ucTarget.close, ctrl);
          if (!cbTarget) return;
          // Prefer a user-provided URL from the insert UI
          let templateLine = null;
          let customUrl = null;
          try { if (result && result.insertUrlMap instanceof Map) customUrl = String(result.insertUrlMap.get(key) || '').trim(); } catch(_) {}
          if (customUrl) {
            templateLine = buildLauncherLineForUrl(customUrl);
            try { logDiag(`[Launcher] Using custom URL for ${tool}.${ctrl}: ${customUrl}`); } catch(_) {}
          }
          // Otherwise try template from YouTubeButton either under same tool or any
          // same tool first
          if (!templateLine && tbTarget) {
            templateLine = extractBtncsExecuteString(out, ucTarget.open, ucTarget.close, 'YouTubeButton');
          }
          // scan all tools if not found
          if (!templateLine) {
            // naive search: try the polygon tool commonly used
            const tbPoly = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, 'Polygon1');
            if (tbPoly) {
              const ucPoly = findUserControlsInTool(out, tbPoly.open, tbPoly.close);
              if (ucPoly) templateLine = extractBtncsExecuteString(out, ucPoly.open, ucPoly.close, 'YouTubeButton');
            }
          }
          if (!templateLine) {
            templateLine = buildBuiltinYouTubeLine2(); // fallback to built-in template
            try { logDiag(`[Launcher] Using built-in template for ${tool}.${ctrl}`); } catch(_) {}
          }
          // Insert into tool-level control
          out = insertExactYouTubeLauncher(out, cbTarget.open, cbTarget.close, eol, templateLine, pageName);
          try { logDiag(`[Launcher] Inserted for ${tool}.${ctrl} at tool-level`); } catch(_) {}
          // Also ensure a macro GroupOperator.UserControls control exists and has launcher
          try { out = ensureGroupControlHasLauncher(out, grp.groupOpenIndex, grp.groupCloseIndex, ctrl, templateLine, eol, pageName); } catch(_) {}
        } catch(_) {}
      });
      return out;
    } catch(_) { return resultText; }
  }

  // Build a launcher property line for a specific URL
  function buildLauncherLineForUrl(url) {
    try {
      const parts = [
        `local url = "${String(url)}"`,
        'local sep = package.config:sub(1,1)',
        'if sep == "\\\\" then',
        '  os.execute(\'start "" "\'..url..\'"\')',
        'else',
        '  local uname = io.popen("uname"):read("*l")',
        '  if uname == "Darwin" then',
        '    os.execute(\'open "\'..url..\'"\')',
        '  else',
        '    os.execute(\'xdg-open "\'..url..\'"\')',
        '  end',
        'end'
      ];
      const lua = parts.join('\n') + '\n';
      const esc = escapeSettingString(lua);
      return 'BTNCS_Execute = "' + esc + '",';
    } catch(_) { return buildBuiltinYouTubeLine2(); }
  }

  function ensureGroupControlHasLauncher(text, groupOpen, groupClose, ctrl, templateLine, eol, pageName) {
    try {
      let out = text;
      // Ensure UserControls block exists at group-level
      let ucGroup = findUserControlsInGroup(out, groupOpen, groupClose);
      if (!ucGroup) {
        const indent = (getLineIndent(out, groupOpen) || '') + '\t';
        const block = eol + indent + 'UserControls = ordered() {' + eol + indent + '},' + eol;
        out = out.slice(0, groupOpen + 1) + block + out.slice(groupOpen + 1);
        // Re-locate positions after insertion
        const grp2Close = findMatchingBrace(out, groupOpen);
        ucGroup = findUserControlsInGroup(out, groupOpen, grp2Close >= 0 ? grp2Close : groupClose) || null;
      }
      if (!ucGroup) return out;
      // If control exists, insert launcher there; else create minimal control with launcher
      let cb = findControlBlockInUc(out, ucGroup.open, ucGroup.close, ctrl);
      if (cb) {
        out = insertExactYouTubeLauncher(out, cb.open, cb.close, eol, templateLine, pageName);
        try { logDiag(`[Launcher] Inserted for Group.${ctrl} at group-level`); } catch(_) {}
        return out;
      }
      // Create minimal control
      try {
        const baseIndent = (getLineIndent(out, ucGroup.open) || '') + '\t';
        const safePage = (pageName && String(pageName).trim()) ? String(pageName).trim() : 'Controls';
        const line = baseIndent + ctrl + ' = { ' + templateLine + ' ICS_ControlPage = "' + escapeQuotes(safePage) + '", INPID_InputControl = "ButtonControl", LINKS_Name = "' + ctrl + '", },' + eol;
        out = out.slice(0, ucGroup.open + 1) + eol + line + out.slice(ucGroup.open + 1);
        try { logDiag(`[Launcher] Created group-level control ${ctrl} with launcher`); } catch(_) {}
        return out;
      } catch(_) { return out; }
    } catch(_) { return text; }
  }

  function applyMacroNameRename(text, result) {
    try {
      if (!result) return text;
      const originalName = (result.macroNameOriginal || '').trim();
      const newName = (result.macroName || '').trim();
      const targetName = newName || originalName;
      if (!originalName || !targetName) return text;
      const desiredType = (result.operatorType || result.operatorTypeOriginal || 'GroupOperator').trim() || 'GroupOperator';
      const nameEsc = originalName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp('(^|\\n)(\\s*)' + nameEsc + '(\\s*=\\s*)(GroupOperator|MacroOperator)(\\s*\\{)', 'm');
      return text.replace(re, (match, prefix, spaces, equalsPart, existingType, bracePart) => {
        const typeOut = desiredType || existingType;
        return `${prefix}${spaces}${targetName}${equalsPart}${typeOut}${bracePart}`;
      });
    } catch(_) { return text; }
  }


  // Replace per-control URL launches inside specific UserControls blocks
  function applyLaunchUrlOverrides(text, result, eol) {
    try {
      const overrides = (result && result.buttonOverrides && typeof result.buttonOverrides.forEach === 'function') ? result.buttonOverrides : null;
      if (!overrides || result.buttonOverrides.size === 0) return text;
      const grp = locateMacroGroupBounds(text, result);
      if (!grp) return text;
      let out = text;
      overrides.forEach((newUrl, key) => {
        try {
          const v = String(newUrl || '').trim();
          if (!v) return;
          const dot = String(key).indexOf('.');
          if (dot < 0) return;
          const tool = key.slice(0, dot);
          const ctrl = key.slice(dot + 1);
          let tb = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, tool);
          let uc = tb ? findUserControlsInTool(out, tb.open, tb.close)
                      : findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
          if (!uc) return;
          const cb = findControlBlockInUc(out, uc.open, uc.close, ctrl);
          if (!cb) return;
          const before = out;
          out = rewriteBtncsExecuteInRange(out, cb.open, cb.close, v, eol);
          if (out !== before) { try { logDiag(`[URL override] ${key} -> rewritten canonical`); } catch(_) {} }
        } catch(_) {}
      });
      return out;
    } catch(_) { return text; }
  }
  function findToolBlockInGroup(text, groupOpen, groupClose, toolName) {
    try {
      const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const groupText = text.slice(groupOpen, groupClose);
      const re = new RegExp('(^|\\n)\\s*' + esc(String(toolName)) + '\\s*=\\s*[A-Za-z_][A-Za-z0-9_]*\\s*\\{');
      const m = re.exec(groupText);
      if (!m) return null;
      const relOpen = m.index + m[0].lastIndexOf('{');
      const absOpen = groupOpen + relOpen;
      const absClose = findMatchingBrace(text, absOpen);
      if (absClose < 0 || absClose > groupClose) return null;
      return { open: absOpen, close: absClose };
    } catch(_) { return null; }
  }

  function findUserControlsInGroup(text, groupOpen, groupClose) {
    try {
      let i = groupOpen + 1;
      let depth = 1;
      let inStr = false;
      const toolsPos = text.indexOf('Tools = ordered()', groupOpen);
      const limit = (toolsPos >= 0 && toolsPos < groupClose) ? toolsPos : groupClose;
      while (i < limit) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && text[i - 1] !== '\\') inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; if (depth <= 0) break; continue; }
        if (depth === 1 && text.slice(i, i + 12) === 'UserControls') {
          let j = i + 12;
          while (j < limit && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < limit && isSpace(text[j])) j++;
          if (text.slice(j, j + 8) === 'ordered(') {
            j += 8;
            while (j < limit && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < limit && isSpace(text[j])) j++;
          }
          if (text[j] !== '{') { i++; continue; }
          const ucOpen = j;
          const ucClose = findMatchingBrace(text, ucOpen);
          if (ucClose < 0 || ucClose > groupClose) return null;
          return { open: ucOpen, close: ucClose };
        }
        i++;
      }
      return null;
    } catch(_) { return null; }
  }
  function findUserControlsInTool(text, toolOpen, toolClose) {
    try {
      if (toolOpen == null || toolClose == null || toolClose <= toolOpen) return null;
      const segment = text.slice(toolOpen, toolClose);
      const match = /UserControls\s*=\s*(?:ordered\(\))?\s*\{/m.exec(segment);
      if (!match) return null;
      const ucOpen = toolOpen + match.index + match[0].lastIndexOf('{');
      if (ucOpen < toolOpen || ucOpen > toolClose) return null;
      const ucClose = findMatchingBrace(text, ucOpen);
      if (ucClose < 0 || ucClose > toolClose) return null;
      return { open: ucOpen, close: ucClose };
    } catch(_) { return null; }
  }

  function findControlBlockInUc(text, ucOpen, ucClose, controlId) {
    try {
      let i = ucOpen + 1, depth = 0, inStr = false;
      while (i < ucClose) {
        const ch = text[i];
        if (inStr) { if (ch === '"' && text[i-1] !== '\\') inStr = false; i++; continue; }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; continue; }
        if (depth === 0 && (isIdentStart(ch) || ch === '[')) {
          let idStart = i;
          let idStr = '';
          if (ch === '[') {
            let j = i + 1;
            while (j < ucClose && text[j] !== ']') j++;
            idStr = text.slice(i, Math.min(j + 1, ucClose));
            i = Math.min(j + 1, ucClose);
          } else {
            i++;
            while (i < ucClose && isIdentPart(text[i])) i++;
            idStr = text.slice(idStart, i);
          }
          const norm = normalizeId(idStr);
          while (i < ucClose && isSpace(text[i])) i++;
          if (text[i] !== '=') { i++; continue; }
          i++;
          while (i < ucClose && isSpace(text[i])) i++;
          if (text[i] !== '{') { i++; continue; }
          const cOpen = i;
          const cClose = findMatchingBrace(text, cOpen);
          if (cClose < 0) break;
          if (String(norm) === String(controlId)) {
            return { open: cOpen, close: cClose };
          }
          i = cClose + 1; continue;
        }
        i++;
      }
      return null;
    } catch(_) { return null; }
  }

  function ensureToolUserControlsBlock(text, bounds, toolName, eol) {
    try {
      if (!toolName || !bounds) return { text, toolBlock: null, ucBlock: null };
      const newline = eol || '\n';
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return { text, toolBlock: null, ucBlock: null };
      const existing = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
      if (existing) return { text, toolBlock, ucBlock: existing };
      const indent = (getLineIndent(text, toolBlock.open) || '') + '	';
      const insertPos = toolBlock.close;
      let insertPrefix = '';
      let probe = insertPos - 1;
      while (probe >= toolBlock.open && /\s/.test(text[probe])) probe--;
      if (probe >= toolBlock.open && text[probe] !== '{' && text[probe] !== ',') {
        insertPrefix = ',';
      }
      const snippet = `${insertPrefix}${newline}${indent}UserControls = ordered() {${newline}${indent}},${newline}`;
      const updated = text.slice(0, insertPos) + snippet + text.slice(insertPos);
      const ucOpen = insertPos + snippet.indexOf('{');
      const ucClose = insertPos + snippet.lastIndexOf('}');
      const delta = snippet.length;
      return {
        text: updated,
        toolBlock: { open: toolBlock.open, close: toolBlock.close + delta },
        ucBlock: { open: ucOpen, close: ucClose },
        delta,
      };
    } catch (_) {
      return { text, toolBlock: null, ucBlock: null };
    }
  }

  function ensureControlBlockInUserControls(text, ucRange, controlName, eol) {
    try {
      if (!ucRange) return { text, block: null, uc: ucRange, delta: 0 };
      const existing = findControlBlockInUc(text, ucRange.open, ucRange.close, controlName);
      if (existing) return { text, block: existing, uc: ucRange, delta: 0 };
      const newline = eol || '\n';
      const indent = (getLineIndent(text, ucRange.open) || '') + '\t';
      const innerIndent = indent + '\t';
      const safeName = sanitizeIdent(controlName || 'Control');
      const displayName = escapeQuotes(controlName || safeName);
      let insertPrefix = '';
      let probe = ucRange.close - 1;
      while (probe >= ucRange.open && /\s/.test(text[probe])) probe--;
      if (probe >= ucRange.open && text[probe] !== '{' && text[probe] !== ',' ) {
        insertPrefix = ',';
      }
      const blockText = `${insertPrefix}${newline}${indent}${safeName} = {${newline}${innerIndent}LINKS_Name = "${displayName}",${newline}${indent}},${newline}`;
      const insertPos = ucRange.close;
      const updated = text.slice(0, insertPos) + blockText + text.slice(insertPos);
      const open = insertPos + blockText.indexOf('{');
      const close = insertPos + blockText.lastIndexOf('}');
      const delta = blockText.length;
      return {
        text: updated,
        block: { open, close },
        uc: { open: ucRange.open, close: ucRange.close + delta },
        delta,
      };
    } catch (_) {
      return { text, block: null, uc: ucRange, delta: 0 };
    }
  }

  function applyLabelCountEdits(text, result, eol) {

    try {

      const labels = (result.entries || []).filter(x => x && x.isLabel && Number.isFinite(x.labelCount));

      if (!labels.length) return text;

      const group = locateMacroGroupBounds(text, result);

      if (!group) return text;

      const toolsPos = text.indexOf('Tools = ordered()', group.groupOpenIndex);

      if (toolsPos < 0 || toolsPos > group.groupCloseIndex) return text;

      const tOpen = text.indexOf('{', toolsPos);

      if (tOpen < 0) return text;

      const tClose = findMatchingBrace(text, tOpen);

      if (tClose < 0) return text;

      let toolsInner = text.slice(tOpen + 1, tClose);

      let changed = false;

      for (const e of labels) {

        const ni = replaceLabelNumInputsInTools(toolsInner, e.sourceOp, e.source, e.labelCount, eol);

        if (ni !== toolsInner) { toolsInner = ni; changed = true; }

      }

      if (!changed) return text;

      return text.slice(0, tOpen + 1) + toolsInner + text.slice(tClose);

    } catch (_) {

      return text;

    }

  }



  function replaceLabelNumInputsInTools(toolsInner, toolName, controlName, newCount, eol) {

    let i = 0, depth = 0, inStr = false;

    while (i < toolsInner.length) {

      const ch = toolsInner[i];

      if (inStr) { if (ch === '"' && toolsInner[i - 1] !== '\\') inStr = false; i++; continue; }

      if (ch === '"') { inStr = true; i++; continue; }

      if (ch === '{') { depth++; i++; continue; }

      if (ch === '}') { depth--; i++; continue; }

      if (depth === 0 && isIdentStart(ch)) {

        const start = i; i++;

        while (i < toolsInner.length && isIdentPart(toolsInner[i])) i++;

        const name = toolsInner.slice(start, i);

        while (i < toolsInner.length && isSpace(toolsInner[i])) i++;

        if (toolsInner[i] !== '=') { i++; continue; }

        i++;

        while (i < toolsInner.length && isSpace(toolsInner[i])) i++;

        while (i < toolsInner.length && isIdentPart(toolsInner[i])) i++;

        while (i < toolsInner.length && isSpace(toolsInner[i])) i++;

        if (toolsInner[i] !== '{') { i++; continue; }

        const bOpen = i;

        const bClose = findMatchingBrace(toolsInner, bOpen);

        if (bClose < 0) break;

        if (name === toolName) {

          const body = toolsInner.slice(bOpen + 1, bClose);

          const newBody = replaceLabelNumInputsInToolBody(body, controlName, newCount, eol);

          if (newBody !== body) {

            return toolsInner.slice(0, bOpen + 1) + newBody + toolsInner.slice(bClose);

          }

        }

        i = bClose + 1;

        continue;

      }

      i++;

    }

    return toolsInner;

  }



  function replaceLabelNumInputsInToolBody(body, controlName, newCount, eol) {

    const ucPos = body.indexOf('UserControls = ordered()');

    if (ucPos < 0) return body;

    const ucOpen = body.indexOf('{', ucPos);

    if (ucOpen < 0) return body;

    const ucClose = findMatchingBrace(body, ucOpen);

    if (ucClose < 0) return body;

    const before = body.slice(0, ucOpen + 1);

    let uc = body.slice(ucOpen + 1, ucClose);

    const after = body.slice(ucClose);

    let i = 0, depth = 0, inStr = false;

    while (i < uc.length) {

      const ch = uc[i];

      if (inStr) { if (ch === '"' && uc[i - 1] !== '\\') inStr = false; i++; continue; }

      if (ch === '"') { inStr = true; i++; continue; }

      if (ch === '{') { depth++; i++; continue; }

      if (ch === '}') { depth--; i++; continue; }

      if (depth === 0 && isIdentStart(ch)) {

        const start = i; i++;

        while (i < uc.length && isIdentPart(uc[i])) i++;

        const name = uc.slice(start, i);

        while (i < uc.length && isSpace(uc[i])) i++;

        if (uc[i] !== '=') { i++; continue; }

        i++;

        while (i < uc.length && isSpace(uc[i])) i++;

        if (uc[i] !== '{') { i++; continue; }

        const cOpen = i;

        const cClose = findMatchingBrace(uc, cOpen);

        if (cClose < 0) break;

        if (name === controlName) {

          const cBody = uc.slice(cOpen + 1, cClose);

          if (!/INPID_InputControl\s*=\s*"LabelControl"/.test(cBody)) { i = cClose + 1; continue; }

          let newCBody;

          if (/LBLC_NumInputs\s*=\s*\d+/.test(cBody)) {

            newCBody = cBody.replace(/LBLC_NumInputs\s*=\s*\d+/, `LBLC_NumInputs = ${newCount}`);

          } else {

            const insert = `${eol}LBLC_NumInputs = ${newCount},`;

            newCBody = insert + cBody;

          }

          const newUc = uc.slice(0, cOpen + 1) + newCBody + uc.slice(cClose);

          return before + newUc + after;

        }

        i = cClose + 1;

        continue;

      }

      i++;

    }

    return body;

  }

  function applyLabelVisibilityEdits(text, result, eol) {
    try {
      if (!result || !Array.isArray(result.entries)) return text;
      const labels = result.entries.filter(e => e && e.isLabel && typeof e.labelHidden === 'boolean');
      if (!labels.length) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      let updated = text;
      const newline = eol || '\n';
      for (const entry of labels) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        updated = rewriteLabelVisibilityForControl(updated, bounds, entry.sourceOp, entry.source, !!entry.labelHidden, newline);
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function rewriteLabelVisibilityForControl(text, bounds, toolName, controlName, hidden, eol) {
    try {
      if (!toolName || !controlName) return text;
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return text;
      const uc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
      if (!uc) return text;
      const block = findControlBlockInUc(text, uc.open, uc.close, controlName);
      if (!block) return text;
      return rewriteLabelVisibilityBlock(text, block.open, block.close, hidden, eol);
    } catch (_) {
      return text;
    }
  }

  function rewriteLabelVisibilityBlock(text, openIndex, closeIndex, hidden, eol) {
    try {
      const indent = (getLineIndent(text, openIndex) || '') + '\t';
      let body = text.slice(openIndex + 1, closeIndex);
      if (hidden) {
        body = upsertControlProp(body, 'IC_Visible', 'false', indent, eol);
        body = upsertControlProp(body, 'INP_Passive', 'true', indent, eol);
      } else {
        body = removeControlProp(body, 'IC_Visible');
        body = removeControlProp(body, 'INP_Passive');
      }
      return text.slice(0, openIndex + 1) + body + text.slice(closeIndex);
    } catch (_) {
      return text;
    }
  }

  function applyLabelDefaultStateEdits(text, result, eol) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const newline = eol || '\n';
      let updated = text;
      for (const entry of result.entries) {
        if (!entry || !entry.isLabel || !entry.sourceOp || !entry.source) continue;
        const desired = normalizeLabelStateValue(entry);
        if (desired == null) continue;
        if (typeof entry.raw === 'string') {
          entry.raw = rewriteInstanceInputDefault(entry.raw, desired, newline);
        }
        updated = rewriteLabelInputValue(updated, bounds, entry.sourceOp, entry.source, desired, newline);
        updated = rewriteLabelUserControlDefault(updated, bounds, entry.sourceOp, entry.source, desired, newline);
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function normalizeLabelStateValue(entry) {
    try {
      let raw = entry?.controlMeta?.defaultValue;
      if (raw == null) raw = entry?.controlMetaOriginal?.defaultValue;
      if (raw == null && entry?.labelValueOriginal != null) raw = entry.labelValueOriginal;
      if (raw == null) return null;
      const normalized = String(raw).trim().toLowerCase();
      if (!normalized) return null;
      if (normalized === '0' || normalized === 'false' || normalized === 'closed') return '0';
      if (normalized === '1' || normalized === 'true' || normalized === 'open') return '1';
      const num = Number(normalized);
      if (Number.isFinite(num)) return num === 0 ? '0' : '1';
      return null;
    } catch (_) {
      return null;
    }
  }

  function rewriteLabelInputValue(text, bounds, toolName, controlName, value, eol) {
    try {
      if (!toolName || !controlName) return text;
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return text;
      const inputsBlock = findInputsBlockInTool(text, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return text;
      const inputBlock = findInputBlockInInputs(text, inputsBlock.open, inputsBlock.close, controlName);
      if (!inputBlock) return text;
      return rewriteInputBlockValue(text, inputBlock.open, inputBlock.close, value, eol);
    } catch (_) {
      return text;
    }
  }

  function rewriteLabelUserControlDefault(text, bounds, toolName, controlName, value, eol) {
    try {
      if (!toolName || !controlName) return text;
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return text;
      const uc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
      if (!uc) return text;
      const block = findControlBlockInUc(text, uc.open, uc.close, controlName);
      if (!block) return text;
      const indent = (getLineIndent(text, block.open) || '') + '	';
      let body = text.slice(block.open + 1, block.close);
      body = upsertControlProp(body, 'INP_Default', value, indent, eol || '\n');
      return text.slice(0, block.open + 1) + body + text.slice(block.close);
    } catch (_) {
      return text;
    }
  }

  function rewriteInstanceInputDefault(raw, value, eol) {
    try {
      if (!raw || value == null) return raw;
      const open = raw.indexOf('{');
      if (open < 0) return raw;
      const close = findMatchingBrace(raw, open);
      if (close < 0) return raw;
      const indent = (getLineIndent(raw, open) || '') + '	';
      let body = raw.slice(open + 1, close);
      body = setInstanceInputProp(body, 'Default', value, indent, eol || '\n');
      return raw.slice(0, open + 1) + body + raw.slice(close);
    } catch (_) {
      return raw;
    }
  }

  function findInputsBlockInTool(text, toolOpen, toolClose) {
    try {
      const slice = text.slice(toolOpen, toolClose);
      const re = /Inputs\s*=\s*\{/g;
      const match = re.exec(slice);
      if (!match) return null;
      const relOpen = match.index + match[0].lastIndexOf('{');
      const absOpen = toolOpen + relOpen;
      const absClose = findMatchingBrace(text, absOpen);
      if (absClose < 0 || absClose > toolClose) return null;
      return { open: absOpen, close: absClose };
    } catch (_) {
      return null;
    }
  }

  function findInputBlockInInputs(text, inputsOpen, inputsClose, controlName) {
    try {
      let i = inputsOpen + 1;
      let depth = 0;
      let inStr = false;
      while (i < inputsClose) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && text[i - 1] !== '\\') inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; continue; }
        if (depth === 0 && (isIdentStart(ch) || ch === '[')) {
          let idStart = i;
          let idStr = '';
          if (ch === '[') {
            let j = i + 1;
            while (j < inputsClose && text[j] !== ']') j++;
            idStr = text.slice(i, Math.min(j + 1, inputsClose));
            i = Math.min(j + 1, inputsClose);
          } else {
            i++;
            while (i < inputsClose && isIdentPart(text[i])) i++;
            idStr = text.slice(idStart, i);
          }
          const norm = normalizeId(idStr);
          while (i < inputsClose && isSpace(text[i])) i++;
          if (text[i] !== '=') { i++; continue; }
          i++;
          while (i < inputsClose && isSpace(text[i])) i++;
          if (text.slice(i, i + 5) === 'Input') {
            i += 5;
            while (i < inputsClose && isSpace(text[i])) i++;
          }
          if (text[i] !== '{') { i++; continue; }
          const blockOpen = i;
          const blockClose = findMatchingBrace(text, blockOpen);
          if (blockClose < 0) break;
          if (String(norm) === String(controlName)) {
            return { open: blockOpen, close: blockClose };
          }
          i = blockClose + 1;
          continue;
        }
        i++;
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function extractInputBlockValue(text, blockOpen, blockClose) {
    try {
      const body = text.slice(blockOpen + 1, blockClose);
      const range = findControlPropRange(body, 'Value');
      if (!range) return null;
      let snippet = body.slice(range.start, range.end);
      const eq = snippet.indexOf('=');
      if (eq < 0) return null;
      let val = snippet.slice(eq + 1).trim();
      if (val.endsWith(',')) val = val.slice(0, -1).trim();
      if (/^Number\s*\{/.test(val)) {
        const match = val.match(/Value\s*=\s*([-\d\.]+)/);
        if (match && match[1] != null) return match[1];
      }
      if (/^{\s*([-\d\.]+)\s*}$/.test(val)) {
        return RegExp.$1;
      }
      if (val.length) return val;
      return null;
    } catch (_) {
      return null;
    }
  }

  function rewriteInputBlockValue(text, blockOpen, blockClose, valueLiteral, eol) {
    try {
      if (valueLiteral == null) return text;
      const indent = (getLineIndent(text, blockOpen) || '') + '\t';
      let body = text.slice(blockOpen + 1, blockClose);
      body = removeControlProp(body, 'Value');
      body = upsertControlProp(body, 'Value', valueLiteral, indent, eol);
      return text.slice(0, blockOpen + 1) + body + text.slice(blockClose);
    } catch (_) {
      return text;
    }
  }

  function extractToolInputValue(text, bounds, toolName, controlName) {
    try {
      if (!toolName || !controlName) return null;
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return null;
      const inputsBlock = findInputsBlockInTool(text, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return null;
      const block = findInputBlockInInputs(text, inputsBlock.open, inputsBlock.close, controlName);
      if (!block) return null;
      return extractInputBlockValue(text, block.open, block.close);
    } catch (_) {
      return null;
    }
  }

  function normalizePageNameGlobal(name) {
    const raw = (name && String(name).trim()) ? String(name).trim() : 'Controls';
    return raw;
  }

  function getEntryKey(entry) {
    if (!entry) return '';
    return `${entry.sourceOp || ''}::${entry.source || ''}`;
  }

  function derivePageOrderFromEntries(entries, existingOrder) {
    const base = Array.isArray(existingOrder) ? existingOrder.map(normalizePageNameGlobal) : [];
    const seen = new Set(base);
    const order = [...base];
    const add = (name) => {
      const norm = normalizePageNameGlobal(name);
      if (!seen.has(norm)) {
        order.push(norm);
        seen.add(norm);
      }
    };
    if (Array.isArray(entries)) entries.forEach(entry => add(entry?.page));
    if (!seen.has('Controls')) order.unshift('Controls');
    return order.filter(Boolean);
  }

  function hydrateEntryPagesFromUserControls(text, result) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return;
      const bounds = locateMacroGroupBounds(text, result);
      const groupOpen = bounds?.groupOpenIndex ?? 0;
      const groupClose = bounds?.groupCloseIndex ?? text.length;
      const cache = new Map();
      const getUcForTool = (toolName) => {
        if (!toolName) return null;
        if (cache.has(toolName)) return cache.get(toolName);
        const tb = findToolBlockInGroup(text, groupOpen, groupClose, toolName);
        const uc = tb ? findUserControlsInTool(text, tb.open, tb.close) : null;
        cache.set(toolName, uc || null);
        return uc;
      };
      for (const entry of result.entries) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        const uc = getUcForTool(entry.sourceOp);
        if (!uc) continue;
        const block = findControlBlockInUc(text, uc.open, uc.close, entry.source);
        if (!block) continue;
        const body = text.slice(block.open + 1, block.close);
        const page = extractQuotedProp(body, 'ICS_ControlPage');
        if (page && (!entry.page || entry.page === 'Controls')) entry.page = page;
      }
    } catch (_) {}
  }

  function ensurePageIconsMap(result) {
    if (!result) return new Map();
    if (result.pageIcons instanceof Map) return result.pageIcons;
    const map = new Map();
    if (result.pageIcons && typeof result.pageIcons === 'object') {
      Object.entries(result.pageIcons).forEach(([key, value]) => {
        if (value) map.set(normalizePageNameGlobal(key), String(value));
      });
    }
    result.pageIcons = map;
    return map;
  }

  function hydrateControlPageIcons(text, result) {
    try {
      if (!text || !result) return;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) { result.pageIcons = new Map(); return; }
      const uc = findGroupUserControlsBlock(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      const map = new Map();
      if (!uc) { result.pageIcons = map; return; }
      let i = uc.openIndex + 1;
      while (i < uc.closeIndex) {
        let nameStart = i;
        while (nameStart < uc.closeIndex && /\s/.test(text[nameStart])) nameStart++;
        if (nameStart >= uc.closeIndex) break;
        if (!isIdentStart(text[nameStart])) { i = nameStart + 1; continue; }
        let nameEnd = nameStart + 1;
        while (nameEnd < uc.closeIndex && isIdentPart(text[nameEnd])) nameEnd++;
        const pageName = text.slice(nameStart, nameEnd);
        let cursor = nameEnd;
        while (cursor < uc.closeIndex && isSpace(text[cursor])) cursor++;
        if (text[cursor] !== '=') { i = cursor + 1; continue; }
        cursor++;
        while (cursor < uc.closeIndex && isSpace(text[cursor])) cursor++;
        if (text[cursor] === 'C' && text.slice(cursor, cursor + 'ControlPage'.length) === 'ControlPage') {
          cursor += 'ControlPage'.length;
          while (cursor < uc.closeIndex && isSpace(text[cursor])) cursor++;
          if (text[cursor] !== '{') { i = cursor + 1; continue; }
          const open = cursor;
          const close = findMatchingBrace(text, open);
          if (close < 0 || close > uc.closeIndex) break;
          const body = text.slice(open + 1, close);
          const iconId = extractQuotedProp(body, 'CTID_DIB_ID');
          if (iconId) map.set(normalizePageNameGlobal(pageName), iconId);
          i = close + 1;
        } else {
          i = cursor + 1;
        }
      }
      result.pageIcons = map;
    } catch (_) { /* ignore */ }
  }

  function extractInstancePropValue(raw, prop) {
    try {
      if (!raw) return null;
      const re = new RegExp(`${prop}\\s*=\\s*([^,\\n\\r}]*)`, 'i');
      const match = re.exec(raw);
      if (!match) return null;
      return match[1].trim().replace(/,$/, '');
    } catch (_) { return null; }
  }

  function extractControlPropValue(body, prop) {
    try {
      if (!body) return null;
      const re = new RegExp(`${prop}\\s*=\\s*([^,\\n\\r}]*)`, 'i');
      const match = re.exec(body);
      if (!match) return null;
      return match[1].trim().replace(/,$/, '');
    } catch (_) { return null; }
  }

  function hydrateControlMetadata(text, result) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return;
      const cache = new Map();
      const getUcForTool = (toolName) => {
        if (!toolName) return null;
        if (cache.has(toolName)) return cache.get(toolName);
        const tb = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
        const uc = tb ? findUserControlsInTool(text, tb.open, tb.close) : null;
        cache.set(toolName, uc || null);
        return uc;
      };
      for (const entry of result.entries) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        const meta = {};
        const uc = getUcForTool(entry.sourceOp);
        if (uc) {
          const block = findControlBlockInUc(text, uc.open, uc.close, entry.source);
          if (block) {
            const body = text.slice(block.open + 1, block.close);
            meta.inputControl = extractControlPropValue(body, 'INPID_InputControl');
            meta.dataType = extractControlPropValue(body, 'LINKID_DataType');
            meta.defaultValue = extractControlPropValue(body, 'INP_Default');
            meta.minScale = extractControlPropValue(body, 'INP_MinScale');
            meta.maxScale = extractControlPropValue(body, 'INP_MaxScale');
            const execLine = extractBtncsExecuteString(text, uc.open, uc.close, entry.source);
            if (execLine) {
              const first = execLine.indexOf('"');
              const last = execLine.lastIndexOf('"');
              if (last > first) {
                const luaEsc = execLine.slice(first + 1, last);
                entry.buttonExecute = unescapeSettingString(luaEsc);
              }
            }
          }
        }
        const instInputControl = extractInstancePropValue(entry.raw, 'INPID_InputControl');
        if (instInputControl != null && !meta.inputControl) meta.inputControl = instInputControl;
        const instDataType = extractInstancePropValue(entry.raw, 'LINKID_DataType');
        if (instDataType != null && !meta.dataType) meta.dataType = instDataType;
        const instDefault = extractInstancePropValue(entry.raw, 'Default');
        if (instDefault != null) meta.defaultValue = instDefault;
        if (entry.isLabel) {
          const actual = extractToolInputValue(text, bounds, entry.sourceOp, entry.source);
          if (actual != null) {
            meta.defaultValue = actual;
            entry.labelValueOriginal = actual;
          }
        }
        entry.controlMeta = { ...meta };
        entry.controlMetaOriginal = { ...meta };
        entry.controlMetaDirty = false;
        if (meta.inputControl && /buttoncontrol/i.test(meta.inputControl)) {
          entry.isButton = true;
        }
        if (typeof entry.buttonExecute !== 'string') {
          entry.buttonExecute = '';
        }
      }
    } catch (_) {}
  }

  function ensureBlendToggleSet(result) {
    if (!result) return new Set();
    if (result.blendToggles instanceof Set) return result.blendToggles;
    const set = new Set();
    if (result.blendToggles && typeof result.blendToggles.forEach === 'function') {
      result.blendToggles.forEach((v) => set.add(v));
    } else if (Array.isArray(result.blendToggles)) {
      result.blendToggles.forEach((v) => set.add(v));
    }
    result.blendToggles = set;
    return set;
  }

  function syncBlendToggleFlags(result) {
    try {
      if (!result || !Array.isArray(result.entries)) return;
      const set = ensureBlendToggleSet(result);
      for (const entry of result.entries) {
        if (!entry) continue;
        const key = getEntryKey(entry);
        if (!key) continue;
        if (entry.isBlendToggle) {
          set.add(key);
          continue;
        }
        entry.isBlendToggle = set.has(key);
      }
    } catch (_) {}
  }

  function getOrderedEntryIndices(result) {
    if (!result || !Array.isArray(result.entries)) return [];
    const entries = result.entries;
    const baseOrder = Array.isArray(result.order) && result.order.length
      ? result.order
      : entries.map((_, idx) => idx);
    const decorated = baseOrder.map((idx, pos) => {
      const entry = entries[idx];
      const sortIndex = entry && Number.isFinite(entry.sortIndex) ? entry.sortIndex : pos;
      return { idx, sortIndex };
    });
    decorated.sort((a, b) => {
      if (a.sortIndex !== b.sortIndex) return a.sortIndex - b.sortIndex;
      return a.idx - b.idx;
    });
    return decorated.map(item => item.idx);
  }

  function hydrateBlendToggleState(_text, result) {
    try {
      if (!result || !Array.isArray(result.entries)) return;
      const set = ensureBlendToggleSet(result);
      set.clear();
      for (const entry of result.entries) {
        if (!entry || !entry.source) continue;
        const isBlend = String(entry.source).toLowerCase() === 'blend';
        const fromMeta = normalizeInputControlValue(
          entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
        );
        const fromInstance = normalizeInputControlValue(
          extractInstancePropValue(entry.raw || '', 'INPID_InputControl')
        );
        const effective = fromMeta || fromInstance;
        const isToggle = isBlend && !!effective && effective.toLowerCase() === 'checkboxcontrol';
        entry.isBlendToggle = isToggle;
        const key = getEntryKey(entry);
        if (isToggle && key) set.add(key);
      }
      result.blendToggles = set;
      window.__FMR_BLEND_SET = set;
    } catch (_) {
      if (result) result.blendToggles = new Set();
    }
  }

  function hydrateOnChangeScripts(text, result) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return;
      const cache = new Map();
      const getBody = (tool, ctrl) => getCachedControlBody(cache, text, bounds, tool, ctrl);
      for (const entry of result.entries) {
        if (!entry || !entry.sourceOp || !entry.source) {
          if (entry && typeof entry.onChange !== 'string') entry.onChange = '';
          continue;
        }
        const body = getBody(entry.sourceOp, entry.source);
        if (!body) {
          entry.onChange = entry.onChange || '';
          continue;
        }
        const match = body.match(/INPS_ExecuteOnChange\s*=\s*"((?:[^"\\]|\\.)*)"\s*,?/);
        entry.onChange = match && match[1] != null ? unescapeSettingString(match[1]) : (entry.onChange || '');
      }
    } catch (_) {
      if (result && Array.isArray(result.entries)) {
        result.entries.forEach(entry => {
          if (entry && typeof entry.onChange !== 'string') entry.onChange = '';
        });
      }
    }
  }

  function hydrateLabelVisibility(text, result) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return;
      const cache = new Map();
      const getBody = (tool, ctrl) => getCachedControlBody(cache, text, bounds, tool, ctrl);
      for (const entry of result.entries) {
        if (!entry || !entry.isLabel || !entry.sourceOp || !entry.source) continue;
        const body = getBody(entry.sourceOp, entry.source);
        if (!body) {
          entry.labelHidden = false;
          continue;
        }
        const hidden = /IC_Visible\s*=\s*false/.test(body) || /INP_Passive\s*=\s*true/.test(body);
        entry.labelHidden = !!hidden;
      }
    } catch (_) {
      if (result && Array.isArray(result.entries)) {
        result.entries.forEach(entry => {
          if (entry && entry.isLabel) entry.labelHidden = !!entry.labelHidden;
        });
      }
    }
  }

  function syncBlendToggleFlags(result) {
    try {
      if (!result || !Array.isArray(result.entries)) return;
      const set = ensureBlendToggleSet(result);
      for (const entry of result.entries) {
        if (!entry) continue;
        const key = getEntryKey(entry);
        if (!key) continue;
        if (entry.isBlendToggle) {
          set.add(key);
          continue;
        }
        entry.isBlendToggle = set.has(key);
      }
    } catch (_) {}
  }

function applyUserControlPages(text, result, eol) {
    try {
      if (!result || !Array.isArray(result.entries) || !result.entries.length) return text;
      let bounds = locateMacroGroupBounds(text, result);
      const perTool = new Map();
      const groupControls = new Map();
      const pageDefinitions = new Map();
      const pageIcons = ensurePageIconsMap(result);
      const blendTargets = new Map();
      const onChangeTargets = new Map();
      const buttonExecuteTargets = new Map();
      const metaOverrideEntries = [];
      const orderedIndices = getOrderedEntryIndices(result);
      const perToolOrder = new Map();
      const iterableEntries = orderedIndices
        .map(idx => result.entries[idx])
        .filter(entry => entry && entry.sourceOp && entry.source);
      for (const entry of iterableEntries) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        if (!perToolOrder.has(entry.sourceOp)) perToolOrder.set(entry.sourceOp, []);
        const arr = perToolOrder.get(entry.sourceOp);
        const normName = normalizeId(entry.source);
        if (!arr.some(n => normalizeId(n) === normName)) arr.push(entry.source);
        const pageName = (entry.page && String(entry.page).trim()) ? String(entry.page).trim() : 'Controls';
        if (!perTool.has(entry.sourceOp)) perTool.set(entry.sourceOp, new Map());
        perTool.get(entry.sourceOp).set(entry.source, pageName);
        groupControls.set(entry.source, pageName);
        if (!pageDefinitions.has(pageName)) {
          const isCommon = pageName === 'Common';
          pageDefinitions.set(pageName, {
            visible: !isCommon,
            priority: 1,
          });
        }
        const cfg = pageDefinitions.get(pageName);
        const iconId = pageIcons.get(normalizePageNameGlobal(pageName));
        if (iconId) cfg.iconId = iconId;
        const isBlend = String(entry.source).toLowerCase() === 'blend';
        const isToggle = entry.isBlendToggle || ensureBlendToggleSet(result).has(getEntryKey(entry));
        if (isBlend && isToggle) {
          entry.isBlendToggle = true;
          blendTargets.set(entry.sourceOp, (blendTargets.get(entry.sourceOp) || new Set()).add(entry.source));
        }
        const script = (entry.onChange && String(entry.onChange).trim()) ? String(entry.onChange).trim() : '';
        if (script) {
          if (!onChangeTargets.has(entry.sourceOp)) onChangeTargets.set(entry.sourceOp, []);
          onChangeTargets.get(entry.sourceOp).push({ control: entry.source, script });
        }
        if (entry.isButton) {
          const execScript = (entry.buttonExecute && String(entry.buttonExecute).trim()) ? String(entry.buttonExecute).trim() : '';
          if (execScript) {
            if (!buttonExecuteTargets.has(entry.sourceOp)) buttonExecuteTargets.set(entry.sourceOp, []);
            buttonExecuteTargets.get(entry.sourceOp).push({ control: entry.source, script: execScript });
          }
        }
        const metaDiff = entry.pendingMetaDiff || getControlMetaDiff(entry);
        if (entry.pendingMetaDiff) delete entry.pendingMetaDiff;
        if (metaDiff) {
          metaOverrideEntries.push({ entry, diff: metaDiff });
        } else {
          const fallbackInput = normalizeInputControlValue(
            entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
          );
          if (fallbackInput) {
            metaOverrideEntries.push({ entry, diff: { inputControl: fallbackInput } });
          }
        }
      }
      let out = text;
      for (const [toolName, controls] of perTool.entries()) {
        if (!toolName || !controls.size) continue;
        out = rewriteToolUserControls(out, toolName, controls, bounds, eol, result);
      }
      out = rewriteGroupUserControls(out, groupControls, bounds, eol);
      const orderedPages = derivePageOrderFromEntries(result.entries, result.pageOrder);
      result.pageOrder = orderedPages;
      const priorityMap = new Map();
      const pageCount = orderedPages.length;
      orderedPages.forEach((name, idx) => {
        const inverted = Math.max(0, pageCount - idx - 1);
        priorityMap.set(name, inverted);
      });
      let fallbackPriority = orderedPages.length;
      pageDefinitions.forEach((cfg, name) => {
        const pr = priorityMap.has(name) ? priorityMap.get(name) : fallbackPriority++;
        cfg.priority = pr;
      });
      out = ensureControlPagesDeclared(out, bounds, pageDefinitions, eol, result);
      const nameRes = applyDisplayNameOverrides(out, result, bounds, eol, result);
      out = nameRes.text;
      bounds = nameRes.bounds || bounds;
      let metaBounds = bounds;
      for (const item of metaOverrideEntries) {
        const res = applyControlMetaOverride(out, item.entry, item.diff, metaBounds, eol, result);
        out = res.text;
        metaBounds = res.bounds || metaBounds;
        item.entry.controlMetaOriginal = { ...(item.entry.controlMeta || {}) };
        item.entry.controlMetaDirty = false;
      }
      bounds = metaBounds || bounds;
      out = reorderToolUserControlsBlocks(out, bounds, perToolOrder, eol, result);
      blendTargets.forEach((controls, toolName) => {
        out = applyBlendCheckboxesToTool(out, bounds, toolName, Array.from(controls), eol);
      });
      onChangeTargets.forEach((items, toolName) => {
        out = applyOnChangeToTool(out, bounds, toolName, items, eol);
      });
      buttonExecuteTargets.forEach((items, toolName) => {
        out = applyButtonExecuteToTool(out, bounds, toolName, items, eol);
      });
      out = pruneRedundantUserControlEntries(out, result);
      out = stripEmptyToolUserControls(out, result);
      out = stripDanglingRootUserControls(out);
      out = stripDanglingRootInputs(out);
      return out;
    } catch (_) {
      return text;
    }
  }

  function rewriteToolUserControls(text, toolName, controlMap, bounds, eol, resultRef) {
    try {
      if (!controlMap || !controlMap.size) return text;
      const newline = eol || '\n';
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(updated, resultRef);
      if (!currentBounds) return updated;
      let ensureRes = ensureToolUserControlsBlock(updated, currentBounds, toolName, newline);
      updated = ensureRes.text;
      currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      let toolBlock = ensureRes.toolBlock || findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
      if (!toolBlock) return updated;
      let uc = ensureRes.ucBlock || findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
      if (!uc) return updated;
      for (const [controlName, pageName] of controlMap.entries()) {
        const ensuredBlock = ensureControlBlockInUserControls(updated, uc, controlName, newline);
        updated = ensuredBlock.text;
        uc = ensuredBlock.uc || uc;
        let block = ensuredBlock.block;
        if (!block && uc) block = findControlBlockInUc(updated, uc.open, uc.close, controlName);
        if (!block) continue;
        updated = rewriteControlBlockPage(updated, block.open, block.close, pageName, newline);
        currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
        toolBlock = findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
        if (!toolBlock) break;
        uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close) || uc;
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function rewriteGroupUserControls(text, controlMap, bounds, eol) {
    try {
      if (!controlMap || !controlMap.size) return text;
      const groupOpen = bounds?.groupOpenIndex ?? 0;
      const groupClose = bounds?.groupCloseIndex ?? text.length;
      const uc = findGroupUserControlsBlock(text, groupOpen, groupClose);
      if (!uc) return text;
      const segments = [];
      controlMap.forEach((pageName, controlName) => {
        const block = findControlBlockInUc(text, uc.openIndex, uc.closeIndex, controlName);
        if (block) segments.push({ open: block.open, close: block.close, pageName });
      });
      if (!segments.length) return text;
      segments.sort((a, b) => b.open - a.open);
      let updated = text;
      for (const seg of segments) {
        updated = rewriteControlBlockPage(updated, seg.open, seg.close, seg.pageName, eol);
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function ensureControlPagesDeclared(text, bounds, pageConfigs, eol, resultRef) {
    try {
      if (!pageConfigs || !pageConfigs.size) return text;
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(updated, resultRef);
      for (const [pageName, config] of pageConfigs.entries()) {
        if (!currentBounds) currentBounds = locateMacroGroupBounds(updated, resultRef);
        if (!currentBounds) break;
        const res = upsertControlPageDefinition(updated, currentBounds, pageName, config, eol || '\n', resultRef);
        updated = res.text;
        currentBounds = res.bounds;
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function reorderToolUserControlsBlocks(text, bounds, orderMap, eol, resultRef) {
    try {
      if (!orderMap || !orderMap.size) return text;
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(updated, resultRef);
      const newline = eol || '\n';
      orderMap.forEach((controls, toolName) => {
        if (!controls || !controls.length) return;
        if (!currentBounds) currentBounds = locateMacroGroupBounds(updated, resultRef);
        if (!currentBounds) return;
        const ensured = ensureToolUserControlsBlock(updated, currentBounds, toolName, newline);
        updated = ensured.text;
        currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
        const tb = ensured.toolBlock || findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
        if (!tb) return;
        const uc = ensured.ucBlock || findUserControlsInTool(updated, tb.open, tb.close);
        if (!uc) return;
        const segments = collectUserControlSegments(updated, uc.open, uc.close);
        if (!segments.length) return;
        const map = new Map();
        segments.forEach(seg => map.set(seg.name, seg));
        const used = new Set();
        const finalOrder = [];
        controls.forEach(name => {
          const seg = map.get(normalizeId(name));
          if (seg) {
            finalOrder.push(seg);
            used.add(seg);
            map.delete(normalizeId(name));
          }
        });
        segments.forEach(seg => {
          if (!used.has(seg)) finalOrder.push(seg);
        });
        const chunks = finalOrder.map((seg, idx) => {
          const needsComma = idx < finalOrder.length - 1;
          return ensureSegmentComma(updated.slice(seg.start, seg.end), needsComma);
        });
        const inner = chunks.join('');
        updated = updated.slice(0, uc.open + 1) + inner + updated.slice(uc.close);
        const delta = inner.length - (uc.close - (uc.open + 1));
        uc.close += delta;
        currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function ensureSegmentComma(segmentText, needsComma) {
    try {
      if (!needsComma) return segmentText;
      const trailingMatch = segmentText.match(/(\s*)$/) || [''];
      const trailing = trailingMatch[0] || '';
      const body = segmentText.slice(0, segmentText.length - trailing.length);
      if (/,(\s*)$/.test(body)) return segmentText;
      return body + ',' + trailing;
    } catch (_) {
      return segmentText;
    }
  }

  function collectUserControlSegments(text, ucOpen, ucClose) {
    try {
      const segments = [];
      let cursor = ucOpen + 1;
      while (cursor < ucClose) {
        const leadingStart = cursor;
        let segmentStart = cursor;
        while (segmentStart < ucClose && /\s/.test(text[segmentStart])) segmentStart++;
        if (segmentStart >= ucClose) break;
        let nameStart = segmentStart;
        let nameEnd = nameStart;
        if (text[nameStart] === '[') {
          nameEnd = text.indexOf(']', nameStart);
          if (nameEnd < 0 || nameEnd >= ucClose) break;
          nameEnd++;
        } else {
          while (nameEnd < ucClose && isIdentPart(text[nameEnd])) nameEnd++;
        }
        const rawName = text.slice(nameStart, nameEnd);
        let pos = nameEnd;
        while (pos < ucClose && isSpace(text[pos])) pos++;
        if (text[pos] !== '=') { cursor = pos + 1; continue; }
        pos++;
        while (pos < ucClose && isSpace(text[pos])) pos++;
        if (text[pos] !== '{') { cursor = pos + 1; continue; }
        const blockOpen = pos;
        const blockClose = findMatchingBrace(text, blockOpen);
        if (blockClose < 0 || blockClose > ucClose) break;
        let end = blockClose + 1;
        if (text[end] === ',') end++;
        while (end < ucClose && /\s/.test(text[end])) {
          if (text[end] === '\n' || text[end] === '\r') { end++; break; }
          end++;
        }
        segments.push({ name: normalizeId(rawName), start: leadingStart, end });
        cursor = end;
      }
      return segments;
    } catch (_) {
      return [];
    }
  }

  function isUserControlBodyRedundant(body) {
    try {
      if (!body) return true;
      let stripped = body.replace(/LINKS_Name\s*=\s*"[^"]*"\s*,?/gi, '');
      stripped = stripped.replace(/ICS_ControlPage\s*=\s*"[^"]*"\s*,?/gi, '');
      stripped = stripped.replace(/^\s+|\s+$/g, '');
      return stripped.length === 0;
    } catch (_) {
      return false;
    }
  }

  function pruneRedundantUserControlEntries(text, resultRef) {
    try {
      if (!text || !resultRef) return text;
      let updated = text;
      let bounds = locateMacroGroupBounds(updated, resultRef);
      if (!bounds) return updated;
      const toolNames = new Set();
      if (Array.isArray(resultRef.entries)) {
        resultRef.entries.forEach((entry) => {
          if (entry && entry.sourceOp) toolNames.add(entry.sourceOp);
        });
      }
      toolNames.forEach((toolName) => {
        bounds = locateMacroGroupBounds(updated, resultRef) || bounds;
        const toolBlock = findToolBlockInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
        if (!toolBlock) return;
        const uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
        if (!uc) return;
        const segments = collectUserControlSegments(updated, uc.open, uc.close);
        if (!segments.length) return;
        const removable = [];
        segments.forEach((seg) => {
          const block = findControlBlockInUc(updated, uc.open, uc.close, seg.name);
          if (!block) return;
          const body = updated.slice(block.open + 1, block.close);
          if (isUserControlBodyRedundant(body)) removable.push(seg);
        });
        if (!removable.length) return;
        removable.sort((a, b) => b.start - a.start);
        removable.forEach((seg) => {
          updated = updated.slice(0, seg.start) + updated.slice(seg.end);
        });
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function stripEmptyToolUserControls(text, resultRef) {
    try {
      if (!text) return text;
      const bounds = locateMacroGroupBounds(text, resultRef);
      if (!bounds) return text;
      const toolsPos = text.indexOf('Tools = ordered()', bounds.groupOpenIndex);
      if (toolsPos < 0 || toolsPos > bounds.groupCloseIndex) return text;
      const toolsOpen = text.indexOf('{', toolsPos);
      if (toolsOpen < 0) return text;
      const toolsClose = findMatchingBrace(text, toolsOpen);
      if (toolsClose < 0) return text;
      const inner = text.slice(toolsOpen + 1, toolsClose);
      let i = 0;
      let depth = 0;
      let inStr = false;
      const removals = [];
      while (i < inner.length) {
        const ch = inner[i];
        if (inStr) {
          if (ch === '"' && !isQuoteEscaped(inner, i)) inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; continue; }
        if (depth == 0 && isIdentStart(ch)) {
          const nameStart = i; i++;
          while (i < inner.length && isIdentPart(inner[i])) i++;
          while (i < inner.length && isSpace(inner[i])) i++;
          if (inner[i] != '=') { i++; continue; }
          i++;
          while (i < inner.length && isSpace(inner[i])) i++;
          while (i < inner.length && isIdentPart(inner[i])) i++;
          while (i < inner.length && isSpace(inner[i])) i++;
          if (inner[i] != '{') { i++; continue; }
          const toolOpen = toolsOpen + 1 + i;
          const toolClose = findMatchingBrace(text, toolOpen);
          if (toolClose < 0 || toolClose > toolsClose) break;
          const ucPos = text.indexOf('UserControls = ordered()', toolOpen);
          if (ucPos >= 0 && ucPos < toolClose) {
            const ucOpen = text.indexOf('{', ucPos);
            if (ucOpen >= 0 && ucOpen < toolClose) {
              const ucClose = findMatchingBrace(text, ucOpen);
              if (ucClose > ucOpen && ucClose < toolClose) {
                const body = text.slice(ucOpen + 1, ucClose);
                if (body.trim().length === 0) {
                  let end = ucClose + 1;
                  if (text[end] === ',') end++;
                  while (end < text.length && /\s/.test(text[end])) {
                    const c = text[end];
                    end++;
                    if (c === '\\n' || c === '\\r') break;
                  }
                  removals.push({ start: ucPos, end });
                }
              }
            }
          }
          i = (toolClose - (toolsOpen + 1)) + 1;
          continue;
        }
        i++;
      }
      if (!removals.length) return text;
      removals.sort((a, b) => b.start - a.start);
      let updated = text;
      removals.forEach(seg => {
        updated = updated.slice(0, seg.start) + updated.slice(seg.end);
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function stripDanglingRootUserControls(text) {
    try {
      if (!text) return text;
      const toolsPos = text.indexOf('Tools = ordered()');
      if (toolsPos < 0) return text;
      const ucPos = text.indexOf('UserControls = ordered()');
      if (ucPos < 0 || ucPos > toolsPos) return text;
      const open = text.indexOf('{', ucPos);
      if (open < 0 || open > toolsPos) return text;
      const close = findMatchingBrace(text, open);
      if (close < 0) return text;
      let end = close + 1;
      if (text[end] === ',') end++;
      while (end < text.length && /\s/.test(text[end])) {
        const ch = text[end];
        end++;
        if (ch === '\n' || ch === '\r') break;
      }
      return text.slice(0, ucPos) + text.slice(end);
    } catch (_) {
      return text;
    }
  }

  function stripDanglingRootInputs(text) {
    try {
      if (!text) return text;
      const toolsPos = text.indexOf('Tools = ordered()');
      if (toolsPos < 0) return text;
      const idx = text.indexOf('Inputs = ordered()');
      if (idx < 0 || idx > toolsPos) return text;
      const braceOpen = text.indexOf('{', idx);
      if (braceOpen < 0) return text;
      const braceClose = findMatchingBrace(text, braceOpen);
      if (braceClose < 0) return text;
      let end = braceClose + 1;
      if (text[end] === ',') end++;
      while (end < text.length && /\s/.test(text[end])) {
        const ch = text[end];
        end++;
        if (ch === '\n' || ch === '\r') break;
      }
      return text.slice(0, idx) + text.slice(end);
    } catch (_) {
      return text;
    }
  }

  function removeAllGroupInputsBlocks(text, groupOpen, groupClose, resultRef) {
    try {
      const blocks = findGroupInputsBlocks(text, groupOpen, groupClose) || [];
      if (!blocks.length) return text;
      let updated = text;
      const sorted = [...blocks].sort((a, b) => b.inputsHeaderStart - a.inputsHeaderStart);
      for (const blk of sorted) {
        let start = blk.inputsHeaderStart;
        while (start > groupOpen && /\s/.test(updated[start - 1])) {
          const prev = updated[start - 1];
          start--;
          if (prev === '\n' || prev === '\r') break;
        }
        let end = blk.closeIndex + 1;
        while (end < updated.length && /\s/.test(updated[end])) {
          const ch = updated[end];
          end++;
          if (ch === '\n' || ch === '\r') break;
        }
        if (updated[end] === ',') {
          end++;
          while (end < updated.length && /\s/.test(updated[end])) {
            const ch = updated[end];
            end++;
            if (ch === '\n' || ch === '\r') break;
          }
        }
        updated = updated.slice(0, start) + updated.slice(end);
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function applyDisplayNameOverrides(text, result, bounds, eol, resultRef) {
    try {
      if (!result || !Array.isArray(result.entries)) return { text, bounds };
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      for (const entry of result.entries) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        const desired = entry.displayName || entry.displayNameOriginal;
        const original = entry.displayNameOriginal || desired;
        if (!desired || desired === original) {
          entry.displayNameDirty = false;
          continue;
        }
        const res = rewriteToolControlDisplayName(updated, currentBounds, entry.sourceOp, entry.source, desired, eol || '\n', resultRef);
        updated = res.text;
        currentBounds = res.bounds || currentBounds;
        entry.displayNameOriginal = desired;
        entry.displayNameDirty = false;
      }
      return { text: updated, bounds: currentBounds };
    } catch (_) {
      return { text, bounds };
    }
  }

  function rewriteToolControlDisplayName(text, bounds, toolName, controlName, label, eol, resultRef) {
    try {
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      if (!currentBounds) return { text, bounds };
      const tb = findToolBlockInGroup(text, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
      if (!tb) return { text, bounds: currentBounds };
      let uc = findUserControlsInTool(text, tb.open, tb.close);
      if (!uc) {
        const ensured = ensureToolUserControlsBlock(text, currentBounds, toolName, eol);
        text = ensured.text;
        currentBounds = locateMacroGroupBounds(text, resultRef) || currentBounds;
        uc = ensured.ucBlock || (tb ? findUserControlsInTool(text, tb.open, tb.close) : null);
        if (!uc) return { text, bounds: currentBounds };
      }
      let ensuredBlock = ensureControlBlockInUserControls(text, uc, controlName, eol);
      text = ensuredBlock.text;
      if (ensuredBlock.uc) uc = ensuredBlock.uc;
      const block = ensuredBlock.block || findControlBlockInUc(text, uc.open, uc.close, controlName);
      if (!block) return { text, bounds: currentBounds };
      const indent = (getLineIndent(text, block.open) || '') + '\t';
      let body = text.slice(block.open + 1, block.close);
      const newline = eol || '\n';
      if (label) body = upsertControlProp(body, 'LINKS_Name', `"${escapeQuotes(label)}"`, indent, newline);
      else body = removeControlProp(body, 'LINKS_Name');
      const updated = text.slice(0, block.open + 1) + body + text.slice(block.close);
      const newBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      return { text: updated, bounds: newBounds };
    } catch (_) {
      return { text, bounds };
    }
  }

  function getControlMetaDiff(entry) {
    try {
      if (!entry) return null;
      const current = entry.controlMeta || {};
      const original = entry.controlMetaOriginal || {};
      const keys = ['defaultValue', 'minScale', 'maxScale', 'inputControl'];
      const diff = {};
      keys.forEach((key) => {
        const normalize = (val) => {
          if (key === 'inputControl') return normalizeInputControlValue(val);
          return val != null && String(val).trim() !== '' ? String(val).trim() : null;
        };
        const cur = normalize(current[key]);
        const orig = normalize(original[key]);
        if (cur !== orig) diff[key] = cur;
      });
      return Object.keys(diff).length ? diff : null;
    } catch (_) {
      return null;
    }
  }

  function applyControlMetaOverride(text, entry, diff, bounds, eol, resultRef) {
    try {
      if (!entry || !entry.sourceOp || !entry.source || !diff) return { text, bounds };
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      if (!currentBounds) return { text, bounds };
      let toolBlock = findToolBlockInGroup(text, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) return { text, bounds };
      let uc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
      if (!uc) {
        const ensured = ensureToolUserControlsBlock(text, currentBounds, entry.sourceOp, eol || '\n');
        text = ensured.text;
        currentBounds = locateMacroGroupBounds(text, resultRef) || currentBounds;
        toolBlock = ensured.toolBlock || findToolBlockInGroup(text, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, entry.sourceOp);
        uc = ensured.ucBlock || (toolBlock ? findUserControlsInTool(text, toolBlock.open, toolBlock.close) : null);
        if (!uc) return { text, bounds: currentBounds };
      }
      let ensuredBlock = ensureControlBlockInUserControls(text, uc, entry.source, eol);
      text = ensuredBlock.text;
      if (ensuredBlock.uc) uc = ensuredBlock.uc;
      const block = ensuredBlock.block || findControlBlockInUc(text, uc.open, uc.close, entry.source);
      if (!block) return { text, bounds: currentBounds };
      const indent = (getLineIndent(text, block.open) || '') + '\t';
      let body = text.slice(block.open + 1, block.close);
      if (Object.prototype.hasOwnProperty.call(diff, 'defaultValue')) {
        const val = diff.defaultValue;
        if (val == null) body = removeControlProp(body, 'INP_Default');
        else body = upsertControlProp(body, 'INP_Default', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'minScale')) {
        const val = diff.minScale;
        if (val == null) body = removeControlProp(body, 'INP_MinScale');
        else body = upsertControlProp(body, 'INP_MinScale', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'maxScale')) {
        const val = diff.maxScale;
        if (val == null) body = removeControlProp(body, 'INP_MaxScale');
        else body = upsertControlProp(body, 'INP_MaxScale', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'inputControl')) {
        const val = diff.inputControl;
        if (val == null) body = removeControlProp(body, 'INPID_InputControl');
        else body = upsertControlProp(body, 'INPID_InputControl', `"${escapeQuotes(val)}"`, indent, eol);
      }
      const normalizeDefault = (val) => {
        if (val == null) return null;
        let out = String(val).trim();
        if (!out) return null;
        if (out.startsWith('"') && out.endsWith('"')) out = out.slice(1, -1);
        if (/^true$/i.test(out)) return '1';
        if (/^false$/i.test(out)) return '0';
        return out;
      };
      const hasControlProp = (prop) => {
        try { return new RegExp(`${prop}\\s*=`, 'i').test(body); } catch (_) { return false; }
      };
      const inputControl = Object.prototype.hasOwnProperty.call(diff, 'inputControl')
        ? diff.inputControl
        : extractControlPropValue(body, 'INPID_InputControl');
      const normalizedInput = normalizeInputControlValue(inputControl);
      if (normalizedInput && normalizedInput.toLowerCase() === 'checkboxcontrol') {
        const fallbackDefault = normalizeDefault(
          Object.prototype.hasOwnProperty.call(diff, 'defaultValue')
            ? diff.defaultValue
            : (entry?.controlMeta?.defaultValue ?? entry?.controlMetaOriginal?.defaultValue ?? extractInstancePropValue(entry.raw, 'Default') ?? '0')
        ) || '0';
        body = upsertControlProp(body, 'INPID_InputControl', '"CheckboxControl"', indent, eol);
        if (!hasControlProp('INP_Integer')) body = upsertControlProp(body, 'INP_Integer', 'false', indent, eol);
        if (!hasControlProp('INP_Default')) body = upsertControlProp(body, 'INP_Default', fallbackDefault, indent, eol);
        if (!hasControlProp('INP_MinScale')) body = upsertControlProp(body, 'INP_MinScale', '0', indent, eol);
        if (!hasControlProp('INP_MaxScale')) body = upsertControlProp(body, 'INP_MaxScale', '1', indent, eol);
        if (!hasControlProp('INP_MinAllowed')) body = upsertControlProp(body, 'INP_MinAllowed', '0', indent, eol);
        if (!hasControlProp('INP_MaxAllowed')) body = upsertControlProp(body, 'INP_MaxAllowed', '1', indent, eol);
        if (!hasControlProp('CBC_TriState')) body = upsertControlProp(body, 'CBC_TriState', 'false', indent, eol);
        if (!hasControlProp('LINKID_DataType')) body = upsertControlProp(body, 'LINKID_DataType', '"Number"', indent, eol);
      }
      const updated = text.slice(0, block.open + 1) + body + text.slice(block.close);
      const newBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      return { text: updated, bounds: newBounds };
    } catch (_) {
      return { text, bounds };
    }
  }

  function upsertControlPageDefinition(text, bounds, pageName, config, eol, resultRef) {
    try {
      if (!pageName) return { text, bounds };
      const ensured = ensureGroupUserControlsBlockExists(text, bounds, eol || '\n', resultRef);
      let updatedText = ensured.text;
      let currentBounds = ensured.bounds;
      const uc = ensured.block;
      if (!uc) return { text: updatedText, bounds: currentBounds };
      const ucBody = updatedText.slice(uc.openIndex + 1, uc.closeIndex);
      const escName = escapeRegex(pageName);
      const re = new RegExp(`(^|\\n)([\\t ]*)${escName}\\s*=\\s*ControlPage\\s*\\{`, 'm');
      const match = re.exec(ucBody);
      const indentBase = (getLineIndent(updatedText, uc.openIndex) || '') + '\t';
      const block = buildControlPageBlock(pageName, config, indentBase, eol || '\n');
      if (match) {
        const relLineStart = match.index + match[1].length;
        const lineStart = uc.openIndex + 1 + relLineStart;
        const braceOpen = updatedText.indexOf('{', lineStart);
        if (braceOpen < 0) return { text: updatedText, bounds: currentBounds };
        const braceClose = findMatchingBrace(updatedText, braceOpen);
        if (braceClose < 0) return { text: updatedText, bounds: currentBounds };
        let end = braceClose + 1;
        if (updatedText[end] === ',') end++;
        while (end < updatedText.length && /\s/.test(updatedText[end]) && updatedText[end] !== '\n' && updatedText[end] !== '\r') end++;
        updatedText = updatedText.slice(0, lineStart) + block + updatedText.slice(end);
      } else {
        const insertion = eol + block;
        updatedText = updatedText.slice(0, uc.closeIndex) + insertion + updatedText.slice(uc.closeIndex);
      }
      const newBounds = locateMacroGroupBounds(updatedText, resultRef) || currentBounds;
      return { text: updatedText, bounds: newBounds };
    } catch (_) {
      return { text, bounds };
    }
  }

  function ensureGroupUserControlsBlockExists(text, bounds, eol, resultRef) {
    try {
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      if (!currentBounds) return { text, bounds: null, block: null };
      let uc = findGroupUserControlsBlock(text, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex);
      if (uc) return { text, bounds: currentBounds, block: uc };
      const indent = (getLineIndent(text, currentBounds.groupOpenIndex) || '') + '\t';
      const snippet = `${eol}${indent}UserControls = ordered() {${eol}${indent}},${eol}`;
      const insertPos = currentBounds.groupOpenIndex + 1;
      const updated = text.slice(0, insertPos) + snippet + text.slice(insertPos);
      const newBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      const newUc = findGroupUserControlsBlock(updated, newBounds.groupOpenIndex, newBounds.groupCloseIndex);
      return { text: updated, bounds: newBounds, block: newUc };
    } catch (_) {
      return { text, bounds, block: null };
    }
  }

  function buildControlPageBlock(name, config, indentBase, eol) {
    const innerIndent = indentBase + '\t';
    const parts = [];
    parts.push(`${indentBase}${name} = ControlPage {`);
    const visible = config && config.visible === false ? 'false' : 'true';
    parts.push(`${innerIndent}CT_Visible = ${visible},`);
    if (config && config.iconId) {
      parts.push(`${innerIndent}CTID_DIB_ID = "${escapeQuotes(config.iconId)}",`);
    }
    const priority = (config && Number.isFinite(config.priority)) ? config.priority : 1;
    parts.push(`${innerIndent}CT_Priority = ${priority},`);
    parts.push(`${indentBase}},`);
    return parts.join(eol);
  }

  function rewriteControlBlockPage(text, openIndex, closeIndex, pageName, eol) {
    try {
      const safePage = (pageName && String(pageName).trim()) ? String(pageName).trim() : 'Controls';
      const body = text.slice(openIndex + 1, closeIndex);
      const replacement = `ICS_ControlPage = "${escapeQuotes(safePage)}",`;
      let newBody;
      if (/ICS_ControlPage\s*=/.test(body)) {
        newBody = body.replace(/ICS_ControlPage\s*=\s*"([^"]*)"\s*,?/g, replacement);
      } else {
        const indent = (getLineIndent(text, openIndex) || '');
        const propIndent = indent + '\t';
        const insert = `${eol}${propIndent}${replacement}`;
        newBody = insert + body;
      }
      return text.slice(0, openIndex + 1) + newBody + text.slice(closeIndex);
    } catch (_) {
      return text;
    }
  }

  function applyBlendToggleOverride(entry, raw, eol) {
    try {
      return raw;
    } catch (_) { return raw; }
  }

  function applyInstanceInputMetaOverride(raw, diff, eol) {
    try {
      if (!diff) return raw;
      const open = raw.indexOf('{');
      if (open < 0) return raw;
      const close = findMatchingBrace(raw, open);
      if (close < 0) return raw;
      const indent = (getLineIndent(raw, open) || '') + '\t';
      let body = raw.slice(open + 1, close);
      const applyProp = (prop, val, wrapQuotes = false) => {
        if (val == null || val === '') {
          body = removeInstanceInputProp(body, prop);
        } else {
          const formatted = wrapQuotes ? `"${escapeQuotes(val)}"` : val;
          body = setInstanceInputProp(body, prop, formatted, indent, eol);
        }
      };
      if (Object.prototype.hasOwnProperty.call(diff, 'defaultValue')) {
        applyProp('INP_Default', diff.defaultValue, false);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'minScale')) {
        applyProp('INP_MinScale', diff.minScale, false);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'maxScale')) {
        applyProp('INP_MaxScale', diff.maxScale, false);
      }
      return raw.slice(0, open + 1) + body + raw.slice(close);
    } catch (_) {
      return raw;
    }
  }

function applyBlendCheckboxesToTool(text, bounds, toolName, controls, eol, resultRef) {
    try {
      if (!toolName || !controls || !controls.length) return text;
      let scope = bounds;
      let toolBlock = scope ? findToolBlockInGroup(text, scope.groupOpenIndex, scope.groupCloseIndex, toolName) : null;
      if (!toolBlock && resultRef) {
        scope = locateMacroGroupBounds(text, resultRef);
        toolBlock = scope ? findToolBlockInGroup(text, scope.groupOpenIndex, scope.groupCloseIndex, toolName) : null;
      }
      if (!toolBlock) {
        logBlendDebug({ sourceOp: toolName, source: 'Blend' }, 'tool-missing', { toolName });
        return text;
      }
      let updated = text;
      let uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
      if (!uc) {
        const ensured = ensureToolUserControlsBlock(updated, scope || bounds, toolName, eol);
        updated = ensured.text;
        if (ensured.toolBlock) toolBlock = ensured.toolBlock;
        uc = ensured.ucBlock;
        if (ensured.delta) {
          if (scope && typeof scope.groupCloseIndex === 'number') scope.groupCloseIndex += ensured.delta;
          if (bounds && typeof bounds.groupCloseIndex === 'number') bounds.groupCloseIndex += ensured.delta;
        }
        if (uc) logBlendDebug({ sourceOp: toolName, source: 'Blend' }, 'tool-uc-created', { toolName });
      }
      if (!uc) {
        logBlendDebug({ sourceOp: toolName, source: 'Blend' }, 'tool-uc-missing', { toolName });
        return updated;
      }
      const unique = Array.from(new Set(controls));
      const blocks = [];
      unique.forEach(name => {
        let block = findControlBlockInUc(updated, uc.open, uc.close, name);
        if (!block) {
          const ensuredBlock = ensureControlBlockInUserControls(updated, uc, name, eol);
          updated = ensuredBlock.text;
          block = ensuredBlock.block;
          if (ensuredBlock.uc) uc = ensuredBlock.uc;
          if (ensuredBlock.delta) {
            toolBlock = { open: toolBlock.open, close: toolBlock.close + ensuredBlock.delta };
            if (scope && typeof scope.groupCloseIndex === 'number') scope.groupCloseIndex += ensuredBlock.delta;
            if (bounds && typeof bounds.groupCloseIndex === 'number') bounds.groupCloseIndex += ensuredBlock.delta;
          }
        }
        if (block) blocks.push(block);
      });
      blocks.sort((a, b) => b.open - a.open);
      blocks.forEach(block => {
        updated = rewriteControlBlockAsCheckbox(updated, block.open, block.close, eol);
      });
      logBlendDebug({ sourceOp: toolName, source: 'Blend' }, 'tool-controls-rewritten', { controls: unique });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function applyOnChangeToTool(text, bounds, toolName, items, eol) {
    try {
      if (!toolName || !items || !items.length) return text;
      let updated = text;
      const groupOpen = bounds?.groupOpenIndex ?? 0;
      const groupClose = bounds?.groupCloseIndex ?? text.length;
      items.forEach(({ control, script }) => {
        if (!control) return;
        const toolBlock = findToolBlockInGroup(updated, groupOpen, groupClose, toolName);
        if (!toolBlock) return;
        const uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
        if (!uc) return;
        const block = findControlBlockInUc(updated, uc.open, uc.close, control);
        if (!block) return;
        updated = rewriteControlBlockOnChange(updated, block.open, block.close, script, eol);
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function applyButtonExecuteToTool(text, bounds, toolName, items, eol) {
    try {
      if (!toolName || !items || !items.length) return text;
      let updated = text;
      const groupOpen = bounds?.groupOpenIndex ?? 0;
      const groupClose = bounds?.groupCloseIndex ?? text.length;
      items.forEach(({ control, script }) => {
        if (!control) return;
        const toolBlock = findToolBlockInGroup(updated, groupOpen, groupClose, toolName);
        if (!toolBlock) return;
        const uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
        if (!uc) return;
        const block = findControlBlockInUc(updated, uc.open, uc.close, control);
        if (!block) return;
        updated = rewriteControlBlockButtonExecute(updated, block.open, block.close, script, eol);
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function getCachedControlBody(cache, text, bounds, toolName, controlName) {
    if (!cache || !text || !bounds || !toolName || !controlName) return null;
    const key = `${toolName}::${controlName}`;
    if (cache.has(key)) return cache.get(key);
    const body = findControlBody(text, bounds, toolName, controlName);
    cache.set(key, body);
    return body;
  }

  function findControlBody(text, bounds, toolName, controlName) {
    try {
      if (!text || !bounds || !toolName || !controlName) return null;
      const block = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (block) {
        const uc = findUserControlsInTool(text, block.open, block.close);
        if (uc) {
          const cb = findControlBlockInUc(text, uc.open, uc.close, controlName);
          if (cb) return text.slice(cb.open + 1, cb.close);
        }
      }
      const groupUc = findUserControlsInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (groupUc) {
        const cb = findControlBlockInUc(text, groupUc.open, groupUc.close, controlName);
        if (cb) return text.slice(cb.open + 1, cb.close);
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function rewriteControlBlockAsCheckbox(text, openIndex, closeIndex, eol) {
    try {
      const indent = (getLineIndent(text, openIndex) || '') + '\t';
      let body = text.slice(openIndex + 1, closeIndex);
      const props = [
        ['INPID_InputControl', '"CheckboxControl"'],
        ['INP_Integer', 'false'],
        ['INP_Default', '1'],
        ['INP_MinScale', '0'],
        ['INP_MaxScale', '1'],
        ['INP_MinAllowed', '0'],
        ['INP_MaxAllowed', '1'],
        ['CBC_TriState', 'false'],
        ['LINKID_DataType', '"Number"'],
      ];
      props.forEach(([prop, value]) => {
        body = upsertControlProp(body, prop, value, indent, eol);
      });
      return text.slice(0, openIndex + 1) + body + text.slice(closeIndex);
    } catch (_) {
      return text;
    }
  }

  function upsertControlProp(body, prop, value, indent, eol) {
    try {
      body = removeControlProp(body, prop);
      if (body.trim().length) {
        const trimmed = body.replace(/\s+$/g, '');
        if (trimmed && !trimmed.endsWith(',')) {
          body = trimmed + ',' + body.slice(trimmed.length);
        }
      }
      const needsBreak = body.trim().length && !body.endsWith(eol);
      return body + (needsBreak ? eol : '') + indent + `${prop} = ${value},`;
    } catch (_) {
      return body;
    }
  }

  function removeControlProp(body, prop) {
    try {
      const range = findControlPropRange(body, prop);
      if (!range) return body;
      return body.slice(0, range.start) + body.slice(range.end);
    } catch (_) {
      return body;
    }
  }

  function findControlPropRange(body, prop) {
    try {
      if (!body) return null;
      const re = new RegExp(`(^|\\r?\\n)(\\s*)${prop}\\s*=`, 'g');
      const match = re.exec(body);
      if (!match) return null;
      let cursor = match.index + match[0].length;
      // cursor currently sits right after '='
      while (cursor < body.length && /\s/.test(body[cursor])) cursor++;
      if (cursor >= body.length) return null;
      const valueStart = cursor;
      if (body[cursor] === '"') {
        cursor++;
        let escaped = false;
        while (cursor < body.length) {
          const ch = body[cursor];
          if (escaped) {
            escaped = false;
          } else if (ch === '\\') {
            escaped = true;
          } else if (ch === '"') {
            cursor++;
            break;
          }
          cursor++;
        }
      } else if (body[cursor] === '{') {
        const close = findMatchingBrace(body, cursor);
        cursor = close >= 0 ? close + 1 : body.length;
      } else {
        while (cursor < body.length && !/[,\r\n}]/.test(body[cursor])) cursor++;
      }
      while (cursor < body.length && /\s/.test(body[cursor])) cursor++;
      if (cursor < body.length && body[cursor] === ',') cursor++;
      while (cursor < body.length && /[ \t]/.test(body[cursor])) cursor++;
      const start = match.index;
      return { start, end: cursor };
    } catch (_) {
      return null;
    }
  }

  function rewriteControlBlockOnChange(text, openIndex, closeIndex, script, eol) {
    try {
      const indent = (getLineIndent(text, openIndex) || '') + '\t';
      let body = text.slice(openIndex + 1, closeIndex);
      if (script && script.trim()) {
        const escaped = `"${escapeSettingString(script)}"`;
        body = upsertControlProp(body, 'INPS_ExecuteOnChange', escaped, indent, eol);
      } else {
        body = removeControlProp(body, 'INPS_ExecuteOnChange');
      }
      return text.slice(0, openIndex + 1) + body + text.slice(closeIndex);
    } catch (_) {
      return text;
    }
  }

  function rewriteControlBlockButtonExecute(text, openIndex, closeIndex, script, eol) {
    try {
      const indent = (getLineIndent(text, openIndex) || '') + '\t';
      let body = text.slice(openIndex + 1, closeIndex);
      if (script && script.trim()) {
        const escaped = `"${escapeSettingString(script)}"`;
        body = upsertControlProp(body, 'BTNCS_Execute', escaped, indent, eol);
      } else {
        body = removeControlProp(body, 'BTNCS_Execute');
      }
      return text.slice(0, openIndex + 1) + body + text.slice(closeIndex);
    } catch (_) {
      return text;
    }
  }

  function setInstanceInputProp(body, prop, value, indent, eol) {
    try {
      const pattern = new RegExp(`(^|\\r?\\n)\\s*${prop}\\s*=\\s*([^,\\n\\r}]*)\\s*,?`, 'm');
      if (value == null) return body;
      if (pattern.test(body)) {
        return body.replace(pattern, `$1${indent}${prop} = ${value},`);
      }
      const prefix = body.trim().length ? (body.endsWith(eol) ? '' : eol) : '';
      return body + prefix + indent + `${prop} = ${value},`;
    } catch (_) { return body; }
  }

  function removeInstanceInputProp(body, prop) {
    try {
      const pattern = new RegExp(`(^|\\r?\\n)\\s*${prop}\\s*=\\s*([^,\\n\\r}]*)\\s*,?`, 'g');
      return body.replace(pattern, '$1');
    } catch (_) { return body; }
  }

  function locateMacroGroupBounds(text, result) {
    try {
      if (!text) return null;
      const names = [];
      if (result?.macroName) names.push(result.macroName);
      if (result?.macroNameOriginal && result.macroNameOriginal !== result.macroName) {
        names.push(result.macroNameOriginal);
      }
      if (!names.length) names.push('Unknown');
      const operatorTypes = [];
      if (result?.operatorType) operatorTypes.push(result.operatorType);
      if (result?.operatorTypeOriginal && result.operatorTypeOriginal !== result.operatorType) {
        operatorTypes.push(result.operatorTypeOriginal);
      }
      if (!operatorTypes.length) operatorTypes.push('GroupOperator', 'MacroOperator');
      for (const name of names) {
        for (const opType of operatorTypes) {
          const bounds = findGroupByName(text, name, opType);
          if (bounds) return bounds;
        }
      }
      if (typeof result?.inputs?.openIndex === 'number') {
        const fallback = findEnclosingGroupForIndex(text, Math.max(0, Math.min(text.length - 1, result.inputs.openIndex)));
        if (fallback) return { groupOpenIndex: fallback.groupOpenIndex, groupCloseIndex: fallback.groupCloseIndex };
      }
      return { groupOpenIndex: 0, groupCloseIndex: text.length };
    } catch (_) {
      return { groupOpenIndex: 0, groupCloseIndex: text.length };
    }
  }

  function findGroupByName(text, macroName, operatorType) {
    try {
      if (!macroName || !operatorType) return null;
      const re = new RegExp(`(^|\\n|\\r)\\s*${escapeRegex(macroName)}\\s*=\\s*${escapeRegex(operatorType)}\\s*\\{`);
      const match = re.exec(text);
      if (!match) return null;
      const openIndex = match.index + match[0].lastIndexOf('{');
      const closeIndex = findMatchingBrace(text, openIndex);
      if (closeIndex < 0) return null;
      return { groupOpenIndex: openIndex, groupCloseIndex: closeIndex };
    } catch (_) {
      return null;
    }
  }

  function escapeRegex(str) {
    return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }



  function applyNameIfEdited(entry, eol) {

    if (!entry || !entry.name) return entry.raw;

    const raw = entry.raw;

    const open = raw.indexOf('{');

    if (open < 0) return raw;

    const close = findMatchingBrace(raw, open);

    if (close < 0) return raw;

    const head = raw.slice(0, open + 1);

    const body = raw.slice(open + 1, close);

    const tail = raw.slice(close);

    const nameRe = /(^|[\s,])Name\s*=\s*"([^"]*)"/;

    if (nameRe.test(body)) {

      const newBody = body.replace(/Name\s*=\s*"([^"]*)"/, `Name = "${escapeQuotes(entry.name)}"`);

      return head + newBody + tail;

    } else {

      // Insert at start of body with a comma

      const insert = `${eol}Name = "${escapeQuotes(entry.name)}",`;

      return head + insert + body + tail;

    }

  }

  function shouldLogBlendDebug() {
    try {
      if (state && state.debugBlendLogging) return true;
      if (typeof window !== 'undefined' && window.__FMR_DEBUG_BLEND) return true;
    } catch (_) {}
    return false;
  }

  function logBlendDebug(entry, stage, extra = {}) {
    if (!shouldLogBlendDebug()) return;
    try {
      const payload = {
        stage,
        op: entry?.sourceOp || null,
        control: entry?.source || null,
        key: entry ? getEntryKey(entry) : null,
        isToggle: !!(entry && entry.isBlendToggle),
        page: entry?.page || null,
        ...extra,
      };
      const msg = `blend-debug ${JSON.stringify(payload)}`;
      if (typeof logDiag === 'function') logDiag(msg);
      else console.log(msg);
    } catch (_) {}
  }

  function applyEntryOverrides(entry, raw, eol) {
    const newline = eol || '\n';
    const isBlendControl = !!(entry && entry.source && String(entry.source).toLowerCase() === 'blend');
    const blendSet = state?.parseResult ? ensureBlendToggleSet(state.parseResult) : null;
    const keyed = entry ? getEntryKey(entry) : '';
    const inSet = !!(blendSet && keyed ? blendSet.has(keyed) : false);
    const isToggle = !!(entry && entry.isBlendToggle) || inSet;
    if (isBlendControl) logBlendDebug(entry, 'pre-export', { inSet, rawHasCheckbox: /INPID_InputControl\s*=\s*"CheckboxControl"/.test(raw || '') });
    if (isBlendControl && isToggle && entry) {
      entry.isBlendToggle = true;
    }
    let out = raw;
    const metaDiff = getControlMetaDiff(entry);
    out = applyInstanceInputMetaOverride(out, metaDiff, newline);
    if (metaDiff) {
      entry.pendingMetaDiff = metaDiff;
    } else if (entry && entry.pendingMetaDiff) {
      delete entry.pendingMetaDiff;
    }
    out = applyBlendToggleOverride(entry, out, newline);
    try {
      const inputControlChanged = !!(
        (metaDiff && Object.prototype.hasOwnProperty.call(metaDiff, 'inputControl')) || entry?.controlTypeEdited
      );
      if (inputControlChanged) {
        const open = out.indexOf('{');
        const close = out.lastIndexOf('}');
        if (open >= 0 && close > open) {
          let body = out.slice(open + 1, close);
          const props = [
            'INPID_InputControl',
            'INP_Integer',
            'INP_Default',
            'INP_MinScale',
            'INP_MaxScale',
            'INP_MinAllowed',
            'INP_MaxAllowed',
            'CBC_TriState',
            'LINKID_DataType',
          ];
          props.forEach(prop => {
            body = removeInstanceInputProp(body, prop);
          });
          out = out.slice(0, open + 1) + body + out.slice(close);
        }
      }
    } catch (_) {}
    if (isBlendControl) logBlendDebug(entry, 'emit-default', { isToggle, inSet });
    return out;
  }

  function isKeyLikelyValid(key) {
    if (!key) return false;
    const trimmed = String(key).trim();
    if (!trimmed) return false;
    return /[A-Za-z0-9_]/.test(trimmed);
  }

  function deriveFallbackEntryKey(entry) {
    if (!entry) return '';
    if (isKeyLikelyValid(entry.key)) return entry.key;
    const base = (entry.sourceOp && entry.source)
      ? `${entry.sourceOp}_${entry.source}`
      : (entry.source || entry.name || entry.displayName || 'Input');
    return makeUniqueKey(base);
  }

  function ensureInstanceInputKey(raw, entry) {
    if (!raw) return raw;
    const match = raw.match(/^(\s*)([^\s=]+)\s*=\s*InstanceInput\b/);
    if (!match) return raw;
    const currentKey = match[2] || '';
    if (isKeyLikelyValid(currentKey)) return raw;
    const fallback = deriveFallbackEntryKey(entry);
    if (!fallback) return raw;
    return raw.replace(/^(\s*)([^\s=]+)\s*=\s*InstanceInput\b/, `$1${fallback} = InstanceInput`);
  }



  function reindent(raw, indent, eol) {

    const lines = raw.replace(/\r\n/g, '\n').split('\n');

    while (lines.length && lines[0].trim() === '') lines.shift();

    while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();

    const indents = lines.filter(l => l.trim() !== '').map(l => (l.match(/^[\t ]*/) || [''])[0].length);

    const minIndent = indents.length ? Math.min(...indents) : 0;

    const trimmed = lines.map(l => l.slice(minIndent));

    return trimmed.map(l => indent + l).join(eol);

  }



  function getLabelGroupForIndex(idx) {

    const order = state.parseResult?.order || [];

    const entries = state.parseResult?.entries || [];

    const pos = order.indexOf(idx);

    if (pos < 0) return [idx];

    { const ent = entries[idx]; const cnt = (ent && ent.isLabel && Number.isFinite(ent.labelCount)) ? ent.labelCount : 0; return [idx, ...order.slice(pos + 1, pos + 1 + cnt)]; }

  }





  // Color grouping helpers for Published panel

  function deriveColorBaseFromId(id) {

    const suf = ["Red","Green","Blue","Alpha"].find(s => String(id || "").endsWith(s));

    return suf ? String(id).slice(0, String(id).length - suf.length) : null;

  }



  function getOrAssignControlGroup(sourceOp, base) {

    if (!state.parseResult) return null;

    if (!state.parseResult.cgMap) state.parseResult.cgMap = new Map();

    if (!Number.isFinite(state.parseResult.cgNext)) state.parseResult.cgNext = 1;

    const key = `${sourceOp}::${base}`;

    if (state.parseResult.cgMap.has(key)) return state.parseResult.cgMap.get(key);

    const id = state.parseResult.cgNext++;

    state.parseResult.cgMap.set(key, id);

    return id;

  }



  // Remove entries and rebuild order/selection






  // Compute insertion position under lowest selected, and helper to insert indices

  function insertIndicesAt(order, idxs, pos) {

    const set = new Set(idxs);

    const filtered = order.filter(v => !set.has(v));

    const clamped = Math.max(0, Math.min(filtered.length, pos));

    filtered.splice(clamped, 0, ...idxs);

    return filtered;

  }



  // Return true if the catalog most likely contains the given type (by key or toolType/type fields)

  function hasTypeInCatalog(cat, type) {

    try {

      if (!cat || !type) return false;

      const t = String(type).trim().toLowerCase();

      if (Object.prototype.hasOwnProperty.call(cat, type)) return true;

      for (const k of Object.keys(cat)) {

        if (String(k).toLowerCase() === t) return true;

        const v = cat[k];

        const vt = String((v && (v.toolType || v.type)) || '').toLowerCase();

        if (vt && vt === t) return true;

      }

      return false;

    } catch (_) { return false; }

  }





  function normalizeId(id) {

    // Strip brackets and quotes around ["..."] keys

    let s = String(id).trim();

    if (s.startsWith('["') && s.endsWith('"]')) s = s.slice(2, -2);

    if (s.startsWith('[') && s.endsWith(']')) s = s.slice(1, -1);

    if (s.startsWith('"') && s.endsWith('"')) s = s.slice(1, -1);

    return s;

  }







  function applyNodeControlMeta(entry, meta) {
    try {
      if (!entry || !meta) return;
      const kind = String(meta.kind || '').toLowerCase();
      if (kind === 'label') {
        entry.isLabel = true;
        entry.labelCount = Number.isFinite(meta.labelCount) ? Number(meta.labelCount) : 0;
      } else if (kind === 'button') {
        entry.isButton = true;
      }
      if (meta.inputControl) {
        const val = String(meta.inputControl);
        const quoted = /"/.test(val) ? val : `"${val}"`;
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        entry.controlMeta.inputControl = quoted;
        if (!entry.controlMetaOriginal.inputControl) entry.controlMetaOriginal.inputControl = quoted;
        if (/buttoncontrol/i.test(val)) entry.isButton = true;
      }
      if (meta.defaultValue != null) {
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        entry.controlMeta.defaultValue = meta.defaultValue;
        if (!entry.controlMetaOriginal.defaultValue) entry.controlMetaOriginal.defaultValue = meta.defaultValue;
      }
    } catch (_) {}
  }

  function isPublished(sourceOp, source) {

    if (!state.parseResult) return false;

    return (state.parseResult.entries || []).some(e => e && e.sourceOp === sourceOp && e.source === source);

  }



  function ensurePublished(sourceOp, source, displayName, meta, options = {}) {

    if (!state.parseResult) return null;

    let idx = (state.parseResult.entries || []).findIndex(e => e && e.sourceOp === sourceOp && e.source === source);

    if (idx < 0) {

      if (!options.skipHistory) pushHistory('publish control');

      const key = makeUniqueKey(`${sourceOp}_${source}`);

      const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());

      const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;

      const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
      const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
      const targetPage = metaPage || activePage || 'Controls';

      const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);

      const entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };

      state.parseResult.entries.push(entry);
      idx = state.parseResult.entries.length - 1;
      state.parseResult.order.push(idx);

    }
    applyNodeControlMeta(state.parseResult.entries[idx], meta);

    if (!options.skipInsert) {
      try {

        const pos = getInsertionPosUnderSelection();

        state.parseResult.order = insertIndicesAt(state.parseResult.order, [idx], pos);

        try { logDiag(`ensurePublished: idx=${idx} moved to pos ${pos}`); } catch(_) {}

      } catch(_) {}
    }

    return { index: idx };

  }



  // Ensure entry exists without moving, return index

  function ensureEntryExists(sourceOp, source, displayName, meta) {

    if (!state.parseResult) return null;

    let idx = (state.parseResult.entries || []).findIndex(e => e && e.sourceOp === sourceOp && e.source === source);

    if (idx >= 0) return idx;

    const key = makeUniqueKey(`${sourceOp}_${source}`);

    const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());

    const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;

    const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
    const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
    const targetPage = metaPage || activePage || 'Controls';

    const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);

    const entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };

    state.parseResult.entries.push(entry);
    idx = state.parseResult.entries.length - 1;
    state.parseResult.order.push(idx);
    applyNodeControlMeta(state.parseResult.entries[idx], meta);

    return idx;

  }



  function removePublished(sourceOp, source) {

    if (!state.parseResult) return;

    pushHistory('unpublish control');

    const removedIdx = [];

    for (let i = 0; i < state.parseResult.entries.length; i++) {

      const e = state.parseResult.entries[i];

      if (e && e.sourceOp === sourceOp && e.source === source) removedIdx.push(i);

    }

    if (removedIdx.length) removePublishedByIndices(removedIdx);

  }



  function buildInstanceInputRaw(key, sourceOp, source, displayName, page, controlGroup) {

    const n = '\n';

    const b = [];

    b.push(`${key} = InstanceInput {`);

    if (sourceOp) b.push(`SourceOp = "${escapeQuotes(sourceOp)}",`);

    if (source) b.push(`Source = "${escapeQuotes(source)}",`);

    if (displayName) b.push(`Name = "${escapeQuotes(displayName)}",`);

    if (Number.isFinite(controlGroup)) b.push(`ControlGroup = ${controlGroup},`);
    const normalizedPage = page && String(page).trim();
    if (normalizedPage && normalizedPage !== 'Controls') {
      b.push(`Page = "${escapeQuotes(normalizedPage)}",`);
    }

    b.push(`}`);

    return b.join(n);

  }

  function buildBlendCheckboxInstanceInput(entry, eol) {

    const newline = eol || '\n';

    const key = entry?.key || makeUniqueKey(`${entry?.sourceOp || 'Blend'}_${entry?.source || 'Blend'}`);

    const pageName = entry?.page ? String(entry.page).trim() : 'Controls';

    const displayName = entry?.name || entry?.displayName || (entry?.source ? String(entry.source) : key);

    const segments = [];

    segments.push(`${key} = InstanceInput {`);

    if (entry?.sourceOp) segments.push(`SourceOp = "${escapeQuotes(entry.sourceOp)}",`);

    if (entry?.source) segments.push(`Source = "${escapeQuotes(entry.source)}",`);

    if (displayName) segments.push(`Name = "${escapeQuotes(displayName)}",`);

    if (Number.isFinite(entry?.controlGroup)) segments.push(`ControlGroup = ${entry.controlGroup},`);
    if (pageName && pageName !== 'Controls') {
      segments.push(`Page = "${escapeQuotes(pageName)}",`);
    }

    segments.push('INPID_InputControl = "CheckboxControl",');

    segments.push('INP_Integer = false,');

    segments.push('INP_Default = 1,');

    segments.push('INP_MinScale = 0,');

    segments.push('INP_MaxScale = 1,');

    segments.push('INP_MinAllowed = 0,');

    segments.push('INP_MaxAllowed = 1,');

    segments.push('CBC_TriState = false,');

    segments.push('LINKID_DataType = "Number",');

    segments.push('}');

    return segments.join(newline);

  }



  function makeUniqueKey(base) {

    const existing = new Set((state.parseResult?.entries || []).map(e => e.key));

    let k = sanitizeIdent(base);

    if (!existing.has(k)) return k;

    let i = 2;

    while (existing.has(`${k}_${i}`)) i++;

    return `${k}_${i}`;

  }



  function sanitizeIdent(s) {

    return String(s).replace(/[^A-Za-z0-9_\[\]\.]/g, '_');

  }




  function isPublished(sourceOp, source) {

    if (!state.parseResult) return false;

    return (state.parseResult.entries || []).some(e => e && e.sourceOp === sourceOp && e.source === source);

  }



  function ensurePublished(sourceOp, source, displayName, meta, options = {}) {

    if (!state.parseResult) return null;

    let idx = (state.parseResult.entries || []).findIndex(e => e && e.sourceOp === sourceOp && e.source === source);

    if (idx < 0) {

      if (!options.skipHistory) pushHistory('publish control');

      const key = makeUniqueKey(`${sourceOp}_${source}`);

      const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());

      const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;

      const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
      const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
      const targetPage = metaPage || activePage || 'Controls';

      const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);

      const entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };

      state.parseResult.entries.push(entry);
      idx = state.parseResult.entries.length - 1;
      state.parseResult.order.push(idx);

    }
    applyNodeControlMeta(state.parseResult.entries[idx], meta);

    if (!options.skipInsert) {
      try {

        const pos = getInsertionPosUnderSelection();

        state.parseResult.order = insertIndicesAt(state.parseResult.order, [idx], pos);

        try { logDiag(`ensurePublished: idx=${idx} moved to pos ${pos}`); } catch(_) {}

      } catch(_) {}
    }

    return { index: idx };

  }



  // Ensure entry exists without moving, return index

  function ensureEntryExists(sourceOp, source, displayName, meta) {

    if (!state.parseResult) return null;

    let idx = (state.parseResult.entries || []).findIndex(e => e && e.sourceOp === sourceOp && e.source === source);

    if (idx >= 0) return idx;

    const key = makeUniqueKey(`${sourceOp}_${source}`);

    const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());

    const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;

    const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
    const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
    const targetPage = metaPage || activePage || 'Controls';

    const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);

    const entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };

    state.parseResult.entries.push(entry);
    idx = state.parseResult.entries.length - 1;
    state.parseResult.order.push(idx);
    applyNodeControlMeta(state.parseResult.entries[idx], meta);

    return idx;

  }



  function removePublished(sourceOp, source) {

    if (!state.parseResult) return;

    pushHistory('unpublish control');

    const removedIdx = [];

    for (let i = 0; i < state.parseResult.entries.length; i++) {

      const e = state.parseResult.entries[i];

      if (e && e.sourceOp === sourceOp && e.source === source) removedIdx.push(i);

    }

    if (removedIdx.length) removePublishedByIndices(removedIdx);

  }



  function buildInstanceInputRaw(key, sourceOp, source, displayName, page, controlGroup) {

    const n = '\n';

    const b = [];

    b.push(`${key} = InstanceInput {`);

    if (sourceOp) b.push(`SourceOp = "${escapeQuotes(sourceOp)}",`);

    if (source) b.push(`Source = "${escapeQuotes(source)}",`);

    if (displayName) b.push(`Name = "${escapeQuotes(displayName)}",`);

    if (Number.isFinite(controlGroup)) b.push(`ControlGroup = ${controlGroup},`);
    const normalizedPage = page && String(page).trim();
    if (normalizedPage && normalizedPage !== 'Controls') {
      b.push(`Page = "${escapeQuotes(normalizedPage)}",`);
    }

    b.push(`}`);

    return b.join(n);

  }



  function makeUniqueKey(base) {

    const existing = new Set((state.parseResult?.entries || []).map(e => e.key));

    let k = sanitizeIdent(base);

    if (!existing.has(k)) return k;

    let i = 2;

    while (existing.has(`${k}_${i}`)) i++;

    return `${k}_${i}`;

  }



  function sanitizeIdent(s) {

    return String(s).replace(/[^A-Za-z0-9_\[\]\.]/g, '_');

  }




  function removePublishedByIndices(indices) {

    if (!state.parseResult) return;

    const toRemove = new Set(indices);

    const kept = [];

    const remap = new Map();

    for (let i = 0; i < state.parseResult.entries.length; i++) {

      if (!toRemove.has(i)) { remap.set(i, kept.length); kept.push(state.parseResult.entries[i]); }

    }

    const newOrder = state.parseResult.order.map(i => remap.get(i)).filter(i => i != null);

    state.parseResult.entries = kept;

    state.parseResult.order = newOrder;

    state.parseResult.originalOrder = [...newOrder];

    state.parseResult.selected = new Set();

    renderList(state.parseResult.entries, state.parseResult.order);

    updateRemoveSelectedState();

  }

























































  function parseModifiersInGroup(text, groupOpen, groupClose) {

    const out = [];

    const modsPos = text.indexOf('Modifiers = ordered()', groupOpen);

    if (modsPos < 0 || modsPos > groupClose) return out;

    const open = text.indexOf('{', modsPos);

    if (open < 0) return out;

    const close = findMatchingBrace(text, open);

    if (close < 0 || close > groupClose) return out;

    const inner = text.slice(open + 1, close);

    let i = 0, depth = 0, inStr = false;

    while (i < inner.length) {

      const ch = inner[i];

      if (inStr) { if (ch === '"' && inner[i-1] !== '\\') inStr = false; i++; continue; }

      if (ch === '"') { inStr = true; i++; continue; }

      if (ch === '{') { depth++; i++; continue; }

      if (ch === '}') { depth--; i++; continue; }

      if (depth === 0 && isIdentStart(ch)) {

        const nameStart = i; i++;

        while (i < inner.length && isIdentPart(inner[i])) i++;

        const modName = inner.slice(nameStart, i);

        while (i < inner.length && isSpace(inner[i])) i++;

        if (inner[i] !== '=') { i++; continue; }

        i++;

        while (i < inner.length && isSpace(inner[i])) i++;

        const typeStart = i; while (i < inner.length && isIdentPart(inner[i])) i++;

        const modType = inner.slice(typeStart, i);

        while (i < inner.length && isSpace(inner[i])) i++;

        if (inner[i] !== '{') { i++; continue; }

        const mOpen = i;

        const mClose = findMatchingBrace(inner, mOpen);

        if (mClose < 0) break;

        const body = inner.slice(mOpen + 1, mClose);

        out.push({ name: modName, type: modType, body, isModifier: true });

        i = mClose + 1; continue;

      }

      i++;

    }

    return out;

  }
  function hasPageProp(raw) {

    try {

      const open = raw.indexOf('{'); const close = raw.lastIndexOf('}');

      const body = (open >= 0 && close > open) ? raw.slice(open+1, close) : raw;

      return /(^|[\s,])Page\s*=\s*"/.test(body);

    } catch(_) { return false; }

  }

  function hasNameProp(raw) {
    try {
      const open = raw.indexOf('{'); const close = raw.lastIndexOf('}');
      const body = (open >= 0 && close > open) ? raw.slice(open+1, close) : raw;
      return /(^|[\s,])Name\s*=\s*"/.test(body);
    } catch(_) { return false; }
  }

  function applyPageIfMissing(raw, pageName, eol) {
    try {
      const normalized = (pageName && String(pageName).trim()) || '';
      if (!normalized || normalized === 'Controls') return raw;
      const open = raw.indexOf('{'); const close = raw.lastIndexOf('}');
      if (open < 0 || close < 0 || close <= open) return raw;
      const head = raw.slice(0, open+1);
      const body = raw.slice(open+1, close);
      const tail = raw.slice(close);
      if ((/(^|[\s,])Page\s*=\s*"/).test(body)) return raw;
      const insert = `${eol}Page = "${escapeQuotes(normalized)}",`;
      return head + insert + body + tail;
    } catch(_) { return raw; }
  }

  removeCommonPageBtn?.addEventListener('click', () => {
    try {
      if (!state.parseResult || !state.originalText) {
        error('Load a macro before updating the Settings page.');
        return;
      }
      const updated = removeCommonPageFromText(state.originalText, state.parseResult, state.newline || '\n');
      if (!updated || updated === state.originalText) {
        info('Common/Settings page already hidden.');
        return;
      }
      state.originalText = updated;
      removeCommonPageBtn.disabled = true;
      info('Common/Settings page hidden. Future exports will keep it hidden.');
    } catch (err) {
      error('Failed to update Settings page: ' + (err?.message || err));
    }
  });



  // Expose minimal debug helpers for console inspection

  try {

    window.FMR_DEBUG = {

      getAnchorPos: () => { try { return getInsertionPosUnderSelection(); } catch(e){ return String(e); } },

      getPinnedIndex: () => null,

      setPinnedIndex: (_i) => null,

      listSelected: () => state.parseResult ? Array.from(state.parseResult.selected || []) : [],

      order: () => state.parseResult ? Array.from(state.parseResult.order || []) : [],

      version: 'anchor-debug-1',
      setDiag: (on) => { try { setDiagnosticsEnabled(!!on); } catch(_) {} },
      toggleDiag: () => { try { setDiagnosticsEnabled(!diagnosticsController.isEnabled()); } catch(_) {} }

    };

  } catch(_) {}

  // Clipboard export helper (tries async Clipboard API, falls back to execCommand)

  async function writeToClipboard(text) {

    try {

      if (navigator && navigator.clipboard && navigator.clipboard.writeText) {

        await navigator.clipboard.writeText(text);

        return;

      }

    } catch (_) { /* fall through to fallback */ }

    // Fallback: hidden textarea + execCommand

    try {

      const ta = document.createElement('textarea');

      ta.style.position = 'fixed';

      ta.style.opacity = '0';

      ta.style.pointerEvents = 'none';

      ta.value = text;

      document.body.appendChild(ta);

      ta.focus();

      ta.select();

      const ok = document.execCommand('copy');

      document.body.removeChild(ta);

      if (!ok) throw new Error('execCommand copy failed');

    } catch (err) {
      // Last-resort prompt fallback so user can copy manually
      try {
        const msg = 'Copy the .setting content below (Ctrl/Cmd+A, then Ctrl/Cmd+C)';
        const joined = String(text || '');
        const promptShown = window.prompt(msg, joined);
        if (promptShown !== null) return; // user saw the content
      } catch (_) {}
      throw err;

    }

  }



  // --- Overrides for button URL rewrite (canonical insert) ---
  function rewriteBtncsExecuteInRange(text, openBraceIndex, closeBraceIndex, newUrl, eol) {
    try {
      const start = openBraceIndex + 1;
      const end = closeBraceIndex;
      // Remove any existing BTNCS_Execute entries inside the block
      const cleanedInside = stripBtncsEntries(text.slice(start, end));

      // Canonical Lua snippet built raw, then escaped for .setting string
      const url = String(newUrl || '');
      const luaRaw = `local url = "${url}"
local ok,err = pcall(function() return bmd and bmd.openurl and bmd.openurl(url) end)
if not ok then
  local sep = package.config:sub(1,1)
  if sep == "\\" then
    os.execute('start "" "'..url..'"')
  else
    local uname = io.popen("uname"):read("*l")
    if uname == "Darwin" then
      os.execute('open "'..url..'"')
    else
      os.execute('xdg-open "'..url..'"')
    end
  end
end
`;
      const lua = luaRaw.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r?\n/g, '\\n');

      // Indent one level within the control block
      let indent = '\t';
      try { indent = (getLineIndent(text, openBraceIndex) || '') + '\t'; } catch(_) {}
      const line = indent + 'BTNCS_Execute = "' + lua + '",' + eol;

      // Rebuild with cleaned inside, then inject at start of block
      const rebuilt = text.slice(0, start) + cleanedInside + text.slice(end);
      return rebuilt.slice(0, start) + eol + line + rebuilt.slice(start);
    } catch(_) { return text; }
  }

  // Always rewrite Button code canonically using the URL field value
  function applyLaunchUrlOverrides(text, result, eol) {
    try {
      const overrides = (result && result.buttonOverrides && typeof result.buttonOverrides.forEach === 'function') ? result.buttonOverrides : null;
      if (!overrides || result.buttonOverrides.size === 0) return text;
      const grp = locateMacroGroupBounds(text, result);
      if (!grp) return text;
      let out = text;
      overrides.forEach((newUrl, key) => {
        try {
          const v = String(newUrl || '').trim();
          if (!v) return;
          const dot = String(key).indexOf('.');
          if (dot < 0) return;
          const tool = key.slice(0, dot);
          const ctrl = key.slice(dot + 1);
          let tb = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, tool);
          let uc = tb ? findUserControlsInTool(out, tb.open, tb.close)
                      : findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
          if (!uc) return;
          const cb = findControlBlockInUc(out, uc.open, uc.close, ctrl);
          if (!cb) return;
          const before = out;
          out = rewriteBtncsExecuteInRange(out, cb.open, cb.close, v, eol);
          if (out !== before) { try { logDiag(`[URL override] ${key} -> rewritten canonical`); } catch(_) {} }
        } catch(_) {}
      });
      return out;
    } catch(_) { return text; }
  }

  const catalogEditor = createCatalogEditor({
    openBtn: catalogEditorOpenBtn,
    closeBtn: catalogEditorCloseBtn,
    root: catalogEditorRoot,
    datasetSelect: catalogDatasetSelect,
    searchInput: catalogTypeSearch,
    typeList: catalogTypeList,
    controlList: catalogControlList,
    emptyState: catalogEmptyState,
    downloadBtn: catalogDownloadBtn,
    nodesPane,
    logDiag,
    info,
    error,
  });


})();


  // Final safety: directly locate control by id in any UserControls range and inject snippet
  function applyCanonicalLaunchSnippets(text, result, eol) {
    try {
      const overrides = (result && result.buttonOverrides && typeof result.buttonOverrides.forEach === 'function') ? result.buttonOverrides : null;
      if (!overrides || result.buttonOverrides.size === 0) return text;
      let out = text;
      // collect all group-level and tool-level UserControls ranges, scan each for control id
      const ranges = [];
      let pos = 0;
      while (true) {
        const i = out.indexOf('UserControls = ordered()', pos);
        if (i < 0) break;
        const brace = out.indexOf('{', i);
        if (brace < 0) break;
        const close = findMatchingBrace(out, brace);
        if (close < 0) break;
        ranges.push({ open: brace, close });
        pos = close + 1;
      }
      if (!ranges.length) return out;
      overrides.forEach((newUrl, key) => {
        try {
          const v = String(newUrl || '').trim();
          if (!v) return;
          const dot = String(key).indexOf('.');
          if (dot < 0) return;
          const ctrl = key.slice(dot + 1);
          // try each range until found
          for (const rg of ranges) {
            const cb = findControlBlockInUc(out, rg.open, rg.close, ctrl);
            if (cb && cb.open && cb.close) {
              const before = out;
              out = rewriteBtncsExecuteInRange(out, cb.open, cb.close, v, eol);
              if (out !== before) { try { logDiag(`[URL override] ${key} -> direct inject`); } catch(_) {} }
              break;
            }
          }
        } catch(_) {}
      });
      return out;
    } catch(_) { return text; }
  }

  // Last-resort: locate control block by id anywhere and inject canonical BTNCS_Execute
  function forceRewriteButton(text, controlId, newUrl, eol) {
    try {
      const escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const id = String(controlId);
      const re = new RegExp('(?:^|[\\s,])(' + escRe(id) + '|\\[\\s*\"' + escRe(id) + '\"\\s*\\])\\s*=\\s*\\{','g');
      let out = text; let m;
      while ((m = re.exec(out)) !== null) {
        const mat = m[0];
        const braceRel = mat.lastIndexOf('{');
        const open = (m.index + braceRel);
        if (open < 0) continue;
        const close = findMatchingBrace(out, open);
        if (close < 0) continue;
        // Remove any existing BTNCS_Execute entries inside this block
        let body = stripBtncsEntries(out.slice(open + 1, close));
        const url = String(newUrl || '');
        const luaRaw = `local url = "${url}"
local sep = package.config:sub(1,1)
if sep == "\\" then
  os.execute('start "" "'..url..'"')
else
  local uname = io.popen("uname"):read("*l")
  if uname == "Darwin" then
    os.execute('open "'..url..'"')
  else
    os.execute('xdg-open "'..url..'"')
  end
end
`;
        const lua = luaRaw.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r?\n/g, '\\n');
        let indent = '\t';
        try { indent = (getLineIndent(out, open) || '') + '\t'; } catch(_) {}
        const line = indent + 'BTNCS_Execute = "' + lua + '",' + eol;
        const newBody = eol + line + body;
        out = out.slice(0, open + 1) + newBody + out.slice(close);
        // continue scanning after close
        re.lastIndex = open + 1 + newBody.length;
      }
      return out;
    } catch(_) { return text; }
  }
  function runValidation(context) {
    try {
      if (!state.parseResult || !state.originalText) return [];
      const issues = [];
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      const declaredPages = listDeclaredControlPages(state.originalText, bounds);
      const usedPages = new Set();
      (state.parseResult.entries || []).forEach(entry => {
        if (!entry) return;
        const pageName = (entry.page && String(entry.page).trim()) ? String(entry.page).trim() : 'Controls';
        if (pageName) usedPages.add(pageName);
      });
      usedPages.forEach(page => {
        if (page === 'Controls') return;
        if (!declaredPages.has(page)) {
          issues.push({ type: 'page', page, message: `Control page "${page}" is referenced but has no ControlPage definition.` });
        }
      });
      const missingControls = [];
      for (const entry of (state.parseResult.entries || [])) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        if (!shouldValidateControlEntry(entry)) continue;
        if (!controlBlockExists(state.originalText, bounds, entry.sourceOp, entry.source)) {
          missingControls.push(`${entry.sourceOp}.${entry.source}`);
        }
      }
      missingControls.forEach(ctrl => {
        issues.push({ type: 'control', control: ctrl, message: `Published control "${ctrl}" has no matching UserControls entry.` });
      });
      const invalidKeys = [];
      for (const entry of (state.parseResult.entries || [])) {
        if (!entry) continue;
        if (isKeyLikelyValid(entry.key)) continue;
        invalidKeys.push(`${entry.sourceOp || '?'}:${entry.source || '?'} (${entry.key || 'blank'})`);
      }
      invalidKeys.forEach(keyInfo => {
        issues.push({ type: 'key', key: keyInfo, message: `Invalid InstanceInput key detected for ${keyInfo}.` });
      });
      state.validationIssues = issues;
      if (issues.length) {
        const label = context ? `Validation (${context})` : 'Validation';
        info(`${label}: ${issues.length} issue${issues.length === 1 ? '' : 's'} detected. See diagnostics for details.`);
        issues.forEach(issue => {
          try {
            logDiag(`[Validation] ${issue.message}`);
          } catch (_) {}
        });
      } else if (context === 'manual') {
        info('Validation: no issues detected.');
      }
      return issues;
    } catch (_) {
      return [];
    }
  }

  function shouldValidateControlEntry(entry) {
    try {
      if (!entry) return false;
      if (entry.isLabel || entry.isButton) return true;
      if (entry.controlMeta && entry.controlMeta.inputControl) return true;
      if (entry.controlMetaOriginal && entry.controlMetaOriginal.inputControl) return true;
      if (entry.source && /^[A-Za-z_]/.test(entry.source)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function listDeclaredControlPages(text, bounds) {
    const pages = new Set(['Controls']);
    try {
      if (!text || !bounds) return pages;
      const segment = text.slice(bounds.groupOpenIndex, bounds.groupCloseIndex);
      const re = /(^|\n)\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*ControlPage\s*\{/g;
      let m;
      while ((m = re.exec(segment)) !== null) {
        const name = (m[2] || '').trim();
        if (name) pages.add(name);
        if (re.lastIndex === m.index) re.lastIndex++;
      }
      return pages;
    } catch (_) {
      return pages;
    }
  }

  function controlBlockExists(text, bounds, toolName, controlId) {
    try {
      if (!text || !bounds || !toolName || !controlId) return false;
      const block = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (block) {
        const uc = findUserControlsInTool(text, block.open, block.close);
        if (uc) {
          const cb = findControlBlockInUc(text, uc.open, uc.close, controlId);
          if (cb) return true;
        }
      }
      const groupUc = findUserControlsInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (groupUc) {
        const cb = findControlBlockInUc(text, groupUc.open, groupUc.close, controlId);
        if (cb) return true;
      }
      return false;
    } catch (_) {
      return false;
    }
  }
