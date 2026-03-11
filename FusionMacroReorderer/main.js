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
import { buildLabelMarkup, normalizeLabelStyle, labelStyleEquals } from './src/labelMarkup.js';
import { highlightLua } from './src/luaHighlighter.js';
import { createIntroPanelController } from './src/ui/introPanel.js';
import { createExportMenuController } from './src/ui/exportMenu.js';
import { createDocTabContextMenu } from './src/ui/docTabContextMenu.js';
import { createDocTabsController } from './src/ui/docTabs.js';
import { createGlobalShortcuts } from './src/ui/shortcuts.js';
import { createFileDropController } from './src/ui/fileDrop.js';
import { createAddControlModal } from './src/ui/addControlModal.js';
import { createTextPrompt } from './src/ui/textPrompt.js';

(() => {
  try {
    const stamp = new Date().toISOString();
    console.log('[FMR Build]', stamp);
  } catch (_) {}

  const state = appState;
  const getNativeApi = () => (typeof window !== 'undefined' ? window.FusionMacroReordererNative : null);

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
  const miniCreateFolderBtn = document.getElementById('miniCreateFolderBtn');
  const macroExplorerMiniFrame = document.getElementById('macroExplorerMiniFrame');
  const macroExplorerFrame = document.getElementById('macroExplorerFrame');
  const macroExplorerPanel = document.getElementById('macroExplorerPanel');
  const macroExplorerMini = document.getElementById('macroExplorerMini');
  const closeExplorerPanelBtn = document.getElementById('closeExplorerPanelBtn');
  const exportMenuBtn = document.getElementById('exportMenuBtn');
  const exportMenuCaretBtn = document.getElementById('exportMenuCaretBtn');
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
  const {
    setIntroCollapsed,
    setIntroToggleVisible,
    updateDropHint,
  } = createIntroPanelController({
    introPanel,
    introToggleBtn,
    openExplorerBtn,
    closeExplorerPanelBtn,
    macroExplorerPanel,
    macroExplorerMini,
    macroExplorerFrame,
    macroExplorerMiniFrame,
    dropHint,
    isElectron: IS_ELECTRON,
    onClearSelections: () => {
      if (typeof clearSelections === 'function') clearSelections();
    },
    onHideDetailDrawer: () => hideDetailDrawer(),
    onSyncDrawerAnchor: () => syncDetailDrawerAnchor(),
    onInfo: (msg) => info(msg),
  });

  const documents = [];
  let activeDocId = null;
  let docCounter = 1;
  let suppressDocDirty = false;
  let suppressDocTabsRender = false;
  let pendingDocRenderTimer = null;
  let draggingDocId = null;
  let activeCsvBatchId = null;
  const FMR_HEADER_LABEL_CONTROL = 'FMR_HeaderVisual';
  const FMR_PATH_DRAW_MODE_META_MARKER = '-- FMR_PATH_DRAW_MODE_META';
  const FMR_PATH_DRAW_MODE_LEGACY_SCRIPT_MARKER = '-- FMR_PATH_DRAW_MODE';
  const FMR_PATH_DRAW_MODE_SOURCEOP_PROP = 'FMR_PathDrawModeSourceOp';
  const FMR_PATH_DRAW_MODE_TARGET_PROP = 'FMR_PathDrawModeTarget';
  const FMR_PATH_DRAW_MODE_INDEX_PROP = 'FMR_PathDrawModeIndex';
  const PATH_DRAW_MODE_OPTIONS = ['ClickAppend', 'Freehand', 'InsertAndModify', 'ModifyOnly', 'Done'];
  let exportMenuController = null;
  let docTabContextMenu = null;
  let pendingReloadAfterImport = false;
  let docTabsController = null;
  const textPrompt = createTextPrompt();

  function openConfirmModal({ title, message, confirmText, cancelText } = {}) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal';
      overlay.setAttribute('role', 'dialog');
      overlay.setAttribute('aria-modal', 'true');

      const form = document.createElement('form');
      form.className = 'add-control-form';

      const header = document.createElement('header');
      const headerWrap = document.createElement('div');
      const titleEl = document.createElement('h3');
      titleEl.textContent = title || 'Confirm';
      headerWrap.appendChild(titleEl);

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = 'x';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headerWrap);
      header.appendChild(closeBtn);

      const body = document.createElement('div');
      body.className = 'form-body';
      const p = document.createElement('p');
      p.style.whiteSpace = 'pre-line';
      p.textContent = message || 'Are you sure?';
      body.appendChild(p);

      const actions = document.createElement('div');
      actions.className = 'modal-actions';
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = cancelText || 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.textContent = confirmText || 'OK';
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);

      form.appendChild(header);
      form.appendChild(body);
      form.appendChild(actions);
      overlay.appendChild(form);
      document.body.appendChild(overlay);

      const cleanup = (value) => {
        try { overlay.remove(); } catch (_) {}
        resolve(value);
      };

      const onCancel = () => cleanup(false);
      const onSubmit = (ev) => { ev.preventDefault(); cleanup(true); };
      const onOverlayClick = (ev) => { if (ev.target === overlay) onCancel(); };
      const onKeyDown = (ev) => { if (ev.key === 'Escape') { ev.preventDefault(); onCancel(); } };

      cancelBtn.addEventListener('click', onCancel);
      closeBtn.addEventListener('click', onCancel);
      form.addEventListener('submit', onSubmit);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeyDown, { once: true });
    });
  }

  function createDocumentMeta({ name, fileName } = {}) {
    return {
      id: `doc_${docCounter++}`,
      name: name || fileName || 'Untitled',
      fileName: fileName || '',
      isDirty: false,
      selected: false,
      createdAt: Date.now(),
      snapshot: null,
      isCsvBatch: false,
      csvBatch: null,
    };
  }

  function addDocument(doc, makeActive = true) {
    if (!doc) return null;
    documents.push(doc);
    if (makeActive || !activeDocId) activeDocId = doc.id;
    renderDocTabs();
    return doc;
  }

  function createLazyDocumentFromText({ name, fileName, text }) {
    const doc = createDocumentMeta({ name, fileName });
    doc.isDirty = true;
    doc.snapshot = {
      parseResult: null,
      originalText: text || '',
      originalFileName: fileName || name || 'Imported.setting',
      originalFilePath: '',
      newline: detectNewline(text || ''),
      generatedText: text || '',
      lastDiffRange: null,
      exportFolder: state.exportFolder || '',
      exportFolderSelected: !!state.exportFolderSelected,
      lastExportPath: '',
      drfxLink: null,
      csvData: state.csvData || null,
      lazyParse: true,
      csvGenerated: true,
    };
    return addDocument(doc, false);
  }

  function addCsvBatchDocument(info = {}) {
    const count = Number.isFinite(info.count) ? info.count : 0;
    const label = `CSV Batch (${count})`;
    const doc = createDocumentMeta({ name: label, fileName: '' });
    doc.isCsvBatch = true;
    doc.csvBatch = {
      count,
      folderPath: info.folderPath || '',
      baseName: info.baseName || '',
      sourceDocId: info.sourceDocId || null,
      createdAt: Date.now(),
    };
    doc.csvBatchSnapshot = info.snapshot || null;
    doc.snapshot = info.snapshot || null;
    doc.isDirty = false;
    return addDocument(doc, false);
  }

  function selectCsvBatchDoc(doc) {
    if (!doc) return;
    activeDocId = doc.id;
    doc.selected = true;
    activeCsvBatchId = doc.id;
    renderDocTabs();
    info('CSV batch selected. Use Export to regenerate files.');
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

  function selectAllDocs() {
    let changed = false;
    documents.forEach((doc) => {
      if (!doc.selected) {
        doc.selected = true;
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
    state.csvData = null;
    updateDataMenuState();
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
        setIntroCollapsed(false);
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
      setIntroCollapsed(false);
    }
    renderDocTabs();
  }

  docTabContextMenu = createDocTabContextMenu({
    getDocument: (docId) => documents.find((doc) => doc.id === docId) || null,
    onCreateDoc: () => createBlankDocument(),
    onRenameDoc: (docId) => promptRenameDocument(docId),
    onToggleDocSelection: (docId) => toggleDocSelection(docId),
    onSelectAll: () => selectAllDocs(),
    onClearSelection: () => clearDocSelections(),
    onCloseDoc: (docId) => closeDocument(docId),
    onCloseOthers: (docId) => closeDocumentsByFilter(d => d.id !== docId, 'Close other tabs without exporting?'),
    onCloseAll: () => closeDocumentsByFilter(() => true, 'Close all tabs without exporting?'),
  });

  docTabsController = createDocTabsController({
    docTabsEl,
    docTabsWrap,
    docTabsPrev,
    docTabsNext,
    getDocuments: () => documents,
    getActiveDocId: () => activeDocId,
    getDraggingDocId: () => draggingDocId,
    setDraggingDocId: (value) => { draggingDocId = value; },
    onReorderDocuments: (fromId, targetId, appendToEnd) => reorderDocuments(fromId, targetId, appendToEnd),
    onToggleDocSelection: (docId) => toggleDocSelection(docId),
    onClearDocSelections: () => clearDocSelections(),
    onSwitchDocument: (docId) => switchToDocument(docId),
    onPromptRename: (docId) => promptRenameDocument(docId),
    onOpenContextMenu: (docId, x, y) => docTabContextMenu?.open?.(docId, x, y),
    onCloseDocument: (docId) => closeDocument(docId),
    onCreateBlankDocument: () => handleNativeOpen(),
    onCreateFromFile: () => handleNativeOpen(),
    onCreateFromClipboard: () => handleImportFromClipboard(),
    onUpdateExportPathDisplay: () => updateDocExportPathDisplay(),
    onSelectCsvBatch: (doc) => selectCsvBatchDoc(doc),
    getDocDisplayName: (doc, docs) => getDocDisplayName(doc, docs),
  });

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
      csvData: state.csvData || null,
    };
  }

  function storeDocumentSnapshot(doc) {
    if (!doc) return null;
    const snap = buildDocumentSnapshot();
    if (!snap) return doc;
    doc.snapshot = snap;
    return doc;
  }

  function applyDocumentSnapshot(doc, options = {}) {
    const snap = doc && doc.snapshot ? doc.snapshot : null;
    const prevSuppress = suppressDocDirty;
    suppressDocDirty = true;
    try {
      if (snap && snap.lazyParse && snap.originalText) {
        loadMacroFromText(snap.originalFileName || doc?.fileName || doc?.name || 'Imported.setting', snap.originalText, {
          createDoc: false,
          preserveFileInfo: true,
          preserveFilePath: true,
          allowAutoUtility: false,
          silentAuto: true,
          skipClear: false,
        });
        storeDocumentSnapshot(doc);
        return;
      }
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
      state.csvData = snap.csvData || null;
      updateDataMenuState();

      const codeText = snap.generatedText != null ? snap.generatedText : (state.originalText || '');
      updateCodeView(codeText);
      clearCodeHighlight();
      if (fileInfo) {
        fileInfo.textContent = `${state.originalFileName} (${(state.originalText || '').length.toLocaleString()} chars)`;
      }
      setMacroNameDisplay(state.parseResult.macroName || 'Unknown');
      const entries = state.parseResult.entries || [];
      if (ctrlCountEl) ctrlCountEl.textContent = String(entries.length);
      if (inputCountEl) inputCountEl.textContent = String(countRecognizedInputs(entries));
      controlsSection.hidden = false;
      resetBtn.disabled = false;
      exportBtn.disabled = false;
      exportTemplatesBtn && (exportTemplatesBtn.disabled = false);
      exportClipboardBtn && (exportClipboardBtn.disabled = false);
      updateExportMenuButtonState();
      removeCommonPageBtn && (removeCommonPageBtn.disabled = false);
      publishSelectedBtn && (publishSelectedBtn.disabled = true);
      clearNodeSelectionBtn && (clearNodeSelectionBtn.disabled = true);
      setCreateControlActionsEnabled(true);
      pendingControlMeta.clear();
      pendingOpenNodes.clear();
      updateExportButtonLabelFromPath(state.originalFilePath);
      updateIntroToggleVisibility();
      if (options.deferHeavy) {
        scheduleDeferredDocRender();
      } else {
        runHeavyDocumentRender();
      }
    } finally {
      suppressDocDirty = prevSuppress;
    }
  }

  function runHeavyDocumentRender() {
    publishedControls?.resetPageOptions?.();
    renderActiveList();
    refreshPageTabs?.();
    updateRemoveSelectedState();
    hideDetailDrawer();
    clearHighlights?.();
    nodesPane?.clearNodeSelection?.();
    nodesPane?.parseAndRenderNodes?.();
    updateUndoRedoState();
    updateUtilityActionsState();
  }

  function scheduleDeferredDocRender() {
    if (pendingDocRenderTimer) {
      clearTimeout(pendingDocRenderTimer);
      pendingDocRenderTimer = null;
    }
    const targetId = activeDocId;
    pendingDocRenderTimer = setTimeout(() => {
      pendingDocRenderTimer = null;
      if (activeDocId !== targetId) return;
      runHeavyDocumentRender();
    }, 180);
  }

  function switchToDocument(docId) {
    if (!docId || docId === activeDocId) return;
    const current = getActiveDocument();
    if (current) storeDocumentSnapshot(current);
    const next = documents.find((doc) => doc.id === docId);
    if (!next) return;
    activeDocId = next.id;
    applyDocumentSnapshot(next, { deferHeavy: true });
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
    if (suppressDocTabsRender) return;
    docTabsController?.render?.();
    updateCsvBatchBanner();
    syncDataLinkPanel();
  }

  function getActiveCsvBatch() {
    const active = getActiveDocument();
    if (active && active.isCsvBatch) return active;
    const selected = documents.filter(doc => doc && doc.selected && doc.isCsvBatch);
    return selected.length ? selected[0] : null;
  }

  function updateCsvBatchBanner() {
    if (!csvBatchBanner) return;
    const batch = getActiveCsvBatch();
    if (!batch) {
      csvBatchBanner.hidden = true;
      return;
    }
    const count = batch.csvBatch && Number.isFinite(batch.csvBatch.count) ? batch.csvBatch.count : 0;
    csvBatchBanner.textContent = `CSV Batch selected (${count}) — Export will regenerate files.`;
    csvBatchBanner.hidden = false;
  }



  function openDataSourceModal(source) {
    const dataSource = String(source || '').trim();
    if (!dataSource) return;
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal csv-modal';
    const modal = document.createElement('form');
    modal.className = 'add-control-form';
    const header = document.createElement('header');
    const headerText = document.createElement('div');
    const eyebrow = document.createElement('div');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Data';
    const title = document.createElement('h3');
    title.textContent = 'Data source URL';
    headerText.appendChild(eyebrow);
    headerText.appendChild(title);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.innerHTML = '&times;';
    header.appendChild(headerText);
    header.appendChild(closeBtn);
    const body = document.createElement('div');
    body.className = 'form-body';
    const input = document.createElement('input');
    input.type = 'text';
    input.value = dataSource;
    input.readOnly = true;
    body.appendChild(input);
    const actions = document.createElement('footer');
    actions.className = 'modal-actions';
    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.textContent = 'Copy';
    copyBtn.addEventListener('click', async () => {
      try {
        if (navigator?.clipboard?.writeText) {
          await navigator.clipboard.writeText(dataSource);
        }
      } catch (_) {}
    });
    const doneBtn = document.createElement('button');
    doneBtn.type = 'button';
    doneBtn.className = 'primary';
    doneBtn.textContent = 'Close';
    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    closeBtn.addEventListener('click', close);
    doneBtn.addEventListener('click', close);
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay) close();
    });
    actions.appendChild(copyBtn);
    actions.appendChild(doneBtn);
    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    input.focus();
    input.select();
  }

  function openDataLinkConfigModal() {
    if (!state.parseResult) return;
    const link = ensureDataLink(state.parseResult);
    const dataSource = link?.source || state.csvData?.sourceName || '';
    if (!dataSource) {
      error('No data source set for this macro.');
      return;
    }
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal csv-modal';
    const modal = document.createElement('form');
    modal.className = 'add-control-form';
    const header = document.createElement('header');
    const headerText = document.createElement('div');
    const eyebrow = document.createElement('div');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Data';
    const title = document.createElement('h3');
    title.textContent = 'Data Link Settings';
    headerText.appendChild(eyebrow);
    headerText.appendChild(title);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.innerHTML = '&times;';
    header.appendChild(headerText);
    header.appendChild(closeBtn);
    const body = document.createElement('div');
    body.className = 'form-body';

    const sourceRow = document.createElement('div');
    sourceRow.className = 'detail-actions';
    const sourceBtn = document.createElement('button');
    sourceBtn.type = 'button';
    sourceBtn.textContent = 'View source URL';
    sourceBtn.addEventListener('click', () => openDataSourceModal(dataSource));
    sourceRow.appendChild(sourceBtn);
    body.appendChild(sourceRow);

    const modeLabel = document.createElement('label');
    modeLabel.textContent = 'Row mode';
    const modeSelect = document.createElement('select');
    ['first', 'index', 'key'].forEach((value) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = value === 'first' ? 'First row'
        : value === 'index' ? 'Row number'
          : 'Key match';
      modeSelect.appendChild(opt);
    });
    modeSelect.value = link.rowMode || 'first';
    modeLabel.appendChild(modeSelect);
    body.appendChild(modeLabel);

    const indexLabel = document.createElement('label');
    indexLabel.textContent = 'Row #';
    const indexInput = document.createElement('input');
    indexInput.type = 'number';
    indexInput.min = '1';
    indexInput.value = String(link.rowIndex || 1);
    indexLabel.appendChild(indexInput);
    body.appendChild(indexLabel);

    const keyLabel = document.createElement('label');
    keyLabel.textContent = 'Key column';
    const keyInput = document.createElement('input');
    keyInput.type = 'text';
    keyInput.value = link.rowKey || '';
    keyLabel.appendChild(keyInput);
    body.appendChild(keyLabel);

    const valueLabel = document.createElement('label');
    valueLabel.textContent = 'Key value';
    const valueInput = document.createElement('input');
    valueInput.type = 'text';
    valueInput.value = link.rowValue || '';
    valueLabel.appendChild(valueInput);
    body.appendChild(valueLabel);

    const toggleVisibility = () => {
      const mode = modeSelect.value || 'first';
      indexLabel.style.display = mode === 'index' ? '' : 'none';
      keyLabel.style.display = mode === 'key' ? '' : 'none';
      valueLabel.style.display = mode === 'key' ? '' : 'none';
    };
    toggleVisibility();
    modeSelect.addEventListener('change', toggleVisibility);

    const actions = document.createElement('footer');
    actions.className = 'modal-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Close';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'submit';
    saveBtn.className = 'primary';
    saveBtn.textContent = 'Save';
    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    cancelBtn.addEventListener('click', close);
    closeBtn.addEventListener('click', close);
    modal.addEventListener('submit', (ev) => {
      ev.preventDefault();
      link.rowMode = modeSelect.value || 'first';
      const val = Number(indexInput.value || 1);
      link.rowIndex = Number.isFinite(val) ? val : 1;
      link.rowKey = keyInput.value || '';
      link.rowValue = valueInput.value || '';
      markActiveDocumentDirty();
      close();
    });
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay) close();
    });
    actions.appendChild(cancelBtn);
    actions.appendChild(saveBtn);
    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  function resolveReloadPath() {
    const direct =
      state.parseResult?.fileMeta?.exportPath ||
      state.originalFilePath ||
      state.lastExportPath ||
      '';
    if (direct) return direct;
    const folder = state.exportFolder || '';
    if (!folder || !state.parseResult) return '';
    const fileName = buildMacroExportName();
    return buildExportPath(folder, fileName);
  }

  function countPresetPackEntries(result) {
    const countPresetChoicesFromText = (textValue) => {
      try {
        const text = String(textValue || '');
        if (!text) return 0;
        let best = 0;
        const controlRe = /(^|\n)\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*\{/g;
        let match;
        while ((match = controlRe.exec(text))) {
          const controlId = String(match[2] || '').trim();
          if (!/preset/i.test(controlId)) continue;
          const open = text.indexOf('{', match.index);
          if (open < 0) continue;
          const close = findMatchingBrace(text, open);
          if (close < 0) continue;
          const body = text.slice(open + 1, close);
          if (!/INPID_InputControl\s*=\s*"(?:ComboControl|MultiButtonControl)"/i.test(body)) continue;
          const comboCount = (body.match(/CCS_AddString\s*=/g) || []).length;
          const buttonCount = (body.match(/MBTNC_AddButton\s*=/g) || []).length;
          const count = Math.max(comboCount, buttonCount);
          if (count > best) best = count;
        }
        return best;
      } catch (_) {
        return 0;
      }
    };
    try {
      const engine = result?.presetEngine && typeof result.presetEngine === 'object' ? result.presetEngine : null;
      if (engine) {
        const presetCount = Object.keys(engine.presets || {}).filter((name) => String(name || '').trim()).length;
        const scopeCount = Array.isArray(engine.scopeEntries) ? engine.scopeEntries.length : 0;
        if (presetCount > 0 && scopeCount > 0) return presetCount;
      }
    } catch (_) {
      // fallback below
    }
    try {
      let best = 0;
      const groupControls = Array.isArray(result?.groupUserControls) ? result.groupUserControls : [];
      groupControls.forEach((control) => {
        if (!control) return;
        const id = String(control.id || '').trim();
        const name = String(control.name || '').trim();
        const inputControl = String(control.inputControl || '').trim();
        const options = Array.isArray(control.choiceOptions)
          ? control.choiceOptions.filter((option) => String(option || '').trim())
          : [];
        if (!options.length) return;
        const isChoiceLike = /(?:combo|multibutton)/i.test(inputControl);
        const isPresetNamed = /preset/i.test(id) || /preset/i.test(name);
        if (isChoiceLike && isPresetNamed) best = Math.max(best, options.length);
      });
      // Fallback for imported macros that only carry selector-style published controls.
      const entries = Array.isArray(result?.entries) ? result.entries : [];
      entries.forEach((entry) => {
        if (!entry) return;
        const source = String(entry.source || '').trim();
        const display = String(entry.displayName || entry.name || '').trim();
        const options = getChoiceOptions(entry);
        const count = Array.isArray(options)
          ? options.filter((option) => String(option || '').trim()).length
          : 0;
        if (!count) return;
        if (source === 'FMR_Preset' || /preset/i.test(source) || /preset/i.test(display)) {
          best = Math.max(best, count);
        }
      });
      if (best <= 0) {
        best = Math.max(best, countPresetChoicesFromText(state.originalText || ''));
      }
      return best;
    } catch (_) {
      return 0;
    }
  }

  function parseNativeVersionsInfoFromText(text) {
    try {
      const raw = String(text || '');
      if (!raw) return { count: 0, currentSettings: 0, maxSlotIndex: 0 };
      const countTopLevelTableEntries = (body) => {
        try {
          const source = String(body || '');
          if (!source) return 0;
          let depth = 0;
          let inString = false;
          let awaitingValue = false;
          let count = 0;
          for (let i = 0; i < source.length; i += 1) {
            const ch = source[i];
            const prev = source[i - 1];
            if (ch === '"' && prev !== '\\') {
              inString = !inString;
              continue;
            }
            if (inString) continue;
            if (depth === 0) {
              if (awaitingValue) {
                if (/\s/.test(ch)) continue;
                if (ch === '{') {
                  count += 1;
                  awaitingValue = false;
                  depth = 1;
                  continue;
                }
                awaitingValue = false;
              }
              if (ch === '=') {
                awaitingValue = true;
                continue;
              }
            }
            if (ch === '{') depth += 1;
            else if (ch === '}') depth = Math.max(0, depth - 1);
          }
          return count;
        } catch (_) {
          return 0;
        }
      };
      let currentSettings = 0;
      const currentRe = /\bCurrentSettings\s*=\s*(\d+)/ig;
      let currentMatch = null;
      while ((currentMatch = currentRe.exec(raw))) {
        const parsed = Number.parseInt(currentMatch[1], 10);
        if (Number.isFinite(parsed) && parsed > currentSettings) currentSettings = parsed;
      }
      let maxSlotIndex = 0;
      let settingsEntryCount = 0;
      const customDataRe = /\bCustomData\s*=\s*\{/ig;
      let customDataMatch = null;
      while ((customDataMatch = customDataRe.exec(raw))) {
        const customOpen = raw.indexOf('{', customDataMatch.index);
        const customClose = customOpen >= 0 ? findMatchingBrace(raw, customOpen) : -1;
        if (!(customOpen >= 0 && customClose > customOpen)) continue;
        const customBody = raw.slice(customOpen + 1, customClose);
        const settingsRe = /\bSettings\s*=\s*\{/ig;
        let settingsMatch = null;
        while ((settingsMatch = settingsRe.exec(customBody))) {
          const settingsOpenLocal = customBody.indexOf('{', settingsMatch.index);
          const settingsCloseLocal = settingsOpenLocal >= 0 ? findMatchingBrace(customBody, settingsOpenLocal) : -1;
          if (!(settingsOpenLocal >= 0 && settingsCloseLocal > settingsOpenLocal)) continue;
          const settingsBody = customBody.slice(settingsOpenLocal + 1, settingsCloseLocal);
          const slotRe = /\[(\d+)\]\s*=/g;
          let slotMatch = null;
          while ((slotMatch = slotRe.exec(settingsBody))) {
            const idx = Number.parseInt(slotMatch[1], 10);
            if (Number.isFinite(idx) && idx > maxSlotIndex) maxSlotIndex = idx;
          }
          const topLevelCount = countTopLevelTableEntries(settingsBody);
          if (topLevelCount > settingsEntryCount) settingsEntryCount = topLevelCount;
        }
      }
      const slotCount = maxSlotIndex > 0 ? maxSlotIndex + 1 : 0;
      const namedSlotCount = settingsEntryCount > 1 ? settingsEntryCount : 0;
      if (!currentSettings && !slotCount && !namedSlotCount) {
        return { count: 0, currentSettings: 0, maxSlotIndex: 0, settingsEntryCount: 0 };
      }
      const count = currentSettings > 0
        ? Math.max(currentSettings, slotCount, namedSlotCount)
        : Math.max(slotCount, namedSlotCount);
      return {
        count: Number.isFinite(count) ? Math.max(0, Math.floor(count)) : 0,
        currentSettings,
        maxSlotIndex,
        settingsEntryCount,
      };
    } catch (_) {
      return { count: 0, currentSettings: 0, maxSlotIndex: 0, settingsEntryCount: 0 };
    }
  }

  function detectNativeVersionsCountFromText(text) {
    try {
      const info = parseNativeVersionsInfoFromText(text);
      return Number.isFinite(info?.count) ? Math.max(0, Math.floor(info.count)) : 0;
    } catch (_) {
      return 0;
    }
  }

  function openNativeVersionsSummaryDialog() {
    try {
      if (!state.parseResult) return;
      const infoObj = parseNativeVersionsInfoFromText(state.originalText || '');
      const count = Number.isFinite(state.parseResult.nativeVersionsCount) && state.parseResult.nativeVersionsCount > 0
        ? Math.floor(state.parseResult.nativeVersionsCount)
        : (Number.isFinite(infoObj?.count) ? infoObj.count : 0);
      if (!count) {
        info('No native versions detected on this macro.');
        return;
      }
      const currentSettings = Number.isFinite(infoObj?.currentSettings) ? infoObj.currentSettings : 0;
      const maxSlotIndex = Number.isFinite(infoObj?.maxSlotIndex) ? infoObj.maxSlotIndex : 0;
      const settingsEntryCount = Number.isFinite(infoObj?.settingsEntryCount) ? infoObj.settingsEntryCount : 0;
      const lines = [
        `Detected ${count} native version slot${count === 1 ? '' : 's'}.`,
      ];
      if (currentSettings > 0) lines.push(`CurrentSettings: ${currentSettings}`);
      if (maxSlotIndex > 0) lines.push(`Settings max slot index: ${maxSlotIndex} (slot count: ${maxSlotIndex + 1})`);
      if (settingsEntryCount > 1) lines.push(`Settings entry blocks: ${settingsEntryCount}`);
      lines.push('This is read-only detection for visibility right now.');
      const msg = lines.join('\n');
      try { window.alert(msg); } catch (_) { info(msg); }
    } catch (_) {
      info('Unable to inspect native versions for this macro.');
    }
  }

  function updatePublishedFeatureBadges() {
    if (!presetCountBadge && !versionsCountBadge) return;
    if (!state.parseResult) {
      if (presetCountBadge) {
        presetCountBadge.hidden = true;
        presetCountBadge.classList.remove('is-actionable');
      }
      if (versionsCountBadge) {
        versionsCountBadge.hidden = true;
        versionsCountBadge.classList.remove('is-actionable');
      }
      return;
    }
    const presetCount = countPresetPackEntries(state.parseResult);
    if (presetCountBadge) {
      if (presetCount > 0) {
        presetCountBadge.textContent = `Presets: ${presetCount}`;
        presetCountBadge.title = `${presetCount} preset${presetCount === 1 ? '' : 's'} recognized (click to open Presets Engine)`;
        presetCountBadge.classList.add('is-actionable');
        presetCountBadge.hidden = false;
      } else {
        presetCountBadge.hidden = true;
        presetCountBadge.classList.remove('is-actionable');
      }
    }
    const cachedVersions = Number(state.parseResult.nativeVersionsCount);
    const versionCount = Number.isFinite(cachedVersions) && cachedVersions > 0
      ? Math.floor(cachedVersions)
      : detectNativeVersionsCountFromText(state.originalText || '');
    if (versionsCountBadge) {
      if (versionCount > 0) {
        versionsCountBadge.textContent = `Versions: ${versionCount}`;
        versionsCountBadge.title = `${versionCount} native version slot${versionCount === 1 ? '' : 's'} recognized (click for details)`;
        versionsCountBadge.classList.add('is-actionable');
        versionsCountBadge.hidden = false;
      } else {
        versionsCountBadge.hidden = true;
        versionsCountBadge.classList.remove('is-actionable');
      }
    }
    try {
      const groupControls = Array.isArray(state.parseResult?.groupUserControls) ? state.parseResult.groupUserControls.length : 0;
      logDiag(`[Badges] preset=${presetCount}, versions=${versionCount}, cachedVersions=${Number.isFinite(cachedVersions) ? cachedVersions : 0}, groupUC=${groupControls}, hasText=${(state.originalText || '').length > 0 ? 1 : 0}, presetBadgeEl=${presetCountBadge ? 1 : 0}, versionsBadgeEl=${versionsCountBadge ? 1 : 0}`);
    } catch (_) {}
  }

  function updateDataLinkStatus() {
    if (!dataLinkStatus) return;
    if (!state.parseResult) {
      dataLinkStatus.hidden = true;
      return;
    }
    const path = resolveReloadPath();
    if (!path) {
      dataLinkStatus.textContent = 'Reload path: Not set';
      dataLinkStatus.title = '';
    } else {
      dataLinkStatus.textContent = 'Reload path: Set';
      dataLinkStatus.title = '';
    }
    dataLinkStatus.hidden = false;
  }

  function syncDataLinkPanel() {
    updatePublishedFeatureBadges();
    if (!dataLinkPanel) return;
    if (!state.parseResult) {
      dataLinkPanel.hidden = true;
      if (dataLinkStatus) dataLinkStatus.hidden = true;
      return;
    }
    const hasCsv = !!(state.csvData && Array.isArray(state.csvData.headers) && state.csvData.headers.length);
    const hasLinkSource = !!(state.parseResult && state.parseResult.dataLink && state.parseResult.dataLink.source);
    if (!hasCsv && !hasLinkSource) {
      dataLinkPanel.hidden = true;
      if (dataLinkStatus) dataLinkStatus.hidden = true;
      return;
    }
    dataLinkPanel.hidden = false;
    updateDataLinkStatus();
  }

  function getSelectedCsvBatch() {
    return documents.find(doc => doc && doc.selected && doc.isCsvBatch) || null;
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
    textPrompt.open({
      title: 'Rename macro tab',
      label: 'Tab name',
      initialValue: current,
      confirmText: 'Rename',
    }).then((nextRaw) => {
      if (nextRaw == null) return;
      const next = String(nextRaw).trim();
      if (!next || next === current) return;
      if (docId === activeDocId && state.parseResult) {
        state.parseResult.macroName = next;
        setMacroNameDisplay(next);
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
    });
  }

  function promptRenameNode(nodeInfo) {
    if (!state.parseResult || !state.originalText) {
      info('Load a macro before renaming a node.');
      return;
    }
    const current = String(nodeInfo && nodeInfo.name ? nodeInfo.name : nodeInfo || '').trim();
    if (!current) return;
    const typeLabel = nodeInfo && nodeInfo.type ? ` (${nodeInfo.type})` : '';
    textPrompt.open({
      title: `Rename node${typeLabel}`,
      label: `Node name`,
      initialValue: current,
      confirmText: 'Rename',
    }).then((raw) => {
      if (raw == null) return;
      let next = String(raw).trim();
      if (!next || next === current) return;
      next = sanitizeIdent(next);
      if (!next) return;
      if (!isIdentStart(next[0])) next = `_${next}`;
      if (next === current) return;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) {
        error('Unable to locate the macro group for renaming.');
        return;
      }
      const existing = collectToolNamesInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (existing.has(next)) {
        error(`A node named "${next}" already exists.`);
        return;
      }
      const res = renameToolInGroupText(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, current, next);
      if (!res || !res.renamed || res.text === state.originalText) {
        info('Rename skipped: no tool definition was updated.');
        return;
      }
      trackToolRename(current, next);
      pushHistory('rename node');
      state.originalText = res.text;
      updateCodeView(state.originalText || '');
      applyNodeRenameToEntries(current, next);
      applyNodeRenameToStateMaps(current, next);
      markContentDirty();
      renderActiveList();
      nodesPane?.parseAndRenderNodes?.();
      info(`Renamed ${current} to ${next}.`);
    });
  }

  function collectToolNamesInGroup(text, groupOpen, groupClose) {
    const names = new Set();
    const blocks = ['Tools', 'Modifiers'];
    blocks.forEach((blockName) => {
      const block = findOrderedBlock(text, groupOpen, groupClose, blockName);
      if (!block) return;
      const entries = parseOrderedBlockEntries(text, block.open, block.close);
      entries.forEach(entry => names.add(entry.name));
    });
    return names;
  }

  function findOrderedBlock(text, groupOpen, groupClose, blockName) {
    try {
      const groupText = text.slice(groupOpen, groupClose);
      const re = new RegExp(`${escapeRegex(blockName)}\\s*=\\s*ordered\\(\\)`, 'i');
      const match = re.exec(groupText);
      if (!match) return null;
      const matchIndex = groupOpen + match.index;
      const open = text.indexOf('{', matchIndex);
      if (open < 0 || open > groupClose) return null;
      const close = findMatchingBrace(text, open);
      if (close < 0 || close > groupClose) return null;
      return { open, close };
    } catch (_) {
      return null;
    }
  }

  function parseOrderedBlockEntries(text, open, close) {
    const inner = text.slice(open + 1, close);
    const entries = [];
    let i = 0;
    let depth = 0;
    let inStr = false;
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
      if (depth === 0 && isIdentStart(ch)) {
        const nameStart = i;
        i++;
        while (i < inner.length && isIdentPart(inner[i])) i++;
        const nameEnd = i;
        const name = inner.slice(nameStart, nameEnd);
        let j = i;
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner[j] !== '=') { i++; continue; }
        j++;
        while (j < inner.length && isSpace(inner[j])) j++;
        const typeStart = j;
        while (j < inner.length && isIdentPart(inner[j])) j++;
        if (j <= typeStart) { i++; continue; }
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner[j] !== '{') { i++; continue; }
        const absNameStart = open + 1 + nameStart;
        const absNameEnd = open + 1 + nameEnd;
        const absOpen = open + 1 + j;
        const absClose = findMatchingBrace(text, absOpen);
        entries.push({ name, nameStart: absNameStart, nameEnd: absNameEnd, blockOpen: absOpen, blockClose: absClose });
        if (absClose > 0) {
          i = absClose - (open + 1) + 1;
          continue;
        }
      }
      i++;
    }
    return entries;
  }

  function getOrderedBlockEntryType(text, entry) {
    try {
      if (!text || !entry || !Number.isFinite(entry.nameEnd) || !Number.isFinite(entry.blockOpen)) return '';
      let i = entry.nameEnd;
      while (i < entry.blockOpen && isSpace(text[i])) i++;
      if (text[i] !== '=') return '';
      i++;
      while (i < entry.blockOpen && isSpace(text[i])) i++;
      const start = i;
      while (i < entry.blockOpen && isIdentPart(text[i])) i++;
      return text.slice(start, i).trim();
    } catch (_) {
      return '';
    }
  }

  function extractOrderedBlockEntryChunk(text, entry) {
    try {
      if (!text || !entry || !Number.isFinite(entry.nameStart) || !Number.isFinite(entry.blockClose)) return '';
      let end = entry.blockClose + 1;
      let cursor = end;
      while (cursor < text.length && (text[cursor] === ' ' || text[cursor] === '\t')) cursor++;
      if (text[cursor] === ',') end = cursor + 1;
      return text.slice(entry.nameStart, end);
    } catch (_) {
      return '';
    }
  }

  function renameToolInGroupText(text, groupOpen, groupClose, oldName, newName) {
    let updated = text;
    let renamed = false;
    let currentClose = groupClose;
    const renameInBlock = (blockName) => {
      const block = findOrderedBlock(updated, groupOpen, currentClose, blockName);
      if (!block) return;
      const entries = parseOrderedBlockEntries(updated, block.open, block.close);
      const target = entries.find(entry => entry.name === oldName);
      if (!target) return;
      updated = updated.slice(0, target.nameStart) + newName + updated.slice(target.nameEnd);
      renamed = true;
      const refreshed = findMatchingBrace(updated, groupOpen);
      if (refreshed != null && refreshed >= 0) currentClose = refreshed;
    };
    renameInBlock('Tools');
    renameInBlock('Modifiers');
    if (!renamed) {
      const tryRegexRename = (blockName) => {
        const block = findOrderedBlock(updated, groupOpen, currentClose, blockName);
        if (!block) return false;
        const res = renameToolInBlockByRegex(updated, block.open, block.close, oldName, newName);
        if (!res.renamed) return false;
        updated = res.text;
        renamed = true;
        const refreshed = findMatchingBrace(updated, groupOpen);
        if (refreshed != null && refreshed >= 0) currentClose = refreshed;
        return true;
      };
      tryRegexRename('Tools');
      if (!renamed) tryRegexRename('Modifiers');
    }
    const groupEnd = (currentClose != null && currentClose >= 0) ? currentClose : groupClose;
    const beforeRefs = updated;
    updated = replaceSourceOpInRange(updated, groupOpen, groupEnd, oldName, newName);
    updated = replaceExpressionsInRange(updated, groupOpen, groupEnd, oldName, newName);
    if (!renamed && updated !== beforeRefs) {
      renamed = true;
    }
    if (!renamed) return { text, renamed: false };
    return { text: updated, renamed };
  }

  function renameToolInBlockByRegex(text, blockOpen, blockClose, oldName, newName) {
    try {
      const prefix = text.slice(0, blockOpen + 1);
      const body = text.slice(blockOpen + 1, blockClose);
      const suffix = text.slice(blockClose);
      const re = new RegExp(`(^|\\n)(\\s*)${escapeRegex(oldName)}(\\s*=\\s*[A-Za-z_][A-Za-z0-9_]*\\s*\\{)`, 'g');
      let changed = false;
      const updatedBody = body.replace(re, (match, lead, indent, rest) => {
        changed = true;
        return `${lead}${indent}${newName}${rest}`;
      });
      if (!changed) return { text, renamed: false };
      return { text: prefix + updatedBody + suffix, renamed: true };
    } catch (_) {
      return { text, renamed: false };
    }
  }

  function trackToolRename(oldName, newName) {
    try {
      if (!state.parseResult || !oldName || !newName || oldName === newName) return;
      if (!(state.parseResult.toolRenameMap instanceof Map)) {
        state.parseResult.toolRenameMap = new Map();
      }
      const map = state.parseResult.toolRenameMap;
      map.forEach((value, key) => {
        if (value === oldName) map.set(key, newName);
      });
      map.set(oldName, newName);
    } catch (_) {}
  }

  function applyToolRenameMap(text, result) {
    try {
      const map = result?.toolRenameMap;
      if (!(map instanceof Map) || !map.size) return text;
      let updated = text;
      const bounds = locateMacroGroupBounds(updated, result);
      if (!bounds) return updated;
      map.forEach((newName, oldName) => {
        if (!oldName || !newName || oldName === newName) return;
        const res = renameToolInGroupText(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, oldName, newName);
        updated = res.text;
        // Always update lingering references even if the tool name already changed.
        updated = replaceSourceOpInRange(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, oldName, newName);
        updated = replaceExpressionsInRange(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, oldName, newName);
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function stripLegacySuffix(name) {
    if (!name) return '';
    const match = String(name).match(/^(.+?)(?:_\d+)+$/);
    return match ? match[1] : String(name);
  }

  function normalizeLegacyToolRefs(text, result) {
    try {
      if (!text || !result) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      let updated = text;
      // First, normalize legacy tool names when safe (unique base + no collision).
      updated = normalizeLegacyToolDefinitions(updated, bounds);
      const toolNames = collectToolNamesInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!toolNames || toolNames.size === 0) return updated;
      updated = replaceLegacyRefsInSourceOps(updated, bounds, toolNames);
      updated = replaceLegacyRefsInExpressions(updated, bounds, toolNames);
      return updated;
    } catch (_) {
      return text;
    }
  }

  async function normalizeLegacyNamesMenu() {
    try {
      if (!state.parseResult || !state.originalText) {
        info('Load a macro before normalizing legacy names.');
        return;
      }
      const ok = window.confirm(
        'Normalize legacy tool names now?\n\nThis replaces names like "XFANIM_1_1_1" with "XFANIM" only when the base name is unique.'
      );
      if (!ok) return;
      const updated = normalizeLegacyToolRefs(state.originalText, state.parseResult);
      if (updated === state.originalText) {
        info('No legacy names found to normalize.');
        return;
      }
      await loadMacroFromText(state.originalFileName || 'Imported.setting', updated, {
        preserveFileInfo: true,
        preserveFilePath: true,
        skipClear: false,
        createDoc: false,
        allowAutoUtility: false,
        silentAuto: true,
      });
      markContentDirty();
      info('Legacy tool names normalized.');
    } catch (err) {
      error(err?.message || err || 'Failed to normalize legacy names.');
    }
  }

  function normalizeLegacyToolDefinitions(text, bounds) {
    try {
      const { groupOpenIndex: open, groupCloseIndex: close } = bounds;
      const toolNames = collectToolNamesInGroup(text, open, close);
      if (!toolNames || toolNames.size === 0) return text;
      const legacyGroups = new Map();
      toolNames.forEach((name) => {
        const base = stripLegacySuffix(name);
        if (!base || base === name) return;
        if (!legacyGroups.has(base)) legacyGroups.set(base, []);
        legacyGroups.get(base).push(name);
      });
      let updated = text;
      let currentClose = close;
      legacyGroups.forEach((names, base) => {
        if (toolNames.has(base)) return;
        if (names.length !== 1) return;
        const oldName = names[0];
        const res = renameToolInGroupText(updated, open, currentClose, oldName, base);
        updated = res.text;
        const refreshed = findMatchingBrace(updated, open);
        if (refreshed != null && refreshed >= 0) currentClose = refreshed;
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function replaceLegacyRefsInSourceOps(text, bounds, toolNames) {
    const { groupOpenIndex: open, groupCloseIndex: close } = bounds;
    const prefix = text.slice(0, open);
    const body = text.slice(open, close + 1);
    const suffix = text.slice(close + 1);
    const re = /SourceOp\s*=\s*"([^"]+)"/g;
    const updated = body.replace(re, (match, name) => {
      if (!name || toolNames.has(name)) return match;
      const base = stripLegacySuffix(name);
      if (base && base !== name && toolNames.has(base)) {
        return match.replace(name, base);
      }
      return match;
    });
    return prefix + updated + suffix;
  }

  function replaceLegacyRefsInExpressions(text, bounds, toolNames) {
    const { groupOpenIndex: open, groupCloseIndex: close } = bounds;
    const prefix = text.slice(0, open);
    const body = text.slice(open, close + 1);
    const suffix = text.slice(close + 1);
    let out = '';
    let i = 0;
    const token = 'Expression';
    const tokenRe = /\b([A-Za-z][A-Za-z0-9_]*?)(?:_\d+)+\b/g;
    while (i < body.length) {
      const idx = body.indexOf(token, i);
      if (idx < 0) {
        out += body.slice(i);
        break;
      }
      const before = idx > 0 ? body[idx - 1] : '';
      const validBoundary = !before || !isIdentPart(before);
      if (!validBoundary) {
        out += body.slice(i, idx + 1);
        i = idx + 1;
        continue;
      }
      out += body.slice(i, idx);
      let j = idx + token.length;
      while (j < body.length && isSpace(body[j])) j++;
      if (body[j] !== '=') {
        out += body.slice(idx, j + 1);
        i = j + 1;
        continue;
      }
      j++;
      while (j < body.length && isSpace(body[j])) j++;
      if (body[j] !== '"') {
        out += body.slice(idx, j + 1);
        i = j + 1;
        continue;
      }
      out += body.slice(idx, j + 1);
      j++;
      const exprStart = j;
      while (j < body.length) {
        const ch = body[j];
        if (ch === '"' && !isQuoteEscaped(body, j)) break;
        j++;
      }
      const expr = body.slice(exprStart, j);
      const next = expr.replace(tokenRe, (name) => {
        if (!name || toolNames.has(name)) return name;
        const base = stripLegacySuffix(name);
        if (base && base !== name && toolNames.has(base)) return base;
        return name;
      });
      out += next;
      if (j < body.length && body[j] === '"') {
        out += '"';
        j++;
      }
      i = j;
    }
    return prefix + out + suffix;
  }

  function replaceSourceOpInRange(text, open, close, oldName, newName) {
    const prefix = text.slice(0, open);
    const body = text.slice(open, close + 1);
    const suffix = text.slice(close + 1);
    const re = new RegExp(`(SourceOp\\s*=\\s*\")${escapeRegex(oldName)}(\")`, 'g');
    return prefix + body.replace(re, `$1${newName}$2`) + suffix;
  }

  function replaceExpressionsInRange(text, open, close, oldName, newName) {
    const prefix = text.slice(0, open);
    const body = text.slice(open, close + 1);
    const suffix = text.slice(close + 1);
    let out = '';
    let i = 0;
    const token = 'Expression';
    while (i < body.length) {
      const idx = body.indexOf(token, i);
      if (idx < 0) {
        out += body.slice(i);
        break;
      }
      const before = idx > 0 ? body[idx - 1] : '';
      const validBoundary = !before || !isIdentPart(before);
      if (!validBoundary) {
        out += body.slice(i, idx + 1);
        i = idx + 1;
        continue;
      }
      out += body.slice(i, idx);
      let j = idx + token.length;
      while (j < body.length && isSpace(body[j])) j++;
      if (body[j] !== '=') {
        out += body.slice(idx, j + 1);
        i = j + 1;
        continue;
      }
      j++;
      while (j < body.length && isSpace(body[j])) j++;
      if (body[j] !== '"') {
        out += body.slice(idx, j + 1);
        i = j + 1;
        continue;
      }
      out += body.slice(idx, j + 1);
      j++;
      const exprStart = j;
      while (j < body.length) {
        const ch = body[j];
        if (ch === '"' && !isQuoteEscaped(body, j)) break;
        j++;
      }
      const expr = body.slice(exprStart, j);
      const next = replaceNameInExpression(expr, oldName, newName);
      out += next;
      if (j < body.length && body[j] === '"') {
        out += '"';
        j++;
      }
      i = j;
    }
    return prefix + out + suffix;
  }

  function replaceNameInExpression(expr, oldName, newName) {
    try {
      if (!expr || !oldName) return expr;
      if (!expr.includes(oldName)) return expr;
      return expr.split(oldName).join(newName);
    } catch (_) {
      return expr;
    }
  }


  function applyNodeRenameToEntries(oldName, newName) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const nameExpr = new RegExp(`\\b${escapeRegex(oldName)}\\b`, 'g');
    state.parseResult.entries.forEach((entry) => {
      if (!entry || entry.sourceOp !== oldName) return;
      entry.sourceOp = newName;
      if (typeof entry.raw === 'string') {
        const re = new RegExp(`(SourceOp\\s*=\\s*\")${escapeRegex(oldName)}(\")`);
        entry.raw = entry.raw.replace(re, `$1${newName}$2`);
      }
      if (entry.displayName && entry.displayName === `${oldName}.${entry.source}`) {
        entry.displayName = `${newName}.${entry.source}`;
      }
      if (entry.displayNameOriginal && entry.displayNameOriginal === `${oldName}.${entry.source}`) {
        entry.displayNameOriginal = `${newName}.${entry.source}`;
      }
      if (entry.onChange && entry.onChange.includes(oldName)) {
        entry.onChange = entry.onChange.replace(nameExpr, newName);
      }
      if (entry.buttonExecute && entry.buttonExecute.includes(oldName)) {
        entry.buttonExecute = entry.buttonExecute.replace(nameExpr, newName);
      }
    });
  }

  function applyNodeRenameToStateMaps(oldName, newName) {
    if (!state.parseResult) return;
    const renamePrefixInSet = (set, oldPrefix, newPrefix) => {
      if (!(set instanceof Set)) return set;
      const next = new Set();
      set.forEach((key) => {
        if (typeof key === 'string' && key.startsWith(oldPrefix)) {
          next.add(newPrefix + key.slice(oldPrefix.length));
        } else {
          next.add(key);
        }
      });
      return next;
    };
    const renamePrefixInMap = (map, oldPrefix, newPrefix) => {
      if (!(map instanceof Map)) return map;
      const next = new Map();
      map.forEach((value, key) => {
        if (typeof key === 'string' && key.startsWith(oldPrefix)) {
          next.set(newPrefix + key.slice(oldPrefix.length), value);
        } else {
          next.set(key, value);
        }
      });
      return next;
    };
    const renameToolNameSet = (set) => {
      if (!(set instanceof Set)) return set;
      const next = new Set();
      set.forEach((name) => {
        next.add(name === oldName ? newName : name);
      });
      return next;
    };
    const renameNodeSelection = (set) => {
      if (!(set instanceof Set)) return set;
      const next = new Set();
      set.forEach((key) => {
        if (typeof key !== 'string') { next.add(key); return; }
        const parts = key.split('|');
        if (parts.length >= 3 && parts[1] === oldName) {
          parts[1] = newName;
          next.add(parts.join('|'));
        } else {
          next.add(key);
        }
      });
      return next;
    };
    const renameEntryKeySet = (set) => renamePrefixInSet(set, `${oldName}::`, `${newName}::`);
    state.parseResult.nodesCollapsed = renameToolNameSet(state.parseResult.nodesCollapsed);
    state.parseResult.nodesPublishedOnly = renameToolNameSet(state.parseResult.nodesPublishedOnly);
    state.parseResult.nodeSelection = renameNodeSelection(state.parseResult.nodeSelection);
    state.parseResult.insertClickedKeys = renamePrefixInSet(state.parseResult.insertClickedKeys, `${oldName}.`, `${newName}.`);
    state.parseResult.buttonExactInsert = renamePrefixInSet(state.parseResult.buttonExactInsert, `${oldName}.`, `${newName}.`);
    state.parseResult.insertUrlMap = renamePrefixInMap(state.parseResult.insertUrlMap, `${oldName}.`, `${newName}.`);
    state.parseResult.buttonOverrides = renamePrefixInMap(state.parseResult.buttonOverrides, `${oldName}.`, `${newName}.`);
    state.parseResult.blendToggles = renameEntryKeySet(state.parseResult.blendToggles);
    if (state.parseResult.luaToolNames instanceof Set) {
      const toolNames = new Set();
      state.parseResult.luaToolNames.forEach((name) => {
        toolNames.add(name === oldName ? newName : name);
      });
      state.parseResult.luaToolNames = toolNames;
    }
    if (pendingControlMeta.size) {
      const next = new Map();
      pendingControlMeta.forEach((value, key) => {
        if (typeof key === 'string' && key.startsWith(`${oldName}::`)) {
          next.set(`${newName}::${key.slice(oldName.length + 2)}`, value);
        } else {
          next.set(key, value);
        }
      });
      pendingControlMeta.clear();
      next.forEach((value, key) => pendingControlMeta.set(key, value));
    }
    if (pendingOpenNodes.size) {
      const next = new Set();
      pendingOpenNodes.forEach((name) => {
        next.add(name === oldName ? newName : name);
      });
      pendingOpenNodes.clear();
      next.forEach((name) => pendingOpenNodes.add(name));
    }
  }

  function updateExportMenuButtonState() {
    exportMenuController?.setEnabled?.(!!state.parseResult);
  }


  function updateDataMenuState() {
    if (!nativeBridge || !nativeBridge.ipcRenderer) return;
    const source = state.csvData?.sourceName || state.parseResult?.dataLink?.source || '';
    const reloadEnabled = /^https?:/i.test(String(source || ''));
    try {
      nativeBridge.ipcRenderer.invoke('set-data-menu-state', { reloadEnabled });
    } catch (_) {}
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

  const EXPORT_PRIMARY_MODE_DEFAULT = 'file';
  const EXPORT_PRIMARY_MODE_LABELS = Object.freeze({
    clipboard: 'Export to Clipboard',
    file: 'Export to File',
    edit: 'Export to Edit Page',
    source: 'Export to Source',
    'source-bulk': 'Export to Source DRFX',
    drfx: 'Export to DRFX',
  });
  let exportPrimaryMode = EXPORT_PRIMARY_MODE_DEFAULT;

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

  function getExportPrimaryLabel(modeId) {
    const key = String(modeId || '').trim();
    return EXPORT_PRIMARY_MODE_LABELS[key] || EXPORT_PRIMARY_MODE_LABELS[EXPORT_PRIMARY_MODE_DEFAULT];
  }

  function resolveCurrentPrimaryExportItem() {
    try {
      const items = buildExportMenuItems()
        .filter((item) => item && item.id && item.type !== 'sep' && typeof item.action === 'function');
      if (!items.length) return null;
      const byId = new Map(items.map((item) => [String(item.id || '').trim(), item]));
      const currentId = String(exportPrimaryMode || '').trim();
      const currentItem = byId.get(currentId);
      if (currentItem && currentItem.enabled) return currentItem;
      const fileItem = byId.get(EXPORT_PRIMARY_MODE_DEFAULT);
      if (currentId !== EXPORT_PRIMARY_MODE_DEFAULT && fileItem && fileItem.enabled) {
        exportPrimaryMode = EXPORT_PRIMARY_MODE_DEFAULT;
        return fileItem;
      }
      // Keep existing mode on startup/disabled states instead of drifting to first menu item.
      if (currentItem) return currentItem;
      if (fileItem) return fileItem;
      return items[0] || null;
    } catch (_) {
      return null;
    }
  }

  function syncPrimaryExportButtonLabel() {
    const current = resolveCurrentPrimaryExportItem();
    const label = getExportPrimaryLabel(current?.id || exportPrimaryMode);
    if (exportMenuBtn) exportMenuBtn.textContent = label;
  }

  function setPrimaryExportMode(modeId) {
    const key = String(modeId || '').trim();
    if (!key) return;
    exportPrimaryMode = key;
    syncPrimaryExportButtonLabel();
  }

  exportMenuController = createExportMenuController({
    button: exportMenuBtn,
    toggleButton: exportMenuCaretBtn,
    menu: exportMenuEl,
    getItems: buildExportMenuItems,
    onItemInvoke: (item) => {
      if (!item || item.type === 'sep' || !item.id) return;
      setPrimaryExportMode(item.id);
    },
  });
  syncPrimaryExportButtonLabel();

  function requestMiniExplorerCreateFolder() {
    try {
      if (!macroExplorerMiniFrame || !macroExplorerMiniFrame.contentWindow) {
        info('Mini explorer is not ready yet.');
        return;
      }
      macroExplorerMiniFrame.contentWindow.postMessage({ type: 'fmr-mini-create-folder' }, '*');
    } catch (err) {
      info(`Unable to trigger folder create: ${err?.message || err}`);
    }
  }

  miniCreateFolderBtn?.addEventListener('click', () => requestMiniExplorerCreateFolder());

  function updateIntroToggleVisibility() {
    setIntroToggleVisible(!!state.parseResult);
  }

  createGlobalShortcuts({
    isElectron: IS_ELECTRON,
    getDocuments: () => documents,
    getActiveDocId: () => activeDocId,
    isInteractiveTarget,
    onSwitchByIndex: (idx) => switchDocumentByIndex(idx),
    onSwitchByOffset: (offset) => switchDocumentByOffset(offset),
    onCloseDoc: (docId) => closeDocument(docId),
    onCreateDoc: () => createBlankDocument(),
  });

  const controlsSection = document.getElementById('controlsSection');

  const controlsList = document.getElementById('controlsList');

  const publishedSearch = document.getElementById('publishedSearch');
  const showCurrentValuesToggle = document.getElementById('showCurrentValues');
  const pageTabsEl = document.getElementById('pageTabs');

  // Nodes pane elements

  const nodesList = document.getElementById('nodesList');
  const nodesPaneEl = document.getElementById('nodesPane');
  const routingBtn = document.getElementById('routingBtn');
  const hideNodesBtn = document.getElementById('hideNodesBtn');
  const showNodesBtn = document.getElementById('showNodesBtn');

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
  const addComboOptions = document.getElementById('addComboOptions');
  const addControlLabelCountInput = document.getElementById('addControlLabelCount');
  const addControlLabelDefaultSelect = document.getElementById('addControlLabelDefault');
  const addControlComboOptionsInput = document.getElementById('addControlComboOptions');
  const addControlPageInput = document.getElementById('addControlPage');
  const addControlPageOptions = document.getElementById('addControlPageOptions');
  const addControlTargetSelect = document.getElementById('addControlTarget');
  const addControlError = document.getElementById('addControlError');
  const addControlCancelBtn = document.getElementById('addControlCancelBtn');
  const addControlCloseBtn = document.getElementById('addControlCloseBtn');
  const addControlSubmitBtn = document.getElementById('addControlSubmitBtn');
  let nodesPane = null;
  let historyController = null;
  let detailDrawerCollapsed = false;
  const DRAWER_MODE = {
    HIDDEN: 'hidden',
    COLLAPSED: 'collapsed',
    MACRO: 'macro',
    CONTROL: 'control',
  };
  let drawerMode = DRAWER_MODE.HIDDEN;
  let codePaneCollapsed = true;
  let codePaneRefreshTimer = null;
  let codeHighlightActive = false;
  let activeDetailEntryIndex = null;
  let renderIconHtml = () => '';
  let showCurrentValues = !!(showCurrentValuesToggle && showCurrentValuesToggle.checked);
  const DETAIL_DRAWER_COLLAPSED_TITLE = 'Details';
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
  let nodesHidden = false;

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
        extras.routingInputs = Array.isArray(state.parseResult.routingInputs)
          ? state.parseResult.routingInputs.map((row) => ({ ...(row || {}) }))
          : [];
        extras.routingManagedNames = Array.from(state.parseResult.routingManagedNames || []);
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
        state.parseResult.routingInputs = Array.isArray(extra.routingInputs)
          ? extra.routingInputs.map((row) => ({ ...(row || {}) }))
          : [];
        state.parseResult.routingManagedNames = new Set(
          Array.isArray(extra.routingManagedNames)
            ? extra.routingManagedNames
            : state.parseResult.routingInputs.map((row) => row?.macroInput).filter(Boolean)
        );
      }
      setMacroNameDisplay(state.parseResult?.macroName || 'Unknown');
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

  setNodesHidden(false);
  hideNodesBtn?.addEventListener('click', () => setNodesHidden(true));
  showNodesBtn?.addEventListener('click', () => setNodesHidden(false));
  routingBtn?.addEventListener('click', () => openRoutingIOModal());


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
  const addControlGlobalBtn = document.getElementById('addControlGlobalBtn');
  const addControlQuickBtn = document.getElementById('addControlQuickBtn');
  const addControlQuickMenu = document.getElementById('addControlQuickMenu');

  const hideReplacedEl = document.getElementById('hideReplaced');
  const quickClickHintEl = document.getElementById('quickClickHint');
  const viewControlsBtn = document.getElementById('viewControlsBtn');
  const viewControlsMenu = document.getElementById('viewControlsMenu');
  const showAllNodesBtn = document.getElementById('showAllNodesBtn');
  const showPublishedNodesBtn = document.getElementById('showPublishedNodesBtn');
  const collapseAllNodesBtn = document.getElementById('collapseAllNodesBtn');
  const showNodeTypeLabelsEl = document.getElementById('showNodeTypeLabels');
  const showInstancedConnectionsEl = document.getElementById('showInstancedConnections');
  const showNextNodeLinksEl = document.getElementById('showNextNodeLinks');
  const autoGroupQuickSetsEl = document.getElementById('autoGroupQuickSets');
  const nameClickQuickSetEl = document.getElementById('nameClickQuickSet');
  const quickSetBlendToCheckboxEl = document.getElementById('quickSetBlendToCheckbox');

  const macroNameEl = document.getElementById('macroName');
  const setMacroNameDisplay = (value) => {
    if (!macroNameEl) return;
    const next = (value && String(value).trim()) ? String(value).trim() : 'Unknown';
    if ('value' in macroNameEl) {
      macroNameEl.value = next;
    } else {
      macroNameEl.textContent = next;
    }
  };
  const getMacroNameDisplay = () => {
    if (!macroNameEl) return '';
    if ('value' in macroNameEl) return macroNameEl.value || '';
    return macroNameEl.textContent || '';
  };

  const setCreateControlActionsEnabled = (enabled) => {
    const on = !!enabled;
    if (addControlGlobalBtn) addControlGlobalBtn.disabled = !on;
    if (addControlQuickBtn) addControlQuickBtn.disabled = !on;
    if (!on) closeAddControlQuickMenu();
  };

  const ctrlCountEl = document.getElementById('ctrlCount');
  const inputCountEl = document.getElementById('inputCount');
  const csvBatchBanner = document.getElementById('csvBatchBanner');
  const dataLinkPanel = document.getElementById('dataLinkPanel');
  const dataLinkConfigure = document.getElementById('dataLinkConfigure');
  const dataLinkReload = document.getElementById('dataLinkReload');
  const dataLinkStatus = document.getElementById('dataLinkStatus');
  const presetCountBadge = document.getElementById('presetCountBadge');
  const versionsCountBadge = document.getElementById('versionsCountBadge');
  const operatorTypeSelect = document.getElementById('operatorType');

  if (dataLinkPanel) dataLinkPanel.hidden = true;
  if (dataLinkStatus) dataLinkStatus.hidden = true;

  dataLinkConfigure?.addEventListener('click', () => openDataLinkConfigModal());
  dataLinkReload?.addEventListener('click', () => reloadDataLinkForCurrentMacro());
  presetCountBadge?.addEventListener('click', () => {
    if (!state.parseResult) return;
    openPresetEngineModal();
  });
  versionsCountBadge?.addEventListener('click', () => {
    if (!state.parseResult) return;
    openNativeVersionsSummaryDialog();
  });

  const exportBtn = document.getElementById('exportBtn');
  const exportTemplatesBtn = document.getElementById('exportTemplatesBtn');
  ensureDefaultMacroExplorerFolders();

  const exportClipboardBtn = document.getElementById('exportClipboardBtn');
  const importClipboardBtn = document.getElementById('importClipboardBtn');
  const openNativeBtn = document.getElementById('openNativeBtn');
  const legacyFileLabel = document.querySelector('label[for="fileInput"]');

  function configureIntroImportActions() {
    try {
      const native = getNativeApi();
      const canNativeOpen = !!(native && native.isElectron && typeof native.openSettingFile === 'function');
      if (openNativeBtn) {
        openNativeBtn.hidden = false;
        openNativeBtn.textContent = canNativeOpen ? 'Open .setting' : 'Import from file';
        openNativeBtn.title = canNativeOpen ? 'Open .setting file' : 'Import .setting file';
      }
      if (legacyFileLabel) legacyFileLabel.style.display = '';
      if (fileInput) fileInput.style.display = 'none';
    } catch (_) {}
  }
  configureIntroImportActions();

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

  const addControlModalController = createAddControlModal({
    state,
    addControlModal,
    addControlForm,
    addControlTitle,
    addControlNameInput,
    addControlTypeSelect,
    addLabelOptions,
    addComboOptions,
    addControlLabelCountInput,
    addControlLabelDefaultSelect,
    addControlComboOptionsInput,
    addControlPageInput,
    addControlPageOptions,
    addControlTargetSelect,
    addControlError,
    addControlCancelBtn,
    addControlCloseBtn,
    addControlSubmitBtn,
    onAddControl: (nodeName, config) => addControlToNode(nodeName, applyLabelSelectionDefaults(config)),
    getTargetNodes: () => getControlTargetNodeNames(),
    getSuggestedLabelCount: () => {
      const span = getSuggestedLabelSelectionSpan();
      return Number.isFinite(span?.count) ? span.count : 0;
    },
    getSuggestedLabelSelectionSpan: () => getSuggestedLabelSelectionSpan(),
  });

  function pickPrimaryToolNameFromCurrentMacro() {
    try {
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) return '';
      const toolsBlock = findOrderedBlock(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
      if (!toolsBlock) return '';
      const entries = parseOrderedBlockEntries(state.originalText, toolsBlock.open, toolsBlock.close);
      return entries && entries.length ? entries[0].name : '';
    } catch (_) { return ''; }
  }

  function getControlTargetNodeNames() {
    try {
      const names = nodesPane?.getNodeNames?.() || [];
      if (Array.isArray(names) && names.length) return names;
    } catch (_) {}
    const primary = pickPrimaryToolNameFromCurrentMacro();
    return primary ? [primary] : [];
  }

  function getSuggestedLabelSelectionSpan() {
    try {
      return typeof getPublishedSelectionSpan === 'function' ? getPublishedSelectionSpan() : null;
    } catch (_) {
      return null;
    }
  }

  function applyLabelSelectionDefaults(config) {
    try {
      const next = { ...(config || {}) };
      if (String(next.type || '').toLowerCase() !== 'label') return next;
      const span = getSuggestedLabelSelectionSpan();
      if (!span || !Number.isFinite(span.count) || span.count <= 0) return next;
      if (!Number.isFinite(next.labelCount) || next.labelCount <= 0) {
        next.labelCount = span.count;
      }
      next.labelSelectionSpan = span;
      return next;
    } catch (_) {
      return config || {};
    }
  }

  function openAddControlDialog(nodeName) {
    addControlModalController?.open?.(nodeName);
  }

  function closeAddControlQuickMenu() {
    if (addControlQuickMenu) addControlQuickMenu.hidden = true;
    if (addControlQuickBtn) addControlQuickBtn.setAttribute('aria-expanded', 'false');
  }

  function toggleAddControlQuickMenu() {
    if (!addControlQuickMenu || !addControlQuickBtn || addControlQuickBtn.disabled) return;
    const willOpen = !!addControlQuickMenu.hidden;
    addControlQuickMenu.hidden = !willOpen;
    addControlQuickBtn.setAttribute('aria-expanded', willOpen ? 'true' : 'false');
  }

  async function quickCreateControl(type) {
    const safeType = String(type || '').toLowerCase();
    if (!['label', 'slider', 'combo', 'button', 'separator'].includes(safeType)) return;
    if (!state.parseResult || !state.originalText) {
      error('Load a macro before creating controls.');
      return;
    }
    const targets = getControlTargetNodeNames();
    const nodeName = targets && targets.length ? targets[0] : '';
    if (!nodeName) {
      error('No target node available for control creation.');
      return;
    }
    const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
    const config = {
      name: safeType === 'label'
        ? 'Label'
        : safeType === 'slider'
          ? 'Slider'
          : safeType === 'combo'
            ? 'Combo'
          : safeType === 'button'
            ? 'Button'
            : 'Separator',
      type: safeType,
      page: activePage || 'Controls',
    };
    if (safeType === 'label') {
      config.labelCount = 0;
      config.labelDefault = 'closed';
    }
    try {
      await addControlToNode(nodeName, applyLabelSelectionDefaults(config));
    } catch (err) {
      error(err?.message || 'Unable to create control.');
    }
  }

  addControlGlobalBtn?.addEventListener('click', () => {
    closeAddControlQuickMenu();
    openAddControlDialog();
  });
  addControlQuickBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    toggleAddControlQuickMenu();
  });
  addControlQuickMenu?.addEventListener('click', async (ev) => {
    const btn = ev.target && ev.target.closest ? ev.target.closest('button[data-quick-type]') : null;
    if (!btn) return;
    ev.preventDefault();
    ev.stopPropagation();
    closeAddControlQuickMenu();
    await quickCreateControl(btn.dataset.quickType || '');
  });
  document.addEventListener('mousedown', (ev) => {
    if (!addControlQuickMenu || !addControlQuickBtn) return;
    if (addControlQuickMenu.hidden) return;
    const target = ev.target;
    const inMenu = addControlQuickMenu.contains(target);
    const inBtn = addControlQuickBtn.contains(target);
    if (!inMenu && !inBtn) closeAddControlQuickMenu();
  });
  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') closeAddControlQuickMenu();
  });

function hideDetailDrawer() {
    activeDetailEntryIndex = null;
    const syncDrawerBodyClasses = () => {
      try {
        document.body.classList.remove('fmr-drawer-open', 'fmr-drawer-collapsed');
        if (!detailDrawer || detailDrawer.hidden) return;
        if (detailDrawerCollapsed) document.body.classList.add('fmr-drawer-collapsed');
        else document.body.classList.add('fmr-drawer-open');
      } catch (_) {}
    };
    if (!detailDrawer) return;
    if (!state.parseResult || controlsSection?.hidden) {
      detailDrawer.hidden = true;
      detailDrawer.classList.remove('open');
      detailDrawer.classList.remove('collapsed');
      detailDrawerCollapsed = false;
      drawerMode = DRAWER_MODE.HIDDEN;
      syncDrawerBodyClasses();
      return;
    }
    collapseDetailDrawer();
    syncDrawerBodyClasses();
  }

  function applyCollapsedDetailDrawerTitle() {
    if (!detailDrawerTitle) return;
    detailDrawerTitle.textContent = DETAIL_DRAWER_COLLAPSED_TITLE;
    detailDrawerTitle.contentEditable = 'false';
    detailDrawerTitle.spellcheck = false;
    detailDrawerTitle.onblur = null;
    detailDrawerTitle.onkeydown = null;
  }

  function updateDetailDrawerToggleVisibility() {
    if (!detailDrawerToggleBtn) return;
    const shouldHide = !!((detailDrawer && detailDrawer.hidden) || (nodesHidden && !detailDrawerCollapsed));
    detailDrawerToggleBtn.hidden = shouldHide;
  }

  function syncDetailDrawerAnchor() {
    if (!detailDrawer) return;
    if (!detailDrawerCollapsed || !state.parseResult || controlsSection?.hidden) {
      detailDrawer.style.top = '';
      detailDrawer.style.right = '';
      return;
    }
    const paneLeft = document.querySelector('.controls .pane-left');
    if (!paneLeft || typeof paneLeft.getBoundingClientRect !== 'function') {
      detailDrawer.style.top = '';
      detailDrawer.style.right = '';
      return;
    }
    const rect = paneLeft.getBoundingClientRect();
    if (!Number.isFinite(rect.top) || !Number.isFinite(rect.right)) return;
    const rightOffset = Math.max(12, Math.round(window.innerWidth - rect.right));
    const topOffset = Math.max(80, Math.round(rect.top + 12));
    detailDrawer.style.right = `${rightOffset}px`;
    detailDrawer.style.top = `${topOffset}px`;
  }

  function applyDrawerState(expanded) {
    if (!detailDrawer) return;
    const syncDrawerBodyClasses = () => {
      try {
        document.body.classList.remove('fmr-drawer-open', 'fmr-drawer-collapsed');
        if (!detailDrawer || detailDrawer.hidden) return;
        if (detailDrawerCollapsed) document.body.classList.add('fmr-drawer-collapsed');
        else document.body.classList.add('fmr-drawer-open');
      } catch (_) {}
    };
    detailDrawer.hidden = false;
    detailDrawerCollapsed = !expanded;
    detailDrawer.classList.toggle('open', expanded);
    detailDrawer.classList.toggle('collapsed', !expanded);
    syncDrawerBodyClasses();
    updateDrawerToggleLabel();
    updateDetailDrawerToggleVisibility();
    syncDetailDrawerAnchor();
  }

  function collapseDetailDrawer() {
    if (!detailDrawer) return;
    const syncDrawerBodyClasses = () => {
      try {
        document.body.classList.remove('fmr-drawer-open', 'fmr-drawer-collapsed');
        if (!detailDrawer || detailDrawer.hidden) return;
        if (detailDrawerCollapsed) document.body.classList.add('fmr-drawer-collapsed');
        else document.body.classList.add('fmr-drawer-open');
      } catch (_) {}
    };
    if (!state.parseResult || controlsSection?.hidden) {
      detailDrawer.hidden = true;
      detailDrawer.classList.remove('open');
      detailDrawer.classList.remove('collapsed');
      detailDrawerCollapsed = false;
      drawerMode = DRAWER_MODE.HIDDEN;
      syncDrawerBodyClasses();
      return;
    }
    drawerMode = DRAWER_MODE.COLLAPSED;
    applyDrawerState(false);
    applyCollapsedDetailDrawerTitle();
    syncDrawerBodyClasses();
  }

  function renderMacroInfoDrawer(options = {}) {
    if (!detailDrawer || !detailDrawerTitle || !detailDrawerSubtitle || !detailDrawerBody) return;
    const expanded = !!options.expanded;
    drawerMode = expanded ? DRAWER_MODE.MACRO : DRAWER_MODE.COLLAPSED;
    applyDrawerState(expanded);
    const fileLabel = state.originalFileName || 'No file loaded';
    detailDrawerSubtitle.textContent = fileLabel;
    detailDrawerBody.innerHTML = '';
    if (!expanded) {
      applyCollapsedDetailDrawerTitle();
    } else if (!state.parseResult) {
      detailDrawerTitle.textContent = 'Macro Info';
      detailDrawerTitle.contentEditable = 'false';
      detailDrawerTitle.spellcheck = false;
      detailDrawerTitle.onblur = null;
      detailDrawerTitle.onkeydown = null;
    }
    if (!state.parseResult) {
      const empty = document.createElement('p');
      empty.className = 'detail-placeholder';
      empty.textContent = 'Import a .setting to see macro details.';
      detailDrawerBody.appendChild(empty);
      return;
    }
    const doc = getActiveDocument && getActiveDocument();
    const snap = doc && doc.snapshot ? doc.snapshot : null;
    const macroName = state.parseResult.macroName || snap?.parseResult?.macroName || '';
    const operatorType = state.parseResult.operatorType || snap?.parseResult?.operatorType || '';
    if (expanded) {
      const fallbackName = state.parseResult.macroNameOriginal || macroName || 'Macro Info';
      detailDrawerTitle.textContent = macroName || fallbackName;
      detailDrawerTitle.contentEditable = 'true';
      detailDrawerTitle.spellcheck = false;
      const commitMacroName = () => {
        if (!state.parseResult) return;
        const newName = (detailDrawerTitle.textContent || '').trim();
        const prevName = state.parseResult.macroName || '';
        if (newName) {
          state.parseResult.macroName = newName;
          setMacroNameDisplay(newName);
          updateActiveDocumentMeta({ name: newName });
          if (newName !== prevName) markActiveDocumentDirty();
        } else {
          const fallback = state.parseResult.macroNameOriginal || prevName || 'Unknown';
          detailDrawerTitle.textContent = fallback;
          setMacroNameDisplay(fallback);
          updateActiveDocumentMeta({ name: fallback });
          if (fallback !== prevName) markActiveDocumentDirty();
        }
      };
      detailDrawerTitle.onblur = commitMacroName;
      detailDrawerTitle.onkeydown = (ev) => {
        if (ev.key === 'Enter') { ev.preventDefault(); commitMacroName(); detailDrawerTitle.blur(); }
        if (ev.key === 'Escape') {
          ev.preventDefault();
          detailDrawerTitle.textContent = state.parseResult.macroName || state.parseResult.macroNameOriginal || 'Macro Info';
          detailDrawerTitle.blur();
        }
      };
    }
    const addInfoField = (label, value) => {
      const field = document.createElement('div');
      field.className = 'detail-field';
      const lbl = document.createElement('label');
      lbl.textContent = label;
      const val = document.createElement('div');
      val.className = 'detail-info-value';
      if (!value) {
        val.textContent = '--';
        val.classList.add('muted');
      } else {
        val.textContent = value;
      }
      field.appendChild(lbl);
      field.appendChild(val);
      detailDrawerBody.appendChild(field);
    };
    const addOperatorField = () => {
      const field = document.createElement('div');
      field.className = 'detail-field';
      const lbl = document.createElement('label');
      lbl.textContent = 'Operator type';
      const select = document.createElement('select');
      select.className = 'detail-type-select';
      if (operatorTypeSelect && operatorTypeSelect.options.length) {
        Array.from(operatorTypeSelect.options).forEach((opt) => {
          const clone = document.createElement('option');
          clone.value = opt.value;
          clone.textContent = opt.textContent;
          select.appendChild(clone);
        });
      } else {
        ['GroupOperator', 'MacroOperator'].forEach((value) => {
          const opt = document.createElement('option');
          opt.value = value;
          opt.textContent = value;
          select.appendChild(opt);
        });
      }
      select.value = operatorType === 'MacroOperator' ? 'MacroOperator' : 'GroupOperator';
      select.addEventListener('change', () => {
        if (!state.parseResult) return;
        const next = select.value === 'MacroOperator' ? 'MacroOperator' : 'GroupOperator';
        const prev = state.parseResult.operatorType || '';
        state.parseResult.operatorType = next;
        if (operatorTypeSelect) operatorTypeSelect.value = next;
        if (next !== prev) markActiveDocumentDirty();
      });
      field.appendChild(lbl);
      field.appendChild(select);
      detailDrawerBody.appendChild(field);
    };
    const controlsCount = state.parseResult.entries ? String(state.parseResult.entries.length) : '';
    addOperatorField();
    addInfoField('Controls found', controlsCount);
    addInfoField('Inputs found', String(countRecognizedInputs(state.parseResult.entries || [])));
    const link = state.drfxLink || snap?.drfxLink || null;
    if (link && link.linked) {
      addInfoField('DRFX preset', link.presetName || 'Linked');
      addInfoField('DRFX pack', link.drfxPath || '');
    } else {
      addInfoField('DRFX link', 'Not linked');
    }
  }

  function refreshDetailDrawerState() {
    if (!detailDrawer) return;
    if (!state.parseResult || controlsSection?.hidden) {
      detailDrawer.hidden = true;
      detailDrawer.classList.remove('open');
      detailDrawer.classList.remove('collapsed');
      drawerMode = DRAWER_MODE.HIDDEN;
      return;
    }
    if (activeDetailEntryIndex != null) {
      renderDetailDrawer(activeDetailEntryIndex);
      return;
    }
    if (nodesHidden) {
      renderMacroInfoDrawer({ expanded: true });
      return;
    }
    if (drawerMode === DRAWER_MODE.MACRO) {
      renderMacroInfoDrawer({ expanded: true });
      return;
    }
    collapseDetailDrawer();
  }

  function isLuaIdentifier(value) {
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(String(value || ''));
  }

  function extractToolNamesAndTypesFromSetting(text, result) {
    const names = new Set();
    const types = new Map();
    try {
      if (!text) return { names, types };
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return { names, types };
      const blocks = ['Tools', 'Modifiers'];
      const scanBlock = (blockName) => {
        const blockPos = text.indexOf(`${blockName} = ordered()`, bounds.groupOpenIndex);
        if (blockPos < 0 || blockPos > bounds.groupCloseIndex) return;
        const open = text.indexOf('{', blockPos);
        if (open < 0) return;
        const close = findMatchingBrace(text, open);
        if (close < 0 || close > bounds.groupCloseIndex) return;
        const inner = text.slice(open + 1, close);
        let i = 0;
        let depth = 0;
        let inStr = false;
        while (i < inner.length) {
          const ch = inner[i];
          if (inStr) { if (ch === '"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
          if (ch === '"') { inStr = true; i++; continue; }
          if (ch === '{') { depth++; i++; continue; }
          if (ch === '}') { depth--; i++; continue; }
          if (depth === 0 && isIdentStart(ch)) {
            const nameStart = i; i++;
            while (i < inner.length && isIdentPart(inner[i])) i++;
            const toolName = inner.slice(nameStart, i).trim();
            if (isLuaIdentifier(toolName)) names.add(toolName);
            while (i < inner.length && isSpace(inner[i])) i++;
            if (inner[i] !== '=') { i++; continue; }
            i++;
            while (i < inner.length && isSpace(inner[i])) i++;
            const typeStart = i;
            while (i < inner.length && isIdentPart(inner[i])) i++;
            const toolType = inner.slice(typeStart, i).trim();
            if (toolName && toolType) types.set(toolName, toolType);
            while (i < inner.length && isSpace(inner[i])) i++;
            if (inner[i] !== '{') { i++; continue; }
            const tOpen = i;
            const tClose = findMatchingBrace(inner, tOpen);
            if (tClose < 0) break;
            i = tClose + 1;
            continue;
          }
          i++;
        }
      };
      blocks.forEach(scanBlock);
      return { names, types };
    } catch (_) {
      return { names, types };
    }
  }

  function extractToolNamesFromSetting(text, result) {
    return extractToolNamesAndTypesFromSetting(text, result).names;
  }

  function extractToolTypesFromSetting(text, result) {
    return extractToolNamesAndTypesFromSetting(text, result).types;
  }

  function createLuaNameOverlay(getToolNames, getControlNames) {
    return {
      token(stream) {
        if (stream.match('--')) {
          stream.skipToEnd();
          return null;
        }
        const ch = stream.peek();
        if (ch === '"' || ch === "'") {
          const quote = stream.next();
          let escaped = false;
          while (!stream.eol()) {
            const c = stream.next();
            if (escaped) {
              escaped = false;
              continue;
            }
            if (c === '\\') {
              escaped = true;
              continue;
            }
            if (c === quote) break;
          }
          return null;
        }
        if (stream.match('[[', false)) {
          stream.match('[[', true);
          if (!stream.skipTo(']]')) {
            stream.skipToEnd();
            return null;
          }
          stream.match(']]', true);
          return null;
        }
        if (stream.match(/^[A-Za-z_][A-Za-z0-9_]*/)) {
          const word = stream.current();
          const tools = typeof getToolNames === 'function' ? getToolNames() : getToolNames;
          if (tools && tools.has && tools.has(word)) return 'lua-tool';
          const controls = typeof getControlNames === 'function' ? getControlNames() : getControlNames;
          if (controls && controls.has && controls.has(word)) return 'lua-control';
          return null;
        }
        stream.next();
        return null;
      },
    };
  }

  function validateLuaBasic(source) {
    const warnings = [];
    const text = String(source || '');
    if (!text.trim()) return warnings;
    let i = 0;
    let quote = null;
    let inLong = false;
    let longType = null;
    let paren = 0;
    let brace = 0;
    let bracket = 0;
    const blockStack = [];
    const pushBlock = (token) => blockStack.push(token);
    const popBlock = (token) => {
      if (!blockStack.length) {
        warnings.push(`Unexpected "${token}"`);
        return;
      }
      const last = blockStack[blockStack.length - 1];
      if (token === 'until') {
        if (last === 'repeat') blockStack.pop();
        else warnings.push('Found "until" without matching "repeat"');
        return;
      }
      if (last === 'repeat') {
        warnings.push('Found "end" but the last block is "repeat" (use "until")');
        blockStack.pop();
        return;
      }
      blockStack.pop();
    };
    while (i < text.length) {
      const ch = text[i];
      const next = text[i + 1];
      if (inLong) {
        if (ch === ']' && next === ']') {
          inLong = false;
          longType = null;
          i += 2;
          continue;
        }
        i += 1;
        continue;
      }
      if (quote) {
        if (ch === '\\') {
          i += 2;
          continue;
        }
        if (ch === quote) {
          quote = null;
          i += 1;
          continue;
        }
        i += 1;
        continue;
      }
      if (ch === '-' && next === '-') {
        if (text[i + 2] === '[' && text[i + 3] === '[') {
          inLong = true;
          longType = 'comment';
          i += 4;
          continue;
        }
        while (i < text.length && text[i] !== '\n') i++;
        continue;
      }
      if (ch === '[' && next === '[') {
        inLong = true;
        longType = 'string';
        i += 2;
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        i += 1;
        continue;
      }
      if (ch === '(') paren++;
      else if (ch === ')') paren--;
      else if (ch === '{') brace++;
      else if (ch === '}') brace--;
      else if (ch === '[') bracket++;
      else if (ch === ']') bracket--;
      if (paren < 0) { warnings.push('Extra ")"'); paren = 0; }
      if (brace < 0) { warnings.push('Extra "}"'); brace = 0; }
      if (bracket < 0) { warnings.push('Extra "]"'); bracket = 0; }
      if (isIdentStart(ch)) {
        let j = i + 1;
        while (j < text.length && isIdentPart(text[j])) j++;
        const word = text.slice(i, j);
        if (word === 'function' || word === 'if' || word === 'do' || word === 'repeat') {
          pushBlock(word);
        } else if (word === 'end') {
          popBlock('end');
        } else if (word === 'until') {
          popBlock('until');
        }
        i = j;
        continue;
      }
      i += 1;
    }
    if (quote) warnings.push('Unclosed string literal');
    if (inLong) warnings.push(`Unclosed long ${longType || 'string'}`);
    if (paren > 0) warnings.push('Unclosed ")"');
    if (brace > 0) warnings.push('Unclosed "}"');
    if (bracket > 0) warnings.push('Unclosed "]"');
    if (blockStack.length) warnings.push('Missing "end" or "until"');
    return warnings.slice(0, 3);
  }

  function renderLuaWarnings(container, warnings) {
    if (!container) return;
    const items = Array.isArray(warnings) ? warnings.filter(Boolean) : [];
    container.innerHTML = '';
    if (!items.length) {
      container.hidden = true;
      return;
    }
    container.hidden = false;
    items.forEach((msg) => {
      const row = document.createElement('div');
      row.className = 'lua-warning';
      row.textContent = msg;
      container.appendChild(row);
    });
  }

  function getLuaWordBounds(line, ch) {
    let start = ch;
    while (start > 0 && /[A-Za-z0-9_]/.test(line[start - 1])) start--;
    let end = ch;
    while (end < line.length && /[A-Za-z0-9_]/.test(line[end])) end++;
    return { start, end, prefix: line.slice(start, ch) };
  }

  function collectLuaNameMap(getToolNames, getControlNames, getToolTypes) {
    const map = new Map();
    const addName = (name, type) => {
      if (!name) return;
      const key = String(name);
      const existing = map.get(key);
      if (!existing) map.set(key, { name: key, type });
      else if (existing.type !== type) existing.type = 'both';
    };
    const tools = typeof getToolNames === 'function' ? getToolNames() : getToolNames;
    if (tools && tools.forEach) tools.forEach((name) => addName(name, 'tool'));
    const controls = typeof getControlNames === 'function' ? getControlNames() : getControlNames;
    if (controls && controls.forEach) controls.forEach((name) => addName(name, 'control'));
    const toolTypes = typeof getToolTypes === 'function' ? getToolTypes() : getToolTypes;
    if (toolTypes && toolTypes.get) {
      map.forEach((item) => {
        if (item.type === 'tool' || item.type === 'both') {
          const t = toolTypes.get(item.name);
          if (t) item.toolType = t;
        }
      });
    }
    return map;
  }

  function getLuaAutocompleteItems(prefix, getToolNames, getControlNames, getToolTypes) {
    if (!prefix) return [];
    const lower = prefix.toLowerCase();
    const map = collectLuaNameMap(getToolNames, getControlNames, getToolTypes);
    const out = [];
    map.forEach((item) => {
      if (item.name.toLowerCase().startsWith(lower)) out.push(item);
    });
    out.sort((a, b) => a.name.localeCompare(b.name));
    return out.slice(0, 8);
  }

  function createLuaAutocompleteMenu(wrapper, onSelect) {
    const menu = document.createElement('div');
    menu.className = 'lua-autocomplete';
    menu.hidden = true;
    const list = document.createElement('div');
    list.className = 'lua-autocomplete-list';
    menu.appendChild(list);
    wrapper.appendChild(menu);
    const state = { items: [], activeIndex: 0 };
    const render = () => {
      list.innerHTML = '';
      state.items.forEach((item, idx) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'lua-autocomplete-item';
        if (idx === state.activeIndex) btn.classList.add('active');
        btn.dataset.type = item.type;
        const nameSpan = document.createElement('span');
        nameSpan.className = 'lua-autocomplete-name';
        nameSpan.textContent = item.name;
        const typeSpan = document.createElement('span');
        typeSpan.className = 'lua-autocomplete-type';
        if (item.type === 'tool') {
          typeSpan.textContent = item.toolType ? item.toolType : 'Tool';
        } else if (item.type === 'both') {
          const label = item.toolType ? `${item.toolType} + Control` : 'Tool + Control';
          typeSpan.textContent = label;
        } else {
          typeSpan.textContent = 'Control';
        }
        btn.appendChild(nameSpan);
        btn.appendChild(typeSpan);
        btn.addEventListener('mousedown', (ev) => {
          ev.preventDefault();
          onSelect(item);
        });
        list.appendChild(btn);
      });
    };
    return {
      show(items) {
        state.items = items;
        state.activeIndex = 0;
        render();
        menu.hidden = items.length === 0;
      },
      hide() {
        menu.hidden = true;
      },
      isOpen() {
        return !menu.hidden;
      },
      selectNext() {
        if (!state.items.length) return;
        state.activeIndex = (state.activeIndex + 1) % state.items.length;
        render();
      },
      selectPrev() {
        if (!state.items.length) return;
        state.activeIndex = (state.activeIndex - 1 + state.items.length) % state.items.length;
        render();
      },
      getActive() {
        return state.items[state.activeIndex] || null;
      },
    };
  }

  function setupLuaAutocompleteForCodeMirror(cm, wrapper, getToolNames, getControlNames, getToolTypes) {
    if (!cm || !wrapper) return;
    const menu = createLuaAutocompleteMenu(wrapper, (item) => {
      const cursor = cm.getCursor();
      const line = cm.getLine(cursor.line);
      const bounds = getLuaWordBounds(line, cursor.ch);
      if (!bounds.prefix) return;
      cm.replaceRange(item.name, { line: cursor.line, ch: bounds.start }, { line: cursor.line, ch: cursor.ch });
      cm.focus();
      menu.hide();
    });
    const update = () => {
      if (document?.body?.classList.contains('pick-mode')) {
        menu.hide();
        return;
      }
      const cursor = cm.getCursor();
      const line = cm.getLine(cursor.line);
      const bounds = getLuaWordBounds(line, cursor.ch);
      if (bounds.end !== cursor.ch || bounds.prefix.length < 2) {
        menu.hide();
        return;
      }
      const items = getLuaAutocompleteItems(bounds.prefix, getToolNames, getControlNames, getToolTypes);
      if (!items.length) {
        menu.hide();
        return;
      }
      menu.show(items);
    };
    cm.on('inputRead', update);
    cm.on('cursorActivity', update);
    cm.on('keydown', (editor, event) => {
      if (!menu.isOpen()) return;
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        menu.selectNext();
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        menu.selectPrev();
      } else if (event.key === 'Enter' || event.key === 'Tab') {
        event.preventDefault();
        const active = menu.getActive();
        if (active) {
          const cursor = editor.getCursor();
          const line = editor.getLine(cursor.line);
          const bounds = getLuaWordBounds(line, cursor.ch);
          if (bounds.prefix) {
            editor.replaceRange(active.name, { line: cursor.line, ch: bounds.start }, { line: cursor.line, ch: cursor.ch });
          }
        }
        menu.hide();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        menu.hide();
      }
    });
    cm.on('blur', () => {
      setTimeout(() => menu.hide(), 60);
    });
  }

  function setupLuaAutocompleteForTextarea(textarea, wrapper, getToolNames, getControlNames, getToolTypes) {
    if (!textarea || !wrapper) return;
    const menu = createLuaAutocompleteMenu(wrapper, (item) => {
      const value = textarea.value || '';
      const pos = textarea.selectionStart || 0;
      const lineStart = value.lastIndexOf('\n', pos - 1) + 1;
      const lineEnd = value.indexOf('\n', pos);
      const line = value.slice(lineStart, lineEnd < 0 ? value.length : lineEnd);
      const bounds = getLuaWordBounds(line, pos - lineStart);
      if (!bounds.prefix) return;
      const insertStart = lineStart + bounds.start;
      textarea.setRangeText(item.name, insertStart, lineStart + (pos - lineStart), 'end');
      textarea.focus();
      menu.hide();
    });
    const update = () => {
      if (document?.body?.classList.contains('pick-mode')) {
        menu.hide();
        return;
      }
      const value = textarea.value || '';
      const pos = textarea.selectionStart || 0;
      const lineStart = value.lastIndexOf('\n', pos - 1) + 1;
      const lineEnd = value.indexOf('\n', pos);
      const line = value.slice(lineStart, lineEnd < 0 ? value.length : lineEnd);
      const bounds = getLuaWordBounds(line, pos - lineStart);
      if (bounds.end !== (pos - lineStart) || bounds.prefix.length < 2) {
        menu.hide();
        return;
      }
      const items = getLuaAutocompleteItems(bounds.prefix, getToolNames, getControlNames, getToolTypes);
      if (!items.length) {
        menu.hide();
        return;
      }
      menu.show(items);
    };
    textarea.addEventListener('input', update);
    textarea.addEventListener('keyup', update);
    textarea.addEventListener('click', update);
    textarea.addEventListener('keydown', (event) => {
      if (!menu.isOpen()) return;
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        menu.selectNext();
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        menu.selectPrev();
      } else if (event.key === 'Enter' || event.key === 'Tab') {
        event.preventDefault();
        const active = menu.getActive();
        if (active) {
          const value = textarea.value || '';
          const pos = textarea.selectionStart || 0;
          const lineStart = value.lastIndexOf('\n', pos - 1) + 1;
          const lineEnd = value.indexOf('\n', pos);
          const line = value.slice(lineStart, lineEnd < 0 ? value.length : lineEnd);
          const bounds = getLuaWordBounds(line, pos - lineStart);
          if (bounds.prefix) {
            const insertStart = lineStart + bounds.start;
            textarea.setRangeText(active.name, insertStart, lineStart + (pos - lineStart), 'end');
          }
        }
        menu.hide();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        menu.hide();
      }
    });
    textarea.addEventListener('blur', () => {
      setTimeout(() => menu.hide(), 60);
    });
  }

  const LUA_OVERLAY_MODE_NAME = 'lua-overlay';

  function ensureLuaOverlayMode() {
    if (typeof window === 'undefined' || !window.CodeMirror || !window.CodeMirror.defineMode) return false;
    if (window.CodeMirror.modes && window.CodeMirror.modes[LUA_OVERLAY_MODE_NAME]) return true;
    window.CodeMirror.defineMode(LUA_OVERLAY_MODE_NAME, (config, parserConfig) => {
      const cfg = parserConfig || {};
      let baseMode;
      if (cfg.baseMode) {
        baseMode = window.CodeMirror.getMode(config, cfg.baseMode);
      } else if (window.CodeMirror.modes && window.CodeMirror.modes.lua) {
        baseMode = window.CodeMirror.getMode(config, 'lua');
      } else {
        baseMode = createFallbackLuaMode();
      }
      if (cfg.overlay && window.CodeMirror.overlayMode) {
        return window.CodeMirror.overlayMode(baseMode, cfg.overlay, true);
      }
      return baseMode;
    });
    return true;
  }

  function createFallbackLuaMode() {
    const keywords = new Set([
      'and', 'break', 'do', 'else', 'elseif', 'end', 'false', 'for', 'function',
      'if', 'in', 'local', 'nil', 'not', 'or', 'repeat', 'return', 'then', 'true',
      'until', 'while',
    ]);
    const atoms = new Set(['true', 'false', 'nil']);
    return {
      startState() { return { inLongComment: false, inLongString: false }; },
      token(stream, state) {
        if (state.inLongComment) {
          if (stream.skipTo(']]')) {
            stream.match(']]', true);
            state.inLongComment = false;
          } else {
            stream.skipToEnd();
          }
          return 'comment';
        }
        if (state.inLongString) {
          if (stream.skipTo(']]')) {
            stream.match(']]', true);
            state.inLongString = false;
          } else {
            stream.skipToEnd();
          }
          return 'string';
        }
        if (stream.match('--[[')) {
          state.inLongComment = true;
          return 'comment';
        }
        if (stream.match('[[', true)) {
          state.inLongString = true;
          return 'string';
        }
        if (stream.match('--')) {
          stream.skipToEnd();
          return 'comment';
        }
        if (stream.match(/^(0x[0-9a-fA-F]+)/)) {
          return 'number';
        }
        if (stream.match(/^(\d+(\.\d+)?([eE][+-]?\d+)?)/)) {
          return 'number';
        }
        if (stream.match(/^"([^"\\]|\\.)*"/) || stream.match(/^'([^'\\]|\\.)*'/)) {
          return 'string';
        }
        if (stream.match(/^[A-Za-z_][A-Za-z0-9_]*/)) {
          const word = stream.current();
          if (keywords.has(word)) return atoms.has(word) ? 'atom' : 'keyword';
          return null;
        }
        stream.next();
        return null;
      },
    };
  }

  function createLuaEditor({
    value = '',
    placeholder = '',
    onBlur,
    getToolNames = null,
    getControlNames = null,
    getToolTypes = null,
  } = {}) {
    const wrapper = document.createElement('div');
    wrapper.className = 'lua-editor';
    const textarea = document.createElement('textarea');
    textarea.className = 'lua-input';
    textarea.spellcheck = false;
    textarea.value = value || '';
    if (placeholder) textarea.placeholder = placeholder;
    wrapper.appendChild(textarea);

    if (typeof window !== 'undefined' && window.CodeMirror && window.CodeMirror.fromTextArea) {
      wrapper.classList.add('lua-editor-cm');
      const hasLuaMode = !!(window.CodeMirror.modes && window.CodeMirror.modes.lua);
      const overlay = createLuaNameOverlay(getToolNames, getControlNames);
      let mode = null;
      if (window.CodeMirror.overlayMode && ensureLuaOverlayMode()) {
        mode = {
          name: LUA_OVERLAY_MODE_NAME,
          baseMode: hasLuaMode ? 'lua' : null,
          overlay,
        };
      } else if (hasLuaMode) {
        mode = 'lua';
      } else {
        mode = createFallbackLuaMode();
      }
      const cm = window.CodeMirror.fromTextArea(textarea, {
        mode,
        lineWrapping: true,
        tabSize: 2,
        indentUnit: 2,
        viewportMargin: 10,
        placeholder: placeholder || undefined,
      });
      try {
        const modeName = (cm.getMode && cm.getMode().name) || (typeof mode === 'string' ? mode : (mode && mode.name)) || 'unknown';
        console.log('[LuaEditor] CodeMirror mode:', modeName, 'overlay:', !!window.CodeMirror.overlayMode);
      } catch (_) {}
      setupLuaAutocompleteForCodeMirror(cm, wrapper, getToolNames, getControlNames, getToolTypes);
      cm.on('blur', () => {
        if (typeof onBlur === 'function') onBlur(cm.getValue());
      });
      if (typeof ResizeObserver !== 'undefined') {
        const ro = new ResizeObserver(() => {
          try { cm.setSize('100%', '100%'); } catch (_) {}
        });
        ro.observe(wrapper);
      }
      return {
        wrapper,
        getValue: () => cm.getValue(),
        setValue: (next) => cm.setValue(next || ''),
        insertAtCursor: (text) => {
          const value = text == null ? '' : String(text);
          cm.focus();
          cm.replaceSelection(value, 'end');
        },
      };
    }

    const pre = document.createElement('pre');
    pre.className = 'lua-highlight';
    const code = document.createElement('code');
    code.className = 'language-lua';
    pre.appendChild(code);
    wrapper.appendChild(pre);
    if (placeholder) {
      const placeholderEl = document.createElement('div');
      placeholderEl.className = 'lua-placeholder';
      placeholderEl.textContent = placeholder;
      wrapper.appendChild(placeholderEl);
    }
    const syncScroll = () => {
      pre.scrollTop = textarea.scrollTop;
      pre.scrollLeft = textarea.scrollLeft;
    };
    const syncHighlight = () => {
      code.innerHTML = highlightLua(textarea.value, {
        toolNames: (typeof getToolNames === 'function') ? getToolNames() : getToolNames,
        controlNames: (typeof getControlNames === 'function') ? getControlNames() : getControlNames,
      });
      wrapper.classList.toggle('is-empty', !textarea.value);
      syncScroll();
    };
    syncHighlight();
    setupLuaAutocompleteForTextarea(textarea, wrapper, getToolNames, getControlNames, getToolTypes);
    textarea.addEventListener('input', syncHighlight);
    textarea.addEventListener('scroll', syncScroll);
    textarea.addEventListener('blur', () => {
      if (typeof onBlur === 'function') onBlur(textarea.value);
    });
    return {
      wrapper,
      textarea,
      getValue: () => textarea.value,
      setValue: (next) => {
        textarea.value = next || '';
        syncHighlight();
      },
      insertAtCursor: (text) => {
        const value = text == null ? '' : String(text);
        const start = textarea.selectionStart ?? textarea.value.length;
        const end = textarea.selectionEnd ?? start;
        textarea.setRangeText(value, start, end, 'end');
        syncHighlight();
        textarea.focus();
      },
    };
  }

  let activePickSession = null;
  let activePickHighlight = null;

  function stopPickSession() {
    activePickSession = null;
    if (activePickHighlight) {
      try { activePickHighlight.classList.remove('pick-target'); } catch (_) {}
      activePickHighlight = null;
    }
    try { document.body && document.body.classList.remove('pick-mode'); } catch (_) {}
  }

  function findPickHighlightTarget(target) {
    try {
      if (nodesList && nodesList.contains(target)) {
        const row = target.closest('.node-row');
        if (row) return row;
        const title = target.closest('.node-title');
        if (title) return title;
      }
      if (controlsList && controlsList.contains(target)) {
        const li = target.closest('li[data-index]');
        if (li) return li;
      }
    } catch (_) {}
    return null;
  }

  function updatePickHighlight(target) {
    const next = findPickHighlightTarget(target);
    if (next === activePickHighlight) return;
    if (activePickHighlight) {
      try { activePickHighlight.classList.remove('pick-target'); } catch (_) {}
    }
    activePickHighlight = next;
    if (activePickHighlight) {
      try { activePickHighlight.classList.add('pick-target'); } catch (_) {}
    }
  }

  function startPickSession(insertFn, options = {}) {
    stopPickSession();
    activePickSession = { insertFn, pendingValue: null, sticky: !!options.sticky, owner: options.owner || null };
    try {
      document.querySelectorAll('.lua-autocomplete').forEach((el) => {
        el.hidden = true;
      });
    } catch (_) {}
    try { document.body && document.body.classList.add('pick-mode'); } catch (_) {}
    info('Pick mode: click a node or published control to insert its name. Press Esc to cancel.');
  }

  function resolvePickValue(target) {
    try {
      if (nodesList && nodesList.contains(target)) {
        const row = target.closest('.node-row');
        if (row) {
          if (row.dataset.kind === 'control' && row.dataset.source && row.dataset.sourceOp) {
            return `${row.dataset.sourceOp}.${row.dataset.source}`;
          }
          if (row.dataset.kind === 'group' && row.dataset.groupBase && row.dataset.sourceOp) {
            return `${row.dataset.sourceOp}.${row.dataset.groupBase}`;
          }
        }
        const title = target.closest('.node-title');
        if (title) {
          const wrap = title.closest('.node');
          const op = wrap && wrap.dataset ? wrap.dataset.op : null;
          if (op) return op;
        }
      }
      if (controlsList && controlsList.contains(target)) {
        const li = target.closest('li[data-index]');
        if (!li) return null;
        const idx = parseInt(li.dataset.index || '-1', 10);
        const entry = (state.parseResult && Array.isArray(state.parseResult.entries)) ? state.parseResult.entries[idx] : null;
        if (!entry) return null;
        if (entry.sourceOp && entry.source) return `${entry.sourceOp}.${entry.source}`;
        return entry.source || entry.sourceOp || null;
      }
    } catch (_) {}
      return null;
    }

  function normalizePickValue(value) {
    const raw = value == null ? '' : String(value).trim();
    if (!raw) return raw;
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return raw;
    const dot = raw.indexOf('.');
    if (dot > 0) {
      const left = raw.slice(0, dot);
      const right = raw.slice(dot + 1);
      const entry = state.parseResult.entries.find(e => e && e.sourceOp === left && (e.source === right || e.displayName === right || e.name === right || e.displayNameOriginal === right));
      if (entry && entry.source) return `${entry.sourceOp}.${entry.source}`;
      return raw;
    }
    const entry = state.parseResult.entries.find(e => e && (e.displayName === raw || e.name === raw || e.displayNameOriginal === raw));
    if (entry && entry.sourceOp && entry.source) return `${entry.sourceOp}.${entry.source}`;
    if (entry && entry.source) return entry.source;
    return raw;
  }

  function handlePickPointer(ev) {
    if (!activePickSession) return;
    const value = normalizePickValue(activePickSession.pendingValue || resolvePickValue(ev.target));
    if (!value) return;
    ev.preventDefault();
    ev.stopPropagation();
    if (ev.type === 'mousedown') {
      activePickSession.pendingValue = String(value);
      return;
    }
    try {
      if (activePickSession && typeof activePickSession.insertFn === 'function') {
        activePickSession.insertFn(String(value));
      }
    } catch (_) {}
    if (activePickSession && activePickSession.sticky) {
      activePickSession.pendingValue = null;
    } else {
      stopPickSession();
    }
  }

  function handlePickMove(ev) {
    if (!activePickSession) return;
    updatePickHighlight(ev.target);
  }

  function handlePickKey(ev) {
    if (!activePickSession) return;
    if (ev.key === 'Escape') {
      ev.preventDefault();
      stopPickSession();
    }
  }

  if (typeof document !== 'undefined') {
    document.addEventListener('mousedown', handlePickPointer, true);
    document.addEventListener('click', handlePickPointer, true);
    document.addEventListener('mousemove', handlePickMove, true);
    document.addEventListener('keydown', handlePickKey, true);
  }

  function buildPickRow(editor) {
    const row = document.createElement('div');
    row.className = 'detail-actions detail-actions-pick';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'toggle';
    btn.textContent = 'Live Linking';
    btn.setAttribute('aria-pressed', 'false');
    let liveActive = false;
    let hasFocus = false;
    const wrapper = editor && editor.wrapper;
    const insertFn = (value) => {
      if (editor && typeof editor.insertAtCursor === 'function') {
        editor.insertAtCursor(value);
      }
    };
    const focusEditor = () => {
      if (editor && editor.textarea) {
        editor.textarea.focus();
        return;
      }
      if (wrapper) {
        const cmEl = wrapper.querySelector('.CodeMirror');
        if (cmEl && cmEl.CodeMirror && typeof cmEl.CodeMirror.focus === 'function') {
          cmEl.CodeMirror.focus();
          return;
        }
      }
    };
    const setLiveActive = (next) => {
      liveActive = !!next;
      btn.setAttribute('aria-pressed', liveActive ? 'true' : 'false');
      btn.classList.toggle('is-active', liveActive);
      if (liveActive && hasFocus) {
        startPickSession(insertFn, { sticky: true, owner: wrapper || null });
      } else if (!liveActive && activePickSession && activePickSession.sticky) {
        stopPickSession();
      }
    };
    if (wrapper) {
      wrapper.addEventListener('focusin', () => {
        hasFocus = true;
        if (liveActive) startPickSession(insertFn, { sticky: true, owner: wrapper || null });
      });
      wrapper.addEventListener('focusout', () => {
        setTimeout(() => {
          const active = document.activeElement;
          const stillInsideEditor = wrapper.contains(active);
          const stillInsidePickControls = row.contains(active);
          if (!stillInsideEditor && !stillInsidePickControls) {
            hasFocus = false;
          }
        }, 0);
      });
    }
    btn.addEventListener('click', () => {
      if (liveActive) {
        setLiveActive(false);
        return;
      }
      focusEditor();
      setLiveActive(true);
    });
    row.appendChild(btn);
    return row;
  }

  function renderDetailDrawer(index) {
    if (!detailDrawer || !detailDrawerTitle || !detailDrawerSubtitle || !detailDrawerBody) return;
    if (!state.parseResult || !Array.isArray(state.parseResult.entries) || !state.parseResult.entries[index]) {
      hideDetailDrawer();
      return;
    }
    activeDetailEntryIndex = index;
    drawerMode = DRAWER_MODE.CONTROL;
    const entry = state.parseResult.entries[index];
    hydrateTextControlMeta(entry);
    applyDrawerState(true);
    detailDrawerTitle.textContent = entry.displayName || entry.name || entry.source || 'Control';
    detailDrawerTitle.contentEditable = 'true';
    detailDrawerTitle.spellcheck = false;
    const commitDrawerName = () => {
      setEntryDisplayNameApi(index, (detailDrawerTitle.textContent || '').trim());
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
    const drawerNav = getSelectedDrawerNavigation(index);
    if (drawerNav) {
      const navRow = document.createElement('div');
      navRow.className = 'detail-selection-nav';
      const prevBtn = document.createElement('button');
      prevBtn.type = 'button';
      prevBtn.className = 'detail-selection-nav-btn';
      prevBtn.innerHTML = renderIconHtml ? renderIconHtml('chevron-left', 14) : '&lt;';
      prevBtn.title = 'Previous selected control';
      prevBtn.disabled = !(typeof drawerNav.previousIndex === 'number');
      prevBtn.addEventListener('click', () => {
        if (typeof drawerNav.previousIndex === 'number') focusDrawerSelectedEntry(drawerNav.previousIndex);
      });
      const status = document.createElement('div');
      status.className = 'detail-selection-nav-status';
      status.textContent = `${drawerNav.currentPos + 1} of ${drawerNav.total} selected`;
      const nextBtn = document.createElement('button');
      nextBtn.type = 'button';
      nextBtn.className = 'detail-selection-nav-btn';
      nextBtn.innerHTML = renderIconHtml ? renderIconHtml('chevron-right', 14) : '&gt;';
      nextBtn.title = 'Next selected control';
      nextBtn.disabled = !(typeof drawerNav.nextIndex === 'number');
      nextBtn.addEventListener('click', () => {
        if (typeof drawerNav.nextIndex === 'number') focusDrawerSelectedEntry(drawerNav.nextIndex);
      });
      navRow.appendChild(prevBtn);
      navRow.appendChild(status);
      navRow.appendChild(nextBtn);
      detailDrawerBody.appendChild(navRow);
    }
    const buildCollapsibleField = (title, defaultOpen = false, stateKey = null) => {
      const field = document.createElement('div');
      field.className = 'detail-field detail-collapsible';
      const header = document.createElement('button');
      header.type = 'button';
      header.className = 'detail-collapsible-header';
      const iconSpan = document.createElement('span');
      iconSpan.className = 'icon';
      if (typeof createIcon === 'function') iconSpan.innerHTML = createIcon('chevron-right', 12);
      else iconSpan.textContent = '>';
      const titleSpan = document.createElement('span');
      titleSpan.textContent = title;
      header.appendChild(iconSpan);
      header.appendChild(titleSpan);
      const body = document.createElement('div');
      body.className = 'detail-collapsible-body';
      let open = !!defaultOpen;
      if (stateKey) {
        entry._ui = entry._ui || {};
        if (typeof entry._ui[stateKey] === 'boolean') {
          open = entry._ui[stateKey];
        }
      }
      const sync = () => {
        body.hidden = !open;
        field.classList.toggle('open', open);
        if (stateKey) {
          entry._ui = entry._ui || {};
          entry._ui[stateKey] = open;
        }
        if (typeof createIcon === 'function') {
          iconSpan.innerHTML = createIcon(open ? 'chevron-down' : 'chevron-right', 12);
        } else {
          iconSpan.textContent = open ? 'v' : '>';
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
    const effectiveDataType = entry.isButton ? '"Number"' : (inferEntryDataType(entry) || null);
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
    const isPathDrawModeEntry = isPathEntryForDrawMode(entry);
    const showDefaultsField = entry.isLabel || !isPathDrawModeEntry;
    if (showDefaultsField) {
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
      const currentRow = document.createElement('div');
      currentRow.className = 'detail-default-row';
      const currentGroup = document.createElement('div');
      currentGroup.className = 'detail-default-item';
      const currentLabel = document.createElement('span');
      currentLabel.textContent = 'Current';
      const isChoice = !isTextControl(entry) && isChoiceControl(entry);
      const comboOptions = isChoice ? getChoiceOptions(entry) : [];
      const useComboSelect = isChoice && comboOptions.length > 0;
      let currentInput = null;
      let currentSelect = null;
      if (useComboSelect) {
        currentSelect = document.createElement('select');
        currentSelect.className = 'detail-type-select current-input current-combo-input';
        comboOptions.forEach((label, idx) => {
          const opt = document.createElement('option');
          opt.value = String(idx);
          opt.textContent = String(label == null ? `Option ${idx}` : label);
          currentSelect.appendChild(opt);
        });
      } else {
        currentInput = document.createElement('input');
        currentInput.type = isChoice ? 'number' : 'text';
        if (isChoice) {
          currentInput.step = '1';
          currentInput.min = '0';
          currentInput.placeholder = 'Option index (0-based)';
        }
        currentInput.className = 'current-input';
      }
      const currentNote = document.createElement('span');
      currentNote.className = 'detail-default-note';
      const currentWrap = document.createElement('div');
      currentWrap.className = 'detail-input-wrap';
      currentWrap.appendChild(useComboSelect ? currentSelect : currentInput);
      const hasCsv = !!(state.csvData && Array.isArray(state.csvData.headers) && state.csvData.headers.length);
      const hasLinkSource = !!(state.parseResult && state.parseResult.dataLink && state.parseResult.dataLink.source);
      const canLinkCsv = isCsvLinkableControl(entry);
      if (canLinkCsv && (hasCsv || hasLinkSource || entry?.csvLink?.column)) {
        currentWrap.classList.add('has-action');
        const csvBtn = document.createElement('button');
        csvBtn.type = 'button';
        csvBtn.className = 'csv-link-btn';
        const linked = entry && entry.csvLink && entry.csvLink.column;
        if (linked) csvBtn.classList.add('linked');
        csvBtn.innerHTML = getCsvLinkIcon();
        csvBtn.title = linked
          ? `Linked to ${entry.csvLink.column}`
          : 'Link CSV column';
        csvBtn.addEventListener('click', async (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          const headers = state.csvData?.headers || [];
          if (!headers.length) {
            error('Import a CSV or Google Sheet first.');
            return;
          }
          const picked = await openCsvColumnPicker({
            title: 'Link CSV column',
            headers,
            current: entry?.csvLink?.column || '',
            allowNone: true,
          });
          if (picked == null) return;
        if (!picked) {
          delete entry.csvLink;
          renderDetailDrawer(index);
          syncDataLinkPanel();
          return;
        }
        entry.csvLink = { column: picked };
        const link = ensureDataLink(state.parseResult);
        syncDataLinkMappings(state.parseResult);
        renderDetailDrawer(index);
        syncDataLinkPanel();
        if (state.csvData && Array.isArray(state.csvData.rows) && state.csvData.rows.length) {
          const rows = state.csvData.rows || [];
          if (rows.length) {
            await applyCsvRowsToCurrentMacro(rows, link, { silent: true });
          }
        }
        });
        currentWrap.appendChild(csvBtn);
      }
      const refreshCurrent = () => {
        const info = getCurrentInputInfo(entry);
        const rawValue = String(info.value || '').trim().replace(/^\"|\"$/g, '');
        if (useComboSelect) {
          const idx = parseComboIndexValue(entry, rawValue);
          currentSelect.value = idx != null ? String(idx) : '';
        } else {
          currentInput.value = info.value || '';
          if (!isChoice) {
            currentInput.placeholder = info.value ? '' : (info.note || 'Current value');
          }
        }
        const csvNote = entry?.csvLink?.column ? `Linked to CSV: ${entry.csvLink.column}` : '';
        const noteParts = [];
        if (useComboSelect) {
          const idx = parseComboIndexValue(entry, rawValue);
          const label = idx != null ? getComboOptionLabel(entry, idx) : '';
          if (label) noteParts.push(`Current option: ${label} (${idx})`);
        } else if (isChoice) {
          noteParts.push('Choice control uses numeric index values starting at 0.');
          noteParts.push('No option labels found on this control yet.');
        }
        if (info.note) noteParts.push(info.note);
        if (csvNote) noteParts.push(csvNote);
        currentNote.textContent = noteParts.join(' ');
        currentNote.hidden = noteParts.length === 0;
      };
      const commitCurrent = () => {
        const nextValue = useComboSelect ? currentSelect.value : currentInput.value;
        setEntryCurrentValue(index, nextValue);
        refreshCurrent();
      };
      if (useComboSelect) {
        currentSelect.addEventListener('change', commitCurrent);
      } else {
        currentInput.addEventListener('keydown', (ev) => {
          if (ev.key === 'Enter') { ev.preventDefault(); commitCurrent(); }
        });
        currentInput.addEventListener('blur', commitCurrent);
      }
      const cgBlock = getColorGroupBlockByIndex(index);
      if (cgBlock && cgBlock.firstIndex === index && cgBlock.count >= 2) {
        const components = getColorGroupComponents(cgBlock);
        if (components) {
          const hexRow = document.createElement('div');
          hexRow.className = 'csv-hex-row';
          const hexLabel = document.createElement('span');
          hexLabel.textContent = 'Hex';
          const hexWrap = document.createElement('div');
          hexWrap.className = 'detail-input-wrap csv-hex-wrap';
          const hexInput = document.createElement('input');
          hexInput.type = 'text';
          hexInput.placeholder = '#RRGGBB or #RRGGBBAA';
          const hexPreview = document.createElement('div');
          hexPreview.className = 'csv-hex-preview';
          hexWrap.appendChild(hexInput);
          hexWrap.appendChild(hexPreview);
          const readComponent = (idx, fallback) => {
            if (!Number.isFinite(idx)) return fallback;
            const info = getCurrentInputInfo(state.parseResult.entries[idx]);
            const cleaned = String(info.value || '').trim().replace(/^\"|\"$/g, '');
            const num = parseFloat(cleaned);
            return Number.isFinite(num) ? clamp01(num) : fallback;
          };
          const refreshHex = () => {
            const r = readComponent(components.red, 0);
            const g = readComponent(components.green, 0);
            const b = readComponent(components.blue, 0);
            const a = readComponent(components.alpha, 1);
            const hex = formatHexColor({ r, g, b, a });
            hexInput.value = hex;
            hexPreview.style.background = `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${clamp01(a)})`;
          };
          const applyHex = (raw) => {
            const parsed = parseHexColor(raw);
            if (!parsed) return false;
            const r = clamp01(parsed.r);
            const g = clamp01(parsed.g);
            const b = clamp01(parsed.b);
            const a = parsed.hasAlpha ? clamp01(parsed.a) : 1;
            const nextText = applyHexToColorGroup({
              entries: state.parseResult.entries,
              order: state.parseResult.order,
              baseText: state.originalText || '',
              groupIndex: index,
              r: r.toFixed(6),
              g: g.toFixed(6),
              b: b.toFixed(6),
              a: a.toFixed(6),
              includeAlpha: parsed.hasAlpha,
              eol: state.newline || '\n',
              resultRef: state.parseResult,
            });
            if (nextText && nextText !== state.originalText) {
              state.originalText = nextText;
              markContentDirty();
            }
            hexInput.value = formatHexColor({ r, g, b, a, includeAlpha: parsed.hasAlpha });
            hexPreview.style.background = `rgba(${Math.round(r * 255)}, ${Math.round(g * 255)}, ${Math.round(b * 255)}, ${clamp01(a)})`;
            return true;
          };
          hexInput.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') { ev.preventDefault(); applyHex(hexInput.value); }
          });
          hexInput.addEventListener('blur', () => { applyHex(hexInput.value); });
          if (hasCsv) {
            hexWrap.classList.add('has-action');
            const hexBtn = document.createElement('button');
            hexBtn.type = 'button';
            hexBtn.className = 'csv-link-btn';
            const linked = entry && entry.csvLinkHex && entry.csvLinkHex.column;
            if (linked) hexBtn.classList.add('linked');
            hexBtn.innerHTML = getCsvLinkIcon();
            hexBtn.title = linked
              ? `Linked to ${entry.csvLinkHex.column}`
              : 'Link CSV column';
            hexBtn.addEventListener('click', async (ev) => {
              ev.preventDefault();
              ev.stopPropagation();
              const headers = state.csvData?.headers || [];
              const picked = await openCsvColumnPicker({
                title: 'Link CSV column',
                headers,
                current: entry?.csvLinkHex?.column || '',
                allowNone: true,
              });
              if (picked == null) return;
              if (!picked) {
                delete entry.csvLinkHex;
              } else {
                entry.csvLinkHex = entry.csvLinkHex || {};
                entry.csvLinkHex.column = picked;
              }
              syncDataLinkMappings(state.parseResult);
              syncDataLinkPanel();
              renderDetailDrawer(index);
              if (state.csvData && Array.isArray(state.csvData.rows) && state.csvData.rows.length) {
                await applyCsvRowsToCurrentMacro(state.csvData.rows, ensureDataLink(state.parseResult), { silent: true });
              }
            });
            hexWrap.appendChild(hexBtn);
          }
          hexRow.appendChild(hexLabel);
          hexRow.appendChild(hexWrap);
          currentGroup.appendChild(hexRow);

          if (entry?.csvLinkHex?.column) {
            const overrideRow = document.createElement('div');
            overrideRow.className = 'csv-override-row';
            const overrideLabel = document.createElement('span');
            overrideLabel.textContent = 'Row override';
            const overrideSelect = document.createElement('select');
            overrideSelect.innerHTML = `
              <option value="default">Default (macro)</option>
              <option value="index">Row index</option>
              <option value="key">Key/value</option>
            `;
            const overrideInputs = document.createElement('div');
            overrideInputs.className = 'csv-override-inputs';
            const indexInput = document.createElement('input');
            indexInput.type = 'number';
            indexInput.min = '1';
            indexInput.placeholder = 'Row #';
            const keyInput = document.createElement('input');
            keyInput.type = 'text';
            keyInput.placeholder = 'Column';
            const valueInput = document.createElement('input');
            valueInput.type = 'text';
            valueInput.placeholder = 'Value';
            overrideInputs.appendChild(indexInput);
            overrideInputs.appendChild(keyInput);
            overrideInputs.appendChild(valueInput);
            const mode = entry.csvLinkHex.rowMode || 'default';
            overrideSelect.value = mode;
            if (Number.isFinite(entry.csvLinkHex.rowIndex)) indexInput.value = String(entry.csvLinkHex.rowIndex);
            if (entry.csvLinkHex.rowKey) keyInput.value = String(entry.csvLinkHex.rowKey);
            if (entry.csvLinkHex.rowValue) valueInput.value = String(entry.csvLinkHex.rowValue);
            const syncOverrideVisibility = () => {
              const val = overrideSelect.value || 'default';
              indexInput.style.display = val === 'index' ? '' : 'none';
              keyInput.style.display = val === 'key' ? '' : 'none';
              valueInput.style.display = val === 'key' ? '' : 'none';
            };
            syncOverrideVisibility();
            const applyOverride = async () => {
              const val = overrideSelect.value || 'default';
              entry.csvLinkHex = entry.csvLinkHex || {};
              if (val === 'default') {
                delete entry.csvLinkHex.rowMode;
                delete entry.csvLinkHex.rowIndex;
                delete entry.csvLinkHex.rowKey;
                delete entry.csvLinkHex.rowValue;
              } else {
                entry.csvLinkHex.rowMode = val;
                if (val === 'index') {
                  const idx = Math.max(1, Number(indexInput.value || 1));
                  entry.csvLinkHex.rowIndex = idx;
                  delete entry.csvLinkHex.rowKey;
                  delete entry.csvLinkHex.rowValue;
                } else if (val === 'key') {
                  entry.csvLinkHex.rowKey = keyInput.value || '';
                  entry.csvLinkHex.rowValue = valueInput.value || '';
                  delete entry.csvLinkHex.rowIndex;
                }
              }
              syncDataLinkMappings(state.parseResult);
              if (state.csvData && Array.isArray(state.csvData.rows) && state.csvData.rows.length) {
                await applyCsvRowsToCurrentMacro(state.csvData.rows, ensureDataLink(state.parseResult), { silent: true });
              }
              if (state.parseResult && state.parseResult.entries && state.parseResult.entries[index]) {
                renderDetailDrawer(index);
                syncDataLinkPanel();
              }
            };
            overrideSelect.addEventListener('change', async () => {
              syncOverrideVisibility();
              await applyOverride();
            });
            indexInput.addEventListener('blur', applyOverride);
            keyInput.addEventListener('blur', applyOverride);
            valueInput.addEventListener('blur', applyOverride);
            overrideRow.appendChild(overrideLabel);
            overrideRow.appendChild(overrideSelect);
            overrideRow.appendChild(overrideInputs);
            currentGroup.appendChild(overrideRow);
          }
          refreshHex();
        }
      }
      currentGroup.appendChild(currentLabel);
      currentGroup.appendChild(currentWrap);
      currentGroup.appendChild(currentNote);
      if (canLinkCsv && entry?.csvLink?.column) {
        const overrideRow = document.createElement('div');
        overrideRow.className = 'csv-override-row';
        const overrideLabel = document.createElement('span');
        overrideLabel.textContent = 'Row override';
        const overrideSelect = document.createElement('select');
        overrideSelect.innerHTML = `
          <option value="default">Default (macro)</option>
          <option value="index">Row index</option>
          <option value="key">Key/value</option>
        `;
        const overrideInputs = document.createElement('div');
        overrideInputs.className = 'csv-override-inputs';
        const indexInput = document.createElement('input');
        indexInput.type = 'number';
        indexInput.min = '1';
        indexInput.placeholder = 'Row #';
        const keyInput = document.createElement('input');
        keyInput.type = 'text';
        keyInput.placeholder = 'Column';
        const valueInput = document.createElement('input');
        valueInput.type = 'text';
        valueInput.placeholder = 'Value';
        overrideInputs.appendChild(indexInput);
        overrideInputs.appendChild(keyInput);
        overrideInputs.appendChild(valueInput);

        const link = ensureDataLink(state.parseResult);
        const mode = entry.csvLink.rowMode || 'default';
        overrideSelect.value = mode;
        if (Number.isFinite(entry.csvLink.rowIndex)) indexInput.value = String(entry.csvLink.rowIndex);
        if (entry.csvLink.rowKey) keyInput.value = String(entry.csvLink.rowKey);
        if (entry.csvLink.rowValue) valueInput.value = String(entry.csvLink.rowValue);

        const syncOverrideVisibility = () => {
          const val = overrideSelect.value || 'default';
          indexInput.style.display = val === 'index' ? '' : 'none';
          keyInput.style.display = val === 'key' ? '' : 'none';
          valueInput.style.display = val === 'key' ? '' : 'none';
        };
        syncOverrideVisibility();

        const applyOverride = async () => {
          const val = overrideSelect.value || 'default';
          if (val === 'default') {
            delete entry.csvLink.rowMode;
            delete entry.csvLink.rowIndex;
            delete entry.csvLink.rowKey;
            delete entry.csvLink.rowValue;
          } else {
            entry.csvLink.rowMode = val;
            if (val === 'index') {
              const idx = Math.max(1, Number(indexInput.value || 1));
              entry.csvLink.rowIndex = idx;
              delete entry.csvLink.rowKey;
              delete entry.csvLink.rowValue;
            } else if (val === 'key') {
              entry.csvLink.rowKey = keyInput.value || '';
              entry.csvLink.rowValue = valueInput.value || '';
              delete entry.csvLink.rowIndex;
            }
          }
          syncDataLinkMappings(state.parseResult);
          if (state.csvData && Array.isArray(state.csvData.rows) && state.csvData.rows.length) {
            await applyCsvRowsToCurrentMacro(state.csvData.rows, link, { silent: true });
          }
          if (state.parseResult && state.parseResult.entries && state.parseResult.entries[index]) {
            renderDetailDrawer(index);
            syncDataLinkPanel();
          }
        };
        overrideSelect.addEventListener('change', async () => {
          syncOverrideVisibility();
          await applyOverride();
        });
        indexInput.addEventListener('blur', applyOverride);
        keyInput.addEventListener('blur', applyOverride);
        valueInput.addEventListener('blur', applyOverride);

        overrideRow.appendChild(overrideLabel);
        overrideRow.appendChild(overrideSelect);
        overrideRow.appendChild(overrideInputs);
        currentGroup.appendChild(overrideRow);
      }
      currentRow.appendChild(currentGroup);
      refreshCurrent();
      defaultsField.appendChild(currentRow);
      if (isTextControl(entry)) {
        const linesRow = document.createElement('div');
        linesRow.className = 'detail-default-row';
        const linesItem = document.createElement('div');
        linesItem.className = 'detail-default-item';
        const linesLabel = document.createElement('span');
        linesLabel.textContent = 'Lines';
        const linesInput = document.createElement('input');
        linesInput.type = 'number';
        linesInput.min = '1';
        linesInput.max = '20';
        const metaLines = entry?.controlMeta?.textLines ?? entry?.controlMetaOriginal?.textLines;
        const parsed = metaLines != null ? Number(String(metaLines).replace(/"/g, '').trim()) : NaN;
        const initial = Number.isFinite(parsed) ? parsed : 8;
        linesInput.value = String(initial);
        linesInput.addEventListener('change', () => {
          setEntryTextLines(index, linesInput.value);
        });
        linesInput.addEventListener('blur', () => {
          setEntryTextLines(index, linesInput.value);
        });
        linesItem.appendChild(linesLabel);
        linesItem.appendChild(linesInput);
        linesRow.appendChild(linesItem);
        defaultsField.appendChild(linesRow);
      }
      if (!isTextControl(entry)) {
        const defaultRow = document.createElement('div');
        defaultRow.className = 'detail-default-row';
        const rangeRow = document.createElement('div');
        rangeRow.className = 'detail-default-row';
        const comboDefaults = isChoiceControl(entry) ? getChoiceOptions(entry) : [];
        const buildDefaultInput = (title, key, handler) => {
          const group = document.createElement('div');
          group.className = 'detail-default-item';
          const lbl = document.createElement('span');
          lbl.textContent = title;
          const isComboDefault = key === 'defaultValue' && comboDefaults.length > 0;
          if (isComboDefault) {
            const select = document.createElement('select');
            select.className = 'detail-type-select current-input current-combo-input';
            comboDefaults.forEach((label, idx) => {
              const opt = document.createElement('option');
              opt.value = String(idx);
              opt.textContent = String(label == null ? `Option ${idx}` : label);
              select.appendChild(opt);
            });
            const parsedDefault = parseComboIndexValue(entry, meta[key]);
            if (parsedDefault != null) {
              select.value = String(parsedDefault);
            }
            select.addEventListener('change', () => handler(select.value));
            group.appendChild(lbl);
            group.appendChild(select);
          } else {
            const input = document.createElement('input');
            const isComboNumeric = key === 'defaultValue' && isChoiceControl(entry) && !isTextControl(entry);
            input.type = isComboNumeric ? 'number' : 'text';
            if (isComboNumeric) {
              input.step = '1';
              input.min = '0';
              input.placeholder = 'Index (0-based)';
            }
            input.value = formatDefaultForDisplay(entry, meta[key]) || '';
            const originalMeta = entry.controlMetaOriginal || {};
            if (originalMeta[key] != null) {
              const originalDisplay = formatDefaultForDisplay(entry, originalMeta[key]);
              if (!isComboNumeric) input.placeholder = `Original: ${originalDisplay || ''}`;
            }
            const commit = () => handler(input.value);
            input.addEventListener('keydown', (ev) => {
              if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
            });
            input.addEventListener('blur', commit);
            group.appendChild(lbl);
            group.appendChild(input);
          }
          return group;
        };
        defaultRow.appendChild(buildDefaultInput('Default', 'defaultValue', (val) => setEntryDefaultValue(index, val)));
        defaultsField.appendChild(defaultRow);
        if (!isChoiceControl(entry)) {
          rangeRow.appendChild(buildDefaultInput('Min', 'minScale', (val) => setEntryRangeValue(index, 'minScale', val)));
          rangeRow.appendChild(buildDefaultInput('Max', 'maxScale', (val) => setEntryRangeValue(index, 'maxScale', val)));
          defaultsField.appendChild(rangeRow);
        }
      }
      detailDrawerBody.appendChild(defaultsField);
    }
    if (!entry.isLabel && isChoiceControl(entry)) {
      const comboField = document.createElement('div');
      comboField.className = 'detail-field';
      const comboLabel = document.createElement('label');
      comboLabel.textContent = 'Options';
      comboField.appendChild(comboLabel);
      if (isMultiButtonControl(entry)) {
        const styleRow = document.createElement('div');
        styleRow.className = 'detail-default-row';
        const styleItem = document.createElement('div');
        styleItem.className = 'detail-default-item detail-choice-toggle';
        const styleLabel = document.createElement('span');
        styleLabel.textContent = 'Show Basic Button';
        const styleToggle = document.createElement('input');
        styleToggle.type = 'checkbox';
        styleToggle.checked = getMultiButtonShowBasic(entry);
        styleToggle.addEventListener('change', () => {
          setEntryMultiButtonShowBasic(index, !!styleToggle.checked);
        });
        styleItem.appendChild(styleLabel);
        styleItem.appendChild(styleToggle);
        styleRow.appendChild(styleItem);
        comboField.appendChild(styleRow);
      }
      const comboInput = document.createElement('textarea');
      comboInput.className = 'detail-combo-options-input';
      const options = getChoiceOptions(entry);
      comboInput.value = options.length ? options.join('\n') : '';
      comboInput.placeholder = 'Option 0\nOption 1\nOption 2';
      comboField.appendChild(comboInput);
      const comboNote = document.createElement('span');
      comboNote.className = 'detail-default-note';
      comboNote.textContent = 'One option per line. Option order defines numeric values starting at 0.';
      comboField.appendChild(comboNote);
      const comboActions = document.createElement('div');
      comboActions.className = 'detail-actions';
      const comboSaveBtn = document.createElement('button');
      comboSaveBtn.type = 'button';
      comboSaveBtn.textContent = 'Save';
      comboSaveBtn.addEventListener('click', () => {
        const parsed = parseComboOptionsText(comboInput.value);
        setEntryComboOptions(index, parsed);
      });
      const comboResetBtn = document.createElement('button');
      comboResetBtn.type = 'button';
      comboResetBtn.textContent = 'Reset';
      comboResetBtn.addEventListener('click', () => {
        const original = Array.isArray(entry?.controlMetaOriginal?.choiceOptions)
          ? entry.controlMetaOriginal.choiceOptions
          : [];
        comboInput.value = (original.length ? original : ['Option 1', 'Option 2', 'Option 3']).join('\n');
        const parsed = parseComboOptionsText(comboInput.value);
        setEntryComboOptions(index, parsed);
      });
      comboActions.appendChild(comboSaveBtn);
      comboActions.appendChild(comboResetBtn);
      comboField.appendChild(comboActions);
      comboInput.addEventListener('keydown', (ev) => {
        const metaEnter = ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey);
        if (!metaEnter) return;
        ev.preventDefault();
        const parsed = parseComboOptionsText(comboInput.value);
        setEntryComboOptions(index, parsed);
      });
      detailDrawerBody.appendChild(comboField);
    }
    if (!entry.isLabel) {
      const expressionField = document.createElement('div');
      expressionField.className = 'detail-field';
      const expressionLabel = document.createElement('label');
      expressionLabel.textContent = 'Expression';
      expressionField.appendChild(expressionLabel);

      const expressionInput = document.createElement('textarea');
      expressionInput.className = 'detail-expression-input';
      expressionInput.placeholder = 'Example: Transform1.Angle * 2';
      expressionInput.value = getEntryExpression(entry) || '';
      expressionField.appendChild(expressionInput);

      const expressionNote = document.createElement('span');
      expressionNote.className = 'detail-default-note';
      expressionNote.textContent = expressionInput.value.trim()
        ? 'Expression drives this control. Setting Current value removes it.'
        : 'Set an expression to drive this control from other controls.';
      expressionField.appendChild(expressionNote);

      const expressionActions = document.createElement('div');
      expressionActions.className = 'detail-actions';
      const applyExprBtn = document.createElement('button');
      applyExprBtn.type = 'button';
      applyExprBtn.textContent = 'Apply';
      applyExprBtn.addEventListener('click', () => {
        setEntryExpression(index, expressionInput.value);
      });
      const clearExprBtn = document.createElement('button');
      clearExprBtn.type = 'button';
      clearExprBtn.textContent = 'Clear';
      clearExprBtn.addEventListener('click', () => {
        expressionInput.value = '';
        setEntryExpression(index, '');
      });
      expressionActions.appendChild(applyExprBtn);
      expressionActions.appendChild(clearExprBtn);
      expressionField.appendChild(expressionActions);

      expressionInput.addEventListener('keydown', (ev) => {
        const metaEnter = ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey);
        if (!metaEnter) return;
        ev.preventDefault();
        setEntryExpression(index, expressionInput.value);
      });
      detailDrawerBody.appendChild(expressionField);
    }
    if (!entry.isLabel && isPathDrawModeEntry) {
      const pathField = document.createElement('div');
      pathField.className = 'detail-field';
      const pathLabel = document.createElement('label');
      pathLabel.textContent = 'Path Draw Mode (Export)';
      pathField.appendChild(pathLabel);
      const status = document.createElement('span');
      status.className = 'detail-default-note';
      const topology = detectPathTopologyForTarget(
        state.originalText,
        state.parseResult,
        String(entry.sourceOp || '').trim(),
        String(entry.source || '').trim()
      );
      status.textContent = `Selection is baked during export. ${describePathDrawModeTopology(topology)} This does not live-update Fusion while editing in MM.`;
      pathField.appendChild(status);

      const controlsRow = document.createElement('div');
      controlsRow.className = 'detail-default-row';
      const modeItem = document.createElement('div');
      modeItem.className = 'detail-default-item';
      const modeCaption = document.createElement('span');
      modeCaption.textContent = 'Mode';
      modeItem.appendChild(modeCaption);
      const modeSelect = document.createElement('select');
      modeSelect.className = 'current-combo-input';
      PATH_DRAW_MODE_OPTIONS.forEach((modeName, modeIdx) => {
        const option = document.createElement('option');
        option.value = String(modeIdx);
        option.textContent = modeName;
        if (topology === 'closed' && modeIdx < 2) option.disabled = true;
        if (topology === 'none' && modeIdx > 1) option.disabled = true;
        modeSelect.appendChild(option);
      });
      const explicitIndex = getPathDrawModeStoredIndex(entry, state.originalText, state.parseResult, { onlyExplicit: false });
      const fallbackIndex = Number.isFinite(explicitIndex) ? explicitIndex : 2;
      const selectedIndex = normalizePathDrawModeIndexByTopology(fallbackIndex, topology);
      modeSelect.value = String(selectedIndex);
      modeSelect.addEventListener('change', () => {
        const chosen = Number(modeSelect.value);
        const normalized = normalizePathDrawModeIndexByTopology(chosen, topology);
        if (normalized !== chosen) modeSelect.value = String(normalized);
        setEntryPathDrawModeIndex(index, normalized);
      });
      modeItem.appendChild(modeSelect);
      controlsRow.appendChild(modeItem);
      pathField.appendChild(controlsRow);

      const pathActions = document.createElement('div');
      pathActions.className = 'detail-actions';
      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.textContent = 'Clear Override';
      clearBtn.addEventListener('click', () => {
        clearEntryPathDrawModeIndex(index);
      });
      pathActions.appendChild(clearBtn);
      pathField.appendChild(pathActions);
      detailDrawerBody.appendChild(pathField);
    }
    if (entry.isLabel) {
      const styleField = document.createElement('div');
      styleField.className = 'detail-field';
      const styleLabel = document.createElement('label');
      styleLabel.textContent = 'Style';
      styleField.appendChild(styleLabel);
      const styleRow = document.createElement('div');
      styleRow.className = 'detail-style-row';
      const currentStyle = normalizeLabelStyle(entry.labelStyle);
      const makeToggle = (key, label) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'detail-style-toggle';
        const active = !!currentStyle[key];
        btn.classList.toggle('active', active);
        btn.setAttribute('aria-pressed', active ? 'true' : 'false');
        btn.textContent = label;
        btn.addEventListener('click', () => {
          setEntryLabelStyle(index, { [key]: !currentStyle[key] });
        });
        return btn;
      };
      styleRow.appendChild(makeToggle('bold', 'Bold'));
      styleRow.appendChild(makeToggle('italic', 'Italic'));
      styleRow.appendChild(makeToggle('underline', 'Underline'));
      styleRow.appendChild(makeToggle('center', 'Center'));
      styleField.appendChild(styleRow);
      const colorRow = document.createElement('div');
      colorRow.className = 'detail-style-color';
      const colorToggle = document.createElement('input');
      colorToggle.type = 'checkbox';
      colorToggle.checked = !!currentStyle.color;
      const colorLabel = document.createElement('span');
      colorLabel.textContent = 'Color';
      const colorInput = document.createElement('input');
      colorInput.type = 'color';
      colorInput.value = currentStyle.color || '#ffffff';
      colorInput.disabled = !colorToggle.checked;
      colorToggle.addEventListener('change', () => {
        if (!colorToggle.checked) {
          colorInput.disabled = true;
          setEntryLabelStyle(index, { color: null });
        } else {
          colorInput.disabled = false;
          setEntryLabelStyle(index, { color: colorInput.value });
        }
      });
      colorInput.addEventListener('input', () => {
        if (colorToggle.checked) {
          setEntryLabelStyle(index, { color: colorInput.value });
        }
      });
      colorInput.addEventListener('change', () => {
        if (colorToggle.checked) {
          setEntryLabelStyle(index, { color: colorInput.value });
        }
      });
      colorRow.appendChild(colorToggle);
      colorRow.appendChild(colorLabel);
      colorRow.appendChild(colorInput);
      styleField.appendChild(colorRow);
      detailDrawerBody.appendChild(styleField);
    }
    const onChangeHasValue = !!(entry.onChange && String(entry.onChange).trim());
    const onChangeSection = buildCollapsibleField('On-Change Script', onChangeHasValue, 'onChangeOpen');
    const onChangeBody = onChangeSection.body;
    const onChangeWarnings = document.createElement('div');
    onChangeWarnings.className = 'lua-warnings';
    onChangeWarnings.hidden = true;
    const onChangeEditor = createLuaEditor({
      value: entry.onChange || '',
      placeholder: 'Lua script to execute when this control changes.',
      onBlur: (value) => {
        updateEntryOnChange(index, value, { silent: true });
        renderLuaWarnings(onChangeWarnings, validateLuaBasic(value));
      },
      getToolNames: () => state.parseResult?.luaToolNames,
      getControlNames: () => state.parseResult?.luaControlNames,
      getToolTypes: () => state.parseResult?.luaToolTypes,
    });
    onChangeBody.appendChild(onChangeEditor.wrapper);
    onChangeBody.appendChild(onChangeWarnings);
    renderLuaWarnings(onChangeWarnings, validateLuaBasic(onChangeEditor.getValue()));
    onChangeBody.appendChild(buildPickRow(onChangeEditor));
    const actions = document.createElement('div');
    actions.className = 'detail-actions';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.textContent = 'Save';
    saveBtn.addEventListener('click', () => {
      updateEntryOnChange(index, onChangeEditor.getValue());
    });
    const clearBtn = document.createElement('button');
    clearBtn.type = 'button';
    clearBtn.textContent = 'Clear';
    clearBtn.addEventListener('click', () => {
      onChangeEditor.setValue('');
      updateEntryOnChange(index, '');
    });
    actions.appendChild(saveBtn);
    actions.appendChild(clearBtn);
    onChangeBody.appendChild(actions);
    detailDrawerBody.appendChild(onChangeSection.field);
    if (entry.isButton) {
      const hasExecuteScript = !!(entry.buttonExecute && String(entry.buttonExecute).trim());
      const executeSection = buildCollapsibleField('Button Execute Script', hasExecuteScript, 'executeOpen');
      const execBody = executeSection.body;
      const execWarnings = document.createElement('div');
      execWarnings.className = 'lua-warnings';
      execWarnings.hidden = true;
      const execEditor = createLuaEditor({
        value: entry.buttonExecute || '',
        placeholder: 'Lua script executed when this button fires.',
        onBlur: (value) => {
          updateEntryButtonExecute(index, value, { silent: true, skipHistory: true });
          renderLuaWarnings(execWarnings, validateLuaBasic(value));
        },
        getToolNames: () => state.parseResult?.luaToolNames,
        getControlNames: () => state.parseResult?.luaControlNames,
        getToolTypes: () => state.parseResult?.luaToolTypes,
      });
      const applyLauncherScript = (script) => {
        const val = script || '';
        execEditor.setValue(val);
        updateEntryButtonExecute(index, val, { silent: true, skipHistory: true });
      };
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
      execBody.appendChild(execEditor.wrapper);
      execBody.appendChild(execWarnings);
      renderLuaWarnings(execWarnings, validateLuaBasic(execEditor.getValue()));
      execBody.appendChild(buildPickRow(execEditor));
      const execActions = document.createElement('div');
      execActions.className = 'detail-actions';
      const execSave = document.createElement('button');
      execSave.type = 'button';
      execSave.textContent = 'Save';
      execSave.addEventListener('click', () => {
        updateEntryButtonExecute(index, execEditor.getValue());
      });
      const execClear = document.createElement('button');
      execClear.type = 'button';
      execClear.textContent = 'Clear';
      let clearPointerHandled = false;
      const markClearPointer = () => {
        clearPointerHandled = true;
        setTimeout(() => { clearPointerHandled = false; }, 300);
      };
      const runExecClear = () => {
        const hadValue = !!(entry.buttonExecute && String(entry.buttonExecute).trim());
        execEditor.setValue('');
        if (hadValue) {
          pushHistory('edit button execute');
        }
        updateEntryButtonExecute(index, '', { silent: true, skipHistory: true });
        if (state.parseResult && entry.sourceOp && entry.source) {
          const key = `${entry.sourceOp}.${entry.source}`;
          if (state.parseResult.insertClickedKeys instanceof Set) {
            state.parseResult.insertClickedKeys.delete(key);
          }
          if (state.parseResult.buttonExactInsert instanceof Set) {
            state.parseResult.buttonExactInsert.delete(key);
          }
        }
        entry._ui = entry._ui || {};
        entry._ui.executeOpen = true;
        renderDetailDrawer(index);
      };
      execClear.addEventListener('pointerdown', () => {
        markClearPointer();
        runExecClear();
      });
      execClear.addEventListener('click', () => {
        if (clearPointerHandled) {
          clearPointerHandled = false;
          return;
        }
        runExecClear();
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
  }

  function updateEntryOnChange(index, script, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    const newVal = script != null ? String(script) : '';
    if (entry.onChange === newVal) return;
    const silent = !!opts.silent;
    const skipHistory = !!opts.skipHistory;
    if (!skipHistory) pushHistory('edit on-change');
    entry.onChange = newVal;
    if (!silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
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
    if (!skipHistory) pushHistory('edit button execute');
    entry.buttonExecute = newVal;
    if (state.parseResult && entry.sourceOp && entry.source) {
      const key = `${entry.sourceOp}.${entry.source}`;
      if (newVal.trim()) {
        if (!(state.parseResult.buttonExactInsert instanceof Set)) state.parseResult.buttonExactInsert = new Set();
        state.parseResult.buttonExactInsert.add(key);
      } else {
        if (state.parseResult.buttonExactInsert instanceof Set) state.parseResult.buttonExactInsert.delete(key);
      }
    }
    if (!silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    markContentDirty();
  }

  function normalizeMetaValue(value) {
    if (value == null) return null;
    const trimmed = String(value).trim();
    return trimmed.length ? trimmed : null;
  }

  function inferEntryDataType(entry) {
    try {
      const explicit = normalizeMetaValue(
        entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType
      );
      if (explicit) return explicit;
      const inputControl = normalizeInputControlValue(
        entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
      );
      const lower = String(inputControl || '').toLowerCase();
      if (/texteditcontrol/.test(lower) || /styledtext/i.test(String(entry?.source || ''))) return '"Text"';
      if (/checkbox/.test(lower)) return '"Boolean"';
      if (/combo|slider|screw|button|label|separator/.test(lower)) return '"Number"';
      if (POINT_INPUT_CONTROLS.includes(inputControl)) return '"Point"';
      return null;
    } catch (_) {
      return null;
    }
  }

  function isTextControl(entry) {
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    if (inputControl && /text/i.test(inputControl)) return true;
    if (String(entry?.source || '').toLowerCase().includes('styledtext')) return true;
    if (String(entry?.source || '') === 'Text') {
      const toolTypes = state.parseResult?.luaToolTypes;
      const toolType = (toolTypes && entry?.sourceOp && toolTypes.get)
        ? toolTypes.get(entry.sourceOp)
        : '';
      if (String(toolType || '').toLowerCase().includes('follower')) return true;
    }
    const dataType = (entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType || '')
      .replace(/"/g, '')
      .trim()
      .toLowerCase();
    if (!dataType) return false;
    if (dataType === 'text' || dataType === 'string' || dataType.includes('text')) return true;
    const defaultRaw = entry?.controlMeta?.defaultValue ?? entry?.controlMetaOriginal?.defaultValue;
    const defaultStr = defaultRaw != null ? String(defaultRaw).trim() : '';
    if (defaultStr.toLowerCase().includes('styledtext')) return true;
    if (defaultStr.startsWith('"') && defaultStr.endsWith('"')) return true;
    return false;
  }

  function hydrateTextControlMeta(entry) {
    try {
      if (!entry || !state.parseResult || !state.originalText) return;
      if (!isTextControl(entry)) return;
      const hasLines = entry?.controlMeta?.textLines != null || entry?.controlMetaOriginal?.textLines != null;
      const hasInput = entry?.controlMeta?.inputControl != null || entry?.controlMetaOriginal?.inputControl != null;
      if (hasLines && hasInput) return;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult) || {
        groupOpenIndex: 0,
        groupCloseIndex: state.originalText.length
      };
      let toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) toolBlock = findToolBlockAnywhere(state.originalText, entry.sourceOp);
      if (!toolBlock) return;
      let controlBlock = findControlBlockInTool(state.originalText, toolBlock.open, toolBlock.close, entry.source);
      if (!controlBlock) {
        const toolUc = findUserControlsInTool(state.originalText, toolBlock.open, toolBlock.close);
        if (toolUc) controlBlock = findControlBlockInUc(state.originalText, toolUc.open, toolUc.close, entry.source);
      }
      if (!controlBlock) return;
      const body = state.originalText.slice(controlBlock.open + 1, controlBlock.close);
      const meta = entry.controlMeta || {};
      const metaOrig = entry.controlMetaOriginal || {};
      const inputControl = extractControlPropValue(body, 'INPID_InputControl');
      const textLines = extractControlPropValue(body, 'TEC_Lines');
      if (inputControl != null) {
        meta.inputControl = meta.inputControl || inputControl;
        metaOrig.inputControl = metaOrig.inputControl || inputControl;
      }
      if (textLines != null) {
        meta.textLines = meta.textLines || textLines;
        metaOrig.textLines = metaOrig.textLines || textLines;
      }
      entry.controlMeta = meta;
      entry.controlMetaOriginal = metaOrig;
    } catch (_) {}
  }

  function isCheckboxControl(entry) {
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    if (inputControl && /checkbox/i.test(inputControl)) return true;
    const dataType = (entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType || '')
      .replace(/"/g, '')
      .trim()
      .toLowerCase();
    if (dataType.includes('bool')) return true;
    return false;
  }

  function isComboControl(entry) {
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    return !!(inputControl && /combo/i.test(inputControl));
  }

  function isMultiButtonControl(entry) {
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    return !!(inputControl && /multibutton/i.test(inputControl));
  }

  function isChoiceControl(entry) {
    return isComboControl(entry) || isMultiButtonControl(entry);
  }

  function isMaskToolTypeForPath(type) {
    try {
      const t = String(type || '').trim().toLowerCase();
      if (!t) return false;
      return /(?:multipoly|mask|polygon|polyline|bspline|spolygon|spolyline|sbspline)/.test(t);
    } catch (_) {
      return false;
    }
  }

  function isPathControlId(id) {
    try {
      const key = String(id || '').trim().toLowerCase();
      if (!key) return false;
      return /(?:polyline|polygon|bspline|spline|path)/.test(key);
    } catch (_) {
      return false;
    }
  }

  function isPathEntryForDrawMode(entry) {
    try {
      if (!entry || entry.isLabel || entry.isButton) return false;
      if (!isPathControlId(entry.source)) return false;
      const map = state.parseResult?.luaToolTypes;
      const toolType = (map && typeof map.get === 'function' && entry.sourceOp) ? map.get(entry.sourceOp) : '';
      if (!toolType) return true;
      return isMaskToolTypeForPath(toolType);
    } catch (_) {
      return false;
    }
  }

  function parsePathDrawModeMetaFromLegacyScript(script) {
    try {
      const raw = String(script || '');
      if (!raw) return null;
      const hasMarker = raw.includes(FMR_PATH_DRAW_MODE_META_MARKER) || raw.includes(FMR_PATH_DRAW_MODE_LEGACY_SCRIPT_MARKER);
      if (!hasMarker) return null;
      const sourceOpMatch = raw.match(/--\s*sourceOp\s*=\s*([^\r\n]+)/i);
      const targetMatch = raw.match(/--\s*target\s*=\s*([^\r\n]+)/i);
      const sourceOp = String(sourceOpMatch?.[1] || '').trim();
      const target = String(targetMatch?.[1] || '').trim();
      if (!sourceOp || !target) return null;
      return { sourceOp, target, legacy: true };
    } catch (_) {
      return null;
    }
  }

  function unquoteSettingValue(raw) {
    try {
      if (raw == null) return '';
      const text = String(raw).trim();
      if (!text) return '';
      if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
        return text.slice(1, -1).trim();
      }
      return text;
    } catch (_) {
      return '';
    }
  }

  function findEntryControlBlockBoundsInText(text, result, entry) {
    try {
      if (!text || !result || !entry || !entry.source) return null;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return null;
      let toolBlock = null;
      if (entry.sourceOp) {
        toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
        if (!toolBlock) toolBlock = findToolBlockAnywhere(text, entry.sourceOp);
      }
      if (toolBlock) {
        let cb = findControlBlockInTool(text, toolBlock.open, toolBlock.close, entry.source);
        if (!cb) {
          const toolUc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
          if (toolUc) cb = findControlBlockInUc(text, toolUc.open, toolUc.close, entry.source);
        }
        if (cb) return cb;
      }
      const gcb = findGroupUserControlBlockById(text, result, entry.source);
      if (gcb) return gcb;
      return null;
    } catch (_) {
      return null;
    }
  }

  function parsePathDrawModeMetaFromEntry(entry, text = state.originalText, result = state.parseResult) {
    try {
      if (!entry) return null;
      const metaSourceOp = unquoteSettingValue(
        entry?.controlMeta?.fmrPathDrawModeSourceOp ??
        entry?.controlMetaOriginal?.fmrPathDrawModeSourceOp
      );
      const metaTarget = unquoteSettingValue(
        entry?.controlMeta?.fmrPathDrawModeTarget ??
        entry?.controlMetaOriginal?.fmrPathDrawModeTarget
      );
      if (metaSourceOp && metaTarget) return { sourceOp: metaSourceOp, target: metaTarget, legacy: false };

      const cb = findEntryControlBlockBoundsInText(text, result, entry);
      if (cb && text) {
        const body = text.slice(cb.open + 1, cb.close);
        const fromPropSourceOp = unquoteSettingValue(extractControlPropValue(body, FMR_PATH_DRAW_MODE_SOURCEOP_PROP));
        const fromPropTarget = unquoteSettingValue(extractControlPropValue(body, FMR_PATH_DRAW_MODE_TARGET_PROP));
        if (fromPropSourceOp && fromPropTarget) return { sourceOp: fromPropSourceOp, target: fromPropTarget, legacy: false };
      }
      return parsePathDrawModeMetaFromLegacyScript(entry.onChange || '');
    } catch (_) {
      return null;
    }
  }

  function extractFirstPolylineBodyFromText(segment) {
    try {
      const source = String(segment || '');
      if (!source) return '';
      const re = /Value\s*=\s*Polyline\s*\{/ig;
      let match = null;
      let fallback = '';
      while ((match = re.exec(source))) {
        const open = source.indexOf('{', match.index);
        if (open < 0) continue;
        const close = findMatchingBrace(source, open);
        if (close <= open) continue;
        const body = source.slice(open + 1, close);
        if (!fallback) fallback = body;
        if (/Points\s*=\s*\{/i.test(body)) return body;
      }
      return fallback;
    } catch (_) {
      return '';
    }
  }

  function extractPathPolylineBodyForTarget(text, result, targetSourceOp, targetSource) {
    try {
      if (!text || !result || !targetSourceOp || !targetSource) return '';
      const bounds = locateMacroGroupBounds(text, result);
      const visited = new Set();
      const walk = (sourceOp, source, depth = 0) => {
        const cleanSourceOp = unquoteSettingValue(sourceOp);
        const cleanSource = unquoteSettingValue(source);
        if (!cleanSourceOp || !cleanSource || depth > 8) return '';
        const key = `${cleanSourceOp}::${cleanSource}`;
        if (visited.has(key)) return '';
        visited.add(key);

        let toolBlock = bounds
          ? findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, cleanSourceOp)
          : null;
        if (!toolBlock) toolBlock = findToolBlockAnywhere(text, cleanSourceOp);
        if (!toolBlock) return '';

        const toolBody = text.slice(toolBlock.open + 1, toolBlock.close);
        const toolDirect = extractFirstPolylineBodyFromText(toolBody);
        if (String(cleanSource).trim().toLowerCase() === 'value' && toolDirect) return toolDirect;

        const inputs = findInputsInTool(text, toolBlock.open, toolBlock.close);
        if (!inputs) return '';
        const inputBlock = findInputBlockInInputs(text, inputs.open, inputs.close, cleanSource);
        if (!inputBlock) return '';

        const inputBody = text.slice(inputBlock.open + 1, inputBlock.close);
        const direct = extractFirstPolylineBodyFromText(inputBody);
        if (direct) return direct;

        const nextOp = unquoteSettingValue(extractControlPropValue(inputBody, 'SourceOp'));
        const nextSource = unquoteSettingValue(extractControlPropValue(inputBody, 'Source'));
        if (!nextOp || !nextSource) return '';
        return walk(nextOp, nextSource, depth + 1);
      };
      return walk(targetSourceOp, targetSource, 0);
    } catch (_) {
      return '';
    }
  }

  function detectPathTopologyForTarget(text, result, targetSourceOp, targetSource) {
    try {
      const body = extractPathPolylineBodyForTarget(text, result, targetSourceOp, targetSource);
      if (!body) return 'none';
      const hasPoints = /Points\s*=\s*\{[\s\S]*\{\s*X\s*=/i.test(body);
      if (!hasPoints) return 'none';
      const isClosed = /Closed\s*=\s*(?:true|1)\b/i.test(body);
      return isClosed ? 'closed' : 'open';
    } catch (_) {
      return 'none';
    }
  }

  function getPathDrawModeHelperSelectedIndex(entry, text, result) {
    try {
      if (!entry) return 0;
      let raw = '';
      const sourceOp = String(entry.sourceOp || '').trim();
      const source = String(entry.source || '').trim();
      if (sourceOp && source && text && result) {
        const bounds = locateMacroGroupBounds(text, result);
        let toolBlock = bounds
          ? findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, sourceOp)
          : null;
        if (!toolBlock) toolBlock = findToolBlockAnywhere(text, sourceOp);
        if (toolBlock) {
          const inputs = findInputsInTool(text, toolBlock.open, toolBlock.close);
          if (inputs) {
            const inputBlock = findInputBlockInInputs(text, inputs.open, inputs.close, source);
            if (inputBlock) {
              const inputBody = text.slice(inputBlock.open + 1, inputBlock.close);
              raw = String(extractControlPropValue(inputBody, 'Value') || '').trim();
            }
          }
        }
      }
      if (!raw) {
        raw = normalizeMetaValue(
          entry?.controlMeta?.defaultValue ??
          entry?.controlMetaOriginal?.defaultValue ??
          extractInstancePropValue(entry?.raw || '', 'Default')
        ) || '0';
      }
      const parsed = parseComboIndexValue(entry, raw);
      if (!Number.isFinite(parsed) || parsed == null) return 0;
      const max = PATH_DRAW_MODE_OPTIONS.length - 1;
      return Math.max(0, Math.min(max, Number(parsed)));
    } catch (_) {
      return 0;
    }
  }

  function normalizePathDrawModeIndexByTopology(index, topology) {
    const idx = Math.max(0, Math.min(PATH_DRAW_MODE_OPTIONS.length - 1, Number(index) || 0));
    if (topology === 'closed') {
      // Closed paths can only use Insert/Modify, Modify only, and Done.
      return idx < 2 ? 2 : idx;
    }
    if (topology === 'none') {
      // With no path drawn, only ClickAppend/Freehand are valid.
      return idx > 1 ? 0 : idx;
    }
    return idx;
  }

  function describePathDrawModeTopology(topology) {
    if (topology === 'closed') {
      return 'Closed path detected. Valid modes: Insert/Modify, Modify only, Done.';
    }
    if (topology === 'none') {
      return 'No path points detected. Valid modes: ClickAppend, Freehand.';
    }
    return 'Open path detected. All draw modes are valid.';
  }

  function normalizePathDrawModeName(value) {
    return String(value || '').replace(/[^A-Za-z0-9]/g, '').toLowerCase();
  }

  function pathDrawModeNameToIndex(value) {
    try {
      const normalized = normalizePathDrawModeName(value);
      if (!normalized) return null;
      const idx = PATH_DRAW_MODE_OPTIONS.findIndex((mode) => normalizePathDrawModeName(mode) === normalized);
      return idx >= 0 ? idx : null;
    } catch (_) {
      return null;
    }
  }

  function parsePathDrawModeIndex(raw) {
    try {
      if (raw == null) return null;
      let text = String(raw).trim();
      if (!text) return null;
      if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
        text = text.slice(1, -1).trim();
      }
      if (!text) return null;
      const numeric = Number(text);
      if (Number.isFinite(numeric)) {
        const rounded = Math.max(0, Math.min(PATH_DRAW_MODE_OPTIONS.length - 1, Math.round(numeric)));
        return rounded;
      }
      return pathDrawModeNameToIndex(text);
    } catch (_) {
      return null;
    }
  }

  function getPathDrawModeFromTool(text, result, sourceOp) {
    try {
      if (!text || !result || !sourceOp) return null;
      const bounds = locateMacroGroupBounds(text, result);
      let toolBlock = bounds
        ? findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, sourceOp)
        : null;
      if (!toolBlock) toolBlock = findToolBlockAnywhere(text, sourceOp);
      if (!toolBlock) return null;
      const toolBody = text.slice(toolBlock.open + 1, toolBlock.close);
      const rawMode = extractControlPropValue(toolBody, 'DrawMode');
      return parsePathDrawModeIndex(rawMode);
    } catch (_) {
      return null;
    }
  }

  function getPathDrawModeStoredIndex(entry, text = state.originalText, result = state.parseResult, options = {}) {
    try {
      if (!entry) return null;
      const onlyExplicit = options?.onlyExplicit === true;
      const hasMetaOverride = !!(entry.controlMeta && Object.prototype.hasOwnProperty.call(entry.controlMeta, 'fmrPathDrawModeIndex'));
      if (hasMetaOverride) {
        const fromMeta = parsePathDrawModeIndex(entry?.controlMeta?.fmrPathDrawModeIndex);
        if (Number.isFinite(fromMeta)) return fromMeta;
        if (onlyExplicit) return null;
      } else {
        const fromOriginal = parsePathDrawModeIndex(entry?.controlMetaOriginal?.fmrPathDrawModeIndex);
        if (Number.isFinite(fromOriginal)) return fromOriginal;
      }

      const cb = findEntryControlBlockBoundsInText(text, result, entry);
      if (cb && text) {
        const body = text.slice(cb.open + 1, cb.close);
        const fromProp = parsePathDrawModeIndex(extractControlPropValue(body, FMR_PATH_DRAW_MODE_INDEX_PROP));
        if (Number.isFinite(fromProp)) return fromProp;
      }

      if (onlyExplicit) return null;
      const fromTool = getPathDrawModeFromTool(text, result, String(entry.sourceOp || '').trim());
      if (Number.isFinite(fromTool)) return fromTool;
      return null;
    } catch (_) {
      return null;
    }
  }

  function setEntryPathDrawModeIndex(index, selectedIndex, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
    const entry = state.parseResult.entries[index];
    if (!entry || !isPathEntryForDrawMode(entry)) return false;
    const parsed = parsePathDrawModeIndex(selectedIndex);
    if (!Number.isFinite(parsed)) {
      if (!opts.silent) error('Invalid draw mode selection.');
      return false;
    }
    const nextValue = String(parsed);
    const currentValue = getPathDrawModeStoredIndex(entry, state.originalText, state.parseResult, { onlyExplicit: true });
    if (Number.isFinite(currentValue) && String(currentValue) === nextValue) return false;
    if (!opts.skipHistory) pushHistory('path draw mode');

    let changedText = false;
    const cb = findEntryControlBlockBounds(entry);
    if (cb && state.originalText) {
      const indent = (getLineIndent(state.originalText, cb.open) || '') + '\t';
      const eol = state.newline || '\n';
      let body = state.originalText.slice(cb.open + 1, cb.close);
      body = upsertControlProp(body, FMR_PATH_DRAW_MODE_INDEX_PROP, nextValue, indent, eol);
      const nextText = state.originalText.slice(0, cb.open + 1) + body + state.originalText.slice(cb.close);
      if (nextText !== state.originalText) {
        state.originalText = nextText;
        changedText = true;
      }
    }

    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.fmrPathDrawModeIndex = nextValue;
    if (
      entry.controlMetaOriginal.fmrPathDrawModeIndex == null ||
      String(entry.controlMetaOriginal.fmrPathDrawModeIndex).trim() === ''
    ) {
      entry.controlMetaOriginal.fmrPathDrawModeIndex = nextValue;
    }
    entry.controlMetaDirty = true;

    if (!opts.silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    if (changedText || !opts.skipMarkDirty) markContentDirty();
    return true;
  }

  function clearEntryPathDrawModeIndex(index, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
    const entry = state.parseResult.entries[index];
    if (!entry || !isPathEntryForDrawMode(entry)) return false;
    const currentValue = getPathDrawModeStoredIndex(entry, state.originalText, state.parseResult, { onlyExplicit: true });
    if (!Number.isFinite(currentValue)) return false;
    if (!opts.skipHistory) pushHistory('path draw mode');

    let changedText = false;
    const cb = findEntryControlBlockBounds(entry);
    if (cb && state.originalText) {
      const body = state.originalText.slice(cb.open + 1, cb.close);
      const rebuilt = removeControlProp(body, FMR_PATH_DRAW_MODE_INDEX_PROP);
      const nextText = state.originalText.slice(0, cb.open + 1) + rebuilt + state.originalText.slice(cb.close);
      if (nextText !== state.originalText) {
        state.originalText = nextText;
        changedText = true;
      }
    }

    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.fmrPathDrawModeIndex = null;
    entry.controlMetaOriginal.fmrPathDrawModeIndex = null;
    entry.controlMetaDirty = true;

    if (!opts.silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    if (changedText || !opts.skipMarkDirty) markContentDirty();
    return true;
  }

  function applyPathDrawModeExportPatches(text, result, eol, options = {}) {
    try {
      if (!text || !result || !Array.isArray(result.entries) || !result.entries.length) return text;
      if (options.safeEditExport === true) return text;
      let updated = text;
      const newline = eol || '\n';
      let direct = 0;
      let legacy = 0;
      let patched = 0;
      let clamped = 0;
      const patchedTargets = new Set();
      const applyForTarget = (targetSourceOp, targetSource, selectedIndex) => {
        const sourceOp = String(targetSourceOp || '').trim();
        const source = String(targetSource || '').trim();
        if (!sourceOp || !source) return false;
        const topology = detectPathTopologyForTarget(updated, result, sourceOp, source);
        const effectiveIndex = normalizePathDrawModeIndexByTopology(selectedIndex, topology);
        if (effectiveIndex !== selectedIndex) {
          clamped += 1;
          try {
            logDiag(`[Path DrawMode clamp] ${sourceOp}::${source} ${selectedIndex} -> ${effectiveIndex} (${topology})`);
          } catch (_) {}
        }
        const mode = PATH_DRAW_MODE_OPTIONS[effectiveIndex] || PATH_DRAW_MODE_OPTIONS[0];
        const bounds = locateMacroGroupBounds(updated, result);
        let toolBlock = bounds
          ? findToolBlockInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, sourceOp)
          : null;
        if (!toolBlock) toolBlock = findToolBlockAnywhere(updated, sourceOp);
        if (!toolBlock) return false;
        const toolIndent = (getLineIndent(updated, toolBlock.open) || '') + '\t';
        let toolBody = updated.slice(toolBlock.open + 1, toolBlock.close);
        toolBody = upsertControlProp(toolBody, 'DrawMode', `"${escapeQuotes(mode)}"`, toolIndent, newline);
        updated = updated.slice(0, toolBlock.open + 1) + toolBody + updated.slice(toolBlock.close);

        // Do not patch Input-level draw-mode props here. Path inputs often contain
        // nested Polyline literals, and generic block upserts can corrupt syntax.
        // Tool-level DrawMode is sufficient for persisted behavior in Fusion.
        patchedTargets.add(`${sourceOp}::${source}`);
        patched += 1;
        return true;
      };

      for (const targetEntry of result.entries) {
        if (!targetEntry || !isPathEntryForDrawMode(targetEntry)) continue;
        const explicitIndex = getPathDrawModeStoredIndex(targetEntry, updated, result, { onlyExplicit: true });
        if (!Number.isFinite(explicitIndex)) continue;
        direct += 1;
        applyForTarget(targetEntry.sourceOp, targetEntry.source, explicitIndex);
      }

      for (const helper of result.entries) {
        if (!helper || !isChoiceControl(helper)) continue;
        const meta = parsePathDrawModeMetaFromEntry(helper, updated, result);
        if (!meta) continue;
        const key = `${String(meta.sourceOp || '').trim()}::${String(meta.target || '').trim()}`;
        if (!key || patchedTargets.has(key)) continue;
        let helperIndex = getPathDrawModeHelperSelectedIndex(helper, updated, result);
        if (!Number.isFinite(helperIndex)) helperIndex = 0;
        legacy += 1;
        applyForTarget(meta.sourceOp, meta.target, helperIndex);
      }

      if (direct > 0 || legacy > 0) {
        try {
          logDiag(`[Path DrawMode export] direct=${direct}, legacy=${legacy}, patched=${patched}, clamped=${clamped}`);
        } catch (_) {}
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function appendEntryChunkToOrderedBlock(text, blockOpen, blockClose, chunk, eol) {
    try {
      if (!text || !chunk || !Number.isFinite(blockOpen) || !Number.isFinite(blockClose)) return text;
      const newline = eol || '\n';
      let cursor = Math.max(blockOpen + 1, blockClose - 1);
      while (cursor > blockOpen && /\s/.test(text[cursor])) cursor -= 1;
      const prev = text[cursor];
      const isEmpty = prev === '{';
      const needsComma = !isEmpty && prev !== ',';
      const needsLeadingNl = text[blockClose - 1] !== '\n' && text[blockClose - 1] !== '\r';
      const prefix = (needsComma ? ',' : '') + (needsLeadingNl ? newline : '');
      return text.slice(0, blockClose) + prefix + chunk + newline + text.slice(blockClose);
    } catch (_) {
      return text;
    }
  }

  function ensurePresetPathSwitchToolInGroup(text, result, switchToolName, defaultIndex, numberOfInputs, eol) {
    try {
      const cleanName = String(switchToolName || '').trim();
      if (!text || !result || !cleanName) return { text, toolBlock: null, created: false };
      let updated = text;
      const newline = eol || '\n';
      let bounds = locateMacroGroupBounds(updated, result);
      if (!bounds) return { text, toolBlock: null, created: false };
      let toolsBlock = findOrderedBlock(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
      if (!toolsBlock) return { text, toolBlock: null, created: false };
      let entries = parseOrderedBlockEntries(updated, toolsBlock.open, toolsBlock.close) || [];
      let existing = entries.find((entry) => String(entry?.name || '').trim() === cleanName);
      if (!existing) {
        const entryIndent = (getLineIndent(updated, toolsBlock.open) || '') + '\t';
        const switchChunk = [
          `${entryIndent}${cleanName} = SwitchPolyLine {`,
          `${entryIndent}\tCtrlWZoom = false,`,
          `${entryIndent}\tInputs = {`,
          `${entryIndent}\t\tNumberOfInputs = Input { Value = ${Math.max(2, Number(numberOfInputs) || 2)}, },`,
          `${entryIndent}\t\tSource = Input { Value = ${Math.max(0, Number(defaultIndex) || 0)}, },`,
          `${entryIndent}\t},`,
          `${entryIndent}},`,
        ].join(newline);
        updated = appendEntryChunkToOrderedBlock(updated, toolsBlock.open, toolsBlock.close, switchChunk, newline);
        bounds = locateMacroGroupBounds(updated, result) || bounds;
        toolsBlock = findOrderedBlock(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
        if (!toolsBlock) return { text: updated, toolBlock: null, created: true };
        entries = parseOrderedBlockEntries(updated, toolsBlock.open, toolsBlock.close) || [];
        existing = entries.find((entry) => String(entry?.name || '').trim() === cleanName);
        if (!existing) return { text: updated, toolBlock: null, created: true };
        return {
          text: updated,
          toolBlock: { open: existing.blockOpen, close: existing.blockClose },
          created: true,
        };
      }
      return {
        text: updated,
        toolBlock: { open: existing.blockOpen, close: existing.blockClose },
        created: false,
      };
    } catch (_) {
      return { text, toolBlock: null, created: false };
    }
  }

  function upsertPresetPathSwitchInputValue(text, result, toolName, toolBlock, inputName, valueLiteral, eol) {
    try {
      if (!text || !result || !toolBlock || !inputName) return { text, toolBlock, changed: false };
      const cleanToolName = String(toolName || '').trim();
      if (!cleanToolName) return { text, toolBlock, changed: false };
      const newline = eol || '\n';
      let updated = text;
      const ensured = ensureInputsBlockInToolBlock(updated, toolBlock, newline);
      updated = ensured.text;
      let currentToolBlock = ensured.toolBlock || toolBlock;
      let inputsBlock = ensured.inputsBlock || findInputsInTool(updated, currentToolBlock.open, currentToolBlock.close);
      if (!inputsBlock) return { text: updated, toolBlock: currentToolBlock, changed: false };
      let inputBlock = findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, inputName);
      let changed = false;
      if (!inputBlock) {
        const itemIndent = (getLineIndent(updated, inputsBlock.open) || '') + '\t';
        const inputChunk = [
          `${itemIndent}${inputName} = Input {`,
          `${itemIndent}\tValue = ${valueLiteral},`,
          `${itemIndent}},`,
        ].join(newline);
        updated = appendEntryChunkToOrderedBlock(updated, inputsBlock.open, inputsBlock.close, inputChunk, newline);
        changed = true;
      } else {
        const indent = (getLineIndent(updated, inputBlock.open) || '') + '\t';
        let body = updated.slice(inputBlock.open + 1, inputBlock.close);
        const nextBody = upsertControlProp(body, 'Value', valueLiteral, indent, newline);
        if (nextBody !== body) {
          updated = updated.slice(0, inputBlock.open + 1) + nextBody + updated.slice(inputBlock.close);
          changed = true;
        }
      }
      const group = locateMacroGroupBounds(updated, result);
      if (group) {
        const refreshed = findToolBlockInGroup(updated, group.groupOpenIndex, group.groupCloseIndex, cleanToolName);
        if (refreshed) currentToolBlock = refreshed;
      }
      return { text: updated, toolBlock: currentToolBlock, changed };
    } catch (_) {
      return { text, toolBlock, changed: false };
    }
  }

  function rewirePresetPathTargetInputToSwitch(text, result, sourceOp, sourceInput, switchToolName, eol) {
    try {
      const toolName = String(sourceOp || '').trim();
      const inputName = String(sourceInput || '').trim();
      const switchName = String(switchToolName || '').trim();
      if (!text || !result || !toolName || !inputName || !switchName) return { text, changed: false };
      const newline = eol || '\n';
      let updated = text;
      let bounds = locateMacroGroupBounds(updated, result);
      if (!bounds) return { text, changed: false };
      let toolBlock = findToolBlockInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
      if (!toolBlock) return { text, changed: false };
      const ensured = ensureInputsBlockInToolBlock(updated, toolBlock, newline);
      updated = ensured.text;
      toolBlock = ensured.toolBlock || toolBlock;
      let inputsBlock = ensured.inputsBlock || findInputsInTool(updated, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return { text: updated, changed: false };
      let inputBlock = findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, inputName);
      if (!inputBlock) {
        updated = insertToolInputStubBlock(updated, inputsBlock.open, inputsBlock.close, inputName, newline);
        bounds = locateMacroGroupBounds(updated, result) || bounds;
        toolBlock = findToolBlockInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName) || toolBlock;
        inputsBlock = findInputsInTool(updated, toolBlock.open, toolBlock.close) || inputsBlock;
        inputBlock = findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, inputName);
      }
      if (!inputBlock) return { text: updated, changed: false };
      const indent = (getLineIndent(updated, inputBlock.open) || '') + '\t';
      let body = updated.slice(inputBlock.open + 1, inputBlock.close);
      const original = body;
      body = removeControlProp(body, 'Value');
      body = upsertControlProp(body, 'SourceOp', `"${escapeQuotes(switchName)}"`, indent, newline);
      body = upsertControlProp(body, 'Source', '"Output"', indent, newline);
      if (body === original) return { text: updated, changed: false };
      updated = updated.slice(0, inputBlock.open + 1) + body + updated.slice(inputBlock.close);
      return { text: updated, changed: true };
    } catch (_) {
      return { text, changed: false };
    }
  }

  function applyPresetPathSwitchExportPatches(text, result, presetRuntime, eol, options = {}) {
    try {
      if (!text || !result || options?.safeEditExport === true) return text;
      const plans = Array.isArray(presetRuntime?.pathSwitchPlan) ? presetRuntime.pathSwitchPlan : [];
      if (!plans.length) return text;
      let updated = text;
      const newline = eol || '\n';
      let createdTools = 0;
      let rewiredTargets = 0;
      let upsertedInputs = 0;
      plans.forEach((plan) => {
        try {
          const switchName = String(plan?.switchToolName || '').trim();
          const literals = Array.isArray(plan?.literals) ? plan.literals : [];
          if (!switchName || !literals.length) return;
          const ensuredTool = ensurePresetPathSwitchToolInGroup(
            updated,
            result,
            switchName,
            plan?.defaultIndex,
            literals.length,
            newline
          );
          updated = ensuredTool.text;
          if (ensuredTool.created) createdTools += 1;
          let switchToolBlock = ensuredTool.toolBlock;
          if (!switchToolBlock) return;
          const countWrite = upsertPresetPathSwitchInputValue(
            updated,
            result,
            switchName,
            switchToolBlock,
            'NumberOfInputs',
            String(Math.max(2, literals.length || 2)),
            newline
          );
          updated = countWrite.text;
          switchToolBlock = countWrite.toolBlock || switchToolBlock;
          if (countWrite.changed) upsertedInputs += 1;
          const sourceWrite = upsertPresetPathSwitchInputValue(
            updated,
            result,
            switchName,
            switchToolBlock,
            'Source',
            String(Math.max(0, Number(plan?.defaultIndex) || 0)),
            newline
          );
          updated = sourceWrite.text;
          switchToolBlock = sourceWrite.toolBlock || switchToolBlock;
          if (sourceWrite.changed) upsertedInputs += 1;
          literals.forEach((literal, idx) => {
            const inputWrite = upsertPresetPathSwitchInputValue(
              updated,
              result,
              switchName,
              switchToolBlock,
              `Input${idx}`,
              String(literal || 'Polyline {}'),
              newline
            );
            updated = inputWrite.text;
            switchToolBlock = inputWrite.toolBlock || switchToolBlock;
            if (inputWrite.changed) upsertedInputs += 1;
          });
          const rewired = rewirePresetPathTargetInputToSwitch(
            updated,
            result,
            plan?.sourceOp,
            plan?.sourceInput,
            switchName,
            newline
          );
          updated = rewired.text;
          if (rewired.changed) rewiredTargets += 1;
        } catch (_) {}
      });
      try {
        if (typeof diagnosticsController?.isEnabled === 'function' && diagnosticsController.isEnabled()) {
          logDiag?.(`[Preset path switch] plans=${plans.length}, tools=${createdTools}, inputs=${upsertedInputs}, rewired=${rewiredTargets}`);
        }
      } catch (_) {}
      return updated;
    } catch (_) {
      return text;
    }
  }

  function getChoiceListPropForEntry(entry) {
    return isMultiButtonControl(entry) ? 'MBTNC_AddButton' : 'CCS_AddString';
  }

  function isIntegerControl(entry) {
    const integerFlag = entry?.controlMeta?.integer ?? entry?.controlMetaOriginal?.integer;
    if (integerFlag != null) {
      const normalized = String(integerFlag).replace(/"/g, '').trim().toLowerCase();
      if (normalized === 'true' || normalized === '1') return true;
      if (normalized === 'false' || normalized === '0') return false;
    }
    const dataType = (entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType || '')
      .replace(/"/g, '')
      .trim()
      .toLowerCase();
    if (dataType.includes('int')) return true;
    return false;
  }

  function isCsvLinkableControl(entry) {
    if (!entry || entry.isButton || entry.isLabel) return false;
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    if (inputControl && POINT_INPUT_CONTROLS.includes(inputControl)) return false;
    if (isTextControl(entry)) return true;
    if (isCheckboxControl(entry) || isChoiceControl(entry)) return true;
    const dataType = (entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType || '')
      .replace(/"/g, '')
      .trim()
      .toLowerCase();
    if (dataType.includes('number') || dataType.includes('float') || dataType.includes('int')) return true;
    if (inputControl) {
      const lower = inputControl.toLowerCase();
      if (lower.includes('slider') || lower.includes('screw')) return true;
    }
    const def = entry?.controlMeta?.defaultValue ?? entry?.controlMetaOriginal?.defaultValue;
    if (def != null && /^-?\d+(\.\d+)?$/.test(String(def).trim())) return true;
    if (!dataType) return true;
    return false;
  }

  function parseCsvBooleanValue(raw) {
    if (raw == null) return null;
    const cleaned = String(raw).trim().toLowerCase();
    if (!cleaned) return null;
    if (['true', 'yes', 'y', 'on'].includes(cleaned)) return 1;
    if (['false', 'no', 'n', 'off'].includes(cleaned)) return 0;
    if (cleaned === '1') return 1;
    if (cleaned === '0') return 0;
    return null;
  }

  function parseCsvNumericValue(raw) {
    if (raw == null) return null;
    let text = String(raw).trim();
    if (!text) return null;
    if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
      text = text.slice(1, -1).trim();
    }
    if (!text) return null;
    const paren = text.startsWith('(') && text.endsWith(')');
    if (paren) text = `-${text.slice(1, -1)}`;
    text = text.replace(/[$€£¥₩₹]/g, '').replace(/,/g, '').replace(/\s+/g, '');
    if (!text) return null;
    const numeric = /^[-+]?(\d+(\.\d+)?|\.\d+)([eE][-+]?\d+)?$/.test(text);
    if (!numeric) return null;
    const num = Number(text);
    return Number.isFinite(num) ? num : null;
  }

  function getChoiceOptions(entry) {
    try {
      const fromMeta = entry?.controlMeta?.choiceOptions;
      if (Array.isArray(fromMeta) && fromMeta.length) return fromMeta;
      const fromOriginal = entry?.controlMetaOriginal?.choiceOptions;
      if (Array.isArray(fromOriginal) && fromOriginal.length) return fromOriginal;
      if (entry && state.originalText) {
        const cb = findEntryControlBlockBounds(entry);
        if (cb) {
          const body = state.originalText.slice(cb.open + 1, cb.close);
          const parsed = extractChoiceOptionsFromBody(body);
          if (Array.isArray(parsed) && parsed.length) {
            entry.controlMeta = entry.controlMeta || {};
            entry.controlMetaOriginal = entry.controlMetaOriginal || {};
            entry.controlMeta.choiceOptions = [...parsed];
            if (!Array.isArray(entry.controlMetaOriginal.choiceOptions) || !entry.controlMetaOriginal.choiceOptions.length) {
              entry.controlMetaOriginal.choiceOptions = [...parsed];
            }
            return parsed;
          }
        }
      }
      return [];
    } catch (_) {
      return [];
    }
  }

  function getMultiButtonShowBasic(entry) {
    try {
      if (!entry || !isMultiButtonControl(entry)) return false;
      const fromMeta = entry?.controlMeta?.multiButtonShowBasic;
      if (fromMeta != null && String(fromMeta).trim() !== '') {
        return parseControlBooleanValue(fromMeta, false);
      }
      const fromOriginal = entry?.controlMetaOriginal?.multiButtonShowBasic;
      if (fromOriginal != null && String(fromOriginal).trim() !== '') {
        return parseControlBooleanValue(fromOriginal, false);
      }
      const cb = findEntryControlBlockBounds(entry);
      if (!cb || !state.originalText) return false;
      const body = state.originalText.slice(cb.open + 1, cb.close);
      const raw = extractControlPropValue(body, 'MBTNC_ShowBasicButton');
      return parseControlBooleanValue(raw, false);
    } catch (_) {
      return false;
    }
  }

  function parseComboIndexValue(entry, raw) {
    if (raw == null) return null;
    let text = String(raw).trim();
    if (!text) return null;
    if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
      text = text.slice(1, -1).trim();
    }
    if (!text) return null;
    const num = Number(text);
    if (!Number.isFinite(num)) return null;
    const rounded = Math.round(num);
    const clamped = clampNumericToEntry(entry, rounded);
    return Number.isFinite(clamped) ? clamped : null;
  }

  function getComboOptionLabel(entry, index) {
    try {
      const options = getChoiceOptions(entry);
      if (!options.length) return '';
      if (!Number.isFinite(index)) return '';
      const idx = Math.trunc(index);
      if (idx < 0 || idx >= options.length) return '';
      return String(options[idx] == null ? '' : options[idx]);
    } catch (_) {
      return '';
    }
  }

  function clampNumericToEntry(entry, value) {
    let next = value;
    const minRaw = entry?.controlMeta?.minAllowed ?? entry?.controlMetaOriginal?.minAllowed;
    const maxRaw = entry?.controlMeta?.maxAllowed ?? entry?.controlMetaOriginal?.maxAllowed;
    const min = minRaw != null ? Number(String(minRaw).replace(/"/g, '').trim()) : null;
    const max = maxRaw != null ? Number(String(maxRaw).replace(/"/g, '').trim()) : null;
    if (isChoiceControl(entry)) {
      const options = getChoiceOptions(entry);
      if (options.length) {
        next = Math.max(0, Math.min(options.length - 1, next));
      }
    }
    if (Number.isFinite(min)) next = Math.max(min, next);
    if (Number.isFinite(max)) next = Math.min(max, next);
    return next;
  }

  function coerceCsvNumericValue(entry, raw) {
    const boolVal = parseCsvBooleanValue(raw);
    if (boolVal != null) {
      if (isCheckboxControl(entry)) return boolVal ? 1 : 0;
      return boolVal;
    }
    const num = parseCsvNumericValue(raw);
    if (!Number.isFinite(num)) return null;
    let next = num;
    if (isCheckboxControl(entry)) next = num ? 1 : 0;
    if (isChoiceControl(entry) || isIntegerControl(entry)) next = Math.round(next);
    next = clampNumericToEntry(entry, next);
    return next;
  }

  function clamp01(value) {
    const num = Number(value);
    if (!Number.isFinite(num)) return 0;
    return Math.max(0, Math.min(1, num));
  }

  function parseHexColor(raw) {
    if (!raw) return null;
    let hex = String(raw).trim();
    if (!hex) return null;
    if ((hex.startsWith('"') && hex.endsWith('"')) || (hex.startsWith("'") && hex.endsWith("'"))) {
      hex = hex.slice(1, -1).trim();
    }
    const match = hex.match(/#?([0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{4}|[0-9a-fA-F]{3})/);
    if (match && match[1]) {
      hex = match[1];
    }
    if (hex.startsWith('#')) hex = hex.slice(1);
    if (hex.length === 3 || hex.length === 4) {
      hex = hex.split('').map(ch => ch + ch).join('');
    }
    if (hex.length !== 6 && hex.length !== 8) return null;
    if (!/^[0-9a-fA-F]+$/.test(hex)) return null;
    const r = parseInt(hex.slice(0, 2), 16) / 255;
    const g = parseInt(hex.slice(2, 4), 16) / 255;
    const b = parseInt(hex.slice(4, 6), 16) / 255;
    const hasAlpha = hex.length === 8;
    const a = hasAlpha ? (parseInt(hex.slice(6, 8), 16) / 255) : 1;
    return { r, g, b, a, hasAlpha };
  }

  function toHexByte(value) {
    const v = Math.round(clamp01(value) * 255);
    return v.toString(16).padStart(2, '0').toUpperCase();
  }

  function formatHexColor(color) {
    if (!color) return '';
    const r = toHexByte(color.r);
    const g = toHexByte(color.g);
    const b = toHexByte(color.b);
    const a = toHexByte(color.a);
    const includeAlpha = color.includeAlpha || (Number.isFinite(color.a) && Math.abs(color.a - 1) > 0.0001);
    return `#${r}${g}${b}${includeAlpha ? a : ''}`;
  }

  function stripDefaultQuotes(value) {
    if (value == null) return null;
    let str = String(value).trim();
    if (str.length >= 2 && str.startsWith('"') && str.endsWith('"')) {
      str = unescapeSettingString(str.slice(1, -1));
    }
    return str;
  }

  function unescapeSettingString(value) {
    return String(value || '')
      .replace(/\\\\/g, '\\')
      .replace(/\\"/g, '"')
      .replace(/\\n/g, '\n')
      .replace(/\\r/g, '\r')
      .replace(/\\t/g, '\t');
  }

  function stripOuterQuotes(value) {
    let str = String(value || '').trim();
    if (str.startsWith('"') && str.endsWith('"')) str = str.slice(1, -1);
    return str;
  }

  function isStyledTextValue(value) {
    const stripped = stripOuterQuotes(value);
    return /^StyledText\b/.test(stripped);
  }

  function extractStyledTextPlainText(value) {
    const stripped = stripOuterQuotes(value);
    if (!/^StyledText\b/.test(stripped)) return null;
    const match = stripped.match(/(?:Text|Value)\s*=\s*"((?:\\.|[^"])*)"/);
    if (!match) return null;
    return unescapeSettingString(match[1]);
  }

  function updateStyledTextValue(value, plainText) {
    const stripped = stripOuterQuotes(value);
    if (!/^StyledText\b/.test(stripped)) return value;
    const safeText = escapeQuotes(String(plainText || ''));
    let updated = stripped;
    if (/Text\s*=\s*"/.test(updated)) {
      updated = updated.replace(/Text\s*=\s*"((?:\\.|[^"])*)"/, `Text = "${safeText}"`);
    } else if (/Value\s*=\s*"/.test(updated)) {
      updated = updated.replace(/Value\s*=\s*"((?:\\.|[^"])*)"/, `Value = "${safeText}"`);
    }
    return `"${escapeSettingString(updated)}"`;
  }

  function buildStyledTextBlock(plainText) {
    const safeText = escapeSettingString(String(plainText || ''));
    return `StyledText { Value = "${safeText}" }`;
  }

  function updateToolInputBlockBody(body, entry, value, indent, eol) {
    try {
      const nextText = String(value || '');
      const isText = isTextControl(entry);
      const useStyled = isText && (String(entry?.source || '').toLowerCase().includes('styledtext') || /StyledText\s*\{/.test(body));
      if (useStyled) {
        const replacement = buildStyledTextBlock(nextText);
        if (/Value\s*=\s*StyledText\s*\{[\s\S]*?\}/i.test(body)) {
          return body.replace(/Value\s*=\s*StyledText\s*\{[\s\S]*?\}/i, `Value = ${replacement}`);
        }
        return setInstanceInputProp(body, 'Value', replacement, indent, eol || '\n');
      }
      if (isText) {
        const safe = `"${escapeSettingString(nextText)}"`;
        return setInstanceInputProp(body, 'Value', safe, indent, eol || '\n');
      }
      const raw = nextText.trim();
      let updated = body;
      updated = removeControlProp(updated, 'Expression');
      updated = removeControlProp(updated, 'SourceOp');
      updated = removeControlProp(updated, 'Source');
      return setInstanceInputProp(updated, 'Value', raw, indent, eol || '\n');
    } catch (_) {
      return body;
    }
  }

  function buildToolInputValueLiteral(entry, value) {
    try {
      const nextText = String(value == null ? '' : value);
      const isText = isTextControl(entry);
      const useStyled = isText && String(entry?.source || '').toLowerCase().includes('styledtext');
      if (useStyled) {
        return buildStyledTextBlock(nextText);
      }
      if (isText) {
        return `"${escapeSettingString(nextText)}"`;
      }
      const raw = nextText.trim();
      return raw || '0';
    } catch (_) {
      return String(value == null ? '' : value);
    }
  }

  function formatToolInputId(id) {
    try {
      const raw = String(id == null ? '' : id).trim();
      if (!raw) return '';
      if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(raw)) return raw;
      const normalized = normalizeId(raw);
      if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(normalized)) return normalized;
      return `["${escapeSettingString(normalized)}"]`;
    } catch (_) {
      return String(id == null ? '' : id);
    }
  }

  function insertToolInputBlock(text, inputsOpen, inputsClose, entry, value, eol) {
    try {
      if (!text || !entry || !entry.source) return text;
      const nl = eol || '\n';
      const itemIndent = (getLineIndent(text, inputsOpen) || '') + '\t';
      const propIndent = itemIndent + '\t';
      const inputId = formatToolInputId(entry.source);
      if (!inputId) return text;
      const valueLiteral = buildToolInputValueLiteral(entry, value);
      const block =
        `${itemIndent}${inputId} = Input {${nl}` +
        `${propIndent}Value = ${valueLiteral},${nl}` +
        `${itemIndent}},`;
      let cursor = Math.max(inputsOpen + 1, inputsClose - 1);
      while (cursor > inputsOpen && /\s/.test(text[cursor])) cursor -= 1;
      const prev = text[cursor];
      const isEmptyInputs = prev === '{';
      const needsSeparatorComma = !isEmptyInputs && prev !== ',';
      const needsLeadingNl = text[inputsClose - 1] !== '\n' && text[inputsClose - 1] !== '\r';
      const prefix = needsLeadingNl ? nl : '';
      const separator = needsSeparatorComma ? ',' : '';
      return text.slice(0, inputsClose) + separator + prefix + block + nl + text.slice(inputsClose);
    } catch (_) {
      return text;
    }
  }

  function setInputValueInToolText(text, entry, value, eol, resultRef) {
    try {
      if (!text || !entry || !entry.sourceOp || !entry.source) return text;
      const bounds = locateMacroGroupBounds(text, resultRef || state.parseResult);
      if (!bounds) return text;
      const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) return text;
      const inputsBlock = findInputsInTool(text, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return text;
      const inputBlock = findInputBlockInInputs(text, inputsBlock.open, inputsBlock.close, entry.source);
      if (!inputBlock) {
        return insertToolInputBlock(text, inputsBlock.open, inputsBlock.close, entry, value, eol);
      }
      const indent = (getLineIndent(text, inputBlock.open) || '') + '\t';
      let body = text.slice(inputBlock.open + 1, inputBlock.close);
      body = updateToolInputBlockBody(body, entry, value, indent, eol);
      return text.slice(0, inputBlock.open + 1) + body + text.slice(inputBlock.close);
    } catch (_) {
      return text;
    }
  }

  function setInstanceDefaultInText(text, entry, value, eol, resultRef) {
    try {
      if (!text || !entry || !entry.key) return text;
      const bounds = locateMacroGroupBounds(text, resultRef || state.parseResult);
      if (!bounds || !resultRef || !resultRef.inputs) return text;
      const inputs = resultRef.inputs;
      const inputBlock = findInstanceInputBlockInInputs(text, inputs.openIndex, inputs.closeIndex, entry.key);
      if (!inputBlock) return text;
      const indent = (getLineIndent(text, inputBlock.open) || '') + '\t';
      let body = text.slice(inputBlock.open + 1, inputBlock.close);
      const formatted = formatDefaultForStorage(entry, String(value));
      if (formatted == null) {
        body = removeInstanceInputProp(body, 'Default');
      } else {
        body = setInstanceInputProp(body, 'Default', formatted, indent, eol || '\n');
      }
      return text.slice(0, inputBlock.open + 1) + body + text.slice(inputBlock.close);
    } catch (_) {
      return text;
    }
  }

  function applyNumericDefaultToEntry(entry, value, eol) {
    if (!entry) return;
    const formatted = formatDefaultForStorage(entry, String(value));
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.defaultValue = formatted;
    entry.controlMetaDirty = true;
    applyDefaultToEntryRaw(entry, formatted, eol);
  }

  function applyHexToColorGroup({ entries, order, baseText, groupIndex, r, g, b, a, includeAlpha, eol, resultRef }) {
    if (!entries || !order) return baseText || '';
    const cgBlock = getColorGroupBlockByIndexFrom(entries, order, groupIndex);
    if (!cgBlock) return baseText || '';
    const components = getColorGroupComponentsFrom(entries, cgBlock);
    if (!components) return baseText || '';
    let updated = baseText || '';
    const applyComponent = (compIdx, value) => {
      if (!Number.isFinite(compIdx)) return;
      const compEntry = entries[compIdx];
      if (!compEntry) return;
      const val = String(value);
      updated = setInputValueInToolText(updated, compEntry, val, eol, resultRef);
      updated = setInstanceDefaultInText(updated, compEntry, val, eol, resultRef);
      applyNumericDefaultToEntry(compEntry, val, eol);
    };
    applyComponent(components.red, r);
    applyComponent(components.green, g);
    applyComponent(components.blue, b);
    if (includeAlpha && a != null) applyComponent(components.alpha, a);
    return updated;
  }

  function resolveTextDefaultSource(entry, value) {
    if (!entry || !isTextControl(entry)) return value;
    try {
      const inst = extractInstancePropValue(entry?.raw || '', 'Default');
      if (inst != null && String(inst).trim()) return inst;
      const toolValue = getInputValueFromTool(entry);
      if (toolValue != null && String(toolValue).trim()) return toolValue;
    } catch (_) {}
    return value;
  }

  function extractStyledTextBlockFromInput(body) {
    try {
      if (!body) return null;
      const match = body.match(/Value\s*=\s*StyledText\s*\{/i);
      if (!match || match.index == null) return null;
      const styledIdx = body.indexOf('StyledText', match.index);
      const openIdx = body.indexOf('{', styledIdx);
      if (openIdx < 0) return null;
      const closeIdx = findMatchingBrace(body, openIdx);
      if (closeIdx < 0) return null;
      return body.slice(styledIdx, closeIdx + 1);
    } catch (_) {
      return null;
    }
  }

  function extractStyledTextPlainTextFromInput(body) {
    try {
      if (!body) return null;
      const match = body.match(/Value\s*=\s*StyledText\s*\{[\s\S]*?Value\s*=\s*"((?:\\.|[^"])*)"/i);
      if (!match) return null;
      return unescapeSettingString(match[1]);
    } catch (_) {
      return null;
    }
  }

  function getInputValueFromTool(entry) {
    try {
      if (!entry || !entry.sourceOp || !entry.source || !state.originalText || !state.parseResult) return null;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) return null;
      const toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) return null;
      const inputsBlock = findInputsInTool(state.originalText, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return null;
      const inputBlock = findInputBlockInInputs(state.originalText, inputsBlock.open, inputsBlock.close, entry.source);
      if (!inputBlock) return null;
      const body = state.originalText.slice(inputBlock.open + 1, inputBlock.close);
      const styled = extractStyledTextBlockFromInput(body);
      if (styled) return styled;
      return extractControlPropValue(body, 'Value') || null;
    } catch (_) {
      return null;
    }
  }

  function formatDefaultForDisplay(entry, value) {
    const next = normalizeMetaValue(resolveTextDefaultSource(entry, value));
    if (next == null) return null;
    if (!isTextControl(entry)) return next;
    const styled = extractStyledTextPlainText(next);
    if (styled != null) return styled;
    return stripDefaultQuotes(next);
  }

  function formatDefaultForStorage(entry, value) {
    const next = normalizeMetaValue(value);
    if (next == null) return null;
    if (!isTextControl(entry)) {
      const numeric = coerceCsvNumericValue(entry, next);
      if (numeric == null) return null;
      return String(numeric);
    }
    const template = resolveTextDefaultSource(entry, entry?.controlMeta?.defaultValue ?? entry?.controlMetaOriginal?.defaultValue);
    if (template && isStyledTextValue(template)) {
      return updateStyledTextValue(template, next);
    }
    const stripped = stripDefaultQuotes(next);
    return `"${escapeQuotes(stripped)}"`;
  }

  function detectCsvDelimiter(text) {
    try {
      const line = String(text || '').split(/\r?\n/)[0] || '';
      const candidates = [',', '\t', ';', '|'];
      let best = ',';
      let bestCount = -1;
      candidates.forEach((ch) => {
        const count = (line.match(new RegExp(`\\${ch}`, 'g')) || []).length;
        if (count > bestCount) { best = ch; bestCount = count; }
      });
      return best || ',';
    } catch (_) { return ','; }
  }

  function findInstanceInputBlockInInputs(text, inputsOpen, inputsClose, key) {
    try {
      if (!text || key == null) return null;
      let i = inputsOpen + 1;
      while (i < inputsClose) {
        if (!isIdentStart(text[i])) { i++; continue; }
        const start = i;
        i++;
        while (i < inputsClose && isIdentPart(text[i])) i++;
        const ident = text.slice(start, i);
        if (ident !== key) continue;
        while (i < inputsClose && isSpace(text[i])) i++;
        if (text[i] !== '=') continue;
        i++;
        while (i < inputsClose && isSpace(text[i])) i++;
        if (text.slice(i, i + 13) !== 'InstanceInput') continue;
        i += 13;
        while (i < inputsClose && isSpace(text[i])) i++;
        if (text[i] !== '{') continue;
        const open = i;
        const close = findMatchingBrace(text, open);
        if (close < 0) return null;
        return { open, close };
      }
    } catch (_) {}
    return null;
  }

  function buildCsvPatchPlan(baseText, result, linkedEntries) {
    try {
      if (!baseText || !result || !Array.isArray(linkedEntries) || !linkedEntries.length) return null;
      const bounds = locateMacroGroupBounds(baseText, result);
      if (!bounds) return null;
      const inputs = result.inputs;
      const plan = [];
      const toolBlockCache = new Map();
      const inputsBlockCache = new Map();
      const inputBlockCache = new Map();
      const instanceBlockCache = new Map();
      const getToolBlockCached = (toolName) => {
        if (!toolName) return null;
        if (toolBlockCache.has(toolName)) return toolBlockCache.get(toolName);
        const block = findToolBlockInGroup(baseText, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName) || null;
        toolBlockCache.set(toolName, block);
        return block;
      };
      const getInputsBlockCached = (toolName, toolBlock) => {
        const key = toolName || '';
        if (inputsBlockCache.has(key)) return inputsBlockCache.get(key);
        const block = toolBlock ? (findInputsInTool(baseText, toolBlock.open, toolBlock.close) || null) : null;
        inputsBlockCache.set(key, block);
        return block;
      };
      const getInputBlockCached = (toolName, inputName, inputsBlock) => {
        const key = `${toolName || ''}::${inputName || ''}`;
        if (inputBlockCache.has(key)) return inputBlockCache.get(key);
        const block = inputsBlock
          ? (findInputBlockInInputs(baseText, inputsBlock.open, inputsBlock.close, inputName) || null)
          : null;
        inputBlockCache.set(key, block);
        return block;
      };
      const getInstanceBlockCached = (instanceKey) => {
        if (!instanceKey) return null;
        if (instanceBlockCache.has(instanceKey)) return instanceBlockCache.get(instanceKey);
        const block = (!inputs || typeof inputs.openIndex !== 'number' || typeof inputs.closeIndex !== 'number')
          ? null
          : (findInstanceInputBlockInInputs(baseText, inputs.openIndex, inputs.closeIndex, instanceKey) || null);
        instanceBlockCache.set(instanceKey, block);
        return block;
      };
      for (const entry of linkedEntries) {
        if (!entry || !entry.csvLink || !entry.csvLink.column) continue;
        if (isTextControl(entry)) {
          if (!entry.sourceOp || !entry.source) continue;
          const toolBlock = getToolBlockCached(entry.sourceOp);
          if (!toolBlock) continue;
          const inputsBlock = getInputsBlockCached(entry.sourceOp, toolBlock);
          if (!inputsBlock) continue;
          const inputBlock = getInputBlockCached(entry.sourceOp, entry.source, inputsBlock);
          if (!inputBlock) continue;
          const indent = (getLineIndent(baseText, inputBlock.open) || '') + '\t';
          plan.push({ kind: 'tool', entry, start: inputBlock.open, end: inputBlock.close, indent });
          continue;
        }
        if (!inputs || typeof inputs.openIndex !== 'number' || typeof inputs.closeIndex !== 'number') continue;
        if (!entry.key) continue;
        const inst = getInstanceBlockCached(entry.key);
        if (!inst) continue;
        const indent = (getLineIndent(baseText, inst.open) || '') + '\t';
        plan.push({ kind: 'instance', entry, start: inst.open, end: inst.close, indent });
      }
      return plan;
    } catch (_) {
      return null;
    }
  }

  function applyCsvRowEdits(baseText, plan, row, eol) {
    if (!baseText || !Array.isArray(plan) || !plan.length) return baseText;
    const edits = [];
      for (const patch of plan) {
        const entry = patch.entry;
        if (!entry || !entry.csvLink || !entry.csvLink.column) continue;
        const rawVal = row[entry.csvLink.column] != null ? String(row[entry.csvLink.column]) : '';
        if (patch.kind === 'tool') {
          const body = baseText.slice(patch.start + 1, patch.end);
          const next = updateToolInputBlockBody(body, entry, rawVal, patch.indent, eol);
          if (next !== body) edits.push({ start: patch.start + 1, end: patch.end, text: next });
        } else if (patch.kind === 'instance') {
          const body = baseText.slice(patch.start + 1, patch.end);
          const formatted = formatDefaultForStorage(entry, rawVal);
          if (formatted != null) {
            entry.controlMeta = entry.controlMeta || {};
            entry.controlMetaOriginal = entry.controlMetaOriginal || {};
            entry.controlMeta.defaultValue = formatted;
            entry.controlMetaDirty = true;
            applyDefaultToEntryRaw(entry, formatted, eol);
          }
          let next = body;
          if (formatted == null) next = removeInstanceInputProp(body, 'Default');
          else next = setInstanceInputProp(body, 'Default', formatted, patch.indent, eol || '\n');
          if (next !== body) edits.push({ start: patch.start + 1, end: patch.end, text: next });
        }
      }
    if (!edits.length) return baseText;
    edits.sort((a, b) => b.start - a.start);
    let out = baseText;
    for (const edit of edits) {
      out = out.slice(0, edit.start) + edit.text + out.slice(edit.end);
    }
    return out;
  }

  function deriveSafeMacroName(rowLabel, result) {
    const originalName = (result?.macroNameOriginal || result?.macroName || '').trim();
    const desired = String(rowLabel || '').trim();
    let safeName = sanitizeIdent(desired);
    if (!safeName) safeName = sanitizeIdent(originalName) || 'Macro';
    if (!isIdentStart(safeName[0])) safeName = `_${safeName}`;
    return safeName;
  }

  function parseCsvText(rawText) {
    const text = String(rawText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const delimiter = detectCsvDelimiter(text);
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i += 1) {
      const ch = text[i];
      const next = text[i + 1];
      if (inQuotes) {
        if (ch === '"' && next === '"') {
          field += '"';
          i += 1;
          continue;
        }
        if (ch === '"' && next !== '"') {
          inQuotes = false;
          continue;
        }
        field += ch;
        continue;
      }
      if (ch === '"') { inQuotes = true; continue; }
      if (ch === delimiter) {
        row.push(field);
        field = '';
        continue;
      }
      if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        continue;
      }
      field += ch;
    }
    if (field.length || row.length) {
      row.push(field);
      rows.push(row);
    }
    const cleanedRows = rows.filter(r => r.some(v => String(v || '').trim().length));
    if (!cleanedRows.length) return { headers: [], rows: [], delimiter };
    const rawHeaders = cleanedRows[0].map(v => String(v || '').trim());
    const headers = [];
    const seen = new Map();
    rawHeaders.forEach((h, idx) => {
      let base = h || `Column ${idx + 1}`;
      const key = base.toLowerCase();
      const count = (seen.get(key) || 0) + 1;
      seen.set(key, count);
      if (count > 1) base = `${base} ${count}`;
      headers.push(base);
    });
    const dataRows = cleanedRows.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = (r && r[idx] != null) ? String(r[idx]) : '';
      });
      return obj;
    });
    return { headers, rows: dataRows, delimiter };
  }

  function normalizeCsvUrl(raw) {
    let url = String(raw || '').trim();
    if (!url) return '';
    if (!/^[a-z][a-z0-9+.-]*:/i.test(url)) url = `https://${url}`;
    if (/\.csv(\?|#|$)/i.test(url)) return url;
    const match = url.match(/docs\.google\.com\/spreadsheets\/d\/([^/]+)/i);
    if (!match) return url;
    const id = match[1];
    const gidMatch = url.match(/[?&]gid=([0-9]+)/i);
    const gid = gidMatch ? gidMatch[1] : '';
    const base = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`;
    return gid ? `${base}&gid=${gid}` : base;
  }

  function setCsvData(data, sourceName) {
    if (!data || !Array.isArray(data.headers)) {
      state.csvData = null;
      updateDataMenuState();
      return;
    }
    state.csvData = {
      headers: data.headers || [],
      rows: data.rows || [],
      delimiter: data.delimiter || ',',
      sourceName: sourceName || '',
      nameColumn: data.nameColumn || null,
    };
    try { info(`CSV loaded: ${state.csvData.rows.length} rows.`); } catch (_) {}
    if (activeDetailEntryIndex != null) renderDetailDrawer(activeDetailEntryIndex);
    updateDataMenuState();
    syncDataLinkPanel();
  }

  function applyFmrDataLinkToEntries(result) {
    try {
      const link = result?.dataLink;
      if (!link || !Array.isArray(result?.entries)) return;
      const mappings = link.mappings || {};
      const overrides = link.overrides || {};
      const hexMappings = link.hex || {};
      const normalizeToolNameForLink = (name) => String(name || '').replace(/_\d+$/, '');
      const buildLinkLookup = (obj) => {
        const exact = new Map();
        const normalized = new Map();
        Object.keys(obj || {}).forEach((key) => {
          exact.set(key, obj[key]);
          const dot = key.indexOf('.');
          if (dot <= 0) return;
          const tool = key.slice(0, dot);
          const control = key.slice(dot + 1);
          const nk = `${normalizeToolNameForLink(tool)}.${control}`;
          if (!normalized.has(nk)) normalized.set(nk, obj[key]);
        });
        return { exact, normalized };
      };
      const mappingLookup = buildLinkLookup(mappings);
      const overridesLookup = buildLinkLookup(overrides);
      const hexLookup = buildLinkLookup(hexMappings);
      result.entries.forEach((entry) => {
        if (!entry || !entry.sourceOp || !entry.source) return;
        const key = `${entry.sourceOp}.${entry.source}`;
        const nkey = `${normalizeToolNameForLink(entry.sourceOp)}.${entry.source}`;
        const col = mappingLookup.exact.get(key) ?? mappingLookup.normalized.get(nkey);
        if (col) {
          entry.csvLink = { column: col };
          const override = overridesLookup.exact.get(key) ?? overridesLookup.normalized.get(nkey);
          if (override && override.rowMode) {
            entry.csvLink.rowMode = override.rowMode;
            if (Number.isFinite(override.rowIndex)) entry.csvLink.rowIndex = override.rowIndex;
            if (override.rowKey != null) entry.csvLink.rowKey = override.rowKey;
            if (override.rowValue != null) entry.csvLink.rowValue = override.rowValue;
          }
        }
        const hex = hexLookup.exact.get(key) ?? hexLookup.normalized.get(nkey);
        if (hex && hex.column) {
          entry.csvLinkHex = { column: hex.column };
          if (hex.rowMode) entry.csvLinkHex.rowMode = hex.rowMode;
          if (Number.isFinite(hex.rowIndex)) entry.csvLinkHex.rowIndex = hex.rowIndex;
          if (hex.rowKey != null) entry.csvLinkHex.rowKey = hex.rowKey;
          if (hex.rowValue != null) entry.csvLinkHex.rowValue = hex.rowValue;
        }
      });
      if (link.source) {
        if (!state.csvData) {
          state.csvData = { headers: [], rows: [], delimiter: ',', sourceName: link.source, nameColumn: link.nameColumn || null };
        } else {
          state.csvData.sourceName = link.source;
          if (link.nameColumn) state.csvData.nameColumn = link.nameColumn;
        }
        updateDataMenuState();
      }
    } catch (_) {}
  }

  function ensureInstanceInputControlGroup(raw, controlGroup, eol) {
    try {
      if (!raw || !Number.isFinite(controlGroup)) return raw;
      const open = raw.indexOf('{');
      if (open < 0) return raw;
      const close = findMatchingBrace(raw, open);
      if (close < 0) return raw;
      const indent = (getLineIndent(raw, open) || '') + '\t';
      let body = raw.slice(open + 1, close);
      body = setInstanceInputProp(body, 'ControlGroup', String(controlGroup), indent, eol || '\n');
      return raw.slice(0, open + 1) + body + raw.slice(close);
    } catch (_) {
      return raw;
    }
  }

  function hydrateFlipPairGroups(result, eol) {
    try {
      if (!result || !Array.isArray(result.entries) || !result.entries.length) return;
      const byTool = new Map();
      result.entries.forEach((entry, idx) => {
        if (!entry || !entry.sourceOp || !entry.source || entry.isLabel) return;
        if (!byTool.has(entry.sourceOp)) byTool.set(entry.sourceOp, []);
        byTool.get(entry.sourceOp).push({ entry, idx });
      });
      const existing = result.entries
        .map((entry) => (entry && Number.isFinite(entry.controlGroup) ? entry.controlGroup : null))
        .filter((v) => Number.isFinite(v));
      let nextGroup = existing.length ? (Math.max(...existing) + 1) : 1;
      byTool.forEach((items) => {
        if (!Array.isArray(items) || items.length < 2) return;
        const horiz = items.find((item) => normalizeId(item.entry.source).toLowerCase() === 'fliphoriz');
        const vert = items.find((item) => normalizeId(item.entry.source).toLowerCase() === 'flipvert');
        if (!horiz || !vert) return;
        const hGroup = Number.isFinite(horiz.entry.controlGroup) ? horiz.entry.controlGroup : null;
        const vGroup = Number.isFinite(vert.entry.controlGroup) ? vert.entry.controlGroup : null;
        const group = hGroup != null ? hGroup : (vGroup != null ? vGroup : nextGroup++);
        if (!Number.isFinite(horiz.entry.controlGroup)) {
          horiz.entry.controlGroup = group;
          if (typeof horiz.entry.raw === 'string') {
            horiz.entry.raw = ensureInstanceInputControlGroup(horiz.entry.raw, group, eol || '\n');
          }
        }
        if (!Number.isFinite(vert.entry.controlGroup)) {
          vert.entry.controlGroup = group;
          if (typeof vert.entry.raw === 'string') {
            vert.entry.raw = ensureInstanceInputControlGroup(vert.entry.raw, group, eol || '\n');
          }
        }
      });
    } catch (_) {}
  }

  function cloneParseResultForCsv(result) {
    if (!result) return null;
    const clone = {
      ...result,
      entries: Array.isArray(result.entries)
        ? result.entries.map((entry) => {
          if (!entry) return entry;
          return {
            ...entry,
            controlMeta: entry.controlMeta ? { ...entry.controlMeta } : {},
            controlMetaOriginal: entry.controlMetaOriginal ? { ...entry.controlMetaOriginal } : {},
          };
        })
        : [],
      order: Array.isArray(result.order) ? [...result.order] : [],
      pageOrder: Array.isArray(result.pageOrder) ? [...result.pageOrder] : [],
      selected: result.selected ? new Set(Array.from(result.selected)) : new Set(),
      collapsed: result.collapsed ? new Set(Array.from(result.collapsed)) : new Set(),
      collapsedLabels: result.collapsedLabels ? new Set(Array.from(result.collapsedLabels)) : new Set(),
      collapsedCG: result.collapsedCG ? new Set(Array.from(result.collapsedCG)) : new Set(),
      blendToggles: result.blendToggles ? new Set(Array.from(result.blendToggles)) : new Set(),
      buttonExactInsert: result.buttonExactInsert ? new Set(Array.from(result.buttonExactInsert)) : new Set(),
      insertUrlMap: result.insertUrlMap ? new Map(result.insertUrlMap) : new Map(),
      buttonOverrides: result.buttonOverrides ? new Map(result.buttonOverrides) : new Map(),
    };
    if (result.pageIcons instanceof Map) {
      clone.pageIcons = new Map(result.pageIcons);
    } else if (result.pageIcons && typeof result.pageIcons === 'object') {
      clone.pageIcons = { ...result.pageIcons };
    }
    return clone;
  }

  function applyDefaultToEntryRaw(entry, next, newline) {
    if (!entry || typeof entry.raw !== 'string') return;
    const open = entry.raw.indexOf('{');
    const close = entry.raw.lastIndexOf('}');
    if (open < 0 || close <= open) return;
    const indent = (getLineIndent(entry.raw, open) || '') + '\t';
    let body = entry.raw.slice(open + 1, close);
    if (next == null) body = removeInstanceInputProp(body, 'Default');
    else body = setInstanceInputProp(body, 'Default', next, indent, newline || '\n');
    entry.raw = entry.raw.slice(0, open + 1) + body + entry.raw.slice(close);
  }

  async function openCsvColumnPicker({ title, headers, current, allowNone = true } = {}) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal csv-modal';
      const modal = document.createElement('form');
      modal.className = 'add-control-form';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      const header = document.createElement('header');
      const headingWrap = document.createElement('div');
      const eyebrow = document.createElement('p');
      eyebrow.className = 'eyebrow';
      eyebrow.textContent = 'CSV';
      const heading = document.createElement('h3');
      heading.textContent = title || 'Select CSV column';
      headingWrap.appendChild(eyebrow);
      headingWrap.appendChild(heading);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = '×';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headingWrap);
      header.appendChild(closeBtn);
      const body = document.createElement('div');
      body.className = 'form-body';
      const label = document.createElement('label');
      label.textContent = 'Column';
      const select = document.createElement('select');
      if (allowNone) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = 'Unlink (no column)';
        select.appendChild(opt);
      }
      (headers || []).forEach((h) => {
        const opt = document.createElement('option');
        opt.value = h;
        opt.textContent = h;
        select.appendChild(opt);
      });
      if (current) select.value = current;
      label.appendChild(select);
      body.appendChild(label);
      const actions = document.createElement('footer');
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.className = 'primary';
      okBtn.textContent = 'Link';
      const close = (value) => {
        try { overlay.remove(); } catch (_) {}
        resolve(value);
      };
      cancelBtn.addEventListener('click', () => close(null));
      closeBtn.addEventListener('click', () => close(null));
      modal.addEventListener('submit', (ev) => {
        ev.preventDefault();
        close(select.value);
      });
      overlay.addEventListener('click', (ev) => {
        if (ev.target === overlay) close(null);
      });
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);
      modal.appendChild(header);
      modal.appendChild(body);
      modal.appendChild(actions);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
      select.focus();
    });
  }

  async function openCsvNameColumnPicker(headers, current) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal csv-modal';
      const modal = document.createElement('form');
      modal.className = 'add-control-form';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      const header = document.createElement('header');
      const headingWrap = document.createElement('div');
      const eyebrow = document.createElement('p');
      eyebrow.className = 'eyebrow';
      eyebrow.textContent = 'CSV';
      const heading = document.createElement('h3');
      heading.textContent = 'Choose name column';
      headingWrap.appendChild(eyebrow);
      headingWrap.appendChild(heading);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = '×';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headingWrap);
      header.appendChild(closeBtn);
      const body = document.createElement('div');
      body.className = 'form-body';
      const label = document.createElement('label');
      label.textContent = 'Name column';
      const select = document.createElement('select');
      const rowOpt = document.createElement('option');
      rowOpt.value = '__row__';
      rowOpt.textContent = 'Row number';
      select.appendChild(rowOpt);
      (headers || []).forEach((h) => {
        const opt = document.createElement('option');
        opt.value = h;
        opt.textContent = h;
        select.appendChild(opt);
      });
      if (current) select.value = current;
      label.appendChild(select);
      body.appendChild(label);
      const actions = document.createElement('footer');
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.className = 'primary';
      okBtn.textContent = 'Continue';
      const close = (value) => {
        try { overlay.remove(); } catch (_) {}
        resolve(value);
      };
      cancelBtn.addEventListener('click', () => close(null));
      closeBtn.addEventListener('click', () => close(null));
      modal.addEventListener('submit', (ev) => {
        ev.preventDefault();
        close(select.value);
      });
      overlay.addEventListener('click', (ev) => {
        if (ev.target === overlay) close(null);
      });
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);
      modal.appendChild(header);
      modal.appendChild(body);
      modal.appendChild(actions);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
      select.focus();
    });
  }

  function getCsvLinkIcon() {
    return '+';
  }

  function setEntryDefaultValue(index, value) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    let next = normalizeMetaValue(value);
    if (next != null && isChoiceControl(entry) && !isTextControl(entry)) {
      const comboIndex = parseComboIndexValue(entry, next);
      if (comboIndex == null) {
        error('Default must be a numeric option index.');
        return;
      }
      next = String(comboIndex);
    }
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
    if (entry.isLabel) {
      renderActiveList();
    }
    markContentDirty();
  }

  function getEntryInputContext(entry) {
    try {
      if (!entry || !entry.sourceOp || !entry.source || !state.originalText || !state.parseResult) return null;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) return null;
      const toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) return null;
      const inputsBlock = findInputsInTool(state.originalText, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return null;
      const inputBlock = findInputBlockInInputs(state.originalText, inputsBlock.open, inputsBlock.close, entry.source);
      if (!inputBlock) return null;
      const body = state.originalText.slice(inputBlock.open + 1, inputBlock.close);
      return { bounds, toolBlock, inputsBlock, inputBlock, body };
    } catch (_) {
      return null;
    }
  }

  function getEntryExpression(entry) {
    const ctx = getEntryInputContext(entry);
    if (!ctx) return '';
    return extractControlStringProp(ctx.body, 'Expression') || '';
  }

  function setEntryCurrentValue(index, value, options = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry || !state.originalText) return;
    const ctx = getEntryInputContext(entry);
    if (!ctx) return;
    let trimmed = String(value || '').trim();
    if (trimmed && isChoiceControl(entry) && !isTextControl(entry)) {
      const comboIndex = parseComboIndexValue(entry, trimmed);
      if (comboIndex == null) {
        if (!options.silent) error('Value must be a numeric option index.');
        return;
      }
      trimmed = String(comboIndex);
    }
    let body = ctx.body;
    const indent = (getLineIndent(state.originalText, ctx.inputBlock.open) || '') + '\t';
    const eol = state.newline || '\n';
    if (trimmed) {
      body = removeControlProp(body, 'Expression');
      body = removeControlProp(body, 'SourceOp');
      body = removeControlProp(body, 'Source');
      body = upsertControlProp(body, 'Value', trimmed, indent, eol);
    } else {
      body = removeControlProp(body, 'Value');
    }
    const updated = state.originalText.slice(0, ctx.inputBlock.open + 1) + body + state.originalText.slice(ctx.inputBlock.close);
    if (updated === state.originalText) return false;
    if (!options.skipHistory) pushHistory('edit current value');
    state.originalText = updated;
    if (!options.skipRender) renderActiveList();
    if (!options.skipMarkDirty) markContentDirty();
    return true;
  }

  function setEntryExpression(index, expression) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
    const entry = state.parseResult.entries[index];
    if (!entry || !state.originalText) return false;
    const ctx = getEntryInputContext(entry);
    if (!ctx) return false;
    const nextExpr = String(expression || '').trim();
    const indent = (getLineIndent(state.originalText, ctx.inputBlock.open) || '') + '\t';
    const eol = state.newline || '\n';
    let body = ctx.body;
    if (nextExpr) {
      body = removeControlProp(body, 'Value');
      body = removeControlProp(body, 'SourceOp');
      body = removeControlProp(body, 'Source');
      body = upsertControlProp(body, 'Expression', `"${escapeSettingString(nextExpr)}"`, indent, eol);
    } else {
      body = removeControlProp(body, 'Expression');
    }
    const updated = state.originalText.slice(0, ctx.inputBlock.open + 1) + body + state.originalText.slice(ctx.inputBlock.close);
    if (updated === state.originalText) return false;
    pushHistory('edit expression');
    state.originalText = updated;
    renderActiveList();
    markContentDirty();
    return true;
  }

  function setEntryLabelStyle(index, style) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry || !entry.isLabel) return;
    const current = normalizeLabelStyle(entry.labelStyle);
    const next = normalizeLabelStyle({ ...current, ...(style || {}) });
    if (labelStyleEquals(current, next)) return;
    const original = normalizeLabelStyle(entry.labelStyleOriginal);
    pushHistory('label style');
    entry.labelStyle = next;
    entry.labelStyleDirty = !labelStyleEquals(original, next);
    entry.labelStyleEdited = true;
    renderActiveList();
    if (activeDetailEntryIndex === index) renderDetailDrawer(index);
    markContentDirty();
  }

  function normalizeLabelImagePathInput(value) {
    let raw = String(value || '').trim();
    if (!raw) return null;
    if ((raw.startsWith('"') && raw.endsWith('"')) || (raw.startsWith("'") && raw.endsWith("'"))) {
      raw = raw.slice(1, -1).trim();
    }
    return raw || null;
  }

  function extractImageSrcFromLabelMarkup(value) {
    const raw = normalizeLabelImagePathInput(value);
    if (!raw) return null;
    if (isSupportedLabelImageDataUri(raw)) return raw;
    let match = String(raw).match(/<\s*img\b[^>]*\bsrc\s*=\s*(['"])(.*?)\1/i);
    if (!match) {
      match = String(raw).match(/<\s*img\b[^>]*\bsrc\s*=\s*([^'"\s>]+)/i);
    }
    if (!match) {
      // Resolve header-image labels are often exported as a truncated tag:
      // <center><img src='data:image/...base64,AAAA
      // with no trailing quote or angle bracket. Accept that form.
      match = String(raw).match(/<\s*img\b[^>]*\bsrc\s*=\s*['"]([^'"]+)$/i);
    }
    if (!match) return null;
    let candidate = normalizeLabelImagePathInput(match[2] != null ? match[2] : match[1]);
    if (!candidate) return null;
    // Imported Resolve header payloads can carry trailing tag residue.
    candidate = String(candidate).trim().replace(/[>'"]+$/g, '');
    return isSupportedLabelImageDataUri(candidate) ? candidate : null;
  }

  function buildHeaderCarrierImageMarkup(value) {
    const src = extractImageSrcFromLabelMarkup(value);
    if (!src) return '';
    const compact = String(src).replace(/\s+/g, '');
    // Resolve header-image labels are sensitive to this legacy formatting:
    // padded leading spaces, no closing quote, and trailing '> '.
    return `        <center><img src='${compact}> `;
  }

  function isSupportedLabelImageDataUri(value) {
    const raw = normalizeLabelImagePathInput(value);
    if (!raw) return false;
    const lower = raw.toLowerCase();
    return lower.startsWith('data:image/png;base64,')
      || lower.startsWith('data:image/jpeg;base64,')
      || lower.startsWith('data:image/jpg;base64,');
  }

  function isSupportedLabelImageFilePath(value) {
    const raw = normalizeLabelImagePathInput(value);
    if (!raw) return false;
    const lower = raw.toLowerCase();
    if (lower.startsWith('data:')) return false;
    if (lower.startsWith('javascript:')) return false;
    const stem = raw.split(/[?#]/, 1)[0] || raw;
    return /\.(png|jpe?g)$/i.test(stem);
  }

  function isSupportedLabelImagePath(value) {
    return isSupportedLabelImageDataUri(value) || isSupportedLabelImageFilePath(value);
  }

  function findHeaderLabelCarrierTargets(result = state.parseResult) {
    const out = [];
    if (!result || !Array.isArray(result.entries) || !result.entries.length) return out;
    for (let i = 0; i < result.entries.length; i++) {
      const entry = result.entries[i];
      if (!entry || entry.source !== FMR_HEADER_LABEL_CONTROL) continue;
      out.push({ index: i, entry });
    }
    return out;
  }

  function findHeaderLabelCarrierTarget(result = state.parseResult, preferredSourceOp = '') {
    const all = findHeaderLabelCarrierTargets(result);
    if (!all.length) return null;
    const preferred = String(preferredSourceOp || '').trim();
    if (preferred) {
      const match = all.find((item) => String(item.entry?.sourceOp || '') === preferred);
      if (match) return match;
    }
    const macroName = result?.macroName || result?.macroNameOriginal || '';
    if (!macroName) return all[0];
    const exact = all.find((item) => item.entry?.sourceOp === macroName);
    return exact || all[0];
  }

  function findHeaderCarrierControlBlockInText(text = state.originalText, result = state.parseResult, preferredSourceOp = '') {
    try {
      if (!text || !result) return null;
      const groupBlock = findGroupUserControlBlockById(text, result, FMR_HEADER_LABEL_CONTROL);
      if (groupBlock) {
        return {
          sourceOp: String(result.macroName || result.macroNameOriginal || '').trim(),
          open: groupBlock.open,
          close: groupBlock.close,
        };
      }
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return null;
      const candidates = [];
      const seen = new Set();
      const addCandidate = (name) => {
        const toolName = String(name || '').trim();
        if (!toolName || seen.has(toolName)) return;
        seen.add(toolName);
        candidates.push(toolName);
      };
      addCandidate(preferredSourceOp);
      addCandidate(pickPrimaryToolNameForHeader(text, result));
      const toolsBlock = findOrderedBlock(text, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
      if (toolsBlock) {
        const toolEntries = parseOrderedBlockEntries(text, toolsBlock.open, toolsBlock.close);
        toolEntries.forEach((item) => addCandidate(item?.name));
      }
      for (const toolName of candidates) {
        const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
        if (!toolBlock) continue;
        const uc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
        if (!uc) continue;
        const block = findControlBlockInUc(text, uc.open, uc.close, FMR_HEADER_LABEL_CONTROL);
        if (block) return { sourceOp: toolName, open: block.open, close: block.close };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function parseHeaderCarrierStyleFromTextBlock(text, block) {
    try {
      if (!text || !block) return normalizeLabelStyle(null);
      const body = text.slice(block.open + 1, block.close);
      const raw = normalizeMetaValue(
        extractControlPropValue(body, 'LINKS_Name')
        || extractControlPropValue(body, 'INP_Default')
      ) || '';
      const parsed = parseLabelMarkup(raw);
      return normalizeLabelStyle(parsed?.style || null);
    } catch (_) {
      return normalizeLabelStyle(null);
    }
  }

  function ensureHeaderLabelCarrierEntry(result = state.parseResult, preferredSourceOp = '') {
    try {
      if (!result || !Array.isArray(result.entries)) return null;
      const existing = findHeaderLabelCarrierTarget(result, preferredSourceOp);
      if (existing && existing.entry) return existing;
      const block = findHeaderCarrierControlBlockInText(state.originalText, result, preferredSourceOp);
      if (!block) return null;
      const sourceOp = String(
        block.sourceOp
        || preferredSourceOp
        || result.macroName
        || result.macroNameOriginal
        || 'Macro'
      ).trim();
      const key = makeUniqueKey(`${sourceOp}_${FMR_HEADER_LABEL_CONTROL}`);
      const style = parseHeaderCarrierStyleFromTextBlock(state.originalText, block);
      const entry = {
        key,
        name: null,
        page: null,
        sourceOp,
        source: FMR_HEADER_LABEL_CONTROL,
        displayName: '',
        displayNameOriginal: '',
        raw: buildInstanceInputRaw(key, sourceOp, FMR_HEADER_LABEL_CONTROL, '', 'Controls', null),
        controlGroup: null,
        onChange: '',
        buttonExecute: '',
        isLabel: true,
        labelCount: 0,
        labelStyle: style,
        labelStyleOriginal: { ...style },
        labelStyleDirty: false,
        labelStyleEdited: false,
        headerCarrierSynthetic: true,
      };
      markHeaderCarrierEntryInternal(entry);
      result.entries.push(entry);
      return { index: result.entries.length - 1, entry };
    } catch (_) {
      return null;
    }
  }

  function removeHeaderCarrierInstanceInputs(text, result = state.parseResult) {
    try {
      if (!text || !result || !result.inputs) return text;
      let open = result.inputs.openIndex;
      if (!Number.isFinite(open) || open < 0 || open >= text.length || text[open] !== '{') return text;
      let updated = text;
      while (true) {
        const close = findMatchingBrace(updated, open);
        if (!Number.isFinite(close) || close <= open) break;
        const entries = parseOrderedBlockEntries(updated, open, close);
        let targetName = null;
        for (const item of entries) {
          const body = updated.slice(item.blockOpen + 1, item.blockClose);
          const src = normalizeMetaValue(extractControlPropValue(body, 'Source'));
          if (String(src || '') === FMR_HEADER_LABEL_CONTROL) {
            targetName = item.name;
            break;
          }
        }
        if (!targetName) break;
        const next = removeControlBlockFromUcById(updated, { open, close }, targetName);
        if (next === updated) break;
        updated = next;
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function pickPrimaryToolNameForHeader(text = state.originalText, result = state.parseResult) {
    try {
      if (!text || !result) return '';
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return '';
      const toolsBlock = findOrderedBlock(text, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
      if (!toolsBlock) return '';
      const entries = parseOrderedBlockEntries(text, toolsBlock.open, toolsBlock.close);
      return entries && entries.length ? String(entries[0].name || '') : '';
    } catch (_) {
      return '';
    }
  }

  function removeControlBlockFromUcById(text, ucRange, controlId) {
    try {
      if (!text || !ucRange) return text;
      const ucOpen = ucRange.open != null ? ucRange.open : ucRange.openIndex;
      const ucClose = ucRange.close != null ? ucRange.close : ucRange.closeIndex;
      if (ucOpen == null || ucClose == null) return text;
      const entries = parseOrderedBlockEntries(text, ucOpen, ucClose);
      const target = entries.find((item) => String(item.name || '') === String(controlId || ''));
      if (!target) return text;
      let delStart = target.nameStart;
      while (delStart > ucOpen + 1 && /[ \t]/.test(text[delStart - 1])) delStart--;
      if (delStart > ucOpen + 1 && text[delStart - 1] === '\r') delStart--;
      if (delStart > ucOpen + 1 && text[delStart - 1] === '\n') delStart--;
      let delEnd = target.blockClose + 1;
      while (delEnd < text.length && /[ \t]/.test(text[delEnd])) delEnd++;
      if (text[delEnd] === ',') delEnd++;
      while (delEnd < text.length && /[ \t]/.test(text[delEnd])) delEnd++;
      if (text[delEnd] === '\r') delEnd++;
      if (text[delEnd] === '\n') delEnd++;
      return text.slice(0, delStart) + text.slice(delEnd);
    } catch (_) {
      return text;
    }
  }

  function removeHeaderCarrierFromGroup(text, result) {
    try {
      return removeGroupUserControlById(text, result, FMR_HEADER_LABEL_CONTROL);
    } catch (_) {
      return text;
    }
  }

  function removeHeaderCarrierFromOtherTools(text, result, keepToolName = '') {
    try {
      let out = text;
      const keep = String(keepToolName || '').trim();
      while (true) {
        const bounds = locateMacroGroupBounds(out, result);
        if (!bounds) break;
        const toolsBlock = findOrderedBlock(out, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
        if (!toolsBlock) break;
        const tools = parseOrderedBlockEntries(out, toolsBlock.open, toolsBlock.close);
        let removed = false;
        for (const item of tools) {
          const toolName = String(item?.name || '').trim();
          if (!toolName) continue;
          if (keep && toolName === keep) continue;
          const tb = findToolBlockInGroup(out, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
          if (!tb) continue;
          const uc = findUserControlsInTool(out, tb.open, tb.close);
          if (!uc) continue;
          const next = removeControlBlockFromUcById(out, uc, FMR_HEADER_LABEL_CONTROL);
          if (next !== out) {
            out = next;
            removed = true;
            break;
          }
        }
        if (!removed) break;
      }
      return out;
    } catch (_) {
      return text;
    }
  }

  function normalizeHeaderCarrierControlBody(body, indent, newline) {
    try {
      let next = body;
      next = upsertControlProp(next, 'LINKS_Name', '""', indent, newline);
      next = upsertControlProp(next, 'INP_Integer', 'false', indent, newline);
      next = upsertControlProp(next, 'INPID_InputControl', '"LabelControl"', indent, newline);
      next = upsertControlProp(next, 'LBLC_MultiLine', 'true', indent, newline);
      next = upsertControlProp(next, 'INP_External', 'false', indent, newline);
      next = upsertControlProp(next, 'LINKID_DataType', '"Number"', indent, newline);
      next = upsertControlProp(next, 'IC_NoReset', 'true', indent, newline);
      next = upsertControlProp(next, 'INP_Passive', 'true', indent, newline);
      next = upsertControlProp(next, 'IC_NoLabel', 'true', indent, newline);
      next = upsertControlProp(next, 'IC_ControlPage', '-1', indent, newline);
      next = removeControlProp(next, 'INP_Default');
      next = removeControlProp(next, 'INP_SplineType');
      next = removeControlProp(next, 'LBLC_NumInputs');
      next = removeControlProp(next, 'LBLC_DropDownButton');
      next = removeControlProp(next, 'IC_Visible');
      next = removeControlProp(next, 'ICS_ControlPage');
      return next;
    } catch (_) {
      return body;
    }
  }

  function markHeaderCarrierEntryInternal(entry) {
    try {
      if (!entry) return;
      entry.locked = true;
      entry.isLabel = true;
      if (!Number.isFinite(entry.labelCount)) entry.labelCount = 0;
      entry.sortIndex = -1000000;
      entry.displayName = '';
      entry.displayNameOriginal = '';
      entry.displayNameDirty = true;
    } catch (_) {}
  }

  function hideHeaderCarrierEntriesFromPublishedOrder(result = state.parseResult) {
    try {
      if (!result || !Array.isArray(result.entries)) return;
      const targets = findHeaderLabelCarrierTargets(result);
      if (!targets.length) return;
      const hidden = new Set(targets.map((item) => item.index));
      if (Array.isArray(result.order)) {
        result.order = result.order.filter((index) => !hidden.has(index));
      }
      if (Array.isArray(result.originalOrder)) {
        result.originalOrder = result.originalOrder.filter((index) => !hidden.has(index));
      }
      if (result.selected && typeof result.selected.delete === 'function') {
        hidden.forEach((index) => result.selected.delete(index));
      }
      if (typeof activeDetailEntryIndex === 'number' && hidden.has(activeDetailEntryIndex)) {
        activeDetailEntryIndex = null;
      }
    } catch (_) {}
  }

  async function ensureHeaderLabelCarrier() {
    try {
      if (!state.parseResult || !state.originalText) return null;
      pushHistory('add header image carrier');
      const newline = state.newline || detectNewline(state.originalText) || '\n';
      const macroName = state.parseResult.macroName || state.parseResult.macroNameOriginal || 'Macro';
      const preferredSourceOp = macroName;

      let workingText = state.originalText;
      workingText = removeHeaderCarrierInstanceInputs(workingText, state.parseResult);
      while (true) {
        const next = removeGroupUserControlById(workingText, state.parseResult, FMR_HEADER_LABEL_CONTROL);
        if (next === workingText) break;
        workingText = next;
      }
      const lines = buildControlDefinitionLines('label', {
        name: '',
        page: 'Controls',
        labelCount: 0,
        labelDefault: 'closed',
      });
      workingText = upsertGroupUserControl(workingText, state.parseResult, FMR_HEADER_LABEL_CONTROL, {
        createLines: lines,
        updateBody: (body, indent, nl) => normalizeHeaderCarrierControlBody(body, indent, nl),
      }, newline);
      workingText = removeHeaderCarrierFromOtherTools(workingText, state.parseResult, '');
      const meta = {
        kind: 'label',
        labelCount: 0,
        defaultValue: '0',
        inputControl: 'LabelControl',
        page: 'Controls',
        locked: true,
      };
      rememberPendingControlMeta(macroName, FMR_HEADER_LABEL_CONTROL, meta);

      const persisted = rebuildContentWithNewOrder(workingText, state.parseResult, newline);
      if (persisted !== state.originalText) {
        state.originalText = persisted;
        await reloadMacroFromCurrentText({ skipClear: true });
      }
      const refreshed = ensureHeaderLabelCarrierEntry(state.parseResult, preferredSourceOp);
      if (refreshed && refreshed.entry) {
        markHeaderCarrierEntryInternal(refreshed.entry);
      }
      hideHeaderCarrierEntriesFromPublishedOrder(state.parseResult);
      return refreshed;
    } catch (_) {
      return null;
    }
  }

  function getHeaderLabelTarget() {
    const preferred = state.parseResult?.macroName
      || state.parseResult?.macroNameOriginal
      || '';
    const target = ensureHeaderLabelCarrierEntry(state.parseResult, preferred);
    if (!target || !preferred) return target;
    if (String(target.entry?.sourceOp || '') !== String(preferred)) return null;
    return target;
  }

  async function applyHeaderImageValue(rawInput) {
    let target = getHeaderLabelTarget();
    const raw = normalizeLabelImagePathInput(rawInput);
    if (!target && raw) {
      target = await ensureHeaderLabelCarrier();
    }
    if (!target && !raw) {
      return { ok: true, cleared: false, noop: true };
    }
    if (!target) {
      return { ok: false, error: 'Unable to create header image carrier label.' };
    }
    if (!raw) {
      setEntryLabelStyle(target.index, { imagePath: null });
      return { ok: true, cleared: true };
    }
    if (!isSupportedLabelImagePath(raw)) {
      return { ok: false, error: 'Use PNG/JPG or a PNG/JPG base64 data URI.' };
    }
    let finalImagePath = raw;
    let embedded = false;
    if (!isSupportedLabelImageDataUri(raw)) {
      const api = getNativeApi();
      if (!api || typeof api.readImageDataUri !== 'function') {
        return { ok: false, error: 'Native image embedding unavailable.' };
      }
      let encoded = null;
      try {
        encoded = await api.readImageDataUri({ filePath: raw });
      } catch (err) {
        const msg = String((err && err.message) || err || '');
        if (/No handler registered/i.test(msg)) {
          return { ok: false, error: 'App restart/update required (image embed handler missing).' };
        }
        return { ok: false, error: 'Unable to embed image.' };
      }
      if (!encoded || encoded.ok !== true || !encoded.dataUri) {
        return { ok: false, error: String((encoded && encoded.error) || 'Unable to embed image.') };
      }
      if (!isSupportedLabelImageDataUri(encoded.dataUri)) {
        return { ok: false, error: 'Unsupported embedded image format.' };
      }
      finalImagePath = encoded.dataUri;
      embedded = true;
    }
    setEntryLabelStyle(target.index, { imagePath: finalImagePath });
    return { ok: true, embedded };
  }

  function openHeaderImagePickerModal({ initialValue, statusText } = {}) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal';
      overlay.setAttribute('role', 'dialog');
      overlay.setAttribute('aria-modal', 'true');

      const form = document.createElement('form');
      form.className = 'add-control-form';

      const header = document.createElement('header');
      const headerWrap = document.createElement('div');
      const titleEl = document.createElement('h3');
      titleEl.textContent = 'Header Image';
      headerWrap.appendChild(titleEl);

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = 'x';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headerWrap);
      header.appendChild(closeBtn);

      const body = document.createElement('div');
      body.className = 'form-body';

      const field = document.createElement('label');
      field.textContent = 'Image path or data URI';
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'header-image-input';
      input.placeholder = 'Path (.png/.jpg/.jpeg) or data:image/...;base64,...';
      input.spellcheck = false;
      input.value = initialValue || '';
      field.appendChild(input);
      body.appendChild(field);

      const status = document.createElement('p');
      status.className = 'header-image-status';
      status.style.margin = '0';
      status.style.whiteSpace = 'pre-line';
      status.textContent = statusText || 'Target: internal header label';
      body.appendChild(status);

      const actions = document.createElement('div');
      actions.className = 'modal-actions';

      const browseBtn = document.createElement('button');
      browseBtn.type = 'button';
      browseBtn.textContent = 'Browse...';

      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.textContent = 'Clear';

      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';

      const applyBtn = document.createElement('button');
      applyBtn.type = 'submit';
      applyBtn.textContent = 'Apply';

      actions.appendChild(browseBtn);
      actions.appendChild(clearBtn);
      actions.appendChild(cancelBtn);
      actions.appendChild(applyBtn);

      form.appendChild(header);
      form.appendChild(body);
      form.appendChild(actions);
      overlay.appendChild(form);
      document.body.appendChild(overlay);

      let done = false;
      const finish = (result) => {
        if (done) return;
        done = true;
        try { overlay.remove(); } catch (_) {}
        document.removeEventListener('keydown', onKeyDown);
        resolve(result);
      };

      const onCancel = () => finish(null);
      const onSubmit = (ev) => {
        ev.preventDefault();
        finish({ action: 'apply', value: input.value });
      };
      const onOverlayClick = (ev) => {
        if (ev.target === overlay) onCancel();
      };
      const onKeyDown = (ev) => {
        if (ev.key === 'Escape') {
          ev.preventDefault();
          onCancel();
        }
      };

      browseBtn.addEventListener('click', async () => {
        try {
          const api = getNativeApi();
          if (!api || typeof api.pickImageFile !== 'function') {
            status.textContent = 'Native image picker unavailable. Paste path or data URI manually.';
            return;
          }
          const currentInput = normalizeLabelImagePathInput(input.value);
          const res = await api.pickImageFile({
            defaultPath: currentInput && !isSupportedLabelImageDataUri(currentInput) ? currentInput : '',
          });
          if (!res || res.canceled || !res.filePath) return;
          input.value = String(res.filePath || '');
          status.textContent = 'Selected file path. Click Apply to embed/set.';
        } catch (_) {
          status.textContent = 'Unable to open image picker.';
        }
      });
      clearBtn.addEventListener('click', () => finish({ action: 'clear', value: '' }));
      cancelBtn.addEventListener('click', onCancel);
      closeBtn.addEventListener('click', onCancel);
      form.addEventListener('submit', onSubmit);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeyDown);

      input.focus();
      input.select();
    });
  }

  async function openHeaderImageDialog() {
    if (!state.parseResult || !state.originalText) {
      error('Load a macro before setting a header image.');
      return;
    }
    const target = getHeaderLabelTarget();
    const style = normalizeLabelStyle(target?.entry?.labelStyle);
    const currentRaw = extractImageSrcFromLabelMarkup(style.imagePath)
      || normalizeLabelImagePathInput(style.imagePath)
      || '';
    const currentIsEmbedded = isSupportedLabelImageDataUri(currentRaw);
    const initialValue = currentIsEmbedded ? '' : currentRaw;
    const statusText = target
      ? (currentIsEmbedded
        ? 'Embedded PNG/JPG currently set on internal header label.'
        : 'Target: internal header label')
      : 'Will create internal header label on first set';
    const modalResult = await openHeaderImagePickerModal({ initialValue, statusText });
    if (!modalResult) return;
    const nextValue = modalResult.action === 'clear' ? '' : modalResult.value;
    const applied = await applyHeaderImageValue(nextValue);
    if (!applied.ok) {
      error(applied.error || 'Unable to set header image.');
      return;
    }
    if (applied.cleared) {
      info('Header image cleared.');
      return;
    }
    info(applied.embedded ? 'Header image applied (embedded PNG/JPG).' : 'Header image applied.');
  }

  async function reloadMacroFromCurrentText(options = {}) {
    try {
      const sourceName = state.originalFileName || 'Imported.setting';
      const keepHistory = options.preserveHistory !== false;
      const prevHistory = keepHistory && state.parseResult ? state.parseResult.history : null;
      const prevFuture = keepHistory && state.parseResult ? state.parseResult.future : null;
      await loadMacroFromText(sourceName, state.originalText || '', {
        allowAutoUtility: false,
        preserveFileInfo: true,
        preserveFilePath: true,
        skipClear: !!options.skipClear,
        silentAuto: true,
        createDoc: false,
      });
      if (state.parseResult) {
        if (keepHistory) {
          if (prevHistory) state.parseResult.history = prevHistory;
          if (prevFuture) state.parseResult.future = prevFuture;
        } else {
          state.parseResult.history = [];
          state.parseResult.future = [];
        }
        updateUndoRedoState();
      }
    } catch (_) {
      return Promise.reject(_);
    }
  }

  function buildControlDefinitionLines(type, config = {}) {
    const name = config.name ? escapeQuotes(config.name) : 'Custom';
    const page = config.page ? `ICS_ControlPage = "${escapeQuotes(config.page)}",` : null;
    if (type === 'slider' || type === 'screw' || type === 'combo') {
      const control = type === 'slider' ? 'SliderControl' : type === 'screw' ? 'ScrewControl' : 'ComboControl';
      const lines = [
        `LINKS_Name = "${name}",`,
        'LINKID_DataType = "Number",',
        'INP_Integer = false,',
        'INP_Default = 0,',
        'INP_SplineType = "Default",',
        `INPID_InputControl = "${control}",`,
      ];
      if (type !== 'combo') {
        lines.splice(4, 0,
          'INP_MinScale = 0,',
          'INP_MaxScale = 1,',
          'INP_MinAllowed = -1000000,',
          'INP_MaxAllowed = 1000000,'
        );
      }
      if (type === 'combo') {
        const rawOptions = Array.isArray(config.comboOptions) ? config.comboOptions : [];
        const options = rawOptions
          .map((opt) => String(opt || '').trim())
          .filter((opt) => opt.length > 0);
        const safeOptions = options.length ? options : ['Option 1', 'Option 2', 'Option 3'];
        safeOptions.forEach((opt) => {
          lines.push(`{ CCS_AddString = "${escapeQuotes(opt)}" },`);
        });
      }
      if (page) lines.push(page);
      return lines;
    }
      if (type === 'button') {
        const lines = [
          `LINKS_Name = "${name}",`,
          'LINKID_DataType = "Number",',
          'INP_Default = 0,',
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
    const supportedTypes = new Set(['label', 'separator', 'button', 'slider', 'combo', 'screw']);
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
      comboOptions: Array.isArray(config?.comboOptions) ? config.comboOptions : [],
    });
    const pendingMeta = {
      kind: type === 'label' ? 'label' : type === 'button' ? 'button' : null,
      labelCount: type === 'label' && Number.isFinite(config?.labelCount) ? Number(config.labelCount) : null,
      selectionInsertStartPos: type === 'label' && Number.isFinite(config?.labelSelectionSpan?.startPos)
        ? Number(config.labelSelectionSpan.startPos)
        : null,
      defaultValue: type === 'label' ? (labelDefault === 'closed' ? '0' : '1') : null,
      inputControl:
        type === 'slider' ? 'SliderControl'
        : type === 'combo' ? 'ComboControl'
        : type === 'screw' ? 'ScrewControl'
        : type === 'button' ? 'ButtonControl'
        : type === 'separator' ? 'SeparatorControl'
        : 'LabelControl',
      choiceOptions: type === 'combo' && Array.isArray(config?.comboOptions)
        ? config.comboOptions.map((opt) => String(opt || '').trim()).filter((opt) => opt.length > 0)
        : [],
    };
    workingText = insertUserControlBlock(workingText, ensured.ucBlock, controlId, lines, newline);
    rememberPendingControlMeta(nodeName, controlId, pendingMeta);
    autoPublishCreatedControl(nodeName, controlId, trimmedName, normalizedPage, pendingMeta);
    pendingOpenNodes.add(nodeName);
    const persisted = rebuildContentWithNewOrder(workingText, state.parseResult, newline, { safeEditExport: true });
    state.originalText = persisted;
    await reloadMacroFromCurrentText({ skipClear: true });
    runValidation('add-control');
    const typeLabel = type === 'button'
      ? 'button'
      : type === 'slider'
        ? 'slider'
        : type === 'combo'
          ? 'combo control'
          : type === 'screw'
            ? 'screw control'
            : type === 'separator'
              ? 'separator control'
              : 'label';
    info(`Added ${typeLabel} "${trimmedName}" to ${nodeName}.`);
    return { nodeName, controlId, displayName: trimmedName };
  }

  async function addControlToGroup(config) {
    if (!state.originalText || !state.parseResult) throw new Error('Load a macro before adding controls.');
    pushHistory('add control');
    const newline = state.newline || detectNewline(state.originalText);
    const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
    if (!bounds) throw new Error('Unable to locate the macro group block.');
    const trimmedName = (config?.name || '').trim();
    if (!trimmedName) throw new Error('Control name is required.');
    const normalizedPage = (config?.page && String(config.page).trim()) ? String(config.page).trim() : 'Controls';
    const typeRaw = (config?.type || 'label').toLowerCase();
    const supportedTypes = new Set(['label', 'separator', 'button', 'slider', 'combo', 'screw']);
    const type = supportedTypes.has(typeRaw) ? typeRaw : 'label';
    let workingText = state.originalText;
    const ensured = ensureGroupUserControlsBlockExists(workingText, bounds, newline, state.parseResult);
    if (!ensured || !ensured.block) throw new Error('Unable to locate Group UserControls.');
    workingText = ensured.text;
    const ucRange = ensured.block.open != null
      ? ensured.block
      : { open: ensured.block.openIndex, close: ensured.block.closeIndex };
    if (ucRange.open == null || ucRange.close == null) throw new Error('Unable to locate Group UserControls.');
    const controlId = generateUniqueControlId(workingText, ucRange, trimmedName);
    const labelDefault = config?.labelDefault === 'open' ? 'open' : 'closed';
    const lines = buildControlDefinitionLines(type, {
      name: trimmedName,
      page: normalizedPage,
      labelCount: config?.labelCount,
      labelDefault,
      comboOptions: Array.isArray(config?.comboOptions) ? config.comboOptions : [],
    });
    const pendingMeta = {
      kind: type === 'label' ? 'label' : type === 'button' ? 'button' : null,
      labelCount: type === 'label' && Number.isFinite(config?.labelCount) ? Number(config.labelCount) : null,
      selectionInsertStartPos: type === 'label' && Number.isFinite(config?.labelSelectionSpan?.startPos)
        ? Number(config.labelSelectionSpan.startPos)
        : null,
      defaultValue: type === 'label' ? (labelDefault === 'closed' ? '0' : '1') : null,
      inputControl:
        type === 'slider' ? 'SliderControl'
        : type === 'combo' ? 'ComboControl'
        : type === 'screw' ? 'ScrewControl'
        : type === 'button' ? 'ButtonControl'
        : type === 'separator' ? 'SeparatorControl'
        : 'LabelControl',
      choiceOptions: type === 'combo' && Array.isArray(config?.comboOptions)
        ? config.comboOptions.map((opt) => String(opt || '').trim()).filter((opt) => opt.length > 0)
        : [],
    };
    workingText = insertUserControlBlock(workingText, ucRange, controlId, lines, newline);
    const macroName = state.parseResult.macroName || state.parseResult.macroNameOriginal || 'Macro';
    rememberPendingControlMeta(macroName, controlId, pendingMeta);
    autoPublishCreatedControl(macroName, controlId, trimmedName, normalizedPage, pendingMeta);
    const persisted = rebuildContentWithNewOrder(workingText, state.parseResult, newline, { safeEditExport: true });
    state.originalText = persisted;
    await reloadMacroFromCurrentText({ skipClear: true });
    runValidation('add-control');
    const typeLabel = type === 'button'
      ? 'button'
      : type === 'slider'
        ? 'slider'
        : type === 'combo'
          ? 'combo control'
          : type === 'screw'
            ? 'screw control'
            : type === 'separator'
              ? 'separator control'
              : 'label';
    info(`Added ${typeLabel} "${trimmedName}" to the macro.`);
    return { nodeName: macroName, controlId, displayName: trimmedName };
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
      if (String(metaWithPage?.kind || '').toLowerCase() === 'label') {
        const selectedSet = state.parseResult?.selected instanceof Set ? state.parseResult.selected : null;
        const selectedPositions = selectedSet
          ? Array.from(selectedSet).map((idx) => currentOrder.indexOf(idx)).filter((idx) => idx >= 0).sort((a, b) => a - b)
          : [];
        if (selectedPositions.length) {
          pos = selectedPositions[0];
          anchored = true;
        }
      }
      if (!anchored && Number.isFinite(metaWithPage?.selectionInsertStartPos)) {
        pos = Math.max(0, Math.min(currentOrder.length, Number(metaWithPage.selectionInsertStartPos)));
        anchored = true;
      }
      if (!anchored && typeof activeDetailEntryIndex === 'number') {
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
      syncEntrySortIndicesToOrder(state.parseResult);
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

  // Macro/group UserControls are their own ownership plane.
  // They belong to the macro body itself and must not be treated as published Inputs.
  // All macro-level mutations should flow through this context + writer path rather than
  // ad hoc text surgery elsewhere in export/rebuild code.
  function ensureGroupUserControlsWriterContext(text, result, eol) {
    try {
      const normalized = normalizeGroupUserControlsBlocks(text, result, eol || '\n');
      const bounds = locateMacroGroupBounds(normalized, result);
      if (!bounds) return { text: normalized, bounds: null, uc: null };
      const ensured = ensureGroupUserControlsBlockExists(normalized, bounds, eol || '\n', result);
      const updated = ensured.text || normalized;
      const finalBounds = ensured.bounds || locateMacroGroupBounds(updated, result) || bounds;
      const uc = ensured.block || findGroupUserControlsBlock(updated, finalBounds.groupOpenIndex, finalBounds.groupCloseIndex);
      return { text: updated, bounds: finalBounds, uc };
    } catch (_) {
      return { text, bounds: null, uc: null };
    }
  }

  function findGroupUserControlBlockById(text, result, controlId) {
    try {
      if (!text || !controlId) return null;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return null;
      const uc = findGroupUserControlsBlock(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!uc) return null;
      return findControlBlockInUc(text, uc.openIndex, uc.closeIndex, controlId);
    } catch (_) {
      return null;
    }
  }

  function getGroupUserControlBodyById(text, result, controlId) {
    try {
      const block = findGroupUserControlBlockById(text, result, controlId);
      if (!block) return null;
      return text.slice(block.open + 1, block.close);
    } catch (_) {
      return null;
    }
  }

  // Group-level controls are preserved/edited separately from published entries.
  // Unknown controls default to preserve-only and are blocked here unless a caller
  // explicitly opts into mutating them.
  function upsertGroupUserControl(text, result, controlId, options, eol) {
    try {
      if (!text || !result || !controlId) return text;
      const newline = eol || '\n';
      const policy = canMutateGroupUserControl(result, controlId, {
        ...(options || {}),
        action: 'update',
      });
      if (!policy.allowed) {
        recordGroupUserControlMutationIssue(result, {
          control: controlId,
          action: 'update',
          message: policy.message,
        });
        return text;
      }
      let context;
      if (options?.createIfMissing === false) {
        const normalized = normalizeGroupUserControlsBlocks(text, result, newline);
        const bounds = locateMacroGroupBounds(normalized, result);
        const uc = bounds ? findGroupUserControlsBlock(normalized, bounds.groupOpenIndex, bounds.groupCloseIndex) : null;
        context = { text: normalized, bounds, uc };
      } else {
        context = ensureGroupUserControlsWriterContext(text, result, newline);
      }
      let updated = context.text;
      if (!context.bounds || !context.uc) return updated;
      let block = findControlBlockInUc(updated, context.uc.openIndex, context.uc.closeIndex, controlId);
      if (!block) {
        if (options?.createIfMissing === false) return updated;
        const lines = (typeof options?.createLines === 'function')
          ? options.createLines()
          : (Array.isArray(options?.createLines) ? options.createLines : []);
        updated = insertUserControlBlock(updated, { open: context.uc.openIndex, close: context.uc.closeIndex }, controlId, lines, newline);
        block = findGroupUserControlBlockById(updated, result, controlId);
        if (!block) return updated;
      }
      const indent = (getLineIndent(updated, block.open) || '') + '\t';
      let body = updated.slice(block.open + 1, block.close);
      if (typeof options?.updateBody === 'function') {
        body = options.updateBody(body, indent, newline);
      }
      updated = updated.slice(0, block.open + 1) + body + updated.slice(block.close);
      return normalizeGroupUserControlsBlocks(updated, result, newline);
    } catch (_) {
      return text;
    }
  }

  // Removal follows the same ownership boundary as updates: macro-level controls are
  // never removed through the published Inputs pipeline.
  function removeGroupUserControlById(text, result, controlId) {
    try {
      if (!text || !controlId) return text;
      const newline = detectNewline(text) || '\n';
      const policy = canMutateGroupUserControl(result, controlId, { action: 'remove' });
      if (!policy.allowed) {
        recordGroupUserControlMutationIssue(result, {
          control: controlId,
          action: 'remove',
          message: policy.message,
        });
        return text;
      }
      const context = ensureGroupUserControlsWriterContext(text, result, newline);
      let updated = context.text;
      if (!context.bounds || !context.uc) return updated;
      const block = findControlBlockInUc(updated, context.uc.openIndex, context.uc.closeIndex, controlId);
      if (!block) return updated;
      let start = Number.isFinite(block.idStart) ? block.idStart : block.open;
      while (start > context.uc.openIndex && /\s/.test(updated[start - 1])) start--;
      let end = block.close + 1;
      while (end < context.uc.closeIndex && /\s/.test(updated[end])) end++;
      if (updated[end] === ',') {
        end++;
        while (end < context.uc.closeIndex && /\s/.test(updated[end])) end++;
      }
      updated = updated.slice(0, start) + updated.slice(end);
      return normalizeGroupUserControlsBlocks(updated, result, newline);
    } catch (_) {
      return text;
    }
  }

  function findEntryControlBlockBounds(entry) {
    try {
      if (!entry || !entry.source || !state.originalText || !state.parseResult) return null;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) return null;
      let toolBlock = null;
      if (entry.sourceOp) {
        toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
        if (!toolBlock) toolBlock = findToolBlockAnywhere(state.originalText, entry.sourceOp);
      }
      if (toolBlock) {
        let cb = findControlBlockInTool(state.originalText, toolBlock.open, toolBlock.close, entry.source);
        if (!cb) {
          const toolUc = findUserControlsInTool(state.originalText, toolBlock.open, toolBlock.close);
          if (toolUc) cb = findControlBlockInUc(state.originalText, toolUc.open, toolUc.close, entry.source);
        }
        if (cb) return cb;
      }
      const gcb = findGroupUserControlBlockById(state.originalText, state.parseResult, entry.source);
      if (gcb) return gcb;
      return null;
    } catch (_) {
      return null;
    }
  }

  function findUserControlBlockForToolOrGroup(text, groupOpen, groupClose, toolName, controlId) {
    try {
      if (!text || controlId == null) return null;
      if (toolName) {
        const toolBlock = findToolBlockInGroup(text, groupOpen, groupClose, toolName);
        if (toolBlock) {
          const toolUc = findUserControlsInTool(text, toolBlock.open, toolBlock.close);
          if (toolUc) {
            const toolControl = findControlBlockInUc(text, toolUc.open, toolUc.close, controlId);
            if (toolControl) {
              return {
                scope: 'tool',
                toolBlock,
                uc: toolUc,
                controlBlock: toolControl,
              };
            }
          }
        }
      }
      const groupUc = findGroupUserControlsBlock(text, groupOpen, groupClose);
      if (!groupUc) return null;
      const groupControl = findControlBlockInUc(text, groupUc.openIndex, groupUc.closeIndex, controlId);
      if (!groupControl) return null;
      return {
        scope: 'group',
        toolBlock: null,
        uc: { open: groupUc.openIndex, close: groupUc.closeIndex },
        controlBlock: groupControl,
      };
    } catch (_) {
      return null;
    }
  }

  function parseComboOptionsText(raw) {
    const text = String(raw || '').replace(/\r\n?/g, '\n');
    const rows = text.includes('\n') ? text.split('\n') : text.split(',');
    return rows
      .map((line) => String(line || '').trim().replace(/^"(.*)"$/, '$1'))
      .filter((line) => line.length > 0);
  }

  function setEntryComboOptions(index, options, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
    const entry = state.parseResult.entries[index];
    if (!entry) return false;
    const normalized = Array.isArray(options)
      ? options.map((opt) => String(opt || '').trim()).filter((opt) => opt.length > 0)
      : [];
    if (!normalized.length) {
      if (!opts.silent) error('Choice controls require at least one option.');
      return false;
    }
    const current = getChoiceOptions(entry);
    if (current.length === normalized.length && current.every((v, i) => String(v) === String(normalized[i]))) {
      return false;
    }
    const applyChoicesToBody = (body, indent, eol) => {
      let nextBody = removeControlListPropEntries(body, 'CCS_AddString');
      nextBody = removeControlListPropEntries(nextBody, 'MBTNC_AddButton');
      const listProp = getChoiceListPropForEntry(entry);
      if (nextBody.trim().length) {
        const trimmed = nextBody.replace(/\s+$/g, '');
        if (trimmed && !trimmed.endsWith(',')) {
          nextBody = trimmed + ',' + nextBody.slice(trimmed.length);
        }
        if (!nextBody.endsWith(eol)) nextBody += eol;
      }
      normalized.forEach((opt) => {
        nextBody += `${indent}{ ${listProp} = "${escapeQuotes(opt)}" },${eol}`;
      });
      return nextBody;
    };

    let changed = false;
    let pendingText = null;
    let pendingRaw = null;
    const cb = findEntryControlBlockBounds(entry);
    const eol = state.newline || '\n';
    if (cb) {
      const indent = (getLineIndent(state.originalText, cb.open) || '') + '\t';
      const body = state.originalText.slice(cb.open + 1, cb.close);
      const rebuilt = applyChoicesToBody(body, indent, eol);
      const nextText = state.originalText.slice(0, cb.open + 1) + rebuilt + state.originalText.slice(cb.close);
      if (nextText !== state.originalText) {
        pendingText = nextText;
        changed = true;
      }
    } else if (typeof entry.raw === 'string') {
      const rawOpen = entry.raw.indexOf('{');
      const rawClose = entry.raw.lastIndexOf('}');
      if (rawOpen >= 0 && rawClose > rawOpen) {
        const rawIndent = (getLineIndent(entry.raw, rawOpen) || '') + '\t';
        const rawBody = entry.raw.slice(rawOpen + 1, rawClose);
        const rebuiltRaw = applyChoicesToBody(rawBody, rawIndent, eol);
        const nextRaw = entry.raw.slice(0, rawOpen + 1) + rebuiltRaw + entry.raw.slice(rawClose);
        if (nextRaw !== entry.raw) {
          pendingRaw = nextRaw;
          changed = true;
        }
      }
    }
    if (!changed) {
      if (!opts.silent) error('Unable to locate control definition for options.');
      return false;
    }
    if (!opts.skipHistory) pushHistory('edit combo options');
    if (pendingText != null) state.originalText = pendingText;
    if (pendingRaw != null) entry.raw = pendingRaw;
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.choiceOptions = [...normalized];
    if (!Array.isArray(entry.controlMetaOriginal.choiceOptions) || !entry.controlMetaOriginal.choiceOptions.length) {
      entry.controlMetaOriginal.choiceOptions = [...normalized];
    }
    entry.controlMetaDirty = true;
    if (!opts.silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    markContentDirty();
    return true;
  }

  function setEntryMultiButtonShowBasic(index, enabled, opts = {}) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
    const entry = state.parseResult.entries[index];
    if (!entry || !isMultiButtonControl(entry)) return false;
    const nextValue = enabled ? 'true' : 'false';
    const currentValue = parseControlBooleanValue(
      entry?.controlMeta?.multiButtonShowBasic ?? entry?.controlMetaOriginal?.multiButtonShowBasic,
      false
    ) ? 'true' : 'false';
    if (currentValue === nextValue) return false;

    let changed = false;
    let pendingText = null;
    let pendingRaw = null;
    const cb = findEntryControlBlockBounds(entry);
    const eol = state.newline || '\n';
    if (cb) {
      const indent = (getLineIndent(state.originalText, cb.open) || '') + '\t';
      const body = state.originalText.slice(cb.open + 1, cb.close);
      const rebuilt = upsertControlProp(body, 'MBTNC_ShowBasicButton', nextValue, indent, eol);
      const nextText = state.originalText.slice(0, cb.open + 1) + rebuilt + state.originalText.slice(cb.close);
      if (nextText !== state.originalText) {
        pendingText = nextText;
        changed = true;
      }
    } else if (typeof entry.raw === 'string') {
      const rawOpen = entry.raw.indexOf('{');
      const rawClose = entry.raw.lastIndexOf('}');
      if (rawOpen >= 0 && rawClose > rawOpen) {
        const rawIndent = (getLineIndent(entry.raw, rawOpen) || '') + '\t';
        const rawBody = entry.raw.slice(rawOpen + 1, rawClose);
        const rebuiltRaw = upsertControlProp(rawBody, 'MBTNC_ShowBasicButton', nextValue, rawIndent, eol);
        const nextRaw = entry.raw.slice(0, rawOpen + 1) + rebuiltRaw + entry.raw.slice(rawClose);
        if (nextRaw !== entry.raw) {
          pendingRaw = nextRaw;
          changed = true;
        }
      }
    }
    if (!changed) {
      if (!opts.silent) error('Unable to locate control definition for multi button style.');
      return false;
    }
    if (!opts.skipHistory) pushHistory('edit multibutton style');
    if (pendingText != null) state.originalText = pendingText;
    if (pendingRaw != null) entry.raw = pendingRaw;
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.multiButtonShowBasic = nextValue;
    if (entry.controlMetaOriginal.multiButtonShowBasic == null || String(entry.controlMetaOriginal.multiButtonShowBasic).trim() === '') {
      entry.controlMetaOriginal.multiButtonShowBasic = nextValue;
    }
    entry.controlMetaDirty = true;
    if (!opts.silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    markContentDirty();
    return true;
  }

  function setEntryTextLines(index, value) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    if (!isTextControl(entry)) return;
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
    );
    if (inputControl && !/texteditcontrol/i.test(inputControl)) return;
    const num = Number(String(value || '').trim());
    if (!Number.isFinite(num)) return;
    const nextVal = Math.max(1, Math.min(20, Math.round(num)));
    const currentRaw = entry.controlMeta?.textLines ?? entry.controlMetaOriginal?.textLines ?? null;
    const current = currentRaw != null ? Number(String(currentRaw).replace(/"/g, '').trim()) : null;
    if (Number.isFinite(current) && current === nextVal) return;
    pushHistory('edit text lines');
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.textLines = String(nextVal);
    entry.controlMetaDirty = true;

    const cb = findEntryControlBlockBounds(entry);
    if (!cb) return;
    const indent = (getLineIndent(state.originalText, cb.open) || '') + '\t';
    let body = state.originalText.slice(cb.open + 1, cb.close);
    body = upsertControlProp(body, 'TEC_Lines', String(nextVal), indent, state.newline || '\n');
    state.originalText = state.originalText.slice(0, cb.open + 1) + body + state.originalText.slice(cb.close);
    updateCodeView(state.originalText || '');
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

  const NUMBER_INPUT_CONTROLS = ['SliderControl','ScrewControl','CheckboxControl','ButtonControl','ComboControl','MultiButtonControl','LabelControl','SeparatorControl'];
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
    if (norm && /multibuttoncontrol/i.test(norm)) {
      const options = getChoiceOptions(entry);
      const seed = options.length ? options : ['Option 1', 'Option 2', 'Option 3'];
      setEntryComboOptions(index, seed, { skipHistory: true, silent: true });
    } else if (norm && /combocontrol/i.test(norm)) {
      const options = getChoiceOptions(entry);
      if (!options.length) {
        setEntryComboOptions(index, ['Option 1', 'Option 2', 'Option 3'], { skipHistory: true, silent: true });
      }
    }
    if (!silent) {
      renderActiveList({ safe: true });
      renderDetailDrawer(index);
    }
    markContentDirty();
  }

  function handleDetailTargetChange(index) {
    if (index == null || index < 0) {
      activeDetailEntryIndex = null;
      if (nodesHidden) {
        renderMacroInfoDrawer({ expanded: true });
      } else {
        collapseDetailDrawer();
      }
      return;
    }
    renderDetailDrawer(index);
  }

  function getSelectedDrawerNavigation(index) {
    try {
      if (!state.parseResult || !Array.isArray(state.parseResult.order)) return null;
      const selected = state.parseResult.selected instanceof Set ? state.parseResult.selected : null;
      if (!selected || selected.size < 2) return null;
      const orderedSelected = state.parseResult.order.filter((entryIndex) => selected.has(entryIndex));
      if (orderedSelected.length < 2) return null;
      const currentPos = orderedSelected.indexOf(index);
      if (currentPos < 0) return null;
      return {
        total: orderedSelected.length,
        currentPos,
        previousIndex: currentPos > 0 ? orderedSelected[currentPos - 1] : null,
        nextIndex: currentPos < orderedSelected.length - 1 ? orderedSelected[currentPos + 1] : null,
      };
    } catch (_) {
      return null;
    }
  }

  function focusDrawerSelectedEntry(index) {
    try {
      if (typeof index !== 'number' || index < 0) return;
      if (typeof setPublishedDetailTarget === 'function') {
        setPublishedDetailTarget(index);
        return;
      }
      handleDetailTargetChange(index);
    } catch (_) {}
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
    if (nodesHidden && !detailDrawerCollapsed) return;
    if (detailDrawerCollapsed) {
      if (activeDetailEntryIndex == null) {
        renderMacroInfoDrawer({ expanded: true });
        return;
      }
      renderDetailDrawer(activeDetailEntryIndex);
      return;
    }
    collapseDetailDrawer();
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

  try {
    let syncDrawerAnchorRaf = 0;
    const scheduleSyncDetailDrawerAnchor = () => {
      if (syncDrawerAnchorRaf) return;
      syncDrawerAnchorRaf = requestAnimationFrame(() => {
        syncDrawerAnchorRaf = 0;
        try { syncDetailDrawerAnchor(); } catch (_) {}
      });
    };
    window.addEventListener('resize', scheduleSyncDetailDrawerAnchor);
  } catch (_) {}

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
    getCurrentValueInfo: (entry) => getCurrentInputInfo(entry),
    shouldShowCurrentValues: () => showCurrentValues,
    onDetailTargetChange: (index) => {
      try { handleDetailTargetChange(typeof index === 'number' ? index : null); } catch (_) {}
    },
    applyNodeControlMeta: (entry, meta) => applyNodeControlMeta(entry, meta),
    onRenderList: () => markContentDirty(),
    onEntryMutated: () => markContentDirty(),
  });

  const renderList = publishedControls.renderList;
  const setPublishedFilter = publishedControls.setFilter;
  const updateRemoveSelectedState = publishedControls.updateRemoveSelectedState;
  const cleanupDropIndicator = publishedControls.cleanupDropIndicator;
  const updateDropIndicatorPosition = publishedControls.updateDropIndicatorPosition;
  const getInsertionPosUnderSelection = publishedControls.getInsertionPosUnderSelection;
  const addOrMovePublishedItemsAt = publishedControls.addOrMovePublishedItemsAt;
  const createIcon = publishedControls.createIcon;
  const getCurrentDropIndex = publishedControls.getCurrentDropIndex;
  const getPublishedSelectionSpan = publishedControls.getSelectionSpan;
  const setPublishedDetailTarget = publishedControls.setDetailTargetIndex;
  const refreshPageTabs = publishedControls.refreshPageTabs;
  const setEntryDisplayNameApi = publishedControls.setEntryDisplayName;
  const appendLauncherUiApi = publishedControls.appendLauncherUi;
  renderIconHtml = typeof createIcon === 'function' ? createIcon : (() => '');
  updateDrawerToggleLabel();

  function renderActiveList(options = {}) {
    if (!state.parseResult) return;
    if (options.safe) {
      try {
        renderList(state.parseResult.entries, state.parseResult.order);
        updatePublishedFeatureBadges();
      } catch (_) {}
      return;
    }
    renderList(state.parseResult.entries, state.parseResult.order);
    updatePublishedFeatureBadges();
  }

  function getAdaptiveHistoryLimit() {
    try {
      const textLen = (state.originalText || '').length;
      const controls = Array.isArray(state.parseResult?.entries) ? state.parseResult.entries.length : 0;
      if (textLen >= 1_200_000 || controls >= 500) return 20;
      if (textLen >= 700_000 || controls >= 300) return 35;
      if (textLen >= 300_000 || controls >= 160) return 50;
      return 80;
    } catch (_) {
      return 80;
    }
  }

  nodesPane = createNodesPane({
    state,
    nodesList,
    nodesSearch,
    hideReplacedEl,
    quickClickHintEl,
    viewControlsBtn,
    viewControlsMenu,
    showAllNodesBtn,
    showPublishedNodesBtn,
    collapseAllNodesBtn,
    showNodeTypeLabelsEl,
    showInstancedConnectionsEl,
    showNextNodeLinksEl,
    autoGroupQuickSetsEl,
    nameClickQuickSetEl,
    quickSetBlendToCheckboxEl,
    publishSelectedBtn,
    clearNodeSelectionBtn,
    importCatalogBtn,
    catalogInput,
    importModifierCatalogBtn,
    modifierCatalogInput,
    logDiag,
    logTag,
    error,
    info,
    highlightNode,
    clearHighlights,
    renderList,
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
    requestRenameNode: (nodeName) => promptRenameNode(nodeName),
    getPendingControlMeta: (op, id) => getPendingControlMeta(op, id),
    consumePendingControlMeta: (op, id) => consumePendingControlMeta(op, id),
    isPickSessionActive: () => !!activePickSession,
    focusPublishedControl: (sourceOp, source) => focusPublishedControl(sourceOp, source),
  });

  historyController = createHistoryController({
    state,
    renderList,
    ctrlCountEl,
    getNodesPane: () => nodesPane,
    undoBtn,
    redoBtn,
    logDiag,
    doc: (typeof document !== 'undefined') ? document : null,
    captureExtraState: () => captureHistoryExtras(),
    restoreExtraState: (extra, context) => restoreHistoryExtras(extra, context),
    getHistoryLimit: () => getAdaptiveHistoryLimit(),
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
      onNormalizeLegacyNames: () => normalizeLegacyNamesMenu(),
      onImportCsvFile: () => importCsvFromFile(),
      onImportCsvUrl: () => importCsvFromUrl(),
      onImportGoogleSheet: () => importGoogleSheetFromUrl(),
      onReloadCsv: () => reloadCsvSource(),
      onGenerateFromCsv: () => generateSettingsFromCsv(),
      onOpenPresetEngine: () => openPresetEngineModal(),
      onRefreshSession: () => refreshCurrentSession(),
      onInsertUpdateDataButton: () => insertUpdateDataButton(),
      onHeaderImageDialog: () => openHeaderImageDialog(),
      onProtocolUrl: (url) => handleProtocolUrl(url),
    });

  function setExportButtonLabel(label) {
    const next = label || DEFAULT_EXPORT_LABEL;
    if (exportBtn) exportBtn.textContent = next;
    syncPrimaryExportButtonLabel();
  }

  async function updateExportButtonLabelFromPath(filePath) {
    if (!filePath || !nativeBridge || !nativeBridge.ipcRenderer) {
      if (exportBtn) setExportButtonLabel(DEFAULT_EXPORT_LABEL);
      state.drfxLink = null;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      updateDataLinkStatus();
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
      updateDataLinkStatus();
    } catch (_) {
      if (exportBtn) setExportButtonLabel(DEFAULT_EXPORT_LABEL);
      state.drfxLink = null;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      updateDataLinkStatus();
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
      const selectedDocs = documents.filter(doc => doc && doc.selected);
      if (selectedDocs.length) {
        selectedDocs.forEach((doc) => {
          if (doc.snapshot) {
            doc.snapshot.exportFolder = folderPath;
            doc.snapshot.exportFolderSelected = true;
          }
        });
        state.exportFolder = folderPath;
        state.exportFolderSelected = true;
        updateDocExportPathDisplay();
        updateDataLinkStatus();
        info(`Export folder set for ${selectedDocs.length} selected tab(s).`);
      } else {
        setExportFolder(folderPath, { selected: true });
      }
    });
  }

  // Make macro name editable in-place
  try {
    if (macroNameEl) {
      macroNameEl.spellcheck = false;
      macroNameEl.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter') { ev.preventDefault(); macroNameEl.blur(); }
      });
      macroNameEl.addEventListener('blur', () => {
        try {
          const newName = (getMacroNameDisplay() || '').trim();
          if (!state.parseResult) return;
          const prevName = state.parseResult.macroName || '';
          if (newName) {
            state.parseResult.macroName = newName;
            updateActiveDocumentMeta({ name: newName });
            if (newName !== prevName) markActiveDocumentDirty();
          } else {
            // Revert to original if cleared
            const fallback = state.parseResult.macroNameOriginal || state.parseResult.macroName || 'Unknown';
            setMacroNameDisplay(fallback);
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

  // Nodes state handled in nodesPane.js
 
  // Published search/filter

  publishedSearch?.addEventListener('input', (e) => {
    setPublishedFilter(e.target.value || '');
    renderActiveList();
  });
  showCurrentValuesToggle?.addEventListener('change', () => {
    showCurrentValues = !!showCurrentValuesToggle.checked;
    renderActiveList();
  });


  // File picker
  fileInput?.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length) {
      logDiag(`fileInput change - files: ${files.length}`);
      await handleFiles(files);
    }
  });

  // Native "Open .setting" (Electron)
  if (openNativeBtn) {
    openNativeBtn.addEventListener('click', async () => {
      await handleNativeOpen();
    });
  }

  // Shared handler for native open (Electron)
  async function handleNativeOpen() {
    try {
      const native = getNativeApi();
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
      const files = Array.isArray(res.files) && res.files.length
        ? res.files
        : [{ filePath: res.filePath || '', baseName: res.baseName || '', content: res.content || '' }];
      let loaded = 0;
      for (const file of files) {
        const name = file?.baseName || file?.filePath || 'Imported.setting';
        const text = file?.content || '';
        if (!text) continue;
        // eslint-disable-next-line no-await-in-loop
        await loadSettingText(name, text);
        setExportFolderFromFilePath(file?.filePath || '', { silent: true });
        loaded += 1;
      }
      if (!loaded) {
        error('Selected file was empty or could not be read.');
        return;
      }
      info(loaded === 1 ? `Loaded macro from file: ${files[0]?.baseName || files[0]?.filePath || 'Imported.setting'}` : `Loaded ${loaded} macros from files.`);
    } catch (err) {
      error('Native open failed: ' + (err?.message || err));
    }
  }

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

  createFileDropController({
    dropZone,
    updateDropHint,
    logDiag,
    handleFiles,
    handleDroppedPath,
    handleDrfxFiles,
    extractFilePathFromData,
    isDrfxPath,
    isElectron: IS_ELECTRON,
  });

  function setExportFolder(folderPath, options = {}) {
    const { silent = false, selected = false } = options || {};
    if (!folderPath) return;
    state.exportFolder = folderPath;
    if (selected) state.exportFolderSelected = true;
    if (!silent) info(`Export folder set to ${folderPath}`);
    const doc = getActiveDocument();
    if (doc) storeDocumentSnapshot(doc);
    updateDocExportPathDisplay();
    updateDataLinkStatus();
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

  function ensureDefaultMacroExplorerFolders() {
    if (!IS_ELECTRON) return;
    const templates = getFusionTemplatesPath();
    if (!templates) return;
    try {
      // eslint-disable-next-line global-require
      const fs = require('fs');
      // eslint-disable-next-line global-require
      const path = require('path');
      if (!fs.existsSync(templates)) return;
      const editRoot = path.join(templates, 'Edit');
      fs.mkdirSync(editRoot, { recursive: true });
      ['Transitions', 'Effects', 'Titles', 'Generators'].forEach((name) => {
        fs.mkdirSync(path.join(editRoot, name), { recursive: true });
      });
    } catch (_) {}
  }

  function resolveExportFolder() {
    return state.exportFolder || getFusionTemplatesPath();
  }

  function resolveExportFolderStrict() {
    if (state.exportFolder) return state.exportFolder;
    if (state.exportFolderSelected) return getFusionTemplatesPath();
    return '';
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

  function buildExportPath(folderPath, fileName) {
    if (!folderPath) return '';
    try {
      // eslint-disable-next-line global-require
      const path = require('path');
      return path.join(folderPath, fileName);
    } catch (_) {
      const sep = folderPath.endsWith('/') || folderPath.endsWith('\\') ? '' : '/';
      return `${folderPath}${sep}${fileName}`;
    }
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
      // Avoid stripping root blocks on import; parse can safely ignore them.
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
        const hasInputs = /Inputs\s*=/.test(text);
        const hasInstance = text.includes('InstanceInput');
        const hasMacro = /=\s*(GroupOperator|MacroOperator)\s*\{/.test(text);
        const preview = String(text || '').slice(0, 120).replace(/\s+/g, ' ').trim();
        logDiag(`Inputs scan - len: ${text.length}, inputs: ${hasInputs}, instance: ${hasInstance}, macro: ${hasMacro}, head: "${preview}"`);
      } catch (e) {
        logDiag(`Diagnostics error while scanning Inputs: ${e.message || e}`);
      }
      state.parseResult = parseSetting(text);
      state.parseResult.nativeVersionsCount = detectNativeVersionsCountFromText(text);
      const toolNames = extractToolNamesFromSetting(text, state.parseResult);
      const toolTypes = extractToolTypesFromSetting(text, state.parseResult);
      const controlNames = new Set();
      for (const entry of (state.parseResult.entries || [])) {
        const toolName = entry?.sourceOp;
        const controlName = entry?.source;
        if (isLuaIdentifier(toolName)) toolNames.add(toolName);
        if (isLuaIdentifier(controlName)) controlNames.add(controlName);
      }
      state.parseResult.luaToolNames = toolNames;
      state.parseResult.luaToolTypes = toolTypes;
      state.parseResult.luaControlNames = controlNames;
      hydrateGroupUserControlsState(state.originalText, state.parseResult);
      syncVisibleGroupUserControlsIntoEntries(state.originalText, state.parseResult);
      state.parseResult.pageOrder = derivePageOrderFromEntries(state.parseResult.entries);
      state.parseResult.activePage = null;
      hydrateEntryPagesFromUserControls(state.originalText, state.parseResult);
      hydrateControlPageIcons(state.originalText, state.parseResult);
      runValidation('load');
      hydrateBlendToggleState(state.originalText, state.parseResult);
      hydrateOnChangeScripts(state.originalText, state.parseResult);
      hydrateLabelVisibility(state.originalText, state.parseResult);
      hydrateControlMetadata(state.originalText, state.parseResult);
      hydrateFlipPairGroups(state.parseResult, state.newline || '\n');
      if (!state.parseResult.operatorType) state.parseResult.operatorType = state.parseResult.operatorTypeOriginal || 'GroupOperator';
      const hiddenLink = extractHiddenDataLink(state.originalText, state.parseResult)
        || extractHiddenDataLinkFallback(state.originalText);
      if (hiddenLink) {
        state.parseResult.dataLink = hiddenLink;
      }
      if (state.parseResult.dataLink) {
        applyFmrDataLinkToEntries(state.parseResult);
      }
      const headerTargetsOnLoad = findHeaderLabelCarrierTargets(state.parseResult);
      headerTargetsOnLoad.forEach((item) => markHeaderCarrierEntryInternal(item.entry));
      ensureHeaderLabelCarrierEntry(state.parseResult);
      hideHeaderCarrierEntriesFromPublishedOrder(state.parseResult);
      const hiddenFileMeta = extractHiddenFileMeta(state.originalText, state.parseResult);
      if (hiddenFileMeta) {
        state.parseResult.fileMeta = hiddenFileMeta;
        if (!state.originalFilePath && hiddenFileMeta.exportPath) {
          state.originalFilePath = hiddenFileMeta.exportPath;
        }
      }
      const hiddenPresetData = extractHiddenPresetData(state.originalText, state.parseResult);
      if (hiddenPresetData && typeof hiddenPresetData === 'object') {
        state.parseResult.presetEngine = hiddenPresetData;
      }
      ensurePresetEngine(state.parseResult);
      syncPresetSelectorEntryToMacro(
        state.parseResult,
        resolveMacroGroupSourceOp(state.originalText, state.parseResult)
          || String(state.parseResult.macroName || state.parseResult.macroNameOriginal || '').trim()
      );
      hydrateRoutingInputsState(state.originalText, state.parseResult);
      updateDataMenuState();
      syncDataLinkPanel();
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
      state.parseResult.collapsedLabels = new Set();
      state.parseResult.collapsedCG = new Set();
      state.parseResult.selected = new Set();
      if (!state.parseResult.entries.length) {
        info('No published controls found. Use Create New Control in the Published Controls panel to add controls.');
      }
      setMacroNameDisplay(state.parseResult.macroName || 'Unknown');
      if (ctrlCountEl) ctrlCountEl.textContent = String(state.parseResult.entries.length);
      if (inputCountEl) inputCountEl.textContent = String(countRecognizedInputs(state.parseResult.entries));
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
      setCreateControlActionsEnabled(true);
      state.parseResult.history = [];
      state.parseResult.future = [];
      updateUndoRedoState();
      renderActiveList();
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
        updateIntroToggleVisibility();
        syncDataLinkPanel();
        if (pendingReloadAfterImport) {
          pendingReloadAfterImport = false;
          const source = state.parseResult?.dataLink?.source || '';
          if (/^https?:/i.test(String(source || ''))) {
            try { await reloadDataLinkForCurrentMacro(); } catch (_) {}
          }
        }
        if (IS_ELECTRON) setIntroCollapsed(true);
      } catch (err) {
      const msg = err.message || String(err);
      error(msg);
      logDiag(`Parse error: ${msg}`);
      controlsSection.hidden = true;
      setCreateControlActionsEnabled(false);
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
      const native = getNativeApi();
      const isElectron = !!(native && native.isElectron);
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

      if (!nodesPane?.getNodeProfiles?.()) {
        pending.push((async () => {
          try {
            const resp = await fetch('FusionNodeProfiles.json');
            if (!resp.ok) return;
            const json = await resp.json();
            nodesPane?.setNodeProfiles?.(json || null);
            const count = Object.keys((json && json.types) || {}).length;
            logTag('Catalog', 'Loaded node profiles via fetch: ' + count + ' tool types');
          } catch (err) {
            logTag('Catalog', 'Node profiles fetch failed: ' + (err?.message || err));
          }
        })());
      }

      if (!nodesPane?.getModifierProfiles?.()) {
        pending.push((async () => {
          try {
            const resp = await fetch('FusionModifierProfiles.json');
            if (!resp.ok) return;
            const json = await resp.json();
            nodesPane?.setModifierProfiles?.(json || null);
            const count = Object.keys((json && json.types) || {}).length;
            logTag('Catalog', 'Loaded modifier profiles via fetch: ' + count + ' modifier types');
          } catch (err) {
            logTag('Catalog', 'Modifier profiles fetch failed: ' + (err?.message || err));
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

    renderActiveList();

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
    const native = getNativeApi();
    if (!native || typeof native.saveDrfxPreset !== 'function' || !state.originalFilePath) {
      error('Active macro is not linked to a DRFX preset.');
      return;
    }
    try {
      const activeDoc = getActiveDocument();
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
        safeEditExport: true,
        includeDataLink: true,
        includePresetRuntime: true,
        preserveAuthoredInputsBlock: !!activeDoc && !activeDoc.isDirty,
      });
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
    const native = getNativeApi();
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
        const content = rebuildContentWithNewOrder(snap.originalText || '', snap.parseResult, snap.newline, {
          includeDataLink: true,
          includePresetRuntime: true,
          preserveAuthoredInputsBlock: !doc.isDirty,
          preserveAuthoredUserControls: !doc.isDirty,
        });
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
      const macroName = (state.parseResult.macroName || '').trim();
      const baseName = macroName ? `${macroName}.setting` : state.originalFileName;
      const outName = suggestOutputName(baseName);
      const defaultPath = buildExportDefaultPath(outName);
      const native = getNativeApi();
      if (native && typeof native.pickSavePath === 'function' && typeof native.writeSettingFile === 'function') {
        try {
          const res = await native.pickSavePath({ defaultPath });
          if (!res || !res.filePath) return;
          const activeDoc = getActiveDocument();
          const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
            debugContext: 'export-file-native',
            includeDataLink: true,
            includePresetRuntime: true,
            fileMetaPath: res.filePath,
            preserveAuthoredInputsBlock: !!activeDoc && !activeDoc.isDirty,
            preserveAuthoredUserControls: !!activeDoc && !activeDoc.isDirty,
          });
          const writeRes = await native.writeSettingFile({ filePath: res.filePath, content: newContent });
          if (writeRes && writeRes.filePath) {
            let savedFileName = '';
            try {
              // eslint-disable-next-line global-require
              const path = require('path');
              savedFileName = path.basename(writeRes.filePath || '');
            } catch (_) {
              savedFileName = String(writeRes.filePath || '').split(/[\\/]/).pop() || '';
            }
            if (savedFileName) {
              state.originalFileName = savedFileName;
              const doc = getActiveDocument();
              if (doc) doc.fileName = savedFileName;
            }
            state.lastExportPath = writeRes.filePath;
            state.originalFilePath = writeRes.filePath;
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
        const activeDoc = getActiveDocument();
        const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
          debugContext: 'export-file-download',
          includeDataLink: true,
          includePresetRuntime: true,
          preserveAuthoredInputsBlock: !!activeDoc && !activeDoc.isDirty,
          preserveAuthoredUserControls: !!activeDoc && !activeDoc.isDirty,
        });
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
      const selectedDocs = documents.filter(doc => doc && doc.selected);
      if (selectedDocs.length) {
        if (!IS_ELECTRON) {
          error('Bulk Export to Edit Page is available in the desktop app.');
          return;
        }
        const batchDoc = selectedDocs.find(doc => doc && doc.isCsvBatch) || null;
        const targetDocs = selectedDocs.filter(doc => doc.snapshot && (doc.snapshot.parseResult || doc.snapshot.originalText) && !doc.isCsvBatch);
        if (!targetDocs.length) {
          if (batchDoc) {
            const snap = batchDoc.csvBatchSnapshot || batchDoc.snapshot || null;
            if (snap) {
              const folderPath = batchDoc.csvBatch?.folderPath || snap.exportFolder || resolveExportFolder();
              await generateSettingsFromCsvCore(snap, {
                filesOnly: true,
                promptNameColumn: false,
                folderPath,
              });
              return;
            }
          }
          error(batchDoc ? 'CSV batch selected. Use Export to regenerate files.' : 'No valid selected tabs to export.');
          return;
        }
        const missing = targetDocs.filter(doc => !(doc.snapshot.exportFolder || state.exportFolder));
        if (missing.length) {
          const msg = 'Choose a destination folder in Macro Explorer before exporting to the Edit Page.';
          try { window.alert(msg); } catch (_) {}
          error(msg);
          return;
        }
        let saved = 0;
        let failed = 0;
        const results = [];
        targetDocs.forEach((doc) => {
          try {
            const snap = doc.snapshot;
            const outName = buildMacroExportNameForSnapshot(snap);
            const targetFolder = snap.exportFolder || state.exportFolder || resolveExportFolder();
            const destPath = buildExportPath(targetFolder, outName);
            const baseText = snap.originalText || '';
            let parsed = snap.parseResult || null;
            if (!parsed && baseText) {
              try {
                parsed = parseSetting(baseText);
                const hiddenLink = extractHiddenDataLink(baseText, parsed) || extractHiddenDataLinkFallback(baseText);
                if (hiddenLink) {
                  parsed.dataLink = hiddenLink;
                  applyFmrDataLinkToEntries(parsed);
                }
                hydrateGroupUserControlsState(baseText, parsed);
                const hiddenPreset = extractHiddenPresetData(baseText, parsed);
                if (hiddenPreset && typeof hiddenPreset === 'object') {
                  parsed.presetEngine = hiddenPreset;
                }
                ensurePresetEngine(parsed);
              } catch (_) {
                parsed = null;
              }
            }
            if (parsed && !parsed.fileMeta) {
              parsed.fileMeta = { exportPath: destPath, buildMarker: FMR_BUILD_MARKER };
            }
            const newContent = parsed
              ? rebuildContentWithNewOrder(baseText, parsed, snap.newline || '\n', {
                  debugContext: 'export-edit-bulk',
                  includeDataLink: true,
                  includePresetRuntime: true,
                  fileMetaPath: destPath,
                  preserveAuthoredInputsBlock: !doc.isDirty,
                  preserveAuthoredUserControls: !doc.isDirty,
                })
              : baseText;
            const finalPath = writeSettingToFolder(targetFolder, outName, newContent);
            snap.lastExportPath = finalPath;
            doc.isDirty = false;
            results.push({ name: outName, path: finalPath, ok: true });
            saved += 1;
          } catch (_) {
            results.push({ name: doc?.name || doc?.fileName || 'Untitled', path: '', ok: false });
            failed += 1;
          }
        });
        const summary = `Exported ${saved} preset(s) to the Edit Page${failed ? `, ${failed} failed.` : '.'}`;
        info(summary);
        try {
          openBulkExportSummaryModal(summary, results);
        } catch (_) {}
        const active = getActiveDocument();
        if (active) storeDocumentSnapshot(active);
        updateDocExportPathDisplay();
        return;
      }
      if (!state.exportFolderSelected) {
        const msg = 'Choose a destination folder in Macro Explorer before exporting to the Edit Page.';
        try { window.alert(msg); } catch (_) {}
        error(msg);
        return;
      }
      const outName = buildMacroExportName();
      if (!IS_ELECTRON) {
        const activeDoc = getActiveDocument();
        const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
          debugContext: 'export-edit-download',
          includeDataLink: true,
          includePresetRuntime: true,
          preserveAuthoredInputsBlock: !!activeDoc && !activeDoc.isDirty,
          preserveAuthoredUserControls: !!activeDoc && !activeDoc.isDirty,
        });
        triggerDownload(outName, newContent);
        info('Exported reordered .setting');
        markActiveDocumentClean();
        return;
      }
      const targetFolder = resolveExportFolder();
      const destPath = buildExportPath(targetFolder, outName);
      if (state.parseResult && !state.parseResult.fileMeta) {
        state.parseResult.fileMeta = { exportPath: destPath, buildMarker: FMR_BUILD_MARKER };
      }
      const activeDoc = getActiveDocument();
      const newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
        debugContext: 'export-edit-single',
        includeDataLink: true,
        includePresetRuntime: true,
        fileMetaPath: destPath,
        preserveAuthoredInputsBlock: !!activeDoc && !activeDoc.isDirty,
        preserveAuthoredUserControls: !!activeDoc && !activeDoc.isDirty,
      });
      const finalPath = writeSettingToFolder(targetFolder, outName, newContent);
      info(`Exported to ${finalPath}`);
      state.lastExportPath = finalPath;
      const doc = getActiveDocument();
      if (doc) storeDocumentSnapshot(doc);
      updateDocExportPathDisplay();
      markActiveDocumentClean();
    } catch (err) {
      error(err.message || String(err));
    }
  }

  function normalizeClipboardSettingText(text) {
    try {
      return String(text || '')
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n/g, '\r\n');
    } catch (_) {
      return String(text || '');
    }
  }

  async function exportToClipboard() {
    if (!state.parseResult) return;
    try {
      synchronizeAutoQuickSetLabelEntries(state.parseResult, state.originalText);
      const stateQuickSetCount = Array.isArray(state.parseResult?.entries)
        ? state.parseResult.entries.filter((entry) => /^MM_QuickSetLabel_/i.test(String(entry?.source || '').trim())).length
        : 0;
      const expectedQuickSetLabels = stateQuickSetCount;
      let newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
        debugContext: 'export-clipboard',
        includeDataLink: true,
        includePresetRuntime: true,
        // Clipboard export should always reflect full MM state, including
        // synthetic macro-root controls (quick-set labels, preset helpers, etc).
        preserveAuthoredInputsBlock: false,
        preserveAuthoredUserControls: false,
      });
      if (expectedQuickSetLabels > 0 && !/MM_QuickSetLabel_/i.test(newContent)) {
        synchronizeAutoQuickSetLabelEntries(state.parseResult, state.originalText);
        newContent = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
          debugContext: 'export-clipboard-retry',
          includeDataLink: true,
          includePresetRuntime: true,
          preserveAuthoredInputsBlock: false,
          preserveAuthoredUserControls: false,
        });
      }
      const clipboardContent = normalizeClipboardSettingText(newContent);
      const builtQuickSetMatches = newContent.match(/MM_QuickSetLabel_[A-Za-z0-9_]+/g) || [];
      const builtQuickSetCount = new Set(builtQuickSetMatches).size;
      const expectsQuickSetLabel = builtQuickSetCount > 0;
      const hasGroupUserControls = /UserControls\s*=\s*ordered\(\)\s*\{/i.test(newContent);
      const native = getNativeApi();
      let copied = false;
      if (native && typeof native.writeClipboard === 'function') {
        try {
          native.writeClipboard(clipboardContent);
          copied = true;
          let readBack = '';
          if (typeof native.readClipboard === 'function') {
            try { readBack = String(native.readClipboard() || ''); } catch (_) { readBack = ''; }
          }
          const readQuickSetMatches = readBack.match(/MM_QuickSetLabel_[A-Za-z0-9_]+/g) || [];
          const readQuickSetCount = new Set(readQuickSetMatches).size;
          const readBackHasQuickSetLabel = readQuickSetCount > 0;
          const readBackHasGroupUc = /UserControls\s*=\s*ordered\(\)\s*\{/i.test(readBack);
          try {
            if (typeof logDiag === 'function') {
              logDiag(`[Clipboard debug] stateQuickSetCount=${stateQuickSetCount}, builtQuickSetCount=${builtQuickSetCount}, builtQuickSetLabel=${expectsQuickSetLabel ? 1 : 0}, builtGroupUc=${hasGroupUserControls ? 1 : 0}, readQuickSetCount=${readQuickSetCount}, readQuickSetLabel=${readBackHasQuickSetLabel ? 1 : 0}, readGroupUc=${readBackHasGroupUc ? 1 : 0}, builtLen=${newContent.length}, readLen=${readBack.length}`);
            }
          } catch (_) {}
          // If native clipboard write appears to drop critical content, fall back to web clipboard API.
          if ((expectsQuickSetLabel && !readBackHasQuickSetLabel) || (hasGroupUserControls && !readBackHasGroupUc)) {
            try {
              await writeToClipboard(clipboardContent);
              copied = true;
            } catch (_) {}
          }
          if (copied) info('Copied reordered .setting to clipboard (native).');
        } catch (err2) {
          error('Native clipboard copy failed: ' + (err2?.message || err2));
        }
      }
      if (!copied) {
        await writeToClipboard(clipboardContent);
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

  function openBulkExportSummaryModal(summary, results = []) {
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal csv-modal';
    const modal = document.createElement('form');
    modal.className = 'add-control-form';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    const header = document.createElement('header');
    const headingWrap = document.createElement('div');
    const eyebrow = document.createElement('p');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Export';
    const heading = document.createElement('h3');
    heading.textContent = 'Bulk Export Summary';
    headingWrap.appendChild(eyebrow);
    headingWrap.appendChild(heading);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = '×';
    closeBtn.setAttribute('aria-label', 'Close');
    header.appendChild(headingWrap);
    header.appendChild(closeBtn);
    const body = document.createElement('div');
    body.className = 'form-body';
    const summaryEl = document.createElement('p');
    summaryEl.textContent = summary;
    const list = document.createElement('div');
    list.style.display = 'flex';
    list.style.flexDirection = 'column';
    list.style.gap = '6px';
    const maxItems = 30;
    const shown = results.slice(0, maxItems);
    shown.forEach((r) => {
      const line = document.createElement('div');
      line.textContent = r.ok ? r.name : `${r.name} (failed)`;
      if (!r.ok) line.style.color = 'var(--danger)';
      list.appendChild(line);
    });
    if (results.length > maxItems) {
      const remaining = results.length - maxItems;
      const more = document.createElement('div');
      more.textContent = `+ ${remaining} more...`;
      more.style.color = 'var(--muted)';
      list.appendChild(more);
    }
    body.appendChild(summaryEl);
    body.appendChild(list);
    const actions = document.createElement('footer');
    const okBtn = document.createElement('button');
    okBtn.type = 'submit';
    okBtn.className = 'primary';
    okBtn.textContent = 'OK';
    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.className = 'secondary';
    copyBtn.textContent = 'Copy List';
    copyBtn.addEventListener('click', async () => {
      try {
        const lines = results.map(r => (r.ok ? r.name : `${r.name} (failed)`)).join('\n');
        if (navigator?.clipboard?.writeText) {
          await navigator.clipboard.writeText(lines);
        } else {
          const ta = document.createElement('textarea');
          ta.value = lines;
          document.body.appendChild(ta);
          ta.select();
          document.execCommand('copy');
          ta.remove();
        }
        info('Copied export list to clipboard.');
      } catch (_) {
        error('Unable to copy export list.');
      }
    });
    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    closeBtn.addEventListener('click', () => close());
    modal.addEventListener('submit', (ev) => {
      ev.preventDefault();
      close();
    });
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay) close();
    });
    if (results.length) actions.appendChild(copyBtn);
    actions.appendChild(okBtn);
    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    okBtn.focus();
  }

  const DRFX_ROOT_OPTIONS = ['Edit', 'Fusion'];
  const DRFX_EDIT_OPTIONS = ['Effects', 'Titles', 'Generators', 'Transitions'];

  function parseDrfxCategoryDefaults(raw) {
    const cleaned = String(raw || '').replace(/\\/g, '/').trim();
    const parts = cleaned.split('/').map(p => p.trim()).filter(Boolean);
    let root = 'Edit';
    let sub = 'Effects';
    let tail = '';
    if (!parts.length) return { root, sub, tail };
    const first = parts[0];
    if (/^fusion$/i.test(first)) {
      root = 'Fusion';
      tail = parts.slice(1).join('/');
      return { root, sub, tail };
    }
    if (/^edit$/i.test(first)) {
      root = 'Edit';
      const second = parts[1];
      if (second && DRFX_EDIT_OPTIONS.some(opt => opt.toLowerCase() === second.toLowerCase())) {
        sub = DRFX_EDIT_OPTIONS.find(opt => opt.toLowerCase() === second.toLowerCase()) || sub;
        tail = parts.slice(2).join('/');
      } else {
        tail = parts.slice(1).join('/');
      }
      return { root, sub, tail };
    }
    // Unknown prefix: keep defaults and treat the full path as the tail.
    tail = parts.join('/');
    return { root, sub, tail };
  }

  function buildDrfxCategoryPath(root, sub, tail) {
    const segments = [];
    const base = String(root || '').trim();
    if (base) segments.push(base);
    if (base === 'Edit') {
      const editSegment = String(sub || '').trim();
      if (editSegment) segments.push(editSegment);
    }
    let rest = String(tail || '').trim();
    if (rest) {
      rest = rest.replace(/^[\\/]+/, '').replace(/\\/g, '/');
    }
    const prefix = segments.join('/');
    if (!rest) return prefix;
    return prefix ? `${prefix}/${rest}` : rest;
  }

  function openDrfxCategoryPrompt({ defaultCategory } = {}) {
    return new Promise((resolve) => {
      const defaults = parseDrfxCategoryDefaults(defaultCategory);
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal';
      overlay.setAttribute('role', 'dialog');
      overlay.setAttribute('aria-modal', 'true');
      const form = document.createElement('form');
      form.className = 'add-control-form';
      const header = document.createElement('header');
      const headerWrap = document.createElement('div');
      const title = document.createElement('h3');
      title.textContent = 'Export DRFX pack';
      headerWrap.appendChild(title);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = 'x';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headerWrap);
      header.appendChild(closeBtn);
      const body = document.createElement('div');
      body.className = 'form-body';

      const row = document.createElement('div');
      row.className = 'field-row path-row';
      const rootLabel = document.createElement('label');
      rootLabel.textContent = 'Page';
      const rootSelect = document.createElement('select');
      DRFX_ROOT_OPTIONS.forEach((opt) => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        rootSelect.appendChild(option);
      });
      rootSelect.value = defaults.root;
      rootLabel.appendChild(rootSelect);

      const rootSeparator = document.createElement('span');
      rootSeparator.className = 'path-separator';
      rootSeparator.textContent = '/';

      const subLabel = document.createElement('label');
      subLabel.textContent = 'Category';
      const subSelect = document.createElement('select');
      DRFX_EDIT_OPTIONS.forEach((opt) => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        subSelect.appendChild(option);
      });
      subSelect.value = defaults.sub;
      subLabel.appendChild(subSelect);

      const subSeparator = document.createElement('span');
      subSeparator.className = 'path-separator';
      subSeparator.textContent = '/';

      const tailLabel = document.createElement('label');
      tailLabel.textContent = 'Additional folders';
      const tailInput = document.createElement('input');
      tailInput.type = 'text';
      tailInput.placeholder = 'Stirling Supply Co';
      tailInput.value = defaults.tail || '';
      tailLabel.appendChild(tailInput);

      row.appendChild(rootLabel);
      row.appendChild(rootSeparator);
      row.appendChild(subLabel);
      row.appendChild(subSeparator);
      row.appendChild(tailLabel);
      body.appendChild(row);

      const errorEl = document.createElement('div');
      errorEl.className = 'form-error';

      const actions = document.createElement('div');
      actions.className = 'modal-actions';
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.textContent = 'Export';
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);

      form.appendChild(header);
      form.appendChild(body);
      form.appendChild(errorEl);
      form.appendChild(actions);
      overlay.appendChild(form);
      document.body.appendChild(overlay);

      const syncSubState = () => {
        const isEdit = rootSelect.value === 'Edit';
        subSelect.disabled = !isEdit;
      };
      syncSubState();

      const cleanup = () => {
        rootSelect.removeEventListener('change', syncSubState);
        cancelBtn.removeEventListener('click', onCancel);
        closeBtn.removeEventListener('click', onCancel);
        form.removeEventListener('submit', onSubmit);
        overlay.removeEventListener('click', onOverlayClick);
        document.removeEventListener('keydown', onKeyDown);
        try { overlay.remove(); } catch (_) {}
      };
      const finish = (value) => {
        cleanup();
        resolve(value);
      };
      const onCancel = () => finish(null);
      const onSubmit = (ev) => {
        ev.preventDefault();
        const categoryPath = buildDrfxCategoryPath(rootSelect.value, subSelect.value, tailInput.value);
        if (!categoryPath) {
          errorEl.textContent = 'Category path is required.';
          return;
        }
        finish(categoryPath);
      };
      const onOverlayClick = (ev) => {
        if (ev.target === overlay) onCancel();
      };
      const onKeyDown = (ev) => {
        if (ev.key === 'Escape') {
          ev.preventDefault();
          onCancel();
        }
      };

      rootSelect.addEventListener('change', syncSubState);
      cancelBtn.addEventListener('click', onCancel);
      closeBtn.addEventListener('click', onCancel);
      form.addEventListener('submit', onSubmit);
      overlay.addEventListener('click', onOverlayClick);
      document.addEventListener('keydown', onKeyDown);

      rootSelect.focus();
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
    const drfxNameRaw = await textPrompt.open({
      title: 'Export DRFX pack',
      label: 'DRFX name',
      initialValue: defaultName,
      confirmText: 'Continue',
    });
    if (drfxNameRaw == null) return;
    const drfxName = String(drfxNameRaw || '').trim();
    if (!drfxName) return;
    const categoryPath = await openDrfxCategoryPrompt({ defaultCategory });
    if (categoryPath == null) return;
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
      const content = rebuildContentWithNewOrder(snap.originalText || '', snap.parseResult, snap.newline, {
        includeDataLink: true,
        includePresetRuntime: true,
        preserveAuthoredInputsBlock: !doc.isDirty,
        preserveAuthoredUserControls: !doc.isDirty,
      });
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

  async function refreshCurrentSession() {
    exportMenuController?.close?.();
    if (!state.parseResult || !state.originalText) {
      error('Load a macro before refreshing.');
      return;
    }
    const started = Date.now();
    try {
      await reloadMacroFromCurrentText({
        skipClear: false,
        preserveHistory: false,
      });
      try { messages.textContent = ''; } catch (_) {}
      try { diagnosticsController?.clear?.(); } catch (_) {}
      historyController?.clearHistory?.();
      info(`Session refreshed in ${Date.now() - started}ms. Undo history reset.`);
    } catch (err) {
      error(err?.message || err || 'Unable to refresh current session.');
    }
  }

  const runDefaultExportAction = async () => {
    exportMenuController?.close?.();
    const activeDoc = getActiveDocument();
    const selectedBatch = getSelectedCsvBatch();
    let batchDoc = (activeDoc && activeDoc.isCsvBatch) ? activeDoc : selectedBatch;
    if (!batchDoc && activeCsvBatchId) {
      batchDoc = documents.find(doc => doc && doc.id === activeCsvBatchId && doc.isCsvBatch) || null;
    }
    const batchSnap = batchDoc ? (batchDoc.csvBatchSnapshot || batchDoc.snapshot) : null;
    if (batchDoc && batchSnap) {
      const snap = batchSnap;
      const folderPath = batchDoc.csvBatch?.folderPath || snap.exportFolder || resolveExportFolder();
      await generateSettingsFromCsvCore(snap, {
        filesOnly: true,
        promptNameColumn: false,
        folderPath,
      });
      return;
    }
    const actionItem = resolveCurrentPrimaryExportItem();
    if (!actionItem || typeof actionItem.action !== 'function' || actionItem.enabled === false) {
      error('No export action is currently available.');
      return;
    }
    syncPrimaryExportButtonLabel();
    await Promise.resolve(actionItem.action());
  };

  exportBtn?.addEventListener('click', runDefaultExportAction);
  exportMenuBtn?.addEventListener('click', runDefaultExportAction);

  exportTemplatesBtn?.addEventListener('click', exportToEditPage);



  // Import .setting content from clipboard (with native/browse fallback)
  async function handleImportFromClipboard() {
    try {
      let text = '';
      const native = getNativeApi();

      // 1) Try native clipboard (Electron)
      if (native && typeof native.readClipboard === 'function') {
        try { text = native.readClipboard() || ''; } catch (_) { text = ''; }
      }

      // 2) If still empty, try browser clipboard
      if ((!text || !text.trim()) && navigator && navigator.clipboard && navigator.clipboard.readText) {
        try { text = await navigator.clipboard.readText(); } catch (_) { /* ignore */ }
      }

      // 3) If still empty: prompt for manual paste
      if (!text || !text.trim()) {
        const pasted = await textPrompt.open({
          title: 'Import from clipboard',
          label: 'Paste .setting content',
          initialValue: '',
          confirmText: 'Import',
          multiline: true,
          placeholder: 'Paste .setting content here',
        });
        if (!pasted || !pasted.trim()) {
          error('Clipboard empty or read denied.');
          return;
        }
        text = pasted;
      }

      await loadSettingText('Clipboard.setting', text);
      info('Loaded macro from clipboard');
    } catch (err) {
      error('Clipboard import failed: ' + (err?.message || err));
    }
  }
  importClipboardBtn?.addEventListener('click', handleImportFromClipboard);

  async function importCsvFromFile() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv,.tsv,.txt';
    input.addEventListener('change', () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const text = String(reader.result || '');
          const parsed = parseCsvText(text);
          if (!parsed.headers.length) {
            error('CSV appears to be empty or invalid.');
            return;
          }
          setCsvData(parsed, file.name || 'CSV');
        } catch (err) {
          error(err?.message || err || 'Failed to parse CSV.');
        }
      };
      reader.onerror = () => {
        error('Failed to read CSV file.');
      };
      reader.readAsText(file);
    });
    input.click();
  }

  async function fetchCsvFromUrl(url, sourceName) {
    try {
      const res = await fetch(url);
      if (!res.ok) {
        error(`CSV fetch failed (${res.status}).`);
        return false;
      }
      const text = await res.text();
      const parsed = parseCsvText(text);
      if (!parsed.headers.length) {
        error('CSV appears to be empty or invalid.');
        return false;
      }
      setCsvData(parsed, sourceName || url);
      return true;
    } catch (err) {
      error(err?.message || err || 'CSV fetch failed.');
      return false;
    }
  }

  function isGoogleSheetsUrl(url) {
    return /docs\.google\.com\/spreadsheets\/d\//i.test(String(url || ''));
  }

  async function importCsvFromUrl() {
    const raw = await openCsvUrlPrompt();
    if (!raw) return;
    const url = normalizeCsvUrl(raw);
    if (!url) return;
    await fetchCsvFromUrl(url, url);
  }

  async function openGoogleSheetsUrlPrompt() {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal csv-modal';
      const modal = document.createElement('form');
      modal.className = 'add-control-form';
      const header = document.createElement('header');
      const headerText = document.createElement('div');
      const eyebrow = document.createElement('div');
      eyebrow.className = 'eyebrow';
      eyebrow.textContent = 'Data';
      const title = document.createElement('h3');
      title.textContent = 'Import Google Sheet (Public URL)';
      headerText.appendChild(eyebrow);
      headerText.appendChild(title);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.innerHTML = '&times;';
      header.appendChild(headerText);
      header.appendChild(closeBtn);
      const body = document.createElement('div');
      body.className = 'form-body';
      const note = document.createElement('p');
      note.className = 'detail-default-note';
      note.textContent = 'Sheet must be shared as "Anyone with the link".';
      body.appendChild(note);
      const label = document.createElement('label');
      label.textContent = 'Google Sheets URL';
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = 'https://docs.google.com/spreadsheets/d/...';
      label.appendChild(input);
      body.appendChild(label);
      const actions = document.createElement('footer');
      actions.className = 'modal-actions';
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.className = 'primary';
      okBtn.textContent = 'Import';
      const close = (value) => {
        try { overlay.remove(); } catch (_) {}
        resolve(value);
      };
      cancelBtn.addEventListener('click', () => close(null));
      closeBtn.addEventListener('click', () => close(null));
      modal.addEventListener('submit', (ev) => {
        ev.preventDefault();
        close(input.value);
      });
      overlay.addEventListener('click', (ev) => {
        if (ev.target === overlay) close(null);
      });
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);
      modal.appendChild(header);
      modal.appendChild(body);
      modal.appendChild(actions);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
      input.focus();
    });
  }

  async function importGoogleSheetFromUrl() {
    const raw = await openGoogleSheetsUrlPrompt();
    if (!raw) return;
    if (!isGoogleSheetsUrl(raw)) {
      error('Please paste a Google Sheets URL.');
      return;
    }
    const url = normalizeCsvUrl(raw);
    if (!url) return;
    await fetchCsvFromUrl(url, raw);
  }

  async function reloadCsvSource() {
    const sourceName = state.csvData?.sourceName || '';
    if (!sourceName || !/^https?:/i.test(sourceName)) {
      error('No URL source to reload.');
      return;
    }
    const url = normalizeCsvUrl(sourceName);
    if (!url) return;
    await fetchCsvFromUrl(url, sourceName);
  }

  const UPDATE_DATA_PROTOCOL = 'macromachine://reload';

  function buildUpdateDataUrl(filePath) {
    const base = UPDATE_DATA_PROTOCOL;
    const pathValue = String(filePath || '').trim();
    if (!pathValue) return base;
    const encoded = encodeURIComponent(pathValue);
    return `${base}?path=${encoded}`;
  }

  function rewriteUpdateDataProtocolPaths(text, filePath) {
    try {
      const url = buildUpdateDataUrl(filePath);
      if (!url || url === UPDATE_DATA_PROTOCOL) return text;
      const re = new RegExp(`${UPDATE_DATA_PROTOCOL}(?!\\?path=)`, 'g');
      return String(text || '').replace(re, url);
    } catch (_) {
      return text;
    }
  }

  function buildLauncherLuaForUrl(url) {
    const safeUrl = String(url || '');
    return [
      `local url = "${safeUrl}"`,
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
      'end',
      '',
    ].join('\n');
  }

  function isUpdateDataButtonEntry(entry) {
    try {
      if (!entry) return false;
      const display = String(entry.name || entry.displayName || '').trim().toLowerCase();
      if (display === 'update data') return true;
      const source = String(entry.source || '').trim().toLowerCase();
      if (source === 'update_data' || source.startsWith('update_data')) return true;
      return false;
    } catch (_) { return false; }
  }

  async function insertUpdateDataButton() {
    if (!state.parseResult || !state.originalText) {
      error('Load a macro before inserting the Update Data button.');
      return;
    }
    const hasLink = !!(state.parseResult?.dataLink?.source || state.csvData?.sourceName);
    if (!hasLink) {
      const proceed = await openConfirmModal({
        title: 'No data link found',
        message: 'This macro does not have a data link yet. Insert the Update Data button anyway?',
        confirmText: 'Insert Button',
        cancelText: 'Cancel',
      });
      if (!proceed) return;
    }
    const primaryTool = pickPrimaryToolNameFromCurrentMacro();
    const useNode = primaryTool && primaryTool !== state.parseResult.macroName;
    const res = useNode
      ? await addControlToNode(primaryTool, { name: 'Update Data', type: 'button', page: 'Controls' })
      : await addControlToGroup({ name: 'Update Data', type: 'button', page: 'Controls' });
    if (!res || !res.controlId) return;
    const macroName = res.nodeName;
    const idx = (state.parseResult.entries || []).findIndex(e => e && e.sourceOp === macroName && e.source === res.controlId);
    if (idx < 0) {
      error('Unable to locate the newly created Update Data button.');
      return;
    }
      const pathHint = state.originalFilePath || state.lastExportPath || state.parseResult?.fileMeta?.exportPath || '';
      const lua = buildLauncherLuaForUrl(buildUpdateDataUrl(pathHint));
      updateEntryButtonExecute(idx, lua, { silent: true, skipHistory: true });
      renderActiveList({ safe: true });
      renderDetailDrawer(idx);
    info('Inserted Update Data button. Use it in Fusion to open Macro Machine.');
  }

    function parseUpdateProtocolUrl(raw) {
      const text = String(raw || '').trim();
      if (!text) return null;
      if (!text.toLowerCase().startsWith(UPDATE_DATA_PROTOCOL)) return null;
      try {
        const u = new URL(text);
        const path = u.searchParams.get('path') || '';
        return { path: path ? decodeURIComponent(path) : '' };
      } catch (_) {
        const q = text.indexOf('?');
        if (q < 0) return { path: '' };
        const query = text.slice(q + 1);
        const params = new URLSearchParams(query);
        const path = params.get('path') || '';
        return { path: path ? decodeURIComponent(path) : '' };
      }
    }

    async function updateFromProtocolPath(filePath) {
      const pathValue = String(filePath || '').trim();
      if (!pathValue) return false;
      const native = getNativeApi();
      if (!native || typeof native.readSettingFile !== 'function') {
        error('Native file access unavailable for Update Data.');
        return false;
      }
      try {
        const res = await native.readSettingFile({ filePath: pathValue });
        if (!res || !res.ok || !res.content) {
          error(res?.error || 'Failed to read the macro file.');
          return false;
        }
        const name = res.baseName || pathValue.split(/[\\/]/).pop() || 'Imported.setting';
        await loadMacroFromText(name, res.content, { preserveFilePath: true, preserveFileInfo: false });
        state.originalFilePath = pathValue;
        state.lastExportPath = pathValue;
        setExportFolderFromFilePath(pathValue, { silent: true });
        const summary = await reloadDataLinkForCurrentMacro({ summary: true, maxLines: 4 });
        const updated = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, {
          includeDataLink: true,
          includePresetRuntime: true,
          fileMetaPath: pathValue,
        });
        let message = 'Overwrite the source file with these changes?';
        if (summary && summary.lines && summary.lines.length) {
          const more = summary.total > summary.lines.length ? summary.total - summary.lines.length : 0;
          message = `Changes:\n${summary.lines.map(line => `- ${line}`).join('\n')}`;
          if (more > 0) message += `\n...and ${more} more.`;
          message += '\n\nOverwrite the source file with these changes?';
        }
        const confirmSave = await openConfirmModal({
          title: 'Data updated',
          message,
          confirmText: 'Overwrite',
          cancelText: 'Not now',
        });
        if (!confirmSave) {
          info('Data updated. Source file not overwritten.');
          return true;
        }
        const writeRes = await native.writeSettingFile({ filePath: pathValue, content: updated });
        if (writeRes && writeRes.filePath) {
          info('Updated macro from data link.');
          return true;
        }
        error(writeRes?.error || 'Failed to write the updated macro.');
        return false;
      } catch (err) {
        error(err?.message || err || 'Failed to update macro from data.');
        return false;
      }
    }

    function handleProtocolUrl(url) {
      const raw = String(url || '').trim();
      if (!raw) return;
      if (!raw.toLowerCase().startsWith(UPDATE_DATA_PROTOCOL)) return;
      const parsed = parseUpdateProtocolUrl(raw);
      if (parsed && parsed.path) {
        updateFromProtocolPath(parsed.path).then((ok) => {
          if (!ok) openUpdateFromMacroModal();
        });
        return;
      }
      openUpdateFromMacroModal();
    }

  function openUpdateFromMacroModal() {
    if (!document || !document.body) return;
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const form = document.createElement('form');
    form.className = 'add-control-form';

    const header = document.createElement('header');
    const headerWrap = document.createElement('div');
    const titleEl = document.createElement('h3');
    titleEl.textContent = 'Update Data from Fusion';
    headerWrap.appendChild(titleEl);

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = 'x';
    closeBtn.setAttribute('aria-label', 'Close');
    header.appendChild(headerWrap);
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'form-body';
    const p = document.createElement('p');
    p.textContent = 'Paste the macro from Fusion, then click Import to reload data links and export again.';
    body.appendChild(p);

    const actions = document.createElement('div');
    actions.className = 'modal-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    const importBtn = document.createElement('button');
    importBtn.type = 'button';
    importBtn.textContent = 'Import from Clipboard';
    actions.appendChild(cancelBtn);
    actions.appendChild(importBtn);

    form.appendChild(header);
    form.appendChild(body);
    form.appendChild(actions);
    overlay.appendChild(form);
    document.body.appendChild(overlay);

    const cleanup = () => {
      try { overlay.remove(); } catch (_) {}
    };
    const onCancel = () => cleanup();
    const onSubmit = (ev) => { ev.preventDefault(); };
    const onOverlayClick = (ev) => { if (ev.target === overlay) onCancel(); };
    const onKeyDown = (ev) => { if (ev.key === 'Escape') { ev.preventDefault(); onCancel(); } };

    cancelBtn.addEventListener('click', onCancel);
    closeBtn.addEventListener('click', onCancel);
    form.addEventListener('submit', onSubmit);
    overlay.addEventListener('click', onOverlayClick);
    document.addEventListener('keydown', onKeyDown, { once: true });

    importBtn.addEventListener('click', () => {
      pendingReloadAfterImport = true;
      try { importClipboardBtn && importClipboardBtn.click(); } catch (_) {}
      cleanup();
    });
  }

  function ensureDataLink(result) {
    if (!result) return null;
    if (!result.dataLink) {
      result.dataLink = {
        version: 1,
        source: state.csvData?.sourceName || '',
        nameColumn: state.csvData?.nameColumn || null,
        rowMode: 'first',
        rowIndex: 1,
        rowKey: '',
        rowValue: '',
      };
    } else {
      if (!result.dataLink.source && state.csvData?.sourceName) {
        result.dataLink.source = state.csvData.sourceName;
      }
      if (!result.dataLink.nameColumn && state.csvData?.nameColumn) {
        result.dataLink.nameColumn = state.csvData.nameColumn;
      }
      if (!result.dataLink.rowMode) result.dataLink.rowMode = 'first';
      if (!Number.isFinite(result.dataLink.rowIndex)) result.dataLink.rowIndex = 1;
    }
    return result.dataLink;
  }

  function ensurePresetEngine(result) {
    if (!result) return null;
    if (!result.presetEngine || typeof result.presetEngine !== 'object') {
      result.presetEngine = {
        version: 1,
        buildMarker: FMR_BUILD_MARKER,
        scope: [],
        scopeEntries: [],
        base: {},
        presets: {},
        activePreset: '',
      };
    }
    const engine = result.presetEngine;
    if (!Array.isArray(engine.scope)) engine.scope = [];
    if (!Array.isArray(engine.scopeEntries)) engine.scopeEntries = [];
    if (!engine.base || typeof engine.base !== 'object') engine.base = {};
    if (!engine.presets || typeof engine.presets !== 'object') engine.presets = {};
    if (typeof engine.activePreset !== 'string') engine.activePreset = '';
    if (typeof engine.buildMarker !== 'string' || !engine.buildMarker.trim()) engine.buildMarker = FMR_BUILD_MARKER;
    engine.version = Math.max(2, Number(engine.version) || 1);
    engine.scopeEntries = normalizePresetScopeEntries(result, engine.scopeEntries, engine.scope);
    engine.scope = engine.scopeEntries.map((item) => item.key);
    return engine;
  }

  function buildPresetScopeEntry(result, entry) {
    try {
      if (!entry) return null;
      const key = String(getEntryKey(entry) || '').trim();
      if (!key) return null;
      const page = String(entry.page || 'Controls').trim() || 'Controls';
      const displayName = String(entry.displayName || entry.name || `${entry.sourceOp || ''}.${entry.source || ''}`).trim() || key;
      const inputControl = String(entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl || '').trim();
      const dataType = String(entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || '').trim();
      const targetType = isGroupUserControlEntry(entry) ? 'groupUserControl' : 'publishedInput';
      return {
        key,
        targetType,
        sourceOp: unquoteSettingValue(entry.sourceOp),
        source: unquoteSettingValue(entry.source),
        entryKey: key,
        page,
        displayName,
        inputControl,
        dataType,
      };
    } catch (_) {
      return null;
    }
  }

  function normalizePresetScopeEntry(result, raw) {
    try {
      if (!raw) return null;
      const rawKey = typeof raw === 'string' ? raw : String(raw.key || raw.entryKey || '').trim();
      if (!rawKey) return null;
      const entries = Array.isArray(result?.entries) ? result.entries : [];
      const matchedEntry = entries.find((entry) => getEntryKey(entry) === rawKey) || null;
      if (matchedEntry) {
        const derived = buildPresetScopeEntry(result, matchedEntry);
        if (derived) return derived;
      }
      const sourceOp = unquoteSettingValue(raw.sourceOp);
      const source = unquoteSettingValue(raw.source);
      const targetType = String(raw.targetType || (sourceOp && source ? 'publishedInput' : 'unknown')).trim();
      return {
        key: rawKey,
        targetType,
        sourceOp,
        source,
        entryKey: rawKey,
        page: String(raw.page || 'Controls').trim() || 'Controls',
        displayName: String(raw.displayName || raw.label || rawKey).trim() || rawKey,
        inputControl: String(raw.inputControl || '').trim(),
        dataType: String(raw.dataType || '').trim(),
      };
    } catch (_) {
      return null;
    }
  }

  function normalizePresetScopeEntries(result, scopeEntries, fallbackScope) {
    try {
      const normalized = [];
      const seen = new Set();
      const pushEntry = (candidate) => {
        const entry = normalizePresetScopeEntry(result, candidate);
        const key = String(entry?.key || '').trim();
        if (!key || seen.has(key)) return;
        seen.add(key);
        normalized.push(entry);
      };
      if (Array.isArray(scopeEntries)) scopeEntries.forEach(pushEntry);
      if (Array.isArray(fallbackScope)) fallbackScope.forEach(pushEntry);
      return normalized;
    } catch (_) {
      return [];
    }
  }

  function getPresetEligibleEntries(result) {
    const out = [];
    if (!result || !Array.isArray(result.entries)) return out;
    const order = Array.isArray(result.order) ? result.order : result.entries.map((_, i) => i);
    const seen = new Set();
    order.forEach((idx) => {
      const entry = result.entries[idx];
      if (!entry || entry.locked || entry.isLabel) return;
      const key = getEntryKey(entry);
      if (!key || seen.has(key)) return;
      if (entry.source === FMR_PRESET_SELECT_CONTROL || key === FMR_PRESET_SELECT_CONTROL || key === FMR_PRESET_DATA_CONTROL) return;
      const inputControl = String(entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl || '').toLowerCase();
      if (inputControl.includes('buttoncontrol') || inputControl.includes('labelcontrol')) return;
      const display = String(entry.displayName || entry.name || `${entry.sourceOp || ''}.${entry.source || ''}`).trim();
      out.push({
        key,
        index: idx,
        label: display,
        scopeEntry: buildPresetScopeEntry(result, entry),
      });
      seen.add(key);
    });
    return out;
  }

  function getEntryFallbackPresetValue(entry) {
    try {
      const profileDefault = getProfileDefaultForEntry(entry);
      const raw = normalizeMetaValue(
        entry?.controlMeta?.defaultValue ??
        entry?.controlMetaOriginal?.defaultValue ??
        profileDefault ??
        extractInstancePropValue(entry?.raw || '', 'Default')
      );
      const formatted = formatDefaultForDisplay(entry, raw);
      return formatted != null ? String(formatted) : '';
    } catch (_) {
      return '';
    }
  }

  function getGroupUserControlFallbackPresetValue(result, scopeEntry, control) {
    try {
      const groupControl = control || getParsedGroupUserControlById(result, scopeEntry?.source || scopeEntry?.key || '');
      const raw = normalizeMetaValue(
        groupControl?.defaultValue ??
        extractControlPropLiteral(groupControl?.rawBody || '', 'INP_Default')
      );
      return raw != null ? String(raw) : '';
    } catch (_) {
      return '';
    }
  }

  function getProfileDefaultForEntry(entry) {
    try {
      if (!entry || !entry.sourceOp || !entry.source) return null;
      const toolType = String(state.parseResult?.luaToolTypes?.get?.(entry.sourceOp) || '').trim();
      if (!toolType) return null;
      const candidateTypes = new Set([toolType]);
      if (/Number$/i.test(toolType)) candidateTypes.add(toolType.replace(/Number$/i, ''));
      if (/Point$/i.test(toolType)) candidateTypes.add(toolType.replace(/Point$/i, ''));
      if (/Image$/i.test(toolType)) candidateTypes.add(toolType.replace(/Image$/i, ''));
      if (/Text$/i.test(toolType)) candidateTypes.add(toolType.replace(/Text$/i, ''));
      const nodeProfiles = nodesPane?.getNodeProfiles?.() || null;
      const modifierProfiles = nodesPane?.getModifierProfiles?.() || null;
      for (const candidateType of candidateTypes) {
        const types = [
          modifierProfiles?.types?.[candidateType],
          nodeProfiles?.types?.[candidateType],
        ].filter(Boolean);
        for (const profile of types) {
          const controls = Array.isArray(profile?.controls) ? profile.controls : [];
          const match = controls.find((item) => String(item?.id || '') === String(entry.source || ''));
          const raw = normalizeMetaValue(match?.defaultValue);
          if (raw != null) return raw;
        }
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function getToolTypeForPresetEntry(text, result, entry) {
    try {
      const sourceOp = String(entry?.sourceOp || '').trim();
      if (!sourceOp) return '';
      const cached = String(
        result?.luaToolTypes?.get?.(sourceOp) ||
        state.parseResult?.luaToolTypes?.get?.(sourceOp) ||
        ''
      ).trim();
      if (cached) return cached;
      if (!text) return '';
      const esc = sourceOp.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp('(^|\\n)\\s*' + esc + '\\s*=\\s*([A-Za-z_][A-Za-z0-9_]*)\\s*\\{');
      const match = re.exec(text);
      return String(match?.[2] || '').trim();
    } catch (_) {
      return '';
    }
  }

  function getSpecialPresetFallbackForEntry(text, result, entry) {
    try {
      if (!entry || !result) return null;
      const source = String(entry.source || '').trim().toLowerCase();
      if (source !== 'source') return null;
      const toolType = String(getToolTypeForPresetEntry(text, result, entry) || '').trim().toLowerCase();
      if (!toolType.includes('switch')) return null;
      return '0';
    } catch (_) {
      return null;
    }
  }

  function getPublishedPresetFallbackValue(text, result, entry) {
    const specialFallback = getSpecialPresetFallbackForEntry(text, result, entry);
    if (specialFallback != null) return String(specialFallback);
    return getEntryFallbackPresetValue(entry);
  }

  function getExplicitSwitchSourcePresetValue(text, result, entry) {
    try {
      const toolType = String(getToolTypeForPresetEntry(text, result, entry) || '').trim().toLowerCase();
      if (!toolType.includes('switch')) return null;
      if (String(entry?.source || '').trim().toLowerCase() !== 'source') return null;
      const toolBlock = findToolBlockAnywhere(text, entry.sourceOp);
      if (!toolBlock) return '0';
      const inputsBlock = findInputsInTool(text, toolBlock.open, toolBlock.close);
      if (!inputsBlock) return '0';
      const items = parseOrderedBlockEntries(text, inputsBlock.open, inputsBlock.close)
        .filter((item) => item && item.name && getOrderedBlockEntryType(text, item) === 'Input');
      const sourceItem = items.find((item) => String(item.name || '').trim() === 'Source');
      if (!sourceItem) return '0';
      const body = text.slice(sourceItem.blockOpen + 1, sourceItem.blockClose);
      const value = extractControlPropLiteral(body, 'Value');
      return value != null && String(value).trim() !== '' ? String(value) : '0';
    } catch (_) {
      return '0';
    }
  }

  function listGroupInputEntries(text, result) {
    try {
      if (!text || !result) return [];
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return [];
      let blocks = findGroupInputsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!Array.isArray(blocks) || !blocks.length) {
        const fallback = findPrimaryInputsBlockFallback(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
        blocks = fallback ? [fallback] : [];
      }
      const inputs = blocks[0];
      if (!inputs) return [];
      return parseOrderedBlockEntries(text, inputs.openIndex, inputs.closeIndex)
        .filter((item) => item && item.name && getOrderedBlockEntryType(text, item) === 'Input');
    } catch (_) {
      return [];
    }
  }

  function findGroupInputBlockForScopeEntry(text, result, scopeEntry) {
    try {
      if (!text || !result || !scopeEntry) return null;
      const controlId = String(scopeEntry.source || scopeEntry.key || '').trim();
      if (!controlId) return null;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return null;
      let blocks = findGroupInputsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!Array.isArray(blocks) || !blocks.length) {
        const fallback = findPrimaryInputsBlockFallback(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
        blocks = fallback ? [fallback] : [];
      }
      const inputs = blocks[0];
      if (!inputs) return null;
      const exact = findInputBlockInInputs(text, inputs.openIndex, inputs.closeIndex, controlId);
      if (exact) return exact;
      const all = listGroupInputEntries(text, result);
      const normControl = normalizeId(controlId);
      const prefixMatches = all.filter((item) => {
        const normName = normalizeId(item.name);
        return normName === `${normControl}1` || normName.startsWith(`${normControl}_`) || normName.startsWith(`${normControl}.`);
      });
      if (prefixMatches.length === 1) {
        return { open: prefixMatches[0].blockOpen, close: prefixMatches[0].blockClose };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function resolvePresetScopeTarget(result, scopeCandidate) {
    try {
      if (!result || !scopeCandidate) return null;
      const scopeEntry = normalizePresetScopeEntry(result, scopeCandidate);
      if (!scopeEntry) return null;
      const entries = Array.isArray(result.entries) ? result.entries : [];
      const resolvedEntry = entries.find((entry) => getEntryKey(entry) === scopeEntry.key) || null;
      const groupControl = scopeEntry.targetType === 'groupUserControl'
        ? getParsedGroupUserControlById(result, scopeEntry.source || scopeEntry.key)
        : null;
      return {
        scopeEntry,
        entry: resolvedEntry,
        groupControl,
        targetType: scopeEntry.targetType || (groupControl ? 'groupUserControl' : 'publishedInput'),
      };
    } catch (_) {
      return null;
    }
  }

  function getPresetScopeValueFromSnapshot(snapshot, result, scopeCandidate) {
    try {
      const resolved = resolvePresetScopeTarget(result, scopeCandidate);
      if (!resolved || !result) return '';
      const entry = resolved.entry;
      const text = String(snapshot?.originalText || '');
      if (!text) {
        return resolved.targetType === 'groupUserControl'
          ? getGroupUserControlFallbackPresetValue(result, resolved.scopeEntry, resolved.groupControl)
          : getEntryFallbackPresetValue(entry);
      }
      let body = '';
      if (resolved.targetType === 'groupUserControl') {
        const inputBlock = findGroupInputBlockForScopeEntry(text, result, resolved.scopeEntry);
        if (inputBlock) {
          body = text.slice(inputBlock.open + 1, inputBlock.close);
        } else {
          return getGroupUserControlFallbackPresetValue(result, resolved.scopeEntry, resolved.groupControl);
        }
      } else {
        if (!entry) return '';
        // Switch-family Source controls act as branch selectors and must be captured
        // from the authored switch tool itself instead of the published proxy input.
        const explicitSwitchSource = getExplicitSwitchSourcePresetValue(text, result, entry);
        if (explicitSwitchSource != null) return explicitSwitchSource;
        const sourceOp = unquoteSettingValue(entry.sourceOp);
        const sourceInput = unquoteSettingValue(entry.source);
        if (!sourceOp || !sourceInput) return getPublishedPresetFallbackValue(text, result, entry);
        const bounds = locateMacroGroupBounds(text, result);
        if (!bounds) return getPublishedPresetFallbackValue(text, result, entry);
        const toolBlock = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, sourceOp);
        if (!toolBlock) return getPublishedPresetFallbackValue(text, result, entry);
        const inputsBlock = findInputsInTool(text, toolBlock.open, toolBlock.close);
        if (!inputsBlock) return getPublishedPresetFallbackValue(text, result, entry);
        const inputBlock = findInputBlockInInputs(text, inputsBlock.open, inputsBlock.close, sourceInput);
        if (!inputBlock) return getPublishedPresetFallbackValue(text, result, entry);
        body = text.slice(inputBlock.open + 1, inputBlock.close);
      }
      let value = extractControlPropLiteral(body, 'Value') || '';
      if (resolved.targetType !== 'groupUserControl' && entry && isPathEntryForDrawMode(entry)) {
        const sourceOp = unquoteSettingValue(entry?.sourceOp);
        const source = unquoteSettingValue(entry?.source);
        if (sourceOp && source) {
          const polyBody = extractPathPolylineBodyForTarget(text, result, sourceOp, source);
          if (polyBody && String(polyBody).trim()) {
            value = `Polyline {${polyBody}}`;
          }
        }
      }
      if (isTextControl(entry)) {
        const directPlain = extractStyledTextPlainTextFromInput(body);
        if (directPlain != null) value = directPlain;
        const styled = extractStyledTextBlockFromInput(body);
        if (styled) {
          const plain = extractStyledTextPlainText(styled);
          if (plain != null) value = plain;
        }
        if (/^StyledText\b/.test(String(value).trim())) {
          const plain = extractStyledTextPlainText(String(value).trim());
          if (plain != null) value = plain;
        }
        value = stripDefaultQuotes(value);
      }
      if (value == null || String(value).trim() === '') {
        return resolved.targetType === 'groupUserControl'
          ? getGroupUserControlFallbackPresetValue(result, resolved.scopeEntry, resolved.groupControl)
          : getPublishedPresetFallbackValue(text, result, entry);
      }
      return String(value);
    } catch (_) {
      return '';
    }
  }

  function captureScopeValuesFromSnapshot(snapshot, result, scopeTargets) {
    const values = {};
    if (!result || !Array.isArray(scopeTargets) || !scopeTargets.length) return values;
    scopeTargets.forEach((target) => {
      const resolved = resolvePresetScopeTarget(result, target);
      const key = String(resolved?.scopeEntry?.key || '').trim();
      if (!key) return;
      values[key] = getPresetScopeValueFromSnapshot(snapshot, result, resolved.scopeEntry);
    });
    return values;
  }

  function buildPresetScopeEntriesForKeys(result, scopeKeys) {
    try {
      if (!result || !Array.isArray(scopeKeys) || !scopeKeys.length) return [];
      const keyToEntry = new Map();
      (result.entries || []).forEach((entry) => {
        const key = getEntryKey(entry);
        if (!key || keyToEntry.has(key)) return;
        keyToEntry.set(key, entry);
      });
      return scopeKeys
        .map((key) => {
          const entry = keyToEntry.get(key);
          return entry ? buildPresetScopeEntry(result, entry) : normalizePresetScopeEntry(result, key);
        })
        .filter(Boolean);
    } catch (_) {
      return [];
    }
  }

  function makeUniquePresetName(rawName, used) {
    const base = String(rawName || '').trim() || 'Preset';
    if (!used.has(base)) {
      used.add(base);
      return base;
    }
    let n = 2;
    while (used.has(`${base} ${n}`)) n += 1;
    const next = `${base} ${n}`;
    used.add(next);
    return next;
  }

  function escapeLuaString(value) {
    return String(value || '')
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"')
      .replace(/\r/g, '\\r')
      .replace(/\n/g, '\\n');
  }

  function compactPolylinePresetLiteral(rawValue) {
    try {
      const raw = String(rawValue == null ? '' : rawValue);
      const trimmed = raw.trim();
      if (!/^polyline\b/i.test(trimmed)) return raw;
      let out = trimmed
        .replace(/\r\n?/g, '\n')
        .replace(/\s+/g, ' ')
        .replace(/\s*([{}(),=])\s*/g, '$1')
        .trim();
      if (!out) return raw;
      return out;
    } catch (_) {
      return String(rawValue == null ? '' : rawValue);
    }
  }

  function normalizePresetPayloadValue(rawValue) {
    try {
      const raw = String(rawValue == null ? '' : rawValue);
      const compact = compactPolylinePresetLiteral(raw);
      return String(compact == null ? '' : compact);
    } catch (_) {
      return String(rawValue == null ? '' : rawValue);
    }
  }

  function isPolylinePresetLiteral(rawValue) {
    try {
      const trimmed = String(rawValue == null ? '' : rawValue).trim();
      if (!trimmed) return false;
      if (/^Polyline\s*\{/i.test(trimmed)) return true;
      if (/^\{\s*(?:Closed\s*=|Points\s*=)/i.test(trimmed)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function normalizePolylinePresetLiteralForSwitch(rawValue) {
    try {
      const normalized = normalizePresetPayloadValue(rawValue);
      let trimmed = String(normalized == null ? '' : normalized).trim();
      if (!trimmed) return '';
      if (/^Polyline\s*\{/i.test(trimmed)) return trimmed;
      if (trimmed.startsWith('{')) return `Polyline ${trimmed}`;
      return '';
    } catch (_) {
      return '';
    }
  }

  function sanitizePresetPathSwitchToken(value) {
    return String(value || '').replace(/[^A-Za-z0-9_]/g, '_').replace(/^_+|_+$/g, '');
  }

  function collectKnownToolNamesForPresetSwitch(result) {
    try {
      const names = new Set();
      if (Array.isArray(result?.entries)) {
        result.entries.forEach((entry) => {
          const sourceOp = String(entry?.sourceOp || '').trim();
          if (sourceOp) names.add(sourceOp);
        });
      }
      if (result?.luaToolTypes && typeof result.luaToolTypes.forEach === 'function') {
        result.luaToolTypes.forEach((_type, toolName) => {
          const name = String(toolName || '').trim();
          if (name) names.add(name);
        });
      }
      return names;
    } catch (_) {
      return new Set();
    }
  }

  function makeUniquePresetPathSwitchToolName(baseName, takenNames) {
    try {
      const base = sanitizePresetPathSwitchToken(baseName) || 'FMR_PresetPathSwitch';
      const taken = takenNames instanceof Set ? takenNames : new Set();
      if (!taken.has(base)) {
        taken.add(base);
        return base;
      }
      let n = 2;
      while (taken.has(`${base}${n}`)) n += 1;
      const next = `${base}${n}`;
      taken.add(next);
      return next;
    } catch (_) {
      return 'FMR_PresetPathSwitch';
    }
  }

  function buildPresetPathSwitchPlan(result, payload, runtimeTargets, presetNames, defaultIndex = 0) {
    try {
      if (!Array.isArray(runtimeTargets) || !runtimeTargets.length) return [];
      if (!Array.isArray(presetNames) || !presetNames.length) return [];
      const presetObj = payload?.presets && typeof payload.presets === 'object' ? payload.presets : {};
      const baseObj = payload?.base && typeof payload.base === 'object' ? payload.base : {};
      const takenNames = collectKnownToolNamesForPresetSwitch(result);
      const plan = [];
      runtimeTargets.forEach((target) => {
        const key = String(target?.key || '').trim();
        const sourceOp = String(target?.sourceOp || '').trim();
        const sourceInput = String(target?.sourceInput || '').trim();
        if (!key || !sourceOp || !sourceInput) return;
        const pathLikeTarget = isPathControlId(sourceInput)
          || isPathControlId(target?.runtimeId)
          || isPathControlId(key);
        if (!pathLikeTarget) return;
        const byPreset = presetNames.map((name) => {
          const values = presetObj?.[name] || {};
          if (Object.prototype.hasOwnProperty.call(values, key)) return normalizePolylinePresetLiteralForSwitch(values[key]);
          if (Object.prototype.hasOwnProperty.call(baseObj, key)) return normalizePolylinePresetLiteralForSwitch(baseObj[key]);
          return '';
        });
        const firstValid = byPreset.find((item) => isPolylinePresetLiteral(item)) || '';
        if (!firstValid) return;
        const literals = byPreset.map((item) => (isPolylinePresetLiteral(item) ? item : firstValid));
        const unique = new Set(literals.map((item) => String(item || '').trim()));
        if (unique.size <= 1) return;
        const switchToolName = makeUniquePresetPathSwitchToolName(
          `${sourceOp}_${sourceInput}_PresetPathSwitch`,
          takenNames
        );
        plan.push({
          key,
          sourceOp,
          sourceInput,
          runtimeId: String(target?.runtimeId || '').trim(),
          switchToolName,
          defaultIndex: Math.max(0, Number(defaultIndex) || 0),
          literals,
        });
      });
      return plan;
    } catch (_) {
      return [];
    }
  }

  function serializePresetLuaValue(entry, rawValue) {
    const raw = normalizePresetPayloadValue(rawValue);
    const trimmed = raw.trim();
    const inputControl = normalizeInputControlValue(
      entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl || entry?.inputControl
    );
    const dataType = String(entry?.controlMeta?.dataType || entry?.controlMetaOriginal?.dataType || entry?.dataType || '')
      .replace(/"/g, '')
      .trim()
      .toLowerCase();
    const textLike = (!!inputControl && /text/i.test(inputControl))
      || String(entry?.source || entry?.id || '').toLowerCase().includes('styledtext')
      || dataType === 'text'
      || dataType === 'string'
      || dataType.includes('text')
      || isTextControl(entry);
    if (textLike) {
      return `"${escapeLuaString(raw)}"`;
    }
    if (!trimmed) return '0';
    if (/^(?:-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|true|false|nil)$/i.test(trimmed)) {
      return trimmed;
    }
    const fuIdMatch = /^FuID\s*\{\s*"((?:\\.|[^"])*)"\s*\}$/.exec(trimmed);
    if (fuIdMatch) {
      return `FuID("${escapeLuaString(unescapeSettingString(fuIdMatch[1]))}")`;
    }
    if (
      trimmed.startsWith('{') ||
      /^[A-Za-z_][A-Za-z0-9_]*\s*\{/.test(trimmed) ||
      /^[A-Za-z_][A-Za-z0-9_]*\s*\(/.test(trimmed) ||
      (trimmed.startsWith('"') && trimmed.endsWith('"'))
    ) {
      return trimmed;
    }
    return `"${escapeLuaString(raw)}"`;
  }

  function isUnsafePresetRuntimeConstructorLiteral(literal, target) {
    try {
      const trimmed = String(literal || '').trim();
      if (!trimmed) return false;
      if (/^Polyline\s*\{/i.test(trimmed)) return true;
      if (/^Path\s*\{/i.test(trimmed)) return true;
      const constructorLike = /^[A-Za-z_][A-Za-z0-9_]*\s*[\(\{]/.test(trimmed);
      if (!constructorLike) return false;
      if (/^FuID\s*\(/i.test(trimmed)) return false;
      const idHint = [
        target?.runtimeId,
        target?.key,
        target?.valueMeta?.source,
        target?.valueMeta?.id,
      ].map((v) => String(v || '').toLowerCase()).join(' ');
      if (/(^|[^a-z])(polyline|path|mask)([^a-z]|$)/i.test(idHint)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function resolvePresetRuntimeTargets(result, scopeCandidates) {
    try {
      const targets = [];
      const seen = new Set();
      const scopeEntries = normalizePresetScopeEntries(result, scopeCandidates, []);
      scopeEntries.forEach((scopeEntry) => {
        const resolved = resolvePresetScopeTarget(result, scopeEntry);
        if (!resolved) return;
        const runtimeId = resolved.targetType === 'groupUserControl'
          ? String(resolved.scopeEntry?.source || resolved.groupControl?.id || '').trim()
          : String(resolved.entry?.key || '').trim();
        const valueMeta = resolved.entry || {
          source: String(resolved.scopeEntry?.source || resolved.groupControl?.id || '').trim(),
          inputControl: String(resolved.groupControl?.inputControl || resolved.scopeEntry?.inputControl || '').trim(),
          dataType: String(resolved.groupControl?.dataType || resolved.scopeEntry?.dataType || '').trim(),
          id: String(resolved.groupControl?.id || resolved.scopeEntry?.source || '').trim(),
        };
        const key = String(resolved.scopeEntry?.key || '').trim();
        if (!key || !runtimeId || seen.has(key)) return;
        seen.add(key);
        targets.push({
          key,
          runtimeId,
          targetType: resolved.targetType,
          sourceOp: unquoteSettingValue(resolved.scopeEntry?.sourceOp || resolved.entry?.sourceOp || ''),
          sourceInput: unquoteSettingValue(resolved.scopeEntry?.source || resolved.entry?.source || ''),
          valueMeta,
        });
      });
      return targets;
    } catch (_) {
      return [];
    }
  }

  function buildPresetApplyOnChangeScript(result, payload, selectorInputId = FMR_PRESET_SELECT_CONTROL) {
    try {
      const scopeEntries = Array.isArray(payload?.scopeEntries) ? payload.scopeEntries : [];
      const presetObj = payload?.presets && typeof payload.presets === 'object' ? payload.presets : {};
      const presetNames = Object.keys(presetObj);
      if (!scopeEntries.length || !presetNames.length) return '';
      const runtimeTargets = resolvePresetRuntimeTargets(result, scopeEntries);
      if (!runtimeTargets.length) return '';
      const traceRuntime = !!FMR_PRESET_RUNTIME_TRACE;
      const pathSwitchByKey = new Map();
      (Array.isArray(payload?.pathSwitchPlan) ? payload.pathSwitchPlan : []).forEach((item) => {
        const key = String(item?.key || '').trim();
        const switchToolName = String(item?.switchToolName || '').trim();
        if (!key || !switchToolName) return;
        pathSwitchByKey.set(key, {
          switchToolName,
        });
      });
      const variationByKey = new Map();
      const skippedUnsafe = [];
      runtimeTargets.forEach((target) => {
        const seen = new Set();
        presetNames.forEach((name) => {
          const values = presetObj?.[name] || {};
          if (Object.prototype.hasOwnProperty.call(values, target.key)) {
            seen.add(String(values[target.key]));
          }
        });
        variationByKey.set(target.key, seen.size > 1);
      });
      const lines = [];
      lines.push('local __ok, __err = pcall(function()');
      lines.push(`  local preset = math.floor(tonumber(tool:GetInput("${escapeLuaString(selectorInputId || FMR_PRESET_SELECT_CONTROL)}")) or 0)`);
      presetNames.forEach((name, idx) => {
        const values = presetObj?.[name] || {};
        lines.push(`  ${idx === 0 ? 'if' : 'elseif'} preset == ${idx} then`);
        runtimeTargets.forEach((target) => {
          const pathSwitch = pathSwitchByKey.get(String(target.key || '').trim());
          if (pathSwitch?.switchToolName) {
            lines.push('    do');
            lines.push(`      local __switchName = "${escapeLuaString(pathSwitch.switchToolName)}"`);
            lines.push(`      local __switchIndex = ${idx}`);
            lines.push('      local __comp = comp or (tool and tool.Comp) or nil');
            lines.push('      local __switchTool = nil');
            lines.push('      if __comp and __comp.FindTool and __switchName ~= "" then');
            lines.push('        local __okFind, __found = pcall(function() return __comp:FindTool(__switchName) end)');
            lines.push('        if __okFind then __switchTool = __found end');
            lines.push('      end');
            lines.push('      if __switchTool then');
            lines.push('        pcall(function() __switchTool.Source = __switchIndex end)');
            lines.push('        pcall(function() __switchTool:SetInput("Source", __switchIndex) end)');
            if (traceRuntime) {
              lines.push(`        print("[Preset runtime dbg] switch=${escapeLuaString(pathSwitch.switchToolName)} index=" .. tostring(__switchIndex))`);
            }
            lines.push('      end');
            lines.push('    end');
            return;
          }
          if (!Object.prototype.hasOwnProperty.call(values, target.key)) return;
          const literal = serializePresetLuaValue(target.valueMeta, values[target.key]);
          const variesAcrossPresets = !!variationByKey.get(target.key);
          const isConstructorLike = /^[A-Za-z_][A-Za-z0-9_]*\s*[\(\{]/.test(String(literal || '').trim());
          const unsafeConstructor = isUnsafePresetRuntimeConstructorLiteral(literal, target);
          if (unsafeConstructor && !FMR_PRESET_RUNTIME_INCLUDE_CONSTRUCTORS) {
            skippedUnsafe.push(String(target.runtimeId || target.key || 'unknown'));
            return;
          }
          if (isConstructorLike && !variesAcrossPresets) return;
          lines.push(`    tool["${escapeLuaString(target.runtimeId)}"] = ${literal}`);
        });
      });
      lines.push('  end');
      lines.push('end)');
      lines.push('if not __ok and __err then');
      lines.push('  print(__err)');
      lines.push('end');
      if (skippedUnsafe.length && typeof logDiag === 'function') {
        try {
          const unique = [...new Set(skippedUnsafe)];
          logDiag(`[Preset runtime] skipped unsafe constructor assignments: ${unique.join(', ')}`);
        } catch (_) {}
      }
      return lines.join('\n').trim();
    } catch (_) {
      return '';
    }
  }

  function appendControlListEntries(body, prop, values, indent, eol) {
    try {
      let next = String(body || '');
      const items = Array.isArray(values) ? values : [];
      items.forEach((value) => {
        const escaped = escapeQuotes(String(value || '').trim());
        if (!escaped) return;
        const trimmed = next.replace(/\s+$/g, '');
        if (trimmed && !trimmed.endsWith(',')) {
          next = trimmed + ',' + next.slice(trimmed.length);
        }
        const needsBreak = next.trim().length && !next.endsWith(eol);
        next += (needsBreak ? eol : '') + indent + `{ ${prop} = "${escaped}" },`;
      });
      return next;
    } catch (_) {
      return body;
    }
  }

  function upsertPresetSelectorControl(text, result, presetNames, script, eol, defaultIndex = 0) {
    try {
      const safeDefaultIndex = Math.max(0, Number(defaultIndex) || 0);
      const scriptLiteral = script && script.trim()
        ? buildSettingLuaScriptLiteral(script, { longThreshold: 900 })
        : '';
      const lines = buildControlDefinitionLines('combo', {
        name: 'Preset',
        page: 'Controls',
        comboOptions: presetNames,
      });
      lines.push(`INP_Default = ${safeDefaultIndex},`);
      if (scriptLiteral) {
        lines.push(`INPS_ExecuteOnChange = ${scriptLiteral},`);
      }
      return upsertGroupUserControl(text, result, FMR_PRESET_SELECT_CONTROL, {
        createLines: lines,
        updateBody: (body, indent, newline) => {
          let next = body;
          next = upsertControlProp(next, 'LINKS_Name', '"Preset"', indent, newline);
          next = upsertControlProp(next, 'LINKID_DataType', '"Number"', indent, newline);
          next = upsertControlProp(next, 'INP_Integer', 'false', indent, newline);
          next = upsertControlProp(next, 'INP_Default', String(safeDefaultIndex), indent, newline);
          next = upsertControlProp(next, 'INPID_InputControl', '"ComboControl"', indent, newline);
          next = upsertControlProp(next, 'ICS_ControlPage', '"Controls"', indent, newline);
          next = removeControlListPropEntries(next, 'CCS_AddString');
          next = appendControlListEntries(next, 'CCS_AddString', presetNames, indent, newline);
          next = scriptLiteral
            ? upsertControlProp(next, 'INPS_ExecuteOnChange', scriptLiteral, indent, newline)
            : removeControlProp(next, 'INPS_ExecuteOnChange');
          return next;
        },
      }, eol || '\n');
    } catch (_) {
      return text;
    }
  }

  function syncPresetSelectorEntryToMacro(result, macroSourceOp) {
    try {
      if (!result || !macroSourceOp) return null;
      result.entries = Array.isArray(result.entries) ? result.entries : [];
      result.order = Array.isArray(result.order) ? result.order : result.entries.map((_, i) => i);
      const engine = ensurePresetEngine(result);
      const presetNames = Object.keys(engine?.presets || {}).filter((name) => String(name || '').trim());
      const hasPack = presetNames.length > 0 && Array.isArray(engine?.scopeEntries) && engine.scopeEntries.length > 0;
      const removeEntryAtIndex = (removeIdx) => {
        if (!Number.isFinite(removeIdx) || removeIdx < 0 || removeIdx >= result.entries.length) return;
        result.entries.splice(removeIdx, 1);
        result.order = (Array.isArray(result.order) ? result.order : [])
          .filter((idx) => idx !== removeIdx)
          .map((idx) => (idx > removeIdx ? idx - 1 : idx));
      };
      let idx = result.entries.findIndex((e) => e && e.source === FMR_PRESET_SELECT_CONTROL);
      if (!hasPack) {
        if (idx >= 0) removeEntryAtIndex(idx);
        return null;
      }
      const activeName = String(engine?.activePreset || presetNames[0] || '').trim();
      const defaultIndex = Math.max(0, presetNames.indexOf(activeName));
      const targetPage = 'Controls';
      const meta = {
        page: targetPage,
        inputControl: 'ComboControl',
        defaultValue: String(defaultIndex),
        choiceOptions: presetNames,
        publishTarget: 'groupUserControl',
      };
      if (idx < 0) {
        idx = ensurePublishedInResult(result, macroSourceOp, FMR_PRESET_SELECT_CONTROL, 'Preset', meta);
      } else {
        applyNodeControlMeta(result.entries[idx], meta);
      }
      if (!Number.isFinite(idx) || idx < 0 || !result.entries[idx]) return null;
      const entry = result.entries[idx];
      const key = makeUniqueKey(`${macroSourceOp}_${FMR_PRESET_SELECT_CONTROL}`);
      entry.key = key;
      entry.sourceOp = macroSourceOp;
      entry.source = FMR_PRESET_SELECT_CONTROL;
      entry.name = 'Preset';
      entry.displayName = 'Preset';
      entry.displayNameOriginal = 'Preset';
      entry.displayNameDirty = false;
      entry.page = targetPage;
      entry.raw = buildInstanceInputRaw(key, macroSourceOp, FMR_PRESET_SELECT_CONTROL, 'Preset', targetPage, null);
      entry.syntheticUserControlOnly = true;
      entry.publishTarget = 'groupUserControl';
      entry.syntheticGroupUserControl = true;
      entry.controlMeta = entry.controlMeta || {};
      entry.controlMetaOriginal = entry.controlMetaOriginal || {};
      entry.controlMeta.choiceOptions = [...presetNames];
      entry.controlMeta.inputControl = '"ComboControl"';
      entry.controlMeta.defaultValue = String(defaultIndex);
      entry.controlMetaOriginal.choiceOptions = [...presetNames];
      entry.controlMetaOriginal.inputControl = '"ComboControl"';
      entry.controlMetaOriginal.defaultValue = String(defaultIndex);
      result.order = result.order.filter((orderIdx) => orderIdx !== idx);
      const insertAt = result.order.findIndex((orderIdx) => {
        const current = result.entries[orderIdx];
        return String(current?.page || 'Controls').trim() === targetPage;
      });
      result.order.splice(insertAt >= 0 ? insertAt : 0, 0, idx);
      return idx;
    } catch (_) {
      return null;
    }
  }

  function isGroupUserControlEntry(entry) {
    try {
      return !!(entry && (entry.syntheticGroupUserControl === true || entry.publishTarget === 'groupUserControl'));
    } catch (_) {
      return false;
    }
  }

  function buildGroupUserControlEntry(result, sourceOp, source, displayName, meta) {
    const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
    const targetPage = metaPage || result.activePage || 'Controls';
    const entry = {
      key: makeUniqueKey(`${sourceOp}_${source}`),
      name: displayName || null,
      page: targetPage,
      sourceOp,
      source,
      displayName: displayName || `${sourceOp}.${source}`,
      raw: '',
      controlGroup: null,
      onChange: '',
      buttonExecute: '',
      syntheticGroupUserControl: true,
      publishTarget: 'groupUserControl',
    };
    applyNodeControlMeta(entry, meta);
    return entry;
  }

  function synchronizeAutoQuickSetLabelEntries(result, sourceText) {
    try {
      if (!result) return;
      result.entries = Array.isArray(result.entries) ? result.entries : [];
      result.order = Array.isArray(result.order) ? result.order : result.entries.map((_, i) => i);
      if (!result.entries.length) return;
      const normalizeQuickLabelId = (raw) => {
        const clean = String(raw || '').trim();
        if (!/^MM_QuickSetLabel_/i.test(clean)) return '';
        return clean;
      };
      const makeUniqueQuickLabelId = (baseId, taken) => {
        const base = normalizeQuickLabelId(baseId) || 'MM_QuickSetLabel_Node';
        if (!taken.has(base)) return base;
        let suffix = 2;
        let candidate = `${base}_${suffix}`;
        while (taken.has(candidate)) {
          suffix += 1;
          candidate = `${base}_${suffix}`;
        }
        return candidate;
      };
      try {
        // Guard against duplicate quick-label ids in state; duplicate ids collapse
        // in export maps and only one label survives.
        const allIds = new Set(
          result.entries
            .map((entry) => String(entry?.source || '').trim())
            .filter(Boolean)
        );
        const usedQuickIds = new Set();
        result.entries.forEach((entry) => {
          if (!entry) return;
          const sourceId = normalizeQuickLabelId(entry.source);
          if (!sourceId) return;
          if (!usedQuickIds.has(sourceId)) {
            usedQuickIds.add(sourceId);
            return;
          }
          allIds.delete(sourceId);
          const uniqueId = makeUniqueQuickLabelId(sourceId, allIds);
          entry.source = uniqueId;
          usedQuickIds.add(uniqueId);
          allIds.add(uniqueId);
        });
      } catch (_) {}
      try {
        if ((!Array.isArray(result.groupUserControls) || !result.groupUserControls.length) && sourceText) {
          hydrateGroupUserControlsState(sourceText, result);
        }
      } catch (_) {}
      const macroSourceOp = resolveMacroGroupSourceOp(sourceText || state.originalText || '', result);
      try {
        const groupControls = Array.isArray(result.groupUserControls) ? result.groupUserControls : [];
        groupControls.forEach((control) => {
          if (!control || !control.id) return;
          const sourceId = String(control.id || '').trim();
          if (!/^MM_QuickSetLabel_/i.test(sourceId)) return;
          let idx = result.entries.findIndex((entry) => entry && String(entry.source || '').trim() === sourceId);
          const displayName = extractQuotedProp(control.rawBody || '', 'LINKS_Name') || humanizeName(sourceId);
          const parsedInputControl = String(control.inputControl || '').trim() || 'LabelControl';
          const defaultValue = control.defaultValue != null && String(control.defaultValue).trim() !== ''
            ? String(control.defaultValue)
            : '0';
          const meta = {
            kind: 'label',
            page: control.page || 'Controls',
            inputControl: parsedInputControl,
            labelCount: Number.isFinite(control.labelCount) ? Number(control.labelCount) : 0,
            defaultValue,
            publishTarget: 'groupUserControl',
          };
          if (idx < 0) {
            idx = ensurePublishedInResult(result, macroSourceOp, sourceId, displayName, meta);
          } else {
            const entry = result.entries[idx];
            if (entry) {
              if (!entry.displayName || entry.displayName === `${entry.sourceOp || ''}.${entry.source || ''}`) {
                entry.displayName = displayName;
              }
              if (!entry.name) entry.name = displayName;
              applyNodeControlMeta(entry, meta);
            }
          }
          if (!Number.isFinite(idx) || idx < 0 || !result.entries[idx]) return;
          if (!result.order.includes(idx)) result.order.push(idx);
        });
      } catch (_) {}
      result.entries.forEach((entry) => {
        if (!entry) return;
        const sourceId = String(entry.source || '').trim();
        if (!/^MM_QuickSetLabel_/i.test(sourceId)) return;
        entry.isLabel = true;
        if (!entry.kind) entry.kind = 'Label';
        if (!entry.page) entry.page = 'Controls';
        const parsedGroupControl = getParsedGroupUserControlById(result, sourceId);
        const shouldBeGroupOwned = !!parsedGroupControl || isGroupUserControlEntry(entry);
        if (shouldBeGroupOwned) {
          entry.publishTarget = 'groupUserControl';
          entry.syntheticGroupUserControl = true;
          if (!entry.sourceOp && macroSourceOp) entry.sourceOp = macroSourceOp;
        }
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        if (!entry.controlMeta.inputControl) entry.controlMeta.inputControl = '"LabelControl"';
        if (!entry.controlMetaOriginal.inputControl) entry.controlMetaOriginal.inputControl = entry.controlMeta.inputControl;
        const labelCount = Number.isFinite(entry.labelCount)
          ? Number(entry.labelCount)
          : (Number.isFinite(entry.controlMeta.labelCount) ? Number(entry.controlMeta.labelCount) : 0);
        entry.labelCount = Math.max(0, labelCount);
        entry.controlMeta.labelCount = entry.labelCount;
        if (!Number.isFinite(entry.controlMetaOriginal.labelCount)) {
          entry.controlMetaOriginal.labelCount = entry.labelCount;
        }
        if (entry.controlMeta.defaultValue == null || entry.controlMeta.defaultValue === '') {
          entry.controlMeta.defaultValue = '0';
        }
        if (entry.controlMetaOriginal.defaultValue == null || entry.controlMetaOriginal.defaultValue === '') {
          entry.controlMetaOriginal.defaultValue = entry.controlMeta.defaultValue;
        }
      });
    } catch (_) {}
  }

  function parseTruthyGroupControlVisible(value) {
    try {
      if (value == null || value === '') return true;
      const raw = String(value).replace(/"/g, '').trim().toLowerCase();
      if (!raw) return true;
      return !(raw === 'false' || raw === '0' || raw === 'no');
    } catch (_) {
      return true;
    }
  }

  function syncVisibleGroupUserControlsIntoEntries(text, result) {
    try {
      if (!result || !Array.isArray(result.groupUserControls) || !result.groupUserControls.length) return;
      result.entries = Array.isArray(result.entries) ? result.entries : [];
      result.order = Array.isArray(result.order) ? result.order : result.entries.map((_, i) => i);
      const inserted = [];
      const macroSourceOp = resolveMacroGroupSourceOp(text, result) || String(result.macroName || result.macroNameOriginal || '').trim();
      if (!macroSourceOp) return;
      result.groupUserControls.forEach((control) => {
        if (!control || !control.id) return;
        if (control.kind === 'system') return;
        if (!String(control.inputControl || '').trim()) return;
        if (!parseTruthyGroupControlVisible(control.visible)) return;
        const inputControl = String(control.inputControl || '').trim();
        const lowerInput = inputControl.toLowerCase();
        const displayName = extractQuotedProp(control.rawBody || '', 'LINKS_Name') || humanizeName(control.id);
        const meta = {
          kind: lowerInput === 'buttoncontrol' ? 'button' : (lowerInput === 'labelcontrol' ? 'label' : ''),
          page: control.page || 'Controls',
          inputControl,
          labelCount: Number.isFinite(control.labelCount) ? Number(control.labelCount) : null,
          dataType: control.dataType || '',
          defaultValue: control.defaultValue != null ? control.defaultValue : '',
          choiceOptions: Array.isArray(control.choiceOptions) ? [...control.choiceOptions] : [],
          textLines: extractControlPropValue(control.rawBody || '', 'TEC_Lines'),
          integer: extractControlPropValue(control.rawBody || '', 'INP_Integer'),
          minAllowed: extractControlPropValue(control.rawBody || '', 'INP_MinAllowed'),
          maxAllowed: extractControlPropValue(control.rawBody || '', 'INP_MaxAllowed'),
          minScale: extractControlPropValue(control.rawBody || '', 'INP_MinScale'),
          maxScale: extractControlPropValue(control.rawBody || '', 'INP_MaxScale'),
          multiButtonShowBasic: extractControlPropValue(control.rawBody || '', 'MBTNC_ShowBasicButton'),
          publishTarget: 'groupUserControl',
        };
        const idx = ensurePublishedInResult(result, macroSourceOp, control.id, displayName, meta);
        if (!Number.isFinite(idx) || idx < 0 || !result.entries[idx]) return;
        const entry = result.entries[idx];
        entry.displayName = displayName;
        entry.displayNameOriginal = displayName;
        entry.displayNameDirty = false;
        entry.page = meta.page || entry.page || 'Controls';
        entry.syntheticGroupUserControl = true;
        entry.publishTarget = 'groupUserControl';
        entry.sortIndex = -900000 + Number(control.orderIndex || 0);
        inserted.push({ index: idx, page: entry.page || 'Controls', orderIndex: Number(control.orderIndex || 0) });
      });
      if (inserted.length) {
        const baseOrder = Array.isArray(result.order) ? [...result.order] : result.entries.map((_, i) => i);
        const moved = new Set(inserted.map((item) => item.index));
        let nextOrder = baseOrder.filter((idx) => !moved.has(idx));
        const grouped = new Map();
        inserted
          .slice()
          .sort((a, b) => {
            const pageCmp = String(a.page || '').localeCompare(String(b.page || ''));
            if (pageCmp !== 0) return pageCmp;
            if (a.orderIndex !== b.orderIndex) return a.orderIndex - b.orderIndex;
            return a.index - b.index;
          })
          .forEach((item) => {
            const page = String(item.page || 'Controls').trim() || 'Controls';
            if (!grouped.has(page)) grouped.set(page, []);
            grouped.get(page).push(item.index);
          });
        grouped.forEach((indices, page) => {
          const insertAt = nextOrder.findIndex((idx) => {
            const entry = result.entries[idx];
            const entryPage = (entry?.page && String(entry.page).trim()) ? String(entry.page).trim() : 'Controls';
            return entryPage === page;
          });
          const pos = insertAt >= 0 ? insertAt : nextOrder.length;
          nextOrder.splice(pos, 0, ...indices);
        });
        result.order = nextOrder;
      }
    } catch (_) {}
  }

  function ensurePublishedInResult(result, sourceOp, source, displayName, meta) {
    if (!result) return null;
    let idx = (result.entries || []).findIndex((e) => e && e.sourceOp === sourceOp && e.source === source);
    if (idx < 0 && source === FMR_PRESET_SELECT_CONTROL) {
      idx = (result.entries || []).findIndex((e) => e && e.source === source);
      if (idx >= 0) {
        const key = makeUniqueKey(`${sourceOp}_${source}`);
        const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
        const targetPage = metaPage || result.activePage || 'Controls';
        const entry = result.entries[idx];
        entry.key = key;
        entry.name = displayName || entry.name || null;
        entry.page = targetPage;
        entry.sourceOp = sourceOp;
        entry.source = source;
        entry.displayName = displayName || entry.displayName || `${sourceOp}.${source}`;
        entry.raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, null);
      }
    }
    if (idx < 0) {
      const entry = (meta && meta.publishTarget === 'groupUserControl')
        ? buildGroupUserControlEntry(result, sourceOp, source, displayName, meta)
        : (() => {
            const key = makeUniqueKey(`${sourceOp}_${source}`);
            const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
            const targetPage = metaPage || result.activePage || 'Controls';
            const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, null);
            return {
              key,
              name: displayName || null,
              page: targetPage,
              sourceOp,
              source,
              displayName: displayName || `${sourceOp}.${source}`,
              raw,
              controlGroup: null,
              onChange: '',
              buttonExecute: '',
            };
          })();
      result.entries.push(entry);
      idx = result.entries.length - 1;
      result.order = Array.isArray(result.order) ? result.order : result.entries.map((_, i) => i);
      result.order.push(idx);
    }
    applyNodeControlMeta(result.entries[idx], meta);
    return idx;
  }

  function sanitizeRoutingInputId(value) {
    try {
      let out = String(value || '').trim().replace(/[^A-Za-z0-9_]/g, '_');
      out = out.replace(/_+/g, '_');
      out = out.replace(/^_+/, '');
      const inputMatch = /^input_*(\d+)$/i.exec(out);
      if (inputMatch) out = `Input${Number(inputMatch[1])}`;
      if (!out) return '';
      if (!isIdentStart(out[0])) out = `_${out}`;
      return out;
    } catch (_) {
      return '';
    }
  }

  function makeUniqueRoutingInputId(base, used) {
    const taken = used instanceof Set ? used : new Set();
    const cleanBase = sanitizeRoutingInputId(base) || 'Input';
    const inputMatch = /^Input(\d+)$/i.exec(cleanBase);
    if (inputMatch) {
      let n = Number(inputMatch[1]);
      if (!Number.isFinite(n) || n <= 0) n = 1;
      let candidate = `Input${n}`;
      while (taken.has(candidate)) {
        n += 1;
        candidate = `Input${n}`;
      }
      return candidate;
    }
    if (!taken.has(cleanBase)) return cleanBase;
    let i = 2;
    while (taken.has(`${cleanBase}_${i}`)) i += 1;
    return `${cleanBase}_${i}`;
  }

  function normalizeRoutingInputRows(rows, options = {}) {
    const opts = options && typeof options === 'object' ? options : {};
    const allowEmptyMacroInput = opts.allowEmptyMacroInput === true;
    const out = [];
    (rows || []).forEach((row) => {
      if (!row) return;
      const sourceOp = String(row.sourceOp || '').trim();
      const source = String(row.source || '').trim();
      if (!sourceOp || !source) return;
      const macroInputRaw = String(row.macroInput || '').trim();
      const macroInput = allowEmptyMacroInput ? sanitizeRoutingInputId(macroInputRaw) : (sanitizeRoutingInputId(macroInputRaw) || '');
      out.push({ macroInput, sourceOp, source });
    });
    return out;
  }

  function routingRowsEqual(a, b) {
    try {
      const left = normalizeRoutingInputRows(a || [], { allowEmptyMacroInput: false });
      const right = normalizeRoutingInputRows(b || [], { allowEmptyMacroInput: false });
      if (left.length !== right.length) return false;
      for (let i = 0; i < left.length; i += 1) {
        const l = left[i];
        const r = right[i];
        if (!l || !r) return false;
        if (l.macroInput !== r.macroInput) return false;
        if (l.sourceOp !== r.sourceOp) return false;
        if (l.source !== r.source) return false;
      }
      return true;
    } catch (_) {
      return false;
    }
  }

  function ensureRoutingInputsState(result) {
    if (!result || typeof result !== 'object') return { rows: [], managed: new Set() };
    const rows = normalizeRoutingInputRows(result.routingInputs || [], { allowEmptyMacroInput: false });
    const managed = result.routingManagedNames instanceof Set
      ? new Set(result.routingManagedNames)
      : new Set(rows.map((row) => row.macroInput).filter(Boolean));
    result.routingInputs = rows;
    result.routingManagedNames = managed;
    return { rows, managed };
  }

  function setRoutingInputsState(result, rows, options = {}) {
    if (!result || typeof result !== 'object') return [];
    const opts = options && typeof options === 'object' ? options : {};
    const preserveManagedHistory = opts.preserveManagedHistory !== false;
    const normalized = normalizeRoutingInputRows(rows || [], { allowEmptyMacroInput: true });
    const prevManaged = preserveManagedHistory
      ? (result.routingManagedNames instanceof Set ? new Set(result.routingManagedNames) : new Set())
      : new Set();
    const used = new Set();
    const nextRows = [];
    normalized.forEach((row) => {
      const fallback = sanitizeRoutingInputId(`${row.sourceOp}_${row.source}`) || 'Input';
      const macroInput = makeUniqueRoutingInputId(row.macroInput || fallback, used);
      used.add(macroInput);
      nextRows.push({
        macroInput,
        sourceOp: String(row.sourceOp || '').trim(),
        source: String(row.source || '').trim(),
      });
    });
    const managed = preserveManagedHistory ? prevManaged : new Set();
    nextRows.forEach((row) => managed.add(row.macroInput));
    result.routingInputs = nextRows;
    result.routingManagedNames = managed;
    return nextRows;
  }

  function getPrimaryGroupInputsBlock(text, result) {
    try {
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return null;
      let blocks = findGroupInputsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!Array.isArray(blocks) || !blocks.length) {
        const fallback = findPrimaryInputsBlockFallback(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
        blocks = fallback ? [fallback] : [];
      }
      return Array.isArray(blocks) && blocks.length ? blocks[0] : null;
    } catch (_) {
      return null;
    }
  }

  function parseRoutingInputsFromText(text, result) {
    try {
      if (!text || !result) return [];
      const block = getPrimaryGroupInputsBlock(text, result);
      if (!block) return [];
      const entries = parseOrderedBlockEntries(text, block.openIndex, block.closeIndex);
      const rows = [];
      entries.forEach((entry) => {
        if (!entry || !entry.name) return;
        const type = getOrderedBlockEntryType(text, entry);
        if (type !== 'Input' && type !== 'InstanceInput') return;
        const routing = extractRoutingInputFromEntry(text, entry);
        if (!routing) return;
        rows.push({
          macroInput: sanitizeRoutingInputId(entry.name),
          sourceOp: String(routing.sourceOp).trim(),
          source: String(routing.source).trim(),
        });
      });
      return rows;
    } catch (_) {
      return [];
    }
  }

  function hydrateRoutingInputsState(text, result) {
    try {
      if (!result) return;
      const rows = parseRoutingInputsFromText(text, result);
      setRoutingInputsState(result, rows, { preserveManagedHistory: false });
    } catch (_) {
      ensureRoutingInputsState(result);
    }
  }

  function isLikelyRoutingTargetControlId(idRaw) {
    try {
      const id = String(idRaw || '').trim();
      if (!id) return false;
      const key = id.toLowerCase();
      if (
        key === 'input' ||
        key === 'background' ||
        key === 'foreground' ||
        key === 'effectmask' ||
        key === 'mask' ||
        key === 'sceneinput'
      ) return true;
      if (key === 'source') return false;
      if (/^input\d+$/.test(key)) return true;
      if (/^background\d+$/.test(key)) return true;
      if (/^foreground\d+$/.test(key)) return true;
      if (/^mask\d+$/.test(key)) return true;
      if (/^image\d+$/.test(key)) return true;
      if (/^materialinput\d*$/.test(key)) return true;
      if (/^layer\d+\.(?:foreground|center|mask)$/.test(key)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function isMaskRoutingTargetControlId(idRaw) {
    try {
      const key = String(idRaw || '').trim().toLowerCase();
      if (!key) return false;
      if (key === 'mask' || key === 'effectmask') return true;
      if (/^mask\d*$/.test(key)) return true;
      if (/^effectmask\d*$/.test(key)) return true;
      if (/^layer\d+\.mask$/.test(key)) return true;
      if (/\.mask$/.test(key)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function flattenRoutingNodeControls(node, options = {}) {
    const includeConnected = options && options.includeConnected === true;
    const includeMaskInputs = options && options.includeMaskInputs === true;
    const out = [];
    const seen = new Set();
    const pushControl = (idRaw, nameRaw, meta = null) => {
      const id = String(idRaw || '').trim();
      if (!id) return;
      if (/^group:controlGroup:/i.test(id)) return;
      if (!isLikelyRoutingTargetControlId(id)) return;
      const isConnected = !!meta?.isConnected;
      const isMaskInput = isMaskRoutingTargetControlId(id);
      if (!includeConnected && isConnected) return;
      if (!includeMaskInputs && isMaskInput) return;
      if (seen.has(id)) return;
      seen.add(id);
      out.push({
        id,
        name: String(nameRaw || '').trim() || humanizeName(id),
        isConnected,
        isMaskInput,
        inputSourceOp: String(meta?.inputSourceOp || '').trim() || null,
        inputSource: String(meta?.inputSource || '').trim() || null,
      });
    };
    (node?.controls || []).forEach((control) => {
      if (!control) return;
      if (Array.isArray(control.channels) && control.channels.length) {
        control.channels.forEach((channel) => {
          if (!channel) return;
          pushControl(channel.id, channel.name, channel);
        });
        return;
      }
      pushControl(control.id, control.name, control);
    });
    return out;
  }

  function buildRoutingNodeCandidates(result, options = {}) {
    const nodes = Array.isArray(result?.nodes) ? result.nodes : [];
    const out = [];
    nodes.forEach((node) => {
      if (!node || node.isMacroRoot) return;
      if (node.isModifier) return;
      if (node.external) return;
      const sourceOp = String(node.name || '').trim();
      if (!sourceOp) return;
      const controls = flattenRoutingNodeControls(node, options);
      if (!controls.length) return;
      out.push({
        sourceOp,
        label: `${sourceOp} (${String(node.type || 'Node').trim() || 'Node'})`,
        controls,
      });
    });
    return out;
  }

  function buildManualRoutingRow(existingRows = []) {
    const usedNumbers = new Set();
    (existingRows || []).forEach((row) => {
      const label = String(row?.macroInput || '').trim();
      const match = /^Input\s+(\d+)$/i.exec(label);
      if (match) usedNumbers.add(Number(match[1]));
    });
    let next = 1;
    while (usedNumbers.has(next)) next += 1;
    return {
      macroInput: `Input ${next}`,
      sourceOp: '',
      source: '',
    };
  }

  function isDefaultRoutingMacroInputLabel(value) {
    try {
      return /^input\s*\d+$/i.test(String(value || '').trim());
    } catch (_) {
      return false;
    }
  }

  function openRoutingIOModal() {
    if (!state.parseResult) {
      error('Load a macro before opening Routing.');
      return;
    }
    try { nodesPane?.parseAndRenderNodes?.(); } catch (_) {}
    const result = state.parseResult;
    const active = getActiveDocument();
    if (active) storeDocumentSnapshot(active);
    const allCandidates = buildRoutingNodeCandidates(result, { includeConnected: true });
    if (!allCandidates.length) {
      error('No node inputs found for routing.');
      return;
    }
    let showConnected = false;
    let showMaskInputs = false;
    let candidateByNode = new Map();
    const getVisibleCandidates = () => buildRoutingNodeCandidates(result, {
      includeConnected: showConnected,
      includeMaskInputs: showMaskInputs,
    });
    const seedState = ensureRoutingInputsState(result);
    let rows = seedState.rows.length
      ? seedState.rows.map((row) => ({ ...row }))
      : [];
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal csv-modal';
    const modal = document.createElement('form');
    modal.className = 'add-control-form routing-io-modal';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');

    const header = document.createElement('header');
    const headerText = document.createElement('div');
    const eyebrow = document.createElement('div');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Routing';
    const title = document.createElement('h3');
    title.textContent = 'Macro IO Routing';
    headerText.appendChild(eyebrow);
    headerText.appendChild(title);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.innerHTML = '&times;';
    closeBtn.setAttribute('aria-label', 'Close');
    header.appendChild(headerText);
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'form-body';
    const hint = document.createElement('p');
    hint.className = 'form-hint';
    hint.textContent = 'Choose which internal node inputs should be wired as macro-level Input routes.';
    body.appendChild(hint);

    const filterRow = document.createElement('div');
    filterRow.className = 'routing-io-filters';
    const buildSwitch = (labelText, checked) => {
      const wrap = document.createElement('label');
      wrap.className = 'pane-switch';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.checked = !!checked;
      const slider = document.createElement('span');
      slider.className = 'slider';
      slider.setAttribute('aria-hidden', 'true');
      const label = document.createElement('span');
      label.className = 'label';
      label.textContent = labelText;
      wrap.appendChild(input);
      wrap.appendChild(slider);
      wrap.appendChild(label);
      return { wrap, input };
    };
    const connectedSwitch = buildSwitch('Show connected inputs', showConnected);
    const maskSwitch = buildSwitch('Show mask inputs', showMaskInputs);
    filterRow.appendChild(connectedSwitch.wrap);
    filterRow.appendChild(maskSwitch.wrap);
    body.appendChild(filterRow);

    const suggestedWrap = document.createElement('div');
    suggestedWrap.className = 'routing-io-suggested-wrap';
    const suggestedHead = document.createElement('div');
    suggestedHead.className = 'routing-io-suggested-head';
    suggestedHead.textContent = 'Suggested Inputs';
    const suggestedList = document.createElement('div');
    suggestedList.className = 'routing-io-suggested-list';
    suggestedWrap.appendChild(suggestedHead);
    suggestedWrap.appendChild(suggestedList);
    body.appendChild(suggestedWrap);

    const manualHead = document.createElement('div');
    manualHead.className = 'routing-io-manual-head';
    const manualHeadTitle = document.createElement('span');
    manualHeadTitle.textContent = 'Manual Inputs';
    const manualHeadActions = document.createElement('div');
    manualHeadActions.className = 'routing-io-manual-actions';
    const flipABBtn = document.createElement('button');
    flipABBtn.type = 'button';
    flipABBtn.className = 'routing-io-flip';
    flipABBtn.textContent = 'Flip A/B';
    manualHeadActions.appendChild(flipABBtn);
    manualHead.appendChild(manualHeadTitle);
    manualHead.appendChild(manualHeadActions);
    body.appendChild(manualHead);

    const tableWrap = document.createElement('div');
    tableWrap.className = 'routing-io-table-wrap detail-combo-options-input';
    const table = document.createElement('div');
    table.className = 'routing-io-table';
    tableWrap.appendChild(table);
    body.appendChild(tableWrap);

    const bodyActions = document.createElement('div');
    bodyActions.className = 'detail-actions';
    const addRowBtn = document.createElement('button');
    addRowBtn.type = 'button';
    addRowBtn.textContent = 'Add Input';
    bodyActions.appendChild(addRowBtn);
    body.appendChild(bodyActions);

    const actions = document.createElement('footer');
    actions.className = 'modal-actions';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.className = 'primary';
    saveBtn.textContent = 'Save Routing';
    actions.appendChild(cancelBtn);
    actions.appendChild(saveBtn);

    const routeKey = (sourceOp, source) => `${String(sourceOp || '').trim()}::${String(source || '').trim()}`;
    const findRouteRowIndex = (sourceOp, source) => rows.findIndex((row) => routeKey(row?.sourceOp, row?.source) === routeKey(sourceOp, source));
    const maybeRenumberDefaultRows = () => {
      try {
        if (!Array.isArray(rows) || !rows.length) return;
        if (!rows.every((row) => isDefaultRoutingMacroInputLabel(row?.macroInput))) return;
        rows = rows.map((row, idx) => ({
          ...row,
          macroInput: `Input ${idx + 1}`,
        }));
      } catch (_) {}
    };
    const refreshManualActions = () => {
      flipABBtn.hidden = rows.length !== 2;
      flipABBtn.disabled = rows.length !== 2;
    };
    const renderSuggested = (visibleCandidates) => {
      suggestedList.innerHTML = '';
      const items = [];
      visibleCandidates.forEach((candidate) => {
        const sourceOp = String(candidate?.sourceOp || '').trim();
        if (!sourceOp) return;
        const controls = Array.isArray(candidate?.controls) ? candidate.controls : [];
        controls.forEach((control) => {
          const source = String(control?.id || '').trim();
          if (!source) return;
          items.push({
            sourceOp,
            source,
            nodeLabel: String(candidate?.label || sourceOp).trim(),
            controlName: String(control?.name || source).trim() || source,
            isConnected: !!control?.isConnected,
            isMaskInput: !!control?.isMaskInput,
          });
        });
      });
      if (!items.length) {
        suggestedHead.textContent = 'Suggested Inputs';
        const empty = document.createElement('div');
        empty.className = 'routing-io-empty';
        empty.textContent = 'No suggested inputs with current filters.';
        suggestedList.appendChild(empty);
        return;
      }
      items.sort((a, b) => {
        const opCmp = String(a.sourceOp || '').localeCompare(String(b.sourceOp || ''));
        if (opCmp !== 0) return opCmp;
        return String(a.source || '').localeCompare(String(b.source || ''));
      });
      const selectedCount = items.filter((item) => findRouteRowIndex(item.sourceOp, item.source) >= 0).length;
      suggestedHead.textContent = `Suggested Inputs (${selectedCount}/${items.length} selected)`;
      items.forEach((item) => {
        const row = document.createElement('label');
        row.className = 'routing-io-suggested-item';
        const toggle = document.createElement('input');
        toggle.type = 'checkbox';
        toggle.checked = findRouteRowIndex(item.sourceOp, item.source) >= 0;
        row.classList.toggle('is-selected', !!toggle.checked);
        const text = document.createElement('div');
        text.className = 'routing-io-suggested-text';
        const title = document.createElement('div');
        title.className = 'routing-io-suggested-title';
        const tags = [];
        if (item.isConnected) tags.push('connected');
        if (item.isMaskInput) tags.push('mask');
        title.textContent = tags.length
          ? `${item.controlName} [${tags.join(', ')}]`
          : item.controlName;
        const meta = document.createElement('div');
        meta.className = 'routing-io-suggested-meta';
        meta.textContent = `${item.sourceOp}.${item.source}`;
        text.appendChild(title);
        text.appendChild(meta);
        row.appendChild(toggle);
        row.appendChild(text);
        toggle.addEventListener('change', () => {
          const idx = findRouteRowIndex(item.sourceOp, item.source);
          if (toggle.checked) {
            if (idx < 0) {
              const seed = buildManualRoutingRow(rows);
              rows.push({
                macroInput: String(seed?.macroInput || '').trim() || `Input ${rows.length + 1}`,
                sourceOp: item.sourceOp,
                source: item.source,
              });
            }
          } else if (idx >= 0) {
            rows.splice(idx, 1);
          }
          maybeRenumberDefaultRows();
          renderRows();
        });
        suggestedList.appendChild(row);
      });
    };

    const moveRoutingRow = (fromIdx, toIdx) => {
      try {
        if (!Array.isArray(rows) || rows.length < 2) return;
        if (!Number.isFinite(fromIdx) || !Number.isFinite(toIdx)) return;
        if (fromIdx < 0 || fromIdx >= rows.length) return;
        if (toIdx < 0 || toIdx >= rows.length || toIdx === fromIdx) return;
        const [moved] = rows.splice(fromIdx, 1);
        rows.splice(toIdx, 0, moved);
        maybeRenumberDefaultRows();
        renderRows();
      } catch (_) {}
    };

    const renderRows = () => {
      const visibleCandidates = getVisibleCandidates();
      candidateByNode = new Map(visibleCandidates.map((item) => [item.sourceOp, item]));
      renderSuggested(visibleCandidates);
      table.innerHTML = '';
      const addHead = (text) => {
        const cell = document.createElement('div');
        cell.className = 'routing-io-head';
        cell.textContent = text;
        table.appendChild(cell);
      };
      addHead('Order');
      addHead('Macro Input');
      addHead('Node');
      addHead('Target Input');
      addHead('');
      if (!rows.length) {
        refreshManualActions();
        const empty = document.createElement('div');
        empty.className = 'routing-io-empty';
        if (visibleCandidates.length) {
          empty.textContent = 'No inputs defined. Click Add Input to create one.';
        } else if (!showConnected && !showMaskInputs) {
          empty.textContent = 'No open non-mask inputs found. Enable Show connected inputs or Show mask inputs.';
        } else if (!showConnected && showMaskInputs) {
          empty.textContent = 'No open inputs found. Enable Show connected inputs to include occupied sockets.';
        } else if (showConnected && !showMaskInputs) {
          empty.textContent = 'No non-mask inputs found. Enable Show mask inputs.';
        } else {
          empty.textContent = 'No inputs found for routing.';
        }
        empty.style.gridColumn = '1 / -1';
        table.appendChild(empty);
        return;
      }
      refreshManualActions();
      rows.forEach((row, idx) => {
        const handleCell = document.createElement('div');
        handleCell.className = 'routing-io-cell routing-io-cell-handle';
        const orderBtns = document.createElement('div');
        orderBtns.className = 'routing-io-order-buttons';
        const moveUpBtn = document.createElement('button');
        moveUpBtn.type = 'button';
        moveUpBtn.className = 'routing-io-order-btn';
        moveUpBtn.title = 'Move up';
        moveUpBtn.textContent = '^';
        moveUpBtn.disabled = idx === 0;
        moveUpBtn.addEventListener('click', () => moveRoutingRow(idx, idx - 1));
        const moveDownBtn = document.createElement('button');
        moveDownBtn.type = 'button';
        moveDownBtn.className = 'routing-io-order-btn';
        moveDownBtn.title = 'Move down';
        moveDownBtn.textContent = 'v';
        moveDownBtn.disabled = idx >= rows.length - 1;
        moveDownBtn.addEventListener('click', () => moveRoutingRow(idx, idx + 1));
        orderBtns.appendChild(moveUpBtn);
        orderBtns.appendChild(moveDownBtn);
        handleCell.appendChild(orderBtns);
        table.appendChild(handleCell);

        const macroCell = document.createElement('div');
        macroCell.className = 'routing-io-cell';
        const macroInput = document.createElement('input');
        macroInput.type = 'text';
        macroInput.value = row.macroInput || '';
        macroInput.placeholder = 'InputName';
        macroInput.addEventListener('input', () => {
          rows[idx].macroInput = macroInput.value;
        });
        macroCell.appendChild(macroInput);
        table.appendChild(macroCell);

        const nodeCell = document.createElement('div');
        nodeCell.className = 'routing-io-cell';
        const nodeSelect = document.createElement('select');
        const currentNode = String(row.sourceOp || '').trim();
        const hasCurrentNode = candidateByNode.has(currentNode);
        const chooseNode = document.createElement('option');
        chooseNode.value = '';
        chooseNode.textContent = 'Select node...';
        nodeSelect.appendChild(chooseNode);
        if (!hasCurrentNode && currentNode) {
          const missing = document.createElement('option');
          missing.value = currentNode;
          missing.textContent = `${currentNode} (missing)`;
          nodeSelect.appendChild(missing);
        }
        visibleCandidates.forEach((candidate) => {
          const opt = document.createElement('option');
          opt.value = candidate.sourceOp;
          opt.textContent = candidate.label;
          nodeSelect.appendChild(opt);
        });
        nodeSelect.value = hasCurrentNode ? currentNode : (currentNode || '');
        nodeCell.appendChild(nodeSelect);
        table.appendChild(nodeCell);

        const portCell = document.createElement('div');
        portCell.className = 'routing-io-cell';
        const portSelect = document.createElement('select');
        const syncPortOptions = () => {
          const selectedNode = String(nodeSelect.value || '').trim();
          const controls = candidateByNode.get(selectedNode)?.controls || [];
          portSelect.innerHTML = '';
          const choosePort = document.createElement('option');
          choosePort.value = '';
          choosePort.textContent = selectedNode ? 'Select input...' : 'Select node first...';
          portSelect.appendChild(choosePort);
          const currentPort = String(rows[idx].source || '').trim();
          const hasCurrentPort = controls.some((control) => String(control.id || '').trim() === currentPort);
          if (!hasCurrentPort && currentPort) {
            const missing = document.createElement('option');
            missing.value = currentPort;
            missing.textContent = `${currentPort} (missing)`;
            portSelect.appendChild(missing);
          }
          controls.forEach((control) => {
            const opt = document.createElement('option');
            opt.value = control.id;
            const tags = [];
            if (control.isConnected) tags.push('connected');
            if (control.isMaskInput) tags.push('mask');
            const suffix = tags.length ? ` [${tags.join(', ')}]` : '';
            opt.textContent = `${control.name} (${control.id})${suffix}`;
            portSelect.appendChild(opt);
          });
          const fallback = hasCurrentPort ? currentPort : '';
          portSelect.value = fallback;
          rows[idx].sourceOp = selectedNode;
          rows[idx].source = fallback;
        };
        syncPortOptions();
        nodeSelect.addEventListener('change', syncPortOptions);
        portSelect.addEventListener('change', () => {
          rows[idx].sourceOp = String(nodeSelect.value || '').trim();
          rows[idx].source = String(portSelect.value || '').trim();
        });
        portCell.appendChild(portSelect);
        table.appendChild(portCell);

        const removeCell = document.createElement('div');
        removeCell.className = 'routing-io-cell';
        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'routing-io-remove';
        removeBtn.textContent = 'Remove';
        removeBtn.addEventListener('click', () => {
          rows.splice(idx, 1);
          maybeRenumberDefaultRows();
          renderRows();
        });
        removeCell.appendChild(removeBtn);
        table.appendChild(removeCell);
        if (idx < rows.length - 1) {
          const divider = document.createElement('div');
          divider.className = 'routing-io-row-divider';
          table.appendChild(divider);
        }
      });
    };

    addRowBtn.addEventListener('click', () => {
      rows.push(buildManualRoutingRow(rows));
      maybeRenumberDefaultRows();
      renderRows();
    });
    flipABBtn.addEventListener('click', () => {
      if (rows.length !== 2) return;
      const a = rows[0];
      const b = rows[1];
      const aOp = String(a?.sourceOp || '').trim();
      const aSrc = String(a?.source || '').trim();
      rows[0].sourceOp = String(b?.sourceOp || '').trim();
      rows[0].source = String(b?.source || '').trim();
      rows[1].sourceOp = aOp;
      rows[1].source = aSrc;
      renderRows();
    });
    connectedSwitch.input.addEventListener('change', () => {
      showConnected = !!connectedSwitch.input.checked;
      renderRows();
    });
    maskSwitch.input.addEventListener('change', () => {
      showMaskInputs = !!maskSwitch.input.checked;
      renderRows();
    });

    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    let overlayPointerDown = false;
    overlay.addEventListener('mousedown', (ev) => {
      overlayPointerDown = ev.target === overlay;
    });
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay && overlayPointerDown) close();
      overlayPointerDown = false;
    });
    closeBtn.addEventListener('click', close);
    cancelBtn.addEventListener('click', close);
    modal.addEventListener('submit', (ev) => ev.preventDefault());

    saveBtn.addEventListener('click', () => {
      const nextRowsRaw = normalizeRoutingInputRows(rows, { allowEmptyMacroInput: true });
      const usedNames = new Set();
      const nextRows = [];
      nextRowsRaw.forEach((row) => {
        const base = row.macroInput || sanitizeRoutingInputId(`${row.sourceOp}_${row.source}`) || 'Input';
        const macroInput = makeUniqueRoutingInputId(base, usedNames);
        usedNames.add(macroInput);
        nextRows.push({
          macroInput,
          sourceOp: String(row.sourceOp || '').trim(),
          source: String(row.source || '').trim(),
        });
      });
      const prevRows = ensureRoutingInputsState(result).rows;
      if (routingRowsEqual(prevRows, nextRows)) {
        close();
        return;
      }
      pushHistory('routing io');
      setRoutingInputsState(result, nextRows, { preserveManagedHistory: true });
      markContentDirty();
      info(`Routing updated (${nextRows.length} input${nextRows.length === 1 ? '' : 's'}).`);
      close();
    });

    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    renderRows();
  }

  function openPresetEngineModal() {
    const active = getActiveDocument();
    if (active) storeDocumentSnapshot(active);
    const allDocs = documents.filter((doc) => doc && !doc.isCsvBatch && doc.snapshot && doc.snapshot.parseResult);
    const docHasPresetPack = (doc) => {
      try {
        const result = doc?.snapshot?.parseResult;
        const engine = ensurePresetEngine(result);
        const presetCount = Object.keys(engine?.presets || {}).filter((name) => String(name || '').trim()).length;
        const scopeCount = Array.isArray(engine?.scopeEntries) ? engine.scopeEntries.length : 0;
        return presetCount > 0 && scopeCount > 0;
      } catch (_) {
        return false;
      }
    };
    const selectedDocs = allDocs.filter((doc) => !!doc.selected);
    const candidateDocs = selectedDocs.length >= 2 ? selectedDocs : allDocs;
    if (!candidateDocs.length) {
      error('Load a macro before opening the presets engine.');
      return;
    }
    if (candidateDocs.length < 2 && !candidateDocs.some(docHasPresetPack)) {
      error('Select at least two macro tabs (root + variants) to build a preset pack.');
      return;
    }
    let rootDocId = String((candidateDocs.find((doc) => doc.id === activeDocId) || candidateDocs[0]).id || '');
    const overlay = document.createElement('div');
    overlay.className = 'add-control-modal csv-modal';
    const modal = document.createElement('form');
    modal.className = 'add-control-form';
    modal.classList.add('preset-engine-modal');
    const header = document.createElement('header');
    const headerText = document.createElement('div');
    const eyebrow = document.createElement('div');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'Presets';
    const title = document.createElement('h3');
    title.textContent = 'Presets Engine';
    headerText.appendChild(eyebrow);
    headerText.appendChild(title);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.innerHTML = '&times;';
    header.appendChild(headerText);
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'form-body';
    body.classList.add('preset-engine-body');

    const leftCol = document.createElement('div');
    leftCol.className = 'preset-engine-column preset-engine-column-main';

    const rightCol = document.createElement('div');
    rightCol.className = 'preset-engine-column preset-engine-column-scope';

    const status = document.createElement('p');
    status.className = 'form-hint';
    status.textContent = candidateDocs.length >= 2
      ? 'Choose a root tab, choose variant tabs, select value scope, then build.'
      : 'Review or remove the existing preset pack on this macro. Load more tabs to build a new one.';
    body.appendChild(status);

    const rootLabel = document.createElement('label');
    rootLabel.textContent = 'Root / Base tab';
    const rootSelect = document.createElement('select');
    candidateDocs.forEach((doc) => {
      const opt = document.createElement('option');
      opt.value = doc.id;
      opt.textContent = getDocDisplayName(doc, candidateDocs) || doc.id;
      rootSelect.appendChild(opt);
    });
    rootSelect.value = rootDocId;
    rootLabel.appendChild(rootSelect);
    leftCol.appendChild(rootLabel);

    const baseNameLabel = document.createElement('label');
    baseNameLabel.className = 'preset-engine-field';
    baseNameLabel.textContent = 'Base preset name';
    const baseNameInput = document.createElement('input');
    baseNameInput.type = 'text';
    baseNameInput.className = 'preset-engine-variant-name';
    baseNameInput.placeholder = 'Default';
    baseNameLabel.appendChild(baseNameInput);

    const builtState = document.createElement('section');
    builtState.className = 'preset-engine-state';
    builtState.hidden = true;
    const builtStateHeader = document.createElement('div');
    builtStateHeader.className = 'preset-engine-state-header';
    const builtStateTitle = document.createElement('div');
    builtStateTitle.className = 'preset-engine-state-title';
    builtStateTitle.textContent = 'Current root preset pack';
    const builtStateToggle = document.createElement('button');
    builtStateToggle.type = 'button';
    builtStateToggle.className = 'preset-engine-state-toggle';
    builtStateToggle.textContent = 'Show details';
    const builtStateMeta = document.createElement('div');
    builtStateMeta.className = 'preset-engine-state-meta';
    const builtStatePresets = document.createElement('div');
    builtStatePresets.className = 'preset-engine-state-presets';
    const builtStateScope = document.createElement('div');
    builtStateScope.className = 'preset-engine-state-scope';
    const builtStateDetails = document.createElement('div');
    builtStateDetails.className = 'preset-engine-state-details';
    const builtStateActions = document.createElement('div');
    builtStateActions.className = 'preset-engine-state-actions';
    const clearBuiltBtn = document.createElement('button');
    clearBuiltBtn.type = 'button';
    clearBuiltBtn.textContent = 'Remove Preset Pack';
    let builtStateExpanded = false;
    builtStateHeader.appendChild(builtStateTitle);
    builtStateHeader.appendChild(builtStateToggle);
    builtState.appendChild(builtStateHeader);
    builtState.appendChild(builtStateMeta);
    builtStateDetails.appendChild(builtStatePresets);
    builtStateDetails.appendChild(builtStateScope);
    builtState.appendChild(builtStateDetails);
    builtStateActions.appendChild(clearBuiltBtn);
    builtState.appendChild(builtStateActions);

    const variantLabel = document.createElement('label');
    variantLabel.className = 'preset-engine-field';
    variantLabel.textContent = 'Variant tabs (presets)';
    const variantWrap = document.createElement('div');
    variantWrap.className = 'detail-combo-options-input preset-engine-list preset-engine-variant-list';
    variantLabel.appendChild(variantWrap);
    leftCol.appendChild(variantLabel);
    leftCol.appendChild(builtState);

    const scopeLabel = document.createElement('label');
    scopeLabel.className = 'preset-engine-field';
    scopeLabel.textContent = 'Value scope';
    const scopeWrap = document.createElement('div');
    scopeWrap.className = 'detail-combo-options-input preset-engine-list preset-engine-scope-list';
    scopeLabel.appendChild(scopeWrap);
    rightCol.appendChild(scopeLabel);

    let currentRoot = null;
    let rootEligible = [];
    let scopeSelection = new Set();
    let variantRows = [];

    const setBuiltStateExpanded = (expanded) => {
      builtStateExpanded = !!expanded;
      builtState.classList.toggle('expanded', builtStateExpanded);
      builtStateToggle.textContent = builtStateExpanded ? 'Hide details' : 'Show details';
    };

    const renderBuiltState = () => {
      builtStateMeta.innerHTML = '';
      builtStatePresets.innerHTML = '';
      builtStateScope.innerHTML = '';
      const rootId = String(rootSelect.value || '');
      const rootDoc = candidateDocs.find((doc) => doc.id === rootId) || null;
      const rootResult = rootDoc?.snapshot?.parseResult || null;
      const engine = ensurePresetEngine(rootResult);
      const presetNames = Object.keys(engine?.presets || {});
      const scopeEntries = Array.isArray(engine?.scopeEntries) ? engine.scopeEntries : [];
      const targetCounts = new Map();
      scopeEntries.forEach((item) => {
        const key = String(item?.targetType || 'unknown').trim() || 'unknown';
        targetCounts.set(key, (targetCounts.get(key) || 0) + 1);
      });
      const addMeta = (label, value) => {
        const chip = document.createElement('div');
        chip.className = 'preset-engine-state-chip';
        chip.textContent = `${label}: ${value}`;
        builtStateMeta.appendChild(chip);
      };
      addMeta('Root', getDocDisplayName(rootDoc, candidateDocs) || rootId || 'Unknown');
      addMeta('Presets', presetNames.length);
      addMeta('Scope', scopeEntries.length);
      if (engine?.activePreset) addMeta('Default', engine.activePreset);
      if (targetCounts.size) {
        targetCounts.forEach((count, key) => {
          const label = key === 'groupUserControl' ? 'Macro-root' : (key === 'publishedInput' ? 'Published' : key);
          addMeta(label, count);
        });
      }
      if (!presetNames.length) {
        builtState.hidden = true;
        clearBuiltBtn.disabled = true;
        builtState.classList.add('is-empty');
        setBuiltStateExpanded(false);
        const empty = document.createElement('div');
        empty.className = 'preset-engine-state-empty';
        empty.textContent = 'No preset pack built on this root yet.';
        builtStatePresets.appendChild(empty);
        return;
      }
      builtState.hidden = false;
      builtState.classList.remove('is-empty');
      clearBuiltBtn.disabled = false;
      setBuiltStateExpanded(false);
      presetNames.forEach((name) => {
        const pill = document.createElement('span');
        pill.className = 'preset-engine-preset-pill';
        if (name === engine.activePreset) pill.classList.add('active');
        pill.textContent = name;
        builtStatePresets.appendChild(pill);
      });
      scopeEntries.forEach((item) => {
        const row = document.createElement('div');
        row.className = 'preset-engine-state-scope-row';
        const name = document.createElement('span');
        name.className = 'preset-engine-state-scope-name';
        name.textContent = item.displayName || item.key;
        const meta = document.createElement('span');
        meta.className = 'preset-engine-state-scope-meta';
        const target = item.targetType === 'groupUserControl' ? 'Macro-root' : (item.targetType === 'publishedInput' ? 'Published' : String(item.targetType || 'unknown'));
        const page = String(item.page || 'Controls').trim() || 'Controls';
        meta.textContent = `${target} • ${page}`;
        row.appendChild(name);
        row.appendChild(meta);
        builtStateScope.appendChild(row);
      });
    };

    const renderVariants = () => {
      variantWrap.innerHTML = '';
      variantRows = [];
      const rootId = String(rootSelect.value || '');
      const rootDoc = candidateDocs.find((doc) => doc.id === rootId) || null;

      const rootRow = document.createElement('div');
      rootRow.className = 'preset-engine-variant-row';
      const rootFlag = document.createElement('input');
      rootFlag.type = 'checkbox';
      rootFlag.checked = true;
      rootFlag.disabled = true;
      const rootSource = document.createElement('span');
      rootSource.className = 'preset-engine-variant-source';
      rootSource.textContent = `Root / Base: ${getDocDisplayName(rootDoc, candidateDocs) || rootId}`;
      rootSource.title = rootSource.textContent;
      rootRow.appendChild(rootFlag);
      rootRow.appendChild(rootSource);
      rootRow.appendChild(baseNameInput);
      variantWrap.appendChild(rootRow);

      candidateDocs.forEach((doc) => {
        if (doc.id === rootId) return;
        const row = document.createElement('div');
        row.className = 'preset-engine-variant-row';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = true;
        const source = document.createElement('span');
        source.className = 'preset-engine-variant-source';
        source.textContent = getDocDisplayName(doc, candidateDocs) || doc.id;
        source.title = source.textContent;
        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.className = 'preset-engine-variant-name';
        nameInput.value = getDocDisplayName(doc, candidateDocs) || 'Preset';
        nameInput.placeholder = 'Preset name';
        row.appendChild(cb);
        row.appendChild(source);
        row.appendChild(nameInput);
        variantWrap.appendChild(row);
        variantRows.push({ doc, cb, nameInput });
      });
      if (!variantRows.length) {
        const empty = document.createElement('div');
        empty.className = 'preset-engine-state-empty';
        empty.textContent = 'No variant tabs loaded. Load additional tabs to build a new preset pack.';
        variantWrap.appendChild(empty);
      }
      buildBtn.disabled = variantRows.length === 0;
    };

    const renderScope = () => {
      scopeWrap.innerHTML = '';
      const rootId = String(rootSelect.value || '');
      currentRoot = candidateDocs.find((doc) => doc.id === rootId) || null;
      const rootResult = currentRoot?.snapshot?.parseResult || null;
      const rootEngine = ensurePresetEngine(rootResult);
      const existingPresetNames = Object.keys(rootResult?.presetEngine?.presets || {});
      const existingBaseName =
        String(rootResult?.presetEngine?.activePreset || '').trim() ||
        String(existingPresetNames[0] || '').trim();
      baseNameInput.value = existingBaseName || 'Default';
      rootEligible = getPresetEligibleEntries(rootResult);
      const existingScope = new Set(
        ((Array.isArray(rootEngine?.scopeEntries) && rootEngine.scopeEntries.length)
          ? rootEngine.scopeEntries.map((item) => item?.key)
          : (rootResult?.presetEngine?.scope || []))
          .filter((key) => rootEligible.some((item) => item.key === key))
      );
      scopeSelection = existingScope.size
        ? existingScope
        : new Set(rootEligible.map((item) => item.key));
      rootEligible.forEach((item) => {
        const row = document.createElement('label');
        row.className = 'preset-engine-scope-item';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = scopeSelection.has(item.key);
        cb.addEventListener('change', () => {
          if (cb.checked) scopeSelection.add(item.key);
          else scopeSelection.delete(item.key);
        });
        const textWrap = document.createElement('span');
        textWrap.className = 'preset-engine-scope-text';
        const txt = document.createElement('span');
        txt.className = 'preset-engine-scope-label';
        txt.textContent = item.label;
        const meta = document.createElement('span');
        meta.className = 'preset-engine-scope-key';
        meta.textContent = item.key;
        textWrap.appendChild(txt);
        textWrap.appendChild(meta);
        row.appendChild(cb);
        row.appendChild(textWrap);
        scopeWrap.appendChild(row);
      });
      if (!rootEligible.length) {
        status.textContent = 'Root tab has no eligible controls for presets.';
      } else {
        status.textContent = 'Choose a root tab, choose variant tabs, select value scope, then build.';
      }
      renderBuiltState();
    };

    const scopeActions = document.createElement('div');
    scopeActions.className = 'detail-actions';
    const selectAllBtn = document.createElement('button');
    selectAllBtn.type = 'button';
    selectAllBtn.textContent = 'Select all';
    selectAllBtn.addEventListener('click', () => {
      scopeWrap.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
        cb.checked = true;
      });
      rootEligible.forEach((item) => {
        scopeSelection.add(item.key);
      });
    });
    const clearAllBtn = document.createElement('button');
    clearAllBtn.type = 'button';
    clearAllBtn.textContent = 'Select none';
    clearAllBtn.addEventListener('click', () => {
      scopeWrap.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
        cb.checked = false;
      });
      rootEligible.forEach((item) => {
        scopeSelection.delete(item.key);
      });
    });
    scopeActions.appendChild(selectAllBtn);
    scopeActions.appendChild(clearAllBtn);
    rightCol.appendChild(scopeActions);

    body.appendChild(leftCol);
    body.appendChild(rightCol);

    const actions = document.createElement('footer');
    actions.className = 'modal-actions';
    const closeActionBtn = document.createElement('button');
    closeActionBtn.type = 'button';
    closeActionBtn.textContent = 'Close';
    const buildBtn = document.createElement('button');
    buildBtn.type = 'button';
    buildBtn.className = 'primary';
    buildBtn.textContent = 'Build Preset Pack';

    buildBtn.addEventListener('click', () => {
      const rootId = String(rootSelect.value || '');
      const rootDoc = candidateDocs.find((doc) => doc.id === rootId);
      const rootSnap = rootDoc?.snapshot;
      const rootResult = rootSnap?.parseResult;
      if (!rootDoc || !rootSnap || !rootResult) {
        status.textContent = 'Root tab is invalid.';
        return;
      }
      const scopeKeys = rootEligible
        .map((item) => item.key)
        .filter((key) => scopeSelection.has(key));
      if (!scopeKeys.length) {
        status.textContent = 'Select at least one scope control.';
        return;
      }
      const selectedVariants = variantRows.filter((row) => row.cb.checked);
      if (!selectedVariants.length) {
        status.textContent = 'Select at least one variant tab.';
        return;
      }
      const rootPosByKey = new Map(rootEligible.map((item, idx) => [item.key, idx]));
      const scopeEntries = buildPresetScopeEntriesForKeys(rootResult, scopeKeys);
      const base = captureScopeValuesFromSnapshot(rootSnap, rootResult, scopeEntries);
      Object.keys(base).forEach((key) => {
        base[key] = normalizePresetPayloadValue(base[key]);
      });
      const presets = {};
      const usedPresetNames = new Set();
      const warnings = [];
      const requestedBasePresetName = String(baseNameInput.value || '').trim() || 'Default';
      const defaultPresetName = makeUniquePresetName(requestedBasePresetName, usedPresetNames);
      presets[defaultPresetName] = { ...base };
      selectedVariants.forEach((row) => {
        const variantSnap = row.doc?.snapshot;
        const variantResult = variantSnap?.parseResult;
        if (!variantSnap || !variantResult) return;
        const variantEligible = getPresetEligibleEntries(variantResult);
        const variantByKey = new Map();
        variantEligible.forEach((item) => {
          if (item?.scopeEntry) variantByKey.set(item.key, item.scopeEntry);
        });
        const values = {};
        let matched = 0;
        scopeEntries.forEach((scopeEntry) => {
          const key = String(scopeEntry?.key || '').trim();
          if (!key) return;
          let variantScopeEntry = variantByKey.get(key) || null;
          if (!variantScopeEntry) {
            const rootPos = rootPosByKey.get(key);
            if (Number.isFinite(rootPos) && rootPos >= 0 && rootPos < variantEligible.length) {
              variantScopeEntry = variantEligible[rootPos].scopeEntry || null;
            }
          }
          if (!variantScopeEntry) {
            values[key] = base[key];
            return;
          }
          values[key] = normalizePresetPayloadValue(
            getPresetScopeValueFromSnapshot(variantSnap, variantResult, variantScopeEntry)
          );
          matched += 1;
        });
        const presetName = makeUniquePresetName(row.nameInput.value || getDocDisplayName(row.doc, candidateDocs), usedPresetNames);
        presets[presetName] = values;
        if (matched < scopeKeys.length) {
          const fallbackCount = scopeKeys.length - matched;
          warnings.push(`${presetName}: ${fallbackCount} scope control${fallbackCount === 1 ? '' : 's'} used base fallback.`);
        }
      });
      const presetNames = Object.keys(presets);
      if (!presetNames.length) {
        status.textContent = 'No presets could be built from selected variants.';
        return;
      }
      rootResult.presetEngine = {
        version: 2,
        buildMarker: FMR_BUILD_MARKER,
        scope: scopeKeys,
        scopeEntries,
        base,
        presets,
        activePreset: defaultPresetName,
      };
      ensurePresetEngine(rootResult);
      syncPresetSelectorEntryToMacro(
        rootResult,
        resolveMacroGroupSourceOp(rootSnap.originalText, rootResult)
          || String(rootResult.macroName || rootResult.macroNameOriginal || '').trim()
      );
      rootSnap.parseResult = rootResult;
      rootSnap.newline = rootSnap.newline || detectNewline(rootSnap.originalText || '');
      if (rootDoc.id === activeDocId) {
        state.parseResult = rootResult;
        state.newline = rootSnap.newline;
        renderActiveList();
        refreshPageTabs?.();
        updateRemoveSelectedState();
        nodesPane?.parseAndRenderNodes?.();
        const activeDoc = getActiveDocument();
        if (activeDoc) storeDocumentSnapshot(activeDoc);
      } else {
        rootDoc.snapshot = rootSnap;
      }
      renderDocTabs();
      renderBuiltState();
      try {
        builtState.hidden = false;
        builtState.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      } catch (_) {}
      const warningText = warnings.length ? ` Warnings: ${warnings.slice(0, 3).join(' | ')}${warnings.length > 3 ? ' ...' : ''}` : '';
      const rootLabel = getDocDisplayName(rootDoc, candidateDocs) || rootDoc.id;
      status.textContent = `Built ${presetNames.length} preset${presetNames.length === 1 ? '' : 's'} on "${rootLabel}".${warningText}`;
      info(`Preset pack built: ${presetNames.length} preset(s) attached to ${rootLabel}.`);
    });

    clearBuiltBtn.addEventListener('click', () => {
      const rootId = String(rootSelect.value || '');
      const rootDoc = candidateDocs.find((doc) => doc.id === rootId);
      const rootSnap = rootDoc?.snapshot;
      const rootResult = rootSnap?.parseResult;
      if (!rootDoc || !rootSnap || !rootResult) {
        status.textContent = 'Root tab is invalid.';
        return;
      }
      rootResult.presetEngine = {
        version: 2,
        buildMarker: FMR_BUILD_MARKER,
        scope: [],
        scopeEntries: [],
        base: {},
        presets: {},
        activePreset: '',
      };
      ensurePresetEngine(rootResult);
      syncPresetSelectorEntryToMacro(
        rootResult,
        resolveMacroGroupSourceOp(rootSnap.originalText, rootResult)
          || String(rootResult.macroName || rootResult.macroNameOriginal || '').trim()
      );
      rootSnap.parseResult = rootResult;
      if (rootDoc.id === activeDocId) {
        state.parseResult = rootResult;
        state.newline = rootSnap.newline || state.newline || detectNewline(rootSnap.originalText || '');
        renderActiveList();
        refreshPageTabs?.();
        updateRemoveSelectedState();
        nodesPane?.parseAndRenderNodes?.();
        const activeDoc = getActiveDocument();
        if (activeDoc) storeDocumentSnapshot(activeDoc);
      } else {
        rootDoc.snapshot = rootSnap;
      }
      renderScope();
      renderDocTabs();
      renderBuiltState();
      const rootLabel = getDocDisplayName(rootDoc, candidateDocs) || rootDoc.id;
      status.textContent = `Removed preset pack from "${rootLabel}".`;
      info(`Preset pack removed from ${rootLabel}.`);
    });
    builtStateToggle.addEventListener('click', () => {
      if (builtState.classList.contains('is-empty')) return;
      setBuiltStateExpanded(!builtStateExpanded);
    });
    setBuiltStateExpanded(false);

    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    let overlayPointerDown = false;
    closeActionBtn.addEventListener('click', close);
    closeBtn.addEventListener('click', close);
    overlay.addEventListener('mousedown', (ev) => {
      overlayPointerDown = ev.target === overlay;
    });
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay && overlayPointerDown) close();
      overlayPointerDown = false;
    });
    modal.addEventListener('submit', (ev) => ev.preventDefault());
    rootSelect.addEventListener('change', () => {
      rootDocId = String(rootSelect.value || '');
      renderVariants();
      renderScope();
    });

    actions.appendChild(closeActionBtn);
    actions.appendChild(buildBtn);
    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    renderVariants();
    renderScope();
  }

  function resolveRowFromLink(rows, link) {
    if (!Array.isArray(rows) || !rows.length || !link) return null;
    const mode = link.rowMode || 'first';
    if (mode === 'index') {
      const idx = Math.max(1, Number(link.rowIndex || 1)) - 1;
      return rows[idx] || null;
    }
    if (mode === 'key') {
      const key = String(link.rowKey || '').trim();
      const val = String(link.rowValue || '').trim();
      if (!key) return null;
      return rows.find(r => String(r?.[key] ?? '').trim() === val) || null;
    }
    return rows[0] || null;
  }

  function resolveRowForEntry(rows, link, entry, linkData) {
    const data = linkData || entry?.csvLink;
    if (!data) return resolveRowFromLink(rows, link);
    const mode = data.rowMode || 'default';
    if (mode === 'default' || mode === 'first') return resolveRowFromLink(rows, link);
    const override = {
      rowMode: mode,
      rowIndex: data.rowIndex,
      rowKey: data.rowKey,
      rowValue: data.rowValue,
    };
    return resolveRowFromLink(rows, override);
  }

  async function reloadDataLinkForCurrentMacro() {
    if (!state.parseResult) return;
    const perfStart = Date.now();
    let fetchMs = 0;
    let parseMs = 0;
    let applyMs = 0;
    let relinkMs = 0;
    let summaryMs = 0;
    const options = arguments && arguments[0] ? arguments[0] : {};
    const wantSummary = !!options.summary;
    const summaryMaxLines = Number.isFinite(options.maxLines) ? Math.max(1, options.maxLines) : 4;
    const priorMappings = new Map();
    const beforeValues = new Map();
    (state.parseResult.entries || []).forEach((entry) => {
      if (!entry || !entry.sourceOp || !entry.source) return;
      if (!entry.csvLink || !entry.csvLink.column) return;
      priorMappings.set(`${entry.sourceOp}.${entry.source}`, entry.csvLink.column);
      if (wantSummary) {
        const key = `${entry.sourceOp}.${entry.source}`;
        const info = getCurrentInputInfo(entry) || {};
        const name = String(entry.displayName || entry.name || entry.label || key).trim();
        const value = info.value != null ? String(info.value) : '';
        beforeValues.set(key, { name, value });
      }
    });
    const link = ensureDataLink(state.parseResult);
    const source = String(link?.source || '').trim();
    if (!source) {
      error('No data source set for this macro.');
      return;
    }
    const url = normalizeCsvUrl(source);
    try {
      const fetchStart = Date.now();
      const res = await fetch(url);
      fetchMs = Date.now() - fetchStart;
      if (!res.ok) {
        error(`CSV fetch failed (${res.status}).`);
        return;
      }
      const parseStart = Date.now();
      const text = await res.text();
      const parsed = parseCsvText(text);
      parseMs = Date.now() - parseStart;
      if (!parsed.headers.length) {
        error('CSV appears to be empty or invalid.');
        return;
      }
      setCsvData(parsed, source);
      const rows = parsed.rows || [];
      const baseRow = resolveRowFromLink(rows, link);
      if (!baseRow) {
        error('No matching row found for this macro.');
        return;
      }
      const applyStart = Date.now();
      await applyCsvRowsToCurrentMacro(rows, link, { silent: true });
      applyMs = Date.now() - applyStart;
      if (priorMappings.size && state.parseResult) {
        const relinkStart = Date.now();
        let applied = 0;
        (state.parseResult.entries || []).forEach((entry) => {
          if (!entry || !entry.sourceOp || !entry.source) return;
          if (entry.csvLink && entry.csvLink.column) return;
          const col = priorMappings.get(`${entry.sourceOp}.${entry.source}`);
          if (!col) return;
          entry.csvLink = { column: col };
          applied += 1;
        });
        if (applied) {
          ensureDataLink(state.parseResult);
          syncDataLinkMappings(state.parseResult);
          if (activeDetailEntryIndex != null) renderDetailDrawer(activeDetailEntryIndex);
          if (typeof updateDataLinkPanelVisibility === 'function') {
            updateDataLinkPanelVisibility();
          }
        }
        relinkMs = Date.now() - relinkStart;
      }
      let summary = null;
      if (wantSummary && state.parseResult) {
        const summaryStart = Date.now();
        const afterValues = new Map();
        (state.parseResult.entries || []).forEach((entry) => {
          if (!entry || !entry.sourceOp || !entry.source) return;
          if (!entry.csvLink || !entry.csvLink.column) return;
          const key = `${entry.sourceOp}.${entry.source}`;
          const info = getCurrentInputInfo(entry) || {};
          const name = String(entry.displayName || entry.name || entry.label || key).trim();
          const value = info.value != null ? String(info.value) : '';
          afterValues.set(key, { name, value });
        });
        let total = 0;
        const lines = [];
        const truncate = (value) => {
          const str = String(value ?? '');
          if (str.length <= 60) return str;
          return `${str.slice(0, 57)}...`;
        };
        beforeValues.forEach((before, key) => {
          const after = afterValues.get(key);
          if (!after) return;
          if (String(before.value) === String(after.value)) return;
          total += 1;
          if (lines.length < summaryMaxLines) {
            lines.push(`${before.name}: "${truncate(before.value)}" → "${truncate(after.value)}"`);
          }
        });
        summary = { lines, total };
        summaryMs = Date.now() - summaryStart;
      }
      logDiag(`[CSV reload] rows=${rows.length}, headers=${parsed.headers.length}, fetch=${fetchMs}ms, parse=${parseMs}ms, apply=${applyMs}ms, relink=${relinkMs}ms, summary=${summaryMs}ms, total=${Date.now() - perfStart}ms`);
      info('Data reloaded for current macro.');
      return summary;
    } catch (err) {
      logDiag(`[CSV reload] failed after ${Date.now() - perfStart}ms: ${err?.message || err}`);
      error(err?.message || err || 'Failed to reload data.');
    }
  }

  async function applyCsvRowsToCurrentMacroNow(rows, link, options = {}) {
    if (!state.parseResult) return;
    const perfStart = Date.now();
    const perf = { planMs: 0, patchMs: 0, reloadMs: 0 };
    const stats = {
      rowCacheHits: 0,
      rowCacheMisses: 0,
      planSize: 0,
      mixedMode: false,
      entryPlanCount: 0,
      editsCount: 0,
    };
    const priorMappings = new Map();
    const priorHexMappings = new Map();
    (state.parseResult.entries || []).forEach((entry) => {
      if (!entry || !entry.sourceOp || !entry.source) return;
      const key = `${entry.sourceOp}.${entry.source}`;
      if (entry.csvLink && entry.csvLink.column) {
        priorMappings.set(key, { ...entry.csvLink });
      }
      if (entry.csvLinkHex && entry.csvLinkHex.column) {
        priorHexMappings.set(key, { ...entry.csvLinkHex });
      }
    });
    if (!Array.isArray(rows) || !rows.length) {
      if (!options.silent) error('No matching row found for this macro.');
      return;
    }
    const eol = state.newline || '\n';
    let workingText = state.originalText || '';
    const linkedEntries = (state.parseResult.entries || []).filter(entry => entry && entry.csvLink && entry.csvLink.column);
    const hexEntries = (state.parseResult.entries || [])
      .map((entry, idx) => ({ entry, idx }))
      .filter(({ entry }) => entry && entry.csvLinkHex && entry.csvLinkHex.column);
    const rowCache = new Map();
    const resolveEntryRowCached = (entry, overrideLink = null) => {
      if (!entry || !entry.sourceOp || !entry.source) return null;
      const linkObj = overrideLink || entry.csvLink;
      if (!linkObj || !linkObj.column) return null;
      const rowMode = String(linkObj.rowMode || 'default');
      const rowIndex = Number.isFinite(linkObj.rowIndex) ? Number(linkObj.rowIndex) : '';
      const rowKey = linkObj.rowKey != null ? String(linkObj.rowKey) : '';
      const rowValue = linkObj.rowValue != null ? String(linkObj.rowValue) : '';
      const cacheKey = `${linkObj.column}|${rowMode}|${rowIndex}|${rowKey}|${rowValue}|${overrideLink ? 'hex' : 'csv'}`;
      if (rowCache.has(cacheKey)) {
        stats.rowCacheHits += 1;
        return rowCache.get(cacheKey);
      }
      stats.rowCacheMisses += 1;
      const resolved = resolveRowForEntry(rows, link, entry, overrideLink) || null;
      rowCache.set(cacheKey, resolved);
      return resolved;
    };
    const planStart = Date.now();
    const plan = buildCsvPatchPlan(workingText, state.parseResult, linkedEntries);
    perf.planMs = Date.now() - planStart;
    stats.planSize = Array.isArray(plan) ? plan.length : 0;
    const patchStart = Date.now();
    if (plan && plan.length) {
      const baseRow = resolveRowFromLink(rows, link);
      if (!baseRow) {
        if (!options.silent) error('No matching row found for this macro.');
        return;
      }
      if (linkedEntries.every((entry) => (entry.csvLink?.rowMode || 'default') === 'default')) {
        workingText = applyCsvRowEdits(workingText, plan, baseRow, eol);
      } else {
        stats.mixedMode = true;
        const planByEntry = new Map();
        for (const patch of plan) {
          if (!patch || !patch.entry) continue;
          let list = planByEntry.get(patch.entry);
          if (!list) {
            list = [];
            planByEntry.set(patch.entry, list);
          }
          list.push(patch);
        }
        stats.entryPlanCount = planByEntry.size;
        const edits = [];
        for (const entry of linkedEntries) {
          const row = resolveEntryRowCached(entry);
          if (!row) continue;
          const entryPlan = planByEntry.get(entry);
          if (!entryPlan || !entryPlan.length) continue;
          for (const patch of entryPlan) {
            const rawVal = row[entry.csvLink.column] != null ? String(row[entry.csvLink.column]) : '';
            if (patch.kind === 'tool') {
              const body = workingText.slice(patch.start + 1, patch.end);
              const next = updateToolInputBlockBody(body, entry, rawVal, patch.indent, eol);
              if (next !== body) edits.push({ start: patch.start + 1, end: patch.end, text: next });
            } else if (patch.kind === 'instance') {
              const body = workingText.slice(patch.start + 1, patch.end);
              const formatted = formatDefaultForStorage(entry, rawVal);
              if (formatted != null) {
                entry.controlMeta = entry.controlMeta || {};
                entry.controlMetaOriginal = entry.controlMetaOriginal || {};
                entry.controlMeta.defaultValue = formatted;
                entry.controlMetaDirty = true;
                applyDefaultToEntryRaw(entry, formatted, eol);
              }
              let next = body;
              if (formatted == null) next = removeInstanceInputProp(body, 'Default');
              else next = setInstanceInputProp(body, 'Default', formatted, patch.indent, eol || '\n');
              if (next !== body) edits.push({ start: patch.start + 1, end: patch.end, text: next });
            }
          }
        }
        if (edits.length) {
          stats.editsCount = edits.length;
          edits.sort((a, b) => b.start - a.start);
          let updated = workingText;
          for (const edit of edits) {
            updated = updated.slice(0, edit.start) + edit.text + updated.slice(edit.end);
          }
          workingText = updated;
        }
      }
    } else {
      linkedEntries.forEach((entry) => {
        const row = resolveEntryRowCached(entry);
        if (!row) return;
        const col = entry.csvLink.column;
        const rawVal = row[col] != null ? String(row[col]) : '';
        if (isTextControl(entry)) {
          workingText = setInputValueInToolText(workingText, entry, rawVal, eol, state.parseResult);
        } else {
          const formatted = formatDefaultForStorage(entry, rawVal);
          entry.controlMeta = entry.controlMeta || {};
          entry.controlMetaOriginal = entry.controlMetaOriginal || {};
          entry.controlMeta.defaultValue = formatted;
          entry.controlMetaDirty = true;
          applyDefaultToEntryRaw(entry, formatted, eol);
        }
      });
    }
    if (hexEntries.length) {
      hexEntries.forEach(({ entry, idx }) => {
        const row = resolveEntryRowCached(entry, entry.csvLinkHex);
        if (!row) return;
        const col = entry.csvLinkHex.column;
        const rawVal = row[col] != null ? String(row[col]) : '';
        const parsed = parseHexColor(rawVal);
        if (!parsed) return;
        const r = clamp01(parsed.r);
        const g = clamp01(parsed.g);
        const b = clamp01(parsed.b);
        const a = parsed.hasAlpha ? clamp01(parsed.a) : null;
        workingText = applyHexToColorGroup({
          entries: state.parseResult.entries,
          order: state.parseResult.order,
          baseText: workingText,
          groupIndex: idx,
          r,
          g,
          b,
          a,
          includeAlpha: parsed.hasAlpha,
          eol,
          resultRef: state.parseResult,
        });
      });
    }
    perf.patchMs = Date.now() - patchStart;
    syncDataLinkMappings(state.parseResult);
    const beforePersistLen = workingText.length;
    let updated = workingText;
    // Keep CSV row-apply lightweight. Hidden metadata is persisted on export/save paths.
    const persistMarkers = !!options.persistMarkers;
    if (persistMarkers) {
      const dataLink = buildFmrDataLinkPayload(state.parseResult);
      const fileMetaPath = state.originalFilePath
        || state.lastExportPath
        || (state.exportFolder ? buildExportPath(state.exportFolder, buildMacroExportName()) : '');
      if (dataLink) {
        updated = stripFmrDataLinkBlock(updated);
        updated = upsertHiddenDataLinkControl(updated, state.parseResult, dataLink, eol);
      }
      if (fileMetaPath) {
        updated = upsertHiddenFileMetaControl(updated, state.parseResult, { exportPath: fileMetaPath, buildMarker: FMR_BUILD_MARKER }, eol);
        updated = rewriteUpdateDataProtocolPaths(updated, fileMetaPath);
      }
    }
    const afterPersistLen = updated.length;
    const reloadStart = Date.now();
    const prevGenerated = state.generatedText || '';
    state.originalText = updated;
    state.newline = detectNewline(updated);
    state.lastDiffRange = computeDiffRange(prevGenerated, updated);
    updateCodeView(updated);
    applyPendingHighlight();
    if (activeDetailEntryIndex != null && state.parseResult?.entries?.[activeDetailEntryIndex]) {
      renderDetailDrawer(activeDetailEntryIndex);
    }
    renderActiveList({ safe: true });
    syncDataLinkPanel();
    markActiveDocumentDirty();
    const activeDoc = getActiveDocument();
    if (activeDoc) storeDocumentSnapshot(activeDoc);
    perf.reloadMs = Date.now() - reloadStart;
    logDiag(`[CSV] apply rows: linked=${linkedEntries.length}, hex=${hexEntries.length}, ms=${Date.now() - perfStart}, plan=${perf.planMs}, patch=${perf.patchMs}, reload=${perf.reloadMs}, mixed=${stats.mixedMode ? 1 : 0}, planSize=${stats.planSize}, planEntries=${stats.entryPlanCount}, edits=${stats.editsCount}, rowCache=${rowCache.size}, rowHit=${stats.rowCacheHits}, rowMiss=${stats.rowCacheMisses}, persist=${persistMarkers ? 1 : 0}, len=${beforePersistLen}->${afterPersistLen}`);
    if (!options.silent) info('Data applied from CSV.');
  }

  let csvApplyQueuedJob = null;
  let csvApplyRunning = false;
  let csvApplyWaiters = [];

  async function applyCsvRowsToCurrentMacro(rows, link, options = {}) {
    return new Promise((resolve, reject) => {
      const hadPending = !!csvApplyQueuedJob;
      csvApplyQueuedJob = { rows, link, options };
      csvApplyWaiters.push({ resolve, reject });
      if (hadPending) {
        logDiag(`[CSV queue] coalesced request (waiters=${csvApplyWaiters.length})`);
      }
      if (csvApplyRunning) return;
      const pump = async () => {
        csvApplyRunning = true;
        let lastResult;
        let lastError = null;
        let jobs = 0;
        while (csvApplyQueuedJob) {
          const job = csvApplyQueuedJob;
          csvApplyQueuedJob = null;
          jobs += 1;
          try {
            lastResult = await applyCsvRowsToCurrentMacroNow(job.rows, job.link, job.options || {});
            lastError = null;
          } catch (err) {
            lastError = err;
          }
        }
        const waiters = csvApplyWaiters.splice(0);
        csvApplyRunning = false;
        logDiag(`[CSV queue] drained jobs=${jobs}, resolvedWaiters=${waiters.length}, error=${lastError ? 1 : 0}`);
        waiters.forEach((w) => {
          try {
            if (lastError) w.reject(lastError);
            else w.resolve(lastResult);
          } catch (_) {}
        });
      };
      pump();
    });
  }

  async function openCsvUrlPrompt() {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.className = 'add-control-modal csv-modal';
      const modal = document.createElement('form');
      modal.className = 'add-control-form';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-modal', 'true');
      const header = document.createElement('header');
      const headingWrap = document.createElement('div');
      const eyebrow = document.createElement('p');
      eyebrow.className = 'eyebrow';
      eyebrow.textContent = 'CSV';
      const heading = document.createElement('h3');
      heading.textContent = 'Import CSV from URL';
      headingWrap.appendChild(eyebrow);
      headingWrap.appendChild(heading);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = '×';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(headingWrap);
      header.appendChild(closeBtn);
      const body = document.createElement('div');
      body.className = 'form-body';
      const label = document.createElement('label');
      label.textContent = 'CSV URL';
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = 'https://example.com/data.csv';
      label.appendChild(input);
      body.appendChild(label);
      const actions = document.createElement('footer');
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel';
      const okBtn = document.createElement('button');
      okBtn.type = 'submit';
      okBtn.className = 'primary';
      okBtn.textContent = 'Import';
      const close = (value) => {
        try { overlay.remove(); } catch (_) {}
        resolve(value);
      };
      cancelBtn.addEventListener('click', () => close(null));
      closeBtn.addEventListener('click', () => close(null));
      modal.addEventListener('submit', (ev) => {
        ev.preventDefault();
        close(input.value);
      });
      overlay.addEventListener('click', (ev) => {
        if (ev.target === overlay) close(null);
      });
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);
      modal.appendChild(header);
      modal.appendChild(body);
      modal.appendChild(actions);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
      input.focus();
    });
  }

  async function generateSettingsFromCsvCore(snapshot, options = {}) {
    const snap = snapshot || {};
    const result = snap.parseResult || null;
    const originalText = snap.originalText || '';
    const csvData = snap.csvData || null;
    if (!result || !originalText) {
      error('Load a macro before generating from CSV.');
      return false;
    }
    if (!csvData || !Array.isArray(csvData.headers) || !csvData.headers.length) {
      error('Import a CSV first.');
      return false;
    }
    const linked = (result.entries || []).filter((entry) => entry && (entry.csvLink && entry.csvLink.column || entry.csvLinkHex && entry.csvLinkHex.column));
    if (!linked.length) {
      error('No controls are linked to CSV columns yet.');
      return false;
    }
    const promptName = options.promptNameColumn !== false;
    let nameColumn = csvData.nameColumn || null;
    if (promptName) {
      nameColumn = await openCsvNameColumnPicker(csvData.headers, nameColumn);
      if (nameColumn == null) return false;
      csvData.nameColumn = nameColumn;
    }
    if (!nameColumn) nameColumn = '__row__';
    const baseName = (result.macroName || result.macroNameOriginal || 'Macro').trim();
    const eol = snap.newline || detectNewline(originalText) || '\n';
    const rows = csvData.rows || [];
    if (!rows.length) {
      error('CSV has no data rows.');
      return false;
    }
    const filesOnly = !!options.filesOnly;
    const lazyTabs = options.lazyTabs !== false;
    const openTabLimit = Number.isFinite(options.openTabLimit) ? Math.max(0, options.openTabLimit) : null;
    const overflowToFolder = !!options.overflowToFolder;
    const folderPath = options.folderPath || snap.exportFolder || state.exportFolder || resolveExportFolder();
    if (filesOnly && !folderPath) {
      error('No export folder available. Set an export folder first.');
      return false;
    }
    const prevSuppress = suppressDocDirty;
    const prevTabsSuppress = suppressDocTabsRender;
    suppressDocDirty = true;
    suppressDocTabsRender = true;
    const baseText = rebuildContentWithNewOrder(originalText || '', result, eol, { commitChanges: false });
    const linkedText = (result.entries || []).filter((entry) => entry && entry.csvLink && entry.csvLink.column);
    const linkedHex = (result.entries || []).map((entry, idx) => ({ entry, idx }))
      .filter(({ entry }) => entry && entry.csvLinkHex && entry.csvLinkHex.column);
    const patchPlan = buildCsvPatchPlan(baseText, result, linkedText);
    const useFastPath = Array.isArray(patchPlan) && patchPlan.length > 0;
    const rowLabels = buildUniqueCsvNames(rows, nameColumn, baseName);
    const baseLink = ensureDataLink(result);
    const applyHexRowEdits = (text, row, resultRef) => {
      if (!linkedHex.length) return text;
      let updated = text;
      linkedHex.forEach(({ entry, idx }) => {
        const col = entry.csvLinkHex.column;
        const rawVal = row[col] != null ? String(row[col]) : '';
        const parsed = parseHexColor(rawVal);
        if (!parsed) return;
        const r = clamp01(parsed.r);
        const g = clamp01(parsed.g);
        const b = clamp01(parsed.b);
        const a = parsed.hasAlpha ? clamp01(parsed.a) : null;
        updated = applyHexToColorGroup({
          entries: resultRef?.entries || [],
          order: resultRef?.order || [],
          baseText: updated,
          groupIndex: idx,
          r,
          g,
          b,
          a,
          includeAlpha: parsed.hasAlpha,
          eol,
          resultRef,
        });
      });
      return updated;
    };
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i] || {};
      const rowLabel = rowLabels[i] || `${baseName} ${i + 1}`;
      const rowLink = baseLink
        ? { ...baseLink, rowMode: 'index', rowIndex: i + 1 }
        : null;
      let content = '';
      if (useFastPath) {
        let workingText = applyCsvRowEdits(baseText, patchPlan, row, eol);
        workingText = applyHexRowEdits(workingText, row, result);
        const tmpResult = {
          macroNameOriginal: result.macroNameOriginal || result.macroName || '',
          macroName: rowLabel,
          operatorType: result.operatorType || result.operatorTypeOriginal || 'GroupOperator',
          operatorTypeOriginal: result.operatorTypeOriginal || result.operatorType || 'GroupOperator',
          dataLink: rowLink || null,
          csvData,
          entries: result.entries || [],
        };
        workingText = applyMacroNameRename(workingText, tmpResult);
        const safeMacroName = deriveSafeMacroName(rowLabel, result);
        workingText = ensureActiveToolName(workingText, safeMacroName, eol);
        const linkPayload = buildFmrDataLinkPayload(tmpResult);
        content = workingText;
        if (linkPayload) {
          content = stripFmrDataLinkBlock(content);
          content = upsertHiddenDataLinkControl(content, tmpResult, linkPayload, eol);
        }
      } else {
        const resultClone = cloneParseResultForCsv(result);
        if (!resultClone) continue;
        resultClone.macroName = rowLabel;
        resultClone.dataLink = rowLink || null;
        resultClone.csvData = csvData;
        let workingText = baseText;
        const entries = resultClone.entries || [];
        entries.forEach((entry) => {
          if (!entry || !entry.csvLink || !entry.csvLink.column) return;
          const col = entry.csvLink.column;
          const rawVal = row[col] != null ? String(row[col]) : '';
          if (isTextControl(entry)) {
            workingText = setInputValueInToolText(workingText, entry, rawVal, eol, result);
          } else {
            const formatted = formatDefaultForStorage(entry, rawVal);
            entry.controlMeta = entry.controlMeta || {};
            entry.controlMetaOriginal = entry.controlMetaOriginal || {};
            entry.controlMeta.defaultValue = formatted;
            entry.controlMetaDirty = true;
            applyDefaultToEntryRaw(entry, formatted, eol);
          }
        });
        workingText = applyHexRowEdits(workingText, row, resultClone);
        content = rebuildContentWithNewOrder(workingText, resultClone, eol);
      }
      const safeName = sanitizeFileBaseName(rowLabel || baseName);
      const fileName = `${safeName || 'Macro'}.setting`;
      const shouldOpenTab = !filesOnly && (openTabLimit == null || i < openTabLimit);
      if (!shouldOpenTab) {
        if (overflowToFolder || filesOnly) {
          if (!folderPath) {
            error('No export folder available. Set an export folder first.');
            suppressDocDirty = prevSuppress;
            suppressDocTabsRender = prevTabsSuppress;
            return false;
          }
          try {
            writeSettingToFolder(folderPath, fileName, content);
          } catch (err) {
            error(err?.message || err || 'Failed to write CSV output.');
            suppressDocDirty = prevSuppress;
            suppressDocTabsRender = prevTabsSuppress;
            return false;
          }
          continue;
        }
        continue;
      }
      if (filesOnly) {
        try {
          writeSettingToFolder(folderPath, fileName, content);
        } catch (err) {
          error(err?.message || err || 'Failed to write CSV output.');
          suppressDocDirty = prevSuppress;
          suppressDocTabsRender = prevTabsSuppress;
          return false;
        }
      } else if (lazyTabs) {
        createLazyDocumentFromText({ name: rowLabel || baseName, fileName, text: content });
      } else {
        await loadMacroFromText(fileName, content, {
          preserveFileInfo: true,
          preserveFilePath: false,
          createDoc: true,
          allowAutoUtility: false,
          skipClear: false,
        });
        const doc = getActiveDocument();
        if (doc) {
          doc.name = rowLabel || doc.name;
          doc.fileName = fileName;
          doc.isDirty = true;
        }
      }
    }
    suppressDocDirty = prevSuppress;
    suppressDocTabsRender = prevTabsSuppress;
    renderDocTabs();
    if (filesOnly) {
      addCsvBatchDocument({
        count: rows.length,
        folderPath,
        baseName,
        sourceDocId: options.sourceDocId || null,
        snapshot: options.storeSnapshot ? snap : null,
      });
    }
    info(`Generated ${rows.length} macros from CSV.${filesOnly ? ' Files saved to export folder.' : ''}`);
    return true;
  }

  async function generateSettingsFromCsv() {
    if (!state.parseResult || !state.originalText) {
      error('Load a macro before generating from CSV.');
      return;
    }
    if (!state.csvData || !Array.isArray(state.csvData.headers) || !state.csvData.headers.length) {
      error('Import a CSV first.');
      return;
    }
    const activeDoc = getActiveDocument();
    if (activeDoc) storeDocumentSnapshot(activeDoc);
    const filesOnly = false;
    const openTabLimit = null;
    const overflowToFolder = false;
    const snap = buildDocumentSnapshot();
    if (!snap) return;
    await generateSettingsFromCsvCore(snap, {
      filesOnly,
      promptNameColumn: true,
      sourceDocId: activeDoc?.id || null,
      storeSnapshot: false,
      openTabLimit,
      overflowToFolder,
      lazyTabs: true,
    });
  }



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



  exportClipboardBtn?.addEventListener('click', exportToClipboard);
  // Cross-pane highlight
  function clearHighlights() {
    try {
      document.querySelectorAll('.hl-pub').forEach((el) => el.classList.remove('hl-pub'));
      document.querySelectorAll('.hl-node').forEach((el) => el.classList.remove('hl-node'));
    } catch (_) {}
  }

  function deriveGroupBase(source) {
    if (!source) return null;
    let base = null;
    try {
      base = (typeof deriveColorBaseFromId === 'function') ? deriveColorBaseFromId(source) : null;
    } catch (_) {
      base = null;
    }
    if (!base) {
      const suffix = ['Red', 'Green', 'Blue', 'Alpha'].find((s) => String(source || '').endsWith(s));
      if (suffix) base = String(source).slice(0, String(source).length - suffix.length);
    }
    return base;
  }

  function getColorGroupBlockByIndexFrom(entries, order, idx) {
    const entry = entries[idx];
    if (!entry || !entry.sourceOp) return null;
    const startPos = order.indexOf(idx);
    if (startPos < 0) return null;

    if (Number.isFinite(entry.controlGroup)) {
      const key = `${entry.sourceOp}::${entry.controlGroup}`;
      let first = startPos;
      let last = startPos;
      for (let p = startPos - 1; p >= 0; p--) {
        const e2 = entries[order[p]];
        if (!e2 || e2.isLabel || !Number.isFinite(e2.controlGroup) || e2.sourceOp !== entry.sourceOp) break;
        if (`${e2.sourceOp}::${e2.controlGroup}` !== key) break;
        first = p;
      }
      for (let p = startPos + 1; p < order.length; p++) {
        const e2 = entries[order[p]];
        if (!e2 || e2.isLabel || !Number.isFinite(e2.controlGroup) || e2.sourceOp !== entry.sourceOp) break;
        if (`${e2.sourceOp}::${e2.controlGroup}` !== key) break;
        last = p;
      }
      const indices = order.slice(first, last + 1);
      return { firstIndex: order[first], lastIndex: order[last], indices, count: indices.length, key };
    }

    const base = deriveGroupBase(entry.source);
    if (!base) return null;
    const key = `${entry.sourceOp}::base:${base}`;
    const indices = order.filter((orderIdx) => {
      const e2 = entries[orderIdx];
      if (!e2 || e2.isLabel || e2.sourceOp !== entry.sourceOp) return false;
      const b2 = deriveGroupBase(e2.source);
      return !!b2 && b2 === base;
    });
    if (!indices.length) return null;
    return {
      firstIndex: indices[0],
      lastIndex: indices[indices.length - 1],
      indices,
      count: indices.length,
      key,
    };
  }

  function cleanDocLabelValue(value) {
    return String(value || '').replace(/\s+\*$/, '').trim();
  }

  function getFileLabelBase(fileName) {
    const raw = cleanDocLabelValue(fileName);
    if (!raw) return '';
    const leaf = raw.split(/[\\/]/).pop() || raw;
    return leaf.replace(/\.[^.]+$/, '').trim();
  }

  function isPlaceholderMacroName(name) {
    const value = cleanDocLabelValue(name).toLowerCase();
    if (!value) return true;
    return /^(new\s*macro|newmacro|clipboard|untitled|unknown|imported)$/.test(value);
  }

  function getDocDisplayName(doc, docs = documents) {
    if (!doc) return 'Untitled';
    if (doc.isCsvBatch) return cleanDocLabelValue(doc.name) || 'CSV Batch';
    const name = cleanDocLabelValue(doc.name);
    const fileLabel = getFileLabelBase(doc.fileName || doc?.snapshot?.originalFileName || '');
    const hasRealName = !!name && !isPlaceholderMacroName(name);
    if (hasRealName) {
      const duplicate = (docs || []).filter((item) => cleanDocLabelValue(item?.name).toLowerCase() === name.toLowerCase()).length > 1;
      if (duplicate && fileLabel && fileLabel.toLowerCase() !== name.toLowerCase()) {
        return `${name} - ${fileLabel}`;
      }
      return name;
    }
    if (fileLabel) return fileLabel;
    if (name) return name;
    return 'Untitled';
  }

  function deriveMacroNameFromFileName(fileName) {
    const base = getFileLabelBase(fileName);
    if (!base) return '';
    return base.replace(/[_-]+/g, ' ').trim();
  }

  function resolveExportMacroName(result, fileMetaPath = '') {
    const currentName = cleanDocLabelValue(result?.macroName || '');
    const originalName = cleanDocLabelValue(result?.macroNameOriginal || '');
    if (currentName && !isPlaceholderMacroName(currentName)) return currentName;
    if (!currentName && originalName && !isPlaceholderMacroName(originalName)) return originalName;
    const fallbackFile = fileMetaPath
      || state.originalFilePath
      || state.originalFileName
      || '';
    return deriveMacroNameFromFileName(fallbackFile) || currentName || originalName || '';
  }

  function getColorGroupComponentsFrom(entries, cgBlock) {
    if (!cgBlock) return null;
    const parts = { red: null, green: null, blue: null, alpha: null };
    cgBlock.indices.forEach((idx) => {
      const e = entries[idx];
      if (!e || !e.source) return;
      const name = normalizeId(String(e.source || '')).toLowerCase();
      if (/red(\d*)$/.test(name)) { parts.red = idx; return; }
      if (/green(\d*)$/.test(name)) { parts.green = idx; return; }
      if (/blue(\d*)$/.test(name)) { parts.blue = idx; return; }
      if (/alpha(\d*)$/.test(name)) { parts.alpha = idx; return; }
      if (/r(\d*)$/.test(name)) { parts.red = idx; return; }
      if (/g(\d*)$/.test(name)) { parts.green = idx; return; }
      if (/b(\d*)$/.test(name)) { parts.blue = idx; return; }
      if (/a(\d*)$/.test(name)) { parts.alpha = idx; return; }
      if (name.includes('red')) { parts.red = idx; return; }
      if (name.includes('green')) { parts.green = idx; return; }
      if (name.includes('blue')) { parts.blue = idx; return; }
      if (name.includes('alpha')) { parts.alpha = idx; return; }
    });
    if (parts.red == null || parts.green == null || parts.blue == null) {
      const ordered = cgBlock.indices.slice();
      if (ordered.length >= 3) {
        if (parts.red == null) parts.red = ordered[0];
        if (parts.green == null) parts.green = ordered[1];
        if (parts.blue == null) parts.blue = ordered[2];
        if (ordered.length >= 4 && parts.alpha == null) parts.alpha = ordered[3];
      }
    }
    if (parts.red == null || parts.green == null || parts.blue == null) {
      const ranked = cgBlock.indices.slice().map((idx) => {
        const e = entries[idx];
        const name = e && e.source ? normalizeId(String(e.source)).toLowerCase() : '';
        let score = 10;
        if (name.includes('red')) score = 1;
        else if (name.includes('green')) score = 2;
        else if (name.includes('blue')) score = 3;
        else if (name.includes('alpha')) score = 4;
        return { idx, score };
      }).sort((a, b) => a.score - b.score);
      const ordered = ranked.map(r => r.idx);
      if (ordered.length >= 3) {
        if (parts.red == null) parts.red = ordered[0];
        if (parts.green == null) parts.green = ordered[1];
        if (parts.blue == null) parts.blue = ordered[2];
        if (ordered.length >= 4 && parts.alpha == null) parts.alpha = ordered[3];
      }
    }
    if (parts.red == null || parts.green == null || parts.blue == null) {
      try {
        const names = cgBlock.indices.map((idx) => {
          const e = entries[idx];
          return e && e.source ? normalizeId(String(e.source)) : '?';
        });
        logDiag?.(`[ColorGroup] unresolved mapping key=${cgBlock.key} names=${names.join(', ')} parts=${JSON.stringify(parts)}`);
      } catch (_) {}
    }
    return parts;
  }

  function getColorGroupBlockByIndex(idx) {
    if (!state.parseResult) return null;
    const entries = state.parseResult.entries || [];
    const order = state.parseResult.order || [];
    return getColorGroupBlockByIndexFrom(entries, order, idx);
  }

  function getColorGroupComponents(cgBlock) {
    if (!cgBlock || !state.parseResult) return null;
    const entries = state.parseResult.entries || [];
    return getColorGroupComponentsFrom(entries, cgBlock);
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
          let el = nodesList.querySelector(`.node-published-indicator[data-source-op="${op}"][data-source="${src}"]`);
          if (el) return el;

          const base = deriveGroupBase(src);
          if (base) {
            el = nodesList.querySelector(`.node-published-indicator.group[data-source-op="${op}"][data-group-id="${base}"]`);
            if (el) return el;
          }

          const indicators = nodesList.querySelectorAll('.node-published-indicator');
          for (const indicator of indicators) {
            if (!indicator.classList.contains('group')) {
              if ((indicator.dataset.sourceOp || '') === op && (indicator.dataset.source || '') === src) return indicator;
            } else {
              const dsBase = indicator.dataset.groupId || '';
              if (dsBase && (indicator.dataset.sourceOp || '') === op) {
                if (base && dsBase === base) return indicator;
              }
            }
          }
          return null;
        }

        const el = nodesList.querySelector(`.node-published-indicator.group[data-source-op="${op}"]`);
        if (el) return el;
        const indicators = nodesList.querySelectorAll('.node-published-indicator.group');
        for (const indicator of indicators) {
          if ((indicator.dataset.sourceOp || '') === op) return indicator;
        }
        return null;
      };

      let target = findEl();
      if (!target && nodesPane?.getNodeFilter?.()) {
        nodesPane?.clearFilter?.();
        target = findEl();
      }

      if (!target) {
        // Fallback: highlight the node header by data attribute
        try {
          const wrap = nodesList && nodesList.querySelector(`.node[data-op="${op}"] .node-header`);
          if (wrap) {
            wrap.classList.add('hl-node');
            if (pulse) wrap.classList.add('pulse');
            scrollRowIntoView(wrap, nodesList);
            setTimeout(() => wrap.classList.remove('pulse'), 800);
            return true;
          }
        } catch (_) {}
        return false;
      }

      const row = target.closest('.node-row') || target;
      row.classList.add('hl-node');
      if (pulse) row.classList.add('pulse');
      scrollRowIntoView(row, nodesList);
      setTimeout(() => row.classList.remove('pulse'), 800);
      return true;
    } catch (_) {
      return false;
    }
  }

  function focusPublishedControl(sourceOp, source, pulse = true) {
    try {
      if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return false;
      const op = String(sourceOp || '').trim();
      const src = String(source || '').trim();
      if (!op || !src) return false;
      const idx = state.parseResult.entries.findIndex((entry) => entry && entry.sourceOp === op && entry.source === src);
      if (idx < 0) return false;
      const entry = state.parseResult.entries[idx];
      const targetPage = String(entry?.page || 'Controls').trim() || 'Controls';
      if (state.parseResult.activePage !== targetPage) {
        state.parseResult.activePage = targetPage;
        try { refreshPageTabs(); } catch (_) {}
      }
      if (publishedSearch && publishedSearch.value) {
        publishedSearch.value = '';
        setPublishedFilter('');
      }
      const sel = new Set([idx]);
      state.parseResult.selected = sel;
      if (typeof setPublishedDetailTarget === 'function') {
        try { setPublishedDetailTarget(idx, { render: false }); } catch (_) {}
      }
      renderList(state.parseResult.entries, state.parseResult.order);
      const row = controlsList ? controlsList.querySelector(`li[data-index="${idx}"]`) : null;
      if (!row) return true;
      row.classList.add('hl-pub');
      if (pulse) row.classList.add('pulse');
      try { scrollRowIntoView(row, controlsList); } catch (_) {}
      setTimeout(() => {
        try { row.classList.remove('pulse'); } catch (_) {}
      }, 800);
      return true;
    } catch (_) {
      return false;
    }
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
    if (ctrlCountEl) ctrlCountEl.textContent = '0';
    if (inputCountEl) inputCountEl.textContent = '0';
    if (csvBatchBanner) csvBatchBanner.hidden = true;
    removeCommonPageBtn && (removeCommonPageBtn.disabled = true);
    deselectAllBtn && (deselectAllBtn.disabled = true);
    setCreateControlActionsEnabled(false);
    hideDetailDrawer();
    updateUtilityActionsState();
    updateExportMenuButtonState();
    updateIntroToggleVisibility();
    cancelPendingCodeRefresh();
    updateCodeView('');
    state.lastDiffRange = null;
    pendingControlMeta.clear();
    pendingOpenNodes.clear();
    updatePublishedFeatureBadges();

  }

  function updateUtilityActionsState() {
    if (addUtilityNodeBtn) {
      addUtilityNodeBtn.disabled = !state.parseResult;
    }
    if (routingBtn) {
      routingBtn.disabled = !state.parseResult;
    }
  }

  function countRecognizedInputs(entries) {
    if (!Array.isArray(entries) || !entries.length) return 0;
    let count = 0;
    for (const entry of entries) {
      if (!entry) continue;
      if (entry.locked) { count++; continue; }
      const key = typeof entry.key === 'string' ? entry.key.toLowerCase() : '';
      if (/^maininput\d+$/.test(key) || /^main\d+input$/.test(key)) count++;
    }
    return count;
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
        const rebuilt = rebuildContentWithNewOrder(state.originalText, state.parseResult, state.newline, { commitChanges: false });
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



  const MAX_MESSAGE_LINES = 220;

  function appendMessageLine(text, isError = false) {
    if (!messages) return;
    const div = document.createElement('div');
    if (isError) div.style.color = '#ff9a9a';
    div.textContent = text;
    messages.appendChild(div);
    while (messages.childElementCount > MAX_MESSAGE_LINES) {
      try { messages.removeChild(messages.firstChild); } catch (_) { break; }
    }
  }

  function info(msg) {
    appendMessageLine(msg, false);
  }

  function error(msg) {
    appendMessageLine(`Error: ${msg}`, true);
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

  function setNodesHidden(hidden) {
    nodesHidden = !!hidden;
    const wasHidden = document?.body?.classList?.contains('nodes-hidden');
    if (typeof document !== 'undefined') {
      document.body.classList.toggle('nodes-hidden', nodesHidden);
    }
    if (controlsSection) {
      controlsSection.classList.toggle('nodes-hidden', nodesHidden);
    }
    if (nodesPaneEl) {
      nodesPaneEl.hidden = nodesHidden;
    }
    if (hideNodesBtn) {
      hideNodesBtn.hidden = nodesHidden;
    }
    if (showNodesBtn) {
      showNodesBtn.hidden = !nodesHidden;
    }
    if (wasHidden && !nodesHidden && activeDetailEntryIndex == null) {
      collapseDetailDrawer();
    }
    refreshDetailDrawerState();
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
        const after = text.slice(toolsClose);
        const newInner = wrappedGroup;
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
      const lowerRaw = raw.trim().toLowerCase();
      if (lowerRaw === 'clipboard') {
        return 'NewMacro';
      }
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
      return text;
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
    if (dot > 0) return sanitizeFileBaseName(name.slice(0, dot) + '.reordered' + name.slice(dot));
    return sanitizeFileBaseName(name + '.reordered.setting');
  }

  function buildFmrDataLinkBlock(link, eol) {
    try {
      if (!link) return '';
      const json = JSON.stringify(link);
      const nl = eol || '\n';
      return `${nl}-- FMR_DATA_LINK_BEGIN${nl}-- ${json}${nl}-- FMR_DATA_LINK_END${nl}`;
    } catch (_) {
      return '';
    }
  }

  function stripFmrDataLinkBlock(text) {
    if (!text) return text;
    return text.replace(/\s*--\s*FMR_DATA_LINK_BEGIN[\s\S]*?--\s*FMR_DATA_LINK_END\s*/g, '\n');
  }

  const FMR_DATA_LINK_CONTROL = 'FMR_DataLink';
  const FMR_FILE_META_CONTROL = 'FMR_FileMeta';
  const FMR_PRESET_DATA_CONTROL = 'FMR_PresetData';
  const FMR_PRESET_SCRIPT_CONTROL = 'FMR_PresetScript';
  const FMR_PRESET_SELECT_CONTROL = 'FMR_Preset';
  // Diagnostic safety gate: constructor-style values (Polyline{}, FuID{}, etc.)
  // are the remaining likely blocker for Fusion load stability in preset OnChange.
  const FMR_PRESET_RUNTIME_INCLUDE_CONSTRUCTORS = false;
  // Temporary deep tracing for preset runtime path assignment behavior in Fusion console.
  const FMR_PRESET_RUNTIME_TRACE = false;
  // Hidden preset blob controls are parse-fragile in some Fusion builds.
  // Keep this off by default and drive runtime from direct selector script.
  const FMR_PRESET_EMBED_HIDDEN_BLOBS = false;
  // Keep hidden preset blobs well under Fusion's practical per-line parser limits.
  const FMR_HIDDEN_TEXT_CHUNK_SIZE = 900;
  const FMR_HIDDEN_TEXT_MAX_CHUNKS = 64;
  const FMR_PRESET_DIRECT_SCRIPT_MAX = 1200;
  const FMR_BUILD_MARKER = 'preset-debug-2026-02-28-a';

  function getChunkedControlId(baseId, chunkIndex) {
    const base = String(baseId || '').trim();
    const idx = Number(chunkIndex) || 1;
    if (!base || idx <= 1) return base;
    return `${base}_${idx}`;
  }

  function splitEscapedSettingIntoChunks(escapedValue, chunkSize = FMR_HIDDEN_TEXT_CHUNK_SIZE) {
    try {
      const src = String(escapedValue || '');
      const size = Math.max(256, Number(chunkSize) || FMR_HIDDEN_TEXT_CHUNK_SIZE);
      if (!src) return [];
      const out = [];
      let cursor = 0;
      while (cursor < src.length) {
        out.push(src.slice(cursor, cursor + size));
        cursor += size;
      }
      return out;
    } catch (_) {
      return [];
    }
  }

  function listChunkedControlIds(result, baseId) {
    try {
      const base = String(baseId || '').trim();
      if (!base) return [];
      const controls = Array.isArray(result?.groupUserControls) ? result.groupUserControls : [];
      const matches = controls
        .map((control) => String(control?.id || '').trim())
        .filter((id) => id && (id === base || id.startsWith(`${base}_`)));
      return Array.from(new Set(matches)).sort((a, b) => {
        const ai = a === base ? 1 : (Number(a.slice(base.length + 1)) || 9999);
        const bi = b === base ? 1 : (Number(b.slice(base.length + 1)) || 9999);
        return ai - bi;
      });
    } catch (_) {
      return [];
    }
  }

  function removeChunkedControlSeries(text, result, baseId) {
    try {
      let next = String(text || '');
      const base = String(baseId || '').trim();
      if (!base) return next;
      const known = listChunkedControlIds(result, base);
      if (known.length) {
        known.forEach((id) => {
          next = removeUserControlBlockById(next, result, id);
        });
        return next;
      }
      for (let i = 1; i <= FMR_HIDDEN_TEXT_MAX_CHUNKS; i += 1) {
        const id = getChunkedControlId(base, i);
        const updated = removeUserControlBlockById(next, result, id);
        if (updated === next && i > 1) break;
        next = updated;
      }
      return next;
    } catch (_) {
      return text;
    }
  }

  function collectChunkedControlDefaultValue(result, baseId) {
    try {
      const base = String(baseId || '').trim();
      if (!base) return '';
      const controls = Array.isArray(result?.groupUserControls) ? result.groupUserControls : [];
      const chunks = [];
      controls.forEach((control) => {
        const id = String(control?.id || '').trim();
        if (!id || (id !== base && !id.startsWith(`${base}_`))) return;
        const idx = id === base ? 1 : (Number(id.slice(base.length + 1)) || NaN);
        if (!Number.isFinite(idx) || idx < 1) return;
        const value = typeof control?.defaultValue === 'string' ? control.defaultValue : '';
        chunks.push({ idx, value });
      });
      if (!chunks.length) return '';
      chunks.sort((a, b) => a.idx - b.idx);
      return chunks.map((item) => item.value || '').join('');
    } catch (_) {
      return '';
    }
  }

  function extractHiddenDataLink(text, result) {
    try {
      const parsedLinkControl = getParsedGroupUserControlById(result, FMR_DATA_LINK_CONTROL);
      if (parsedLinkControl && typeof parsedLinkControl.defaultValue === 'string' && parsedLinkControl.defaultValue.trim()) {
        return JSON.parse(parsedLinkControl.defaultValue);
      }
      if (!text) return null;
      const body = getGroupUserControlBodyById(text, result, FMR_DATA_LINK_CONTROL);
      if (!body) return null;
      const raw = extractControlPropValue(body, 'INP_Default');
      if (!raw) return null;
      let json = raw.trim();
      if (json.startsWith('"') && json.endsWith('"')) {
        json = unescapeSettingString(json.slice(1, -1));
      }
      json = (json || '').trim();
      if (!json) return null;
      return JSON.parse(json);
    } catch (_) {
      return null;
    }
  }

  function extractHiddenDataLinkFallback(text) {
    try {
      if (!text) return null;
      const re = /FMR_DataLink\s*=\s*\{[\s\S]*?INP_Default\s*=\s*("([^"\\]|\\.)*")/g;
      let match = null;
      let last = null;
      // eslint-disable-next-line no-cond-assign
      while ((match = re.exec(text)) !== null) {
        last = match[1];
      }
      if (!last) return null;
      let json = last.trim();
      if (json.startsWith('"') && json.endsWith('"')) {
        json = unescapeSettingString(json.slice(1, -1));
      }
      json = (json || '').trim();
      if (!json) return null;
      return JSON.parse(json);
    } catch (_) {
      return null;
    }
  }

  function extractHiddenFileMeta(text, result) {
    try {
      const parsedMetaControl = getParsedGroupUserControlById(result, FMR_FILE_META_CONTROL);
      if (parsedMetaControl && typeof parsedMetaControl.defaultValue === 'string' && parsedMetaControl.defaultValue.trim()) {
        return JSON.parse(parsedMetaControl.defaultValue);
      }
      if (!text) return null;
      const body = getGroupUserControlBodyById(text, result, FMR_FILE_META_CONTROL);
      if (!body) return null;
      const raw = extractControlPropValue(body, 'INP_Default');
      if (!raw) return null;
      let json = raw.trim();
      if (json.startsWith('"') && json.endsWith('"')) {
        json = unescapeSettingString(json.slice(1, -1));
      }
      json = (json || '').trim();
      if (!json) return null;
      return JSON.parse(json);
    } catch (_) {
      return null;
    }
  }

  function extractHiddenPresetData(text, result) {
    try {
      const parsedValue = collectChunkedControlDefaultValue(result, FMR_PRESET_DATA_CONTROL);
      if (parsedValue && parsedValue.trim()) {
        return JSON.parse(parsedValue);
      }
      if (!text) return null;
      let assembled = '';
      for (let i = 1; i <= FMR_HIDDEN_TEXT_MAX_CHUNKS; i += 1) {
        const controlId = getChunkedControlId(FMR_PRESET_DATA_CONTROL, i);
        const body = getGroupUserControlBodyById(text, result, controlId);
        if (!body) {
          if (i === 1) return null;
          break;
        }
        const raw = extractControlPropValue(body, 'INP_Default');
        if (!raw) continue;
        let chunk = raw.trim();
        if (chunk.startsWith('"') && chunk.endsWith('"')) {
          chunk = unescapeSettingString(chunk.slice(1, -1));
        }
        assembled += String(chunk || '');
      }
      const json = String(assembled || '').trim();
      if (!json) return null;
      return JSON.parse(json);
    } catch (_) {
      return null;
    }
  }

  function upsertHiddenFileMetaControl(text, result, meta, eol) {
    try {
      const safeJson = meta ? `"${escapeSettingString(JSON.stringify(meta))}"` : '';
      return upsertGroupUserControl(text, result, FMR_FILE_META_CONTROL, {
        createLines: [
          'LINKS_Name = "FMR File Meta",',
          'INPID_InputControl = "TextEditControl",',
          'ICS_ControlPage = "Controls",',
          'IC_Visible = false,',
          `INP_Default = ${safeJson || '""'},`,
        ],
        updateBody: (body, indent, newline) => {
          let next = body;
          next = upsertControlProp(next, 'LINKS_Name', '"FMR File Meta"', indent, newline);
          next = upsertControlProp(next, 'INPID_InputControl', '"TextEditControl"', indent, newline);
          next = upsertControlProp(next, 'ICS_ControlPage', '"Controls"', indent, newline);
          next = upsertControlProp(next, 'IC_Visible', 'false', indent, newline);
          next = meta
            ? upsertControlProp(next, 'INP_Default', safeJson, indent, newline)
            : removeControlProp(next, 'INP_Default');
          return next;
        },
      }, eol || '\n');
    } catch (_) {
      return text;
    }
  }

  function upsertHiddenDataLinkControl(text, result, link, eol) {
    try {
      const controlId = FMR_DATA_LINK_CONTROL;
      const safeJson = link ? `"${escapeSettingString(JSON.stringify(link))}"` : '';
      return upsertGroupUserControl(text, result, controlId, {
        createLines: [
          'LINKS_Name = "FMR Data Link",',
          'INPID_InputControl = "TextEditControl",',
          'ICS_ControlPage = "Controls",',
          'IC_Visible = false,',
          `INP_Default = ${safeJson || '""'},`,
        ],
        updateBody: (body, indent, newline) => {
          let next = body;
          next = upsertControlProp(next, 'LINKS_Name', '"FMR Data Link"', indent, newline);
          next = upsertControlProp(next, 'INPID_InputControl', '"TextEditControl"', indent, newline);
          next = upsertControlProp(next, 'ICS_ControlPage', '"Controls"', indent, newline);
          next = upsertControlProp(next, 'IC_Visible', 'false', indent, newline);
          next = link
            ? upsertControlProp(next, 'INP_Default', safeJson, indent, newline)
            : removeControlProp(next, 'INP_Default');
          return next;
        },
      }, eol || '\n');
    } catch (_) {
      return text;
    }
  }

  function upsertHiddenPresetDataControl(text, result, data, eol) {
    try {
      let next = String(text || '');
      const encoded = data ? escapeSettingString(JSON.stringify(data)) : '';
      const chunks = splitEscapedSettingIntoChunks(encoded, FMR_HIDDEN_TEXT_CHUNK_SIZE);
      if (!chunks.length) {
        return removeChunkedControlSeries(next, result, FMR_PRESET_DATA_CONTROL);
      }
      chunks.slice(0, FMR_HIDDEN_TEXT_MAX_CHUNKS).forEach((chunk, index) => {
        const chunkIdx = index + 1;
        const controlId = getChunkedControlId(FMR_PRESET_DATA_CONTROL, chunkIdx);
        const label = chunkIdx === 1 ? 'FMR Preset Data' : `FMR Preset Data ${chunkIdx}`;
        const defaultText = `"${String(chunk || '')}"`;
        next = upsertGroupUserControl(next, result, controlId, {
          createLines: [
            `LINKS_Name = "${label}",`,
            'INPID_InputControl = "TextEditControl",',
            'LINKID_DataType = "Text",',
            'ICS_ControlPage = "Controls",',
            'IC_Visible = false,',
            `INP_Default = ${defaultText},`,
          ],
          updateBody: (body, indent, newline) => {
            let bodyNext = body;
            bodyNext = upsertControlProp(bodyNext, 'LINKS_Name', `"${label}"`, indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'INPID_InputControl', '"TextEditControl"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'LINKID_DataType', '"Text"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'ICS_ControlPage', '"Controls"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'IC_Visible', 'false', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'INP_Default', defaultText, indent, newline);
            return bodyNext;
          },
        }, eol || '\n');
      });
      for (let i = chunks.length + 1; i <= FMR_HIDDEN_TEXT_MAX_CHUNKS; i += 1) {
        next = removeUserControlBlockById(next, result, getChunkedControlId(FMR_PRESET_DATA_CONTROL, i));
      }
      return next;
    } catch (_) {
      return text;
    }
  }

  function upsertHiddenPresetScriptControl(text, result, script, eol) {
    try {
      let next = String(text || '');
      const encoded = script && script.trim() ? escapeSettingString(script) : '';
      const chunks = splitEscapedSettingIntoChunks(encoded, FMR_HIDDEN_TEXT_CHUNK_SIZE);
      if (!chunks.length) {
        return removeChunkedControlSeries(next, result, FMR_PRESET_SCRIPT_CONTROL);
      }
      chunks.slice(0, FMR_HIDDEN_TEXT_MAX_CHUNKS).forEach((chunk, index) => {
        const chunkIdx = index + 1;
        const controlId = getChunkedControlId(FMR_PRESET_SCRIPT_CONTROL, chunkIdx);
        const label = chunkIdx === 1 ? 'FMR Preset Script' : `FMR Preset Script ${chunkIdx}`;
        const defaultText = `"${String(chunk || '')}"`;
        next = upsertGroupUserControl(next, result, controlId, {
          createLines: [
            `LINKS_Name = "${label}",`,
            'INPID_InputControl = "TextEditControl",',
            'LINKID_DataType = "Text",',
            'ICS_ControlPage = "Controls",',
            'IC_Visible = false,',
            `INP_Default = ${defaultText},`,
          ],
          updateBody: (body, indent, newline) => {
            let bodyNext = body;
            bodyNext = upsertControlProp(bodyNext, 'LINKS_Name', `"${label}"`, indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'INPID_InputControl', '"TextEditControl"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'LINKID_DataType', '"Text"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'ICS_ControlPage', '"Controls"', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'IC_Visible', 'false', indent, newline);
            bodyNext = upsertControlProp(bodyNext, 'INP_Default', defaultText, indent, newline);
            return bodyNext;
          },
        }, eol || '\n');
      });
      for (let i = chunks.length + 1; i <= FMR_HIDDEN_TEXT_MAX_CHUNKS; i += 1) {
        next = removeUserControlBlockById(next, result, getChunkedControlId(FMR_PRESET_SCRIPT_CONTROL, i));
      }
      return next;
    } catch (_) {
      return text;
    }
  }

  function buildPresetScriptLauncherScript() {
    try {
      const lines = [];
      lines.push(`local src = tostring(tool:GetInput("${escapeLuaString(FMR_PRESET_SCRIPT_CONTROL)}") or "")`);
      lines.push(`for i = 2, ${Math.max(2, Number(FMR_HIDDEN_TEXT_MAX_CHUNKS) || 64)} do`);
      lines.push(`  local part = tool:GetInput("${escapeLuaString(FMR_PRESET_SCRIPT_CONTROL)}_" .. i)`);
      lines.push('  if part == nil then break end');
      lines.push('  src = src .. tostring(part)');
      lines.push('end');
      lines.push('if src ~= "" then');
      lines.push('  local compile = loadstring or load');
      lines.push('  if type(compile) ~= "function" then return end');
      lines.push('  local okCompile, fnOrErr = pcall(compile, src)');
      lines.push('  if not okCompile then');
      lines.push('    print(fnOrErr)');
      lines.push('  elseif type(fnOrErr) == "function" then');
      lines.push('    local okRun, runErr = pcall(fnOrErr)');
      lines.push('    if not okRun then print(runErr) end');
      lines.push('  elseif fnOrErr then');
      lines.push('    print(fnOrErr)');
      lines.push('  end');
      lines.push('end');
      return lines.join('\n');
    } catch (_) {
      return '';
    }
  }

  function removeUserControlBlockById(text, result, controlId) {
    return removeGroupUserControlById(text, result, controlId);
  }

  function removePresetRuntimeControls(text, result) {
    try {
      let next = String(text || '');
      next = removeUserControlBlockById(next, result, FMR_PRESET_SELECT_CONTROL);
      next = removeChunkedControlSeries(next, result, FMR_PRESET_DATA_CONTROL);
      next = removeChunkedControlSeries(next, result, FMR_PRESET_SCRIPT_CONTROL);
      return next;
    } catch (_) {
      return text;
    }
  }

  function buildFmrDataLinkPayload(result) {
    try {
      const csv = result?.csvData || state.csvData || null;
      const source = result?.dataLink?.source || csv?.sourceName || '';
      const hasMappings = Array.isArray(result?.entries) && result.entries.some(e => e?.csvLink?.column || e?.csvLinkHex?.column);
      if (!source || !hasMappings) return null;
      const mappings = (result?.dataLink && result.dataLink.mappings)
        ? { ...result.dataLink.mappings }
        : {};
      const overrides = (result?.dataLink && result.dataLink.overrides)
        ? { ...result.dataLink.overrides }
        : {};
      const hex = (result?.dataLink && result.dataLink.hex)
        ? { ...result.dataLink.hex }
        : {};
      result.entries.forEach((entry) => {
        if (!entry || !entry.sourceOp || !entry.source) return;
        const key = `${entry.sourceOp}.${entry.source}`;
        if (entry.csvLink && entry.csvLink.column) {
          mappings[key] = entry.csvLink.column;
          if (entry.csvLink.rowMode && entry.csvLink.rowMode !== 'default') {
            overrides[key] = {
              rowMode: entry.csvLink.rowMode,
              rowIndex: Number.isFinite(entry.csvLink.rowIndex) ? entry.csvLink.rowIndex : null,
              rowKey: entry.csvLink.rowKey || null,
              rowValue: entry.csvLink.rowValue || null,
            };
          } else if (overrides[key]) {
            delete overrides[key];
          }
        }
        if (entry.csvLinkHex && entry.csvLinkHex.column) {
          hex[key] = {
            column: entry.csvLinkHex.column,
            rowMode: entry.csvLinkHex.rowMode || null,
            rowIndex: Number.isFinite(entry.csvLinkHex.rowIndex) ? entry.csvLinkHex.rowIndex : null,
            rowKey: entry.csvLinkHex.rowKey || null,
            rowValue: entry.csvLinkHex.rowValue || null,
          };
        } else if (hex[key]) {
          delete hex[key];
        }
      });
      if (!Object.keys(mappings).length && !Object.keys(hex).length) return null;
      const link = result?.dataLink || {};
      return {
        version: 1,
        source,
        nameColumn: link.nameColumn || csv?.nameColumn || null,
        rowMode: link.rowMode || 'first',
        rowIndex: Number.isFinite(link.rowIndex) ? link.rowIndex : null,
        rowKey: link.rowKey || null,
        rowValue: link.rowValue || null,
        mappings,
        overrides: Object.keys(overrides).length ? overrides : undefined,
        hex: Object.keys(hex).length ? hex : undefined,
      };
    } catch (_) {
      return null;
    }
  }

  function syncDataLinkMappings(result) {
    if (!result) return;
    const link = ensureDataLink(result);
    if (!link) return;
    const mappings = {};
    const overrides = {};
    const hex = {};
    (result.entries || []).forEach((entry) => {
      if (!entry || !entry.csvLink || !entry.csvLink.column || !entry.sourceOp || !entry.source) return;
      const key = `${entry.sourceOp}.${entry.source}`;
      mappings[key] = entry.csvLink.column;
      if (entry.csvLink.rowMode && entry.csvLink.rowMode !== 'default') {
        overrides[key] = {
          rowMode: entry.csvLink.rowMode,
          rowIndex: Number.isFinite(entry.csvLink.rowIndex) ? entry.csvLink.rowIndex : null,
          rowKey: entry.csvLink.rowKey || null,
          rowValue: entry.csvLink.rowValue || null,
        };
      }
    });
    (result.entries || []).forEach((entry) => {
      if (!entry || !entry.csvLinkHex || !entry.csvLinkHex.column || !entry.sourceOp || !entry.source) return;
      const key = `${entry.sourceOp}.${entry.source}`;
      hex[key] = {
        column: entry.csvLinkHex.column,
        rowMode: entry.csvLinkHex.rowMode || null,
        rowIndex: Number.isFinite(entry.csvLinkHex.rowIndex) ? entry.csvLinkHex.rowIndex : null,
        rowKey: entry.csvLinkHex.rowKey || null,
        rowValue: entry.csvLinkHex.rowValue || null,
      };
    });
    link.mappings = mappings;
    if (Object.keys(overrides).length) link.overrides = overrides;
    else delete link.overrides;
    if (Object.keys(hex).length) link.hex = hex;
    else delete link.hex;
  }

  function sanitizeFileBaseName(value) {
    return String(value || '')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/\s+/g, '_')
      .trim();
  }

  function buildMacroExportName() {
    const macroName = resolveExportMacroName(state.parseResult, state.originalFilePath || state.originalFileName);
    if (macroName) {
      const safe = sanitizeFileBaseName(macroName);
      return (safe || 'Macro') + '.setting';
    }
    return suggestOutputName(state.originalFileName || 'Macro.setting');
  }

  function buildMacroExportNameForSnapshot(snap) {
    if (snap?.csvGenerated && snap?.originalFileName) {
      return sanitizeFileBaseName(String(snap.originalFileName));
    }
    const macroName = (() => {
      const currentName = cleanDocLabelValue(snap?.parseResult?.macroName || '');
      const originalName = cleanDocLabelValue(snap?.parseResult?.macroNameOriginal || '');
      if (currentName && !isPlaceholderMacroName(currentName)) return currentName;
      if (!currentName && originalName && !isPlaceholderMacroName(originalName)) return originalName;
      return deriveMacroNameFromFileName(snap?.originalFilePath || snap?.originalFileName || '') || currentName || originalName || '';
    })();
    if (macroName) {
      const safe = sanitizeFileBaseName(macroName);
      return (safe || 'Macro') + '.setting';
    }
    return suggestOutputName(snap?.originalFileName || 'Macro.setting');
  }

  function buildUniqueCsvNames(rows, nameColumn, baseName) {
    const names = [];
    const counts = new Map();
    for (let i = 0; i < rows.length; i += 1) {
      const row = rows[i] || {};
      let label = '';
      if (nameColumn === '__row__') {
        label = `${baseName} ${i + 1}`;
      } else {
        const raw = String(row[nameColumn] || '').trim();
        label = raw || `${baseName} ${i + 1}`;
      }
      const key = label.toLowerCase();
      const count = (counts.get(key) || 0) + 1;
      counts.set(key, count);
      names.push(count > 1 ? `${label}_${count}` : label);
    }
    return names;
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



  function rebuildContentWithNewOrder(original, result, eol, options = {}) {
    const { entries } = result;
    try { synchronizeAutoQuickSetLabelEntries(result, original); } catch (_) {}
    try { hydrateFlipPairGroups(result, eol || '\n'); } catch (_) {}
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
    const baseText = applyToolRenameMap(original, result);
    let updated = baseText;
    const dataLink = (options.includeDataLink !== false) ? buildFmrDataLinkPayload(result) : null;
    const presetRuntime = options.includePresetRuntime === true
      ? buildPresetRuntimeState(result)
      : null;
    const hasUpdateDataButton = !!(result && Array.isArray(result.entries) && result.entries.some((entry) => isUpdateDataButtonEntry(entry)));
    const fileMetaPath = options.fileMetaPath
      || result?.fileMeta?.exportPath
      || state.originalFilePath
      || (state.exportFolder ? buildExportPath(state.exportFolder, buildMacroExportName()) : '');
    const exportMacroName = resolveExportMacroName(result, fileMetaPath);
    const shouldPersistFileMeta = !!(fileMetaPath && (dataLink || hasUpdateDataButton));
    const fileMeta = shouldPersistFileMeta ? { exportPath: fileMetaPath, buildMarker: FMR_BUILD_MARKER } : null;
    // Clean exports for authored macros should preserve existing tool/user-control
    // blocks verbatim unless MM is explicitly managing a helper system that needs
    // to rewrite them (for example data-link/update-button wiring).
    const preserveUntouchedUserControls = options.preserveAuthoredUserControls === true && !dataLink && !hasUpdateDataButton;
    const refreshGroupUserControlState = (textValue) => {
      try {
        hydrateGroupUserControlsState(textValue, result);
      } catch (_) {}
      return textValue;
    };
    if (fileMetaPath && result && Array.isArray(result.entries)) {
      const overrides = result.buttonOverrides instanceof Map
        ? new Map(result.buttonOverrides)
        : (result.buttonOverrides ? new Map(result.buttonOverrides) : new Map());
      const url = buildUpdateDataUrl(fileMetaPath);
      result.entries.forEach((entry) => {
        if (!entry || !entry.sourceOp || !entry.source) return;
        if (!isUpdateDataButtonEntry(entry)) return;
        overrides.set(`${entry.sourceOp}.${entry.source}`, url);
      });
      if (overrides.size) result.buttonOverrides = overrides;
    }
    const safeEditExport = options.safeEditExport === true;
    if (!safeEditExport && !preserveUntouchedUserControls) {
      updated = applyLabelCountEdits(updated, result, eol);
      updated = applyLabelVisibilityEdits(updated, result, eol);
      updated = applyLabelDefaultStateEdits(updated, result, eol);
      updated = applyUserControlPages(updated, result, eol, options);
      updated = ensureQuickSetLabelsMaterializedInGroupUserControls(updated, result, eol);
      // Safety: strip BTNCS_Execute from managed buttons unless explicitly clicked
      updated = stripUnclickedBtncs(updated, result, eol);
      // Insert exact launcher only for explicitly clicked controls
      updated = ensureExactLauncherInserted(updated, result, eol);
      updated = stripLegacyLauncherArtifacts(updated);
      updated = applyCanonicalLaunchSnippets(updated, result, eol);
      updated = rewriteUpdateDataProtocolPaths(updated, fileMetaPath);
    }
    updated = applyMacroNameRename(updated, result, exportMacroName);
    updated = ensureDefaultExtentSetOnGeneratorTools(updated, result, eol);
    updated = normalizeGroupUserControlsBlocks(updated, result, eol);
    refreshGroupUserControlState(updated);
    updated = ensureGroupInputsBlock(updated, result, eol);
    updated = rewritePrimaryInputsBlock(updated, result, eol, options);
    updated = ensureRoutingSourceInputsMaterialized(updated, result, eol);
    updated = rewriteMacroRoutingInputsBlock(updated, result, eol, options);
    updated = applyPathDrawModeExportPatches(updated, result, eol, options);
    if (options.includePresetRuntime === true) {
      updated = applyPresetPathSwitchExportPatches(updated, result, presetRuntime, eol, options);
    }
    updated = ensureActiveToolName(updated, exportMacroName || (result?.macroName || result?.macroNameOriginal || '').trim(), eol);
    updated = normalizeGroupUserControlsBlocks(updated, result, eol);
    refreshGroupUserControlState(updated);
    if (options.includePresetRuntime === true) {
      updated = removePresetRuntimeControls(updated, result);
      refreshGroupUserControlState(updated);
      if (presetRuntime) {
        const runtimeScript = String(presetRuntime.script || '');
        if (FMR_PRESET_EMBED_HIDDEN_BLOBS) {
          updated = upsertHiddenPresetDataControl(updated, result, presetRuntime.payload, eol);
          refreshGroupUserControlState(updated);
          const useScriptLauncher = runtimeScript.length > FMR_PRESET_DIRECT_SCRIPT_MAX;
          if (useScriptLauncher) {
            updated = upsertHiddenPresetScriptControl(updated, result, runtimeScript, eol);
            refreshGroupUserControlState(updated);
          } else {
            updated = upsertHiddenPresetScriptControl(updated, result, '', eol);
            refreshGroupUserControlState(updated);
          }
          updated = upsertPresetSelectorControl(
            updated,
            result,
            presetRuntime.presetNames,
            useScriptLauncher ? buildPresetScriptLauncherScript() : runtimeScript,
            eol,
            presetRuntime.defaultIndex,
          );
        } else {
          updated = upsertPresetSelectorControl(
            updated,
            result,
            presetRuntime.presetNames,
            runtimeScript,
            eol,
            presetRuntime.defaultIndex,
          );
        }
        refreshGroupUserControlState(updated);
      }
    }
    if (options.includeDataLink !== false && dataLink) {
      updated = stripFmrDataLinkBlock(updated);
      updated = upsertHiddenDataLinkControl(updated, result, dataLink, eol);
      refreshGroupUserControlState(updated);
    }
    if (fileMeta) {
      updated = upsertHiddenFileMetaControl(updated, result, fileMeta, eol);
      refreshGroupUserControlState(updated);
    }
    updated = normalizeGroupUserControlsBlocks(updated, result, eol);
    updated = stripEmptyGroupUserControls(updated, result);
    refreshGroupUserControlState(updated);
    const exportStructure = validateGroupUserControlsStructure(updated, result);
    try {
      if (typeof logDiag === 'function') {
        const emptyUcMatches = updated.match(/\n\s*UserControls\s*=\s*ordered\(\)\s*\{\s*\},/g);
        const emptyUcCount = Array.isArray(emptyUcMatches) ? emptyUcMatches.length : 0;
        const hasPresetSelector = updated.includes(`${FMR_PRESET_SELECT_CONTROL} = {`);
        const hasBuildMarker = updated.includes(FMR_BUILD_MARKER);
        const hasFileMetaMarker = updated.includes(`"buildMarker":"${FMR_BUILD_MARKER}"`);
        const engine = result?.presetEngine && typeof result.presetEngine === 'object' ? result.presetEngine : null;
        const scopeLen = Array.isArray(engine?.scope) ? engine.scope.length : 0;
        const presetLen = engine?.presets && typeof engine.presets === 'object'
          ? Object.keys(engine.presets).length
          : 0;
        const debugContext = String(options?.debugContext || 'generic');
        const hasEngine = engine ? 1 : 0;
        const hasFileMetaPath = fileMetaPath ? 1 : 0;
        const groupUcBlocks = Number(exportStructure?.parsed?.blockCount || 0);
        const groupUcControls = Array.isArray(exportStructure?.parsed?.controls) ? exportStructure.parsed.controls.length : 0;
        const groupUcDupes = Array.isArray(exportStructure?.parsed?.duplicateIds) ? exportStructure.parsed.duplicateIds.length : 0;
        const groupUcUnknown = Array.isArray(exportStructure?.parsed?.controls)
          ? exportStructure.parsed.controls.filter((control) => control?.kind === 'unknown').length
          : 0;
        const groupUcIssues = Array.isArray(exportStructure?.issues) ? exportStructure.issues.length : 0;
        const groupUcMutationIssues = Array.isArray(result?.groupUserControlMutationIssues) ? result.groupUserControlMutationIssues.length : 0;
        const quickSetEntryIds = Array.isArray(result?.entries)
          ? result.entries
              .map((entry) => String(entry?.source || '').trim())
              .filter((id) => /^MM_QuickSetLabel_/i.test(id))
          : [];
        const quickSetEntryUnique = new Set(quickSetEntryIds);
        const quickSetOutputMatches = updated.match(/(^|\n|\r)\s*(MM_QuickSetLabel_[A-Za-z0-9_]+)\s*=\s*\{/g) || [];
        const quickSetOutputUnique = new Set(
          quickSetOutputMatches
            .map((chunk) => {
              const m = String(chunk || '').match(/MM_QuickSetLabel_[A-Za-z0-9_]+/);
              return m ? m[0] : '';
            })
            .filter(Boolean)
        );
        const pathSwitchCount = Array.isArray(presetRuntime?.pathSwitchPlan) ? presetRuntime.pathSwitchPlan.length : 0;
        logDiag(`[Export debug:${debugContext}] engine=${hasEngine}, scope=${scopeLen}, presets=${presetLen}, fileMetaPath=${hasFileMetaPath}, marker=${hasBuildMarker ? 1 : 0}, fileMetaMarker=${hasFileMetaMarker ? 1 : 0}, presetControl=${hasPresetSelector ? 1 : 0}, emptyGroupUc=${emptyUcCount}, groupUcBlocks=${groupUcBlocks}, groupUcControls=${groupUcControls}, groupUcDupes=${groupUcDupes}, groupUcUnknown=${groupUcUnknown}, groupUcIssues=${groupUcIssues}, groupUcMutationIssues=${groupUcMutationIssues}, quickSetEntries=${quickSetEntryUnique.size}, quickSetOutput=${quickSetOutputUnique.size}, pathSwitch=${pathSwitchCount}`);
        (exportStructure?.issues || []).forEach((issue) => {
          logDiag(`[Export structure] ${issue.message}`);
        });
        (result?.groupUserControlMutationIssues || []).forEach((issue) => {
          logDiag(`[Export structure] ${issue.message}`);
        });
      }
    } catch (_) {}
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

  function ensureDefaultExtentSetOnGeneratorTools(text, result, eol) {
    try {
      if (!text || !result) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const toolsBlock = findOrderedBlock(text, bounds.groupOpenIndex, bounds.groupCloseIndex, 'Tools');
      if (!toolsBlock) return text;
      const entries = parseOrderedBlockEntries(text, toolsBlock.open, toolsBlock.close);
      if (!Array.isArray(entries) || !entries.length) return text;
      const eligibleTypes = new Set(['background', 'fastnoise', 'srender']);
      const newline = eol || detectNewline(text) || '\n';
      let updated = text;
      const reversedEntries = entries.slice().reverse();
      reversedEntries.forEach((entry) => {
        if (!entry || !Number.isFinite(entry.blockOpen) || !Number.isFinite(entry.blockClose)) return;
        const toolType = String(getOrderedBlockEntryType(updated, entry) || '').trim().toLowerCase();
        if (!eligibleTypes.has(toolType)) return;
        const body = updated.slice(entry.blockOpen + 1, entry.blockClose);
        if (/\bExtentSet\s*=/.test(body)) return;
        const inputsBlock = findInputsInTool(updated, entry.blockOpen, entry.blockClose);
        if (!inputsBlock) return;
        const hasGlobalIn = !!findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, 'GlobalIn');
        const hasGlobalOut = !!findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, 'GlobalOut');
        if (hasGlobalIn || hasGlobalOut) return;
        const indent = (getLineIndent(updated, entry.blockOpen) || '') + '\t';
        const insert = `${newline}${indent}ExtentSet = true,`;
        updated = updated.slice(0, entry.blockOpen + 1) + insert + updated.slice(entry.blockOpen + 1);
      });
      return updated;
    } catch (_) {
      return text;
    }
  }

  function rewritePrimaryInputsBlock(text, result, eol, options = {}) {
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
      const existingEntries = parseOrderedBlockEntries(text, braceOpen, braceClose) || [];
      // Some authored macros intentionally mix helper Input blocks with normal
      // published InstanceInputs. For untouched docs, preserve that block
      // verbatim instead of canonicalizing it into MM's generated format.
      const preserveAuthoredInputsBlock = options?.preserveAuthoredInputsBlock === true;
      if (preserveAuthoredInputsBlock && existingEntries.some((entry) => getOrderedBlockEntryType(text, entry) !== 'InstanceInput')) {
        return text;
      }
      const segments = [];
      exportOrder.forEach(i => {
        const entry = result.entries[i];
        if (!entry) return;
        const sourceId = String(entry.source || '').trim();
        const isAutoQuickSetLabel = /^MM_QuickSetLabel_/i.test(sourceId);
        if (entry.syntheticUserControlOnly || entry.source === FMR_PRESET_SELECT_CONTROL) return;
        if (isGroupUserControlEntry(entry) && !isAutoQuickSetLabel) return;
        let raw = applyEntryOverrides(entry, applyNameIfEdited(entry, eol), eol);
        raw = ensureInstanceInputKey(raw, entry);
        let chunk = reindent(raw, entryIndent, eol);
        if (!chunk || !chunk.trim().length) return;
        chunk = ensureTrailingComma(chunk);
        segments.push({ index: i, chunk });
      });
      // Preserve the original Inputs block verbatim when the published order and
      // effective InstanceInput payloads are unchanged. This avoids no-op export
      // churn from purely formatting-driven rewrites.
      if (canPreservePrimaryInputsBlock(text, result, segments, newline)) {
        return text;
      }
      let outputSegments = segments.slice();
      if (existingEntries.length) {
        const generatedQueue = segments.slice();
        outputSegments = [];
        existingEntries.forEach((current) => {
          const type = getOrderedBlockEntryType(text, current);
          if (type === 'InstanceInput') {
            const nextGenerated = generatedQueue.shift();
            if (nextGenerated) outputSegments.push(nextGenerated);
            return;
          }
          const preservedRaw = extractOrderedBlockEntryChunk(text, current);
          const preservedChunk = ensureTrailingComma(reindent(preservedRaw, entryIndent, eol));
          if (preservedChunk && preservedChunk.trim().length) {
            outputSegments.push({ index: null, chunk: preservedChunk });
          }
        });
        while (generatedQueue.length) outputSegments.push(generatedQueue.shift());
      }
      const header = `${blockIndent}Inputs = ordered() {`;
      const inner = outputSegments.length ? newline + newline + outputSegments.map(seg => seg.chunk).join(newline) + newline + innerIndent : '';
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

  function extractRoutingInputFromEntry(text, entry) {
    try {
      if (!text || !entry || !Number.isFinite(entry.blockOpen) || !Number.isFinite(entry.blockClose)) return null;
      const body = text.slice(entry.blockOpen + 1, entry.blockClose);
      const sourceOp = unquoteSettingValue(extractControlPropValue(body, 'SourceOp'));
      const source = unquoteSettingValue(extractControlPropValue(body, 'Source'));
      if (!sourceOp || !source) return null;
      return {
        sourceOp: String(sourceOp).trim(),
        source: String(source).trim(),
      };
    } catch (_) {
      return null;
    }
  }

  function buildRoutingInputChunk(row, entryIndent, newline) {
    const sourceOp = escapeQuotes(String(row?.sourceOp || '').trim());
    const source = escapeQuotes(String(row?.source || '').trim());
    const macroInput = sanitizeRoutingInputId(row?.macroInput || '');
    if (!macroInput || !sourceOp || !source) return '';
    return [
      `${entryIndent}${macroInput} = InstanceInput {`,
      `${entryIndent}\tSourceOp = "${sourceOp}",`,
      `${entryIndent}\tSource = "${source}",`,
      `${entryIndent}},`,
    ].filter(Boolean).join(newline);
  }

  function insertToolInputStubBlock(text, inputsOpen, inputsClose, controlName, eol) {
    try {
      if (!text || !controlName) return text;
      const nl = eol || '\n';
      const itemIndent = (getLineIndent(text, inputsOpen) || '') + '\t';
      const inputId = formatToolInputId(controlName);
      if (!inputId) return text;
      const block = `${itemIndent}${inputId} = Input {  },`;
      let cursor = Math.max(inputsOpen + 1, inputsClose - 1);
      while (cursor > inputsOpen && /\s/.test(text[cursor])) cursor -= 1;
      const prev = text[cursor];
      const isEmptyInputs = prev === '{';
      const needsSeparatorComma = !isEmptyInputs && prev !== ',';
      const needsLeadingNl = text[inputsClose - 1] !== '\n' && text[inputsClose - 1] !== '\r';
      const prefix = needsLeadingNl ? nl : '';
      const separator = needsSeparatorComma ? ',' : '';
      return text.slice(0, inputsClose) + separator + prefix + block + nl + text.slice(inputsClose);
    } catch (_) {
      return text;
    }
  }

  function ensureInputsBlockInToolBlock(text, toolBlock, eol) {
    try {
      if (!toolBlock) return { text, toolBlock: null, inputsBlock: null };
      const newline = eol || '\n';
      const existing = findInputsInTool(text, toolBlock.open, toolBlock.close);
      if (existing) return { text, toolBlock, inputsBlock: existing };
      const indent = (getLineIndent(text, toolBlock.open) || '') + '\t';
      const insertPos = toolBlock.close;
      let insertPrefix = '';
      let probe = insertPos - 1;
      while (probe >= toolBlock.open && /\s/.test(text[probe])) probe--;
      if (probe >= toolBlock.open && text[probe] !== '{' && text[probe] !== ',') {
        insertPrefix = ',';
      }
      const snippet = `${insertPrefix}${newline}${indent}Inputs = {${newline}${indent}},${newline}`;
      const updated = text.slice(0, insertPos) + snippet + text.slice(insertPos);
      const open = insertPos + snippet.indexOf('{');
      const close = insertPos + snippet.lastIndexOf('}');
      const delta = snippet.length;
      return {
        text: updated,
        toolBlock: { open: toolBlock.open, close: toolBlock.close + delta },
        inputsBlock: { open, close },
      };
    } catch (_) {
      return { text, toolBlock: null, inputsBlock: null };
    }
  }

  function ensureRoutingSourceInputsMaterialized(text, result, eol) {
    try {
      if (!text || !result) return text;
      const routingState = ensureRoutingInputsState(result);
      const rows = normalizeRoutingInputRows(routingState.rows || [], { allowEmptyMacroInput: false });
      if (!Array.isArray(rows) || !rows.length) return text;
      let updated = text;
      let inserted = 0;
      rows.forEach((row) => {
        try {
          const sourceOp = String(row?.sourceOp || '').trim();
          const source = String(row?.source || '').trim();
          if (!sourceOp || !source) return;
          const bounds = locateMacroGroupBounds(updated, result);
          if (!bounds) return;
          let toolBlock = findToolBlockInGroup(updated, bounds.groupOpenIndex, bounds.groupCloseIndex, sourceOp);
          if (!toolBlock) return;
          const ensured = ensureInputsBlockInToolBlock(updated, toolBlock, eol || '\n');
          updated = ensured.text;
          toolBlock = ensured.toolBlock || toolBlock;
          const inputsBlock = ensured.inputsBlock || findInputsInTool(updated, toolBlock.open, toolBlock.close);
          if (!inputsBlock) return;
          const exists = findInputBlockInInputs(updated, inputsBlock.open, inputsBlock.close, source);
          if (exists) return;
          updated = insertToolInputStubBlock(updated, inputsBlock.open, inputsBlock.close, source, eol || '\n');
          inserted += 1;
        } catch (_) {}
      });
      if (typeof logDiag === 'function' && diagnosticsController?.isEnabled?.()) {
        try { logDiag(`[Routing Inputs] source-stubs inserted=${inserted}`); } catch (_) {}
      }
      return updated;
    } catch (_) {
      return text;
    }
  }

  function rewriteMacroRoutingInputsBlock(text, result, eol, options = {}) {
    try {
      if (!text || !result) return text;
      const routingState = ensureRoutingInputsState(result);
      const desiredRows = Array.isArray(routingState.rows) ? routingState.rows : [];
      const managedNames = routingState.managed instanceof Set ? routingState.managed : new Set();
      if (!desiredRows.length && !managedNames.size) return text;

      const block = getPrimaryGroupInputsBlock(text, result);
      if (!block || !Number.isFinite(block.inputsHeaderStart) || !Number.isFinite(block.openIndex) || !Number.isFinite(block.closeIndex)) return text;
      const idx = block.inputsHeaderStart;
      const braceOpen = block.openIndex;
      const braceClose = block.closeIndex;

      let after = braceClose + 1;
      let trailingWhitespace = '';
      while (after < text.length && /\s/.test(text[after])) {
        trailingWhitespace += text[after];
        after += 1;
      }
      let hasComma = false;
      if (text[after] === ',') {
        hasComma = true;
        after += 1;
        while (after < text.length && /\s/.test(text[after])) {
          trailingWhitespace += text[after];
          after += 1;
        }
      }

      const existingEntries = parseOrderedBlockEntries(text, braceOpen, braceClose) || [];
      const existingInputByName = new Map();
      existingEntries.forEach((entry) => {
        if (!entry || !entry.name) return;
        const type = getOrderedBlockEntryType(text, entry);
        if (type !== 'Input' && type !== 'InstanceInput') return;
        const routing = extractRoutingInputFromEntry(text, entry);
        if (!routing) return;
        existingInputByName.set(entry.name, { entry, routing });
      });
      const desiredByName = new Map();
      desiredRows.forEach((row) => {
        const key = sanitizeRoutingInputId(row?.macroInput || '');
        if (!key) return;
        desiredByName.set(key, {
          macroInput: key,
          sourceOp: String(row?.sourceOp || '').trim(),
          source: String(row?.source || '').trim(),
        });
      });
      if (!desiredByName.size && !managedNames.size) return text;

      let needsRewrite = false;
      managedNames.forEach((name) => {
        const key = sanitizeRoutingInputId(name);
        if (!key || desiredByName.has(key)) return;
        if (existingInputByName.has(key)) needsRewrite = true;
      });
      if (!needsRewrite) {
        desiredByName.forEach((desired, name) => {
          const existingWrap = existingInputByName.get(name);
          if (!existingWrap || !existingWrap.entry) {
            needsRewrite = true;
            return;
          }
          const current = existingWrap.routing || extractRoutingInputFromEntry(text, existingWrap.entry);
          if (!current) {
            needsRewrite = true;
            return;
          }
          if (String(current.sourceOp) !== String(desired.sourceOp) || String(current.source) !== String(desired.source)) {
            needsRewrite = true;
          }
        });
      }
      if (!needsRewrite) return text;

      const blockIndent = getLineIndent(text, idx) || '';
      const innerIndent = blockIndent + '\t';
      const entryIndent = innerIndent + '\t';
      const newline = eol || detectNewline(text) || '\n';
      const outputSegments = [];
      const consumed = new Set();

      existingEntries.forEach((entry) => {
        if (!entry) return;
        const type = getOrderedBlockEntryType(text, entry);
        if (type === 'Input' || type === 'InstanceInput') {
          const routing = extractRoutingInputFromEntry(text, entry);
          const key = sanitizeRoutingInputId(entry.name);
          if (routing && key && desiredByName.has(key)) {
            const chunk = buildRoutingInputChunk(desiredByName.get(key), entryIndent, newline);
            if (chunk) outputSegments.push({ index: null, chunk });
            consumed.add(key);
            return;
          }
          if (routing && key && managedNames.has(key)) {
            return;
          }
        }
        const preservedRaw = extractOrderedBlockEntryChunk(text, entry);
        const preservedChunk = ensureTrailingComma(reindent(preservedRaw, entryIndent, eol));
        if (preservedChunk && preservedChunk.trim().length) {
          outputSegments.push({ index: null, chunk: preservedChunk });
        }
      });

      desiredByName.forEach((row, key) => {
        if (consumed.has(key)) return;
        const chunk = buildRoutingInputChunk(row, entryIndent, newline);
        if (chunk) outputSegments.push({ index: null, chunk });
      });

      const header = `${blockIndent}Inputs = ordered() {`;
      const inner = outputSegments.length
        ? newline + newline + outputSegments.map((segment) => segment.chunk).join(newline) + newline + innerIndent
        : '';
      let replacement = `${header}${inner}${blockIndent}}`;
      if (hasComma) replacement += ',';
      replacement += trailingWhitespace || newline;

      if (typeof logDiag === 'function' && diagnosticsController?.isEnabled?.()) {
        try {
          logDiag(`[Routing Inputs] managed=${managedNames.size}, output=${desiredByName.size}`);
        } catch (_) {}
      }
      return text.slice(0, idx) + replacement + text.slice(after);
    } catch (_) {
      return text;
    }
  }

  function canPreservePrimaryInputsBlock(text, result, segments, eol) {
    try {
      if (!text || !result || !Array.isArray(segments)) return false;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return false;
      let blocks = findGroupInputsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!Array.isArray(blocks) || !blocks.length) {
        const fallback = findPrimaryInputsBlockFallback(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
        blocks = fallback ? [fallback] : [];
      }
      if (!Array.isArray(blocks) || !blocks.length) return false;
      const inputs = blocks[0];
      if (!inputs) return false;
      const currentEntries = parseOrderedBlockEntries(text, inputs.openIndex, inputs.closeIndex)
        .filter((item) => item && item.name && getOrderedBlockEntryType(text, item) === 'InstanceInput');
      if (currentEntries.length !== segments.length) return false;
      for (let i = 0; i < segments.length; i++) {
        const current = currentEntries[i];
        const desired = segments[i];
        if (!current || !desired) return false;
        const desiredKeyMatch = String(desired.chunk || '').match(/^\s*([^\s=]+)\s*=\s*InstanceInput\b/);
        const desiredKey = desiredKeyMatch && desiredKeyMatch[1] ? desiredKeyMatch[1] : '';
        if (!desiredKey || current.name !== desiredKey) return false;
        const currentRaw = text.slice(current.nameStart, current.blockClose + 1);
        const currentNorm = normalizeInstanceInputChunkForCompare(currentRaw);
        const desiredNorm = normalizeInstanceInputChunkForCompare(desired.chunk, eol);
        if (currentNorm !== desiredNorm) return false;
      }
      return true;
    } catch (_) {
      return false;
    }
  }

  function normalizeInstanceInputChunkForCompare(raw, eol) {
    try {
      if (raw == null) return '';
      const normalizedEol = eol || '\n';
      let out = String(raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const lines = out.split('\n');
      while (lines.length && lines[0].trim() === '') lines.shift();
      while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
      out = lines
        .map((line) => line.trim())
        .filter((line, idx, arr) => line.length > 0 || (idx > 0 && idx < arr.length - 1))
        .join(normalizedEol);
      out = out.replace(/,\s*$/, '');
      return out.trim();
    } catch (_) {
      return String(raw || '').trim();
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
          if (ch === '"' && !isQuoteEscaped(text, i)) inStr = false;
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
      let i = groupOpen + 1;
      let depth = 1;
      let inStr = false;
      while (i < groupClose) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && !isQuoteEscaped(text, i)) inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; if (depth <= 0) break; i++; continue; }
        if (depth === 1 && text.slice(i, i + 12) === 'UserControls') {
          let j = i + 12;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text.slice(j, j + 8) === 'ordered(') {
            j += 8;
            while (j < groupClose && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < groupClose && isSpace(text[j])) j++;
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
      const keep = new Set();
        if (result.insertClickedKeys instanceof Set) {
          result.insertClickedKeys.forEach((key) => keep.add(key));
        }
        if (result.buttonExactInsert instanceof Set) {
          result.buttonExactInsert.forEach((key) => keep.add(key));
        }
      let out = text;
      for (const e of (result.entries || [])) {
        if (!e || !e.sourceOp || !e.source) continue;
        // Only manage published ButtonControls and skip YouTubeButton template
          try {
            if (String(e.source) === 'YouTubeButton') continue;
            if (!isButtonControl(out, grp.groupOpenIndex, grp.groupCloseIndex, e.sourceOp, e.source)) continue;
            const key = `${e.sourceOp}.${e.source}`;
            if (isUpdateDataButtonEntry(e)) keep.add(key);
            if (keep.has(key)) continue; // preserve for explicit inserts
            // tool-level
          const tb = findToolBlockInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, e.sourceOp);
          let uc = tb ? findUserControlsInTool(out, tb.open, tb.close) : null;
          if (!uc) uc = findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
          if (uc) {
            const cb = findControlBlockInUc(out, uc.open, uc.close, e.source);
            if (cb) {
              const isGroupLevel = !tb;
              if (isGroupLevel) {
                const policy = canMutateGroupUserControl(result, e.source, { action: 'update' });
                if (!policy.allowed) {
                  recordGroupUserControlMutationIssue(result, {
                    control: e.source,
                    action: 'update',
                    message: policy.message,
                  });
                } else {
                  out = removeBtncsExecuteInRange(out, cb.open, cb.close);
                }
              } else {
                out = removeBtncsExecuteInRange(out, cb.open, cb.close);
              }
            }
          }
          // group-level
          try {
            const ucg = findUserControlsInGroup(out, grp.groupOpenIndex, grp.groupCloseIndex);
            if (ucg) {
              const cb2 = findControlBlockInUc(out, ucg.open, ucg.close, e.source);
              if (cb2) {
                const policy = canMutateGroupUserControl(result, e.source, { action: 'update' });
                if (!policy.allowed) {
                  recordGroupUserControlMutationIssue(result, {
                    control: e.source,
                    action: 'update',
                    message: policy.message,
                  });
                } else {
                  out = removeBtncsExecuteInRange(out, cb2.open, cb2.close);
                }
              }
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

  function decodeLuaLongStringLiteral(literal) {
    try {
      const value = String(literal || '').trim();
      const open = /^\[(=*)\[/.exec(value);
      if (!open) return null;
      const eq = open[1] || '';
      const closeToken = `]${eq}]`;
      const contentStart = open[0].length;
      const closeIndex = value.lastIndexOf(closeToken);
      if (closeIndex < contentStart) return null;
      return value.slice(contentStart, closeIndex);
    } catch (_) {
      return null;
    }
  }

  function buildLuaLongStringLiteral(text) {
    try {
      const src = String(text || '');
      for (let eqCount = 0; eqCount <= 8; eqCount += 1) {
        const eq = '='.repeat(eqCount);
        const open = `[${eq}[`;
        const close = `]${eq}]`;
        if (!src.includes(close)) {
          return `${open}${src}${close}`;
        }
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function buildSettingLuaScriptLiteral(script, options = {}) {
    try {
      const src = String(script || '');
      if (!src.trim()) return '';
      const longThreshold = Math.max(256, Number(options?.longThreshold) || 1200);
      const preferLong = !!options?.preferLong || src.length >= longThreshold;
      if (preferLong) {
        const longLiteral = buildLuaLongStringLiteral(src);
        if (longLiteral) return longLiteral;
      }
      return `"${escapeSettingString(src)}"`;
    } catch (_) {
      return `"${escapeSettingString(String(script || ''))}"`;
    }
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
      const resolved = findUserControlBlockForToolOrGroup(original, groupOpen, groupClose, toolName, controlId);
      if (!resolved || !resolved.controlBlock) return false;
      const slice = original.slice(resolved.controlBlock.open, Math.min(resolved.controlBlock.close + 1, original.length));
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
      const resolved = findUserControlBlockForToolOrGroup(original, groupOpen, groupClose, toolName, controlId);
      if (!resolved || !resolved.uc || !resolved.controlBlock) return false;
      const propLine = extractBtncsExecuteString(original, resolved.uc.open, resolved.uc.close, controlId);
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
      let out = resultText;
      try { logDiag(`[Launcher] queue size: ${req.size}`); } catch(_) {}
      // Attempt to source template from an existing YouTubeButton control under same tool or any tool
      const entries = Array.isArray(result?.entries) ? result.entries : [];
      req.forEach((key) => {
        try {
          const grp = locateMacroGroupBounds(out, result);
          if (!grp) return;
          const dot = String(key).indexOf('.'); if (dot < 0) return;
          const tool = key.slice(0, dot); const ctrl = key.slice(dot + 1);
          const entry = entries.find(e => e && e.sourceOp === tool && e.source === ctrl) || null;
          const pageName = entry && entry.page ? entry.page : 'Controls';
          try { logDiag(`[Launcher] processing ${tool}.${ctrl}`); } catch(_) {}
          // find target control block
          const resolvedTarget = findUserControlBlockForToolOrGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, tool, ctrl);
          if (!resolvedTarget || !resolvedTarget.controlBlock || !resolvedTarget.uc) return;
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
          if (!templateLine && resolvedTarget.scope === 'tool') {
            templateLine = extractBtncsExecuteString(out, resolvedTarget.uc.open, resolvedTarget.uc.close, 'YouTubeButton');
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
          out = insertExactYouTubeLauncher(out, resolvedTarget.controlBlock.open, resolvedTarget.controlBlock.close, eol, templateLine, pageName);
          try { logDiag(`[Launcher] Inserted for ${tool}.${ctrl} at tool-level`); } catch(_) {}
          // Also ensure a macro GroupOperator.UserControls control exists and has launcher
          try {
            const grpAfterInsert = locateMacroGroupBounds(out, result);
            if (grpAfterInsert) {
              out = ensureGroupControlHasLauncher(out, grpAfterInsert.groupOpenIndex, grpAfterInsert.groupCloseIndex, ctrl, templateLine, eol, pageName);
            }
          } catch(_) {}
        } catch(_) {}
      });
      return out;
    } catch(_) { return resultText; }
  }

  function findAllGroupUserControlsBlocks(text, groupOpen, groupClose) {
    try {
      const blocks = [];
      let i = groupOpen + 1;
      let depth = 1;
      let inStr = false;
      while (i < groupClose) {
        const ch = text[i];
        if (inStr) {
          if (ch === '"' && !isQuoteEscaped(text, i)) inStr = false;
          i++;
          continue;
        }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; if (depth <= 0) break; i++; continue; }
        if (depth === 1 && text.slice(i, i + 12) === 'UserControls') {
          const declStart = i;
          let j = i + 12;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text.slice(j, j + 8) === 'ordered(') {
            j += 8;
            while (j < groupClose && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < groupClose && isSpace(text[j])) j++;
          }
          if (text[j] !== '{') { i++; continue; }
          const openIndex = j;
          const closeIndex = findMatchingBrace(text, openIndex);
          if (closeIndex < 0 || closeIndex > groupClose) return blocks;
          let endIndex = closeIndex + 1;
          while (endIndex < groupClose && /[ \t]/.test(text[endIndex])) endIndex++;
          if (text[endIndex] === ',') endIndex++;
          blocks.push({ declStart, openIndex, closeIndex, endIndex });
          i = endIndex;
          continue;
        }
        i++;
      }
      if (!blocks.length) {
        const groupText = text.slice(groupOpen, groupClose);
        const match = /Inputs\s*=\s*(?:ordered\(\))?\s*\{/m.exec(groupText);
        if (match) {
          const openIndex = groupOpen + match.index + match[0].lastIndexOf('{');
          const closeIndex = findMatchingBrace(text, openIndex);
          if (closeIndex > openIndex && closeIndex <= groupClose) {
            blocks.push({ inputsHeaderStart: groupOpen + match.index, openIndex, closeIndex });
          }
        }
      }
      return blocks;
    } catch (_) {
      return [];
    }
  }

  function findPrimaryInputsBlockFallback(text, groupOpen, groupClose) {
    try {
      if (!text) return null;
      const start = Number.isFinite(groupOpen) ? groupOpen : 0;
      const end = Number.isFinite(groupClose) ? groupClose : text.length;
      const toolsPos = text.indexOf('Tools = ordered()', start);
      const searchEnd = (toolsPos >= 0 && toolsPos < end) ? toolsPos : end;
      const segment = text.slice(start, searchEnd);
      const match = /Inputs\s*=\s*(?:ordered\(\))?\s*\{/m.exec(segment);
      if (!match) return null;
      const openIndex = start + match.index + match[0].lastIndexOf('{');
      const closeIndex = findMatchingBrace(text, openIndex);
      if (closeIndex > openIndex && closeIndex <= end) {
        return { inputsHeaderStart: start + match.index, openIndex, closeIndex };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function parseGroupUserControlValueLiteral(body, prop) {
    try {
      const literal = extractControlPropLiteral(body, prop);
      if (!literal) return null;
      const trimmed = String(literal).trim();
      if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
        return unescapeSettingString(trimmed.slice(1, -1));
      }
      const longDecoded = decodeLuaLongStringLiteral(trimmed);
      if (longDecoded != null) return longDecoded;
      return trimmed;
    } catch (_) {
      return null;
    }
  }

  function normalizeGroupUserControlInputControl(value) {
    try {
      let raw = String(value || '').trim();
      if (raw.length >= 2 && raw.startsWith('"') && raw.endsWith('"')) {
        raw = raw.slice(1, -1);
      }
      return raw.trim().toLowerCase();
    } catch (_) {
      return '';
    }
  }

  function ensureQuickSetLabelsMaterializedInGroupUserControls(text, result, eol) {
    try {
      if (!text || !result || !Array.isArray(result.entries)) return text;
      const quickEntries = result.entries.filter((entry) => {
        const sourceId = String(entry?.source || '').trim();
        if (!/^MM_QuickSetLabel_/i.test(sourceId)) return false;
        return isGroupUserControlEntry(entry);
      });
      if (!quickEntries.length) return text;
      synchronizeAutoQuickSetLabelEntries(result, text);
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const map = new Map();
      quickEntries.forEach((entry) => {
        if (!entry || !entry.source) return;
        const sourceId = String(entry.source || '').trim();
        if (!sourceId) return;
        const pageName = (entry.page && String(entry.page).trim()) ? String(entry.page).trim() : 'Controls';
        map.set(sourceId, {
          page: pageName,
          entry,
          isGroupOwned: true,
        });
      });
      if (!map.size) return text;
      return rewriteGroupUserControls(text, map, bounds, eol || '\n', result);
    } catch (_) {
      return text;
    }
  }

  function isKnownGroupUserControlInputControl(value) {
    const known = new Set([
      'buttoncontrol',
      'checkboxcontrol',
      'combocontrol',
      'colorcontrol',
      'labelcontrol',
      'multibuttoncontrol',
      'offsetcontrol',
      'screwcontrol',
      'slidercontrol',
      'texteditcontrol',
    ]);
    return known.has(normalizeGroupUserControlInputControl(value));
  }

  function classifyGroupUserControl(control) {
    try {
      const id = String(control?.id || '').trim();
      const inputControl = String(control?.inputControl || '').trim();
      const hiddenSystem = isMacroLayerOnlyControlId(id);
      const knownType = isKnownGroupUserControlInputControl(inputControl);
      return {
        kind: hiddenSystem ? 'system' : (knownType ? 'known' : 'unknown'),
        preserveOnly: !hiddenSystem && !knownType,
      };
    } catch (_) {
      return { kind: 'unknown', preserveOnly: true };
    }
  }

  function getParsedGroupUserControlById(result, controlId) {
    try {
      if (!result || !controlId) return null;
      const map = result.groupUserControlsById instanceof Map ? result.groupUserControlsById : null;
      if (map && map.has(controlId)) return map.get(controlId) || null;
      const controls = Array.isArray(result.groupUserControls) ? result.groupUserControls : [];
      return controls.find((control) => control && control.id === controlId) || null;
    } catch (_) {
      return null;
    }
  }

  function recordGroupUserControlMutationIssue(result, issue) {
    try {
      if (!result || !issue) return;
      if (!Array.isArray(result.groupUserControlMutationIssues)) {
        result.groupUserControlMutationIssues = [];
      }
      result.groupUserControlMutationIssues.push({
        control: issue.control || '',
        action: issue.action || '',
        message: issue.message || '',
      });
      try {
        if (issue.message) logDiag(`[Group UC mutate] ${issue.message}`);
      } catch (_) {}
    } catch (_) {}
  }

  function canMutateGroupUserControl(result, controlId, options = {}) {
    try {
      const existing = getParsedGroupUserControlById(result, controlId);
      if (!existing) return { allowed: true, control: null };
      if (existing.preserveOnly && options.allowPreserveOnlyMutation !== true) {
        return {
          allowed: false,
          control: existing,
          message: `Blocked ${options.action || 'mutation'} of preserve-only macro-level control "${controlId}".`,
        };
      }
      return { allowed: true, control: existing };
    } catch (_) {
      return { allowed: true, control: null };
    }
  }

  function applyGroupUserControlMeta(control, meta) {
    try {
      if (!control || !meta) return meta;
      meta.inputControl = meta.inputControl || control.inputControl || '';
      meta.dataType = meta.dataType || control.dataType || '';
      if (!Number.isFinite(meta.labelCount) && Number.isFinite(control.labelCount)) {
        meta.labelCount = Number(control.labelCount);
      }
      if (meta.defaultValue == null || meta.defaultValue === '') meta.defaultValue = control.defaultValue;
      meta.textLines = meta.textLines || extractControlPropValue(control.rawBody, 'TEC_Lines');
      meta.integer = meta.integer || extractControlPropValue(control.rawBody, 'INP_Integer');
      meta.minAllowed = meta.minAllowed || extractControlPropValue(control.rawBody, 'INP_MinAllowed');
      meta.maxAllowed = meta.maxAllowed || extractControlPropValue(control.rawBody, 'INP_MaxAllowed');
      meta.minScale = meta.minScale || extractControlPropValue(control.rawBody, 'INP_MinScale');
      meta.maxScale = meta.maxScale || extractControlPropValue(control.rawBody, 'INP_MaxScale');
      meta.multiButtonShowBasic = meta.multiButtonShowBasic || extractControlPropValue(control.rawBody, 'MBTNC_ShowBasicButton');
      if (!Array.isArray(meta.choiceOptions) || !meta.choiceOptions.length) {
        meta.choiceOptions = Array.isArray(control.choiceOptions) ? [...control.choiceOptions] : [];
      }
      return meta;
    } catch (_) {
      return meta;
    }
  }

  function parseGroupUserControls(text, result) {
    try {
      if (!text || !result) {
        return { blocks: [], controls: [], duplicateIds: [], blockCount: 0 };
      }
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) {
        return { blocks: [], controls: [], duplicateIds: [], blockCount: 0 };
      }
      const blocks = findAllGroupUserControlsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      const controls = [];
      const seenIds = new Map();
      const duplicateIds = new Set();
      blocks.forEach((block, blockIndex) => {
        const segments = collectUserControlSegments(text, block.openIndex, block.closeIndex);
        segments.forEach((segment, orderIndex) => {
          const body = text.slice(segment.blockOpen + 1, segment.blockClose);
          const control = {
            id: String(segment.name || '').trim(),
            inputControl: extractControlPropValue(body, 'INPID_InputControl') || '',
            dataType: extractControlPropValue(body, 'LINKID_DataType') || '',
            page: extractQuotedProp(body, 'ICS_ControlPage') || '',
            visible: extractControlPropValue(body, 'IC_Visible'),
            labelCount: (() => {
              const labelCountMatch = body.match(/LBLC_NumInputs\s*=\s*([0-9]+)/i);
              return labelCountMatch ? parseInt(labelCountMatch[1], 10) : null;
            })(),
            defaultValue: parseGroupUserControlValueLiteral(body, 'INP_Default'),
            onChange: parseGroupUserControlValueLiteral(body, 'INPS_ExecuteOnChange'),
            buttonExecute: parseGroupUserControlValueLiteral(body, 'BTNCS_Execute'),
            choiceOptions: extractChoiceOptionsFromBody(body),
            multiButtonShowBasic: extractControlPropValue(body, 'MBTNC_ShowBasicButton'),
            rawBody: body,
            rawText: text.slice(segment.start, segment.end),
            blockIndex,
            orderIndex,
            range: {
              start: segment.start,
              end: segment.end,
              open: segment.blockOpen,
              close: segment.blockClose,
            },
          };
          const classification = classifyGroupUserControl(control);
          control.kind = classification.kind;
          control.preserveOnly = classification.preserveOnly;
          if (control.id) {
            if (seenIds.has(control.id)) duplicateIds.add(control.id);
            else seenIds.set(control.id, true);
          }
          controls.push(control);
        });
      });
      return {
        blocks,
        controls,
        duplicateIds: Array.from(duplicateIds),
        blockCount: blocks.length,
      };
    } catch (_) {
      return { blocks: [], controls: [], duplicateIds: [], blockCount: 0 };
    }
  }

  function hydrateGroupUserControlsState(text, result) {
    try {
      if (!result) return;
      const parsed = parseGroupUserControls(text, result);
      result.groupUserControls = Array.isArray(parsed.controls) ? parsed.controls : [];
      result.groupUserControlsById = new Map();
      result.groupUserControls.forEach((control) => {
        if (!control || !control.id || result.groupUserControlsById.has(control.id)) return;
        result.groupUserControlsById.set(control.id, control);
      });
      result.groupUserControlsBlocks = Array.isArray(parsed.blocks) ? parsed.blocks : [];
      const unknownControls = result.groupUserControls.filter((control) => control?.kind === 'unknown');
      result.groupUserControlIssues = {
        blockCount: Number(parsed.blockCount || 0),
        duplicateIds: Array.isArray(parsed.duplicateIds) ? [...parsed.duplicateIds] : [],
        unknownControls: unknownControls.map((control) => ({
          id: control.id,
          inputControl: control.inputControl || '',
        })),
      };
      result.groupUserControlMutationIssues = [];
      try {
        logDiag(`[Group UC] blocks=${result.groupUserControlIssues.blockCount}, controls=${result.groupUserControls.length}, duplicateIds=${result.groupUserControlIssues.duplicateIds.length}, unknown=${result.groupUserControlIssues.unknownControls.length}`);
      } catch (_) {}
    } catch (_) {
      result.groupUserControls = [];
      result.groupUserControlsById = new Map();
      result.groupUserControlsBlocks = [];
      result.groupUserControlIssues = { blockCount: 0, duplicateIds: [], unknownControls: [] };
      result.groupUserControlMutationIssues = [];
    }
  }

  function isMacroLayerOnlyControlId(controlId) {
    try {
      const id = String(controlId || '').trim();
      if (!id) return false;
      if (id.startsWith(`${FMR_PRESET_DATA_CONTROL}_`)) return true;
      if (id.startsWith(`${FMR_PRESET_SCRIPT_CONTROL}_`)) return true;
      return id === FMR_FILE_META_CONTROL
        || id === FMR_DATA_LINK_CONTROL
        || id === FMR_PRESET_DATA_CONTROL
        || id === FMR_PRESET_SCRIPT_CONTROL
        || id === FMR_PRESET_SELECT_CONTROL;
    } catch (_) {
      return false;
    }
  }

  function isMacroControlLeakedIntoInputs(text, controlId) {
    try {
      const id = String(controlId || '').trim();
      if (!id || !text) return false;
      const re = new RegExp(`(^|\\n|\\r)\\s*${escapeRegex(id)}\\s*=\\s*InstanceInput\\s*\\{`, 'm');
      return re.test(String(text));
    } catch (_) {
      return false;
    }
  }

  function validateGroupUserControlsStructure(text, result) {
    try {
      const parsed = parseGroupUserControls(text, result);
      const issues = [];
      if (Number(parsed.blockCount || 0) > 1) {
        issues.push({
          type: 'multiple-group-usercontrols',
          message: `Macro has ${parsed.blockCount} top-level UserControls blocks. Expected at most 1.`,
        });
      }
      (Array.isArray(parsed.duplicateIds) ? parsed.duplicateIds : []).forEach((id) => {
        issues.push({
          type: 'duplicate-group-usercontrol-id',
          control: id,
          message: `Macro-level UserControls contains duplicate control id "${id}".`,
        });
      });
      const leakedMacroIds = [];
      (result?.entries || []).forEach((entry) => {
        if (!entry) return;
        if (!isMacroLayerOnlyControlId(entry.source)) return;
        const id = String(entry.source || '').trim();
        if (!id) return;
        if (!isMacroControlLeakedIntoInputs(text, id)) return;
        leakedMacroIds.push(id);
      });
      leakedMacroIds.forEach((id) => {
        issues.push({
          type: 'macro-control-leaked-into-inputs',
          control: id,
          message: `Macro-layer control "${id}" is present in published Inputs/entries and should remain group-owned.`,
        });
      });
      return {
        issues,
        parsed,
      };
    } catch (_) {
      return { issues: [], parsed: { blocks: [], controls: [], duplicateIds: [], blockCount: 0 } };
    }
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
        out = out.slice(0, groupClose) + block + out.slice(groupClose);
        // Re-locate positions after insertion
        const grp2Close = findMatchingBrace(out, groupOpen);
        ucGroup = findUserControlsInGroup(out, groupOpen, grp2Close >= 0 ? grp2Close : groupClose) || null;
      }
      if (!ucGroup) return out;
      // If control exists, insert launcher there; else create minimal control with launcher
      let cb = findControlBlockInUc(out, ucGroup.open, ucGroup.close, ctrl);
      if (cb) {
        const policy = canMutateGroupUserControl(state.parseResult, ctrl, { action: 'update' });
        if (!policy.allowed) {
          recordGroupUserControlMutationIssue(state.parseResult, {
            control: ctrl,
            action: 'update',
            message: policy.message,
          });
          return out;
        }
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
      const requestedName = arguments.length > 2 ? String(arguments[2] || '').trim() : '';
      const newName = requestedName || (result.macroName || '').trim();
      const targetName = newName || originalName;
      if (!originalName || !targetName) return text;
      let safeName = sanitizeIdent(targetName);
      if (!safeName) safeName = sanitizeIdent(originalName) || 'Macro';
      if (!isIdentStart(safeName[0])) safeName = `_${safeName}`;
      const desiredType = (result.operatorType || result.operatorTypeOriginal || 'GroupOperator').trim() || 'GroupOperator';
      const nameEsc = originalName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp('(^|\\n)(\\s*)' + nameEsc + '(\\s*=\\s*)(GroupOperator|MacroOperator)(\\s*\\{)', 'm');
      return text.replace(re, (match, prefix, spaces, equalsPart, existingType, bracePart) => {
        const typeOut = desiredType || existingType;
        return `${prefix}${spaces}${safeName}${equalsPart}${typeOut}${bracePart}`;
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
          const resolved = findUserControlBlockForToolOrGroup(out, grp.groupOpenIndex, grp.groupCloseIndex, tool, ctrl);
          if (!resolved || !resolved.controlBlock) return;
          if (resolved.scope === 'group') {
            const policy = canMutateGroupUserControl(result, ctrl, { action: 'update' });
            if (!policy.allowed) {
              recordGroupUserControlMutationIssue(result, {
                control: ctrl,
                action: 'update',
                message: policy.message,
              });
              return;
            }
          }
          const before = out;
          out = rewriteBtncsExecuteInRange(out, resolved.controlBlock.open, resolved.controlBlock.close, v, eol);
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

  function findToolBlockAnywhere(text, toolName) {
    try {
      if (!text || !toolName) return null;
      const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp('(^|\\n)\\s*' + esc(String(toolName)) + '\\s*=\\s*[A-Za-z_][A-Za-z0-9_]*\\s*\\{', 'm');
      const m = re.exec(text);
      if (!m) return null;
      const relOpen = m.index + m[0].lastIndexOf('{');
      const absOpen = relOpen;
      const absClose = findMatchingBrace(text, absOpen);
      if (absClose < 0) return null;
      return { open: absOpen, close: absClose };
    } catch (_) { return null; }
  }

  function findUserControlsInGroup(text, groupOpen, groupClose) {
    try {
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
        if (ch === '}') { depth--; i++; if (depth <= 0) break; continue; }
        if (depth === 1 && text.slice(i, i + 12) === 'UserControls') {
          let j = i + 12;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text[j] !== '=') { i++; continue; }
          j++;
          while (j < groupClose && isSpace(text[j])) j++;
          if (text.slice(j, j + 8) === 'ordered(') {
            j += 8;
            while (j < groupClose && text[j] !== ')') j++;
            if (text[j] === ')') j++;
            while (j < groupClose && isSpace(text[j])) j++;
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

  function findInputsInTool(text, toolOpen, toolClose) {
    try {
      if (toolOpen == null || toolClose == null || toolClose <= toolOpen) return null;
      const segment = text.slice(toolOpen, toolClose);
      const match = /Inputs\s*=\s*(?:ordered\(\))?\s*\{/m.exec(segment);
      if (!match) return null;
      const inputOpen = toolOpen + match.index + match[0].lastIndexOf('{');
      if (inputOpen < toolOpen || inputOpen > toolClose) return null;
      const inputClose = findMatchingBrace(text, inputOpen);
      if (inputClose < 0 || inputClose > toolClose) return null;
      return { open: inputOpen, close: inputClose };
    } catch (_) { return null; }
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
            return { idStart, open: cOpen, close: cClose };
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
      return ensureUserControlsBlockInToolBlock(text, toolBlock, newline);
    } catch (_) {
      return { text, toolBlock: null, ucBlock: null };
    }
  }

  function ensureUserControlsBlockInToolBlock(text, toolBlock, eol) {
    try {
      if (!toolBlock) return { text, toolBlock: null, ucBlock: null };
      const newline = eol || '\n';
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

      const labels = (result.entries || []).filter((x) => x && x.isLabel && Number.isFinite(x.labelCount) && (x.headerCarrierSynthetic || x.controlMetaDirty));

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

            if (!Number(newCount)) { i = cClose + 1; continue; }

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
      const labels = result.entries.filter(e => e && e.isLabel && typeof e.labelHidden === 'boolean' && !!e.controlMetaDirty);
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
      if (!hidden && !/IC_Visible\s*=/.test(body)) return text;
      if (hidden) {
        body = upsertControlProp(body, 'IC_Visible', 'false', indent, eol);
      } else {
        body = upsertControlProp(body, 'IC_Visible', 'true', indent, eol);
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
        if (!entry.controlMetaDirty && !entry.headerCarrierSynthetic) continue;
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
      const wantedName = normalizeId(controlName);
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
          const isInput = text.slice(i, i + 5) === 'Input';
          const isInstanceInput = text.slice(i, i + 13) === 'InstanceInput';
          if (isInstanceInput) {
            i += 13;
            while (i < inputsClose && isSpace(text[i])) i++;
          } else if (isInput) {
            i += 5;
            while (i < inputsClose && isSpace(text[i])) i++;
          } else {
            continue;
          }
          if (text[i] !== '{') { i++; continue; }
          const blockOpen = i;
          const blockClose = findMatchingBrace(text, blockOpen);
          if (blockClose < 0) break;
          if (String(norm) === String(wantedName)) {
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
        let page = null;
        if (uc) {
          const block = findControlBlockInUc(text, uc.open, uc.close, entry.source);
          if (block) {
            const body = text.slice(block.open + 1, block.close);
            page = extractQuotedProp(body, 'ICS_ControlPage');
          }
        }
        if (!page) {
          const groupControl = getParsedGroupUserControlById(result, entry.source);
          if (groupControl && groupControl.page) page = groupControl.page;
        }
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
      const map = new Map();
      const controls = Array.isArray(result.groupUserControls) ? result.groupUserControls : [];
      if (!controls.length) {
        result.pageIcons = map;
        return;
      }
      controls.forEach((control) => {
        if (!control || !control.id || !control.rawBody) return;
        if (control.inputControl) return;
        if (!/\bCTID_DIB_ID\b/.test(control.rawBody)) return;
        const iconId = extractQuotedProp(control.rawBody, 'CTID_DIB_ID');
        if (iconId) map.set(normalizePageNameGlobal(control.id), iconId);
      });
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

  function extractControlPropLiteral(body, prop) {
    try {
      if (!body || !prop) return null;
      const range = findControlPropRange(body, prop);
      if (!range) return null;
      const line = body.slice(range.start, range.end);
      const eq = line.indexOf('=');
      if (eq < 0) return null;
      let value = line.slice(eq + 1).trim();
      if (value.endsWith(',')) value = value.slice(0, -1).trim();
      return value || null;
    } catch (_) {
      return null;
    }
  }

  function extractControlStringProp(body, prop) {
    try {
      if (!body || !prop) return null;
      const range = findControlPropRange(body, prop);
      if (!range) return null;
      const line = body.slice(range.start, range.end);
      const eq = line.indexOf('=');
      if (eq < 0) return null;
      let value = line.slice(eq + 1).trim();
      if (value.endsWith(',')) value = value.slice(0, -1).trim();
      if (!value) return null;
      if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) {
        return unescapeSettingString(value.slice(1, -1));
      }
      const longDecoded = decodeLuaLongStringLiteral(value);
      if (longDecoded != null) return longDecoded;
      return value;
    } catch (_) {
      return null;
    }
  }

  function buildPresetEnginePayload(result) {
    try {
      const engine = ensurePresetEngine(result);
      if (!engine || typeof engine !== 'object') return null;
      const scopeEntries = Array.isArray(engine.scopeEntries) ? engine.scopeEntries : [];
      const scope = Array.isArray(engine.scope)
        ? engine.scope.map((key) => String(key || '').trim()).filter(Boolean)
        : [];
      const presets = {};
      const presetObj = engine.presets && typeof engine.presets === 'object' ? engine.presets : {};
      Object.keys(presetObj).forEach((name) => {
        const cleanName = String(name || '').trim();
        if (!cleanName) return;
        const values = presetObj[name];
        if (!values || typeof values !== 'object') return;
        const cleanValues = {};
        scope.forEach((key) => {
      if (values[key] == null) return;
          cleanValues[key] = normalizePresetPayloadValue(values[key]);
        });
        if (Object.keys(cleanValues).length) {
          presets[cleanName] = cleanValues;
        }
      });
      const base = {};
      const baseObj = engine.base && typeof engine.base === 'object' ? engine.base : {};
      scope.forEach((key) => {
        if (baseObj[key] == null) return;
        base[key] = normalizePresetPayloadValue(baseObj[key]);
      });
      if (!scope.length && !Object.keys(presets).length) return null;
      const activePreset = String(engine.activePreset || '').trim();
      return {
        version: Math.max(2, Number(engine.version) || 1),
        buildMarker: String(engine.buildMarker || FMR_BUILD_MARKER),
        scope,
        scopeEntries: scopeEntries.length ? scopeEntries.map((item) => ({ ...item })) : undefined,
        base: Object.keys(base).length ? base : undefined,
        presets,
        activePreset: activePreset || undefined,
      };
    } catch (_) {
      return null;
    }
  }

  function buildPresetRuntimeState(result) {
    try {
      const payload = buildPresetEnginePayload(result);
      if (!payload) return null;
      const scopeEntries = normalizePresetScopeEntries(result, payload.scopeEntries, payload.scope);
      const presetNames = payload.presets && typeof payload.presets === 'object'
        ? Object.keys(payload.presets).filter((name) => String(name || '').trim())
        : [];
      if (!scopeEntries.length || !presetNames.length) return null;
      const runtimeTargets = resolvePresetRuntimeTargets(result, scopeEntries);
      if (!runtimeTargets.length) return null;
      const defaultPresetName = String(payload.activePreset || presetNames[0] || '').trim();
      const defaultIndex = Math.max(0, presetNames.indexOf(defaultPresetName));
      const pathSwitchPlan = buildPresetPathSwitchPlan(
        result,
        { ...payload, scopeEntries },
        runtimeTargets,
        presetNames,
        defaultIndex
      );
      const script = buildPresetApplyOnChangeScript(result, {
        ...payload,
        scopeEntries,
        pathSwitchPlan,
      }, FMR_PRESET_SELECT_CONTROL);
      if (!script) return null;
      return {
        payload: {
          ...payload,
          scopeEntries,
        },
        presetNames,
        defaultIndex,
        pathSwitchPlan,
        script,
      };
    } catch (_) {
      return null;
    }
  }

  function extractControlStringPropList(body, prop) {
    try {
      if (!body || !prop) return [];
      const escapedProp = String(prop).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(`(?:\\{\\s*)?${escapedProp}\\s*=\\s*("(?:\\\\.|[^"])*"|[^,}\\r\\n]+)\\s*(?:\\})?`, 'ig');
      const values = [];
      let match;
      while ((match = re.exec(body))) {
        let value = (match[1] || '').trim();
        if (!value) continue;
        if (value.endsWith(',')) value = value.slice(0, -1).trim();
        value = value.replace(/\}\s*$/, '').trim();
        if (!value) continue;
        if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) {
          value = unescapeSettingString(value.slice(1, -1));
        }
        values.push(value);
      }
      return values;
    } catch (_) {
      return [];
    }
  }

  function extractChoiceOptionsFromBody(body) {
    try {
      const combo = extractControlStringPropList(body, 'CCS_AddString');
      if (combo.length) return combo;
      const buttons = extractControlStringPropList(body, 'MBTNC_AddButton');
      if (buttons.length) return buttons;
      return [];
    } catch (_) {
      return [];
    }
  }

  function getCurrentInputInfo(entry) {
    const empty = { value: '', note: '' };
    try {
      if (entry && entry.isLabel) {
        const rawDefault = normalizeMetaValue(
          entry?.controlMeta?.defaultValue ??
          entry?.controlMetaOriginal?.defaultValue ??
          extractInstancePropValue(entry?.raw || '', 'Default') ??
          entry?.labelValueOriginal
        );
        const isOpen = rawDefault === '0' ? false : true;
        return { value: isOpen ? 'Open' : 'Closed', note: 'Label default state.' };
      }
      if (!entry || !entry.sourceOp || !entry.source || !state.originalText || !state.parseResult) return empty;
      const bounds = locateMacroGroupBounds(state.originalText, state.parseResult);
      if (!bounds) return empty;
      const toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
      if (!toolBlock) {
        const fallback = normalizeMetaValue(
          entry?.controlMeta?.defaultValue ??
          entry?.controlMetaOriginal?.defaultValue ??
          extractInstancePropValue(entry?.raw || '', 'Default')
        );
        const formatted = formatDefaultForDisplay(entry, fallback);
        return formatted != null ? { value: formatted, note: 'Using Inputs default.' } : (isTextControl(entry) ? empty : { value: '', note: 'Inputs block not found.' });
      }
      if (isTextControl(entry) && String(entry.source || '').toLowerCase().includes('styledtext')) {
        const direct = extractStyledTextValueFromTool(toolBlock.open, toolBlock.close);
        if (direct != null) {
          return { value: direct, note: '' };
        }
      }
      const inputsBlock = findInputsInTool(state.originalText, toolBlock.open, toolBlock.close);
      if (!inputsBlock) {
        const fallback = normalizeMetaValue(
          entry?.controlMeta?.defaultValue ??
          entry?.controlMetaOriginal?.defaultValue ??
          extractInstancePropValue(entry?.raw || '', 'Default')
        );
        const formatted = formatDefaultForDisplay(entry, fallback);
        return formatted != null ? { value: formatted, note: 'Using Inputs default.' } : (isTextControl(entry) ? empty : { value: '', note: 'Inputs block not found.' });
      }
      const inputBlock = findInputBlockInInputs(state.originalText, inputsBlock.open, inputsBlock.close, entry.source);
      if (!inputBlock) {
        const fallback = normalizeMetaValue(
          entry?.controlMeta?.defaultValue ??
          entry?.controlMetaOriginal?.defaultValue ??
          extractInstancePropValue(entry?.raw || '', 'Default')
        );
        const formatted = formatDefaultForDisplay(entry, fallback);
        return formatted != null ? { value: formatted, note: 'Using Inputs default.' } : (isTextControl(entry) ? empty : { value: '', note: 'Input not found on the tool.' });
      }
        const body = state.originalText.slice(inputBlock.open + 1, inputBlock.close);
        let value = extractControlPropValue(body, 'Value') || '';
        if (isTextControl(entry)) {
          const directPlain = extractStyledTextPlainTextFromInput(body);
          if (directPlain != null) value = directPlain;
        const styled = extractStyledTextBlockFromInput(body);
        if (styled) {
          const plain = extractStyledTextPlainText(styled);
          if (plain != null) value = plain;
        }
        if (/^StyledText\b/.test(value.trim())) {
          const plain = extractStyledTextPlainText(value.trim());
          if (plain != null) value = plain;
        }
        value = stripDefaultQuotes(value);
      }
      const sourceOp = extractControlPropValue(body, 'SourceOp');
      const source = extractControlPropValue(body, 'Source');
      const expression = extractControlStringProp(body, 'Expression');
      const strip = (v) => String(v || '').replace(/^\"|\"$/g, '');
      let note = '';
      if (expression) {
        note = 'Driven by Expression. Setting a value overrides the link.';
      } else if (sourceOp || source) {
        const op = strip(sourceOp);
        const src = strip(source);
        note = `Driven by ${op || 'SourceOp'}${src ? '.' + src : ''}. Setting a value overrides the link.`;
      }
        if (!value && !expression && !sourceOp && !source) {
          const fallback = normalizeMetaValue(
            entry?.controlMeta?.defaultValue ??
            entry?.controlMetaOriginal?.defaultValue ??
            extractInstancePropValue(entry?.raw || '', 'Default')
          );
          if (fallback != null) {
            const formatted = formatDefaultForDisplay(entry, fallback);
            if (formatted != null && String(formatted).trim()) value = formatted;
          }
        }
        if (entry?.csvLink?.column && !isTextControl(entry)) {
          const linkedDefault = normalizeMetaValue(
            entry?.controlMeta?.defaultValue ??
            entry?.controlMetaOriginal?.defaultValue ??
            extractInstancePropValue(entry?.raw || '', 'Default')
          );
          const formatted = formatDefaultForDisplay(entry, linkedDefault);
          if (formatted != null && String(formatted).trim()) {
            value = formatted;
            if (!note) note = `Linked to CSV: ${entry.csvLink.column}`;
          }
        }
        return { value, note };
      } catch (_) {
        return empty;
      }
    }

  function extractStyledTextValueFromTool(toolOpen, toolClose) {
    try {
      if (!state.originalText || toolOpen == null || toolClose == null) return null;
      const segment = state.originalText.slice(toolOpen, toolClose);
      const match = segment.match(/StyledText\s*=\s*Input\s*\{/i);
      if (!match || match.index == null) return null;
      const start = toolOpen + match.index;
      const braceIndex = state.originalText.indexOf('{', start);
      if (braceIndex < 0 || braceIndex > toolClose) return null;
      const closeIndex = findMatchingBrace(state.originalText, braceIndex);
      if (closeIndex < 0 || closeIndex > toolClose) return null;
      const body = state.originalText.slice(braceIndex + 1, closeIndex);
      const direct = extractStyledTextPlainTextFromInput(body);
      if (direct != null) return direct;
      const matchValue = body.match(/Value\s*=\s*"((?:\\.|[^"])*)"/i);
      if (matchValue) return unescapeSettingString(matchValue[1]);
      return null;
    } catch (_) {
      return null;
    }
  }

    function hydrateControlMetadata(text, result) {
      try {
        if (!text || !result || !Array.isArray(result.entries)) return;
        const bounds = locateMacroGroupBounds(text, result) || { groupOpenIndex: 0, groupCloseIndex: text.length };
        const cache = new Map();
        const getUcForTool = (toolName) => {
          if (!toolName) return null;
          if (cache.has(toolName)) return cache.get(toolName);
          let tb = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, toolName);
          if (!tb) tb = findToolBlockAnywhere(text, toolName);
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
              meta.textLines = extractControlPropValue(body, 'TEC_Lines');
              meta.integer = extractControlPropValue(body, 'INP_Integer');
              meta.minAllowed = extractControlPropValue(body, 'INP_MinAllowed');
              meta.maxAllowed = extractControlPropValue(body, 'INP_MaxAllowed');
              meta.minScale = extractControlPropValue(body, 'INP_MinScale');
              meta.maxScale = extractControlPropValue(body, 'INP_MaxScale');
              meta.multiButtonShowBasic = extractControlPropValue(body, 'MBTNC_ShowBasicButton');
              meta.choiceOptions = extractChoiceOptionsFromBody(body);
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
            if (!meta.inputControl) {
              let tb = findToolBlockInGroup(text, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
              if (!tb) tb = findToolBlockAnywhere(text, entry.sourceOp);
              if (tb) {
                const tBlock = findControlBlockInTool(text, tb.open, tb.close, entry.source);
                if (tBlock) {
                  const body = text.slice(tBlock.open + 1, tBlock.close);
                  meta.inputControl = meta.inputControl || extractControlPropValue(body, 'INPID_InputControl');
                  meta.dataType = meta.dataType || extractControlPropValue(body, 'LINKID_DataType');
                  meta.defaultValue = meta.defaultValue || extractControlPropValue(body, 'INP_Default');
                  meta.textLines = meta.textLines || extractControlPropValue(body, 'TEC_Lines');
                  meta.integer = meta.integer || extractControlPropValue(body, 'INP_Integer');
                  meta.minAllowed = meta.minAllowed || extractControlPropValue(body, 'INP_MinAllowed');
                  meta.maxAllowed = meta.maxAllowed || extractControlPropValue(body, 'INP_MaxAllowed');
                  meta.minScale = meta.minScale || extractControlPropValue(body, 'INP_MinScale');
                  meta.maxScale = meta.maxScale || extractControlPropValue(body, 'INP_MaxScale');
                  meta.multiButtonShowBasic = meta.multiButtonShowBasic || extractControlPropValue(body, 'MBTNC_ShowBasicButton');
                  if (!Array.isArray(meta.choiceOptions) || !meta.choiceOptions.length) {
                    meta.choiceOptions = extractChoiceOptionsFromBody(body);
                  }
                }
                if (!meta.inputControl) {
                  const toolUc = findUserControlsInTool(text, tb.open, tb.close);
                  if (toolUc) {
                    const uBlock = findControlBlockInUc(text, toolUc.open, toolUc.close, entry.source);
                    if (uBlock) {
                      const body = text.slice(uBlock.open + 1, uBlock.close);
                      meta.inputControl = meta.inputControl || extractControlPropValue(body, 'INPID_InputControl');
                      meta.dataType = meta.dataType || extractControlPropValue(body, 'LINKID_DataType');
                      meta.defaultValue = meta.defaultValue || extractControlPropValue(body, 'INP_Default');
                      meta.textLines = meta.textLines || extractControlPropValue(body, 'TEC_Lines');
                      meta.integer = meta.integer || extractControlPropValue(body, 'INP_Integer');
                      meta.minAllowed = meta.minAllowed || extractControlPropValue(body, 'INP_MinAllowed');
                      meta.maxAllowed = meta.maxAllowed || extractControlPropValue(body, 'INP_MaxAllowed');
                      meta.minScale = meta.minScale || extractControlPropValue(body, 'INP_MinScale');
                      meta.maxScale = meta.maxScale || extractControlPropValue(body, 'INP_MaxScale');
                      meta.multiButtonShowBasic = meta.multiButtonShowBasic || extractControlPropValue(body, 'MBTNC_ShowBasicButton');
                      if (!Array.isArray(meta.choiceOptions) || !meta.choiceOptions.length) {
                        meta.choiceOptions = extractChoiceOptionsFromBody(body);
                      }
                    }
                  }
                }
              }
          }
          if (!meta.inputControl) {
            const groupControl = getParsedGroupUserControlById(result, entry.source);
            if (groupControl) {
              applyGroupUserControlMeta(groupControl, meta);
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
        entry.controlMeta = {
          ...meta,
          choiceOptions: Array.isArray(meta.choiceOptions) ? [...meta.choiceOptions] : [],
        };
        entry.controlMetaOriginal = {
          ...meta,
          choiceOptions: Array.isArray(meta.choiceOptions) ? [...meta.choiceOptions] : [],
        };
        entry.controlMetaDirty = false;
        if (meta.inputControl && /labelcontrol/i.test(meta.inputControl)) {
          entry.isLabel = true;
          entry.labelCount = Number.isFinite(meta.labelCount)
            ? Number(meta.labelCount)
            : (Number.isFinite(entry.labelCount) ? Number(entry.labelCount) : 0);
        }
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

  function syncEntrySortIndicesToOrder(result) {
    try {
      if (!result || !Array.isArray(result.entries) || !Array.isArray(result.order)) return;
      result.order.forEach((idx, pos) => {
        const entry = result.entries[idx];
        if (entry) entry.sortIndex = pos;
      });
    } catch (_) {}
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
        const hidden = /IC_Visible\s*=\s*false/.test(body);
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

  function applyUserControlPages(text, result, eol, options = {}) {
    try {
      if (!result || !Array.isArray(result.entries) || !result.entries.length) return text;
      const commitChanges = options.commitChanges !== false;
      const macroSourceOp = resolveMacroGroupSourceOp(text, result);
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
      const iterableEntries = [];
      const seenEntryRefs = new Set();
      const pushEntry = (entry) => {
        if (!entry || !entry.source) return;
        if (seenEntryRefs.has(entry)) return;
        seenEntryRefs.add(entry);
        iterableEntries.push(entry);
      };
      orderedIndices.forEach((idx) => {
        pushEntry(result.entries[idx]);
      });
      // Safety: macro-root quick-set/group controls should always participate in
      // export mutation even if they were accidentally dropped from result.order.
      (result.entries || []).forEach((entry) => {
        if (isGroupUserControlEntry(entry)) {
          pushEntry(entry);
        }
      });
      for (const entry of iterableEntries) {
        if (!entry || !entry.source) continue;
        const pageName = (entry.page && String(entry.page).trim()) ? String(entry.page).trim() : 'Controls';
        const isGroupOwnedEntry = isGroupUserControlEntry(entry);
        if (isGroupOwnedEntry && !entry.sourceOp && macroSourceOp) {
          entry.sourceOp = macroSourceOp;
        }
        if (!isGroupOwnedEntry && !entry.sourceOp) continue;
        if (isGroupOwnedEntry) {
          groupControls.set(entry.source, {
            page: pageName,
            entry,
            isGroupOwned: true,
          });
        } else if (!groupControls.has(entry.source)) {
          groupControls.set(entry.source, {
            page: pageName,
            entry: null,
            isGroupOwned: false,
          });
        }
        if (!isGroupOwnedEntry) {
          if (!perToolOrder.has(entry.sourceOp)) perToolOrder.set(entry.sourceOp, []);
          const arr = perToolOrder.get(entry.sourceOp);
          const normName = normalizeId(entry.source);
          if (!arr.some(n => normalizeId(n) === normName)) arr.push(entry.source);
          if (!perTool.has(entry.sourceOp)) perTool.set(entry.sourceOp, new Map());
          const isAutoQuickSetLabel = /^MM_QuickSetLabel_/i.test(String(entry.source || '').trim());
          perTool.get(entry.sourceOp).set(entry.source, {
            page: pageName,
            forceCreate: entry.syntheticToolUserControl === true || isAutoQuickSetLabel,
            entry,
          });
        }
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
        const isBlend = !isGroupUserControlEntry(entry) && String(entry.source).toLowerCase() === 'blend';
        const isToggle = entry.isBlendToggle || ensureBlendToggleSet(result).has(getEntryKey(entry));
        if (isBlend && isToggle) {
          entry.isBlendToggle = true;
          blendTargets.set(entry.sourceOp, (blendTargets.get(entry.sourceOp) || new Set()).add(entry.source));
        }
        const script = !isGroupUserControlEntry(entry) && (entry.onChange && String(entry.onChange).trim()) ? String(entry.onChange).trim() : '';
        if (script) {
          if (!onChangeTargets.has(entry.sourceOp)) onChangeTargets.set(entry.sourceOp, []);
          onChangeTargets.get(entry.sourceOp).push({ control: entry.source, script });
        }
        if (!isGroupUserControlEntry(entry) && entry.isButton) {
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
        }
      }
      let out = text;
      for (const [toolName, controls] of perTool.entries()) {
        if (!toolName || !controls.size) continue;
        out = rewriteToolUserControls(out, toolName, controls, bounds, eol, result);
      }
      bounds = locateMacroGroupBounds(out, result) || bounds;
      out = rewriteGroupUserControls(out, groupControls, bounds, eol, result);
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
      const onlyControlsPage = pageDefinitions.size === 1 && pageDefinitions.has('Controls');
      if (!onlyControlsPage) {
        bounds = locateMacroGroupBounds(out, result) || bounds;
        out = ensureControlPagesDeclared(out, bounds, pageDefinitions, eol, result);
      }
      bounds = locateMacroGroupBounds(out, result);
      const nameRes = applyDisplayNameOverrides(out, result, bounds, eol, result, { commitChanges });
      out = nameRes.text;
      bounds = nameRes.bounds || bounds;
      let metaBounds = bounds;
      for (const item of metaOverrideEntries) {
        const res = applyControlMetaOverride(out, item.entry, item.diff, metaBounds, eol, result);
        out = res.text;
        metaBounds = res.bounds || metaBounds;
        if (commitChanges) {
          item.entry.controlMetaOriginal = { ...(item.entry.controlMeta || {}) };
          item.entry.controlMetaDirty = false;
        }
      }
      bounds = metaBounds || bounds;
      // Preserve authored tool-level UserControls order by default.
      // Reordering tool custom controls to mirror macro published order is not
      // a no-op-safe transformation and has caused structural drift on real macros.
      if (options.reorderToolUserControls === true) {
        out = reorderToolUserControlsBlocks(out, bounds, perToolOrder, eol, result);
      }
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
      const parseToolControlPayload = (payload) => {
        if (payload && typeof payload === 'object') {
          const sourceId = String(payload?.entry?.source || '').trim();
          return {
            pageName: (payload.page && String(payload.page).trim()) ? String(payload.page).trim() : 'Controls',
            forceCreate: payload.forceCreate === true || /^MM_QuickSetLabel_/i.test(sourceId),
            sourceEntry: payload.entry || null,
          };
        }
        return {
          pageName: (payload && String(payload).trim()) ? String(payload).trim() : 'Controls',
          forceCreate: false,
          sourceEntry: null,
        };
      };
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(updated, resultRef);
      if (!currentBounds) return updated;
      let toolBlock = findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
      if (!toolBlock) {
        toolBlock = findToolBlockAnywhere(updated, toolName);
        if (!toolBlock) {
          return updated;
        }
      }
      let uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
      if (!uc) {
        const needsUserControlsBlock = Array.from(controlMap.values()).some((payload) => {
          const { pageName, forceCreate } = parseToolControlPayload(payload);
          const normalizedPage = normalizePageNameGlobal(pageName || 'Controls');
          return forceCreate || normalizedPage !== 'Controls';
        });
        if (!needsUserControlsBlock) return updated;
        const ensureRes = ensureUserControlsBlockInToolBlock(updated, toolBlock, newline);
        updated = ensureRes.text;
        currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
        toolBlock = ensureRes.toolBlock
          || findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName)
          || findToolBlockAnywhere(updated, toolName);
        if (!toolBlock) return updated;
        uc = ensureRes.ucBlock || findUserControlsInTool(updated, toolBlock.open, toolBlock.close);
        if (!uc) return updated;
      }
      for (const [controlName, payload] of controlMap.entries()) {
        const { pageName, forceCreate, sourceEntry } = parseToolControlPayload(payload);
        const normalizedPage = normalizePageNameGlobal(pageName || 'Controls');
        let block = findControlBlockInUc(updated, uc.open, uc.close, controlName);
        if (!block) {
          // Do not create a minimal override for default-page controls.
          // Built-in controls (e.g. MultiButton icons) can degrade to plain text if overridden.
          if (normalizedPage === 'Controls' && !forceCreate) continue;
          const ensuredBlock = ensureControlBlockInUserControls(updated, uc, controlName, newline);
          updated = ensuredBlock.text;
          uc = ensuredBlock.uc || uc;
          block = ensuredBlock.block;
          if (!block && uc) block = findControlBlockInUc(updated, uc.open, uc.close, controlName);
        }
        if (!block) continue;
        if (normalizedPage === 'Controls' && !forceCreate) {
          const body = updated.slice(block.open + 1, block.close);
          if (!/ICS_ControlPage\s*=/.test(body)) continue;
        }
        if (forceCreate && sourceEntry && (sourceEntry.isLabel || /^MM_QuickSetLabel_/i.test(String(controlName || '').trim()))) {
          const indent = (getLineIndent(updated, block.open) || '') + '\t';
          let body = updated.slice(block.open + 1, block.close);
          const labelText = String(
            sourceEntry.displayName ||
            sourceEntry.name ||
            humanizeName(sourceEntry.source || controlName) ||
            controlName
          ).trim();
          if (labelText) {
            body = upsertControlProp(body, 'LINKS_Name', `"${escapeQuotes(labelText)}"`, indent, newline);
          }
          const rawDefault = normalizeMetaValue(
            sourceEntry?.controlMeta?.defaultValue ??
            sourceEntry?.controlMetaOriginal?.defaultValue ??
            (sourceEntry?.labelValueOriginal != null ? sourceEntry.labelValueOriginal : null) ??
            '0'
          );
          const normDefault = String(rawDefault || '').trim() === '0' ? '0' : '1';
          const labelCount = Math.max(0, Number.isFinite(sourceEntry?.labelCount) ? Number(sourceEntry.labelCount) : 0);
          body = upsertControlProp(body, 'INPID_InputControl', '"LabelControl"', indent, newline);
          body = upsertControlProp(body, 'LINKID_DataType', '"Number"', indent, newline);
          body = upsertControlProp(body, 'INP_Integer', 'false', indent, newline);
          body = upsertControlProp(body, 'LBLC_DropDownButton', 'true', indent, newline);
          body = upsertControlProp(body, 'LBLC_NumInputs', String(labelCount), indent, newline);
          body = upsertControlProp(body, 'INP_SplineType', '"Default"', indent, newline);
          body = upsertControlProp(body, 'INP_Default', normDefault, indent, newline);
          updated = updated.slice(0, block.open + 1) + body + updated.slice(block.close);
          currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
          toolBlock = findToolBlockInGroup(updated, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
          if (!toolBlock) break;
          uc = findUserControlsInTool(updated, toolBlock.open, toolBlock.close) || uc;
          if (uc) {
            block = findControlBlockInUc(updated, uc.open, uc.close, controlName) || block;
          }
        }
        if (normalizedPage !== 'Controls') {
          updated = rewriteControlBlockPage(updated, block.open, block.close, normalizedPage, newline);
        }
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

  function rewriteGroupUserControls(text, controlMap, bounds, eol, resultRef) {
    try {
      if (!controlMap || !controlMap.size) return text;
      let updated = text;
      const newline = eol || '\n';
      for (const [controlName, payload] of controlMap.entries()) {
        const pageName = (payload && typeof payload === 'object') ? payload.page : payload;
        const sourceEntry = (payload && typeof payload === 'object') ? payload.entry : null;
        const isGroupOwned = !!(payload && typeof payload === 'object' && payload.isGroupOwned && sourceEntry);
        const isAutoQuickSetLabel = /^MM_QuickSetLabel_/i.test(String(controlName || '').trim());
        const createLines = isGroupOwned
          ? buildGroupUserControlCreateLinesFromEntry(sourceEntry, pageName)
          : [];
        updated = upsertGroupUserControl(updated, resultRef, controlName, {
          createIfMissing: isGroupOwned,
          allowPreserveOnlyMutation: isAutoQuickSetLabel,
          createLines,
          updateBody: (body, indent) => {
            const safePage = normalizePageNameGlobal(pageName || 'Controls');
            const parsed = getParsedGroupUserControlById(resultRef, controlName);
            const parsedBody = typeof parsed?.rawBody === 'string' ? parsed.rawBody : '';
            if (!isGroupOwned && parsedBody.trim()) {
              body = parsedBody;
            }
            if (isGroupOwned && sourceEntry) {
              const labelText = String(
                sourceEntry.displayName ||
                sourceEntry.name ||
                humanizeName(sourceEntry.source || controlName) ||
                controlName
              ).trim();
              if (labelText) {
                body = upsertControlProp(body, 'LINKS_Name', `"${escapeQuotes(labelText)}"`, indent, newline);
              }
              if (sourceEntry.isLabel) {
                const rawDefault = normalizeMetaValue(
                  sourceEntry?.controlMeta?.defaultValue ??
                  sourceEntry?.controlMetaOriginal?.defaultValue ??
                  (sourceEntry?.labelValueOriginal != null ? sourceEntry.labelValueOriginal : null) ??
                  '0'
                );
                const normDefault = String(rawDefault || '').trim() === '0' ? '0' : '1';
                const labelCount = Math.max(0, Number.isFinite(sourceEntry?.labelCount) ? Number(sourceEntry.labelCount) : 0);
                body = upsertControlProp(body, 'INPID_InputControl', '"LabelControl"', indent, newline);
                body = upsertControlProp(body, 'LINKID_DataType', '"Number"', indent, newline);
                body = upsertControlProp(body, 'INP_Integer', 'false', indent, newline);
                body = upsertControlProp(body, 'LBLC_DropDownButton', 'true', indent, newline);
                body = upsertControlProp(body, 'LBLC_NumInputs', String(labelCount), indent, newline);
                body = upsertControlProp(body, 'INP_SplineType', '"Default"', indent, newline);
                body = upsertControlProp(body, 'INP_Default', normDefault, indent, newline);
              }
            }
            const replacement = `ICS_ControlPage = "${escapeQuotes(safePage)}",`;
            if (/ICS_ControlPage\s*=/.test(body)) {
              return body.replace(/ICS_ControlPage\s*=\s*"([^"]*)"\s*,?/g, replacement);
            }
            return `${newline}${indent}${replacement}${body}`;
          }
        }, newline);
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
        const orderChanged = finalOrder.length === segments.length && finalOrder.some((seg, idx) => seg !== segments[idx]);
        if (!orderChanged) return;
        const ucIndent = getLineIndent(updated, uc.open) || '';
        const entryIndent = ucIndent + '\t';
        const chunks = finalOrder.map((seg, idx) => {
          const needsComma = idx < finalOrder.length - 1;
          const rawSegment = updated.slice(seg.start, seg.end).trim();
          const normalizedSegment = reindent(rawSegment, entryIndent, newline);
          return ensureSegmentComma(normalizedSegment, needsComma);
        });
        const inner = chunks.length
          ? `${newline}${chunks.join(newline)}${newline}${ucIndent}`
          : '';
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
        segments.push({
          name: normalizeId(rawName),
          start: leadingStart,
          end,
          blockOpen,
          blockClose,
        });
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
          const body = updated.slice(seg.blockOpen + 1, seg.blockClose);
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
                  let start = ucPos;
                  while (start > toolOpen && /[ \t]/.test(text[start - 1])) start--;
                  if (start > toolOpen && (text[start - 1] === '\n' || text[start - 1] === '\r')) start--;
                  let end = ucClose + 1;
                  if (text[end] === ',') end++;
                  while (end < text.length && /\s/.test(text[end])) {
                    const c = text[end];
                    end++;
                    if (c === '\n' || c === '\r') break;
                  }
                  removals.push({ start, end });
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
      if (/=\s*(GroupOperator|MacroOperator)\s*\{/.test(text)) return text;
      const toolsPos = text.indexOf('Tools = ordered()');
      if (toolsPos < 0) return text;
      const ucPos = text.indexOf('UserControls = ordered()');
      if (ucPos < 0 || ucPos > toolsPos) return text;
      const ucGroup = findEnclosingGroupForIndex(text, ucPos);
      if (ucGroup) return text;
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
      if (/=\s*(GroupOperator|MacroOperator)\s*\{/.test(text)) return text;
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

  function applyDisplayNameOverrides(text, result, bounds, eol, resultRef, options = {}) {
    try {
      if (!result || !Array.isArray(result.entries)) return { text, bounds };
      const commitChanges = options.commitChanges !== false;
      let updated = text;
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      for (const entry of result.entries) {
        if (!entry || !entry.sourceOp || !entry.source) continue;
        if (isGroupUserControlEntry(entry)) {
          const desired = entry.displayName || entry.displayNameOriginal;
          const original = entry.displayNameOriginal || desired;
          if (!desired || desired === original) {
            if (commitChanges) entry.displayNameDirty = false;
            continue;
          }
          let currentBoundsForGroup = currentBounds || locateMacroGroupBounds(updated, resultRef);
          if (!currentBoundsForGroup) continue;
          const block = findGroupUserControlBlockById(updated, resultRef, entry.source);
          if (!block) continue;
          const indent = (getLineIndent(updated, block.open) || '') + '\t';
          const newline = eol || '\n';
          let body = updated.slice(block.open + 1, block.close);
          if (desired) body = upsertControlProp(body, 'LINKS_Name', `"${escapeQuotes(desired)}"`, indent, newline);
          else body = removeControlProp(body, 'LINKS_Name');
          updated = updated.slice(0, block.open + 1) + body + updated.slice(block.close);
          currentBounds = locateMacroGroupBounds(updated, resultRef) || currentBoundsForGroup;
          if (commitChanges) {
            entry.displayNameOriginal = desired;
            entry.displayNameDirty = false;
          }
          continue;
        }
        if (entry.isLabel) {
          const isHeaderCarrier = entry.source === FMR_HEADER_LABEL_CONTROL;
          const fallbackText = entry.displayName || entry.displayNameOriginal || entry.name || entry.source || `${entry.sourceOp}.${entry.source}`;
          const nameText = isHeaderCarrier ? '' : fallbackText;
          const style = normalizeLabelStyle(entry.labelStyle);
          const originalStyle = entry.labelStyleOriginal ? normalizeLabelStyle(entry.labelStyleOriginal) : style;
          const styleChanged = !!entry.labelStyleEdited || !!entry.labelStyleDirty || !labelStyleEquals(style, originalStyle);
          const nameChanged = isHeaderCarrier
            ? String(fallbackText || '').trim().length > 0 || !!entry.displayNameDirty
            : !!entry.displayNameDirty;
          if (!styleChanged && !nameChanged) {
            if (commitChanges) {
              entry.displayNameDirty = false;
              entry.labelStyleDirty = false;
              entry.labelStyleEdited = false;
            }
            continue;
          }
          const desiredMarkup = buildLabelMarkup(nameText, style);
          const res = rewriteToolControlDisplayName(
            updated,
            currentBounds,
            entry.sourceOp,
            entry.source,
            desiredMarkup,
            eol || '\n',
            resultRef,
            { headerCarrier: isHeaderCarrier }
          );
          updated = res.text;
          currentBounds = res.bounds || currentBounds;
          if (!res.applied) continue;
          if (commitChanges) {
            entry.displayNameOriginal = nameText;
            entry.displayNameDirty = false;
            entry.labelStyleOriginal = { ...style };
            entry.labelStyleDirty = false;
            entry.labelStyleEdited = false;
          }
          continue;
        }
        const desired = entry.displayName || entry.displayNameOriginal;
        const original = entry.displayNameOriginal || desired;
        if (!desired || desired === original) {
          if (commitChanges) {
            entry.displayNameDirty = false;
          }
          continue;
        }
        const res = rewriteToolControlDisplayName(updated, currentBounds, entry.sourceOp, entry.source, desired, eol || '\n', resultRef);
        updated = res.text;
        currentBounds = res.bounds || currentBounds;
        if (!res.applied) continue;
        if (commitChanges) {
          entry.displayNameOriginal = desired;
          entry.displayNameDirty = false;
        }
      }
      return { text: updated, bounds: currentBounds };
    } catch (_) {
      return { text, bounds };
    }
  }

  function rewriteToolControlDisplayName(text, bounds, toolName, controlName, label, eol, resultRef, options = {}) {
    try {
      let currentBounds = bounds || locateMacroGroupBounds(text, resultRef);
      if (!currentBounds) return { text, bounds, applied: false };
      const isHeaderCarrier = !!options.headerCarrier || String(controlName || '') === FMR_HEADER_LABEL_CONTROL;
      if (isHeaderCarrier) {
        const carrierBlock = findHeaderCarrierControlBlockInText(text, resultRef, toolName);
        if (!carrierBlock) return { text, bounds: currentBounds, applied: false };
        const indent = (getLineIndent(text, carrierBlock.open) || '') + '\t';
        const newline = eol || '\n';
        let body = text.slice(carrierBlock.open + 1, carrierBlock.close);
        const markup = buildHeaderCarrierImageMarkup(label);
        const markupValue = markup ? `"${escapeQuotes(markup)}"` : null;
        body = upsertControlProp(body, 'LINKS_Name', markupValue || '""', indent, newline);
        body = upsertControlProp(body, 'INP_Integer', 'false', indent, newline);
        body = upsertControlProp(body, 'INPID_InputControl', '"LabelControl"', indent, newline);
        body = upsertControlProp(body, 'LBLC_MultiLine', 'true', indent, newline);
        body = upsertControlProp(body, 'INP_External', 'false', indent, newline);
        body = upsertControlProp(body, 'LINKID_DataType', '"Number"', indent, newline);
        body = upsertControlProp(body, 'IC_NoReset', 'true', indent, newline);
        body = upsertControlProp(body, 'INP_Passive', 'true', indent, newline);
        body = upsertControlProp(body, 'IC_NoLabel', 'true', indent, newline);
        body = upsertControlProp(body, 'IC_ControlPage', '-1', indent, newline);
        body = removeControlProp(body, 'INP_Default');
        body = removeControlProp(body, 'INP_SplineType');
        body = removeControlProp(body, 'LBLC_NumInputs');
        body = removeControlProp(body, 'LBLC_DropDownButton');
        body = removeControlProp(body, 'IC_Visible');
        body = removeControlProp(body, 'ICS_ControlPage');
        const updated = text.slice(0, carrierBlock.open + 1) + body + text.slice(carrierBlock.close);
        const newBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
        return { text: updated, bounds: newBounds, applied: true };
      }
      const tb = findToolBlockInGroup(text, currentBounds.groupOpenIndex, currentBounds.groupCloseIndex, toolName);
      if (!tb) return { text, bounds: currentBounds, applied: false };
      let uc = findUserControlsInTool(text, tb.open, tb.close);
      if (!uc) {
        const ensured = ensureToolUserControlsBlock(text, currentBounds, toolName, eol);
        text = ensured.text;
        currentBounds = locateMacroGroupBounds(text, resultRef) || currentBounds;
        uc = ensured.ucBlock || (tb ? findUserControlsInTool(text, tb.open, tb.close) : null);
        if (!uc) return { text, bounds: currentBounds, applied: false };
      }
      let ensuredBlock = ensureControlBlockInUserControls(text, uc, controlName, eol);
      text = ensuredBlock.text;
      if (ensuredBlock.uc) uc = ensuredBlock.uc;
      const block = ensuredBlock.block || findControlBlockInUc(text, uc.open, uc.close, controlName);
      if (!block) return { text, bounds: currentBounds, applied: false };
      const indent = (getLineIndent(text, block.open) || '') + '\t';
      let body = text.slice(block.open + 1, block.close);
      const newline = eol || '\n';
      if (label) body = upsertControlProp(body, 'LINKS_Name', `"${escapeQuotes(label)}"`, indent, newline);
      else body = removeControlProp(body, 'LINKS_Name');
      const updated = text.slice(0, block.open + 1) + body + text.slice(block.close);
      const newBounds = locateMacroGroupBounds(updated, resultRef) || currentBounds;
      return { text: updated, bounds: newBounds, applied: true };
    } catch (_) {
      return { text, bounds, applied: false };
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

  function normalizeControlMetaDefaultLiteral(value) {
    try {
      if (value == null) return null;
      let out = String(value).trim();
      if (!out) return null;
      if (out.startsWith('"') && out.endsWith('"')) out = out.slice(1, -1);
      if (/^true$/i.test(out)) return '1';
      if (/^false$/i.test(out)) return '0';
      return out;
    } catch (_) {
      return null;
    }
  }

  function ensureCheckboxMetaProps(body, diff, entry, indent, eol) {
    try {
      let nextBody = body;
      const hasControlProp = (prop) => {
        try { return new RegExp(`${prop}\\s*=`, 'i').test(nextBody); } catch (_) { return false; }
      };
      const inputControl = Object.prototype.hasOwnProperty.call(diff, 'inputControl')
        ? diff.inputControl
        : extractControlPropValue(nextBody, 'INPID_InputControl');
      const normalizedInput = normalizeInputControlValue(inputControl);
      if (!normalizedInput || normalizedInput.toLowerCase() !== 'checkboxcontrol') return nextBody;
      const fallbackDefault = normalizeControlMetaDefaultLiteral(
        Object.prototype.hasOwnProperty.call(diff, 'defaultValue')
          ? diff.defaultValue
          : (entry?.controlMeta?.defaultValue ?? entry?.controlMetaOriginal?.defaultValue ?? extractInstancePropValue(entry.raw, 'Default') ?? '0')
      ) || '0';
      nextBody = upsertControlProp(nextBody, 'INPID_InputControl', '"CheckboxControl"', indent, eol);
      if (!hasControlProp('INP_Integer')) nextBody = upsertControlProp(nextBody, 'INP_Integer', 'false', indent, eol);
      if (!hasControlProp('INP_Default')) nextBody = upsertControlProp(nextBody, 'INP_Default', fallbackDefault, indent, eol);
      if (!hasControlProp('INP_MinScale')) nextBody = upsertControlProp(nextBody, 'INP_MinScale', '0', indent, eol);
      if (!hasControlProp('INP_MaxScale')) nextBody = upsertControlProp(nextBody, 'INP_MaxScale', '1', indent, eol);
      if (!hasControlProp('INP_MinAllowed')) nextBody = upsertControlProp(nextBody, 'INP_MinAllowed', '0', indent, eol);
      if (!hasControlProp('INP_MaxAllowed')) nextBody = upsertControlProp(nextBody, 'INP_MaxAllowed', '1', indent, eol);
      if (!hasControlProp('CBC_TriState')) nextBody = upsertControlProp(nextBody, 'CBC_TriState', 'false', indent, eol);
      if (!hasControlProp('LINKID_DataType')) nextBody = upsertControlProp(nextBody, 'LINKID_DataType', '"Number"', indent, eol);
      return nextBody;
    } catch (_) {
      return body;
    }
  }

  function applyControlMetaPropsToBody(body, diff, entry, indent, eol) {
    try {
      let nextBody = body;
      if (Object.prototype.hasOwnProperty.call(diff, 'defaultValue')) {
        const val = diff.defaultValue;
        if (val == null) nextBody = removeControlProp(nextBody, 'INP_Default');
        else nextBody = upsertControlProp(nextBody, 'INP_Default', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'minScale')) {
        const val = diff.minScale;
        if (val == null) nextBody = removeControlProp(nextBody, 'INP_MinScale');
        else nextBody = upsertControlProp(nextBody, 'INP_MinScale', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'maxScale')) {
        const val = diff.maxScale;
        if (val == null) nextBody = removeControlProp(nextBody, 'INP_MaxScale');
        else nextBody = upsertControlProp(nextBody, 'INP_MaxScale', val, indent, eol);
      }
      if (Object.prototype.hasOwnProperty.call(diff, 'inputControl')) {
        const val = diff.inputControl;
        if (val == null) nextBody = removeControlProp(nextBody, 'INPID_InputControl');
        else nextBody = upsertControlProp(nextBody, 'INPID_InputControl', `"${escapeQuotes(val)}"`, indent, eol);
      }
      nextBody = ensureCheckboxMetaProps(nextBody, diff, entry, indent, eol);
      return nextBody;
    } catch (_) {
      return body;
    }
  }

  function applyControlMetaOverride(text, entry, diff, bounds, eol, resultRef) {
    try {
      if (!entry || !entry.sourceOp || !entry.source || !diff) return { text, bounds };
      if (isGroupUserControlEntry(entry)) {
        const updated = upsertGroupUserControl(text, resultRef, entry.source, {
          createIfMissing: false,
          updateBody: (body, indent, newline) => applyControlMetaPropsToBody(body, diff, entry, indent, newline),
        }, eol || '\n');
        const newBounds = locateMacroGroupBounds(updated, resultRef) || bounds || locateMacroGroupBounds(text, resultRef);
        return { text: updated, bounds: newBounds };
      }
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
      body = applyControlMetaPropsToBody(body, diff, entry, indent, eol);
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
        const bodyStart = uc.openIndex + 1;
        let insertPos = uc.closeIndex;
        let cursor = uc.closeIndex - 1;
        while (cursor >= bodyStart && /\s/.test(updatedText[cursor])) cursor--;
        if (cursor >= bodyStart && updatedText[cursor] !== ',' && updatedText[cursor] !== '{') {
          updatedText = updatedText.slice(0, cursor + 1) + ',' + updatedText.slice(cursor + 1);
          insertPos += 1;
        }
        const insertion = eol + block;
        updatedText = updatedText.slice(0, insertPos) + insertion + updatedText.slice(insertPos);
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
      try {
        if (typeof logDiag === 'function' && diagnosticsController?.isEnabled?.()) {
          const macroName = String(resultRef?.macroNameOriginal || resultRef?.macroName || '').trim() || 'unknown';
          const stack = String(new Error().stack || '')
            .split(/\r?\n/)
            .slice(2, 6)
            .map((line) => line.trim())
            .join(' | ');
          logDiag(`[Group UC create] macro=${macroName} stack=${stack}`);
        }
      } catch (_) {}
      const indent = (getLineIndent(text, currentBounds.groupOpenIndex) || '') + '\t';
      const snippet = `${eol}${indent}UserControls = ordered() {${eol}${indent}},${eol}`;
      const insertPos = currentBounds.groupCloseIndex;
      const updated = text.slice(0, insertPos) + snippet + text.slice(insertPos);
      const ucOpenRel = snippet.indexOf('{');
      const ucOpen = ucOpenRel >= 0 ? insertPos + ucOpenRel : -1;
      const ucClose = ucOpen >= 0 ? findMatchingBrace(updated, ucOpen) : -1;
      const newBounds = locateMacroGroupBounds(updated, resultRef) || {
        groupOpenIndex: currentBounds.groupOpenIndex,
        groupCloseIndex: currentBounds.groupCloseIndex + snippet.length,
      };
      const newUc = (ucOpen >= 0 && ucClose > ucOpen)
        ? { openIndex: ucOpen, closeIndex: ucClose }
        : findGroupUserControlsBlock(updated, newBounds.groupOpenIndex, newBounds.groupCloseIndex);
      return { text: updated, bounds: newBounds, block: newUc };
    } catch (_) {
      return { text, bounds, block: null };
    }
  }

  function normalizeGroupUserControlsBlocks(text, result, eol) {
    try {
      if (!text) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const blocks = findAllGroupUserControlsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (blocks.length <= 1) return text;
      const newline = eol || '\n';
      const mergedBodies = [];
      blocks.forEach((block) => {
        const body = text.slice(block.openIndex + 1, block.closeIndex).trim();
        if (body) mergedBodies.push(body);
      });
      const baseIndent = (getLineIndent(text, bounds.groupOpenIndex) || '') + '\t';
      let replacement = `${newline}${baseIndent}UserControls = ordered() {`;
      if (mergedBodies.length) {
        replacement += `${newline}${mergedBodies.join(newline)}`;
        if (!replacement.endsWith(newline)) replacement += newline;
      } else {
        replacement += newline;
      }
      replacement += `${baseIndent}},${newline}`;
      const first = blocks[0];
      const last = blocks[blocks.length - 1];
      return text.slice(0, first.declStart) + replacement + text.slice(last.endIndex);
    } catch (_) {
      return text;
    }
  }

  function buildGroupUserControlCreateLinesFromEntry(entry, pageName) {
    try {
      const safePage = normalizePageNameGlobal(pageName || 'Controls');
      const displayName = String(entry?.displayName || entry?.name || humanizeName(entry?.source || 'Control') || 'Control').trim();
      const inputControl = normalizeInputControlValue(
        entry?.controlMeta?.inputControl || entry?.controlMetaOriginal?.inputControl
      ) || '';
      const lowerInput = inputControl.toLowerCase();
      let type = 'label';
      if (lowerInput === 'buttoncontrol') type = 'button';
      else if (lowerInput === 'separatorcontrol') type = 'separator';
      else if (lowerInput === 'slidercontrol') type = 'slider';
      else if (lowerInput === 'screwcontrol') type = 'screw';
      else if (lowerInput === 'combocontrol') type = 'combo';
      else if (lowerInput === 'multibuttoncontrol') type = 'combo';
      const rawDefault = normalizeMetaValue(
        entry?.controlMeta?.defaultValue ??
        entry?.controlMetaOriginal?.defaultValue ??
        (entry?.isLabel ? '0' : null)
      );
      const normalizedDefault = String(rawDefault || '').trim().toLowerCase();
      const labelDefault = (normalizedDefault === '1' || normalizedDefault === 'open' || normalizedDefault === 'true')
        ? 'open'
        : 'closed';
      const choiceOptions = Array.isArray(entry?.controlMeta?.choiceOptions) && entry.controlMeta.choiceOptions.length
        ? [...entry.controlMeta.choiceOptions]
        : (Array.isArray(entry?.controlMetaOriginal?.choiceOptions) ? [...entry.controlMetaOriginal.choiceOptions] : []);
      let lines = buildControlDefinitionLines(type, {
        name: displayName,
        page: safePage,
        labelCount: Number.isFinite(entry?.labelCount) ? Number(entry.labelCount) : 0,
        labelDefault,
        comboOptions: choiceOptions,
      });
      if (inputControl) {
        lines = lines.filter((line) => !/^\s*INPID_InputControl\s*=/.test(String(line || '')));
        lines.push(`INPID_InputControl = "${escapeQuotes(inputControl)}",`);
      }
      if (Number.isFinite(entry?.labelCount) && !lines.some((line) => /^\s*LBLC_NumInputs\s*=/.test(String(line || '')))) {
        lines.push(`LBLC_NumInputs = ${Math.max(0, Number(entry.labelCount) || 0)},`);
      }
      return lines;
    } catch (_) {
      return [];
    }
  }

  function stripEmptyGroupUserControls(text, result) {
    try {
      if (!text) return text;
      const bounds = locateMacroGroupBounds(text, result);
      if (!bounds) return text;
      const blocks = findAllGroupUserControlsBlocks(text, bounds.groupOpenIndex, bounds.groupCloseIndex);
      if (!blocks.length) return text;
      const removable = blocks.filter((block) => {
        try {
          return text.slice(block.openIndex + 1, block.closeIndex).trim().length === 0;
        } catch (_) {
          return false;
        }
      });
      if (!removable.length) return text;
      let updated = text;
      removable.sort((a, b) => b.declStart - a.declStart);
      removable.forEach((block) => {
        let start = block.declStart;
        while (start > bounds.groupOpenIndex && /[ \t]/.test(updated[start - 1])) start--;
        if (start > bounds.groupOpenIndex && (updated[start - 1] === '\n' || updated[start - 1] === '\r')) start--;
        let end = block.endIndex;
        while (end < updated.length && /[ \t]/.test(updated[end])) end++;
        if (end < updated.length && (updated[end] === '\r' || updated[end] === '\n')) {
          const firstBreak = updated[end];
          end++;
          if (firstBreak === '\r' && updated[end] === '\n') end++;
        }
        updated = updated.slice(0, start) + updated.slice(end);
      });
      return updated;
    } catch (_) {
      return text;
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
      const groupBlock = findGroupUserControlBlockById(text, result, controlName);
      if (groupBlock) return text.slice(groupBlock.open + 1, groupBlock.close);
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

  function removeControlPropAll(body, prop) {
    try {
      let next = body;
      let guard = 0;
      while (guard < 500) {
        const range = findControlPropRange(next, prop);
        if (!range) break;
        next = next.slice(0, range.start) + next.slice(range.end);
        guard += 1;
      }
      return next;
    } catch (_) {
      return body;
    }
  }

  function removeControlListPropEntries(body, prop) {
    try {
      let next = String(body || '');
      const escapedProp = String(prop || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (!escapedProp) return next;
      const quoted = '"(?:\\\\.|[^"])*"';
      const bare = '[^,}\\r\\n]+';
      const value = `(?:${quoted}|${bare})`;
      const braced = new RegExp(`\\{\\s*${escapedProp}\\s*=\\s*${value}\\s*\\}\\s*,?\\s*`, 'ig');
      next = next.replace(braced, '');
      next = removeControlPropAll(next, prop);
      return next;
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
      if (body[cursor] === '[') {
        const longOpen = /^\[(=*)\[/.exec(body.slice(cursor));
        if (longOpen) {
          const eq = longOpen[1] || '';
          const closeToken = `]${eq}]`;
          const contentStart = cursor + longOpen[0].length;
          const closeIndex = body.indexOf(closeToken, contentStart);
          cursor = closeIndex >= 0 ? closeIndex + closeToken.length : body.length;
        } else {
          while (cursor < body.length && !/[,\r\n}]/.test(body[cursor])) cursor++;
        }
      } else if (body[cursor] === '"') {
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
        const identMatch = /^[A-Za-z_][A-Za-z0-9_]*/.exec(body.slice(cursor));
        if (identMatch) {
          const identEnd = cursor + identMatch[0].length;
          let probe = identEnd;
          while (probe < body.length && /\s/.test(body[probe])) probe++;
          if (probe < body.length && body[probe] === '{') {
            const close = findMatchingBrace(body, probe);
            cursor = close >= 0 ? close + 1 : body.length;
          } else if (probe < body.length && body[probe] === '(') {
            const close = findMatchingParen(body, probe);
            cursor = close >= 0 ? close + 1 : body.length;
          } else {
            while (cursor < body.length && !/[,\r\n}]/.test(body[cursor])) cursor++;
          }
        } else {
          while (cursor < body.length && !/[,\r\n}]/.test(body[cursor])) cursor++;
        }
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
      const topMatch = text.match(/Tools\s*=\s*ordered\(\)\s*\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{/s);
      if (topMatch && topMatch[1] && topMatch[2]) {
        const bounds = findGroupByName(text, String(topMatch[1]).trim(), String(topMatch[2]).trim());
        if (bounds) return bounds;
      }
      if (typeof result?.inputs?.openIndex === 'number') {
        const fallback = findEnclosingGroupForIndex(text, Math.max(0, Math.min(text.length - 1, result.inputs.openIndex)));
        if (fallback) return { groupOpenIndex: fallback.groupOpenIndex, groupCloseIndex: fallback.groupCloseIndex };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function resolveMacroGroupSourceOp(text, result) {
    try {
      if (text) {
        const bounds = locateMacroGroupBounds(text, result);
        if (bounds && Number.isFinite(bounds.groupOpenIndex)) {
          const lineStart = text.lastIndexOf('\n', bounds.groupOpenIndex);
          const head = text.slice(Math.max(0, lineStart + 1), bounds.groupOpenIndex + 1);
          const match = head.match(/\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{$/);
          if (match && match[1]) return String(match[1]).trim();
        }
        const topMatch = text.match(/Tools\s*=\s*ordered\(\)\s*\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(GroupOperator|MacroOperator)\s*\{/s);
        if (topMatch && topMatch[1]) return String(topMatch[1]).trim();
      }
    } catch (_) {}
    return String(result?.macroNameOriginal || result?.macroName || 'Macro').trim();
  }

  function findGroupUserControlRecord(result, controlId) {
    try {
      if (!result || !controlId) return null;
      const id = String(controlId).trim();
      if (!id) return null;
      const map = result.groupUserControlsById;
      if (map && typeof map.get === 'function') {
        const hit = map.get(id);
        if (hit) return hit;
      }
      const list = Array.isArray(result.groupUserControls) ? result.groupUserControls : [];
      return list.find((item) => item && String(item.id || '').trim() === id) || null;
    } catch (_) {
      return null;
    }
  }

  function syncGroupUserControlEntryBridge(entry, meta) {
    try {
      if (!entry || !state.parseResult) return entry;
      const requestedTarget = String(meta?.publishTarget || entry.publishTarget || '').trim();
      const isGroupOwned = requestedTarget === 'groupUserControl' || entry.syntheticGroupUserControl === true;
      if (!isGroupOwned) return entry;
      entry.publishTarget = 'groupUserControl';
      entry.syntheticGroupUserControl = true;
      entry.sourceOp = resolveMacroGroupSourceOp(state.originalText, state.parseResult);
      const backing = findGroupUserControlRecord(state.parseResult, entry.source);
      const normalizedPage = normalizePageName((meta && meta.page) || entry.page || backing?.page || 'Controls');
      entry.page = normalizedPage;
      if (backing) {
        backing.page = normalizedPage;
        const controlName = String(backing.name || '').trim();
        if (controlName) {
          entry.name = controlName;
          if (!entry.displayName || entry.displayName === `${entry.sourceOp}.${entry.source}` || entry.displayName === entry.source) {
            entry.displayName = controlName;
          }
        }
        const nextMeta = { ...(entry.controlMeta || {}) };
        const nextMetaOriginal = { ...(entry.controlMetaOriginal || {}) };
        const inputControl = String(backing.inputControl || meta?.inputControl || '').trim();
        if (inputControl) {
          nextMeta.inputControl = inputControl;
          nextMetaOriginal.inputControl = nextMetaOriginal.inputControl || inputControl;
        }
        const lowerInput = inputControl.toLowerCase();
        if (!entry.kind) {
          if (lowerInput === 'labelcontrol') entry.kind = 'Label';
          else if (lowerInput === 'buttoncontrol') entry.kind = 'Button';
          else if (lowerInput === 'separatorcontrol') entry.kind = 'Separator';
        }
        if (lowerInput === 'labelcontrol' || String(entry.kind || '').toLowerCase() === 'label') {
          entry.isLabel = true;
          if (!Number.isFinite(entry.labelCount)) entry.labelCount = 0;
        } else if (lowerInput === 'buttoncontrol' || String(entry.kind || '').toLowerCase() === 'button') {
          entry.isButton = true;
        }
        if (!nextMeta.kind && entry.kind) nextMeta.kind = String(entry.kind).toLowerCase();
        if (!nextMetaOriginal.kind && nextMeta.kind) nextMetaOriginal.kind = nextMeta.kind;
        const labelCount = Number.isFinite(backing.labelCount) ? Number(backing.labelCount) : (Number.isFinite(meta?.labelCount) ? Number(meta.labelCount) : null);
        if (Number.isFinite(labelCount)) {
          nextMeta.labelCount = labelCount;
          if (!Number.isFinite(nextMetaOriginal.labelCount)) nextMetaOriginal.labelCount = labelCount;
          entry.labelCount = labelCount;
        }
        const defaultValue = backing.defaultValue != null ? backing.defaultValue : (meta?.defaultValue != null ? meta.defaultValue : null);
        if (defaultValue != null) {
          nextMeta.defaultValue = defaultValue;
          if (nextMetaOriginal.defaultValue == null) nextMetaOriginal.defaultValue = defaultValue;
        }
        entry.controlMeta = nextMeta;
        entry.controlMetaOriginal = nextMetaOriginal;
      }
      return entry;
    } catch (_) {
      return entry;
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

    if (!entry) return entry?.raw;

    const raw = entry.raw;
    if (!raw) return raw;

    const open = raw.indexOf('{');
    if (open < 0) return raw;

    const close = findMatchingBrace(raw, open);
    if (close < 0) return raw;

    const head = raw.slice(0, open + 1);
    const body = raw.slice(open + 1, close);
    const tail = raw.slice(close);

    const rangeBoundLabel = (() => {
      const raw = String(entry?.source || '').trim();
      const m = raw.match(/(?:[._])(Min|Max)$/i);
      if (!m) return '';
      return ' ';
    })();
    const isRangeBound = !!rangeBoundLabel;

    if (entry.isLabel) {
      const cleaned = removeInstanceInputProp(body, 'Name');
      return head + cleaned + tail;
    }

    if (isRangeBound) {
      let normalized = removeInstanceInputProp(body, 'Name');
      const insert = `${eol}Name = "${escapeQuotes(rangeBoundLabel)}",`;
      return head + insert + normalized + tail;
    }

    if (!entry.name) return raw;

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
      const rangeBoundLabel = (() => {
        const rawSource = String(entry?.source || '').trim();
        const m = rawSource.match(/(?:[._])(Min|Max)$/i);
        if (!m) return '';
        return ' ';
      })();
      if (rangeBoundLabel) {
        const open = out.indexOf('{');
        const close = out.lastIndexOf('}');
        if (open >= 0 && close > open) {
          let body = out.slice(open + 1, close);
          body = removeInstanceInputProp(body, 'Name');
          body = setInstanceInputProp(body, 'Name', `"${escapeQuotes(rangeBoundLabel)}"`, getLineIndent(out, open + 1), newline);
          out = out.slice(0, open + 1) + body + out.slice(close);
        }
      }
    } catch (_) {}
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
    if (Number.isFinite(entry?.controlGroup)) {
      out = ensureInstanceInputControlGroup(out, entry.controlGroup, newline);
    }
    if (isBlendControl) logBlendDebug(entry, 'emit-default', { isToggle, inSet });
    return out;
  }

  function isKeyLikelyValid(key) {
    if (!key) return false;
    const trimmed = String(key).trim();
    if (!trimmed) return false;
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmed)) return true;
    if (/^\[\s*"(?:\\.|[^"\\])+"\s*\]$/.test(trimmed)) return true;
    return false;
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
    const fallback = formatToolInputId(deriveFallbackEntryKey(entry));
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

  function parseColorChannelMeta(id) {
    const raw = String(id || '').trim();
    if (!raw) return null;
    const suffix = raw.match(/^(.*?)(Red|Green|Blue|Alpha)$/i);
    if (suffix && suffix[1]) {
      return {
        base: suffix[1],
        channel: suffix[2].charAt(0).toUpperCase() + suffix[2].slice(1).toLowerCase(),
      };
    }
    const prefix = raw.match(/^(Red|Green|Blue|Alpha)(.*)$/i);
    if (prefix) {
      const tail = String(prefix[2] || '');
      if (!tail) {
        return {
          base: 'Color',
          channel: prefix[1].charAt(0).toUpperCase() + prefix[1].slice(1).toLowerCase(),
        };
      }
      if (/^\d+(clone)?$/i.test(tail)) {
        return {
          base: `Color${tail}`,
          channel: prefix[1].charAt(0).toUpperCase() + prefix[1].slice(1).toLowerCase(),
        };
      }
    }
    return null;
  }

  function parseControlBooleanValue(raw, fallback = false) {
    if (raw == null) return fallback;
    const cleaned = String(raw).replace(/"/g, '').trim().toLowerCase();
    if (!cleaned) return fallback;
    if (['true', 'yes', 'y', 'on', '1'].includes(cleaned)) return true;
    if (['false', 'no', 'n', 'off', '0'].includes(cleaned)) return false;
    return fallback;
  }

  function deriveColorBaseFromId(id) {

    const meta = parseColorChannelMeta(id);
    return meta ? meta.base : null;

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
      if (Array.isArray(meta.choiceOptions) && meta.choiceOptions.length) {
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        const normalizedChoices = meta.choiceOptions
          .map((opt) => String(opt || '').trim())
          .filter((opt) => opt.length > 0);
        if (normalizedChoices.length) {
          entry.controlMeta.choiceOptions = [...normalizedChoices];
          if (!Array.isArray(entry.controlMetaOriginal.choiceOptions) || !entry.controlMetaOriginal.choiceOptions.length) {
            entry.controlMetaOriginal.choiceOptions = [...normalizedChoices];
          }
        }
      }
      if (meta.multiButtonShowBasic != null && String(meta.multiButtonShowBasic).trim() !== '') {
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        const normalized = parseControlBooleanValue(meta.multiButtonShowBasic, false) ? 'true' : 'false';
        entry.controlMeta.multiButtonShowBasic = normalized;
        if (!entry.controlMetaOriginal.multiButtonShowBasic) {
          entry.controlMetaOriginal.multiButtonShowBasic = normalized;
        }
      }
      if (meta.defaultX != null) {
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        entry.controlMeta.defaultX = meta.defaultX;
        if (!entry.controlMetaOriginal.defaultX) entry.controlMetaOriginal.defaultX = meta.defaultX;
      }
      if (meta.defaultY != null) {
        entry.controlMeta = entry.controlMeta || {};
        entry.controlMetaOriginal = entry.controlMetaOriginal || {};
        entry.controlMeta.defaultY = meta.defaultY;
        if (!entry.controlMetaOriginal.defaultY) entry.controlMetaOriginal.defaultY = meta.defaultY;
      }
      if (Number.isFinite(meta.controlGroup) && meta.controlGroup > 0) {
        entry.controlGroup = Number(meta.controlGroup);
      }
      if (meta.locked === true) {
        entry.locked = true;
      }
      if (meta.syntheticToolUserControl === true) {
        entry.syntheticToolUserControl = true;
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
      let entry;
      if (meta && meta.publishTarget === 'groupUserControl') {
        entry = buildGroupUserControlEntry(state.parseResult, sourceOp, source, displayName, meta);
      } else {
        const key = makeUniqueKey(`${sourceOp}_${source}`);
        const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());
        const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;
        const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
        const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
        const targetPage = metaPage || activePage || 'Controls';
        const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);
        entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };
      }

      state.parseResult.entries.push(entry);
      idx = state.parseResult.entries.length - 1;
      state.parseResult.order.push(idx);

    }
    applyNodeControlMeta(state.parseResult.entries[idx], meta);
    syncGroupUserControlEntryBridge(state.parseResult.entries[idx], meta);

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

    if (idx >= 0) {
      const existing = state.parseResult.entries[idx];
      if (existing) {
        if (displayName && (!existing.displayName || existing.displayName === `${existing.sourceOp}.${existing.source}` || existing.displayName === existing.source)) {
          existing.displayName = displayName;
        }
        if (displayName && !existing.name) {
          existing.name = displayName;
        }
        applyNodeControlMeta(existing, meta);
        syncGroupUserControlEntryBridge(existing, meta);
      }
      return idx;
    }

    let entry;
    if (meta && meta.publishTarget === 'groupUserControl') {
      entry = buildGroupUserControlEntry(state.parseResult, sourceOp, source, displayName, meta);
    } else {
      const key = makeUniqueKey(`${sourceOp}_${source}`);
      const base = (typeof deriveColorBaseFromId === "function" ? deriveColorBaseFromId(source) : (function(){ const suf=["Red","Green","Blue","Alpha"].find(s => String(source||"").endsWith(s)); return suf ? String(source).slice(0, String(source).length - suf.length) : null; })());
      const controlGroup = base ? getOrAssignControlGroup(sourceOp, base) : null;
      const metaPage = (meta && meta.page && String(meta.page).trim()) ? String(meta.page).trim() : '';
      const activePage = (state.parseResult && state.parseResult.activePage) ? String(state.parseResult.activePage).trim() : '';
      const targetPage = metaPage || activePage || 'Controls';
      const raw = buildInstanceInputRaw(key, sourceOp, source, displayName, targetPage, controlGroup);
      entry = { key, name: displayName || null, page: targetPage, sourceOp, source, displayName: displayName || `${sourceOp}.${source}`, raw, controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null), onChange: '', buttonExecute: '' };
    }

    state.parseResult.entries.push(entry);
    idx = state.parseResult.entries.length - 1;
    state.parseResult.order.push(idx);
    applyNodeControlMeta(state.parseResult.entries[idx], meta);
    syncGroupUserControlEntryBridge(state.parseResult.entries[idx], meta);

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

    const safeKey = formatToolInputId(key);
    b.push(`${safeKey || key} = InstanceInput {`);

    if (sourceOp) b.push(`SourceOp = "${escapeQuotes(sourceOp)}",`);

    if (source) b.push(`Source = "${escapeQuotes(source)}",`);

    const rangeBoundLabel = (() => {
      const raw = String(source || '').trim();
      const m = raw.match(/(?:[._])(Min|Max)$/i);
      if (!m) return '';
      return ' ';
    })();
    const effectiveDisplayName = rangeBoundLabel || displayName;
    if (effectiveDisplayName) b.push(`Name = "${escapeQuotes(effectiveDisplayName)}",`);

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

    const safeKey = formatToolInputId(key);
    segments.push(`${safeKey || key} = InstanceInput {`);

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

    renderActiveList();

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

      getEntry: (idx) => {
        try {
          if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return null;
          return state.parseResult.entries[idx] || null;
        } catch (_) { return null; }
      },

      getCsvState: () => {
        try {
          return {
            hasCsv: !!(state.csvData && Array.isArray(state.csvData.headers) && state.csvData.headers.length),
            source: state.csvData?.sourceName || '',
            headers: state.csvData?.headers || [],
          };
        } catch (_) { return null; }
      },

      testCsvCoerce: (idx, raw) => {
        try {
          const entry = state.parseResult?.entries?.[idx];
          if (!entry) return null;
          return {
            inputControl: entry?.controlMeta?.inputControl,
            dataType: entry?.controlMeta?.dataType,
            raw,
            coerced: coerceCsvNumericValue(entry, raw),
          };
        } catch (_) { return null; }
      },
      findTextLinesInfo: (idx) => {
        try {
          const entry = state.parseResult?.entries?.[idx];
          if (!entry || !state.originalText) return null;
          const bounds = locateMacroGroupBounds(state.originalText, state.parseResult) || { groupOpenIndex: 0, groupCloseIndex: state.originalText.length };
          let toolBlock = findToolBlockInGroup(state.originalText, bounds.groupOpenIndex, bounds.groupCloseIndex, entry.sourceOp);
          if (!toolBlock) toolBlock = findToolBlockAnywhere(state.originalText, entry.sourceOp);
          if (!toolBlock) return { toolBlock: false };
          const toolSegment = state.originalText.slice(toolBlock.open, toolBlock.close);
          let controlBlock = findControlBlockInTool(state.originalText, toolBlock.open, toolBlock.close, entry.source);
          if (!controlBlock) return { toolBlock: true, controlBlock: false, toolSegmentStart: toolSegment.slice(0, 200) };
          const body = state.originalText.slice(controlBlock.open + 1, controlBlock.close);
          return {
            toolBlock: true,
            controlBlock: true,
            hasTextEdit: /TextEditControl/i.test(body),
            tecLines: extractControlPropValue(body, 'TEC_Lines'),
            bodyStart: body.slice(0, 200),
          };
        } catch (err) { return { error: err?.message || String(err) }; }
      },

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
      const msg = 'Copy the .setting content below (Ctrl/Cmd+A, then Ctrl/Cmd+C)';
      const joined = String(text || '');
      await textPrompt.open({
        title: 'Copy .setting content',
        label: msg,
        initialValue: joined,
        confirmText: 'Close',
        multiline: true,
      });
      return;

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

  // Final safety: directly locate control by id in any UserControls range and inject snippet
  function applyCanonicalLaunchSnippets(text, result, eol) {
    try {
      const overrides = (result && result.buttonOverrides && typeof result.buttonOverrides.forEach === 'function') ? result.buttonOverrides : null;
      if (!overrides || result.buttonOverrides.size === 0) return text;
      let out = text;
      const groupBounds = locateMacroGroupBounds(out, result);
      const groupUc = groupBounds ? findUserControlsInGroup(out, groupBounds.groupOpenIndex, groupBounds.groupCloseIndex) : null;
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
              const isGroupLevel = !!groupUc && rg.open === groupUc.open && rg.close === groupUc.close;
              if (isGroupLevel) {
                const policy = canMutateGroupUserControl(result, ctrl, { action: 'update' });
                if (!policy.allowed) {
                  recordGroupUserControlMutationIssue(result, {
                    control: ctrl,
                    action: 'update',
                    message: policy.message,
                  });
                  break;
                }
              }
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
  let validationTimer = null;
  let pendingValidationContext = null;

  function runValidation(context) {
    try {
      const entries = state.parseResult?.entries || [];
      const isLarge = entries.length > 250;
      if (isLarge && context !== 'manual') {
        pendingValidationContext = context || 'auto';
        if (validationTimer) clearTimeout(validationTimer);
        validationTimer = setTimeout(() => {
          validationTimer = null;
          runValidationNow(pendingValidationContext);
          pendingValidationContext = null;
        }, 250);
        return [];
      }
    } catch (_) {}
    return runValidationNow(context);
  }

  function findControlBlockInTool(text, toolOpen, toolClose, controlId) {
    try {
      let i = toolOpen + 1, depth = 0, inStr = false;
      while (i < toolClose) {
        const ch = text[i];
        if (inStr) { if (ch === '"' && text[i - 1] !== '\\') inStr = false; i++; continue; }
        if (ch === '"') { inStr = true; i++; continue; }
        if (ch === '{') { depth++; i++; continue; }
        if (ch === '}') { depth--; i++; continue; }
        if (depth === 0 && isIdentStart(ch)) {
          const idStart = i;
          i++;
          while (i < toolClose && isIdentPart(text[i])) i++;
          const idStr = text.slice(idStart, i);
          while (i < toolClose && isSpace(text[i])) i++;
          if (text[i] !== '=') { i++; continue; }
          i++;
          while (i < toolClose && isSpace(text[i])) i++;
          if (text[i] !== '{') { i++; continue; }
          const cOpen = i;
          const cClose = findMatchingBrace(text, cOpen);
          if (cClose < 0) break;
          if (String(idStr) === String(controlId)) {
            return { open: cOpen, close: cClose };
          }
          i = cClose + 1;
          continue;
        }
        i++;
      }
      return null;
    } catch (_) { return null; }
  }

  function runValidationNow(context) {
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
      const groupUcIssues = state.parseResult.groupUserControlIssues || { blockCount: 0, duplicateIds: [] };
      const groupUcMutationIssues = Array.isArray(state.parseResult.groupUserControlMutationIssues)
        ? state.parseResult.groupUserControlMutationIssues
        : [];
      if (Number(groupUcIssues.blockCount || 0) > 1) {
        issues.push({
          type: 'group-user-controls',
          message: `Macro has ${groupUcIssues.blockCount} top-level UserControls blocks. Expected exactly 1.`,
        });
      }
      (Array.isArray(groupUcIssues.duplicateIds) ? groupUcIssues.duplicateIds : []).forEach((id) => {
        issues.push({
          type: 'group-user-control-id',
          control: id,
          message: `Macro-level UserControls contains duplicate control id "${id}".`,
        });
      });
      (Array.isArray(groupUcIssues.unknownControls) ? groupUcIssues.unknownControls : []).forEach((item) => {
        try {
          logDiag(`[Group UC] Preserve-only control "${item.id}" type="${item.inputControl || 'unknown'}"`);
        } catch (_) {}
      });
      groupUcMutationIssues.forEach((item) => {
        if (!item || !item.message) return;
        issues.push({
          type: 'group-user-control-mutation',
          control: item.control,
          message: item.message,
        });
      });
      for (const entry of (state.parseResult.entries || [])) {
        if (!entry || !isMacroLayerOnlyControlId(entry.source)) continue;
        const leaked = isMacroControlLeakedIntoInputs(state.originalText || state.parseResult.originalText || '', entry.source);
        if (!leaked) continue;
        issues.push({
          type: 'macro-control-leaked-into-inputs',
          control: entry.source,
          message: `Macro-layer control "${entry.source}" is present in published Inputs/entries and should remain group-owned.`,
        });
      }
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
      if (findGroupUserControlBlockById(text, state.parseResult, controlId)) return true;
      return false;
    } catch (_) {
      return false;
    }
  }

})();
