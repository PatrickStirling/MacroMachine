export function createHistoryController({
  state,
  renderList,
  ctrlCountEl,
  getNodesPane = () => null,
  undoBtn,
  redoBtn,
  logDiag = () => {},
  doc,
  captureExtraState = () => ({}),
  restoreExtraState = () => {},
}) {
  const targetDoc = doc || (typeof document !== 'undefined' ? document : null);

  function cloneEntries(entries) {
    return (entries || []).map(e => (e ? { ...e } : e));
  }

  function snapshotState() {
    if (!state.parseResult) return null;
    return {
      entries: cloneEntries(state.parseResult.entries),
      order: [...(state.parseResult.order || [])],
      originalOrder: [...(state.parseResult.originalOrder || [])],
      selected: Array.from(state.parseResult.selected || []),
      collapsed: Array.from(state.parseResult.collapsed || []),
      collapsedCG: Array.from(state.parseResult.collapsedCG || []),
      nodesCollapsed: Array.from(state.parseResult.nodesCollapsed || []),
      nodesPublishedOnly: Array.from(state.parseResult.nodesPublishedOnly || []),
      nodesViewInitialized: !!state.parseResult.nodesViewInitialized,
      cgNext: state.parseResult.cgNext,
      cgMap: Array.from((state.parseResult.cgMap || new Map()).entries()),
      blendToggles: Array.from(state.parseResult.blendToggles || []),
      pageOrder: Array.from(state.parseResult.pageOrder || []),
      activePage: state.parseResult.activePage || null,
      macroName: state.parseResult.macroName || null,
      macroNameOriginal: state.parseResult.macroNameOriginal || null,
      operatorType: state.parseResult.operatorType || null,
      operatorTypeOriginal: state.parseResult.operatorTypeOriginal || null,
      extra: captureExtraState(state),
    };
  }

  function restoreState(snap, context) {
    if (!state.parseResult || !snap) return;
    state.parseResult.entries = cloneEntries(snap.entries);
    state.parseResult.order = [...snap.order];
    state.parseResult.originalOrder = [...snap.originalOrder];
    state.parseResult.selected = new Set(snap.selected || []);
    state.parseResult.collapsed = new Set(snap.collapsed || []);
    state.parseResult.collapsedCG = new Set(snap.collapsedCG || []);
    state.parseResult.nodesCollapsed = new Set(snap.nodesCollapsed || []);
    state.parseResult.nodesPublishedOnly = new Set(snap.nodesPublishedOnly || []);
    state.parseResult.nodesViewInitialized = !!snap.nodesViewInitialized;
    state.parseResult.cgNext = snap.cgNext;
    state.parseResult.cgMap = new Map(snap.cgMap || []);
    state.parseResult.blendToggles = new Set(snap.blendToggles || []);
    state.parseResult.pageOrder = Array.isArray(snap.pageOrder) ? [...snap.pageOrder] : [];
    state.parseResult.activePage = snap.activePage || null;
    state.parseResult.macroName = snap.macroName || state.parseResult.macroName;
    state.parseResult.macroNameOriginal = snap.macroNameOriginal || state.parseResult.macroNameOriginal;
    state.parseResult.operatorType = snap.operatorType || state.parseResult.operatorType;
    state.parseResult.operatorTypeOriginal = snap.operatorTypeOriginal || state.parseResult.operatorTypeOriginal;
    if (ctrlCountEl) ctrlCountEl.textContent = String(state.parseResult.entries.length);
    renderList(state.parseResult.entries, state.parseResult.order);
    const pane = typeof getNodesPane === 'function' ? getNodesPane() : null;
    if (pane && typeof pane.parseAndRenderNodes === 'function') {
      pane.parseAndRenderNodes();
    }
    restoreExtraState(snap.extra || {}, context || 'restore');
  }

  function updateUndoRedoState() {
    if (!undoBtn || !redoBtn) return;
    const h = (state.parseResult && state.parseResult.history) ? state.parseResult.history.length : 0;
    const f = (state.parseResult && state.parseResult.future) ? state.parseResult.future.length : 0;
    undoBtn.disabled = h === 0;
    redoBtn.disabled = f === 0;
  }

  function pushHistory(label) {
    try {
      if (!state.parseResult) return;
      if (!state.parseResult.history) state.parseResult.history = [];
      if (!state.parseResult.future) state.parseResult.future = [];
      const snap = snapshotState();
      state.parseResult.history.push(snap);
      state.parseResult.future = [];
      updateUndoRedoState();
      if (label) logDiag(`History: ${label}`);
    } catch (_) {}
  }

  function undo() {
    if (!state.parseResult || !state.parseResult.history || state.parseResult.history.length === 0) return;
    const current = snapshotState();
    const prev = state.parseResult.history.pop();
    if (!state.parseResult.future) state.parseResult.future = [];
    state.parseResult.future.push(current);
    restoreState(prev, 'undo');
    updateUndoRedoState();
  }

  function redo() {
    if (!state.parseResult || !state.parseResult.future || state.parseResult.future.length === 0) return;
    const current = snapshotState();
    const next = state.parseResult.future.pop();
    if (!state.parseResult.history) state.parseResult.history = [];
    state.parseResult.history.push(current);
    restoreState(next, 'redo');
    updateUndoRedoState();
  }

  undoBtn?.addEventListener('click', () => undo());
  redoBtn?.addEventListener('click', () => redo());

  const keyHandler = (ev) => {
    const key = ev.key && ev.key.toLowerCase();
    if (!key) return;
    const redoCombo = (ev.ctrlKey || ev.metaKey) && (key === 'y' || (key === 'z' && ev.shiftKey));
    const undoCombo = (ev.ctrlKey || ev.metaKey) && key === 'z' && !ev.shiftKey;
    if (undoCombo) { ev.preventDefault(); undo(); }
    if (redoCombo) { ev.preventDefault(); redo(); }
  };

  if (targetDoc) targetDoc.addEventListener('keydown', keyHandler);

  return {
    pushHistory,
    undo,
    redo,
    snapshotState,
    restoreState,
    updateUndoRedoState,
    dispose: () => {
      if (targetDoc) targetDoc.removeEventListener('keydown', keyHandler);
    },
  };
}
