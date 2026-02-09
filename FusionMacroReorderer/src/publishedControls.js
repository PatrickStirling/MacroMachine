import { getLineIndent, findMatchingBrace, isSpace, isIdentStart, isIdentPart, humanizeName } from './textUtils.js';
import { buildLabelMarkup, normalizeLabelStyle } from './labelMarkup.js';

export function createPublishedControls({
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
  refreshNodesChecks,
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
  getCurrentValueInfo,
  shouldShowCurrentValues,
  onDetailTargetChange,
  applyNodeControlMeta = () => {},
  onRenderList = null,
  onEntryMutated = null,
}) {
  let publishedFilter = '';
  let highlightHandler = typeof highlightNode === 'function' ? highlightNode : () => {};
  let insertionAnchorPos = null;
  let dragFromIndex = null;
  let currentDropIndex = null;
  const dropIndicator = document.createElement('li');
  dropIndicator.className = 'drop-indicator';
  const manualPageOptions = new Set();
  const PAGE_ICON_OPTIONS = [
    { id: '', label: 'Default (None)' },
    { id: 'Icons.Tools.Tabs.Controls', label: 'Controls' },
    { id: 'Icons.Tools.Tabs.Custom', label: 'Custom' },
    { id: 'Icons.Tools.Tabs.Text', label: 'Text' },
    { id: 'Icons.Tools.Tabs.Layout', label: 'Layout' },
    { id: 'Icons.Tools.Tabs.Transform', label: 'Transform' },
    { id: 'Icons.Tools.Tabs.Style', label: 'Style' },
    { id: 'Icons.Tools.Tabs.Image', label: 'Image' },
    { id: 'Icons.Tools.Tabs.Common', label: 'Common' },
    { id: 'Icons.Tools.Tabs.Color', label: 'Color' },
    { id: 'Icons.Tools.Tabs.Blur', label: 'Blur' },
    { id: 'Icons.Tools.Tabs.Merge', label: 'Merge' },
    { id: 'Icons.Tools.Tabs.Channels', label: 'Channels' },
    { id: 'Icons.Tools.Tabs.Noise', label: 'Noise' },
    { id: 'Icons.Tools.Tabs.Tracker', label: 'Tracker' },
    { id: 'Icons.Tools.Tabs.Circles', label: 'Circles' },
    { id: 'Icons.Tools.Tabs.Operation', label: 'Operation' },
    { id: 'Icons.Tools.Tabs.Display', label: 'Display' },
    { id: 'Icons.Tools.Tabs.XYZ', label: 'XYZ' },
    { id: 'Icons.Tools.Tabs.Material', label: 'Material' },
    { id: 'Icons.Tools.Tabs.Shader', label: 'Shader' },
    { id: 'Icons.Tools.Tabs.Background', label: 'Background' },
    { id: 'Icons.Tools.Tabs.Projection', label: 'Projection' },
    { id: 'Icons.Tools.Tabs.Render', label: 'Render' },
    { id: 'Icons.Tools.Tabs.CameraTracker', label: 'CameraTracker' },
    { id: 'Icons.Tools.Tabs.PlanarTracker', label: 'PlanarTracker' },
    { id: 'Icons.Tools.Tabs.Region', label: 'Region' },
    { id: 'Icons.Tools.Tabs.Calculation', label: 'Calculation' },
    { id: 'Icons.Tools.Tabs.Gradient', label: 'Gradient' },
    { id: 'Icons.Tools.Tabs.MIDI', label: 'MIDI' },
    { id: 'Icons.Tools.Tabs.MIDIChannel', label: 'MIDIChannel' },
    { id: 'Icons.Tools.Tabs.Offset', label: 'Offset' },
    { id: 'Icons.Tools.Tabs.PolyLine', label: 'PolyLine' },
    { id: 'Icons.Tools.Tabs.NumberProbe', label: 'NumberProbe' },
    { id: 'Icons.Tools.Tabs.Shake', label: 'Shake' },
    { id: 'Icons.Tools.Tabs.Time', label: 'Time' },
    { id: 'Icons.Tools.Tabs.XYPath', label: 'XYPath' },
    { id: 'Icons.Tools.Tabs.pRegion', label: 'pRegion' },
    { id: 'Icons.Tools.Tabs.pSets', label: 'pSets' },
    { id: 'Icons.Tools.Tabs.ColorChannels', label: 'ColorChannels' },
    { id: 'Icons.Tools.Tabs.AuxChannels', label: 'AuxChannels' },
    { id: 'Icons.Tools.Tabs.Crop', label: 'Crop' },
    { id: 'Icons.Tools.Tabs.AutoCrop', label: 'AutoCrop' },
    { id: 'Icons.Tools.Tabs.ColorGain', label: 'ColorGain' },
    { id: 'Icons.Tools.Tabs.Saturation', label: 'Saturation' },
    { id: 'Icons.Tools.Tabs.HotSpot', label: 'HotSpot' },
    { id: 'Icons.Tools.Tabs.ColorScale', label: 'ColorScale' },
    { id: 'Icons.Tools.Tabs.Grain', label: 'Grain' },
    { id: 'Icons.Tools.Tabs.HueCurves', label: 'HueCurves' },
    { id: 'Icons.Tools.Tabs.ChromaKeyer', label: 'ChromaKeyer' },
    { id: 'Icons.Tools.Tabs.CleanPlate', label: 'CleanPlate' },
    { id: 'Icons.Tools.Tabs.Spill', label: 'Spill' },
    { id: 'Icons.Tools.Tabs.Ranges', label: 'Ranges' },
    { id: 'Icons.Tools.Tabs.Mask', label: 'Mask' },
    { id: 'Icons.Tools.Tabs.Key', label: 'Key' },
    { id: 'Icons.Tools.Tabs.Matte', label: 'Matte' },
    { id: 'Icons.Tools.Tabs.Primatte', label: 'Primatte' },
    { id: 'Icons.Tools.Tabs.Replace', label: 'Replace' },
    { id: 'Icons.Tools.Tabs.Degrain', label: 'Degrain' },
    { id: 'Icons.Tools.Tabs.Numbers', label: 'Numbers' },
    { id: 'Icons.Tools.Tabs.Points', label: 'Points' },
    { id: 'Icons.Tools.Tabs.LUTs', label: 'LUTs' },
    { id: 'Icons.Tools.Tabs.Vertex', label: 'Vertex' },
    { id: 'Icons.Tools.Tabs.Red', label: 'Red' },
    { id: 'Icons.Tools.Tabs.Green', label: 'Green' },
    { id: 'Icons.Tools.Tabs.Blue', label: 'Blue' },
    { id: 'Icons.Tools.Tabs.Alpha', label: 'Alpha' },
    { id: 'Icons.Tools.Tabs.Solve', label: 'Solve' },
    { id: 'Icons.Tools.Tabs.Camera', label: 'Camera' },
    { id: 'Icons.Tools.Tabs.WhiteBalance', label: 'WhiteBalance' },
  ];
  const pageTabsContainer = pageTabsEl || null;
  let activePage = null;
  let draggingPage = null;
  let currentIconPicker = null;
  const detailTargetChangeCb = (typeof onDetailTargetChange === 'function') ? onDetailTargetChange : null;
  const entryMutatedCb = (typeof onEntryMutated === 'function') ? onEntryMutated : null;
  let detailTargetIndex = null;
  const getCurrentValueInfoCb = (typeof getCurrentValueInfo === 'function') ? getCurrentValueInfo : null;
  const shouldShowCurrentValuesCb = (typeof shouldShowCurrentValues === 'function') ? shouldShowCurrentValues : () => false;

  function notifyDetailTarget() {
    if (detailTargetChangeCb) detailTargetChangeCb(detailTargetIndex != null ? detailTargetIndex : null);
  }

  function notifyEntryMutated(index) {
    if (entryMutatedCb && typeof index === 'number' && index >= 0) {
      entryMutatedCb(index);
    }
  }

  function syncDetailTarget(preferredIdx = null) {
    try {
      const sel = getSelectedSet();
      if (preferredIdx != null && sel && sel.has(preferredIdx)) {
        detailTargetIndex = preferredIdx;
      } else if (sel && sel.size) {
        if (detailTargetIndex != null && sel.has(detailTargetIndex)) {
          // keep current
        } else {
          const arr = Array.from(sel);
          detailTargetIndex = arr[arr.length - 1];
        }
      } else {
        detailTargetIndex = null;
      }
      notifyDetailTarget();
    } catch (_) {
      detailTargetIndex = null;
      notifyDetailTarget();
    }
  }

  function ensurePageIconMap() {
    if (!state.parseResult) return new Map();
    let icons = state.parseResult.pageIcons;
    if (icons instanceof Map) return icons;
    const map = new Map();
    if (icons && typeof icons === 'object') {
      Object.entries(icons).forEach(([k, v]) => {
        if (v) map.set(normalizePageName(k), String(v));
      });
    }
    state.parseResult.pageIcons = map;
    return map;
  }

  function getPageIcon(pageName) {
    try {
      const map = ensurePageIconMap();
      return map.get(normalizePageName(pageName)) || '';
    } catch (_) { return ''; }
  }

  function setPageIconMapping(pageName, iconId) {
    try {
      const map = ensurePageIconMap();
      const normalized = normalizePageName(pageName);
      if (iconId) map.set(normalized, iconId);
      else map.delete(normalized);
      state.parseResult.pageIcons = map;
    } catch (_) {}
  }

  function getIconOptionLabel(id) {
    if (!id) return 'Default';
    const found = PAGE_ICON_OPTIONS.find(opt => opt.id === id);
    if (found) return found.label;
    return id.split('.').pop() || id;
  }

  function getIconShortLabel(id) {
    if (!id) return '–';
    const label = getIconOptionLabel(id);
    return label.slice(0, 2).toUpperCase();
  }

  function setFilter(value) {
    publishedFilter = (value || '').toLowerCase();
  }

  function updateRemoveSelectedState() {
    const sel = getSelectedSet();
    if (removeBtn) removeBtn.disabled = !(sel && sel.size > 0);
    if (deselectAllBtn) deselectAllBtn.disabled = !(sel && sel.size > 0);
  }

  function refreshAnchorFromSelection(order) {
    const sel = getSelectedSet();
    let maxSel = -1;
    sel.forEach(i2 => {
      const pp = order.indexOf(i2);
      if (pp > maxSel) maxSel = pp;
    });
    insertionAnchorPos = (maxSel >= 0) ? maxSel : null;
  }

  function setAnchor(pos) {
    if (pos == null || pos < 0) {
      insertionAnchorPos = null;
    } else {
      insertionAnchorPos = pos;
    }
  }

  function clearSelectionState() {
    try {
      const sel = getSelectedSet();
      if (sel && typeof sel.clear === 'function') sel.clear();
      state.parseResult.selected = sel instanceof Set ? sel : new Set();
      updateRemoveSelectedState();
      syncDetailTarget();
    } catch (_) {}
  }

  function clearSelection() {
    try {
      clearSelectionState();
      renderList(state.parseResult.entries, state.parseResult.order);
    } catch (_) {}
  }

  function setEntryDisplayNameByIndex(index, newName) {
    if (!state.parseResult || !Array.isArray(state.parseResult.entries)) return;
    const entry = state.parseResult.entries[index];
    if (!entry) return;
    const trimmed = (newName || '').trim();
    const fallback = entry.displayNameOriginal || `${entry.sourceOp || ''}${entry.source ? '.' + entry.source : ''}`;
    const nextDisplay = trimmed || fallback;
    if ((entry.displayName || '') === nextDisplay) {
      entry.displayNameDirty = false;
      entry.displayNameOverride = trimmed ? trimmed : null;
      if (entry.isLabel) {
        entry.name = null;
      } else {
        entry.name = trimmed || null;
      }
      renderList(state.parseResult.entries, state.parseResult.order);
      return;
    }
    if (typeof pushHistory === 'function') pushHistory('rename control');
    entry.displayName = nextDisplay;
    if (entry.isLabel) {
      entry.name = null;
    } else {
      entry.name = trimmed || null;
    }
    entry.displayNameOverride = trimmed || null;
    entry.displayNameDirty = !!trimmed && nextDisplay !== (entry.displayNameOriginal || fallback);
    renderList(state.parseResult.entries, state.parseResult.order);
    notifyEntryMutated(index);
  }

  function createIcon(name, size = 16) {
    const w = size;
    const h = size;
    const ns = 'http://www.w3.org/2000/svg';
    if (name === 'drag') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><g stroke="currentColor" stroke-width="1.2"><circle cx="5" cy="4" r="1"/><circle cx="11" cy="4" r="1"/><circle cx="5" cy="8" r="1"/><circle cx="11" cy="8" r="1"/><circle cx="5" cy="12" r="1"/><circle cx="11" cy="12" r="1"/></g></svg>`;
    }
    if (name === 'chevron-down') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M4.2 6.5 8 10.3l3.8-3.8.9.9L8 12 3.3 7.4z"/></svg>`;
    }
    if (name === 'chevron-right') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M6.5 4.2 10.3 8 6.5 11.8l.9.9L12 8 7.4 3.3z"/></svg>`;
    }
    if (name === 'chevron-left') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="m9.5 11.8-3.8-3.8 3.8-3.8-.9-.9L4 8l4.6 4.7z"/></svg>`;
    }
    if (name === 'arrow-up') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M8 3 3.5 7.5l.9.9L7.3 5.5V13h1.4V5.5l2.9 2.9.9-.9z"/></svg>`;
    }
    if (name === 'arrow-down') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M8 13l4.5-4.5-.9-.9-2.9 2.9V3H7.3v7.5L4.4 7.6l-.9.9z"/></svg>`;
    }
    if (name === 'trash') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M6 2h4l1 1h3v1H2V3h3l1-1Zm-2 4h8l-.6 8.1A2 2 0 0 1 9.4 16H6.6a2 2 0 0 1-1.99-1.9L4 6Zm3 1v6h1V7H7Zm2 0v6h1V7H9Z"/></svg>`;
    }
    if (name === 'pin') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M9 1 6 4l2 2-3 4-2 1 1-2 4-3-2-2 3-3 2 2 1-1 1 1-1 1 2 2-1 1-2-2-3 3 2 2-1 1-2-2-3 2 2-3-2-2 1-1 2 2 3-3z"/></svg>`;
    }
    if (name === 'eye-open') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M8 3.5c3.1 0 5.7 2 7 4.5-1.3 2.5-3.9 4.5-7 4.5s-5.7-2-7-4.5c1.3-2.5 3.9-4.5 7-4.5Zm0 1.4c-2.2 0-4.2 1.4-5.2 3.1 1 1.7 3 3.1 5.2 3.1s4.2-1.4 5.2-3.1c-1-1.7-3-3.1-5.2-3.1Zm0 1.1a2.5 2.5 0 1 1 0 5 2.5 2.5 0 0 1 0-5Zm0 1.2a1.3 1.3 0 1 0 0 2.6 1.3 1.3 0 0 0 0-2.6Z"/></svg>`;
    }
    if (name === 'eye-closed') {
      return `<svg xmlns="${ns}" width="${w}" height="${h}" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M8 3.5c3.1 0 5.7 2 7 4.5-.4.8-.9 1.5-1.5 2.2L12.8 8.5c.3-.4.5-.7.7-1-1-1.7-3-3.1-5.2-3.1-.5 0-1 .1-1.5.2l-1.2-1.2c.7-.3 1.7-.4 2.4-.4Zm-6.5 4.5a8.6 8.6 0 0 1 2.5-3.1L2.2 3.1l.7-.7 11.2 11.2-.7.7-2.1-2.1A11 11 0 0 1 8 12.5c-3.1 0-5.7-2-7-4.5Zm2.3 0c1 1.7 3 3.1 5.2 3.1.4 0 .9 0 1.3-.1L7.9 9.1a2.5 2.5 0 0 1-3.2-3.2L5 7l.9.9a1.3 1.3 0 0 0 1.2 1.2L4 8Z"/></svg>`;
    }
    return '';
  }

  function normalizePageName(name) {
    const n = (name && String(name).trim()) ? String(name).trim() : 'Controls';
    return n;
  }

  function ensurePageOrderIncludes(pageName) {
    try {
      if (!state.parseResult) return;
      const normalized = normalizePageName(pageName);
      let order = Array.isArray(state.parseResult.pageOrder) ? [...state.parseResult.pageOrder] : [];
      if (!order.includes(normalized)) {
        order.push(normalized);
        state.parseResult.pageOrder = order;
      }
    } catch (_) {}
  }

  function syncPageOrderFromState() {
    if (!state.parseResult) return ['Controls'];
    const entries = Array.isArray(state.parseResult.entries) ? state.parseResult.entries : [];
    let order = Array.isArray(state.parseResult.pageOrder) ? state.parseResult.pageOrder.map(normalizePageName) : [];
    const seen = new Set();
    const normalizedOrder = [];
    const add = (name) => {
      const norm = normalizePageName(name);
      if (!seen.has(norm)) {
        normalizedOrder.push(norm);
        seen.add(norm);
      }
    };
    order.forEach(add);
    entries.forEach((entry) => add(getEntryPage(entry)));
    if (!seen.has('Controls')) normalizedOrder.unshift('Controls');
    state.parseResult.pageOrder = normalizedOrder;
    return normalizedOrder;
  }

  function setActivePage(pageName, opts = {}) {
    const normalized = pageName ? normalizePageName(pageName) : null;
    const changed = normalized !== activePage;
    activePage = normalized;
    if (state.parseResult) state.parseResult.activePage = normalized;
    if (changed) closeIconPicker();
    if (changed && state.parseResult) {
      clearSelectionState();
    }
    if (changed && !opts.silent && state.parseResult) {
      renderList(state.parseResult.entries, state.parseResult.order);
    }
  }

  function refreshPageTabsInternal(opts = {}) {
    if (!pageTabsContainer) return;
    const pages = syncPageOrderFromState();
    if (!pages || pages.length <= 1) {
      pageTabsContainer.hidden = true;
      pageTabsContainer.innerHTML = '';
      setActivePage(null, { silent: true });
      return;
    }
    const storedActive = state.parseResult?.activePage || activePage;
    if (!storedActive || !pages.includes(storedActive)) {
      setActivePage(pages[0], { silent: opts.deferRenderView });
    } else {
      setActivePage(storedActive, { silent: opts.deferRenderView });
    }
    pageTabsContainer.hidden = false;
    pageTabsContainer.innerHTML = '';
    closeIconPicker();
    pages.forEach((page) => {
      const tab = document.createElement('div');
      tab.className = 'page-tab';
      if (page === activePage) tab.classList.add('active');
      tab.dataset.page = page;
      tab.draggable = true;
      tab.addEventListener('click', (ev) => {
        if (ev.target && ev.target.closest('.page-tab-icon-btn')) return;
        setActivePage(page);
      });
      tab.addEventListener('dragstart', (e) => {
        draggingPage = page;
        tab.classList.add('dragging');
        e.dataTransfer?.setData('text/plain', page);
      });
      tab.addEventListener('dragend', () => {
        draggingPage = null;
        tab.classList.remove('dragging');
        clearTabDragIndicators();
      });
      tab.addEventListener('dragover', (e) => {
        if (!draggingPage || draggingPage === page) return;
        e.preventDefault();
        tab.classList.add('drag-over');
      });
      tab.addEventListener('dragleave', () => tab.classList.remove('drag-over'));
      tab.addEventListener('drop', (e) => {
        if (!draggingPage || draggingPage === page) return;
        e.preventDefault();
        tab.classList.remove('drag-over');
        reorderPageOrder(draggingPage, page);
      });
      const labelBtn = document.createElement('button');
      labelBtn.type = 'button';
      labelBtn.className = 'page-tab-label';
      labelBtn.textContent = page;
      labelBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        setActivePage(page);
      });
      const iconBtn = document.createElement('button');
      iconBtn.type = 'button';
      iconBtn.className = 'page-tab-icon-btn';
      const iconId = getPageIcon(page);
      iconBtn.textContent = getIconShortLabel(iconId);
      iconBtn.title = `Change icon (${getIconOptionLabel(iconId)})`;
      iconBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        openIconPicker(page, iconBtn);
      });
      tab.appendChild(labelBtn);
      tab.appendChild(iconBtn);
      pageTabsContainer.appendChild(tab);
    });
  }

  function handleTabContainerDragOver(e) {
    if (!draggingPage) return;
    e.preventDefault();
  }

  function handleTabContainerDrop(e) {
    if (!draggingPage) return;
    e.preventDefault();
    const target = e.target;
    if (target && target.classList && target.classList.contains('page-tab')) return;
    reorderPageOrder(draggingPage, null, true);
  }

  function reorderPageOrder(fromPage, targetPage, appendToEnd = false) {
    try {
      if (!state.parseResult) return;
      const order = Array.isArray(state.parseResult.pageOrder) ? [...state.parseResult.pageOrder] : syncPageOrderFromState();
      const fromIndex = order.indexOf(fromPage);
      if (fromIndex < 0) return;
      const [item] = order.splice(fromIndex, 1);
      let insertIndex = appendToEnd ? order.length : order.indexOf(targetPage);
      if (insertIndex < 0) insertIndex = order.length;
      order.splice(insertIndex, 0, item);
      state.parseResult.pageOrder = order;
      refreshPageTabsInternal({ deferRenderView: true });
      renderList(state.parseResult.entries, state.parseResult.order);
    } catch (_) {}
  }

  function clearTabDragIndicators() {
    if (!pageTabsContainer) return;
    pageTabsContainer.querySelectorAll('.page-tab.drag-over').forEach(el => el.classList.remove('drag-over'));
  }

  function closeIconPicker() {
    if (!currentIconPicker) return;
    try {
      currentIconPicker.remove();
    } catch (_) {}
    currentIconPicker = null;
  }

  function openIconPicker(pageName, anchor) {
    try {
      closeIconPicker();
      const picker = document.createElement('select');
      picker.className = 'page-icon-select';
      PAGE_ICON_OPTIONS.forEach((opt) => {
        const optionEl = document.createElement('option');
        optionEl.value = opt.id;
        optionEl.textContent = opt.label;
        picker.appendChild(optionEl);
      });
      picker.value = getPageIcon(pageName) || '';
      const rect = anchor.getBoundingClientRect();
      picker.style.position = 'absolute';
      picker.style.left = `${rect.left + window.scrollX}px`;
      picker.style.top = `${rect.bottom + window.scrollY + 4}px`;
      document.body.appendChild(picker);
      currentIconPicker = picker;
      const cleanup = () => {
        if (!picker) return;
        try { picker.remove(); } catch (_) {}
        if (currentIconPicker === picker) currentIconPicker = null;
      };
      picker.addEventListener('change', () => {
        const val = picker.value || '';
        setPageIconMapping(pageName, val || null);
        refreshPageTabsInternal({ deferRenderView: true });
        renderList(state.parseResult.entries, state.parseResult.order);
        cleanup();
      });
      picker.addEventListener('blur', () => {
        setTimeout(() => cleanup(), 150);
      });
      picker.focus();
    } catch (_) {}
  }

  if (pageTabsContainer) {
    pageTabsContainer.addEventListener('dragover', handleTabContainerDragOver);
    pageTabsContainer.addEventListener('drop', handleTabContainerDrop);
  }

  deselectAllBtn?.addEventListener('click', () => {
    if (!state.parseResult) return;
    if (typeof pushHistory === 'function') pushHistory('clear selection');
    clearSelection();
  });

  const notifyRenderList = typeof onRenderList === 'function' ? onRenderList : null;

  function renderList(entries, order) {
    controlsList.innerHTML = '';
    insertionAnchorPos = null;
    const filter = publishedFilter.trim().toLowerCase();
    const collapsedCG = state.parseResult?.collapsedCG || new Set();
    const collapsed = state.parseResult?.collapsed || new Set();
    const showCurrentValues = !!(shouldShowCurrentValuesCb && shouldShowCurrentValuesCb());
    const indentMap = computeIndentMap(entries, order);
    const labelGroupCounts = new Map();
    try {
      for (let pos = 0; pos < order.length; pos++) {
        const idx = order[pos];
        const e = entries[idx];
        if (!e || !e.isLabel || !Number.isFinite(e.labelCount) || e.labelCount <= 0) continue;
        for (let p = pos; p <= pos + e.labelCount && p < order.length; p++) {
          const targetIdx = order[p];
          labelGroupCounts.set(targetIdx, (labelGroupCounts.get(targetIdx) || 0) + 1);
        }
      }
    } catch (_) {}

    let pos = 0;
    const total = order.length;
    const useChunking = total > 250 && typeof requestAnimationFrame === 'function';
    const finalize = () => {
      refreshNodesChecks();
      refreshPageTabsInternal();
      syncDetailTarget();
      try {
        if (Array.isArray(order)) {
          order.forEach((idx, pos) => {
            const entry = entries[idx];
            if (entry) entry.sortIndex = pos;
          });
        }
      } catch (_) {}
    };
    const runChunk = () => {
      const start = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
      while (pos < total) {
        const idx = order[pos];
        const e = entries[idx];
      if (e && e.locked) continue;
      const filterPage = activePage;
      if (filterPage && getEntryPage(e) !== filterPage) continue;
      if (filter) {
        const t = (e.displayName || '').toLowerCase();
        const s = (e.source || '').toLowerCase();
        const op = (e.sourceOp || '').toLowerCase();
        if (!(t.includes(filter) || s.includes(filter) || op.includes(filter))) continue;
      }

      const li = document.createElement('li');
      li.draggable = true;
      li.dataset.index = String(idx);
      li.setAttribute('aria-grabbed', 'false');
      if (getSelectedSet().has(idx)) {
        li.classList.add('selected');
        insertionAnchorPos = Math.max(insertionAnchorPos ?? -1, pos);
      }
      li.addEventListener('click', (ev) => {
        if (isInteractiveTarget(ev.target)) return;
        const sel = getSelectedSet();
        const alreadySelected = li.classList.contains('selected');
        const hasModifier = ev.shiftKey || ev.metaKey || ev.ctrlKey;
        if (alreadySelected && !hasModifier) {
          if (detailTargetIndex === idx) {
            li.classList.remove('selected');
            sel.delete(idx);
            state.parseResult.selected = sel;
            refreshAnchorFromSelection(order);
            updateRemoveSelectedState();
            syncDetailTarget(null);
          } else {
            syncDetailTarget(idx);
          }
          return;
        }
        const willSelect = !alreadySelected;
        li.classList.toggle('selected', willSelect);
        if (willSelect) sel.add(idx); else sel.delete(idx);
        state.parseResult.selected = sel;
        const p = order.indexOf(idx);
        if (willSelect) insertionAnchorPos = (insertionAnchorPos == null) ? p : Math.max(insertionAnchorPos, p);
        else refreshAnchorFromSelection(order);
        updateRemoveSelectedState();
        syncDetailTarget(willSelect ? idx : null);
      });

      if (e.isLabel) li.classList.add('label');
      const labelDepth = labelGroupCounts.get(idx) || 0;
      if (labelDepth > 0) {
        li.classList.add('label-group');
        li.style.setProperty('--label-group-depth', String(labelDepth));
      } else {
        li.classList.remove('label-group');
        li.style.removeProperty('--label-group-depth');
      }
      try {
        if (Number.isFinite(e.controlGroup) && e.sourceOp) {
          const cg = getContiguousColorGroupBlock(idx, entries, order);
          if (cg && cg.count >= 2) {
            li.classList.add('color-group');
          }
        }
      } catch (_) {}

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.title = 'Select row';
      cb.checked = getSelectedSet().has(idx);
      cb.addEventListener('change', (ev) => {
        ev.stopPropagation();
        const sel = getSelectedSet();
        if (cb.checked) sel.add(idx); else sel.delete(idx);
        state.parseResult.selected = sel;
        li.classList.toggle('selected', cb.checked);
        const p = order.indexOf(idx);
        if (cb.checked) insertionAnchorPos = (insertionAnchorPos == null) ? p : Math.max(insertionAnchorPos, p);
        else refreshAnchorFromSelection(order);
        updateRemoveSelectedState();
        syncDetailTarget(cb.checked ? idx : null);
      });

      const handle = document.createElement('div');
      handle.className = 'handle';
      handle.title = 'Drag to reorder';
      handle.innerHTML = createIcon('drag');

      const twistyCell = document.createElement('div');
      twistyCell.className = 'twisty-cell';
      const textCol = document.createElement('div');
      textCol.className = 'text-col';
      const textBlock = document.createElement('div');
      textBlock.className = 'text-block';
      const titleRow = document.createElement('div');
      titleRow.className = 'title-row';
      const title = document.createElement('div');
      title.className = 'title';
      const titleText = e.displayName || e.name || e.source || 'Control';
      if (e.isLabel) {
        title.innerHTML = buildLabelMarkup(titleText, normalizeLabelStyle(e.labelStyle));
      } else {
        title.textContent = titleText;
      }
      title.contentEditable = 'true';
      title.spellcheck = false;
      title.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter') { ev.preventDefault(); title.blur(); }
        if (ev.key === 'Escape') { ev.preventDefault(); title.textContent = e.displayName; title.blur(); }
      });
      title.addEventListener('blur', () => {
        const newName = (title.textContent || '').trim();
        setEntryDisplayNameByIndex(idx, newName);
      });

      if (e.isLabel) {
        const countInput = document.createElement('input');
        countInput.type = 'number';
        countInput.min = '0';
        countInput.step = '1';
        countInput.value = String(e.labelCount || 0);
        countInput.title = 'Items in this label group';
        countInput.className = 'label-count';
        countInput.addEventListener('change', () => {
          const val = parseInt(countInput.value, 10);
          pushHistory('label count');
          e.labelCount = Number.isFinite(val) && val >= 0 ? val : 0;
          renderList(entries, order);
          notifyEntryMutated(idx);
        });
        titleRow.appendChild(countInput);
        const eyeBtn = document.createElement('button');
        eyeBtn.type = 'button';
        eyeBtn.className = 'label-visibility-btn';
        const refreshEye = () => {
          eyeBtn.innerHTML = createIcon(e.labelHidden ? 'eye-closed' : 'eye-open');
          eyeBtn.title = e.labelHidden ? 'Show label group' : 'Hide label group';
        };
        refreshEye();
        eyeBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          toggleLabelVisibility(e);
        });
        titleRow.appendChild(eyeBtn);
      }

      const cgBlk = getContiguousColorGroupBlock(idx, entries, order);
      const indent = document.createElement('div');
      indent.className = 'indent';
      indent.style.width = '16px';
      const twisty = document.createElement('span');
      twisty.className = e.isLabel ? 'twisty twisty-label' : 'twisty hidden';
      twisty.title = 'Collapse/expand label group';
      twisty.innerHTML = (e.isLabel && collapsed.has(idx)) ? createIcon('chevron-right') : createIcon('chevron-down');
      if (e.isLabel) {
        const toggle = (ev) => {
          ev.stopPropagation();
          if (collapsed.has(idx)) collapsed.delete(idx); else collapsed.add(idx);
          state.parseResult.collapsed = collapsed;
          renderList(entries, order);
        };
        twisty.addEventListener('click', toggle);
        title.addEventListener('click', (ev) => {
          if (ev.target && ev.target.isContentEditable) return;
          toggle(ev);
        });
        twistyCell.appendChild(twisty);
      }
      if (cgBlk && cgBlk.firstIndex === idx && cgBlk.count >= 2 && e.sourceOp && Number.isFinite(e.controlGroup)) {
        const key = getCgKey(e);
        const cgTwisty = document.createElement('span');
        cgTwisty.className = 'twisty twisty-group';
        cgTwisty.innerHTML = collapsedCG.has(key) ? createIcon('chevron-right') : createIcon('chevron-down');
        cgTwisty.title = 'Toggle grouped controls';
        cgTwisty.addEventListener('click', (ev) => {
          ev.stopPropagation();
          if (collapsedCG.has(key)) collapsedCG.delete(key); else collapsedCG.add(key);
          state.parseResult.collapsedCG = collapsedCG;
          renderList(entries, order);
        });
        twistyCell.appendChild(cgTwisty);
        const badge = document.createElement('span');
        badge.style.marginLeft = '6px';
        badge.style.color = 'var(--muted)';
        badge.textContent = '(' + cgBlk.count + ')';
        title.appendChild(badge);
      }

      const meta = document.createElement('div');
      meta.className = 'meta';
      const detailBits = [];
      if (e.sourceOp || e.source) detailBits.push(`${e.sourceOp || ''}${e.source ? (e.sourceOp ? '.' : '') + e.source : ''}`);
      meta.textContent = detailBits.join('  ');
      if (e.onChange && String(e.onChange).trim().length) {
        const badge = document.createElement('span');
        badge.className = 'meta-badge';
        badge.textContent = 'OnChange';
        meta.appendChild(badge);
      }
      titleRow.prepend(title);
      textBlock.appendChild(titleRow);
      if (detailBits.length) textBlock.appendChild(meta);
      textCol.appendChild(indent);
      textCol.appendChild(textBlock);

      const buttons = document.createElement('div');
      buttons.className = 'row-buttons';
      const goBtn = document.createElement('button');
      goBtn.type = 'button';
      goBtn.innerHTML = createIcon('chevron-right');
      goBtn.title = 'Jump to node';
      goBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        if (e && e.sourceOp && e.source) { try { clearHighlights(); highlightHandler(e.sourceOp, e.source); } catch (_) {} }
      });
      const pageCell = document.createElement('div');
      pageCell.className = 'page-select';
      const pageSelect = document.createElement('select');
      pageSelect.title = 'Change page';
      populatePageOptions(pageSelect, getEntryPage(e));
      try {
        const currentPage = getEntryPage(e) || 'Controls';
        pageSelect.value = currentPage;
      } catch (_) {
        pageSelect.value = 'Controls';
      }
      pageSelect.addEventListener('change', (ev) => {
        ev.stopPropagation();
        const selectedSet = new Set(getSelectedSet() || []);
        if (!selectedSet.has(idx)) selectedSet.add(idx);
        const applyPageToTargets = (pageName) => {
          const targets = getEntriesByIndices(selectedSet);
          applyPageToEntries(targets, pageName);
        };
        const val = pageSelect.value;
        if (val === '__new__') {
          pageSelect.value = getEntryPage(e) || 'Controls';
          promptForNewPage(pageCell, pageSelect, { entry: e, applyMultiple: applyPageToTargets });
        } else {
          applyPageToTargets(val);
        }
      });
      const blendCell = document.createElement('div');
      blendCell.className = 'blend-toggle-cell';
      if (isBlendEntry(e)) {
        const blendBtn = document.createElement('button');
        blendBtn.type = 'button';
        const active = isBlendToggleEnabled(e);
        blendBtn.textContent = active ? 'Unset On/Off' : 'Set On/Off';
        blendBtn.title = 'Convert Blend control to on/off toggle';
        blendBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          toggleBlendCheckbox(e);
        });
        blendCell.appendChild(blendBtn);
      }

      let currentCell = null;
      if (showCurrentValues) {
        currentCell = document.createElement('div');
        currentCell.className = 'current-value';
        let value = '';
        let note = '';
        if (getCurrentValueInfoCb) {
          const info = getCurrentValueInfoCb(e) || {};
          value = info.value != null ? String(info.value).trim() : '';
          note = info.note || '';
        }
        if (!value) {
          currentCell.textContent = '--';
          currentCell.classList.add('empty');
          if (note) currentCell.title = note;
        } else {
          currentCell.textContent = value;
          currentCell.title = note || value;
        }
      }

      pageCell.appendChild(pageSelect);
      if (blendCell.childElementCount) buttons.appendChild(blendCell);
      if (currentCell) buttons.appendChild(currentCell);
      buttons.appendChild(pageCell);
      buttons.appendChild(goBtn);

      li.appendChild(cb);
      li.appendChild(handle);
      li.appendChild(twistyCell);
      li.appendChild(textCol);
      li.appendChild(buttons);
      addDragHandlers(li, entries, order);
      controlsList.appendChild(li);

        let skip = 0;
        if (e.isLabel && e.labelCount > 0 && collapsed.has(idx)) skip += e.labelCount;
        if (cgBlk && cgBlk.firstIndex === idx && cgBlk.count >= 2 && e.sourceOp && Number.isFinite(e.controlGroup) && collapsedCG.has(getCgKey(e))) skip += (cgBlk.count - 1);
        if (skip > 0) pos += skip;
        pos += 1;
        if (useChunking) {
          const now = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
          if (now - start > 12) break;
        }
      }
      if (useChunking && pos < total) {
        requestAnimationFrame(runChunk);
      } else {
        finalize();
      }
    };
    runChunk();
    if (notifyRenderList) {
      try { notifyRenderList(); } catch (_) {}
    }
  }

  function computeIndentMap(entries, order) {
    const indentMap = new Map();
    let active = [];
    for (let pos = 0; pos < order.length; pos++) {
      const idx = order[pos];
      const e = entries[idx];
      indentMap.set(idx, active.length);
      if (e && e.isLabel && e.labelCount > 0) {
        active.push(e.labelCount + 1);
      }
      active = active.map(v => v - 1).filter(v => v > 0);
    }
    return indentMap;
  }

  function moveViaButtons(clickedIdx, delta) {
    if (!state.parseResult) return;
    const order = [...state.parseResult.order];
    const group = getDragSelection(clickedIdx);
    let newOrder = null;
    if (group && group.length > 1) {
      newOrder = moveSelectionBlock(order, new Set(group), delta);
    } else {
      const pos = order.indexOf(clickedIdx);
      const newPos = pos + delta;
      if (pos < 0 || newPos < 0 || newPos >= order.length) return;
      newOrder = [...order];
      const [item] = newOrder.splice(pos, 1);
      newOrder.splice(newPos, 0, item);
    }
    logDiag('button move');
    state.parseResult.order = newOrder;
    renderList(state.parseResult.entries, newOrder);
    notifyEntryMutated(clickedIdx);
  }

  function moveSelectionBlock(order, selSet, delta) {
    const selectedInOrder = order.filter(v => selSet.has(v));
    if (!selectedInOrder.length) return order;
    const remaining = order.filter(v => !selSet.has(v));
    const posMap = new Map();
    order.forEach((v, i) => posMap.set(v, i));
    const firstSelIndex = posMap.get(selectedInOrder[0]);
    const blockIndexInRemaining = remaining.reduce((acc, v) => acc + (posMap.get(v) < firstSelIndex ? 1 : 0), 0);
    let newBlockIndex = blockIndexInRemaining + (delta < 0 ? -1 : delta > 0 ? +1 : 0);
    newBlockIndex = Math.max(0, Math.min(remaining.length, newBlockIndex));
    if (newBlockIndex === blockIndexInRemaining) return order;
    const out = remaining.slice();
    out.splice(newBlockIndex, 0, ...selectedInOrder);
    return out;
  }

  function addDragHandlers(li, entries, order) {
    li.addEventListener('dragstart', (e) => {
      li.classList.add('dragging');
      li.setAttribute('aria-grabbed', 'true');
      const idx = parseInt(li.dataset.index || '-1', 10);
      dragFromIndex = idx;
      currentDropIndex = null;
      e.dataTransfer?.setData('text/plain', String(idx));
      e.dataTransfer?.setDragImage(li, 10, 10);
    });
    li.addEventListener('dragend', () => {
      li.classList.remove('dragging');
      li.setAttribute('aria-grabbed', 'false');
      cleanupDropIndicator();
    });
    li.addEventListener('dragover', (e) => {
      if (!state.parseResult) return;
      e.preventDefault();
      e.dataTransfer && (e.dataTransfer.dropEffect = 'move');
      updateDropIndicatorPosition(e, state.parseResult.order);
    });
    li.addEventListener('drop', (e) => {
      if (!state.parseResult || dragFromIndex === null || currentDropIndex === null) return;
      e.preventDefault();
      pushHistory('drag/drop');
      const orderCopy = [...state.parseResult.order];
      const selected = getDragSelection(dragFromIndex);
      const positions = selected.map(v => orderCopy.indexOf(v)).sort((a, b) => a - b);
      const selSet = new Set(selected);
      const remaining = orderCopy.filter(v => !selSet.has(v));
      let toPos = currentDropIndex;
      const countLess = positions.filter(p => p < toPos).length;
      toPos = Math.max(0, Math.min(remaining.length, toPos - countLess));
      const selectedInOrder = orderCopy.filter(v => selSet.has(v));
      remaining.splice(toPos, 0, ...selectedInOrder);
      logDiag('row drop');
      state.parseResult.order = remaining;
      renderList(entries, remaining);
      if (selectedInOrder.length) {
        notifyEntryMutated(selectedInOrder[selectedInOrder.length - 1]);
      }
      cleanupDropIndicator();
    });
  }

  function cleanupDropIndicator() {
    dragFromIndex = null;
    currentDropIndex = null;
    if (dropIndicator.parentElement) dropIndicator.parentElement.removeChild(dropIndicator);
  }

  function updateDropIndicatorPosition(e, order) {
    const children = Array.from(controlsList.children).filter(el => el !== dropIndicator && el.offsetParent !== null);
    let nextEl = null;
    for (const child of children) {
      const rect = child.getBoundingClientRect();
      const offset = e.clientY - rect.top - rect.height / 2;
      if (offset < 0) { nextEl = child; break; }
    }
    if (nextEl) {
      const idx = parseInt(nextEl.dataset.index || '-1', 10);
      const pos = order.indexOf(idx);
      currentDropIndex = Math.max(0, pos);
      controlsList.insertBefore(dropIndicator, nextEl);
    } else {
      currentDropIndex = order.length;
      controlsList.appendChild(dropIndicator);
    }
  }

  function getCurrentDropIndex() {
    return currentDropIndex;
  }

  function getDragSelection(dragIdx) {
    const sel = getSelectedSet();
    const order = state.parseResult?.order || [];
    const entries = state.parseResult?.entries || [];
    const expandLabelGroup = (idx) => {
      const pos = order.indexOf(idx);
      if (pos < 0) return [idx];
      const ent = entries[idx];
      const cnt = (ent && ent.isLabel && Number.isFinite(ent.labelCount)) ? ent.labelCount : 0;
      return [idx, ...order.slice(pos + 1, pos + 1 + cnt)];
    };
    const expandColorGroup = (idx) => {
      const blk = getContiguousColorGroupBlock(idx, entries, order);
      return (blk && blk.count >= 2) ? blk.indices : [idx];
    };
    const expandAll = (idx) => {
      const a = new Set(expandLabelGroup(idx));
      for (const v of expandColorGroup(idx)) a.add(v);
      return Array.from(a);
    };
    if (sel.size > 1 && sel.has(dragIdx)) {
      const items = [];
      for (const s of sel) {
        for (const v of expandAll(s)) if (!items.includes(v)) items.push(v);
      }
      return order.filter(v => items.includes(v));
    }
    return expandAll(dragIdx);
  }

  function getInsertionPosUnderSelection() {
    if (!state.parseResult) return 0;
    const order = state.parseResult.order || [];
    const sel = getSelectedSet();
    if (detailTargetIndex != null) {
      const detailPos = order.indexOf(detailTargetIndex);
      if (detailPos >= 0) {
        const pos = Math.max(0, Math.min(order.length, detailPos + 1));
        try { logDiag(`Anchor pos via detail target = ${detailPos}, insert at ${pos}`); } catch (_) {}
        return pos;
      }
    }
    if (insertionAnchorPos != null) {
      const pos = Math.max(0, Math.min(order.length, insertionAnchorPos + 1));
      try { logDiag(`Anchor pos via insertionAnchorPos = ${insertionAnchorPos}, insert at ${pos}`); } catch (_) {}
      return pos;
    }
    try {
      let maxPosSel = -1;
      if (controlsList) {
        const items = controlsList.querySelectorAll('li.selected[data-index]');
        Array.from(items || []).forEach(li => {
          const idx = parseInt(li.dataset.index || '-1', 10);
          const p = order.indexOf(idx);
          if (p > maxPosSel) maxPosSel = p;
        });
      }
      if (maxPosSel >= 0) {
        try { logDiag(`Anchor pos via li.selected = ${maxPosSel}`); } catch (_) {}
        return maxPosSel + 1;
      }
    } catch (_) {}
    if (sel && sel.size > 0) {
      const maxPos = Math.max(...Array.from(sel).map(v => order.indexOf(v)).filter(v => v >= 0));
      if (maxPos >= 0) {
        try { logDiag(`Anchor pos via model selected = ${maxPos}, insert at ${maxPos + 1}`); } catch (_) {}
        return maxPos + 1;
      }
    }
    try {
      const checked = controlsList ? controlsList.querySelectorAll('input[type="checkbox"]:checked[data-index]') : [];
      if (checked && checked.length) {
        let maxPos = -1;
        Array.from(checked).forEach(cb => {
          const idx = parseInt(cb.dataset.index || '-1', 10);
          const p = order.indexOf(idx);
          if (p > maxPos) maxPos = p;
        });
        if (maxPos >= 0) {
          try { logDiag(`Anchor pos via checked boxes = ${maxPos}, insert at ${maxPos + 1}`); } catch (_) {}
          return maxPos + 1;
        }
      }
    } catch (_) {}
    try { logDiag('Anchor pos default = end'); } catch (_) {}
    return order.length;
  }

  function addOrMovePublishedItemsAt(items, insertPos) {
    if (!state.parseResult || !Array.isArray(items) || !items.length) return;
    const orderCopy = [...state.parseResult.order];
    const indices = [];
    for (const it of items) {
      if (!it || !it.sourceOp || !it.source) continue;
      let idx = (state.parseResult.entries || []).findIndex(entry => entry && entry.sourceOp === it.sourceOp && entry.source === it.source);
      if (idx < 0) {
        const key = makeUniqueKey(`${it.sourceOp}_${it.source}`);
        const controlGroup = Number.isFinite(it.controlGroup) ? it.controlGroup : (it.base ? getOrAssignControlGroup(it.sourceOp, it.base) : null);
        const raw = buildInstanceInputRaw(key, it.sourceOp, it.source, it.displayName || it.source, 'Controls', controlGroup);
        const displayName = it.displayName || `${it.sourceOp}.${it.source}`;
        const entry = {
          key,
          name: it.displayName || null,
          page: 'Controls',
          sourceOp: it.sourceOp,
          source: it.source,
          displayName,
          raw,
          controlGroup: (Number.isFinite(controlGroup) ? controlGroup : null),
          controlMeta: {},
          controlMetaOriginal: {},
          displayNameOriginal: displayName,
        };
        state.parseResult.entries.push(entry);
        idx = state.parseResult.entries.length - 1;
        orderCopy.push(idx);
      }
      if (idx >= 0 && typeof applyNodeControlMeta === 'function') {
        const entry = state.parseResult.entries[idx];
        applyNodeControlMeta(entry, it);
      }
      indices.push(idx);
    }
    const selSet = new Set(indices);
    const positions = indices.map(v => orderCopy.indexOf(v)).sort((a, b) => a - b);
    const remaining = orderCopy.filter(v => !selSet.has(v));
    let toPos = (typeof insertPos === 'number' && insertPos >= 0) ? insertPos : remaining.length;
    const countLess = positions.filter(p => p < toPos).length;
    toPos = Math.max(0, Math.min(remaining.length, toPos - countLess));
    const selectedInOrder = orderCopy.filter(v => selSet.has(v));
    remaining.splice(toPos, 0, ...selectedInOrder);
    state.parseResult.order = remaining;
    if (state.parseResult.entries && ctrlCountEl) {
      ctrlCountEl.textContent = String(state.parseResult.entries.length);
    }
    renderList(state.parseResult.entries, state.parseResult.order);
    updateRemoveSelectedState();
    if (indices.length) {
      notifyEntryMutated(indices[indices.length - 1]);
    }
  }

  function entryLooksLikeButton(entry) {
    try {
      if (entry?.isButton) return true;
      const meta = entry?.controlMeta || entry?.controlMetaOriginal;
      if (!meta) return false;
      const val = String(meta.inputControl || '');
      return /buttoncontrol/i.test(val);
    } catch (_) { return false; }
  }

  function normalizeLauncherUrl(raw) {
    const value = String(raw || '').trim();
    if (!value) return '';
    if (/^[a-z][a-z0-9+.-]*:/i.test(value)) return value;
    return `https://${value}`;
  }

  function buildLauncherLuaScript(url) {
    try {
      const safe = String(url || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      const parts = [
        `local url = "${safe}"`,
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
      return parts.join('\n') + '\n';
    } catch (_) {
      return '';
    }
  }

  function appendLauncherUi(container, entry, options = {}) {
    try {
      const onScriptChange = options && typeof options.onScriptChange === 'function' ? options.onScriptChange : null;
      const emitScript = (script, meta) => {
        if (onScriptChange) {
          try { onScriptChange(script, meta || {}); } catch (_) {}
        }
      };
      const emitScriptFromUrl = (url, meta = {}) => {
        const normalizedUrl = normalizeLauncherUrl(url);
        const script = normalizedUrl ? buildLauncherLuaScript(normalizedUrl) : '';
        emitScript(script, { ...(meta || {}), url: normalizedUrl });
      };
      if (!state.parseResult) state.parseResult = {};
      if (!state.parseResult.buttonExactInsert) state.parseResult.buttonExactInsert = new Set();
      if (!state.parseResult.insertUrlMap) state.parseResult.insertUrlMap = new Map();
      const grp = findEnclosingGroupForIndex(state.originalText, state.parseResult.inputs.openIndex);
      if (!grp || !entry || !entry.sourceOp || !entry.source) return;
      const looksLikeButton = entryLooksLikeButton(entry);
      if (!looksLikeButton && !shouldOfferInsertForControl(state.originalText, grp.groupOpenIndex, grp.groupCloseIndex, entry.sourceOp, entry.source)) return;
      const column = document.createElement('div');
      column.className = 'detail-launcher-column';
      const urlInput = document.createElement('input');
      urlInput.type = 'text';
      urlInput.placeholder = 'https://www.stirlingsupply.co';
      const key = `${entry.sourceOp}.${entry.source}`;
      const ensureManagedSets = () => {
        if (!state.parseResult.insertClickedKeys) state.parseResult.insertClickedKeys = new Set();
        state.parseResult.insertClickedKeys.add(key);
        if (!state.parseResult.buttonExactInsert) state.parseResult.buttonExactInsert = new Set();
        state.parseResult.buttonExactInsert.add(key);
      };
      const hasManagedKey = () => {
        return Boolean(
          (state.parseResult.insertClickedKeys && state.parseResult.insertClickedKeys.has(key)) ||
          (state.parseResult.buttonExactInsert && state.parseResult.buttonExactInsert.has(key))
        );
      };
      let isManaged = hasManagedKey();
      let autoUpdateNotified = false;
      try {
        if (state.parseResult.insertUrlMap.has(key)) {
          const storedUrl = state.parseResult.insertUrlMap.get(key) || '';
          const normalized = normalizeLauncherUrl(storedUrl);
          urlInput.value = normalized || storedUrl;
          if (normalized && normalized !== storedUrl) {
            state.parseResult.insertUrlMap.set(key, normalized);
          }
        } else {
          const grp2 = findEnclosingGroupForIndex(state.originalText, state.parseResult.inputs.openIndex);
          if (grp2) {
            const tb2 = findToolBlockInGroup(state.originalText, grp2.groupOpenIndex, grp2.groupCloseIndex, entry.sourceOp);
            let uc2 = tb2 ? findUserControlsInTool(state.originalText, tb2.open, tb2.close) : findUserControlsInGroup(state.originalText, grp2.groupOpenIndex, grp2.groupCloseIndex);
            if (uc2) {
              const cb2 = findControlBlockInUc(state.originalText, uc2.open, uc2.close, entry.source);
              if (cb2) {
                const body = state.originalText.slice(cb2.open + 1, cb2.close);
                const m = body.match(/BTNCS_Execute\s*=\s*"((?:[^"\\]|\\.)*)"/);
                if (m && m[1]) {
                  const lua = unescapeSettingString(m[1]);
                  const mUrl = lua && lua.match(/local\s+url\s*=\s*"([^"]*)"/);
                  if (mUrl && mUrl[1]) {
                    const normalized = normalizeLauncherUrl(mUrl[1]);
                    urlInput.value = normalized || mUrl[1];
                    state.parseResult.insertUrlMap.set(key, normalized || mUrl[1]);
                    ensureManagedSets();
                    isManaged = true;
                  }
                }
              }
            }
          }
          const lua = unescapeSettingString(m[1]);
          emitScript(lua, { source: 'existing' });
        }
      } catch (_) {}
      let lastQueuedUrl = normalizeLauncherUrl(urlInput.value);
      urlInput.addEventListener('input', () => {
        try {
          const normalized = normalizeLauncherUrl(urlInput.value);
          if (normalized === lastQueuedUrl) return;
          lastQueuedUrl = normalized;
          state.parseResult.insertUrlMap.set(key, normalized);
          if (isManaged && normalized) {
            ensureManagedSets();
            if (!autoUpdateNotified) {
              info('Launcher URL updated. Export to apply.');
              autoUpdateNotified = true;
            }
            emitScriptFromUrl(normalized, { source: 'url-input' });
          }
        } catch (_) {}
      });
      const btn = document.createElement('button');
      btn.type = 'button';
      const updateButtonState = () => {
        if (isManaged) {
          btn.textContent = 'Launcher Auto-Updates';
          btn.disabled = true;
          btn.title = 'Launcher already inserted; editing the URL will update it on export.';
        } else {
          btn.textContent = 'Insert Link Launcher';
          btn.disabled = false;
          btn.title = 'Insert cross-platform URL launcher using the URL field';
        }
      };
      updateButtonState();
      btn.style.alignSelf = 'flex-start';
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        try {
          let vNow = String(urlInput.value || '').trim();
          if (!vNow) {
            const fallback = (state.parseResult.insertUrlMap && state.parseResult.insertUrlMap.get(key)) || 'https://www.stirlingsupply.co';
            vNow = String(fallback);
            urlInput.value = vNow;
          }
          const normalized = normalizeLauncherUrl(vNow);
          if (normalized && normalized !== vNow) {
            urlInput.value = normalized;
          }
          state.parseResult.insertUrlMap.set(key, normalized);
          ensureManagedSets();
          isManaged = true;
          autoUpdateNotified = false;
          updateButtonState();
          info('Launcher will be inserted on export.');
          emitScriptFromUrl(normalized, { source: 'insert-click' });
        } catch (_) {}
      });
      column.appendChild(btn);
      column.appendChild(urlInput);
      container.appendChild(column);
    } catch (_) {}
  }

  function getCgKey(e) {
    return `${e.sourceOp}::${e.controlGroup}`;
  }

  function getContiguousColorGroupBlock(idx, entries, order) {
    const entry = entries[idx];
    if (!entry || !entry.sourceOp || !Number.isFinite(entry.controlGroup)) return null;
    const key = getCgKey(entry);
    const startPos = order.indexOf(idx);
    if (startPos < 0) return null;
    let first = startPos;
    let last = startPos;
    for (let p = startPos - 1; p >= 0; p--) {
      const e2 = entries[order[p]];
      if (!e2 || e2.isLabel || !Number.isFinite(e2.controlGroup) || e2.sourceOp !== entry.sourceOp) break;
      if (getCgKey(e2) !== key) break;
      first = p;
    }
    for (let p = startPos + 1; p < order.length; p++) {
      const e2 = entries[order[p]];
      if (!e2 || e2.isLabel || !Number.isFinite(e2.controlGroup) || e2.sourceOp !== entry.sourceOp) break;
      if (getCgKey(e2) !== key) break;
      last = p;
    }
    const indices = order.slice(first, last + 1);
    return { firstIndex: order[first], lastIndex: order[last], indices, count: indices.length, key };
  }

  function getEntryPage(entry) {
    if (!entry || !entry.page) return 'Controls';
    return entry.page;
  }

  function populatePageOptions(select, current) {
    while (select.firstChild) select.removeChild(select.firstChild);
    const options = buildPageOptions();
    const known = new Set(options);
    if (current && !known.has(current)) options.push(current);
    for (const name of options) {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      select.appendChild(opt);
    }
    const newOpt = document.createElement('option');
    newOpt.value = '__new__';
    newOpt.textContent = 'Add page…';
    select.appendChild(newOpt);
  }
  function promptForNewPage(container, select, options = {}) {
    if (!container) return;
    const entry = options && options.entry ? options.entry : null;
    const onCommitMultiple = options && typeof options.applyMultiple === 'function' ? options.applyMultiple : null;
    const fallbackValue = entry ? (getEntryPage(entry) || 'Controls') : 'Controls';
    let wrapper = container.querySelector('.new-page-input');
    if (wrapper) {
      const inputExisting = wrapper.querySelector('input');
      if (inputExisting) {
        inputExisting.value = '';
        inputExisting.focus();
      }
      return;
    }
    wrapper = document.createElement('div');
    wrapper.className = 'new-page-input';
    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = 'Page name';
    wrapper.appendChild(input);
    container.appendChild(wrapper);
    let finalized = false;
    const cleanup = () => {
      if (wrapper && wrapper.parentNode === container) container.removeChild(wrapper);
    };
    const commit = () => {
      if (finalized) return;
      finalized = true;
      const val = input.value ? input.value.trim() : '';
      if (!val) {
        select.value = fallbackValue;
        cleanup();
        return;
      }
      addPageOption(val);
      if (onCommitMultiple) onCommitMultiple(val);
      else if (entry) setEntryPage(entry, val);
      ensurePageOrderIncludes(val);
      setActivePage(val);
      refreshAllPageSelects();
      cleanup();
    };
    const cancel = () => {
      if (finalized) return;
      finalized = true;
      select.value = fallbackValue;
      cleanup();
    };
    input.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
      if (ev.key === 'Escape') { ev.preventDefault(); cancel(); }
    });
    input.addEventListener('blur', () => {
      setTimeout(() => {
        if (!finalized) cancel();
      }, 0);
    });
    input.focus();
  }

  function refreshAllPageSelects() {
    try {
      const selects = controlsList ? controlsList.querySelectorAll('.page-select select') : [];
      const current = Array.from(selects);
      for (const select of current) {
        const li = select.closest('li');
        const idx = li ? parseInt(li.dataset.index || '-1', 10) : -1;
        const entry = (idx >= 0 && state.parseResult && Array.isArray(state.parseResult.entries)) ? state.parseResult.entries[idx] : null;
        const pageName = entry ? getEntryPage(entry) : 'Controls';
        populatePageOptions(select, pageName);
        select.value = pageName;
      }
    } catch (_) {}
  }

  function buildPageOptions() {
    try {
      pruneManualPageOptions();
      const set = new Set(['Controls']);
      if (state.parseResult && Array.isArray(state.parseResult.entries)) {
        for (const entry of state.parseResult.entries) {
          const pageName = entry && entry.page ? String(entry.page).trim() : '';
          if (pageName) set.add(pageName);
        }
      }
      manualPageOptions.forEach(p => set.add(p));
      return Array.from(set);
    } catch (_) { return ['Controls']; }
  }

  function addPageOption(name) {
    try {
      if (!name) return;
      manualPageOptions.add(name);
    } catch (_) {}
  }

  function pruneManualPageOptions() {
    try {
      const used = new Set();
      if (state.parseResult && Array.isArray(state.parseResult.entries)) {
        state.parseResult.entries.forEach((entry) => {
          if (!entry || !entry.page) return;
          const norm = normalizePageName(entry.page);
          if (norm && norm !== 'Controls') used.add(norm);
        });
      }
      const toDelete = [];
      manualPageOptions.forEach((name) => {
        const norm = normalizePageName(name);
        if (!norm || norm === 'Controls' || !used.has(norm)) toDelete.push(name);
      });
      toDelete.forEach(name => manualPageOptions.delete(name));
      if (state.parseResult && Array.isArray(state.parseResult.pageOrder)) {
        state.parseResult.pageOrder = state.parseResult.pageOrder
          .map(normalizePageName)
          .filter((name, idx, arr) => {
            if (!name || name === 'Controls') return idx === arr.indexOf(name);
            return used.has(name);
          });
      }
      const currentActive = activePage || (state.parseResult ? state.parseResult.activePage : null);
      const normalizedActive = currentActive ? normalizePageName(currentActive) : null;
      if (normalizedActive && normalizedActive !== 'Controls' && !used.has(normalizedActive)) {
        setActivePage('Controls', { silent: true });
      }
      refreshPageTabsInternal({ deferRenderView: true });
    } catch (_) {}
  }

  function setEntryPage(entry, pageName) {
    if (!entry) return;
    applyPageToEntries([entry], pageName);
  }

  function getEntriesByIndices(indices) {
    const results = [];
    try {
      if (!state.parseResult || !Array.isArray(state.parseResult.entries) || !indices) return results;
      indices.forEach((idx) => {
        if (idx == null) return;
        const entry = state.parseResult.entries[idx];
        if (entry) results.push(entry);
      });
    } catch (_) {}
    return results;
  }

  function applyPageToEntries(entriesList, pageName) {
    try {
      if (!entriesList || !entriesList.length || !state.parseResult || !Array.isArray(state.parseResult.entries)) return;
      const value = pageName && pageName !== 'Controls' ? pageName : 'Controls';
      let changed = false;
      for (const entry of entriesList) {
        if (!entry) continue;
        if (entry.page !== value) changed = true;
        entry.page = value;
        if (typeof entry.raw === 'string') {
          const updated = ensurePageProp(entry.raw, value);
          if (updated !== entry.raw) {
            entry.raw = updated;
            changed = true;
          }
        }
      }
      ensurePageOrderIncludes(value);
      syncPageOrderFromState();
      pruneManualPageOptions();
      if (changed) {
        renderList(state.parseResult.entries, state.parseResult.order);
        entriesList.forEach((entry) => {
          if (!entry || !state.parseResult || !Array.isArray(state.parseResult.entries)) return;
          const idx = state.parseResult.entries.indexOf(entry);
          notifyEntryMutated(idx);
        });
      } else {
        refreshPageTabsInternal({ deferRenderView: true });
      }
    } catch (_) {}
  }

  function buildEntryKey(entry) {
    if (!entry) return '';
    return `${entry.sourceOp || ''}::${entry.source || ''}`;
  }

  function ensureBlendToggleSetLocal() {
    try {
      if (!state || !state.parseResult) return new Set();
      if (state.parseResult.blendToggles instanceof Set) return state.parseResult.blendToggles;
      const initial = new Set();
      if (state.parseResult.blendToggles && typeof state.parseResult.blendToggles.forEach === 'function') {
        state.parseResult.blendToggles.forEach((value) => {
          try { initial.add(value); } catch (_) {}
        });
      } else if (Array.isArray(state.parseResult.blendToggles)) {
        state.parseResult.blendToggles.forEach((value) => initial.add(value));
      }
      state.parseResult.blendToggles = initial;
      return initial;
    } catch (_) {
      const fallback = new Set();
      if (state && state.parseResult) state.parseResult.blendToggles = fallback;
      return fallback;
    }
  }

  function setBlendToggleState(entry, enabled) {
    if (!entry || !state || !state.parseResult) return;
    entry.isBlendToggle = !!enabled;
    const set = ensureBlendToggleSetLocal();
    const key = buildEntryKey(entry);
    if (!key) return;
    if (entry.isBlendToggle) set.add(key);
    else set.delete(key);
    state.parseResult.blendToggles = set;
  }

  function isBlendEntry(entry) {
    if (!entry || !entry.source) return false;
    return String(entry.source).toLowerCase() === 'blend';
  }

  function isBlendToggleEnabled(entry) {
    return !!(entry && entry.isBlendToggle);
  }

  function toggleBlendCheckbox(entry) {
    if (typeof pushHistory === 'function') {
      pushHistory('toggle blend control');
    }
    if (!entry) return;
    const nextState = !entry.isBlendToggle;
    entry.isBlendToggle = nextState;
    setBlendToggleState(entry, nextState);
    entry.controlMeta = entry.controlMeta || {};
    entry.controlMetaOriginal = entry.controlMetaOriginal || {};
    entry.controlMeta.inputControl = nextState ? '"CheckboxControl"' : '"SliderControl"';
    entry.controlMetaDirty = true;
    entry.controlTypeEdited = true;
    updateEntryRawForBlend(entry);
    renderList(state.parseResult.entries, state.parseResult.order);
    if (state.parseResult && Array.isArray(state.parseResult.entries)) {
      const idx = state.parseResult.entries.indexOf(entry);
      notifyEntryMutated(idx);
    }
  }

  function updateEntryRawForBlend(entry) {
    try {
      if (!entry || typeof entry.raw !== 'string') return;
      const raw = entry.raw;
      const open = raw.indexOf('{');
      const close = raw.lastIndexOf('}');
      if (open < 0 || close <= open) return;
      let body = raw.slice(open + 1, close);
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
      entry.raw = raw.slice(0, open + 1) + body + raw.slice(close);
    } catch (_) {}
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
    } catch (_) {
      return body;
    }
  }

  function removeInstanceInputProp(body, prop) {
    try {
      const pattern = new RegExp(`(^|\\r?\\n)\\s*${prop}\\s*=\\s*([^,\\n\\r}]*)\\s*,?`, 'g');
      return body.replace(pattern, '$1');
    } catch (_) {
      return body;
    }
  }

  function toggleLabelVisibility(entry) {
    if (!entry) return;
    pushHistory('toggle label visibility');
    entry.labelHidden = !entry.labelHidden;
    renderList(state.parseResult.entries, state.parseResult.order);
    if (state.parseResult && Array.isArray(state.parseResult.entries)) {
      const idx = state.parseResult.entries.indexOf(entry);
      notifyEntryMutated(idx);
    }
  }

  function ensurePageProp(raw, pageName) {
    try {
      const open = raw.indexOf('{'); const close = raw.lastIndexOf('}');
      if (open < 0 || close <= open) return raw;
      const head = raw.slice(0, open + 1);
      let body = raw.slice(open + 1, close);
      const tail = raw.slice(close);
      const normalized = (pageName && String(pageName).trim()) || '';
      const pageRegex = /(\r?\n)?\s*Page\s*=\s*"([^"]*)"\s*,?/g;
      if (!normalized || normalized === 'Controls') {
        if (!pageRegex.test(body)) return raw;
        body = body.replace(pageRegex, '');
        return head + body + tail;
      }
      const pageLine = `Page = "${normalized}",`;
      if (pageRegex.test(body)) {
        body = body.replace(pageRegex, pageLine);
        return head + body + tail;
      }
      return head + pageLine + body + tail;
    } catch (_) { return raw; }
  }

  function getInsertionAnchorPos() {
    return insertionAnchorPos;
  }

  return {
    renderList,
    setFilter,
    updateRemoveSelectedState,
    cleanupDropIndicator,
    updateDropIndicatorPosition,
    getInsertionPosUnderSelection,
    addOrMovePublishedItemsAt,
    createIcon,
    getCurrentDropIndex,
    setHighlightHandler: (fn) => {
      if (typeof fn === 'function') highlightHandler = fn;
    },
    resetPageOptions: () => { manualPageOptions.clear(); },
    refreshPageTabs: () => { refreshPageTabsInternal(); },
    setEntryDisplayName: setEntryDisplayNameByIndex,
    appendLauncherUi: (container, entry, options) => appendLauncherUi(container, entry, options),
  };
}
  function isInteractiveTarget(node) {
    if (!node || typeof node.closest !== 'function') return false;
    if (node.closest('button,select,input,textarea')) return true;
    const editable = node.closest('[contenteditable="true"]');
    return !!editable;
  }
