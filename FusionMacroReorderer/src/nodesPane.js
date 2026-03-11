import { findEnclosingGroupForIndex, extractQuotedProp } from './parser.js';
import { findMatchingBrace, isSpace, isIdentStart, isIdentPart, humanizeName, isQuoteEscaped } from './textUtils.js';

export function createNodesPane(options = {}) {
  const {
    state,
    nodesList,
    nodesSearch,
    hideReplacedEl,
    quickClickHintEl = null,
    viewControlsBtn = null,
    viewControlsMenu = null,
    showAllNodesBtn = null,
    showPublishedNodesBtn = null,
    collapseAllNodesBtn = null,
    showNodeTypeLabelsEl = null,
    showInstancedConnectionsEl = null,
    showNextNodeLinksEl = null,
    autoGroupQuickSetsEl = null,
    nameClickQuickSetEl = null,
    quickSetBlendToCheckboxEl = null,
    publishSelectedBtn,
    clearNodeSelectionBtn,
    importCatalogBtn,
    catalogInput,
    importModifierCatalogBtn,
    modifierCatalogInput,
    logDiag = () => {},
    logTag = () => {},
    error = () => {},
    info = () => {},
    highlightNode: initialHighlightNode = () => {},
    clearHighlights = () => {},
    renderList: renderPublishedList = () => {},
    getInsertionPosUnderSelection = () => 0,
    insertIndicesAt = (order) => order,
    isPublished = () => false,
    ensurePublished = () => null,
    ensureEntryExists = () => null,
    removePublished = () => {},
    createIcon = () => '',
    sanitizeIdent = (s) => String(s),
    normalizeId = (s) => String(s),
    enableNodeDrag = true,
    requestRenameNode = null,
    requestAddControl = null,
    focusPublishedControl = null,
    getPendingControlMeta = () => null,
    consumePendingControlMeta = () => {},
    isPickSessionActive = () => false,
  } = options;
  const UTILITY_NODE_NAME = 'UTILITY';
  const QUICK_SET_STORAGE_KEY = 'fmr.nodeQuickSets.v1';
  const QUICK_SET_OPTIONS_STORAGE_KEY = 'fmr.nodeQuickSetOptions.v1';
  const VIEW_CONTROLS_STORAGE_KEY = 'fmr.nodesViewControls.v1';
  const QUICK_SET_FALLBACK_COUNT = 10;

  let nodeCatalog = null;
  let modifierCatalog = null;
  let nodeProfiles = null;
  let modifierProfiles = null;
  let modifierBindingContext = new Map();
  let modifierContextDiagSeen = new Set();
  let maskPathDriverNodeNames = new Set();
  let maskPathDriverNodeNamesLower = new Set();
  let nodeFilter = '';
  let hideReplaced = false;
  let showNodeTypeLabels = true;
  let showInstancedConnections = true;
  let showNextNodeLinks = true;
  let autoGroupQuickSets = false;
  let nameClickQuickSet = false;
  let quickSetBlendToCheckbox = false;
  let highlightCallback = initialHighlightNode || (() => {});
  let lastNodeNames = [];
  let nodeContextMenu = null;
  let nodeContextMenuCleanup = null;
  let exprInspectorOverlay = null;
  let exprInspectorCleanup = null;
  let quickSetStore = loadQuickSetStore();
  let quickSetOptions = loadQuickSetOptions();
  const savedViewOptions = loadViewControlOptions();
  if (savedViewOptions) {
    hideReplaced = savedViewOptions.hideReplaced === true;
    showNodeTypeLabels = savedViewOptions.showNodeTypeLabels !== false;
    showInstancedConnections = savedViewOptions.showInstancedConnections !== false;
    showNextNodeLinks = savedViewOptions.showNextNodeLinks !== false;
    autoGroupQuickSets = savedViewOptions.autoGroupQuickSets === true;
    nameClickQuickSet = savedViewOptions.nameClickQuickSet === true;
    quickSetBlendToCheckbox = savedViewOptions.quickSetBlendToCheckbox === true;
  }

  function loadQuickSetStore() {
    try {
      if (typeof localStorage === 'undefined') return {};
      const raw = localStorage.getItem(QUICK_SET_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_) {
      return {};
    }
  }

  function saveQuickSetStore() {
    try {
      if (typeof localStorage === 'undefined') return;
      localStorage.setItem(QUICK_SET_STORAGE_KEY, JSON.stringify(quickSetStore || {}));
    } catch (_) {}
  }

  function loadQuickSetOptions() {
    try {
      if (typeof localStorage === 'undefined') return { resolveToDrivers: true };
      const raw = localStorage.getItem(QUICK_SET_OPTIONS_STORAGE_KEY);
      if (!raw) return { resolveToDrivers: true };
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return { resolveToDrivers: true };
      return {
        resolveToDrivers: parsed.resolveToDrivers !== false,
      };
    } catch (_) {
      return { resolveToDrivers: true };
    }
  }

  function saveQuickSetOptions() {
    try {
      if (typeof localStorage === 'undefined') return;
      localStorage.setItem(QUICK_SET_OPTIONS_STORAGE_KEY, JSON.stringify(quickSetOptions || { resolveToDrivers: true }));
    } catch (_) {}
  }

  function getQuickSetResolveToDrivers() {
    return quickSetOptions?.resolveToDrivers !== false;
  }

  function setQuickSetResolveToDrivers(value) {
    quickSetOptions = {
      ...(quickSetOptions || {}),
      resolveToDrivers: value !== false,
    };
    saveQuickSetOptions();
  }

  function loadViewControlOptions() {
    try {
      if (typeof localStorage === 'undefined') return null;
      const raw = localStorage.getItem(VIEW_CONTROLS_STORAGE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return null;
      return parsed;
    } catch (_) {
      return null;
    }
  }

  function saveViewControlOptions() {
    try {
      if (typeof localStorage === 'undefined') return;
      localStorage.setItem(VIEW_CONTROLS_STORAGE_KEY, JSON.stringify({
        hideReplaced: !!hideReplaced,
        showNodeTypeLabels: !!showNodeTypeLabels,
        showInstancedConnections: !!showInstancedConnections,
        showNextNodeLinks: !!showNextNodeLinks,
        autoGroupQuickSets: !!autoGroupQuickSets,
        nameClickQuickSet: !!nameClickQuickSet,
        quickSetBlendToCheckbox: !!quickSetBlendToCheckbox,
      }));
    } catch (_) {}
  }

  function yieldToPickSession(ev) {
    try {
      if (!isPickSessionActive()) return false;
      if (ev) {
        ev.preventDefault();
        ev.stopPropagation();
      }
      return true;
    } catch (_) {
      return false;
    }
  }

  function syncViewControlInputs() {
    try {
      if (hideReplacedEl) hideReplacedEl.checked = !!hideReplaced;
      if (showNodeTypeLabelsEl) showNodeTypeLabelsEl.checked = !!showNodeTypeLabels;
      if (showInstancedConnectionsEl) showInstancedConnectionsEl.checked = !!showInstancedConnections;
      if (showNextNodeLinksEl) showNextNodeLinksEl.checked = !!showNextNodeLinks;
      if (autoGroupQuickSetsEl) autoGroupQuickSetsEl.checked = !!autoGroupQuickSets;
      if (nameClickQuickSetEl) nameClickQuickSetEl.checked = !!nameClickQuickSet;
      if (quickSetBlendToCheckboxEl) quickSetBlendToCheckboxEl.checked = !!quickSetBlendToCheckbox;
      if (quickClickHintEl) {
        quickClickHintEl.hidden = !nameClickQuickSet;
        quickClickHintEl.textContent = 'Quick-click ON';
      }
    } catch (_) {}
  }

  function setViewControlsMenuOpen(open) {
    try {
      if (!viewControlsMenu || !viewControlsBtn) return;
      const next = !!open;
      viewControlsMenu.hidden = !next;
      viewControlsBtn.setAttribute('aria-expanded', next ? 'true' : 'false');
    } catch (_) {}
  }

  function setAllNodeModes(mode = 'open') {
    try {
      if (!state.parseResult) return;
      if (!(state.parseResult.nodesCollapsed instanceof Set)) state.parseResult.nodesCollapsed = new Set();
      if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
      const collapsed = state.parseResult.nodesCollapsed;
      const publishedOnly = state.parseResult.nodesPublishedOnly;
      collapsed.clear();
      publishedOnly.clear();
      const nodes = Array.isArray(state.parseResult.nodes) ? state.parseResult.nodes : [];
      for (const n of nodes) {
        const nodeControlsRaw = ensureDissolveMixControl(n, n.controls || []);
        const nodeControls = dedupeControlsByResolvedSource(nodeControlsRaw);
        const hasPublished = (nodeControls || []).some(c => {
          if (isSectionHeaderControl(c)) return false;
          if (isGroupedControl(c)) {
            return (c.channels || []).some((ch) => {
              const src = resolveControlSource(ch);
              return !!src && isPublished(resolvePublishSourceOp(n, ch), src);
            });
          }
          const src = resolveControlSource(c);
          return src ? isPublished(resolvePublishSourceOp(n, c), src) : false;
        });
        if (mode === 'closed') {
          collapsed.add(n.name);
          continue;
        }
        if (mode === 'published' && hasPublished) {
          publishedOnly.add(n.name);
        }
      }
      state.parseResult.nodesCollapsed = collapsed;
      state.parseResult.nodesPublishedOnly = publishedOnly;
      state.parseResult.nodesViewInitialized = true;
      parseAndRenderNodes();
    } catch (_) {}
  }

  function sanitizeQuickSetKeys(keys, candidatesMap) {
    const out = [];
    const seen = new Set();
    for (const raw of (keys || [])) {
      const key = String(raw || '').trim();
      if (!key || seen.has(key)) continue;
      if (candidatesMap && !candidatesMap.has(key)) continue;
      seen.add(key);
      out.push(key);
    }
    return out;
  }

  function getQuickSetForType(type, candidatesMap) {
    const key = String(type || '').trim();
    if (!key) return [];
    const stored = quickSetStore && Array.isArray(quickSetStore[key]) ? quickSetStore[key] : [];
    return sanitizeQuickSetKeys(stored, candidatesMap);
  }

  function setQuickSetForType(type, keys) {
    const key = String(type || '').trim();
    if (!key) return;
    const clean = sanitizeQuickSetKeys(keys);
    if (!clean.length) {
      delete quickSetStore[key];
    } else {
      quickSetStore[key] = clean;
    }
    saveQuickSetStore();
  }

  function isBoilerplateQuickControlId(id) {
    const text = String(id || '').trim();
    if (!text) return true;
    if (/^Global(In|Out)$/i.test(text)) return true;
    if (/^ProcessMode$/i.test(text)) return true;
    if (/^UseFrameFormatSettings$/i.test(text)) return true;
    if (/^Depth$/i.test(text)) return true;
    if (/^Blank\d+$/i.test(text)) return true;
    if (/^(Comments|HideInputs|BaseLayerName)$/i.test(text)) return true;
    if (/^(FrameRenderScript|StartRenderScript|EndRenderScript)$/i.test(text)) return true;
    if (/^(LayerSpacer|Input_LayerSelect|EffectMask_LayerSelect)$/i.test(text)) return true;
    if (/^(ProcessLayers|ProcessLayersCustom)$/i.test(text)) return true;
    if (/^UseGPU$/i.test(text)) return true;
    return false;
  }

  function isProcessChannelQuickControlId(id) {
    const text = String(id || '').trim();
    if (!text) return false;
    return /^Process(?:Red|Green|Blue|Alpha)$/i.test(text);
  }

  function isBlendQuickControlId(id) {
    const text = String(id || '').trim();
    if (!text) return false;
    return /^Blend$/i.test(text);
  }

  function isBlendQuickControl(control) {
    if (!control) return false;
    if (isBlendQuickControlId(control.id)) return true;
    if (isBlendQuickControlId(control.source)) return true;
    if (isBlendQuickControlId(control.name)) return true;
    return false;
  }

  function isGamutControlId(id) {
    const text = String(id || '').trim();
    if (!text) return false;
    return /^Gamut(?:\.|$)/i.test(text);
  }

  function isGamutQuickControl(control) {
    if (!control) return false;
    if (isGamutControlId(control.id)) return true;
    if (isGamutControlId(control.source)) return true;
    if (isGamutControlId(control.name)) return true;
    return false;
  }

  function isGamutQuickCandidate(candidate) {
    if (!candidate) return false;
    if (candidate.kind === 'group') {
      const channels = Array.isArray(candidate.channels) ? candidate.channels : [];
      if (!channels.length) return false;
      return channels.every((ch) => isGamutQuickControl(ch));
    }
    if (candidate.kind === 'control') {
      return isGamutQuickControl(candidate.control || { id: candidate.id, name: candidate.label });
    }
    return false;
  }

  function isBoilerplateDisplayControlId(id, node = null) {
    try {
      const text = String(id || '').trim();
      if (!text) return false;
      if (/^Global(In|Out)$/i.test(text)) return true;
      if (/^ProcessMode$/i.test(text)) return true;
      if (/^PixelAspect$/i.test(text)) return true;
      if (/^UseFrameFormatSettings$/i.test(text)) return true;
      if (/^Depth$/i.test(text)) return true;
      if (/^(Width|Height)$/i.test(text)) {
        const nodeType = String(node?.type || '').trim();
        if (nodeType && isMaskToolTypeForPath(nodeType)) return false;
        return true;
      }
      return false;
    } catch (_) {
      return false;
    }
  }

  function prioritizeQuickSetCandidates(list) {
    const items = Array.isArray(list) ? [...list] : [];
    if (!items.length) return items;
    const priority = (cand) => {
      if (!cand) return 10;
      if (cand.kind === 'control' && isBlendQuickControl(cand.control || { id: cand.id, name: cand.label })) return 0;
      if (cand.kind === 'group') {
        const channels = Array.isArray(cand.channels) ? cand.channels : [];
        if (channels.some((ch) => isBlendQuickControl(ch))) return 1;
      }
      return 5;
    };
    return items
      .map((cand, idx) => ({ cand, idx, p: priority(cand) }))
      .sort((a, b) => (a.p - b.p) || (a.idx - b.idx))
      .map((it) => it.cand);
  }

  function isBlendNodeControl(control) {
    if (!control) return false;
    if (isGroupedControl(control)) {
      const channels = Array.isArray(control.channels) ? control.channels : [];
      return channels.some((ch) => isBlendQuickControl(ch));
    }
    return isBlendQuickControl(control);
  }

  function isDeprioritizedExpandedControl(control, node = null) {
    try {
      if (!control) return false;
      if (isGroupedControl(control)) {
        const channels = Array.isArray(control.channels) ? control.channels : [];
        if (!channels.length) return false;
        return channels.every((ch) => {
          const id = String(ch?.id || '').trim();
          return isGamutControlId(id) || isBoilerplateDisplayControlId(id, node);
        });
      }
      const id = String(control.id || control.source || '').trim();
      return isGamutControlId(id) || isBoilerplateDisplayControlId(id, node);
    } catch (_) {
      return false;
    }
  }

  function prioritizeBlendControlsFlat(controls, node = null) {
    const list = Array.isArray(controls) ? controls : [];
    if (!list.length) return list;
    const blend = [];
    const rest = [];
    const deprioritized = [];
    for (const control of list) {
      if (isBlendNodeControl(control)) blend.push(control);
      else if (isDeprioritizedExpandedControl(control, node)) deprioritized.push(control);
      else rest.push(control);
    }
    const next = blend.concat(rest, deprioritized);
    return next.length ? next : list;
  }

  function prioritizeExpandedNodeControls(controls, node = null) {
    const list = Array.isArray(controls) ? controls : [];
    if (!list.length) return list;
    const hasSectionHeaders = list.some((control) => isSectionHeaderControl(control));
    if (!hasSectionHeaders) return prioritizeBlendControlsFlat(list, node);
    const out = [];
    let cursor = 0;
    while (cursor < list.length) {
      const current = list[cursor];
      if (isSectionHeaderControl(current)) {
        let end = cursor + 1;
        while (end < list.length && !isSectionHeaderControl(list[end])) end += 1;
        out.push(current, ...prioritizeBlendControlsFlat(list.slice(cursor + 1, end), node));
        cursor = end;
        continue;
      }
      let end = cursor;
      while (end < list.length && !isSectionHeaderControl(list[end])) end += 1;
      out.push(...prioritizeBlendControlsFlat(list.slice(cursor, end), node));
      cursor = end;
    }
    return out;
  }

  function isColorFocusedQuickSetNode(nodeType, nodeName = '') {
    try {
      const typeText = String(nodeType || '').trim();
      const nameText = String(nodeName || '').trim();
      const combined = `${typeText} ${nameText}`.toLowerCase();
      if (!combined) return false;
      // Broad color-tool matcher to keep process channel toggles only where useful.
      return /(color|hue|saturation|gamma|gain|contrast|lut|white\s*balance|chroma|luma|tint)/i.test(combined);
    } catch (_) {
      return false;
    }
  }

  function buildQuickSetCandidates(nodeControls, allowSourceControl, nodeMeta = null) {
    const list = [];
    const map = new Map();
    void nodeMeta;
    for (const c of (nodeControls || [])) {
      if (!c) continue;
      if (isGroupedControl(c)) {
        const base = String(c.base || c.id || '').trim();
        if (!base) continue;
        const key = `group:${base}`;
        const channels = Array.isArray(c.channels) ? c.channels.map(ch => ({ id: ch.id, name: ch.name || ch.id })) : [];
        if (!channels.length) continue;
        const item = {
          key,
          kind: 'group',
          label: c.groupLabel || humanizeName(base),
          base,
          channels,
          count: channels.length,
        };
        list.push(item);
        map.set(key, item);
        continue;
      }
      const id = String(c.id || '').trim();
      if (!id) continue;
      if (id === 'SourceOp') continue;
      if (id === 'Source' && !allowSourceControl) continue;
      const key = `control:${id}`;
      const item = {
        key,
        kind: 'control',
        id,
        label: c.name || humanizeName(id),
        control: c,
      };
      list.push(item);
      map.set(key, item);
    }
    return { list: prioritizeQuickSetCandidates(list), map };
  }

  function buildSuggestedQuickSetKeys(candidates, nodeMeta = null) {
    const typePreferred = buildTypePreferredQuickSetKeys(candidates, nodeMeta);
    if (typePreferred.length) return typePreferred.slice(0, QUICK_SET_FALLBACK_COUNT);
    const keepProcessChannels = isColorFocusedQuickSetNode(nodeMeta?.type, nodeMeta?.name);
    const primary = [];
    for (const cand of (candidates || [])) {
      if (!cand) continue;
      if (isGamutQuickCandidate(cand)) continue;
      if (cand.kind === 'group') {
        if (!keepProcessChannels) {
          const channels = Array.isArray(cand.channels) ? cand.channels : [];
          if (channels.length && channels.every((ch) => isProcessChannelQuickControlId(ch?.id))) continue;
        }
        primary.push(cand.key);
        continue;
      }
      if (!keepProcessChannels && isProcessChannelQuickControlId(cand.id)) continue;
      if (isBoilerplateQuickControlId(cand.id)) continue;
      primary.push(cand.key);
    }
    const fallback = (candidates || []).map(c => c && c.key).filter(Boolean);
    const preferred = primary.length ? primary : fallback;
    return preferred.slice(0, QUICK_SET_FALLBACK_COUNT);
  }

  function buildTypePreferredQuickSetKeys(candidates, nodeMeta = null) {
    try {
      const list = Array.isArray(candidates) ? candidates : [];
      if (!list.length) return [];
      const type = String(nodeMeta?.type || '').trim().toLowerCase();
      const isShapeSystemType = /^s[a-z]/i.test(type);
      const controls = list.filter((entry) => entry && entry.kind === 'control');
      const groups = list.filter((entry) => entry && entry.kind === 'group');
      const out = [];
      const seen = new Set();
      const pushKey = (key) => {
        const next = String(key || '').trim();
        if (!next || seen.has(next)) return;
        seen.add(next);
        out.push(next);
      };
      const pickControl = (id) => {
        const target = String(id || '').trim().toLowerCase();
        if (!target) return '';
        const cand = controls.find((entry) => String(entry.id || '').trim().toLowerCase() === target);
        return cand && cand.key ? cand.key : '';
      };
      const pickFirstControl = (ids = []) => {
        for (const id of (ids || [])) {
          const key = pickControl(id);
          if (key) return key;
        }
        return '';
      };
      const pickGroupByMatcher = (matcher) => {
        const isMatch = (value) => {
          const text = String(value || '').trim();
          if (!text) return false;
          if (matcher instanceof RegExp) return matcher.test(text);
          return text.toLowerCase() === String(matcher || '').trim().toLowerCase();
        };
        const cand = groups.find((entry) => isMatch(entry.label) || isMatch(entry.base));
        return cand && cand.key ? cand.key : '';
      };
      const addControl = (id) => pushKey(pickControl(id));
      const addFirstControl = (ids = []) => pushKey(pickFirstControl(ids));
      const addGroup = (matcher) => pushKey(pickGroupByMatcher(matcher));
      const addControlsMatching = (regex, sorter = null) => {
        const matches = controls.filter((entry) => regex.test(String(entry.id || '')));
        if (!matches.length) return;
        const sorted = sorter ? matches.slice().sort(sorter) : matches;
        sorted.forEach((entry) => pushKey(entry.key));
      };

      if (type === 'transform') {
        ['Center', 'Size', 'Angle'].forEach(addControl);
        return out;
      }

      if (type.includes('textplus')) {
        addControl('StyledText');
        addGroup(/font\s*\/\s*style/i);
        if (!out.some((key) => /^group:/i.test(key))) {
          addControl('Font');
          addControl('Style');
        }
        addGroup(/^color$/i);
        addControl('Size');
        addFirstControl(['Tracking', 'CharacterSpacing']);
        addFirstControl(['VerticalTopCenterBottom', 'VerticallyJustified']);
        addFirstControl(['HorizontalLeftCenterRight', 'HorizontallyJustified']);
        return out;
      }

      if (type.includes('background')) {
        addGroup(/^color$/i);
        if (!out.length) {
          addControl('TopLeftRed');
          addControl('TopLeftGreen');
          addControl('TopLeftBlue');
          addControl('TopLeftAlpha');
        }
        return out;
      }

      if (type === 'merge') {
        ['Blend', 'Center', 'Size', 'Angle'].forEach(addControl);
        return out;
      }

      if (type.includes('brightnesscontrast')) {
        ['Gain', 'Lift', 'Gamma', 'Contrast', 'Brightness', 'Saturation'].forEach(addControl);
        return out;
      }

      if (type.includes('colorcorrector')) {
        ['Hue', 'Saturation', 'Contrast', 'Gain', 'Lift', 'Gamma', 'Brightness'].forEach(addControl);
        return out;
      }

      if (type.includes('softglow')) {
        ['Blend', 'Threshold', 'Gain', 'XGlowSize'].forEach(addControl);
        return out;
      }

      if (type === 'glow') {
        ['Blend', 'XGlowSize', 'Glow'].forEach(addControl);
        return out;
      }

      if (type === 'blur') {
        ['Blend', 'LockXY', 'XBlurSize', 'YBlurSize'].forEach(addControl);
        return out;
      }

      if (type.includes('gaussianblur')) {
        ['Blend', 'IsSplitHV', 'HStrength', 'VStrength'].forEach(addControl);
        return out;
      }

      if (type.includes('directionalblur')) {
        ['Blend', 'Length', 'Angle', 'Glow'].forEach(addControl);
        return out;
      }

      if (type.startsWith('switch')) {
        addControl('Source');
        return out;
      }

      if (type.includes('dissolve')) {
        addControl('Mix');
        return out;
      }

      // Shape System (s*) defaults
      if (type === 'sboolean') {
        addControl('Operation');
        addControl('StyleMode');
        addGroup(/^color\b/i);
        if (!out.some((key) => /^group:/i.test(key))) {
          ['Red', 'Green', 'Blue', 'Alpha'].forEach(addControl);
        }
        addControl('Opacity');
        return out;
      }

      if (type === 'sbspline') {
        ['Solid', 'BorderWidth', 'WritePosition', 'WriteLength', 'Polyline'].forEach(addControl);
        return out;
      }

      if (type === 'schangestyle') {
        addGroup(/^color\b/i);
        if (!out.some((key) => /^group:/i.test(key))) {
          ['Red', 'Green', 'Blue', 'Alpha'].forEach(addControl);
        }
        addControl('Opacity');
        return out;
      }

      if (type === 'sduplicate') {
        ['Copies', 'TimeOffset', 'XOffset', 'YOffset', 'XSize', 'YSize', 'ZRotation'].forEach(addControl);
        return out;
      }

      if (type === 'sellipse') {
        ['Solid', 'BorderWidth', 'Width', 'Height', 'WritePosition', 'WriteLength'].forEach(addControl);
        return out;
      }

      if (type === 'sexpand') {
        ['Amount', 'JoinStyle'].forEach(addControl);
        return out;
      }

      if (type === 'sgrid') {
        ['CellsX', 'CellsY', 'XOffset', 'YOffset'].forEach(addControl);
        return out;
      }

      if (type === 'sjitter') {
        addGroup(/shape\s*offset\s*x/i);
        addGroup(/shape\s*offset\s*y/i);
        addGroup(/shape\s*size\s*x/i);
        addGroup(/shape\s*size\s*y/i);
        addGroup(/shape\s*rotate/i);
        return out;
      }

      if (type === 'smerge') {
        return out;
      }

      if (type === 'sngon') {
        ['Sides', 'Solid', 'BorderWidth', 'WritePosition', 'WriteLength', 'Height', 'Width'].forEach(addControl);
        return out;
      }

      if (type === 'soutline') {
        ['Thickness', 'JoinStyle', 'CapStyle', 'WritePosition', 'WriteLength'].forEach(addControl);
        return out;
      }

      if (type === 'spolygon') {
        ['Polyline', 'BorderWidth', 'Solid', 'CapStyle', 'WritePosition', 'WriteLength'].forEach(addControl);
        return out;
      }

      if (type === 'srectangle') {
        // sRectangle uses per-axis translate controls instead of a point-center control.
        addFirstControl(['Center', 'Translate.X']);
        addControl('Translate.Y');
        ['Width', 'Height', 'CornerRadius', 'BorderWidth', 'Solid', 'CapStyle', 'WritePosition', 'WriteLength'].forEach(addControl);
        return out;
      }

      if (type === 'srender') {
        return out;
      }

      if (type === 'sstar') {
        ['Points', 'Depth', 'Solid', 'BorderWidth', 'WritePosition', 'WriteLength', 'Height', 'Width'].forEach(addControl);
        return out;
      }

      if (type === 'stext') {
        addControl('StyledText');
        addGroup(/font\s*\/\s*style/i);
        if (!out.some((key) => /^group:/i.test(key))) {
          addControl('Font');
          addControl('Style');
        }
        addGroup(/^color\b/i);
        if (!out.some((key) => /^group:/i.test(key))) {
          addFirstControl(['Red1Clone', 'Red1']);
          addFirstControl(['Green1Clone', 'Green1']);
          addFirstControl(['Blue1Clone', 'Blue1']);
          addFirstControl(['Alpha1Clone', 'Alpha1']);
        }
        addControl('Size');
        addFirstControl(['Tracking', 'CharacterSpacingClone', 'CharacterSpacing']);
        addFirstControl(['VerticalJustificationNew', 'VerticalTopCenterBottom', 'VerticallyJustified']);
        addFirstControl(['HorizontalJustificationNew', 'HorizontalLeftCenterRight', 'HorizontallyJustified']);
        return out;
      }

      if (type === 'stransform') {
        ['XOffset', 'YOffset', 'XSize', 'YSize', 'ZRotation'].forEach(addControl);
        return out;
      }

      if (!isShapeSystemType && (type.includes('polygon') || type.includes('polyline') || type.includes('bspline'))) {
        ['Level', 'ShowViewControls', 'Polyline', 'BorderWidth', 'SoftEdge', 'Solid', 'CapStyle', 'WritePosition', 'WriteLength']
          .forEach((id) => addFirstControl([id, id.toLowerCase()]));
        return out;
      }

      if (!isShapeSystemType && type.includes('rectangle')) {
        ['Level', 'ShowViewControls', 'Center', 'Width', 'Height', 'CornerRadius', 'BorderWidth', 'SoftEdge', 'Solid', 'CapStyle', 'WritePosition', 'WriteLength']
          .forEach((id) => addFirstControl([id, id.toLowerCase()]));
        return out;
      }

      if (!isShapeSystemType && type.includes('ellipse')) {
        ['Level', 'ShowViewControls', 'Center', 'Width', 'Height', 'BorderWidth', 'Solid', 'CapStyle', 'WritePosition', 'WriteLength']
          .forEach((id) => addFirstControl([id, id.toLowerCase()]));
        return out;
      }

      if (type.includes('multimerge')) {
        const fieldOrder = { center: 0, size: 1, angle: 2 };
        addControlsMatching(/^Layer\d+\.(Center|Size|Angle)$/i, (a, b) => {
          const ma = String(a.id || '').match(/^Layer(\d+)\.(.+)$/i);
          const mb = String(b.id || '').match(/^Layer(\d+)\.(.+)$/i);
          const ai = ma ? Number(ma[1]) : Number.MAX_SAFE_INTEGER;
          const bi = mb ? Number(mb[1]) : Number.MAX_SAFE_INTEGER;
          if (ai !== bi) return ai - bi;
          const af = ma ? String(ma[2] || '').toLowerCase() : '';
          const bf = mb ? String(mb[2] || '').toLowerCase() : '';
          return (fieldOrder[af] ?? 99) - (fieldOrder[bf] ?? 99);
        });
        return out;
      }

      if (type.includes('multitext')) {
        const fieldOrder = { styledtext: 0, style: 1, size: 2, textfill: 3, center: 4 };
        addControlsMatching(/^Text\d+\.(StyledText|Style|Size|Text\d+Fill|Center)$/i, (a, b) => {
          const ma = String(a.id || '').match(/^Text(\d+)\.(.+)$/i);
          const mb = String(b.id || '').match(/^Text(\d+)\.(.+)$/i);
          const ai = ma ? Number(ma[1]) : Number.MAX_SAFE_INTEGER;
          const bi = mb ? Number(mb[1]) : Number.MAX_SAFE_INTEGER;
          if (ai !== bi) return ai - bi;
          const normalizeField = (field) => {
            const raw = String(field || '').toLowerCase();
            if (/^text\d+fill$/.test(raw)) return 'textfill';
            return raw;
          };
          const af = normalizeField(ma ? ma[2] : '');
          const bf = normalizeField(mb ? mb[2] : '');
          return (fieldOrder[af] ?? 99) - (fieldOrder[bf] ?? 99);
        });
        return out;
      }

      if (type.includes('multipoly')) {
        const fieldOrder = { level: 0, polyline: 1, polyline2: 2 };
        addControlsMatching(/^PolyMask\d+\.(Level|Polyline2?)$/i, (a, b) => {
          const ma = String(a.id || '').match(/^PolyMask(\d+)\.(.+)$/i);
          const mb = String(b.id || '').match(/^PolyMask(\d+)\.(.+)$/i);
          const ai = ma ? Number(ma[1]) : Number.MAX_SAFE_INTEGER;
          const bi = mb ? Number(mb[1]) : Number.MAX_SAFE_INTEGER;
          if (ai !== bi) return ai - bi;
          const af = ma ? String(ma[2] || '').toLowerCase() : '';
          const bf = mb ? String(mb[2] || '').toLowerCase() : '';
          return (fieldOrder[af] ?? 99) - (fieldOrder[bf] ?? 99);
        });
        addFirstControl(['Level']);
        addFirstControl(['Polyline']);
        return out;
      }

      if (type.includes('erodedilate')) {
        addControl('Amount');
        return out;
      }

      if (type.includes('displace')) {
        ['Center', 'Offset', 'RefractionStrength'].forEach(addControl);
        return out;
      }

      if (type.includes('letterbox')) {
        ['Mode', 'Width', 'Height', 'HiQOnly'].forEach(addControl);
        return out;
      }

      return out;
    } catch (_) {
      return [];
    }
  }

  function openQuickSetModal(node, nodeControls, allowSourceControl) {
    if (!node || !Array.isArray(nodeControls)) return;
    const { list, map } = buildQuickSetCandidates(nodeControls, allowSourceControl, node);
    if (!list.length) {
      info('No eligible controls found for quick set editing on this node.');
      return;
    }
    const suggested = buildSuggestedQuickSetKeys(list, node);
    const existing = getQuickSetForType(node.type, map);
    const initial = existing.length ? existing : suggested;
    const selected = new Set(initial);

    const overlay = document.createElement('div');
    overlay.className = 'quick-set-modal';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const panel = document.createElement('form');
    panel.className = 'quick-set-panel';

    const header = document.createElement('header');
    const titleWrap = document.createElement('div');
    const h3 = document.createElement('h3');
    h3.textContent = `Quick Set: ${node.type || node.name}`;
    const note = document.createElement('p');
    note.textContent = `${node.name} - choose controls for one-click quick publish`;
    titleWrap.appendChild(h3);
    titleWrap.appendChild(note);
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.textContent = 'x';
    closeBtn.setAttribute('aria-label', 'Close');
    header.appendChild(titleWrap);
    header.appendChild(closeBtn);

    const body = document.createElement('div');
    body.className = 'quick-set-body';
    const toolbar = document.createElement('div');
    toolbar.className = 'quick-set-toolbar';
    const useSuggestedBtn = document.createElement('button');
    useSuggestedBtn.type = 'button';
    useSuggestedBtn.textContent = 'Use Suggested';
    const selectAllBtn = document.createElement('button');
    selectAllBtn.type = 'button';
    selectAllBtn.textContent = 'Select All';
    const selectNoneBtn = document.createElement('button');
    selectNoneBtn.type = 'button';
    selectNoneBtn.textContent = 'Select None';
    toolbar.appendChild(useSuggestedBtn);
    toolbar.appendChild(selectAllBtn);
    toolbar.appendChild(selectNoneBtn);
    body.appendChild(toolbar);

    const optionsRow = document.createElement('label');
    optionsRow.className = 'quick-set-option';
    const resolveDriversCb = document.createElement('input');
    resolveDriversCb.type = 'checkbox';
    resolveDriversCb.checked = getQuickSetResolveToDrivers();
    const resolveDriversText = document.createElement('span');
    resolveDriversText.textContent = 'Quick publish: resolve selected driven controls to their driver controls';
    optionsRow.appendChild(resolveDriversCb);
    optionsRow.appendChild(resolveDriversText);
    body.appendChild(optionsRow);

    const listEl = document.createElement('div');
    listEl.className = 'quick-set-list';
    const refreshChecks = () => {
      const checks = listEl.querySelectorAll('input[type="checkbox"][data-key]');
      checks.forEach((cb) => {
        const key = cb.dataset.key || '';
        cb.checked = selected.has(key);
      });
    };
    for (const cand of list) {
      if (!cand || !cand.key) continue;
      const row = document.createElement('label');
      row.className = 'quick-set-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.dataset.key = cand.key;
      cb.checked = selected.has(cand.key);
      cb.addEventListener('change', () => {
        if (cb.checked) selected.add(cand.key);
        else selected.delete(cand.key);
      });
      const text = document.createElement('span');
      text.className = 'quick-set-label';
      if (cand.kind === 'group') {
        const count = Number.isFinite(cand.count) ? ` (${cand.count})` : '';
        text.textContent = `${cand.label}${count}`;
      } else {
        text.textContent = cand.label;
      }
      row.appendChild(cb);
      row.appendChild(text);
      listEl.appendChild(row);
    }
    body.appendChild(listEl);

    const footer = document.createElement('footer');
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'submit';
    saveBtn.className = 'primary';
    saveBtn.textContent = 'Save Set';
    footer.appendChild(cancelBtn);
    footer.appendChild(saveBtn);

    panel.appendChild(header);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    closeBtn.addEventListener('click', close);
    cancelBtn.addEventListener('click', close);
    overlay.addEventListener('click', (ev) => {
      if (ev.target === overlay) close();
    });
    useSuggestedBtn.addEventListener('click', () => {
      selected.clear();
      suggested.forEach((key) => selected.add(key));
      refreshChecks();
    });
    selectAllBtn.addEventListener('click', () => {
      selected.clear();
      list.forEach((cand) => selected.add(cand.key));
      refreshChecks();
    });
    selectNoneBtn.addEventListener('click', () => {
      selected.clear();
      refreshChecks();
    });
    panel.addEventListener('submit', (ev) => {
      ev.preventDefault();
      const keys = sanitizeQuickSetKeys(Array.from(selected), map);
      setQuickSetResolveToDrivers(!!resolveDriversCb.checked);
      setQuickSetForType(node.type, keys);
      info(`Saved quick set for ${node.type || node.name}: ${keys.length} control(s).`);
      close();
      if (state.parseResult) parseAndRenderNodes();
    });
  }

  function applyQuickSetToNode(node, nodeControls, allowSourceControl) {
    try {
      if (!state.parseResult || !node || !Array.isArray(nodeControls)) return;
      const resolveToDrivers = getQuickSetResolveToDrivers();
      const { list, map } = buildQuickSetCandidates(nodeControls, allowSourceControl, node);
      if (!list.length) {
        info('No eligible controls to publish from quick set.');
        return;
      }
      let keys = getQuickSetForType(node.type, map);
      if (!keys.length) {
        keys = buildSuggestedQuickSetKeys(list, node);
        if (keys.length) {
          setQuickSetForType(node.type, keys);
        }
      }
      if (!keys.length) {
        info(`No quick-set controls available for ${node.type || node.name}.`);
        return;
      }
      const publishMap = new Map();
      const modifierAdded = new Set();
      let missing = 0;
      let redirected = 0;

      const addPublishItem = (sourceOp, source, displayName, controlDef = null) => {
        const op = String(sourceOp || '').trim();
        const src = String(source || '').trim();
        if (!op || !src) return false;
        const key = `${op}::${src}`;
        if (publishMap.has(key)) return false;
        publishMap.set(key, {
          sourceOp: op,
          source: src,
          displayName: displayName || humanizeName(src) || src,
          controlDef,
        });
        return true;
      };

      const findControlDefinition = (sourceOp, source) => {
        try {
          const tool = findToolByNameAnywhere(state.originalText, sourceOp);
          if (!tool) return null;
          const isMod = isModifierType(tool.type || '');
          const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
          const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
          const contextBindings = isMod ? (modifierBindingContext.get(tool.name) || []) : [];
          const controls = deriveControlsForTool(tool, cat, profileRef, contextBindings) || [];
          const target = controls.find((c) => c && !c.type && String(c.id || '') === String(source));
          return target || null;
        } catch (_) {
          return null;
        }
      };

      const addExpressionTargets = (exprRefs) => {
        let addedAny = false;
        for (const expr of (exprRefs || [])) {
          for (const target of (expr?.targets || [])) {
            if (!target || !target.sourceOp || !target.source) continue;
            const def = findControlDefinition(target.sourceOp, target.source);
            const label = (def && (def.name || def.id)) || humanizeName(target.source) || target.source;
            if (addPublishItem(target.sourceOp, target.source, label, def)) addedAny = true;
          }
        }
        return addedAny;
      };

      const addDriverControl = (driverName, preferredSource = '') => {
        const mod = String(driverName || '').trim();
        if (!mod) return false;
        const requested = String(preferredSource || '').trim();
        const tool = findToolByNameAnywhere(state.originalText, mod);
        if (!tool) return false;
        const isMod = isModifierType(tool.type || '');
        const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
        const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
        const contextBindings = isMod ? (modifierBindingContext.get(tool.name) || []) : [];
        const controls = deriveControlsForTool(tool, cat, profileRef, contextBindings) || [];

        const addById = (id, fallbackLabel = '') => {
          const src = String(id || '').trim();
          if (!src) return false;
          const driverKey = `${mod}::${src}`;
          if (modifierAdded.has(driverKey)) return false;
          modifierAdded.add(driverKey);
          const def = findControlDefinition(mod, src);
          const label = (def && (def.name || def.id)) || fallbackLabel || src;
          return addPublishItem(mod, src, label, def);
        };

        if (isMod) {
          const allowSource = /switch/i.test(String(tool.type || tool.name || ''));
          const quick = buildQuickSetCandidates(controls, allowSource, { type: tool.type, name: tool.name });
          const savedKeys = getQuickSetForType(tool.type, quick.map);
          let changed = false;
          for (const key of (savedKeys || [])) {
            const cand = quick.map.get(key);
            if (!cand) continue;
            if (cand.kind === 'group') {
              for (const ch of (cand.channels || [])) {
                if (!ch || !ch.id) continue;
                changed = addById(ch.id, ch.name || ch.id) || changed;
              }
              continue;
            }
            changed = addById(cand.id, cand.label || cand.id) || changed;
          }
          if (changed) return true;
        }

        const pickFlatControl = (id) => {
          const targetId = String(id || '').trim();
          if (!targetId) return null;
          return controls.find((c) => c && !c.type && String(c.id || '') === targetId) || null;
        };

        let chosen = null;
        if (requested) chosen = pickFlatControl(requested);
        if (!chosen) chosen = pickFlatControl('Value');
        if (!chosen) chosen = controls.find((c) => c && !c.type) || null;
        if (!chosen || !chosen.id) return false;
        return addById(chosen.id, chosen.name || chosen.id);
      };

      const resolveAndAddControl = (sourceOp, source, displayName, controlDef) => {
        const effectiveSource = String(source || '').trim();
        const effectiveSourceOp = String(sourceOp || '').trim();
        const effectiveDisplayName = displayName;
        if (!resolveToDrivers || !effectiveSource) {
          addPublishItem(effectiveSourceOp, effectiveSource, effectiveDisplayName, controlDef);
          return;
        }
        const mods = [];
        for (const [modName, refs] of modifierBindingContext.entries()) {
          for (const ref of (refs || [])) {
            if (!ref) continue;
            if (String(ref.tool || '') !== String(effectiveSourceOp || '')) continue;
            if (String(ref.id || '') !== effectiveSource) continue;
            mods.push({ name: modName, source: String(ref.source || '').trim() });
          }
        }
        const exprRefs = (node && String(node.name || '') === String(effectiveSourceOp || '') && node.expressionDriven)
          ? (typeof node.expressionDriven.get === 'function' ? (node.expressionDriven.get(effectiveSource) || []) : (node.expressionDriven[effectiveSource] || []))
          : [];
        const hasDrivers = (mods.length > 0) || (exprRefs && exprRefs.length > 0);
        if (!hasDrivers) {
          addPublishItem(effectiveSourceOp, effectiveSource, effectiveDisplayName, controlDef);
          return;
        }
        redirected += 1;
        let driverAdded = false;
        if (exprRefs && exprRefs.length > 0) {
          driverAdded = addExpressionTargets(exprRefs) || driverAdded;
        }
        for (const modInfo of mods) {
          driverAdded = addDriverControl(modInfo?.name, modInfo?.source) || driverAdded;
        }
        if (!driverAdded) {
          addPublishItem(effectiveSourceOp, effectiveSource, effectiveDisplayName, controlDef);
        }
      };

      for (const key of keys) {
        const cand = map.get(key);
        if (!cand) {
          missing += 1;
          continue;
        }
        if (cand.kind === 'group') {
          for (const ch of (cand.channels || [])) {
            if (!ch || !ch.id) continue;
            const chDef = ch.control || findControlDefinition(node.name, ch.id);
            const label = (chDef && (chDef.name || chDef.id)) || ch.name || ch.id;
            resolveAndAddControl(node.name, ch.id, label, chDef);
          }
          continue;
        }
        const def = cand.control || findControlDefinition(node.name, cand.id);
        const label = cand.label || (def && (def.name || def.id)) || cand.id;
        resolveAndAddControl(node.name, cand.id, label, def);
      }

      const idxs = [];
      let added = 0;
      let already = 0;
      for (const item of publishMap.values()) {
        if (isPublished(item.sourceOp, item.source)) already += 1;
        else added += 1;
        const meta = buildQuickSetPublishMeta(item);
        const displayName = buildQuickSetPublishDisplayName(item, node);
        const idx = ensureEntryExists(item.sourceOp, item.source, displayName, meta);
        if (idx != null) idxs.push(idx);
      }
      const indicesToInsert = [...idxs];
      if (autoGroupQuickSets && idxs.length) {
        const nodeLabelBase = String(node.name || node.type || 'Node').trim() || 'Node';
        const labelDisplayName = `${nodeLabelBase} Controls`;
        const sourceOp = String(node.name || '').trim();
        const sourceBase = `MM_QuickSetLabel_${sanitizeIdent(nodeLabelBase) || 'Node'}`;
        let sourceId = sourceBase;
        let suffix = 2;
        const entries = Array.isArray(state.parseResult?.entries) ? state.parseResult.entries : [];
        while (entries.some((entry) => entry && String(entry.source || '') === sourceId)) {
          sourceId = `${sourceBase}_${suffix}`;
          suffix += 1;
        }
        let anchorControlGroup = null;
        for (const controlIdx of idxs) {
          const targetEntry = entries[controlIdx];
          if (!targetEntry) continue;
          const cg = Number(targetEntry.controlGroup);
          if (Number.isFinite(cg) && cg > 0) {
            anchorControlGroup = cg;
            break;
          }
        }
        const labelIdx = ensureEntryExists(sourceOp, sourceId, labelDisplayName, {
          kind: 'label',
          labelCount: idxs.length,
          inputControl: 'LabelControl',
          defaultValue: '0',
          page: 'Controls',
          controlGroup: anchorControlGroup,
          syntheticToolUserControl: true,
        });
        if (labelIdx != null) {
          indicesToInsert.unshift(labelIdx);
        }
      }
      if (indicesToInsert.length) {
        const pos = getInsertionPosUnderSelection();
        state.parseResult.order = insertIndicesAt(state.parseResult.order, indicesToInsert, pos);
      }
      renderPublishedList(state.parseResult.entries, state.parseResult.order);
      refreshNodesChecks();
      info(`Quick Publish ${node.name}: added ${added}, already published ${already}${redirected ? `, redirected ${redirected}` : ''}${missing ? `, unavailable ${missing}` : ''}.`);
    } catch (err) {
      error(`Quick publish failed: ${err?.message || err}`);
    }
  }

  function toBlendCheckboxDefault(rawValue) {
    try {
      if (rawValue == null) return '1';
      const text = String(rawValue).trim();
      if (!text) return '1';
      if (/^(true|yes|on)$/i.test(text)) return '1';
      if (/^(false|no|off)$/i.test(text)) return '0';
      const num = Number(text);
      if (Number.isFinite(num)) return num <= 0 ? '0' : '1';
      return '1';
    } catch (_) {
      return '1';
    }
  }

  function shouldForceBlendCheckbox(item) {
    try {
      if (!quickSetBlendToCheckbox) return false;
      if (!item) return false;
      if (isBlendQuickControlId(item.source)) return true;
      return isBlendQuickControl(item.controlDef);
    } catch (_) {
      return false;
    }
  }

  function buildQuickSetPublishMeta(item) {
    const meta = buildControlMetaFromDefinition(item?.controlDef);
    if (!shouldForceBlendCheckbox(item)) return meta;
    const out = { ...(meta || {}) };
    out.inputControl = 'CheckboxControl';
    out.defaultValue = toBlendCheckboxDefault(out.defaultValue);
    return out;
  }

  function buildQuickSetPublishDisplayName(item, node) {
    const fallback = String(item?.displayName || '').trim()
      || humanizeName(item?.source || '')
      || String(item?.source || '').trim()
      || 'Control';
    if (!shouldForceBlendCheckbox(item)) return fallback;
    const nodeName = String(node?.name || '').trim() || String(node?.type || '').trim() || 'Effect';
    return `${nodeName} On/Off`;
  }

  function buildControlMetaFromDefinition(control) {
    if (!control) return null;
    const meta = {};
    let touched = false;
    if (control.kind) { meta.kind = String(control.kind).toLowerCase(); touched = true; }
    if (Number.isFinite(control.labelCount)) { meta.labelCount = Number(control.labelCount); touched = true; }
    if (control.inputControl) { meta.inputControl = control.inputControl; touched = true; }
    if (control.page) { meta.page = String(control.page); touched = true; }
    if (control.defaultValue != null) { meta.defaultValue = control.defaultValue; touched = true; }
    if (Array.isArray(control.choiceOptions) && control.choiceOptions.length) { meta.choiceOptions = [...control.choiceOptions]; touched = true; }
    if (control.multiButtonShowBasic != null && String(control.multiButtonShowBasic).trim() !== '') { meta.multiButtonShowBasic = String(control.multiButtonShowBasic).trim(); touched = true; }
    if (control.defaultX != null) { meta.defaultX = control.defaultX; touched = true; }
    if (control.defaultY != null) { meta.defaultY = control.defaultY; touched = true; }
    if (Number.isFinite(control.controlGroup) && control.controlGroup > 0) { meta.controlGroup = Number(control.controlGroup); touched = true; }
    if (control.isMacroUserControl) { meta.publishTarget = 'groupUserControl'; touched = true; }
    return touched ? meta : null;
  }

  function isGroupedControl(control) {
    return !!(control && (control.type === 'color-group' || control.type === 'linked-group' || control.type === 'range-group'));
  }

  function isSectionHeaderControl(control) {
    return !!(control && (control.type === 'slot-header' || control.type === 'common-header'));
  }

  function shouldDefaultCollapsedSlotHeader(control) {
    try {
      if (!control || !isSectionHeaderControl(control)) return false;
      const slotType = String(control.slotType || '').trim().toLowerCase();
      if (slotType === 'textplus-element') return true;
      if (slotType === 'textplus-shading') return true;
      return false;
    } catch (_) {
      return false;
    }
  }

  function getNodeSelection() {
    try {
      if (!state.parseResult) return new Set();
      if (!(state.parseResult.nodeSelection instanceof Set)) state.parseResult.nodeSelection = new Set();
      return state.parseResult.nodeSelection;
    } catch (_) {
      return new Set();
    }
  }

  function getNodeSlotCollapseState() {
    try {
      if (!state.parseResult) return new Map();
      if (!(state.parseResult.nodeSlotCollapsed instanceof Map)) state.parseResult.nodeSlotCollapsed = new Map();
      return state.parseResult.nodeSlotCollapsed;
    } catch (_) {
      return new Map();
    }
  }

  function clearNodeSelection() {
    if (!state.parseResult) return;
    state.parseResult.nodeSelection = new Set();
    state.parseResult.nodeSelectionMuted = false;
    updateNodeSelectionButtons();
  }

  function updateNodeSelectionButtons() {
    try {} catch (_) {}
  }

  function closeNodeContextMenu() {
    if (nodeContextMenuCleanup) {
      nodeContextMenuCleanup();
      nodeContextMenuCleanup = null;
    }
    if (nodeContextMenu) {
      try { nodeContextMenu.remove(); } catch (_) {}
      nodeContextMenu = null;
    }
  }

  function closeExpressionInspector() {
    if (exprInspectorCleanup) {
      exprInspectorCleanup();
      exprInspectorCleanup = null;
    }
    if (exprInspectorOverlay) {
      try { exprInspectorOverlay.remove(); } catch (_) {}
      exprInspectorOverlay = null;
    }
  }

  function openExpressionInspector(options = {}) {
    try {
      closeExpressionInspector();
      const rawTargets = Array.isArray(options.targets) ? options.targets : [];
      const targets = [];
      const seen = new Set();
      for (const t of rawTargets) {
        if (!t || !t.sourceOp || !t.source) continue;
        const key = `${t.sourceOp}::${t.source}`;
        if (seen.has(key)) continue;
        seen.add(key);
        targets.push({ sourceOp: String(t.sourceOp), source: String(t.source) });
      }

      const overlay = document.createElement('div');
      overlay.className = 'expr-inspector-modal';
      overlay.setAttribute('role', 'dialog');
      overlay.setAttribute('aria-modal', 'true');

      const panel = document.createElement('div');
      panel.className = 'expr-inspector-panel';

      const header = document.createElement('div');
      header.className = 'expr-inspector-header';
      const titleWrap = document.createElement('div');
      const title = document.createElement('h3');
      title.textContent = 'Expression Dependencies';
      const sub = document.createElement('p');
      const op = String(options.sourceOp || '').trim();
      const src = String(options.source || '').trim();
      const displayName = String(options.displayName || '').trim();
      const idText = op && src ? `${op}.${src}` : (displayName || 'Control');
      sub.textContent = `Control: ${displayName || idText}`;
      titleWrap.appendChild(title);
      titleWrap.appendChild(sub);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'expr-inspector-close';
      closeBtn.textContent = 'x';
      closeBtn.setAttribute('aria-label', 'Close');
      header.appendChild(titleWrap);
      header.appendChild(closeBtn);

      const body = document.createElement('div');
      body.className = 'expr-inspector-body';

      const exprLabel = document.createElement('div');
      exprLabel.className = 'expr-inspector-section-label';
      exprLabel.textContent = 'Expression';
      body.appendChild(exprLabel);

      const exprBox = document.createElement('pre');
      exprBox.className = 'expr-inspector-expression';
      exprBox.textContent = String(options.expression || '').trim() || '(empty)';
      body.appendChild(exprBox);

      const refsLabel = document.createElement('div');
      refsLabel.className = 'expr-inspector-section-label';
      refsLabel.textContent = `References (${targets.length})`;
      body.appendChild(refsLabel);

      const refsWrap = document.createElement('div');
      refsWrap.className = 'expr-inspector-refs';
      if (!targets.length) {
        const empty = document.createElement('div');
        empty.className = 'expr-inspector-empty';
        empty.textContent = 'No references detected in this expression.';
        refsWrap.appendChild(empty);
      } else {
        for (const t of targets) {
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'expr-inspector-ref';
          btn.textContent = `${t.sourceOp}.${t.source}`;
          btn.title = 'Jump to referenced control';
          btn.addEventListener('click', () => {
            try {
              clearHighlights();
              highlightCallback(t.sourceOp, t.source || null);
            } catch (_) {}
            closeExpressionInspector();
          });
          refsWrap.appendChild(btn);
        }
      }
      body.appendChild(refsWrap);

      const footer = document.createElement('div');
      footer.className = 'expr-inspector-footer';
      const doneBtn = document.createElement('button');
      doneBtn.type = 'button';
      doneBtn.textContent = 'Close';
      doneBtn.addEventListener('click', closeExpressionInspector);
      footer.appendChild(doneBtn);

      panel.appendChild(header);
      panel.appendChild(body);
      panel.appendChild(footer);
      overlay.appendChild(panel);
      document.body.appendChild(overlay);

      const onKeyDown = (ev) => {
        if (ev.key === 'Escape') {
          ev.preventDefault();
          closeExpressionInspector();
        }
      };
      overlay.addEventListener('click', (ev) => {
        if (ev.target === overlay) closeExpressionInspector();
      });
      closeBtn.addEventListener('click', closeExpressionInspector);
      document.addEventListener('keydown', onKeyDown, true);
      exprInspectorCleanup = () => {
        document.removeEventListener('keydown', onKeyDown, true);
      };
      exprInspectorOverlay = overlay;
    } catch (_) {}
  }

  function openNodeContextMenu(node, x, y) {
    closeNodeContextMenu();
    if (!node || typeof requestRenameNode !== 'function') return;
    const menu = document.createElement('div');
    menu.className = 'doc-tab-menu';
    menu.setAttribute('role', 'menu');
    const addItem = (label, onClick) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'doc-tab-menu-item';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeNodeContextMenu();
        onClick?.();
      });
      menu.appendChild(btn);
    };
    const renameLabel = node?.name ? `Rename "${node.name}"...` : 'Rename node...';
    addItem(renameLabel, () => requestRenameNode(node));
    document.body.appendChild(menu);
    nodeContextMenu = menu;
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
      if (!menu.contains(ev.target)) closeNodeContextMenu();
    };
    const onKeyDown = (ev) => {
      if (ev.key === 'Escape') closeNodeContextMenu();
    };
    const onViewportChange = () => closeNodeContextMenu();
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', onViewportChange);
    window.addEventListener('scroll', onViewportChange, true);
    nodeContextMenuCleanup = () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', onViewportChange);
      window.removeEventListener('scroll', onViewportChange, true);
    };
  }

  nodesSearch?.addEventListener('input', (e) => {
    nodeFilter = (e.target.value || '').toLowerCase();
    if (state.parseResult) parseAndRenderNodes();
  });

  syncViewControlInputs();
  setViewControlsMenuOpen(false);

  const applyViewControlChange = () => {
    try {
      hideReplaced = !!(hideReplacedEl && hideReplacedEl.checked);
      showNodeTypeLabels = !!(showNodeTypeLabelsEl ? showNodeTypeLabelsEl.checked : true);
      showInstancedConnections = !!(showInstancedConnectionsEl ? showInstancedConnectionsEl.checked : true);
      showNextNodeLinks = !!(showNextNodeLinksEl ? showNextNodeLinksEl.checked : true);
      autoGroupQuickSets = !!(autoGroupQuickSetsEl ? autoGroupQuickSetsEl.checked : false);
      nameClickQuickSet = !!(nameClickQuickSetEl ? nameClickQuickSetEl.checked : false);
      quickSetBlendToCheckbox = !!(quickSetBlendToCheckboxEl ? quickSetBlendToCheckboxEl.checked : false);
      saveViewControlOptions();
      syncViewControlInputs();
      if (state.parseResult) parseAndRenderNodes();
    } catch (_) {}
  };

  hideReplacedEl?.addEventListener('change', applyViewControlChange);
  showNodeTypeLabelsEl?.addEventListener('change', applyViewControlChange);
  showInstancedConnectionsEl?.addEventListener('change', applyViewControlChange);
  showNextNodeLinksEl?.addEventListener('change', applyViewControlChange);
  autoGroupQuickSetsEl?.addEventListener('change', applyViewControlChange);
  nameClickQuickSetEl?.addEventListener('change', applyViewControlChange);
  quickSetBlendToCheckboxEl?.addEventListener('change', applyViewControlChange);

  viewControlsBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const willOpen = !!viewControlsMenu?.hidden;
    setViewControlsMenuOpen(willOpen);
  });
  showAllNodesBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    setAllNodeModes('open');
    setViewControlsMenuOpen(false);
  });
  showPublishedNodesBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    setAllNodeModes('published');
    setViewControlsMenuOpen(false);
  });
  collapseAllNodesBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    setAllNodeModes('closed');
    setViewControlsMenuOpen(false);
  });

  if (viewControlsMenu) {
    document.addEventListener('mousedown', (ev) => {
      if (viewControlsMenu.hidden) return;
      if (viewControlsMenu.contains(ev.target) || viewControlsBtn?.contains(ev.target)) return;
      setViewControlsMenuOpen(false);
    });
    document.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Escape') return;
      if (viewControlsMenu.hidden) return;
      setViewControlsMenuOpen(false);
    });
  }


  importCatalogBtn?.addEventListener('click', () => catalogInput?.click());
  catalogInput?.addEventListener('change', handleCatalogImport);

  importModifierCatalogBtn?.addEventListener('click', () => modifierCatalogInput?.click());
  modifierCatalogInput?.addEventListener('change', handleModifierCatalogImport);

  async function handleCatalogImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      let json = JSON.parse(text);
      if (json && typeof json === 'object') {
        if (json.nodes) json = json.nodes;
        else if (json.tools) json = json.tools;
        else if (json.types) json = json.types;
      }
      nodeCatalog = json || null;
      logTag('Catalog', `Loaded node catalog from file: ${Object.keys(nodeCatalog || {}).length} tool types`);
      if (state.parseResult) parseAndRenderNodes();
    } catch (err) {
      error('Catalog parse error: ' + (err.message || err));
    }
  }

  async function handleModifierCatalogImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      let json = JSON.parse(text);
      if (json && typeof json === 'object') {
        if (json.modifiers) json = json.modifiers;
        else if (json.tools) json = json.tools;
        else if (json.types) json = json.types;
      }
      modifierCatalog = json || null;
      logTag('Catalog', `Loaded modifier catalog from file: ${Object.keys(modifierCatalog || {}).length} modifier types`);
      if (state.parseResult) parseAndRenderNodes();
    } catch (err) {
      error('Modifier catalog parse error: ' + (err.message || err));
    }
  }

  function handlePublishSelected() {
    try {
      if (!state.parseResult || !nodesList) return;
      const rows = Array.from(nodesList.querySelectorAll('.node-row.selected'));
      if (!rows.length) return;
      const items = [];
      for (const row of rows) {
        const meta = row && row._mmMeta ? row._mmMeta : null;
        const kind = row.dataset.kind;
        const op = row.dataset.sourceOp;
        if (kind === 'group') {
          const metaChannels = Array.isArray(meta?.channels) ? meta.channels : [];
          for (const ch of metaChannels) {
            const chOp = String(ch?.sourceOp || op || '').trim();
            const chSrc = String(ch?.source || ch?.id || '').trim();
            if (!chOp || !chSrc) continue;
            const chMeta = ch && ch.control ? buildControlMetaFromDefinition(ch.control) : null;
            items.push({ sourceOp: chOp, source: chSrc, name: (ch && ch.name) || chSrc, controlMeta: chMeta });
          }
        } else if (kind === 'control') {
          const id = row.dataset.source;
          const nm = row.dataset.name || id;
          const ctrlKind = row.dataset.controlKind || '';
          const labelCount = row.dataset.labelCount ? parseInt(row.dataset.labelCount, 10) : null;
          const inputControl = row.dataset.inputControl || '';
          const defaultValue = row.dataset.defaultValue || null;
          items.push({
            sourceOp: String(meta?.sourceOp || op || '').trim(),
            source: id,
            name: nm,
            kind: ctrlKind.toLowerCase(),
            labelCount,
            inputControl,
            defaultValue,
          });
        }
      }
      if (!items.length) return;
      const idxs = [];
      for (const it of items) {
        const idx = ensureEntryExists(it.sourceOp, it.source, it.name, {
          kind: it.kind,
          labelCount: it.labelCount,
          inputControl: it.inputControl,
          defaultValue: it.defaultValue,
          controlGroup: it.controlMeta?.controlGroup,
          defaultX: it.controlMeta?.defaultX,
          defaultY: it.controlMeta?.defaultY,
        });
        if (idx != null) idxs.push(idx);
      }
      if (idxs.length) {
        const pos = getInsertionPosUnderSelection();
        state.parseResult.order = insertIndicesAt(state.parseResult.order, idxs, pos);
        try { logDiag(`Batch publish count=${idxs.length} at pos ${pos}`); } catch (_) {}
        renderPublishedList(state.parseResult.entries, state.parseResult.order);
        refreshNodesChecks();
        state.parseResult.nodeSelectionMuted = true;
        rows.forEach((row) => {
          row.classList.remove('selected');
        });
        updateNodeSelectionButtons();
      }
    } catch (e) {
      try { logDiag('Batch publish error: ' + (e.message || e)); } catch (_) {}
    }
  }

  function isNodeDragEvent(ev) {
    try {
      const dt = ev.dataTransfer;
      if (!dt) return false;
      const types = Array.from(dt.types || []);
      if (types.includes('application/x-fmr-node-control')) return true;
      const txt = dt.getData('text/plain');
      return !!(txt && String(txt).startsWith('FMR_NODE:'));
    } catch (_) {
      return false;
    }
  }

  function parseNodeDragData(ev) {
    try {
      const dt = ev.dataTransfer;
      if (!dt) return null;
      let payload = null;
      const raw = dt.getData('application/x-fmr-node-control') || dt.getData('text/plain') || '';
      const s = raw && raw.startsWith('FMR_NODE:') ? raw.slice('FMR_NODE:'.length) : null;
      if (s) payload = JSON.parse(s);
      return payload;
    } catch (_) {
      return null;
    }
  }

  function parseAndRenderNodes() {
    try {
      if (!nodesList) return;
      nodesList.innerHTML = '';
      if (!state.parseResult || !state.originalText) return;
      let grp = findMacroGroupForNodes(
        state.originalText,
        state.parseResult?.macroName,
        state.parseResult?.macroNameOriginal,
        state.parseResult?.operatorType,
        state.parseResult?.operatorTypeOriginal,
      );
      if (!grp && state.parseResult?.inputs?.openIndex != null) {
        grp = findEnclosingGroupForIndex(state.originalText, state.parseResult.inputs.openIndex);
      }
      if (!grp) {
        try { logDiag('Nodes: no group bounds found; falling back to whole file.'); } catch (_) {}
        grp = { name: state.parseResult.macroName || 'Unknown', groupOpenIndex: 0, groupCloseIndex: state.originalText.length };
      }
      const tools = parseToolsInGroup(state.originalText, grp.groupOpenIndex, grp.groupCloseIndex);
      const maskPathDriverOps = collectMaskPathDriverOps(tools);
      maskPathDriverNodeNames = new Set(maskPathDriverOps);
      maskPathDriverNodeNamesLower = new Set(Array.from(maskPathDriverOps).map((name) => String(name || '').trim().toLowerCase()).filter(Boolean));
      const modifiers = parseModifiersInGroup(state.originalText, grp.groupOpenIndex, grp.groupCloseIndex);
      const downstream = buildDownstreamMap(tools);
      const modifierBindings = buildModifierBindings(tools);
      const expressionBindings = buildExpressionBindings(tools);
      modifierBindingContext = modifierBindings;
      modifierContextDiagSeen = new Set();
      let nodes = tools.map(t => {
        const typeStr = String(t.type || '');
        const isMod = isModifierType(typeStr);
        const catalogRef = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
        const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
        const contextBindings = isMod ? (modifierBindings.get(t.name) || []) : [];
        return {
          name: t.name,
          type: t.type,
          controls: deriveControlsForTool(t, catalogRef, profileRef, contextBindings),
          isModifier: isMod,
          external: false,
          instanceSourceOp: t.instanceSourceOp || null,
          isMaskPathDriver: maskPathDriverOps.has(t.name) && /bezierspline/i.test(String(t.type || '')),
        };
      }).concat(modifiers.map(m => ({
        name: m.name,
        type: m.type,
        controls: deriveControlsForTool(
          m,
          (modifierCatalog || nodeCatalog),
          (modifierProfiles || nodeProfiles),
          (modifierBindings.get(m.name) || []),
        ),
        isModifier: true,
        external: false,
        instanceSourceOp: m.instanceSourceOp || null,
        isMaskPathDriver: false,
      })));
      const controlledBy = buildControlledByMap(modifierBindings);
      nodes = nodes.map(n => ({
        ...n,
        downstream: downstream.get(n.name) || [],
        bindings: modifierBindings.get(n.name) || [],
        controlledBy: controlledBy.get(n.name) || new Map(),
        expressionDriven: expressionBindings.get(n.name) || new Map(),
      }));
      augmentWithReferencedNodes(nodes, tools, modifierBindings);
      augmentWithAllTools(nodes, modifierBindings);
      const downstreamAll = buildDownstreamForAll(nodes);
      nodes = nodes.map(n => ({
        ...n,
        downstream: downstreamAll.get(n.name) || [],
        bindings: modifierBindings.get(n.name) || (n.bindings || []),
        controlledBy: (controlledBy.get(n.name) || n.controlledBy || new Map()),
        expressionDriven: (expressionBindings.get(n.name) || n.expressionDriven || new Map()),
        isMaskPathDriver: !!(n.isMaskPathDriver || (maskPathDriverOps.has(n.name) && /bezierspline/i.test(String(n.type || '')))),
      }));
      nodes = nodes.filter((n) => !shouldHideNodeFromList(n));
      lastNodeNames = Array.from(new Set(
        (nodes || [])
          .filter((n) => n && !n.external && n.name)
          .map((n) => String(n.name))
      ));

      const macroControls = buildMacroRootControls(
        state.parseResult,
        state.parseResult?.macroName || grp.name || 'Macro',
      );
      if (macroControls.length) {
        const macroName = state.parseResult?.macroName || grp.name || 'Macro';
        const macroNode = {
          name: macroName,
          type: state.parseResult?.operatorType || 'GroupOperator',
          controls: macroControls,
          isModifier: false,
          external: false,
          isMacroRoot: true,
          downstream: [],
          downstreamAll: [],
          bindings: [],
          controlledBy: new Map(),
          expressionDriven: new Map(),
          instanceSourceOp: null,
        };
        nodes.unshift(macroNode);
      }
      nodes = prioritizeUtilityNode(nodes, UTILITY_NODE_NAME);
      try {
        if (state.parseResult) state.parseResult.nodes = Array.isArray(nodes) ? nodes.slice() : [];
      } catch (_) {}
      const flt = (nodeFilter || '').trim();
      const filtered = !flt ? nodes : nodes.filter(n => {
        if (n.name.toLowerCase().includes(flt) || (n.type || '').toLowerCase().includes(flt)) return true;
        return (n.controls || []).some(c => (c.groupLabel || '').toLowerCase().includes(flt) || (c.name || '').toLowerCase().includes(flt));
      });
      try { if (state.parseResult) state.parseResult.controlMeta = null; } catch (_) {}
      renderNodes(filtered);
      logDiag(`Nodes parsed: ${nodes.length}`);
    } catch (e) {
      logDiag('Nodes parse error: ' + (e.message || e));
    }
  }

  function buildDownstreamMap(tools) {
    const map = new Map();
    const bodies = tools.map(t => ({ name: t.name, body: String(t.body || '') }));
    for (const t of tools) map.set(t.name, new Set());
    for (const b of bodies) {
      const re = /SourceOp\s*=\s*\"([^\"]+)\"/g;
      let m;
      while ((m = re.exec(b.body)) !== null) {
        const src = m[1];
        if (src && map.has(src)) map.get(src).add(b.name);
      }
      const exprRefs = findExpressionRefsInToolBody(b.body) || [];
      for (const ref of exprRefs) {
        for (const target of (ref.targets || [])) {
          const src = target && target.sourceOp ? target.sourceOp : '';
          if (src && map.has(src)) map.get(src).add(b.name);
        }
      }
    }
    const out = new Map();
    for (const [k, v] of map.entries()) out.set(k, Array.from(v));
    return out;
  }

  function buildModifierBindings(tools) {
    const map = new Map();
    for (const t of tools) {
      const refs = findModifierRefsInToolBody(t.body) || [];
      for (const r of refs) {
        if (!map.has(r.modifier)) map.set(r.modifier, []);
        map.get(r.modifier).push({ tool: t.name, toolType: t.type, id: r.id, source: r.source || '' });
      }
    }
    return map;
  }

  function buildExpressionBindings(tools) {
    const map = new Map();
    for (const t of tools) {
      const refs = findExpressionRefsInToolBody(t.body) || [];
      if (!refs.length) continue;
      map.set(t.name, expressionRefsToMap(refs));
    }
    return map;
  }

  function expressionRefsToMap(refs) {
    const map = new Map();
    for (const r of (refs || [])) {
      if (!r || !r.id) continue;
      const key = String(r.id);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
    }
    return map;
  }

  function buildControlledByMap(modifierBindings) {
    const map = new Map();
    for (const [mod, arr] of modifierBindings.entries()) {
      for (const b of arr) {
        if (!map.has(b.tool)) map.set(b.tool, new Map());
        const m = map.get(b.tool);
        if (!m.has(b.id)) m.set(b.id, []);
        m.get(b.id).push(mod);
      }
    }
    return map;
  }

  function augmentWithReferencedNodes(nodes, tools, modifierBindings) {
    try {
      const existing = new Set(nodes.map(n => n.name));
      const refs = new Set((state.parseResult && state.parseResult.entries ? state.parseResult.entries : []).map(e => e && e.sourceOp).filter(Boolean));
      try {
        for (const r of extractReferencedOpsFromTools(tools || [])) refs.add(r);
      } catch (_) {}
      for (const op of refs) {
        if (!existing.has(op)) {
          const t = findToolByNameAnywhere(state.originalText, op);
          if (t) {
            const typeStr = String(t.type || '');
            const isMod = isModifierType(typeStr);
            const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
            const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
            const contextBindings = isMod ? ((modifierBindings && modifierBindings.get(t.name)) || []) : [];
            nodes.push({
              name: t.name,
              type: t.type,
              controls: deriveControlsForTool(t, cat, profileRef, contextBindings),
              isModifier: isMod,
              external: true,
              expressionDriven: expressionRefsToMap(findExpressionRefsInToolBody(t.body) || []),
              instanceSourceOp: t.instanceSourceOp || null,
            });
            existing.add(op);
          }
        }
      }
    } catch (_) {}
  }

  function augmentWithAllTools(nodes, modifierBindings) {
    try {
      const existingNames = new Set(nodes.map(n => n.name));
      const reAll = /(^|\n)\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*\{/g;
      const disallowedTypes = new Set(['groupinfo', 'operatorinfo', 'input', 'instanceinput', 'instanceoutput', 'polyline', 'fuid']);
      let mAll;
      while ((mAll = reAll.exec(state.originalText)) !== null) {
        const name = mAll[2];
        const type = mAll[3];
        if (!name || existingNames.has(name)) continue;
        const typeStr = String(type || '');
        if (disallowedTypes.has(typeStr.toLowerCase())) continue;
        if (/^value$/i.test(String(name || '').trim())) continue;
        const t = findToolByNameAnywhere(state.originalText, name);
        if (!t) continue;
        const isMod = isModifierType(typeStr);
        const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
        const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
        const contextBindings = isMod ? ((modifierBindings && modifierBindings.get(t.name)) || []) : [];
        nodes.push({
          name: t.name,
          type: t.type,
          controls: deriveControlsForTool(t, cat, profileRef, contextBindings),
          isModifier: isMod,
          external: true,
          expressionDriven: expressionRefsToMap(findExpressionRefsInToolBody(t.body) || []),
          instanceSourceOp: t.instanceSourceOp || null,
        });
        existingNames.add(name);
      }
    } catch (_) {}
  }

  function buildDownstreamForAll(nodes) {
    const defs = [];
    for (const n of nodes) {
      try {
        const t = findToolByNameAnywhere(state.originalText, n.name);
        if (t) defs.push({ name: n.name, body: String(t.body || '') });
      } catch (_) {}
    }
    const mapAll = new Map();
    for (const d of defs) mapAll.set(d.name, new Set());
    for (const d of defs) {
      const re = /SourceOp\s*=\s*\"([^\"]+)\"/g;
      let m;
      while ((m = re.exec(d.body)) !== null) {
        const src = m[1];
        if (src && mapAll.has(src)) mapAll.get(src).add(d.name);
      }
      const exprRefs = findExpressionRefsInToolBody(d.body) || [];
      for (const ref of exprRefs) {
        for (const target of (ref.targets || [])) {
          const src = target && target.sourceOp ? target.sourceOp : '';
          if (src && mapAll.has(src)) mapAll.get(src).add(d.name);
        }
      }
    }
    const out = new Map();
    for (const [k, v] of mapAll.entries()) out.set(k, Array.from(v));
    return out;
  }

  function resolveControlSource(control) {
    if (!control) return '';
    const idStr = String(control.id || '');
    if (/^\d+$/.test(idStr) && control.name) {
      try {
        const base = String(control.name).replace(/\s+/g, '');
        return sanitizeIdent(base);
      } catch (_) {
        return idStr;
      }
    }
    return idStr;
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

  function collectMaskPathDriverOps(tools) {
    const out = new Set();
    try {
      for (const tool of (tools || [])) {
        if (!tool || !isMaskToolTypeForPath(tool.type)) continue;
        const entries = parseToolInputEntries(String(tool.body || ''));
        for (const entry of entries) {
          if (!entry || !isPathControlId(entry.id)) continue;
          const body = String(entry.inputBody || '');
          if (!body) continue;
          const sourceOp = String(extractQuotedProp(body, 'SourceOp') || '').trim();
          const source = String(extractQuotedProp(body, 'Source') || '').trim().toLowerCase();
          if (sourceOp && source === 'value') out.add(sourceOp);
        }
      }
    } catch (_) {}
    return out;
  }

  function escapeRegex(text) {
    return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function findMacroGroupByName(text, name, operatorType) {
    try {
      const safeName = String(name || '').trim();
      const safeType = String(operatorType || '').trim();
      if (!safeName || !safeType) return null;
      const re = new RegExp(`(^|\\n|\\r)\\s*${escapeRegex(safeName)}\\s*=\\s*${escapeRegex(safeType)}\\s*\\{`);
      const match = re.exec(text);
      if (!match) return null;
      const openIndex = match.index + match[0].lastIndexOf('{');
      const closeIndex = findMatchingBrace(text, openIndex);
      if (closeIndex < 0) return null;
      return {
        name: safeName,
        groupOpenIndex: openIndex,
        groupCloseIndex: closeIndex,
      };
    } catch (_) {
      return null;
    }
  }

  function findMacroGroupForNodes(text, macroName, macroNameOriginal, operatorType, operatorTypeOriginal) {
    try {
      const names = [];
      const addName = (value) => {
        const next = String(value || '').trim();
        if (!next || names.includes(next)) return;
        names.push(next);
      };
      addName(macroName);
      addName(macroNameOriginal);
      const operatorTypes = [];
      const addType = (value) => {
        const next = String(value || '').trim();
        if (!next || operatorTypes.includes(next)) return;
        operatorTypes.push(next);
      };
      addType(operatorType);
      addType(operatorTypeOriginal);
      addType('GroupOperator');
      addType('MacroOperator');
      for (const name of names) {
        for (const type of operatorTypes) {
          const bounds = findMacroGroupByName(text, name, type);
          if (bounds) return bounds;
        }
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function resolvePublishSourceOp(node, control) {
    try {
      const nodeOp = String(node?.name || '').trim();
      const instanceOp = String(node?.instanceSourceOp || '').trim();
      if (!nodeOp) return instanceOp || '';
      if (!instanceOp) return nodeOp;
      const state = String(control?.instanceState || '').trim().toLowerCase();
      if (state === 'instanced' || state === 'instanced-explicit') return instanceOp;
      return nodeOp;
    } catch (_) {
      return String(node?.name || '').trim();
    }
  }

  function dedupeControlsByResolvedSource(controls) {
    try {
      if (!Array.isArray(controls) || !controls.length) return controls || [];
      const out = [];
      const seen = new Set();
      for (const c of controls) {
        if (!c) continue;
        if (isGroupedControl(c)) {
          const gk = `group:${String(c.base || c.id || '').trim()}`;
          if (!gk || seen.has(gk)) continue;
          seen.add(gk);
          out.push(c);
          continue;
        }
        const src = resolveControlSource(c);
        const ck = `control:${String(src || '').trim()}`;
        if (!ck || seen.has(ck)) continue;
        seen.add(ck);
        out.push(c);
      }
      return out;
    } catch (_) {
      return controls || [];
    }
  }

  function getNodeTypeMeta(node) {
    try {
      if (!node) return null;
      if (node.isMacroRoot) return { key: 'macro-root', label: 'Macro Root' };
      if (node.isModifier) return { key: 'modifier', label: 'Modifier' };
      const type = String(node.type || '').trim();
      const name = String(node.name || '').trim();
      const typeLower = type.toLowerCase();
      const explicitTypeMap = new Map([
        // Matte
        ['alphadivide', { key: 'matte', label: 'Matte' }],
        ['alphamultiply', { key: 'matte', label: 'Matte' }],
        ['chromakeyer', { key: 'matte', label: 'Matte' }],
        ['cleanplate', { key: 'matte', label: 'Matte' }],
        ['cryptomatte', { key: 'matte', label: 'Matte' }],
        ['deltakeyer', { key: 'matte', label: 'Matte' }],
        ['depthmap', { key: 'matte', label: 'Matte' }],
        ['differencekeyer', { key: 'matte', label: 'Matte' }],
        ['lumakeyer', { key: 'matte', label: 'Matte' }],
        ['magicmask', { key: 'matte', label: 'Matte' }],
        ['mattecontrol', { key: 'matte', label: 'Matte' }],
        ['relight', { key: 'matte', label: 'Matte' }],
        ['ultrakeyer', { key: 'matte', label: 'Matte' }],
        // Tracking
        ['cameratracker', { key: 'tracking', label: 'Tracking' }],
        ['dimension.cameratracker', { key: 'tracking', label: 'Tracking' }],
        ['planartracker', { key: 'tracking', label: 'Tracking' }],
        ['dimension.planartracker', { key: 'tracking', label: 'Tracking' }],
        ['surfacetracker', { key: 'tracking', label: 'Tracking' }],
        ['tracker', { key: 'tracking', label: 'Tracking' }],
        // Warp
        ['coordinatespace', { key: 'warp', label: 'Warp' }],
        ['coordspace', { key: 'warp', label: 'Warp' }],
        ['cornerpositioner', { key: 'warp', label: 'Warp' }],
        ['displace', { key: 'warp', label: 'Warp' }],
        ['gridwarp', { key: 'warp', label: 'Warp' }],
        ['lensdistort', { key: 'warp', label: 'Warp' }],
        ['perspectivepositioner', { key: 'warp', label: 'Warp' }],
        ['vectordistortion', { key: 'warp', label: 'Warp' }],
        ['vectorwarp', { key: 'warp', label: 'Warp' }],
        ['vortex', { key: 'warp', label: 'Warp' }],
        // VR
        ['immersivepatcher', { key: 'vr', label: 'VR' }],
        ['latlongpatcher', { key: 'vr', label: 'VR' }],
        ['panomap', { key: 'vr', label: 'VR' }],
        ['sphericalstabilizer', { key: 'vr', label: 'VR' }],
        // Stereo
        ['anaglyph', { key: 'stereo', label: 'Stereo' }],
        ['combiner', { key: 'stereo', label: 'Stereo' }],
        ['disparity', { key: 'stereo', label: 'Stereo' }],
        ['dimension.disparity', { key: 'stereo', label: 'Stereo' }],
        ['disparitytoz', { key: 'stereo', label: 'Stereo' }],
        ['dimension.disparitytoz', { key: 'stereo', label: 'Stereo' }],
        ['globalalign', { key: 'stereo', label: 'Stereo' }],
        ['dimension.globalalign', { key: 'stereo', label: 'Stereo' }],
        ['neweye', { key: 'stereo', label: 'Stereo' }],
        ['dimension.neweye', { key: 'stereo', label: 'Stereo' }],
        ['splitter', { key: 'stereo', label: 'Stereo' }],
        ['stereoalign', { key: 'stereo', label: 'Stereo' }],
        ['dimension.stereoalign', { key: 'stereo', label: 'Stereo' }],
        ['ztodisparity', { key: 'stereo', label: 'Stereo' }],
        ['dimension.ztodisparity', { key: 'stereo', label: 'Stereo' }],
        // Layer
        ['layermuxer', { key: 'layer', label: 'Layer' }],
        ['layerregex', { key: 'layer', label: 'Layer' }],
        ['layerremover', { key: 'layer', label: 'Layer' }],
        ['swizzler', { key: 'layer', label: 'Layer' }],
        // I/O
        ['loader', { key: 'io', label: 'I/O' }],
        ['mediain', { key: 'io', label: 'I/O' }],
        ['mediaout', { key: 'io', label: 'I/O' }],
        ['saver', { key: 'io', label: 'I/O' }],
        // Paint
        ['paint', { key: 'paint', label: 'Paint' }],
        // Film
        ['cineonlog', { key: 'film', label: 'Film' }],
        ['filmgrain', { key: 'film', label: 'Film' }],
        ['grain', { key: 'film', label: 'Film' }],
        ['lighttrim', { key: 'film', label: 'Film' }],
        ['removenoise', { key: 'film', label: 'Film' }],
        // Color
        ['acestransform', { key: 'color', label: 'Color' }],
        ['autogain', { key: 'color', label: 'Color' }],
        ['brightnesscontrast', { key: 'color', label: 'Color' }],
        ['channelboolean', { key: 'color', label: 'Color' }],
        ['channelbooleans', { key: 'color', label: 'Color' }],
        ['chromaticaberrationremoval', { key: 'color', label: 'Color' }],
        ['chromaticadaptation', { key: 'color', label: 'Color' }],
        ['colorspacetransform', { key: 'color', label: 'Color' }],
        ['filmlookcreator', { key: 'color', label: 'Color' }],
        ['filelut', { key: 'color', label: 'Color' }],
        ['colorcorrector', { key: 'color', label: 'Color' }],
        ['colorcurves', { key: 'color', label: 'Color' }],
        ['colorgain', { key: 'color', label: 'Color' }],
        ['colormatrix', { key: 'color', label: 'Color' }],
        ['customcolormatrix', { key: 'color', label: 'Color' }],
        ['colorspace', { key: 'color', label: 'Color' }],
        ['copyaux', { key: 'color', label: 'Color' }],
        ['gamut', { key: 'color', label: 'Color' }],
        ['gamutconvert', { key: 'color', label: 'Color' }],
        ['gamutlimiter', { key: 'color', label: 'Color' }],
        ['gamutmapping', { key: 'color', label: 'Color' }],
        ['huecurves', { key: 'color', label: 'Color' }],
        ['ociocdltransform', { key: 'color', label: 'Color' }],
        ['ociocolorspace', { key: 'color', label: 'Color' }],
        ['ociodisplay', { key: 'color', label: 'Color' }],
        ['ociofiletransform', { key: 'color', label: 'Color' }],
        ['setcanvascolor', { key: 'color', label: 'Color' }],
        ['tintensity', { key: 'color', label: 'Color' }],
        // Deep
        ['dcrop', { key: 'deep', label: 'Deep' }],
        ['deeptoimage', { key: 'deep', label: 'Deep' }],
        ['deeptopoints', { key: 'deep', label: 'Deep' }],
        ['defocus', { key: 'deep', label: 'Deep' }],
        ['dholdout', { key: 'deep', label: 'Deep' }],
        ['dmerge', { key: 'deep', label: 'Deep' }],
        ['drecolor', { key: 'deep', label: 'Deep' }],
        ['dresize', { key: 'deep', label: 'Deep' }],
        ['dtransform', { key: 'deep', label: 'Deep' }],
        ['fog', { key: 'deep', label: 'Deep' }],
        ['imagetodeep', { key: 'deep', label: 'Deep' }],
        ['shader', { key: 'deep', label: 'Deep' }],
        ['ssao', { key: 'deep', label: 'Deep' }],
        ['texture', { key: 'deep', label: 'Deep' }],
        // Miscellaneous
        ['autodomain', { key: 'misc', label: 'Misc' }],
        ['changedepth', { key: 'misc', label: 'Misc' }],
        ['custom', { key: 'misc', label: 'Misc' }],
        ['fields', { key: 'misc', label: 'Misc' }],
        ['frameaverage', { key: 'misc', label: 'Misc' }],
        ['keyframestretcher', { key: 'misc', label: 'Misc' }],
        ['runcommand', { key: 'misc', label: 'Misc' }],
        ['setdomain', { key: 'misc', label: 'Misc' }],
        ['timespeed', { key: 'misc', label: 'Misc' }],
        ['timestretcher', { key: 'misc', label: 'Misc' }],
        // Filter
        ['createbumpmap', { key: 'filter', label: 'Filter' }],
        ['customfilter', { key: 'filter', label: 'Filter' }],
        ['erodedilate', { key: 'filter', label: 'Filter' }],
        ['filter', { key: 'filter', label: 'Filter' }],
        ['rankfilter', { key: 'filter', label: 'Filter' }],
        // Transform
        ['camerashake', { key: 'transform', label: 'Transform' }],
        ['crop', { key: 'transform', label: 'Transform' }],
        ['dve', { key: 'transform', label: 'Transform' }],
        ['letterbox', { key: 'transform', label: 'Transform' }],
        ['planartransform', { key: 'transform', label: 'Transform' }],
        ['resize', { key: 'transform', label: 'Transform' }],
        ['scale', { key: 'transform', label: 'Transform' }],
        ['transform', { key: 'transform', label: 'Transform' }],
        // Effect
        ['duplicate', { key: 'effect', label: 'Effect' }],
        ['highlight', { key: 'effect', label: 'Effect' }],
        ['hotspot', { key: 'effect', label: 'Effect' }],
        ['objectremoval', { key: 'effect', label: 'Effect' }],
        ['pseudocolor', { key: 'effect', label: 'Effect' }],
        ['rays', { key: 'effect', label: 'Effect' }],
        ['shadow', { key: 'effect', label: 'Effect' }],
        ['trails', { key: 'effect', label: 'Effect' }],
        // Explicit one-offs from prior tests
        ['daysky', { key: 'generator', label: 'Generator' }],
        ['mandel', { key: 'generator', label: 'Generator' }],
        ['dent', { key: 'warp', label: 'Warp' }],
        ['distort', { key: 'effect', label: 'Effect' }],
        ['drip', { key: 'warp', label: 'Warp' }],
        ['plasma', { key: 'generator', label: 'Generator' }],
        ['bumpmap', { key: 'three-d', label: '3D' }],
        ['cubemap', { key: 'three-d', label: '3D' }],
        ['falloffoperator', { key: 'three-d', label: '3D' }],
        ['spheremap', { key: 'three-d', label: '3D' }],
        ['texcatcher', { key: 'three-d', label: '3D' }],
        ['rangesmask', { key: 'mask', label: 'Mask' }],
        ['trianglemask', { key: 'mask', label: 'Mask' }],
        ['tv', { key: 'effect', label: 'Effect' }],
        ['volumefog', { key: 'position', label: 'Position' }],
        ['volumemask', { key: 'position', label: 'Position' }],
        ['wandmask', { key: 'mask', label: 'Mask' }],
        ['whitebalance', { key: 'color', label: 'Color' }],
        ['ztoworldpos', { key: 'position', label: 'Position' }],
      ]);
      if (explicitTypeMap.has(typeLower)) return explicitTypeMap.get(typeLower);
      const shapeSystemTypes = new Set([
        'sboolean',
        'sbspline',
        'schangestyle',
        'sduplicate',
        'sellipse',
        'sexpand',
        'sgrid',
        'sjitter',
        'smerge',
        'sngon',
        'soutline',
        'spolygon',
        'srectangle',
        'srender',
        'sstar',
        'stext',
        'stransform',
        'sbitmap',
        'sshape',
      ]);
      const blob = `${type} ${name}`.toLowerCase();
      const has = (re) => re.test(blob);
      const hasType = (re) => re.test(typeLower);

      if (has(/\b3d\b|camera3d|renderer3d|light3d|merge3d|transform3d|material|imageplane3d|shape3d/)) {
        return { key: 'three-d', label: '3D' };
      }
      if (hasType(/^u[a-z0-9]/)) {
        return { key: 'three-d', label: '3D' };
      }
      if (hasType(/^(?:light(?:ambient|directional|dome|point|spot|trim)|mtl|dimension\.)/)) {
        return { key: 'three-d', label: '3D' };
      }
      const looksLikeParticlePrefix = /^p[A-Z]/.test(type);
      if (hasType(/^p(?:emitter|render|turbulence|bounce|directionalforce|drag|friction|imageemitter|kill|line|pointforce|randomforce|sprite|stylize|vortex|wind)\b/) || has(/\bparticle\b/)) {
        return { key: 'particle', label: 'Particle' };
      }
      if (looksLikeParticlePrefix) {
        return { key: 'particle', label: 'Particle' };
      }
      if (hasType(/^ofx\./)) {
        return { key: 'ofx', label: 'OFX' };
      }
      if (hasType(/^fuse[._]/)) {
        return { key: 'fuse', label: 'Fuse' };
      }
      const looksLikeShapeSystemPrefix = /^s[A-Z]/.test(type);
      if (shapeSystemTypes.has(typeLower) || looksLikeShapeSystemPrefix) {
        return { key: 'shape-system', label: 'Shape' };
      }
      if (hasType(/^(?:multipoly|polymask|maskpaint|bitmapmask|paintmask|rectangle(?:mask)?|ellipse(?:mask)?|polygon(?:mask)?|polyline(?:mask)?|bspline(?:mask)?|spline(?:mask)?|polypath|mask|outline)\b/)) {
        return { key: 'mask', label: 'Mask' };
      }
      if (has(/multitext|loader|mediain|background|textplus|text3d|text|generator|solid|checker|gradient|fastnoise|noise/)) {
        return { key: 'generator', label: 'Generator' };
      }
      if (hasType(/^(?:multimerge|merge|dissolve|switch|wireless)\b/)) {
        return { key: 'flow', label: 'Flow' };
      }
      if (has(/transform|blur|glow|color|correct|resize|crop|key|tracker|optical|lens|vignette|channel|levels|brightness|contrast|sharpen|erodedilate|displace|letterbox/)) {
        return { key: 'effect', label: 'Effect' };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function shouldHideNodeFromList(node) {
    try {
      if (!node || node.isMacroRoot) return false;
      const nodeNameRaw = String(node.name || '').trim();
      const nodeNameLower = nodeNameRaw.toLowerCase();
      if (nodeNameRaw && (maskPathDriverNodeNames.has(nodeNameRaw) || maskPathDriverNodeNamesLower.has(nodeNameLower))) return true;
      if (node.isMaskPathDriver) return true;
      const type = String(node.type || '').trim().toLowerCase();
      if (/bezierspline/.test(type)) return true;
      if (type === 'bezierspline' && /(?:polyline|polygon|bspline|spline|path)/i.test(nodeNameRaw)) return true;
      const name = String(node.name || '').trim().toLowerCase();
      if (type === 'groupoperator' || type === 'macrooperator') return true;
      const blob = `${type} ${name}`;
      return /\bpiperouter\b|\baudiodisplay\b/.test(blob);
    } catch (_) {
      return false;
    }
  }

  function renderNodes(nodes, preserveNodeName = '') {
    if (!nodesList) return;
    const scrollHost = nodesList.parentElement || nodesList;
    const previousScrollTop = Number.isFinite(scrollHost.scrollTop) ? scrollHost.scrollTop : 0;
    const escapedNodeName = preserveNodeName
      ? ((typeof CSS !== 'undefined' && CSS && typeof CSS.escape === 'function')
          ? CSS.escape(String(preserveNodeName))
          : String(preserveNodeName).replace(/\\/g, '\\\\').replace(/"/g, '\\"'))
      : '';
    let anchorOffsetBefore = null;
    if (preserveNodeName) {
      try {
        const anchorNode = nodesList.querySelector(`.node[data-op="${escapedNodeName}"] .node-header`);
        if (anchorNode) {
          anchorOffsetBefore = anchorNode.getBoundingClientRect().top;
        }
      } catch (_) {}
    }
    nodesList.innerHTML = '';
    if (!state.parseResult) state.parseResult = {};
    if (!(state.parseResult.nodesCollapsed instanceof Set)) state.parseResult.nodesCollapsed = new Set();
    if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
    if (!state.parseResult.nodesViewInitialized) {
      state.parseResult.nodesCollapsed.clear();
      state.parseResult.nodesPublishedOnly.clear();
      nodes.forEach(n => {
        const hasPublished = (n.controls || []).some(c => {
          if (isGroupedControl(c)) {
            return (c.channels || []).some((ch) => {
              const src = resolveControlSource(ch);
              return !!src && isPublished(resolvePublishSourceOp(n, ch), src);
            });
          }
          const src = resolveControlSource(c);
          return src ? isPublished(resolvePublishSourceOp(n, c), src) : false;
        });
        if (hasPublished) state.parseResult.nodesPublishedOnly.add(n.name);
        else state.parseResult.nodesCollapsed.add(n.name);
      });
      state.parseResult.nodesViewInitialized = true;
    }
    const collapsedNodes = state.parseResult.nodesCollapsed;
    const publishedOnlyNodes = state.parseResult.nodesPublishedOnly;
    const getNodeMode = (name) => {
      if (collapsedNodes.has(name)) return 'closed';
      if (publishedOnlyNodes.has(name)) return 'published';
      return 'open';
    };
    const setNodeMode = (name, mode) => {
      collapsedNodes.delete(name);
      publishedOnlyNodes.delete(name);
      if (mode === 'closed') collapsedNodes.add(name);
      else if (mode === 'published') publishedOnlyNodes.add(name);
      state.parseResult.nodesCollapsed = collapsedNodes;
      state.parseResult.nodesPublishedOnly = publishedOnlyNodes;
    };
    let selectIndex = 0;
    for (const n of nodes) {
      const typeMeta = getNodeTypeMeta(n);
      const wrapper = document.createElement('div');
      wrapper.className = 'node';
      if (typeMeta && typeMeta.key) {
        wrapper.classList.add('node-has-type', `node-type-${typeMeta.key}`);
      }
      wrapper.dataset.op = n.name;
      const header = document.createElement('div');
      header.className = 'node-header';
      const canRenameNode = !!requestRenameNode && !n.isMacroRoot && !n.external;
      if (canRenameNode) {
        header.addEventListener('contextmenu', (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          openNodeContextMenu(n, ev.clientX, ev.clientY);
        });
      }
      const nodeTwisty = document.createElement('span');
      const title = document.createElement('div');
      title.className = 'node-title';
      if (n.name && n.name.toUpperCase() === UTILITY_NODE_NAME) {
        title.textContent = n.name;
      } else if (n.isMacroRoot) {
        title.textContent = n.name || 'Macro';
      } else {
        title.textContent = `${n.name} (${n.type || 'Unknown'})`;
      }
      if (showNodeTypeLabels && typeMeta && typeMeta.label && !n.isMacroRoot) {
        const typeChip = document.createElement('span');
        typeChip.className = 'node-type-chip';
        typeChip.textContent = typeMeta.label;
        title.appendChild(typeChip);
      }
      const flags = document.createElement('div');
      flags.className = 'node-flags';
      (function () {
        if (n.isMacroRoot) {
          const badge = document.createElement('span');
          badge.className = 'node-badge macro-root';
          badge.textContent = 'Macro Root';
          flags.appendChild(badge);
        }
      })();
      const nodeControlsRaw = ensureDissolveMixControl(n, n.controls || []);
      const nodeControls = dedupeControlsByResolvedSource(nodeControlsRaw);
      const nodeControlsDisplay = prioritizeExpandedNodeControls(nodeControls, n);
      const slotCollapseState = getNodeSlotCollapseState();
      let activeSectionHidden = false;
      let activeSectionKey = '';
      const hasAnyPublished = (nodeControls || []).some(c => {
        if (isGroupedControl(c)) {
          return (c.channels || []).some((ch) => {
            const src = resolveControlSource(ch);
            return !!src && isPublished(resolvePublishSourceOp(n, ch), src);
          });
        }
        const src = resolveControlSource(c);
        return src ? isPublished(resolvePublishSourceOp(n, c), src) : false;
      });
      const allowSourceControl = /switch/i.test(String(n.type || n.name || ''));
      if (!hasAnyPublished) publishedOnlyNodes.delete(n.name);
      const mode = getNodeMode(n.name);
      const isCollapsed = mode === 'closed';
      const showPublishedOnly = mode === 'published' && hasAnyPublished;
      nodeTwisty.className = 'twisty' + (showPublishedOnly ? ' published-only' : '');
      nodeTwisty.title = showPublishedOnly ? 'Showing published controls only. Click to collapse.' : (isCollapsed ? 'Expand node' : 'Show published controls only');
      nodeTwisty.innerHTML = isCollapsed ? createIcon('chevron-right') : createIcon('chevron-down');
      nodeTwisty.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const current = getNodeMode(n.name);
        let next = null;
        if (!hasAnyPublished) {
          next = current === 'closed' ? 'open' : 'closed';
        } else {
          next = current === 'closed' ? 'published' : current === 'published' ? 'open' : 'closed';
        }
        setNodeMode(n.name, next);
        renderNodes(nodes, n.name);
      });
      header.appendChild(nodeTwisty);
      header.appendChild(title);
      const metaLinksWrap = document.createElement('div');
      metaLinksWrap.className = 'node-meta-links';
      let hasMetaLinks = false;
      let hasNextNodeLink = false;
      let hasInstanceLink = false;
      const appendMetaLink = (el) => {
        if (!el) return;
        metaLinksWrap.appendChild(el);
        hasMetaLinks = true;
      };
      let nextNodeWrap = null;
      if (showInstancedConnections && n.instanceSourceOp) {
        const instanceWrap = document.createElement('div');
        instanceWrap.className = 'node-next node-instance-link';
        const sep = document.createElement('span'); sep.className = 'sep'; sep.innerHTML = createIcon('chevron-right');
        const prefix = document.createElement('span');
        prefix.className = 'node-instance-prefix';
        prefix.textContent = 'Instanced From';
        const link = document.createElement('a');
        link.href = '#';
        link.className = 'deep-link';
        link.textContent = n.instanceSourceOp;
        link.title = 'Go to source instance node';
        link.addEventListener('click', (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          try { clearHighlights(); highlightCallback(n.instanceSourceOp, null); } catch (_) {}
        });
        instanceWrap.appendChild(sep);
        instanceWrap.appendChild(prefix);
        instanceWrap.appendChild(link);
        hasInstanceLink = true;
        appendMetaLink(instanceWrap);
      }
      if (showNextNodeLinks && Array.isArray(n.bindings) && n.bindings.length > 0) {
        const b = n.bindings[0];
        nextNodeWrap = document.createElement('div');
        nextNodeWrap.className = 'node-next';
        const sep = document.createElement('span'); sep.className = 'sep'; sep.innerHTML = createIcon('chevron-right');
        const link = document.createElement('a');
        link.href = '#'; link.className = 'deep-link';
        link.textContent = `${b.tool}.${b.id}`;
        link.title = 'Go to controlled parameter';
        link.addEventListener('click', (ev) => {
          ev.preventDefault(); ev.stopPropagation();
          try { clearHighlights(); highlightCallback(b.tool, b.id); } catch (_) {}
        });
        nextNodeWrap.appendChild(sep);
        nextNodeWrap.appendChild(link);
        hasNextNodeLink = true;
      } else if (showNextNodeLinks && Array.isArray(n.downstream) && n.downstream.length > 0) {
        nextNodeWrap = document.createElement('div');
        nextNodeWrap.className = 'node-next';
        const sep = document.createElement('span'); sep.className = 'sep'; sep.innerHTML = createIcon('chevron-right');
        const link = document.createElement('a');
        link.href = '#'; link.className = 'deep-link';
        const nextName = n.downstream[0];
        link.textContent = nextName;
        link.title = 'Go to downstream node';
        link.addEventListener('click', (ev) => {
          ev.preventDefault(); ev.stopPropagation();
          try { clearHighlights(); highlightCallback(nextName, null); } catch (_) {}
        });
        nextNodeWrap.appendChild(sep);
        nextNodeWrap.appendChild(link);
        hasNextNodeLink = true;
      }
      if (hasNextNodeLink && nextNodeWrap) {
        if (hasInstanceLink && metaLinksWrap.firstChild) {
          metaLinksWrap.insertBefore(nextNodeWrap, metaLinksWrap.firstChild);
          hasMetaLinks = true;
        } else {
          appendMetaLink(nextNodeWrap);
        }
      }
      if (hasMetaLinks) {
        if (hasNextNodeLink && hasInstanceLink) {
          metaLinksWrap.classList.add('stacked');
        }
        header.appendChild(metaLinksWrap);
      }
      const headerTail = document.createElement('div');
      headerTail.className = 'node-header-tail';
      const quickCandidates = buildQuickSetCandidates(nodeControls, allowSourceControl, n);
      const savedQuickSet = getQuickSetForType(n.type, quickCandidates.map);
      const effectiveQuickSetCount = (savedQuickSet.length ? savedQuickSet : buildSuggestedQuickSetKeys(quickCandidates.list, n)).length;
      const quickPublishBtn = document.createElement('button');
      quickPublishBtn.type = 'button';
      quickPublishBtn.className = 'node-quick-publish-btn';
      quickPublishBtn.textContent = 'Quick';
      quickPublishBtn.title = `Quick Publish (${effectiveQuickSetCount})`;
      quickPublishBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        applyQuickSetToNode(n, nodeControls, allowSourceControl);
      });
      if (effectiveQuickSetCount > 0) {
        header.title = 'Ctrl/Cmd+Click to Quick Publish';
        header.addEventListener('click', (ev) => {
          if (!ev || ev.button !== 0) return;
          if (!(ev.ctrlKey || ev.metaKey)) return;
          ev.preventDefault();
          ev.stopPropagation();
          applyQuickSetToNode(n, nodeControls, allowSourceControl);
        });
        if (nameClickQuickSet) {
          title.classList.add('node-title-quick');
          title.title = 'Click to Quick Publish';
          title.addEventListener('click', (ev) => {
            if (!ev || ev.button !== 0) return;
            ev.preventDefault();
            ev.stopPropagation();
            applyQuickSetToNode(n, nodeControls, allowSourceControl);
          });
        }
      }
      const quickEditBtn = document.createElement('button');
      quickEditBtn.type = 'button';
      quickEditBtn.className = 'node-quick-edit-btn';
      quickEditBtn.textContent = 'Set';
      quickEditBtn.title = `Edit Quick Set (${effectiveQuickSetCount})`;
      quickEditBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        openQuickSetModal(n, nodeControls, allowSourceControl);
      });
      if (typeof requestAddControl === 'function') {
        const addBtn = document.createElement('button');
        addBtn.type = 'button';
        addBtn.className = 'node-add-control-btn';
        addBtn.innerHTML = createIcon('plus') || '+';
        addBtn.title = `Add control to ${n.name}`;
        addBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          requestAddControl(n.name);
        });
        headerTail.appendChild(addBtn);
      }
      headerTail.appendChild(flags);
      header.appendChild(headerTail);
      wrapper.appendChild(header);
      const list = document.createElement('div');
      list.className = 'node-controls';
      list.style.display = isCollapsed ? 'none' : '';
      const quickRow = document.createElement('div');
      quickRow.className = 'node-quick-row';
      quickRow.appendChild(quickPublishBtn);
      quickRow.appendChild(quickEditBtn);
      list.appendChild(quickRow);
      const getControlMods = (controlId) => {
        try {
          const m = n && n.controlledBy;
          if (!m) return [];
          if (typeof m.get === 'function') return Array.isArray(m.get(controlId)) ? m.get(controlId) : [];
          return Array.isArray(m[controlId]) ? m[controlId] : [];
        } catch (_) {
          return [];
        }
      };
      const getControlExprRefs = (controlId) => {
        try {
          const exprMap = n && n.expressionDriven;
          if (!exprMap) return [];
          if (typeof exprMap.get === 'function') return Array.isArray(exprMap.get(controlId)) ? exprMap.get(controlId) : [];
          return Array.isArray(exprMap[controlId]) ? exprMap[controlId] : [];
        } catch (_) {
          return [];
        }
      };
      const filterDisplayMods = (controlId, mods) => {
        try {
          const list = Array.isArray(mods) ? mods : [];
          if (!list.length) return list;
          if (!isMaskToolTypeForPath(n?.type) || !isPathControlId(controlId)) return list;
          return list.filter((modName) => {
            const raw = String(modName || '').trim();
            if (!raw) return false;
            const lower = raw.toLowerCase();
            if (maskPathDriverNodeNames.has(raw)) return false;
            if (maskPathDriverNodeNamesLower.has(lower)) return false;
            return true;
          });
        } catch (_) {
          return Array.isArray(mods) ? mods : [];
        }
      };
      const isControlled = (id) => {
        try {
          const modDriven = filterDisplayMods(id, getControlMods(id)).length > 0;
          const exprDriven = getControlExprRefs(id).length > 0;
          return modDriven || exprDriven;
        } catch (_) { return false; }
      };
        for (const c of (nodeControlsDisplay || [])) {
          if (c.id === 'SourceOp') continue;
          if (c.id === 'Source' && !allowSourceControl) continue;
          if (isSectionHeaderControl(c)) {
            activeSectionKey = `${n.name}::${String(c.groupKey || c.id || '').trim()}`;
            if (!slotCollapseState.has(activeSectionKey) && shouldDefaultCollapsedSlotHeader(c)) {
              slotCollapseState.set(activeSectionKey, true);
            }
            activeSectionHidden = !!slotCollapseState.get(activeSectionKey);
            const sectionKey = activeSectionKey;
            const row = document.createElement('div');
            row.className = 'node-row node-slot-header';
            row.dataset.slotKey = sectionKey;
            const toggle = document.createElement('span');
            toggle.className = 'node-slot-toggle';
            toggle.textContent = activeSectionHidden ? '▸' : '▾';
            const label = document.createElement('span');
            label.className = 'ctrl-name';
            label.textContent = c.groupLabel || c.name || humanizeName(c.id || 'Slot');
            row.appendChild(toggle);
            row.appendChild(label);
            row.addEventListener('click', () => {
              const next = !slotCollapseState.get(sectionKey);
              slotCollapseState.set(sectionKey, next);
              parseAndRenderNodes();
            });
            list.appendChild(row);
            continue;
          }
          if (activeSectionHidden) continue;
          if (isGroupedControl(c)) {
          if (hideReplaced) {
            const anyControlled = (c.channels || []).some(ch => isControlled(ch.id));
            if (anyControlled) continue;
          }
          if (showPublishedOnly) {
            const allPublished = (c.channels || []).length > 0 && (c.channels || []).every((ch) => {
              const src = resolveControlSource(ch);
              return !!src && isPublished(resolvePublishSourceOp(n, ch), src);
            });
            if (!allPublished) continue;
          }
          const row = document.createElement('div'); row.className = 'node-row';
          row.dataset.selectIndex = String(selectIndex++);
          row.dataset.kind = 'group';
          row.dataset.sourceOp = n.name;
          const groupId = c.base || c.id || '';
          row.dataset.groupId = groupId;
          row.dataset.channels = (c.channels || []).map(ch => ch.id).join('|');
          row._mmMeta = {
            kind: 'group',
            sourceOp: n.name,
            groupBase: groupId,
            channels: (c.channels || []).map((ch) => {
              const source = resolveControlSource(ch);
              return {
                sourceOp: resolvePublishSourceOp(n, ch),
                source,
                id: source,
                name: ch.name || ch.id,
                control: ch.control || null,
              };
            }),
          };
          if (enableNodeDrag) {
            row.draggable = true;
            row.addEventListener('dragstart', (ev) => {
              try {
                const payload = {
                  kind: 'group',
                  sourceOp: n.name,
                  base: groupId,
                  channels: (row._mmMeta?.channels || []).map((ch) => ({
                    sourceOp: ch?.sourceOp || n.name,
                    source: ch?.source || ch?.id || '',
                    id: ch?.id || ch?.source || '',
                    name: ch?.name || ch?.id || '',
                  })),
                };
                const txt = 'FMR_NODE:' + JSON.stringify(payload);
                ev.dataTransfer && ev.dataTransfer.setData('text/plain', txt);
                ev.dataTransfer && (ev.dataTransfer.effectAllowed = 'copy');
              } catch (_) {}
            });
          }
          const indicator = document.createElement('span'); indicator.className = 'node-published-indicator group'; indicator.title = 'Published status';
          indicator.dataset.sourceOp = n.name;
          indicator.dataset.groupId = groupId;
          indicator.dataset.channels = (row._mmMeta.channels || []).map(ch => ch.source || ch.id).join('|');
          indicator.addEventListener('click', (ev) => {
            ev.preventDefault();
            ev.stopPropagation();
            if (yieldToPickSession(ev)) return;
            if (typeof focusPublishedControl !== 'function') return;
            const channelMeta = row._mmMeta && Array.isArray(row._mmMeta.channels) ? row._mmMeta.channels : [];
            for (const ch of channelMeta) {
              const chOp = String(ch?.sourceOp || n.name);
              const chSrc = String(ch?.source || ch?.id || '').trim();
              if (chSrc && isPublished(chOp, chSrc)) {
                try { clearHighlights(); } catch (_) {}
                try { focusPublishedControl(chOp, chSrc); } catch (_) {}
                break;
              }
            }
          });
          const publishGroup = () => {
            lastSelectIndex = parseInt(row.dataset.selectIndex || '-1', 10);
            const channelMeta = row._mmMeta && Array.isArray(row._mmMeta.channels) ? row._mmMeta.channels : [];
            const allIdxs = [];
            let firstPublishedTarget = null;
            for (const ch of channelMeta) {
              const chOp = String(ch?.sourceOp || n.name);
              const chSrc = String(ch?.source || ch?.id || '').trim();
              if (!chSrc || isPublished(chOp, chSrc)) continue;
              const meta = ch && ch.control ? buildControlMetaFromDefinition(ch.control) : null;
              const r = ensurePublished(chOp, chSrc, ch.name, meta);
              if (r) allIdxs.push(r.index);
            }
            if (allIdxs.length) {
              const pos = getInsertionPosUnderSelection();
              try { logDiag(`Insert group count=${allIdxs.length} at pos ${pos}`); } catch (_) {}
              state.parseResult.order = insertIndicesAt(state.parseResult.order, allIdxs, pos);
            } else {
              for (const ch of channelMeta) {
                const chOp = String(ch?.sourceOp || n.name);
                const chSrc = String(ch?.source || ch?.id || '').trim();
                if (chSrc && isPublished(chOp, chSrc)) {
                  firstPublishedTarget = { sourceOp: chOp, source: chSrc };
                  break;
                }
              }
            }
            renderPublishedList(state.parseResult.entries, state.parseResult.order);
            refreshNodesChecks();
            if (!allIdxs.length && firstPublishedTarget && typeof focusPublishedControl === 'function') {
              try { clearHighlights(); } catch (_) {}
              try { focusPublishedControl(firstPublishedTarget.sourceOp, firstPublishedTarget.source); } catch (_) {}
            }
          };
          const label = document.createElement('span'); label.className = 'ctrl-name'; label.textContent = c.groupLabel || (groupId + ' (group)');
          label.addEventListener('click', (ev) => {
            if (yieldToPickSession(ev)) return;
            ev.preventDefault(); ev.stopPropagation();
            if (ev.shiftKey) {
              publishRangeFromRow(row, lastSelectIndex, 'publish');
              return;
            }
            publishGroup();
          });
          row.appendChild(indicator); row.appendChild(label);
          if (showInstancedConnections && n.instanceSourceOp) {
            const badgeMeta = getInstanceBadgeMeta(c.instanceState);
            if (badgeMeta) {
              const badge = document.createElement('span');
              badge.className = `ctrl-instance-badge ${badgeMeta.className}`;
              badge.textContent = badgeMeta.text;
              row.appendChild(badge);
            }
          }
          list.appendChild(row);
          continue;
        }
        if (hideReplaced && isControlled(c.id)) continue;
        const srcKey = resolveControlSource(c);
        const srcOpTarget = resolvePublishSourceOp(n, c);
        if (showPublishedOnly && !(srcKey && isPublished(srcOpTarget, srcKey))) continue;
        const row = document.createElement('div'); row.className = 'node-row';
        row.dataset.selectIndex = String(selectIndex++);
        row.dataset.kind = 'control';
        row.dataset.sourceOp = srcOpTarget;
        row.dataset.source = srcKey;
        row.dataset.name = c.name || c.id;
        const pendingMeta = typeof getPendingControlMeta === 'function' ? getPendingControlMeta(srcOpTarget, srcKey) : null;
        if (pendingMeta) {
          if (pendingMeta.kind && !c.kind) c.kind = pendingMeta.kind === 'label' ? 'Label' : pendingMeta.kind === 'button' ? 'Button' : c.kind;
          if (pendingMeta.labelCount != null && c.labelCount == null) c.labelCount = pendingMeta.labelCount;
          if (pendingMeta.defaultValue != null && c.defaultValue == null) c.defaultValue = pendingMeta.defaultValue;
          if (pendingMeta.inputControl && !c.inputControl) c.inputControl = pendingMeta.inputControl;
        }
        if (c.kind) row.dataset.controlKind = String(c.kind).toLowerCase();
        if (typeof c.labelCount === 'number' && Number.isFinite(c.labelCount)) {
          row.dataset.labelCount = String(c.labelCount);
        }
        if (c.inputControl) row.dataset.inputControl = c.inputControl;
        if (c.defaultValue != null) row.dataset.defaultValue = String(c.defaultValue);
        const controlMeta = buildControlMetaFromDefinition(c);
        row._mmMeta = {
          kind: 'control',
          sourceOp: srcOpTarget,
          source: srcKey,
          displayName: c.name || c.id,
          controlMeta,
        };
        if (enableNodeDrag) {
          row.draggable = true;
          row.addEventListener('dragstart', (ev) => {
            try {
              const payload = {
                kind: 'control',
                sourceOp: srcOpTarget,
                source: srcKey,
                name: c.name || c.id,
                controlKind: row.dataset.controlKind || '',
                labelCount: row.dataset.labelCount ? parseInt(row.dataset.labelCount, 10) : null,
                inputControl: row.dataset.inputControl || '',
                defaultValue: row.dataset.defaultValue || null,
              };
              const txt = 'FMR_NODE:' + JSON.stringify(payload);
              ev.dataTransfer && ev.dataTransfer.setData('text/plain', txt);
              ev.dataTransfer && (ev.dataTransfer.effectAllowed = 'copy');
            } catch (_) {}
          });
        }
        const indicator = document.createElement('span'); indicator.className = 'node-published-indicator'; indicator.title = 'Published status';
        indicator.dataset.sourceOp = srcOpTarget; indicator.dataset.source = srcKey;
        indicator.addEventListener('click', (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          if (yieldToPickSession(ev)) return;
          if (!srcKey || typeof focusPublishedControl !== 'function') return;
          if (isPublished(srcOpTarget, srcKey)) {
            try { clearHighlights(); } catch (_) {}
            try { focusPublishedControl(srcOpTarget, srcKey); } catch (_) {}
          }
        });
        const label = document.createElement('span'); label.className = 'ctrl-name'; label.textContent = c.name || c.id;
        if (isControlled(c.id)) label.classList.add('replaced');
        label.addEventListener('click', (ev) => {
          if (yieldToPickSession(ev)) return;
          ev.preventDefault(); ev.stopPropagation();
          if (ev.shiftKey) {
            if (!n.isMacroRoot || !c.isMacroUserControl) {
              publishRangeFromRow(row, lastSelectIndex, 'publish');
            }
            return;
          }
          if (n.isMacroRoot && c.isMacroUserControl) {
            lastSelectIndex = parseInt(row.dataset.selectIndex || '-1', 10);
            const already = isPublished(srcOpTarget, srcKey);
            if (!already) {
              const r = ensurePublished(srcOpTarget, srcKey, c.name, controlMeta);
              if (r) {
                const pos = getInsertionPosUnderSelection();
                try { logDiag(`Insert single idx=${r.index} at pos ${pos}`); } catch (_) {}
                state.parseResult.order = insertIndicesAt(state.parseResult.order, [r.index], pos);
              }
            } else {
              if (typeof focusPublishedControl === 'function') {
                try { clearHighlights(); } catch (_) {}
                try { focusPublishedControl(srcOpTarget, srcKey); } catch (_) {}
              }
            }
            renderPublishedList(state.parseResult.entries, state.parseResult.order);
            refreshNodesChecks();
            return;
          }
          lastSelectIndex = parseInt(row.dataset.selectIndex || '-1', 10);
          const already = isPublished(srcOpTarget, srcKey);
          if (!already) {
            const r = ensurePublished(srcOpTarget, srcKey, c.name, controlMeta);
            if (r) {
              const pos = getInsertionPosUnderSelection();
              try { logDiag(`Insert single idx=${r.index} at pos ${pos}`); } catch (_) {}
              state.parseResult.order = insertIndicesAt(state.parseResult.order, [r.index], pos);
            }
            if (typeof consumePendingControlMeta === 'function') {
              consumePendingControlMeta(srcOpTarget, srcKey);
            }
          } else if (typeof focusPublishedControl === 'function') {
            try { clearHighlights(); } catch (_) {}
            try { focusPublishedControl(srcOpTarget, srcKey); } catch (_) {}
          }
          renderPublishedList(state.parseResult.entries, state.parseResult.order);
          refreshNodesChecks();
        });
        if (showInstancedConnections && n.instanceSourceOp) {
          const badgeMeta = getInstanceBadgeMeta(c.instanceState);
          if (badgeMeta) {
            const badge = document.createElement('span');
            badge.className = `ctrl-instance-badge ${badgeMeta.className}`;
            badge.textContent = badgeMeta.text;
            row.appendChild(badge);
          }
        }
        let byEl = null;
        try {
          const mods = filterDisplayMods(c.id, getControlMods(c.id));
          const exprRefs = getControlExprRefs(c.id);
          if ((mods && mods.length > 0) || (exprRefs && exprRefs.length > 0)) {
            const by = document.createElement('span'); by.className = 'ctrl-by';
            if (mods && mods.length > 0) {
              const modName = mods[0];
              const a = document.createElement('a'); a.href = '#'; a.textContent = modName; a.title = 'Jump to modifier';
              a.addEventListener('click', (ev) => { ev.preventDefault(); ev.stopPropagation(); try { clearHighlights(); highlightCallback(modName, null); } catch (_) {} });
              const icon = document.createElement('span'); icon.className = 'sep'; icon.innerHTML = createIcon('chevron-right', 12); by.appendChild(icon); by.appendChild(a);
            }
            if (exprRefs && exprRefs.length > 0) {
              if (mods && mods.length > 0) {
                const sepText = document.createElement('span');
                sepText.className = 'ctrl-by-dot';
                sepText.textContent = '•';
                by.appendChild(sepText);
              } else {
                const icon = document.createElement('span'); icon.className = 'sep'; icon.innerHTML = createIcon('chevron-right', 12); by.appendChild(icon);
              }
              const expr = exprRefs[0];
              const target = Array.isArray(expr.targets) && expr.targets.length ? expr.targets[0] : null;
              const exprClass = classifyExpressionRef(expr);
              const exprLink = document.createElement('a');
              exprLink.href = '#';
              exprLink.className = 'ctrl-by-expr';
              if (exprClass.kind === 'direct') {
                exprLink.textContent = 'Expr';
                exprLink.title = expr.expression ? `Expression: ${expr.expression}` : 'Expression-driven';
              } else {
                const refs = Array.isArray(expr.targets) ? expr.targets.length : 0;
                exprLink.textContent = refs > 1 ? `Expr fx (${refs})` : 'Expr fx';
                exprLink.title = expr.expression
                  ? `Complex expression (${refs} refs): ${expr.expression}`
                  : `Complex expression-driven control (${refs} refs)`;
              }
              exprLink.addEventListener('click', (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
                try {
                  if (exprClass.kind === 'direct') {
                    clearHighlights();
                    if (target && target.sourceOp) highlightCallback(target.sourceOp, target.source || null);
                    return;
                  }
                  openExpressionInspector({
                    sourceOp: n.name,
                    source: srcKey,
                    displayName: c.name || c.id,
                    expression: expr.expression || '',
                    targets: Array.isArray(expr.targets) ? expr.targets : [],
                  });
                } catch (_) {}
              });
              by.appendChild(exprLink);
            }
            byEl = by;
          }
        } catch (_) {}
        row.appendChild(indicator);
        row.appendChild(label);
        if (byEl) row.appendChild(byEl);
        list.appendChild(row);
      }
      wrapper.appendChild(list);
      nodesList.appendChild(wrapper);
    }
    refreshNodesChecks();
    try {
      if (preserveNodeName) {
        const anchorNode = nodesList.querySelector(`.node[data-op="${escapedNodeName}"] .node-header`);
        if (anchorNode && Number.isFinite(anchorOffsetBefore)) {
          const anchorOffsetAfter = anchorNode.getBoundingClientRect().top;
          scrollHost.scrollTop += (anchorOffsetAfter - anchorOffsetBefore);
          return;
        }
      }
      scrollHost.scrollTop = previousScrollTop;
    } catch (_) {}
  }

  let lastSelectIndex = -1;

  function getRowSelectionKey(row) {
    if (!row) return null;
    if (row.dataset.kind === 'group') {
      return `group|${row.dataset.sourceOp || ''}|${row.dataset.groupId || row.dataset.groupBase || ''}`;
    }
    if (row.dataset.kind === 'control') {
      return `control|${row.dataset.sourceOp || ''}|${row.dataset.source || ''}`;
    }
    return null;
  }

  function setRowSelected(row, shouldSelect, sel) {
    const key = getRowSelectionKey(row);
    if (!key) return;
    if (shouldSelect) sel.add(key);
    else sel.delete(key);
    row.classList.toggle('selected', shouldSelect);
  }

  function setSingleSelection(row) {
    try {
      if (!row || !nodesList || !state.parseResult) return;
      const sel = getNodeSelection();
      sel.clear();
      state.parseResult.nodeSelectionMuted = false;
      const rows = Array.from(nodesList.querySelectorAll('.node-row[data-select-index]'));
      rows.forEach((r) => setRowSelected(r, r === row, sel));
      state.parseResult.nodeSelection = sel;
      updateNodeSelectionButtons();
      lastSelectIndex = parseInt(row.dataset.selectIndex || '-1', 10);
    } catch (_) {}
  }

  function applyRangeSelection(row, shouldSelect) {
    try {
      if (!row || !nodesList) return;
      const currentIndex = parseInt(row.dataset.selectIndex || '-1', 10);
      if (currentIndex < 0) return;
      if (state.parseResult) state.parseResult.nodeSelectionMuted = false;
      const anchor = lastSelectIndex >= 0 ? lastSelectIndex : currentIndex;
      const start = Math.min(anchor, currentIndex);
      const end = Math.max(anchor, currentIndex);
      const rows = Array.from(nodesList.querySelectorAll('.node-row[data-select-index]'));
      const sel = getNodeSelection();
      rows.forEach((r) => {
        const idx = parseInt(r.dataset.selectIndex || '-1', 10);
        if (idx >= start && idx <= end) setRowSelected(r, shouldSelect, sel);
      });
      state.parseResult.nodeSelection = sel;
      updateNodeSelectionButtons();
      lastSelectIndex = currentIndex;
    } catch (_) {}
  }

  function publishRangeFromRow(row, anchorOverride = null, mode = 'toggle') {
    try {
      if (!row || !nodesList) return;
      const currentIndex = parseInt(row.dataset.selectIndex || '-1', 10);
      if (currentIndex < 0) return;
      const anchor = (typeof anchorOverride === 'number' && anchorOverride >= 0) ? anchorOverride
        : (lastSelectIndex >= 0 ? lastSelectIndex : currentIndex);
      const start = Math.min(anchor, currentIndex);
      const end = Math.max(anchor, currentIndex);
      const rows = Array.from(nodesList.querySelectorAll('.node-row[data-select-index]'))
        .filter(r => {
          const idx = parseInt(r.dataset.selectIndex || '-1', 10);
          return idx >= start && idx <= end;
        })
        .sort((a, b) => parseInt(a.dataset.selectIndex || '0', 10) - parseInt(b.dataset.selectIndex || '0', 10));
      if (!rows.length) return;
      const firstMeta = rows.find(r => r && r._mmMeta) ? rows.find(r => r && r._mmMeta)._mmMeta : null;
      const shouldPublish = mode === 'publish'
        ? true
        : (firstMeta && firstMeta.kind === 'group'
          ? !((firstMeta.channels || []).length > 0 && (firstMeta.channels || []).every((ch) => {
            const chOp = String(ch?.sourceOp || firstMeta.sourceOp || '').trim();
            const chSrc = String(ch?.source || ch?.id || '').trim();
            return !!chSrc && isPublished(chOp, chSrc);
          }))
          : !(firstMeta && firstMeta.source && isPublished(firstMeta.sourceOp, firstMeta.source)));
      const indices = [];
      let insertAfterPos = -1;
      const entries = (state.parseResult && Array.isArray(state.parseResult.entries)) ? state.parseResult.entries : [];
      const order = (state.parseResult && Array.isArray(state.parseResult.order)) ? state.parseResult.order : [];
      const findEntryIndex = (op, src) => entries.findIndex(e => e && e.sourceOp === op && e.source === src);
      for (const r of rows) {
        const meta = r._mmMeta;
        if (!meta) continue;
        if (meta.kind === 'group') {
          for (const ch of (meta.channels || [])) {
            const chOp = String(ch?.sourceOp || meta.sourceOp || '').trim();
            const chSrc = String(ch?.source || ch?.id || '').trim();
            if (!chSrc) continue;
            if (shouldPublish) {
              if (!isPublished(chOp, chSrc)) {
                const chMeta = ch && ch.control ? buildControlMetaFromDefinition(ch.control) : null;
                const res = ensurePublished(chOp, chSrc, ch.name, chMeta);
                if (res) indices.push(res.index);
              } else {
                const idx = findEntryIndex(chOp, chSrc);
                if (idx >= 0) {
                  const pos = order.indexOf(idx);
                  if (pos > insertAfterPos) insertAfterPos = pos;
                }
              }
            } else if (mode !== 'publish' && isPublished(chOp, chSrc)) {
              removePublished(chOp, chSrc);
            }
          }
        } else if (meta.kind === 'control') {
          if (shouldPublish) {
            if (!isPublished(meta.sourceOp, meta.source)) {
              const res = ensurePublished(meta.sourceOp, meta.source, meta.displayName, meta.controlMeta || null);
              if (res) indices.push(res.index);
              if (typeof consumePendingControlMeta === 'function') {
                consumePendingControlMeta(meta.sourceOp, meta.source);
              }
            } else {
              const idx = findEntryIndex(meta.sourceOp, meta.source);
              if (idx >= 0) {
                const pos = order.indexOf(idx);
                if (pos > insertAfterPos) insertAfterPos = pos;
              }
            }
          } else if (mode !== 'publish' && isPublished(meta.sourceOp, meta.source)) {
            removePublished(meta.sourceOp, meta.source);
          }
        }
      }
      if (shouldPublish && indices.length) {
        const fallbackPos = getInsertionPosUnderSelection();
        const pos = insertAfterPos >= 0 ? insertAfterPos + 1 : fallbackPos;
        try { logDiag(`Insert range count=${indices.length} at pos ${pos}`); } catch (_) {}
        state.parseResult.order = insertIndicesAt(state.parseResult.order, indices, pos);
      }
      renderPublishedList(state.parseResult.entries, state.parseResult.order);
      refreshNodesChecks();
      lastSelectIndex = currentIndex;
    } catch (_) {}
  }

  function parseToolsInGroup(text, groupOpen, groupClose) {
    const out = [];
    if (groupOpen == null || groupClose == null) return out;
    const segment = text.slice(groupOpen, groupClose);
    const match = /Tools\s*=\s*ordered\(\)\s*\{/.exec(segment);
    if (!match) return out;
    const toolsPos = groupOpen + match.index;
    const open = toolsPos + match[0].lastIndexOf('{');
    if (open < 0) return out;
    const close = findMatchingBrace(text, open);
    if (close < 0 || close > groupClose) return out;
    const inner = text.slice(open + 1, close);
    let i = 0, depth = 0, inStr = false;
    while (i < inner.length) {
      const ch = inner[i];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
      if (ch === '\"') { inStr = true; i++; continue; }
      if (ch === '{') { depth++; i++; continue; }
      if (ch === '}') { depth--; i++; continue; }
      if (depth === 0 && isIdentStart(ch)) {
        const nameStart = i; i++;
        while (i < inner.length && isIdentPart(inner[i])) i++;
        const toolName = inner.slice(nameStart, i);
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner[i] !== '=') { i++; continue; }
        i++;
        while (i < inner.length && isSpace(inner[i])) i++;
        const typeStart = i; while (i < inner.length && isIdentPart(inner[i])) i++;
        const toolType = inner.slice(typeStart, i);
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner[i] !== '{') { i++; continue; }
        const tOpen = i;
        const tClose = findMatchingBrace(inner, tOpen);
        if (tClose < 0) break;
        const body = inner.slice(tOpen + 1, tClose);
        const instanceSourceOp = extractTopLevelQuotedProp(body, 'SourceOp');
        out.push({ name: toolName, type: toolType, body, instanceSourceOp: instanceSourceOp || null });
        i = tClose + 1; continue;
      }
      i++;
    }
    return out;
  }

  function extractTopLevelQuotedProp(body, prop) {
    try {
      if (!body || !prop) return null;
      let i = 0;
      let depth = 0;
      let inStr = false;
      while (i < body.length) {
        const ch = body[i];
        if (inStr) {
          if (ch === '"' && !isQuoteEscaped(body, i)) inStr = false;
          i += 1;
          continue;
        }
        if (ch === '"') {
          inStr = true;
          i += 1;
          continue;
        }
        if (ch === '{') { depth += 1; i += 1; continue; }
        if (ch === '}') { depth -= 1; i += 1; continue; }
        if (depth !== 0) { i += 1; continue; }
        if (!isIdentStart(ch)) { i += 1; continue; }
        const nameStart = i;
        i += 1;
        while (i < body.length && isIdentPart(body[i])) i += 1;
        const key = body.slice(nameStart, i);
        while (i < body.length && isSpace(body[i])) i += 1;
        if (body[i] !== '=') { i += 1; continue; }
        i += 1;
        while (i < body.length && isSpace(body[i])) i += 1;
        if (key !== prop) {
          if (body[i] === '"') {
            i += 1;
            while (i < body.length) {
              if (body[i] === '"' && !isQuoteEscaped(body, i)) { i += 1; break; }
              i += 1;
            }
          }
          continue;
        }
        if (body[i] !== '"') return null;
        i += 1;
        const start = i;
        while (i < body.length) {
          if (body[i] === '"' && !isQuoteEscaped(body, i)) {
            return body.slice(start, i);
          }
          i += 1;
        }
        return null;
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function parseModifiersInGroup(text, groupOpen, groupClose) {
    const out = [];
    if (groupOpen == null || groupClose == null) return out;
    const segment = text.slice(groupOpen, groupClose);
    const match = /Modifiers\s*=\s*ordered\(\)\s*\{/.exec(segment);
    if (!match) return out;
    const modsPos = groupOpen + match.index;
    const open = modsPos + match[0].lastIndexOf('{');
    if (open < 0) return out;
    const close = findMatchingBrace(text, open);
    if (close < 0 || close > groupClose) return out;
    const inner = text.slice(open + 1, close);
    let i = 0, depth = 0, inStr = false;
    while (i < inner.length) {
      const ch = inner[i];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
      if (ch === '\"') { inStr = true; i++; continue; }
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
        const instanceSourceOp = extractTopLevelQuotedProp(body, 'SourceOp');
        out.push({ name: modName, type: modType, body, isModifier: true, instanceSourceOp: instanceSourceOp || null });
        i = mClose + 1; continue;
      }
      i++;
    }
    return out;
  }

  function findModifierRefsInToolBody(body) {
    const res = [];
    if (!body) return res;
    const ip = body.indexOf('Inputs');
    if (ip < 0) return res;
    let i = ip + 'Inputs'.length;
    while (i < body.length && isSpace(body[i])) i++;
    if (body[i] !== '=') return res; i++;
    while (i < body.length && isSpace(body[i])) i++;
    if (body.slice(i, i + 8).toLowerCase() === 'ordered(') {
      i += 8;
      while (i < body.length && isSpace(body[i])) i++;
      if (body[i] === ')') { i++; }
      while (i < body.length && isSpace(body[i])) i++;
    }
    if (body[i] !== '{') return res;
    const open = i; const close = findMatchingBrace(body, open); if (close < 0) return res;
    const inner = body.slice(open + 1, close);
    let j = 0, depth = 0, inStr = false;
    while (j < inner.length) {
      const ch = inner[j];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(inner, j)) inStr = false; j++; continue; }
      if (ch === '\"') { inStr = true; j++; continue; }
      if (ch === '{') { depth++; j++; continue; }
      if (ch === '}') { depth--; j++; continue; }
      if (depth === 0 && isIdentStart(ch)) {
        const idStart = j; j++;
        while (j < inner.length && isIdentPart(inner[j])) j++;
        const id = inner.slice(idStart, j);
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner[j] !== '=') { j++; continue; }
        j++;
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner.slice(j, j + 5) !== 'Input') { continue; }
        while (j < inner.length && inner[j] !== '{') j++;
        if (inner[j] !== '{') { continue; }
        const bOpen = j; const bClose = findMatchingBrace(inner, bOpen); if (bClose < 0) break;
        const ibody = inner.slice(bOpen + 1, bClose);
        const mod = extractQuotedProp(ibody, 'SourceOp');
        const src = extractQuotedProp(ibody, 'Source') || '';
        if (mod) res.push({ id, modifier: mod, source: src });
        j = bClose + 1; continue;
      }
      j++;
    }
    if (res.length === 0) {
      try {
        const re2 = /(\n|\r|^)\s*([A-Za-z_\[\]\.][A-Za-z0-9_\[\]\.]*?)\s*=\s*Input\s*\{[\s\S]*?SourceOp\s*=\s*"([^"]+)"/g;
        let m2;
        while ((m2 = re2.exec(body)) !== null) {
          const id2 = m2[2];
          const mod2 = m2[3];
          if (mod2) res.push({ id: id2, modifier: mod2, source: '' });
        }
      } catch (_) {}
    }
    return res;
  }

  function findExpressionRefsInToolBody(body) {
    const res = [];
    if (!body) return res;
    const ip = body.indexOf('Inputs');
    if (ip < 0) return res;
    let i = ip + 'Inputs'.length;
    while (i < body.length && isSpace(body[i])) i++;
    if (body[i] !== '=') return res; i++;
    while (i < body.length && isSpace(body[i])) i++;
    if (body.slice(i, i + 8).toLowerCase() === 'ordered(') {
      i += 8;
      while (i < body.length && isSpace(body[i])) i++;
      if (body[i] === ')') i++;
      while (i < body.length && isSpace(body[i])) i++;
    }
    if (body[i] !== '{') return res;
    const open = i; const close = findMatchingBrace(body, open); if (close < 0) return res;
    const inner = body.slice(open + 1, close);
    let j = 0, depth = 0, inStr = false;
    while (j < inner.length) {
      const ch = inner[j];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(inner, j)) inStr = false; j++; continue; }
      if (ch === '\"') { inStr = true; j++; continue; }
      if (ch === '{') { depth++; j++; continue; }
      if (ch === '}') { depth--; j++; continue; }
      if (depth === 0 && isIdentStart(ch)) {
        const idStart = j; j++;
        while (j < inner.length && isIdentPart(inner[j])) j++;
        const id = inner.slice(idStart, j);
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner[j] !== '=') { j++; continue; }
        j++;
        while (j < inner.length && isSpace(inner[j])) j++;
        if (inner.slice(j, j + 5) !== 'Input') { continue; }
        while (j < inner.length && inner[j] !== '{') j++;
        if (inner[j] !== '{') { continue; }
        const bOpen = j; const bClose = findMatchingBrace(inner, bOpen); if (bClose < 0) break;
        const ibody = inner.slice(bOpen + 1, bClose);
        const expr = extractQuotedProp(ibody, 'Expression') || '';
        if (expr) {
          const targets = extractExpressionTargets(expr);
          res.push({ id, expression: expr, targets });
        }
        j = bClose + 1; continue;
      }
      j++;
    }
    if (res.length === 0) {
      try {
        const re2 = /(\n|\r|^)\s*([A-Za-z_\[\]\.][A-Za-z0-9_\[\]\.]*?)\s*=\s*Input\s*\{[\s\S]*?Expression\s*=\s*"([^"]+)"/g;
        let m2;
        while ((m2 = re2.exec(body)) !== null) {
          const id2 = m2[2];
          const expr2 = m2[3];
          if (expr2) res.push({ id: id2, expression: expr2, targets: extractExpressionTargets(expr2) });
        }
      } catch (_) {}
    }
    return res;
  }

  function extractExpressionTargets(expression) {
    const out = [];
    try {
      const expr = String(expression || '');
      const re = /([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_\[\]\.][A-Za-z0-9_\[\]\.]*)/g;
      let m;
      const seen = new Set();
      while ((m = re.exec(expr)) !== null) {
        const sourceOp = m[1];
        const source = m[2];
        const key = `${sourceOp}::${source}`;
        if (!sourceOp || !source || seen.has(key)) continue;
        seen.add(key);
        out.push({ sourceOp, source });
      }
    } catch (_) {}
    return out;
  }

  function stripOuterParens(value) {
    let s = String(value || '').trim();
    let changed = true;
    while (changed && s.startsWith('(') && s.endsWith(')')) {
      changed = false;
      let depth = 0;
      let wrapsWhole = true;
      for (let i = 0; i < s.length; i += 1) {
        const ch = s[i];
        if (ch === '(') depth += 1;
        else if (ch === ')') depth -= 1;
        if (depth === 0 && i < s.length - 1) {
          wrapsWhole = false;
          break;
        }
      }
      if (wrapsWhole) {
        s = s.slice(1, -1).trim();
        changed = true;
      }
    }
    return s;
  }

  function classifyExpressionRef(exprRef) {
    try {
      const targets = Array.isArray(exprRef?.targets) ? exprRef.targets : [];
      if (targets.length !== 1) {
        return { kind: 'complex', primaryTarget: targets[0] || null, targetCount: targets.length };
      }
      const t = targets[0];
      if (!t || !t.sourceOp || !t.source) {
        return { kind: 'complex', primaryTarget: null, targetCount: targets.length };
      }
      const exprText = stripOuterParens(String(exprRef?.expression || '').trim());
      const bareRef = `${t.sourceOp}.${t.source}`;
      if (exprText === bareRef) {
        return { kind: 'direct', primaryTarget: t, targetCount: 1 };
      }
      return { kind: 'complex', primaryTarget: t, targetCount: 1 };
    } catch (_) {
      return { kind: 'complex', primaryTarget: null, targetCount: 0 };
    }
  }

  function findToolByNameAnywhere(text, toolName) {
    if (!toolName) return null;
    try {
      const name = String(toolName);
      const re = new RegExp('(^|\\n)\\s*' + name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&') + '\\s*=\\s*([A-Za-z_][A-Za-z0-9_]*)\\s*\\{', 'g');
      const disallowedTypes = new Set(['input', 'instanceinput', 'instanceoutput', 'polyline', 'fuid']);
      let m;
      while ((m = re.exec(text)) !== null) {
        const type = m[2];
        if (disallowedTypes.has(String(type || '').toLowerCase())) continue;
        if (/^value$/i.test(name)) continue;
        const openIndex = m.index + m[0].lastIndexOf('{');
        const closeIndex = findMatchingBrace(text, openIndex);
        if (closeIndex < 0) continue;
        const body = text.slice(openIndex + 1, closeIndex);
        const instanceSourceOp = extractTopLevelQuotedProp(body, 'SourceOp');
        return { name, type, body, instanceSourceOp: instanceSourceOp || null };
      }
      return null;
    } catch (_) { return null; }
  }

  function deriveControlsForTool(tool, catalog, profileRoot, contextBindings = []) {
    const fromUC = parseUserControls(tool.body);
    const fromInputs = parseToolInputs(tool.body);
    const profileControls = getProfileControls(profileRoot || nodeProfiles, tool.type);
    const debugDynamic = isDynamicSlotToolType(tool?.type);
    const isUtility = String(tool.name || '').toUpperCase() === UTILITY_NODE_NAME;
    if (isUtility) {
      return groupColorControls(tool.name, fromUC);
    }
    const map = new Map();
    for (const c of (profileControls || [])) {
      if (!c || !c.id) continue;
      const slotMeta = deriveDynamicSlotMeta(c.id);
      map.set(c.id, {
        id: c.id,
        name: c.name || humanizeName(c.id),
        kind: c.kind || 'Input',
        inputControl: c.inputControl || null,
        defaultValue: c.defaultValue != null ? c.defaultValue : null,
        defaultX: c.defaultX != null ? c.defaultX : null,
        defaultY: c.defaultY != null ? c.defaultY : null,
        controlGroup: Number.isFinite(c.controlGroup) ? Number(c.controlGroup) : null,
        groupKey: c.groupKey || slotMeta?.groupKey || null,
        groupLabel: c.groupLabel || slotMeta?.groupLabel || null,
        slotType: c.slotType || slotMeta?.slotType || null,
        slotIndex: Number.isFinite(c.slotIndex) ? c.slotIndex : (Number.isFinite(slotMeta?.slotIndex) ? slotMeta.slotIndex : null),
      });
    }
    for (const c of fromInputs) {
      const existing = map.get(c.id) || { id: c.id };
      map.set(c.id, {
        ...existing,
        id: c.id,
        name: c.name,
        kind: c.kind || existing.kind || 'Input',
        groupKey: c.groupKey || existing.groupKey || null,
        groupLabel: c.groupLabel || existing.groupLabel || null,
        slotType: c.slotType || existing.slotType || null,
        slotIndex: Number.isFinite(c.slotIndex) ? c.slotIndex : (Number.isFinite(existing.slotIndex) ? existing.slotIndex : null),
        inputSourceOp: c.inputSourceOp || existing.inputSourceOp || null,
        inputSource: c.inputSource || existing.inputSource || null,
        isConnected: c.isConnected === true || existing.isConnected === true,
      });
    }
    if (debugDynamic) {
      try {
        logDiag(`[Dynamic node] ${tool.name} (${tool.type}) fromInputs=${fromInputs.map((c) => c.id).join(', ')}`);
      } catch (_) {}
    }
    for (const c of fromUC) {
      const existing = map.get(c.id) || { id: c.id };
      map.set(c.id, {
        ...existing,
        id: c.id,
        name: c.name,
        kind: c.kind || existing.kind || 'UserControl',
        launchUrl: c.launchUrl || existing.launchUrl || null,
        inputControl: c.inputControl || existing.inputControl || null,
        labelCount: Number.isFinite(c.labelCount) ? c.labelCount : (Number.isFinite(existing.labelCount) ? existing.labelCount : null),
        defaultValue: c.defaultValue != null ? c.defaultValue : existing.defaultValue,
        choiceOptions: Array.isArray(c.choiceOptions) && c.choiceOptions.length ? [...c.choiceOptions] : (Array.isArray(existing.choiceOptions) ? [...existing.choiceOptions] : []),
        multiButtonShowBasic: c.multiButtonShowBasic || existing.multiButtonShowBasic || '',
      });
    }
    const catControls = getCatalogControls(catalog, tool.type);
    if (catControls && Array.isArray(catControls)) {
      for (const c of catControls) {
        if (!c || !c.id) continue;
        if (!map.has(c.id)) map.set(c.id, { id: c.id, name: c.name || humanizeName(c.id), kind: c.kind || 'Input' });
      }
    } else {
      try { logDiag(`[Catalog miss] ${tool.name} type=${tool.type} (no entry in catalog)`); } catch (_) {}
    }
    try {
      const isPerturb = /perturb/i.test(String(tool.type || tool.name || ''));
      if (isPerturb) {
        const fallback = new Map([
          ['1','Value'], ['2','X Scale'], ['3','Y Scale'], ['6','Strength']
        ]);
        for (const [id, name] of fallback) {
          if (map.has(id)) {
            const v = map.get(id) || { id };
            if (!v.name || String(v.name).trim() === '' || v.name === id) {
              v.name = name;
              map.set(id, v);
            }
          }
        }
      }
    } catch (_) {}
    const merged = Array.from(map.values());
    if (fromUC.length) {
      const added = new Set();
      const ordered = [];
      for (const c of fromUC) {
        if (!c || !c.id) continue;
        const key = c.id;
        if (map.has(key) && !added.has(key)) {
          ordered.push(map.get(key));
          added.add(key);
        }
      }
      for (const [key, value] of map.entries()) {
        if (!added.has(key)) ordered.push(value);
      }
      const overrides = getPrimaryColorOverrides(tool, ordered);
      const adjusted = applyPrimaryColorOverrides(ordered, overrides);
      let finalControls = applyDissolveMixOverrides(tool, adjusted);
      finalControls = applySwitchSourceOverrides(tool, finalControls);
      finalControls = applyMaskPathControlOrdering(tool, finalControls);
      finalControls = expandTextPlusElementDefaults(tool, finalControls);
      finalControls = groupColorControls(tool.name, finalControls, overrides);
      finalControls = groupRangeControls(tool.name, finalControls);
      finalControls = applyProfileGroups(finalControls, profileControls);
      finalControls = expandDynamicSlotDefaults(tool, finalControls);
      finalControls = applyDynamicSlotGroups(tool, finalControls);
      finalControls = applyTextPlusShadingGroups(tool, finalControls);
      finalControls = applyModifierContextOverrides(tool, finalControls, contextBindings);
      if (debugDynamic) {
        try {
          logDiag(`[Dynamic node] ${tool.name} (${tool.type}) final=${finalControls.map((c) => c && (c.id || c.base || c.groupLabel || c.name)).join(', ')}`);
        } catch (_) {}
      }
      return applyInstanceStates(tool, finalControls);
    }
    const overrides = getPrimaryColorOverrides(tool, merged);
    const adjusted = applyPrimaryColorOverrides(merged, overrides);
    let finalControls = applyDissolveMixOverrides(tool, adjusted);
    finalControls = applySwitchSourceOverrides(tool, finalControls);
    finalControls = applyMaskPathControlOrdering(tool, finalControls);
    finalControls = expandTextPlusElementDefaults(tool, finalControls);
    finalControls = groupColorControls(tool.name, finalControls, overrides);
    finalControls = groupRangeControls(tool.name, finalControls);
    finalControls = applyProfileGroups(finalControls, profileControls);
    finalControls = expandDynamicSlotDefaults(tool, finalControls);
    finalControls = applyDynamicSlotGroups(tool, finalControls);
    finalControls = applyTextPlusShadingGroups(tool, finalControls);
    finalControls = applyModifierContextOverrides(tool, finalControls, contextBindings);
    if (debugDynamic) {
      try {
        logDiag(`[Dynamic node] ${tool.name} (${tool.type}) final=${finalControls.map((c) => c && (c.id || c.base || c.groupLabel || c.name)).join(', ')}`);
      } catch (_) {}
    }
    return applyInstanceStates(tool, finalControls);
  }

  function applyDynamicSlotGroups(tool, controls) {
    try {
      if (!tool || !Array.isArray(controls) || !controls.length) return controls;
      const type = String(tool.type || '').toLowerCase();
      if (!isDynamicSlotToolType(type)) return controls;
      const slotMap = new Map();
      const staticControls = [];
      for (const c of controls) {
        if (!c || c.type || !c.id) {
          staticControls.push(c);
          continue;
        }
        const groupKey = String(c.groupKey || '').trim();
        const slotType = String(c.slotType || '').trim();
        const slotIndex = Number(c.slotIndex);
        if (!groupKey || !slotType || !Number.isFinite(slotIndex)) {
          staticControls.push(c);
          continue;
        }
        if (!slotMap.has(groupKey)) {
          slotMap.set(groupKey, {
            key: groupKey,
            slotType,
            slotIndex,
            label: String(c.groupLabel || humanizeName(groupKey) || groupKey),
            controls: [],
          });
        }
        slotMap.get(groupKey).controls.push(c);
      }
      if (!slotMap.size) return controls;
      const commonSection = staticControls.length
        ? [{
            id: 'common-header',
            type: 'common-header',
            groupLabel: 'Common',
            groupKey: 'common',
          }, ...staticControls]
        : [];
      const slotSections = Array.from(slotMap.values())
        .sort((a, b) => {
          if (a.slotType !== b.slotType) return a.slotType.localeCompare(b.slotType);
          return a.slotIndex - b.slotIndex;
        })
        .flatMap((group) => ([
          {
            id: `slot-header:${group.key}`,
            type: 'slot-header',
            groupLabel: group.label,
            groupKey: group.key,
            slotType: group.slotType,
            slotIndex: group.slotIndex,
          },
          ...group.controls.sort((a, b) => {
            const aOrder = Number.isFinite(a?.dynamicSlotOrder) ? Number(a.dynamicSlotOrder) : Number.MAX_SAFE_INTEGER;
            const bOrder = Number.isFinite(b?.dynamicSlotOrder) ? Number(b.dynamicSlotOrder) : Number.MAX_SAFE_INTEGER;
            if (aOrder !== bOrder) return aOrder - bOrder;
            return String(a?.name || a?.id || '').localeCompare(String(b?.name || b?.id || ''));
          }),
        ]));
      return [...commonSection, ...slotSections];
    } catch (_) {
      return controls;
    }
  }

  function parseTextPlusElementControlMeta(id) {
    try {
      const raw = String(id || '').trim();
      if (!raw) return null;
      const m = raw.match(/^([A-Za-z_][A-Za-z0-9_]*?)([1-8])(Clone)?$/i);
      if (!m) return null;
      const prefix = String(m[1] || '').trim();
      const slotIndex = Number(m[2]);
      const suffix = String(m[3] || '');
      if (!prefix || !Number.isFinite(slotIndex)) return null;
      if (/^text$/i.test(prefix)) return null;
      return {
        prefix,
        slotIndex,
        suffix,
        templateKey: `${prefix}__${suffix || ''}`,
      };
    } catch (_) {
      return null;
    }
  }

  function applyTextPlusElementSlotMeta(control, slotMeta) {
    if (!control || !slotMeta) return control;
    const slotIndex = Number(slotMeta.slotIndex);
    if (!Number.isFinite(slotIndex)) return control;
    return {
      ...control,
      groupKey: `TextElement${slotIndex}`,
      groupLabel: `Element ${slotIndex}`,
      slotType: 'textplus-element',
      slotIndex,
      dynamicSlotOrder: Number.isFinite(slotMeta.order) ? Number(slotMeta.order) : Number.MAX_SAFE_INTEGER,
    };
  }

  function buildTextPlusElementDisplayName(templateControl, nextId, slotIndex) {
    try {
      const fallback = humanizeName(nextId);
      const baseName = String(templateControl?.name || '').trim();
      if (!baseName) return fallback;
      const replaced = baseName.replace(/\b1\b/g, String(slotIndex));
      return replaced || fallback;
    } catch (_) {
      return humanizeName(nextId);
    }
  }

  function expandTextPlusElementDefaults(tool, controls) {
    try {
      if (!tool || !Array.isArray(controls) || !controls.length) return controls;
      const type = String(tool.type || '').toLowerCase();
      if (!type.includes('textplus')) return controls;
      if (controls.some((control) => isSectionHeaderControl(control))) return controls;

      const parsedById = new Map();
      const existingIds = new Set();
      let hasElementSchema = false;
      for (const control of controls) {
        if (!control || control.type || !control.id) continue;
        const rawId = String(control.id || '').trim();
        if (!rawId) continue;
        existingIds.add(rawId);
        const parsed = parseTextPlusElementControlMeta(rawId);
        if (!parsed) continue;
        parsedById.set(rawId, parsed);
        if (/^(?:Element|Enable|Name)$/i.test(parsed.prefix)) hasElementSchema = true;
      }
      if (!hasElementSchema) return controls;

      const templateOrder = [];
      const templateOrderIndex = new Map();
      const templateControl = new Map();
      for (const control of controls) {
        if (!control || control.type || !control.id) continue;
        const rawId = String(control.id || '').trim();
        const parsed = parsedById.get(rawId);
        if (!parsed) continue;
        if (!templateOrderIndex.has(parsed.templateKey)) {
          templateOrderIndex.set(parsed.templateKey, templateOrder.length);
          templateOrder.push(parsed.templateKey);
        }
        if (parsed.slotIndex === 1 && !templateControl.has(parsed.templateKey)) {
          templateControl.set(parsed.templateKey, control);
        }
      }
      if (!templateOrder.length) return controls;

      const next = [];
      for (const control of controls) {
        if (!control || control.type || !control.id) {
          next.push(control);
          continue;
        }
        const rawId = String(control.id || '').trim();
        const parsed = parsedById.get(rawId);
        if (!parsed) {
          next.push(control);
          continue;
        }
        const slotMeta = {
          slotIndex: parsed.slotIndex,
          order: templateOrderIndex.get(parsed.templateKey),
        };
        next.push(applyTextPlusElementSlotMeta(control, slotMeta));
      }

      const synthetic = [];
      for (const templateKey of templateOrder) {
        const seed = templateControl.get(templateKey);
        if (!seed || !seed.id) continue;
        const parsedSeed = parseTextPlusElementControlMeta(seed.id);
        if (!parsedSeed) continue;
        for (let slot = 2; slot <= 8; slot += 1) {
          const nextId = `${parsedSeed.prefix}${slot}${parsedSeed.suffix || ''}`;
          if (!nextId || existingIds.has(nextId)) continue;
          existingIds.add(nextId);
          synthetic.push(applyTextPlusElementSlotMeta({
            ...seed,
            id: nextId,
            name: buildTextPlusElementDisplayName(seed, nextId, slot),
            isSyntheticDynamicSlot: true,
          }, {
            slotIndex: slot,
            order: templateOrderIndex.get(templateKey),
          }));
        }
      }

      return synthetic.length ? [...next, ...synthetic] : next;
    } catch (_) {
      return controls;
    }
  }

  function getTextPlusShadingSlotFromId(id) {
    try {
      const raw = String(id || '').trim();
      if (!raw) return null;
      const lower = raw.toLowerCase();

      let match = lower.match(/(?:^|\.)(?:red|green|blue|alpha)([123])(?:clone)?$/);
      if (match) return Number(match[1]);

      match = lower.match(/(?:enable|opacity|softnessx|softnessy|softnessglow|softnessblend|softness)([123])$/);
      if (match) return Number(match[1]);

      match = lower.match(/^text([123])fill$/);
      if (match) return Number(match[1]);

      if (lower === 'thickness'
        || lower === 'adaptthicknesstoperspective'
        || lower === 'outsideonly'
        || lower === 'cleanintersections'
        || lower === 'joinstyle'
        || lower === 'miterlimit'
        || lower === 'linestyle'
        || lower === 'separatoroutline') return 2;

      if (lower === 'shadowposition' || lower === 'separatorshadow') return 3;

      return null;
    } catch (_) {
      return null;
    }
  }

  function getTextPlusShadingSlot(control) {
    try {
      if (!control) return null;
      if (isGroupedControl(control)) {
        const channels = Array.isArray(control.channels) ? control.channels : [];
        const slots = new Set();
        for (const ch of channels) {
          const slot = getTextPlusShadingSlotFromId(ch?.id || '');
          if (Number.isFinite(slot)) slots.add(slot);
        }
        return slots.size === 1 ? Array.from(slots)[0] : null;
      }
      return getTextPlusShadingSlotFromId(control.id || '');
    } catch (_) {
      return null;
    }
  }

  function applyTextPlusShadingGroups(tool, controls) {
    try {
      if (!tool || !Array.isArray(controls) || !controls.length) return controls;
      const type = String(tool.type || '').toLowerCase();
      if (!type.includes('textplus')) return controls;
      if (controls.some((c) => isSectionHeaderControl(c))) return controls;

      const elementCommon = [];
      const elementSlots = new Map();
      let hasElementSlots = false;
      for (const control of controls) {
        const slotType = String(control?.slotType || '').trim().toLowerCase();
        const slotIndex = Number(control?.slotIndex);
        if (slotType === 'textplus-element' && Number.isFinite(slotIndex)) {
          if (!elementSlots.has(slotIndex)) elementSlots.set(slotIndex, []);
          elementSlots.get(slotIndex).push(control);
          hasElementSlots = true;
        } else {
          elementCommon.push(control);
        }
      }
      if (hasElementSlots) {
        const out = [];
        if (elementCommon.length) {
          out.push(
            {
              id: 'common-header',
              type: 'common-header',
              groupLabel: 'Common',
              groupKey: 'common',
            },
            ...elementCommon,
          );
        }
        const slotOrder = Array.from(elementSlots.keys()).sort((a, b) => a - b);
        for (const slotIndex of slotOrder) {
          const slotControls = (elementSlots.get(slotIndex) || []).slice();
          slotControls.sort((a, b) => {
            const ao = Number.isFinite(a?.dynamicSlotOrder) ? Number(a.dynamicSlotOrder) : Number.MAX_SAFE_INTEGER;
            const bo = Number.isFinite(b?.dynamicSlotOrder) ? Number(b.dynamicSlotOrder) : Number.MAX_SAFE_INTEGER;
            if (ao !== bo) return ao - bo;
            return String(a?.name || a?.id || '').localeCompare(String(b?.name || b?.id || ''));
          });
          out.push(
            {
              id: `slot-header:textplus-element-${slotIndex}`,
              type: 'slot-header',
              groupLabel: `Element ${slotIndex}`,
              groupKey: `TextElement${slotIndex}`,
              slotType: 'textplus-element',
              slotIndex,
            },
            ...slotControls,
          );
        }
        return out.length ? out : controls;
      }

      const common = [];
      const shading = new Map([[1, []], [2, []], [3, []]]);
      let hasShading = false;

      for (const control of controls) {
        const slot = getTextPlusShadingSlot(control);
        if (Number.isFinite(slot) && shading.has(slot)) {
          shading.get(slot).push(control);
          hasShading = true;
          continue;
        }
        common.push(control);
      }

      if (!hasShading) return controls;

      const out = [];
      if (common.length) {
        out.push(
          {
            id: 'common-header',
            type: 'common-header',
            groupLabel: 'Common',
            groupKey: 'common',
          },
          ...common,
        );
      }

      const labels = new Map([[1, 'Fill'], [2, 'Outline'], [3, 'Shadow']]);
      for (const slotIndex of [1, 2, 3]) {
        const slotControls = shading.get(slotIndex) || [];
        if (!slotControls.length) continue;
        out.push(
          {
            id: `slot-header:textplus-shading-${slotIndex}`,
            type: 'slot-header',
            groupLabel: labels.get(slotIndex) || `Shading ${slotIndex}`,
            groupKey: `textplus-shading-${slotIndex}`,
            slotType: 'textplus-shading',
            slotIndex,
          },
          ...slotControls,
        );
      }
      return out.length ? out : controls;
    } catch (_) {
      return controls;
    }
  }

  function applyModifierContextOverrides(tool, controls, bindings) {
    try {
      if (!tool || !Array.isArray(controls) || !controls.length) return controls;
      const type = String(tool.type || '').toLowerCase();
      const refs = Array.isArray(bindings) ? bindings : [];
      if (!refs.length) return controls;
      const diagOnce = (message) => {
        try {
          const key = `${tool.name || ''}::${tool.type || ''}`;
          if (modifierContextDiagSeen.has(key)) return;
          modifierContextDiagSeen.add(key);
          logDiag(`[Modifier context] ${tool.name} (${tool.type}) ${message}`);
        } catch (_) {}
      };

      if (type === 'offset') {
        const modeScore = { position: 0, distance: 0, angle: 0 };
        for (const ref of refs) {
          const src = String(ref?.source || '').toLowerCase();
          const id = String(ref?.id || '').toLowerCase();
          if (src === 'position' || /center|pivot|position/.test(id)) modeScore.position += 2;
          if (src === 'distance' || /distance|size|scale/.test(id)) modeScore.distance += 2;
          if (src === 'angle' || /angle|aspect|rotation/.test(id)) modeScore.angle += 2;
        }
        const context =
          modeScore.distance > modeScore.angle && modeScore.distance >= modeScore.position ? 'distance'
          : modeScore.angle > modeScore.distance && modeScore.angle >= modeScore.position ? 'angle'
          : 'position';
        diagOnce(`=> ${context} (bindings=${refs.length})`);
        const suffix = context === 'distance' ? 'Distance' : context === 'angle' ? 'Angle' : 'Position';
        return controls.map((c) => {
          if (!c || c.type) return c;
          if (String(c.id || '') !== 'Offset') return c;
          return { ...c, name: `Offset ${suffix}` };
        });
      }

      if (type.startsWith('perturb')) {
        const primary = refs[0] || null;
        const targetName = primary?.id ? humanizeName(primary.id) : '';
        if (!targetName) return controls;
        diagOnce(`=> value targets "${targetName}" (from ${primary?.tool || '?'}:${primary?.id || '?'})`);
        return controls.map((c) => {
          if (!c || c.type) return c;
          if (String(c.id || '') !== 'Value') return c;
          return { ...c, name: targetName };
        });
      }

      if (type.startsWith('switch')) {
        const primary = refs[0] || null;
        const targetName = primary?.id ? humanizeName(primary.id) : '';
        if (!targetName || /^source$/i.test(targetName)) return controls;
        diagOnce(`=> source label "${targetName} Source" (from ${primary?.tool || '?'}:${primary?.id || '?'})`);
        return controls.map((c) => {
          if (!c || c.type) return c;
          if (String(c.id || '') !== 'Source') return c;
          return { ...c, name: `${targetName} Source` };
        });
      }

      return controls;
    } catch (_) {
      return controls;
    }
  }

  function getProfileControls(profileObj, type) {
    try {
      if (!profileObj || !type) return null;
      const root = profileObj.types && typeof profileObj.types === 'object' ? profileObj.types : profileObj;
      const cleanType = String(type).trim();
      let entry = root[cleanType];
      if (!entry) {
        const lower = cleanType.toLowerCase();
        for (const key of Object.keys(root || {})) {
          if (String(key).toLowerCase() === lower) { entry = root[key]; break; }
        }
      }
      if (!entry) {
        const variants = [];
        variants.push(cleanType + 'Mod');
        variants.push(cleanType + 'Modifier');
        if (cleanType.endsWith('Mod')) variants.push(cleanType.slice(0, -3));
        if (cleanType.endsWith('Modifier')) variants.push(cleanType.slice(0, -8));
        for (const v of variants) {
          if (root[v]) { entry = root[v]; break; }
        }
      }
      if (!entry || !Array.isArray(entry.controls)) return null;
      return entry.controls;
    } catch (_) {
      return null;
    }
  }

  function applyProfileGroups(controls, profileControls) {
    try {
      if (!Array.isArray(controls) || !controls.length || !Array.isArray(profileControls) || !profileControls.length) return controls;
      const deriveProfileGroupLabel = (groupKey, explicitLabel, members) => {
        try {
          const cleanExplicit = String(explicitLabel || '').trim();
          if (cleanExplicit && !/^controlGroup:\d+$/i.test(cleanExplicit)) return cleanExplicit;
          const memberList = Array.isArray(members) ? members : [];
          const names = memberList
            .map((m) => String((m && (m.name || humanizeName(m.id))) || '').trim())
            .filter(Boolean);
          const ids = memberList.map((m) => String(m && m.id ? m.id : '').toLowerCase());
          if (ids.includes('fliphoriz') && ids.includes('flipvert')) return 'Flip';
          if (names.length >= 2) {
            const tokenized = names.map((n) => n.split(/\s+/).filter(Boolean));
            const first = tokenized[0] || [];
            const shared = [];
            for (let i = 0; i < first.length; i += 1) {
              const token = first[i];
              if (!tokenized.every((arr) => String(arr[i] || '').toLowerCase() === String(token).toLowerCase())) break;
              shared.push(token);
            }
            if (shared.length) return shared.join(' ');
            if (names.length === 2) return `${names[0]} / ${names[1]}`;
          }
          if (names.length === 1) return names[0];
          if (cleanExplicit) return cleanExplicit;
          return /^controlGroup:\d+$/i.test(String(groupKey || '')) ? 'Grouped Controls' : (humanizeName(groupKey) || groupKey);
        } catch (_) {
          return humanizeName(groupKey) || groupKey;
        }
      };
      const groupDefs = new Map();
      for (const c of profileControls) {
        if (!c || !c.id) continue;
        let key = '';
        if (c.groupKey != null && String(c.groupKey).trim()) {
          key = String(c.groupKey).trim();
        } else if (Number.isFinite(c.controlGroup) && Number(c.controlGroup) > 0) {
          key = `controlGroup:${Number(c.controlGroup)}`;
        }
        if (!key) continue;
        if (!groupDefs.has(key)) groupDefs.set(key, { ids: new Set(), label: c.groupLabel || c.name || null, leader: key });
        const def = groupDefs.get(key);
        def.ids.add(String(c.id));
        if (!def.label && (c.groupLabel || c.name)) def.label = c.groupLabel || c.name;
      }
      if (!groupDefs.size) return controls;
      const controlById = new Map();
      for (const c of controls) {
        if (!c || c.type) continue;
        if (!c.id) continue;
        controlById.set(String(c.id), c);
      }
      const activeGroups = new Map();
      for (const [gk, def] of groupDefs.entries()) {
        const members = Array.from(def.ids).map((id) => controlById.get(id)).filter(Boolean);
        if (members.length < 2) continue;
        activeGroups.set(gk, {
          id: gk,
          label: deriveProfileGroupLabel(gk, def.label, members),
          memberIds: new Set(members.map((m) => String(m.id))),
        });
      }
      if (!activeGroups.size) return controls;
      const emittedGroups = new Set();
      const groupedMembers = new Set();
      const next = [];
      for (const c of controls) {
        if (!c || c.type) {
          next.push(c);
          continue;
        }
        const cid = String(c.id || '');
        if (!cid) {
          next.push(c);
          continue;
        }
        if (groupedMembers.has(cid)) continue;
        let matched = null;
        for (const [gk, g] of activeGroups.entries()) {
          if (g.memberIds.has(cid)) { matched = { gk, g }; break; }
        }
        if (!matched) {
          next.push(c);
          continue;
        }
        if (emittedGroups.has(matched.gk)) {
          groupedMembers.add(cid);
          continue;
        }
        const channels = [];
        for (const c2 of controls) {
          if (!c2 || c2.type || !c2.id) continue;
          const id2 = String(c2.id);
          if (!matched.g.memberIds.has(id2)) continue;
          channels.push({ id: c2.id, name: c2.name || c2.id, control: c2 });
          groupedMembers.add(id2);
        }
        if (channels.length >= 2) {
          next.push({
            id: `group:${matched.gk}`,
            base: matched.gk,
            type: 'linked-group',
            groupLabel: matched.g.label,
            channels,
          });
          emittedGroups.add(matched.gk);
        } else {
          next.push(c);
        }
      }
      return next;
    } catch (_) {
      return controls;
    }
  }

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

  function getPrimaryColorOverrides(tool, controls) {
    try {
      const type = String(tool?.type || '').toLowerCase();
      const isTextPlus = type.includes('textplus');
      const isGenerator = type.includes('generator') || type.includes('colorgeneratorplugin');
      const isBackground = type.includes('background');
      if (!isTextPlus && !isGenerator && !isBackground) return null;
      const channelNameOverrides = new Map();
      const groupLabelOverrides = new Map();
      let primaryBase = '';
      let primaryId = '';
      let primaryScore = -1;
      for (const c of controls || []) {
        if (!c || !c.id) continue;
        const id = String(c.id);
        const channelMeta = parseColorChannelMeta(id);
        if (!channelMeta || channelMeta.channel !== 'Red') continue;
        const norm = id.replace(/[^a-z0-9]/gi, '').toLowerCase();
        let score = 0;
        if (norm === 'red1clone') score = 40;
        else if (norm === 'red1') score = 35;
        else if (norm === 'topleftclonered') score = 30;
        else if (norm === 'topleftred') score = 25;
        else if (/clone/i.test(channelMeta.base)) score = 10;
        if (score > primaryScore) {
          primaryScore = score;
          primaryBase = channelMeta.base;
          primaryId = id;
        }
      }
      if (primaryBase && primaryId) {
        channelNameOverrides.set(primaryId, 'Color');
        groupLabelOverrides.set(primaryBase, 'Color');
      }
      if (!channelNameOverrides.size && !groupLabelOverrides.size) return null;
      const preferTopPlacement = !isTextPlus;
      return {
        channelNameOverrides,
        groupLabelOverrides,
        primaryGroupBase: preferTopPlacement ? (primaryBase || null) : null,
      };
    } catch (_) {
      return null;
    }
  }

  function applyPrimaryColorOverrides(controls, overrides) {
    if (!overrides || !(overrides.channelNameOverrides instanceof Map)) return controls;
    if (!controls || !controls.length) return controls;
    let changed = false;
    const next = controls.map((c) => {
      if (!c || !c.id) return c;
      const nameOverride = overrides.channelNameOverrides.get(c.id);
      if (!nameOverride) return c;
      if (c.name === nameOverride) return c;
      changed = true;
      return { ...c, name: nameOverride };
    });
    return changed ? next : controls;
  }

  function applyDissolveMixOverrides(tool, controls) {
    try {
      const type = String(tool?.type || '').toLowerCase();
      if (!type.includes('dissolve')) return controls;
      if (!Array.isArray(controls)) return controls;
      const mixIndex = controls.findIndex((c) => c && c.id && String(c.id).toLowerCase() === 'mix');
      if (mixIndex >= 0) {
        const next = controls.slice();
        const mixControl = { ...next[mixIndex], name: 'Mix' };
        next.splice(mixIndex, 1);
        next.unshift(mixControl);
        return next;
      }
      return [{ id: 'Mix', name: 'Mix', kind: 'Number' }, ...controls];
    } catch (_) {
      return controls;
    }
  }

  function ensureDissolveMixControl(node, controls) {
    try {
      if (!node) return controls;
      const type = String(node.type || '').toLowerCase();
      if (!type.includes('dissolve')) return controls;
      if (!Array.isArray(controls)) return controls;
      const mixIndex = controls.findIndex((c) => c && c.id && String(c.id).toLowerCase() === 'mix');
      if (mixIndex >= 0) {
        const next = controls.slice();
        const mixControl = { ...next[mixIndex], name: 'Mix' };
        next.splice(mixIndex, 1);
        next.unshift(mixControl);
        return next;
      }
      return [{ id: 'Mix', name: 'Mix', kind: 'Number' }, ...controls];
    } catch (_) {
      return controls;
    }
  }

  function applySwitchSourceOverrides(tool, controls) {
    try {
      const type = String(tool?.type || '').toLowerCase();
      if (!type.includes('switch')) return controls;
      if (!Array.isArray(controls) || !controls.length) return controls;
      const index = controls.findIndex((c) => c && c.id && String(c.id).toLowerCase() === 'source');
      if (index <= 0) return controls;
      const next = controls.slice();
      const [sourceCtrl] = next.splice(index, 1);
      next.unshift(sourceCtrl);
      return next;
    } catch (_) {
      return controls;
    }
  }

  function getCatalogControls(catObj, type) {
    try {
      if (!catObj || !type) return null;
      const cleanType = String(type).trim();
      let entry = catObj[cleanType];
      if (!entry) {
        const tLower = cleanType.toLowerCase();
        for (const k of Object.keys(catObj)) {
          if (String(k).toLowerCase() === tLower) { entry = catObj[k]; break; }
        }
      }
      if (!entry) {
        for (const k of Object.keys(catObj)) {
          const v = catObj[k];
          if (v && String(v.type || '').toLowerCase() === cleanType.toLowerCase()) { entry = v; break; }
        }
      }
      if (!entry) {
        for (const k of Object.keys(catObj)) {
          const v = catObj[k];
          if (v && String(v.toolType || '').toLowerCase() === cleanType.toLowerCase()) { entry = v; break; }
        }
      }
      if (!entry) {
        const variants = [];
        variants.push(cleanType + 'Mod');
        variants.push(cleanType + 'Modifier');
        if (cleanType.endsWith('Mod')) variants.push(cleanType.slice(0, -3));
        if (cleanType.endsWith('Modifier')) variants.push(cleanType.slice(0, -8));
        for (const v of variants) {
          if (catObj[v]) { entry = catObj[v]; break; }
        }
      }
      if (!entry) {
        const SYN = {
          'PerturbPoint': ['PerturbPoint','Perturb'],
          'Perturb': ['Perturb','PerturbPoint'],
          'AnimCurves': ['AnimCurves','LUTLookup','LUTBezier'],
          'LUTLookup': ['LUTLookup','AnimCurves','LUTBezier'],
          'LUTBezier': ['LUTBezier','AnimCurves','LUTLookup']
        };
        const cand = SYN[cleanType];
        if (cand) {
          for (const name of cand) { if (catObj[name]) { entry = catObj[name]; break; } }
        }
      }
      if (!entry) {
        const tLower = cleanType.toLowerCase();
        for (const k of Object.keys(catObj)) {
          const kLower = String(k).toLowerCase();
          if (kLower.includes(tLower) || tLower.includes(kLower)) { entry = catObj[k]; break; }
          const v = catObj[k];
          const tt = String((v && (v.toolType || v.type)) || '').toLowerCase();
          if (tt && (tt.includes(tLower) || tLower.includes(tt))) { entry = catObj[k]; break; }
        }
      }
      if (!entry) return null;
      if (Array.isArray(entry)) return entry;
      if (Array.isArray(entry.controls)) return entry.controls;
      return null;
    } catch (_) { return null; }
  }

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

  function hasTypeInProfiles(profiles, type) {
    try {
      if (!profiles || !type) return false;
      const root = profiles.types && typeof profiles.types === 'object' ? profiles.types : profiles;
      const t = String(type).trim().toLowerCase();
      if (Object.prototype.hasOwnProperty.call(root, type)) return true;
      for (const k of Object.keys(root || {})) {
        if (String(k).toLowerCase() === t) return true;
      }
      return false;
    } catch (_) {
      return false;
    }
  }

  function isModifierType(type) {
    try {
      return hasTypeInCatalog(modifierCatalog, type) || hasTypeInProfiles(modifierProfiles, type);
    } catch (_) {
      return false;
    }
  }

  function parseUserControls(body) {
    const match = /UserControls\s*=\s*(?:ordered\(\))?\s*\{/m.exec(body);
    if (!match) return [];
    const open = match.index + match[0].lastIndexOf('{');
    if (open < 0) return [];
    const close = findMatchingBrace(body, open);
    if (close < 0) return [];
    const uc = body.slice(open + 1, close);
    const controls = [];
    const extractControlStringPropList = (segment, prop) => {
      try {
        if (!segment || !prop) return [];
        const escapedProp = String(prop).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(`(?:\\{\\s*)?${escapedProp}\\s*=\\s*("(?:\\\\.|[^"])*"|[^,}\\r\\n]+)\\s*(?:\\})?`, 'ig');
        const values = [];
        let match;
        while ((match = re.exec(segment))) {
          let value = (match[1] || '').trim();
          if (!value) continue;
          if (value.endsWith(',')) value = value.slice(0, -1).trim();
          value = value.replace(/\}\s*$/, '').trim();
          if (!value) continue;
          if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) {
            value = value.slice(1, -1).replace(/\\"/g, '"').replace(/\\\\/g, '\\');
          }
          if (value) values.push(value);
        }
        return values;
      } catch (_) {
        return [];
      }
    };
    const extractChoiceOptionsFromBody = (segment) => {
      try {
        const combo = extractControlStringPropList(segment, 'CCS_AddString');
        if (combo.length) return combo;
        const buttons = extractControlStringPropList(segment, 'MBTNC_AddButton');
        if (buttons.length) return buttons;
        return [];
      } catch (_) {
        return [];
      }
    };
    let i = 0, depth = 0, inStr = false;
    while (i < uc.length) {
      const ch = uc[i];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(uc, i)) inStr = false; i++; continue; }
      if (ch === '\"') { inStr = true; i++; continue; }
      if (ch === '{') { depth++; i++; continue; }
      if (ch === '}') { depth--; i++; continue; }
      if (depth === 0 && isIdentStart(ch)) {
        const idStart = i; i++;
        while (i < uc.length && isIdentPart(uc[i])) i++;
        const rawId = uc.slice(idStart, i);
        const id = normalizeId(rawId);
        while (i < uc.length && isSpace(uc[i])) i++;
        if (uc[i] !== '=') { i++; continue; }
        i++;
        while (i < uc.length && isSpace(uc[i])) i++;
        if (uc[i] !== '{') { i++; continue; }
        const cOpen = i;
        const cClose = findMatchingBrace(uc, cOpen);
        if (cClose < 0) break;
        const cBody = uc.slice(cOpen + 1, cClose);
        const disp = extractQuotedProp(cBody, 'LINKS_Name') || humanizeName(id);
        const inputControl = extractQuotedProp(cBody, 'INPID_InputControl') || '';
        const lowerInput = String(inputControl || '').toLowerCase();
        const isBtn = lowerInput === 'buttoncontrol';
        const isLabel = lowerInput === 'labelcontrol';
        let kind = 'UserControl';
        if (isBtn) kind = 'Button';
        else if (isLabel) kind = 'Label';
        let launchUrl = null;
        try {
          let m = cBody.match(/bmd\.(?:openurl|openUrl)\s*\(\s*['"](https?:\/\/[^'"\)]+)['"]\s*\)/i);
          if (!m && isBtn) {
            const m2 = cBody.match(/BTNCS_Execute\s*=\s*"([^"]+)"/i);
            if (m2 && m2[1]) {
              const decoded = m2[1].replace(/\\"/g, '"');
              const m3 = decoded.match(/https?:\/\/[^'"\s]+/i);
              if (m3) m = [null, m3[0]];
            }
          }
          if (m && m[1]) launchUrl = m[1];
        } catch (_) {}
        const labelCountMatch = cBody.match(/LBLC_NumInputs\s*=\s*([0-9]+)/i);
        const labelCount = labelCountMatch ? parseInt(labelCountMatch[1], 10) : null;
        const defaultMatch = cBody.match(/INP_Default\s*=\s*([0-9\.\-]+)/i);
        const defaultValue = defaultMatch ? defaultMatch[1] : null;
        const choiceOptions = extractChoiceOptionsFromBody(cBody);
        const basicMatch = cBody.match(/MBTNC_ShowBasicButton\s*=\s*([A-Za-z0-9_."']+)/i);
        const multiButtonShowBasic = basicMatch ? String(basicMatch[1]).replace(/,$/, '').trim() : '';
        controls.push({ id, name: disp, kind, launchUrl, inputControl, labelCount, defaultValue, choiceOptions, multiButtonShowBasic });
        i = cClose + 1; continue;
      }
      i++;
    }
    return controls;
  }

  function parseGroupLevelUserControls(text, groupOpen, groupClose) {
    try {
      if (!text || groupOpen == null || groupClose == null || groupClose <= groupOpen) return [];
      const body = text.slice(groupOpen + 1, groupClose);
      return parseUserControls(body) || [];
    } catch (_) {
      return [];
    }
  }

  function parseTruthyVisible(value) {
    if (value == null || value === '') return true;
    const lowered = String(value).trim().toLowerCase();
    if (!lowered) return true;
    return lowered !== 'false' && lowered !== '0' && lowered !== 'nil';
  }

  function parseLaunchUrlFromGroupControl(control) {
    try {
      if (!control) return null;
      const rawBody = String(control.rawBody || '');
      let match = rawBody.match(/bmd\.(?:openurl|openUrl)\s*\(\s*['"](https?:\/\/[^'"\)]+)['"]\s*\)/i);
      if (!match) {
        const script = String(control.buttonExecute || '').trim();
        if (script) {
          const urlMatch = script.match(/https?:\/\/[^'"\s]+/i);
          if (urlMatch) match = [null, urlMatch[0]];
        }
      }
      return match && match[1] ? match[1] : null;
    } catch (_) {
      return null;
    }
  }

  function buildMacroRootControls(result, macroName) {
    try {
      const controls = Array.isArray(result?.groupUserControls) ? result.groupUserControls : [];
      const mapped = controls
        .filter((control) => control && control.id)
        .filter((control) => control.kind !== 'system')
        .filter((control) => !!String(control.inputControl || '').trim())
        .filter((control) => parseTruthyVisible(control.visible))
        .map((control) => {
          const inputControl = String(control.inputControl || '').trim();
          const lowerInput = inputControl.toLowerCase();
          const isBtn = lowerInput === 'buttoncontrol';
          const isLabel = lowerInput === 'labelcontrol';
          const name = extractQuotedProp(control.rawBody || '', 'LINKS_Name') || humanizeName(control.id);
          const labelCountMatch = String(control.rawBody || '').match(/LBLC_NumInputs\s*=\s*([0-9]+)/i);
          return {
            id: String(control.id),
            name,
            kind: isBtn ? 'Button' : isLabel ? 'Label' : 'UserControl',
            inputControl,
            page: control.page || 'Controls',
            labelCount: labelCountMatch ? parseInt(labelCountMatch[1], 10) : null,
            defaultValue: control.defaultValue != null ? control.defaultValue : null,
            choiceOptions: Array.isArray(control.choiceOptions) ? [...control.choiceOptions] : [],
            launchUrl: parseLaunchUrlFromGroupControl(control),
            isMacroUserControl: true,
            sourceOp: macroName,
          };
        });
      return groupColorControls(macroName, mapped);
    } catch (_) {
      return [];
    }
  }

  function parseToolInputEntries(body) {
    const match = /Inputs\s*=\s*(?:ordered\(\))?\s*\{/m.exec(body);
    if (!match) return [];
    const open = match.index + match[0].lastIndexOf('{');
    if (open < 0) return [];
    const close = findMatchingBrace(body, open);
    if (close < 0) return [];
    const inner = body.slice(open + 1, close);
    const out = [];
    let i = 0, depth = 0, inStr = false;
    while (i < inner.length) {
      const ch = inner[i];
      if (inStr) { if (ch === '\"' && !isQuoteEscaped(inner, i)) inStr = false; i++; continue; }
      if (ch === '\"') { inStr = true; i++; continue; }
      if (ch === '{') { depth++; i++; continue; }
      if (ch === '}') { depth--; i++; continue; }
      if (depth === 0 && (isIdentStart(ch) || ch === '[')) {
        const idStart = i;
        let rawId = '';
        if (ch === '[') {
          let j = i + 1;
          while (j < inner.length && inner[j] !== ']') j++;
          rawId = inner.slice(i, Math.min(j + 1, inner.length));
          i = Math.min(j + 1, inner.length);
        } else {
          i++;
          while (i < inner.length && isIdentPart(inner[i])) i++;
          rawId = inner.slice(idStart, i);
        }
        const id = normalizeId(rawId);
        while (i < inner.length && isSpace(inner[i])) i++;
        if (inner[i] !== '=') { i++; continue; }
        i++;
        while (i < inner.length && isSpace(inner[i])) i++;
        const isInput = inner.slice(i, i + 5) === 'Input';
        const isInstanceInput = inner.slice(i, i + 13) === 'InstanceInput';
        if (!isInput && !isInstanceInput) { continue; }
        while (i < inner.length && inner[i] !== '{') i++;
        if (inner[i] !== '{') { continue; }
        const cOpen = i;
        const cClose = findMatchingBrace(inner, cOpen);
        if (cClose < 0) break;
        const cBody = inner.slice(cOpen + 1, cClose);
        if (!isConnectorInputName(id)) {
          const slotMeta = deriveDynamicSlotMeta(id);
          out.push({
            id,
            name: slotMeta?.displayName || humanizeName(id),
            kind: 'Input',
            inputBody: cBody,
            groupKey: slotMeta?.groupKey || null,
            groupLabel: slotMeta?.groupLabel || null,
            slotType: slotMeta?.slotType || null,
            slotIndex: Number.isFinite(slotMeta?.slotIndex) ? slotMeta.slotIndex : null,
          });
        }
        i = cClose + 1; continue;
      }
      i++;
    }
    return out;
  }

  const DYNAMIC_SLOT_TOOL_TYPES = new Set(['multimerge', 'multitext', 'multipoly']);

  const DYNAMIC_SLOT_SEED_KEYS = {
    multimerge: {
      layer: [
        'Foreground_LayerSelect',
        'Foreground',
        'Center',
        'Size',
        'Angle',
        'ApplyBlank1',
        'ApplyMode',
        'Operator',
        'SubtractiveAdditive',
        'Gain',
        'BurnIn',
        'Blend',
        'ClampCoverage',
        'ApplyBlank2',
        'Edges',
        'FilterMethod',
        'WindowMethod',
        'InvertTransform',
        'LayerName',
      ],
    },
    multitext: {
      text: [
        'StyledText',
        'ClearSelectedStyling',
        'ClearAllStyling',
        'Style',
        'Size',
        'CharacterSpacing',
        'JustifyAll',
        'IndentStep',
        'IndentUnit',
        'LineSpacing',
        'Scroll',
        'ScrollPosition',
        'Tab1Position',
        'Tab1Alignment',
        'Tab2Position',
        'Tab2Alignment',
        'Tab3Position',
        'Tab3Alignment',
        'Tab4Position',
        'Tab4Alignment',
        'Tab5Position',
        'Tab5Alignment',
        'Tab6Position',
        'Tab6Alignment',
        'Tab7Position',
        'Tab7Alignment',
        'Tab8Position',
        'Tab8Alignment',
        'Direction',
        'LineDirection',
        'ReadingDirection',
        'Orientation',
        'UnderlinePosition',
        'ForceMonospaced',
        'UseFontKerning',
        'UseLigatures',
        'SplitLigatures',
        'StylisticSet',
        'KerningSeparator',
        'ManualFontKerning',
        'ClearSelectedKerning',
        'ClearAllKerning',
        'ManualFontPlacement',
        'ClearSelectedPlacement',
        'ClearAllPlacement',
        'ApplyMode',
        'MergeOperator',
        'SubtractiveAdditive',
        'Gain',
        'BurnIn',
        'LayoutType',
        'Wrap',
        'Clip',
        'Center',
        'CenterZ',
        'LayoutSize',
        'LayoutWidth',
        'LayoutHeight',
        'Perspective',
        'FitCharacters',
        'AlignType',
        'RotationOrder',
        'AngleX',
        'AngleY',
        'AngleZ',
        'VerticalTopCenterBottom',
        'VerticallyJustified',
        'HorizontalLeftCenterRight',
        'HorizontallyJustified',
        'Enable1',
        'Opacity1',
        'SoftnessX1',
        'SoftnessY1',
        'SoftnessGlow1',
        'SoftnessBlend1',
        'SeparatorOutline',
        'Enable2',
        'Opacity2',
        'Thickness',
        'AdaptThicknessToPerspective',
        'OutsideOnly',
        'CleanIntersections',
        'JoinStyle',
        'MiterLimit',
        'LineStyle',
        'SoftnessX2',
        'SoftnessY2',
        'SoftnessGlow2',
        'SoftnessBlend2',
        'SeparatorShadow',
        'Enable3',
        'Opacity3',
        'ShadowPosition',
        'SoftnessX3',
        'SoftnessY3',
        'SoftnessGlow3',
        'SoftnessBlend3',
        'CharacterLevelStylingBase',
        'CharacterLevelStyling',
        'Text1Fill',
        'Text2Fill',
        'Text3Fill',
        'Justification',
        'IndentDecrease',
        'Softness1',
        'Softness2',
        'Softness3',
        'Text1',
        'Text1AdvancedControls',
        'Text1AlignLayers',
        'Text1Font',
        'Text1Layout',
        'Text1LayoutHeader',
        'Text1LayoutRotation',
        'Text1Merge',
        'Text1Outline',
        'Text1Paragraph',
        'Text1Scroll',
        'Text1Shading',
        'Text1Shadow',
      ],
    },
    multipoly: {
      poly: [
        'Name',
        'Filter',
        'BorderWidth',
        'Size',
        'Center',
        'Polyline',
        'Polyline2',
      ],
    },
  };

  function isDynamicSlotToolType(type) {
    return DYNAMIC_SLOT_TOOL_TYPES.has(String(type || '').trim().toLowerCase());
  }

  function createDynamicSlotMeta(slotType, slotIndex, groupKey, groupLabel, displayName) {
    return {
      slotType,
      slotIndex: parseInt(slotIndex, 10),
      groupKey,
      groupLabel,
      displayName,
    };
  }

  function deriveDynamicSlotMeta(id) {
    try {
      const raw = String(id || '').trim();
      if (!raw) return null;
      let match = raw.match(/^(Layer)(\d+)\.(.+)$/i);
      if (match) {
        return createDynamicSlotMeta('layer', match[2], `Layer${match[2]}`, `Layer ${match[2]}`, humanizeName(match[3]));
      }
      match = raw.match(/^(LayerName)(\d+)$/i);
      if (match) {
        return createDynamicSlotMeta('layer', match[2], `Layer${match[2]}`, `Layer ${match[2]}`, 'Layer Name');
      }
      match = raw.match(/^(Text)(\d+)\.(.+)$/i);
      if (match) {
        let suffix = String(match[3] || '');
        const repeatedPrefix = new RegExp(`^Text${match[2]}`, 'i');
        suffix = suffix.replace(repeatedPrefix, '') || String(match[3] || '');
        return createDynamicSlotMeta('text', match[2], `Text${match[2]}`, `Text ${match[2]}`, humanizeName(suffix));
      }
      match = raw.match(/^(PolyMask)(\d+)\.(.+)$/i);
      if (match) {
        return createDynamicSlotMeta('poly', match[2], `PolyMask${match[2]}`, `Poly ${match[2]}`, humanizeName(match[3]));
      }
      match = raw.match(/^(PolyMaskName)(\d+)$/i);
      if (match) {
        return createDynamicSlotMeta('poly', match[2], `PolyMask${match[2]}`, `Poly ${match[2]}`, 'Name');
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function getDynamicSlotTemplateDescriptor(id, slotType) {
    try {
      const raw = String(id || '').trim();
      const type = String(slotType || '').trim().toLowerCase();
      if (!raw || !type) return null;
      let match = raw.match(/^Layer(\d+)\.(.+)$/i);
      if (type === 'layer' && match) {
        const suffix = String(match[2] || '').trim();
        if (!suffix) return null;
        return {
          templateKey: suffix,
          buildId: (slotIndex) => `Layer${slotIndex}.${suffix}`,
        };
      }
      match = raw.match(/^LayerName(\d+)$/i);
      if (type === 'layer' && match) {
        return {
          templateKey: 'LayerName',
          buildId: (slotIndex) => `LayerName${slotIndex}`,
        };
      }
      match = raw.match(/^Text(\d+)\.(.+)$/i);
      if (type === 'text' && match) {
        const suffix = String(match[2] || '').trim();
        if (!suffix) return null;
        return {
          templateKey: suffix,
          buildId: (slotIndex) => `Text${slotIndex}.${suffix}`,
        };
      }
      match = raw.match(/^PolyMask(\d+)\.(.+)$/i);
      if (type === 'poly' && match) {
        const suffix = String(match[2] || '').trim();
        if (!suffix) return null;
        return {
          templateKey: suffix,
          buildId: (slotIndex) => `PolyMask${slotIndex}.${suffix}`,
        };
      }
      match = raw.match(/^PolyMaskName(\d+)$/i);
      if (type === 'poly' && match) {
        return {
          templateKey: 'Name',
          buildId: (slotIndex) => `PolyMaskName${slotIndex}`,
        };
      }
      return null;
    } catch (_) {
      return null;
    }
  }

  function getDynamicSlotSeedTemplateKeys(toolType, slotType) {
    try {
      const type = String(toolType || '').trim().toLowerCase();
      const slot = String(slotType || '').trim().toLowerCase();
      return Array.isArray(DYNAMIC_SLOT_SEED_KEYS[type]?.[slot]) ? [...DYNAMIC_SLOT_SEED_KEYS[type][slot]] : [];
    } catch (_) {
      return [];
    }
  }

  function buildDynamicSlotId(slotType, slotIndex, templateKey) {
    try {
      const slot = String(slotType || '').trim().toLowerCase();
      const key = String(templateKey || '').trim();
      const index = Number(slotIndex);
      if (!slot || !key || !Number.isFinite(index)) return '';
      if (slot === 'layer') return key === 'LayerName' ? `LayerName${index}` : `Layer${index}.${key}`;
      if (slot === 'text') {
        const slotKey = key.replace(/^Text\d+/i, `Text${index}`);
        return `Text${index}.${slotKey}`;
      }
      if (slot === 'poly') return key === 'Name' ? `PolyMaskName${index}` : `PolyMask${index}.${key}`;
      return '';
    } catch (_) {
      return '';
    }
  }

  function expandDynamicSlotDefaults(tool, controls) {
    try {
      if (!tool || !Array.isArray(controls) || !controls.length) return controls;
      const type = String(tool.type || '').toLowerCase();
      if (!isDynamicSlotToolType(type)) return controls;
      const slotDefs = new Map();
      const nonSlotControls = [];
      const actualById = new Map();
      const extraOrderBySlotType = new Map();
      for (const c of controls) {
        if (!c || c.type || !c.id) {
          nonSlotControls.push(c);
          continue;
        }
        const id = String(c.id || '').trim();
        if (!id) {
          nonSlotControls.push(c);
          continue;
        }
        const groupKey = String(c.groupKey || '').trim();
        const slotType = String(c.slotType || '').trim().toLowerCase();
        const slotIndex = Number(c.slotIndex);
        if (!groupKey || !slotType || !Number.isFinite(slotIndex)) {
          nonSlotControls.push(c);
          continue;
        }
        if (!slotDefs.has(groupKey)) {
          slotDefs.set(groupKey, {
            groupKey,
            groupLabel: String(c.groupLabel || humanizeName(groupKey) || groupKey),
            slotType,
            slotIndex,
          });
        }
        actualById.set(id, c);
        const descriptor = getDynamicSlotTemplateDescriptor(id, slotType);
        if (!descriptor) continue;
        if (!extraOrderBySlotType.has(slotType)) extraOrderBySlotType.set(slotType, []);
        const orderList = extraOrderBySlotType.get(slotType);
        if (!orderList.includes(descriptor.templateKey)) orderList.push(descriptor.templateKey);
      }
      if (!slotDefs.size) return controls;
      const expanded = [...nonSlotControls];
      const sortedSlots = Array.from(slotDefs.values()).sort((a, b) => {
        if (a.slotType !== b.slotType) return a.slotType.localeCompare(b.slotType);
        return a.slotIndex - b.slotIndex;
      });
      for (const slot of sortedSlots) {
        const slotType = slot.slotType;
        const seeded = getDynamicSlotSeedTemplateKeys(type, slotType);
        const discovered = extraOrderBySlotType.get(slotType) || [];
        const orderedKeys = [...seeded];
        for (const key of discovered) {
          if (!orderedKeys.includes(key)) orderedKeys.push(key);
        }
        const emittedIds = new Set();
        for (let i = 0; i < orderedKeys.length; i += 1) {
          const templateKey = orderedKeys[i];
          const nextId = buildDynamicSlotId(slotType, slot.slotIndex, templateKey);
          if (!nextId || emittedIds.has(nextId)) continue;
          const actual = actualById.get(nextId);
          const slotMeta = deriveDynamicSlotMeta(nextId);
          expanded.push(actual ? {
            ...actual,
            groupKey: slot.groupKey,
            groupLabel: slot.groupLabel,
            slotType: slot.slotType,
            slotIndex: slot.slotIndex,
            dynamicSlotOrder: i,
          } : {
            id: nextId,
            name: slotMeta?.displayName || humanizeName(nextId),
            kind: 'Input',
            groupKey: slot.groupKey,
            groupLabel: slot.groupLabel,
            slotType: slot.slotType,
            slotIndex: slot.slotIndex,
            dynamicSlotOrder: i,
            isSyntheticDynamicSlot: true,
          });
          emittedIds.add(nextId);
        }
      }
      return expanded;
    } catch (_) {
      return controls;
    }
  }

  function applyMaskPathControlOrdering(tool, controls) {
    try {
      if (!tool || !Array.isArray(controls) || controls.length < 2) return controls;
      if (!isMaskToolTypeForPath(tool.type)) return controls;
      const rankFor = (control) => {
        const id = String(control?.id || '').trim();
        if (!id) return null;
        if (/^polyline$/i.test(id)) return 0;
        if (/^(?:polygon|bspline|spline|path)$/i.test(id)) return 1;
        if (isPathControlId(id)) return 2;
        return null;
      };
      const prioritized = [];
      const remainder = [];
      for (const control of controls) {
        const rank = rankFor(control);
        if (rank == null) remainder.push(control);
        else prioritized.push({ control, rank });
      }
      if (!prioritized.length) return controls;
      prioritized.sort((a, b) => a.rank - b.rank);
      return [...prioritized.map((item) => item.control), ...remainder];
    } catch (_) {
      return controls;
    }
  }

  function parseToolInputs(body) {
    const entries = parseToolInputEntries(body);
    return entries.map(({ id, name, kind, groupKey, groupLabel, slotType, slotIndex, inputBody }) => {
      const inputSourceOp = String(extractQuotedProp(String(inputBody || ''), 'SourceOp') || '').trim();
      const inputSource = String(extractQuotedProp(String(inputBody || ''), 'Source') || '').trim();
      return ({
      id,
      name,
      kind,
      groupKey: groupKey || null,
      groupLabel: groupLabel || null,
      slotType: slotType || null,
      slotIndex: Number.isFinite(slotIndex) ? slotIndex : null,
      inputSourceOp: inputSourceOp || null,
      inputSource: inputSource || null,
      isConnected: !!inputSourceOp,
    });
    });
  }

  function parseInstanceInputStateMap(body) {
    const map = new Map();
    try {
      const entries = parseToolInputEntries(body);
      for (const entry of entries) {
        if (!entry || !entry.id) continue;
        const compact = String(entry.inputBody || '').replace(/[\s,]/g, '').trim();
        if (!compact) map.set(String(entry.id), 'instanced-explicit');
        else map.set(String(entry.id), 'overridden');
      }
    } catch (_) {}
    return map;
  }

  function resolveInstanceStateForId(id, stateMap) {
    const key = String(id || '').trim();
    if (!key) return null;
    if (stateMap instanceof Map && stateMap.has(key)) return stateMap.get(key);
    return 'instanced';
  }

  function reduceInstanceState(states) {
    const clean = (states || []).filter(Boolean);
    if (!clean.length) return null;
    const uniq = Array.from(new Set(clean));
    if (uniq.length === 1) return uniq[0];
    return 'mixed';
  }

  function applyInstanceStates(tool, controls) {
    try {
      if (!tool || !tool.instanceSourceOp || !Array.isArray(controls) || !controls.length) return controls;
      const stateMap = parseInstanceInputStateMap(tool.body);
      return controls.map((c) => {
        if (!c) return c;
        if (isGroupedControl(c)) {
          const channels = (c.channels || []).map((ch) => {
            if (!ch) return ch;
            const chState = resolveInstanceStateForId(ch.id, stateMap);
            return { ...ch, instanceState: chState };
          });
          const groupState = reduceInstanceState(channels.map((ch) => ch && ch.instanceState));
          return { ...c, channels, instanceState: groupState };
        }
        const instanceState = resolveInstanceStateForId(c.id, stateMap);
        return { ...c, instanceState };
      });
    } catch (_) {
      return controls;
    }
  }

  function getInstanceBadgeMeta(instanceState) {
    const state = String(instanceState || '').trim().toLowerCase();
    if (!state) return null;
    if (state === 'instanced') {
      return { text: 'Instanced', className: 'instanced' };
    }
    if (state === 'instanced-explicit') {
      return { text: 'Instanced (Explicit)', className: 'instanced-explicit' };
    }
    if (state === 'overridden') {
      return { text: 'Overridden', className: 'overridden' };
    }
    if (state === 'mixed') {
      return { text: 'Mixed', className: 'mixed' };
    }
    return { text: humanizeName(state), className: 'instanced' };
  }

  function isConnectorInputName(id) {
    return false;
  }

  function groupColorControls(sourceOp, controls, options = {}) {
    const opts = options && typeof options === 'object' ? options : {};
    const groupLabelOverrides = opts.groupLabelOverrides instanceof Map ? opts.groupLabelOverrides : new Map();
    const primaryGroupBase = opts.primaryGroupBase ? String(opts.primaryGroupBase) : '';
    const groups = new Map();
    const others = [];
    for (const c of controls) {
      const meta = parseColorChannelMeta(c?.id);
      if (!meta) { others.push({ ...c, sourceOp }); continue; }
      const base = meta.base;
      if (!groups.has(base)) groups.set(base, []);
      groups.get(base).push({ ...c, base, sourceOp, colorChannel: meta.channel });
    }
    const grouped = [];
    for (const [base, ch] of groups.entries()) {
      const sorted = ['Red','Green','Blue','Alpha'].map(s => ch.find(v => v.colorChannel === s)).filter(Boolean);
      const sharedString = (getter) => {
        const values = sorted
          .map((entry) => String(getter(entry) || '').trim())
          .filter(Boolean);
        if (!values.length) return '';
        const first = values[0];
        return values.every((value) => value === first) ? first : '';
      };
      const sharedNumber = (getter) => {
        const values = sorted
          .map((entry) => getter(entry))
          .filter((value) => Number.isFinite(value))
          .map((value) => Number(value));
        if (!values.length) return null;
        const first = values[0];
        return values.every((value) => value === first) ? first : null;
      };
      const sharedGroupKey = sharedString((entry) => entry.groupKey);
      const sharedGroupLabel = sharedString((entry) => entry.groupLabel);
      const sharedSlotType = sharedString((entry) => entry.slotType);
      const sharedSlotIndex = sharedNumber((entry) => entry.slotIndex);
      const sharedDynamicOrder = sharedNumber((entry) => entry.dynamicSlotOrder);
      grouped.push({
        id: base, base, type: 'color-group',
        channels: sorted.map(c => ({ id: c.id, name: c.name })),
        groupKey: sharedGroupKey || null,
        groupLabel: (groupLabelOverrides.get(base) || sharedGroupLabel || `${humanizeName(base)} (RGB)`),
        slotType: sharedSlotType || null,
        slotIndex: Number.isFinite(sharedSlotIndex) ? Number(sharedSlotIndex) : null,
        dynamicSlotOrder: Number.isFinite(sharedDynamicOrder) ? Number(sharedDynamicOrder) : null,
      });
    }
    if (primaryGroupBase) {
      const primaryIndex = grouped.findIndex((item) => item && item.type === 'color-group' && item.base === primaryGroupBase);
      if (primaryIndex > 0) {
        const [primary] = grouped.splice(primaryIndex, 1);
        grouped.unshift(primary);
      }
    }
    for (const o of others) grouped.push(o);
    return grouped;
  }

  function parseRangeBoundMeta(id) {
    const raw = String(id || '').trim();
    if (!raw) return null;
    const separatorMatch = raw.match(/^(.*?)[._](Min|Max)$/i);
    if (separatorMatch && separatorMatch[1]) {
      return {
        base: separatorMatch[1],
        bound: separatorMatch[2].charAt(0).toUpperCase() + separatorMatch[2].slice(1).toLowerCase(),
      };
    }
    return null;
  }

  function groupRangeControls(sourceOp, controls) {
    try {
      if (!Array.isArray(controls) || controls.length < 2) return controls;
      const byBase = new Map();
      for (let i = 0; i < controls.length; i += 1) {
        const control = controls[i];
        if (!control || control.type || !control.id) continue;
        const meta = parseRangeBoundMeta(control.id);
        if (!meta || !meta.base) continue;
        if (!byBase.has(meta.base)) {
          byBase.set(meta.base, { base: meta.base, min: null, max: null, firstIndex: i });
        }
        const entry = byBase.get(meta.base);
        if (meta.bound === 'Min') entry.min = control;
        if (meta.bound === 'Max') entry.max = control;
        if (i < entry.firstIndex) entry.firstIndex = i;
      }
      const eligible = new Map();
      for (const [base, entry] of byBase.entries()) {
        if (entry && entry.min && entry.max) eligible.set(base, entry);
      }
      if (!eligible.size) return controls;

      const emitted = new Set();
      const consumedIds = new Set();
      const out = [];
      for (const control of controls) {
        if (!control) {
          out.push(control);
          continue;
        }
        if (control.type || !control.id) {
          out.push(control);
          continue;
        }
        const id = String(control.id);
        if (consumedIds.has(id)) continue;
        const meta = parseRangeBoundMeta(id);
        if (!meta || !eligible.has(meta.base)) {
          out.push({ ...control, sourceOp });
          continue;
        }
        const entry = eligible.get(meta.base);
        if (!entry || emitted.has(meta.base)) {
          consumedIds.add(id);
          continue;
        }
        const minControl = entry.min;
        const maxControl = entry.max;
        if (!minControl || !maxControl) {
          out.push({ ...control, sourceOp });
          continue;
        }
        // Keep range child labels compact in Fusion inspector.
        const minName = 'Min';
        const maxName = 'Max';
        out.push({
          id: `range:${meta.base}`,
          base: meta.base,
          type: 'range-group',
          groupLabel: `${humanizeName(meta.base)} Range`,
          channels: [
            { id: minControl.id, name: minName, control: minControl },
            { id: maxControl.id, name: maxName, control: maxControl },
          ],
        });
        emitted.add(meta.base);
        consumedIds.add(String(minControl.id));
        consumedIds.add(String(maxControl.id));
      }
      return out;
    } catch (_) {
      return controls;
    }
  }

  function extractReferencedOpsFromTools(tools) {
    try {
      const out = new Set();
      for (const t of (tools || [])) {
        const body = (t && t.body) ? String(t.body) : '';
        const re = /SourceOp\s*=\s*"([^"]+)"/g;
        let m;
        while ((m = re.exec(body)) !== null) {
          if (m[1]) out.add(m[1]);
        }
      }
      return Array.from(out);
    } catch (_) { return []; }
  }

  function buildItemsFromPayload(payload) {
    if (!payload || !payload.sourceOp) return [];
    if (payload.kind === 'group' && Array.isArray(payload.channels)) {
      const items = [];
      const base = payload.base || null;
      const nameById = new Map();
      const opById = new Map();
      try {
        const tool = findToolByNameAnywhere(state.originalText, payload.sourceOp);
        if (tool) {
          const isMod = isModifierType(tool.type || '');
          const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
          const profileRef = isMod ? (modifierProfiles || nodeProfiles) : nodeProfiles;
          const contextBindings = isMod ? (modifierBindingContext.get(tool.name) || []) : [];
          const ctrls = deriveControlsForTool(tool, cat, profileRef, contextBindings) || [];
          for (const c of ctrls) {
            if (c && c.id) nameById.set(String(c.id), c.name || c.id);
          }
          if (/perturb/i.test(tool.type || '')) {
            const fb = new Map([['1','Value'],['2','X Scale'],['3','Y Scale'],['6','Strength']]);
            for (const [id, nm] of fb) if (!nameById.has(id)) nameById.set(id, nm);
          }
        }
      } catch (_) {}
      for (const ch of payload.channels) {
        const rawId = (ch && typeof ch === 'object')
          ? (ch.source || ch.id || '')
          : ch;
        const id = String(rawId || '').trim();
        if (!id) continue;
        const fallbackName = (ch && typeof ch === 'object' && ch.name)
          ? String(ch.name)
          : (humanizeName(id) || id);
        const sourceOp = (ch && typeof ch === 'object' && ch.sourceOp)
          ? String(ch.sourceOp)
          : String(payload.sourceOp || '');
        if (sourceOp) opById.set(id, sourceOp);
        items.push({
          sourceOp: opById.get(id) || String(payload.sourceOp || ''),
          source: id,
          displayName: nameById.get(String(id)) || fallbackName,
          base,
        });
      }
      return items;
    }
    if (payload.kind === 'control' && payload.source) {
      const normalizedKind = (payload.controlKind || payload.kind || '').toString().toLowerCase();
      const labelCount = Number.isFinite(payload.labelCount) ? Number(payload.labelCount) : null;
      return [{
        sourceOp: payload.sourceOp,
        source: payload.source,
        displayName: payload.name || humanizeName(payload.source) || payload.source,
        kind: normalizedKind,
        labelCount,
        inputControl: payload.inputControl || '',
        defaultValue: payload.defaultValue != null ? payload.defaultValue : null,
      }];
    }
    return [];
  }

  function prioritizeUtilityNode(nodes, utilityName) {
    if (!Array.isArray(nodes) || !utilityName) return nodes;
    const special = [];
    const rest = [];
    for (const n of nodes) {
      if (n && n.name && n.name.toUpperCase() === String(utilityName).toUpperCase()) special.push(n);
      else rest.push(n);
    }
    return special.length ? special.concat(rest) : nodes;
  }

  function refreshNodesChecks() {
    if (!nodesList) return;
    const indicators = nodesList.querySelectorAll('.node-published-indicator');
    indicators.forEach((indicator) => {
      const row = indicator.closest ? indicator.closest('.node-row') : null;
      const meta = row && row._mmMeta ? row._mmMeta : null;
      if (meta && meta.kind === 'control') {
        indicator.classList.toggle('is-published', isPublished(meta.sourceOp, meta.source));
        return;
      }
      if (meta && meta.kind === 'group') {
        const channels = Array.isArray(meta.channels) ? meta.channels : [];
        indicator.classList.toggle('is-published', channels.length > 0 && channels.every((ch) => {
          const chOp = String(ch?.sourceOp || meta.sourceOp || '').trim();
          const chSrc = String(ch?.source || ch?.id || '').trim();
          return !!chSrc && isPublished(chOp, chSrc);
        }));
        return;
      }
      const op = indicator.dataset.sourceOp;
      const src = indicator.dataset.source;
      if (src) indicator.classList.toggle('is-published', isPublished(op, src));
      else if (indicator.dataset.groupId || indicator.dataset.groupBase) {
        const listStr = (indicator.dataset.channels || '');
        const ids = listStr ? listStr.split('|').filter(s => s) : [];
        indicator.classList.toggle('is-published', ids.length > 0 && ids.every(id => isPublished(op, id)));
      }
    });
  }

  return {
    parseAndRenderNodes,
    refreshNodesChecks,
    isNodeDragEvent,
    parseNodeDragData,
    clearNodeSelection,
    updateNodeSelectionButtons,
    getNodeSelection,
    buildItemsFromPayload,
    clearFilter: () => {
      nodeFilter = '';
      try { if (nodesSearch) nodesSearch.value = ''; } catch (_) {}
      if (state.parseResult) parseAndRenderNodes();
    },
    getNodeFilter: () => nodeFilter,
    getNodeNames: () => [...lastNodeNames],
    setNodeCatalog(data) { nodeCatalog = data || null; },
    setModifierCatalog(data) { modifierCatalog = data || null; },
    setNodeProfiles(data) { nodeProfiles = data || null; },
    setModifierProfiles(data) { modifierProfiles = data || null; },
    getNodeCatalog() { return nodeCatalog; },
    getModifierCatalog() { return modifierCatalog; },
    getNodeProfiles() { return nodeProfiles; },
    getModifierProfiles() { return modifierProfiles; },
    expandNode: (name, mode = 'open') => {
      if (!state.parseResult) return;
      if (!(state.parseResult.nodesCollapsed instanceof Set)) state.parseResult.nodesCollapsed = new Set();
      if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
      state.parseResult.nodesCollapsed.delete(name);
      state.parseResult.nodesPublishedOnly.delete(name);
      if (mode === 'published') state.parseResult.nodesPublishedOnly.add(name);
    },
    setAllNodeModes,
    setHighlightHandler(fn) {
      if (typeof fn === 'function') highlightCallback = fn;
    },
  };
}


