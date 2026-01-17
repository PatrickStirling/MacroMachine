import { findEnclosingGroupForIndex, extractQuotedProp } from './parser.js';
import { findMatchingBrace, isSpace, isIdentStart, isIdentPart, humanizeName, isQuoteEscaped } from './textUtils.js';

export function createNodesPane(options = {}) {
  const {
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
    logDiag = () => {},
    logTag = () => {},
    error = () => {},
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
    requestAddControl = null,
    getPendingControlMeta = () => null,
    consumePendingControlMeta = () => {},
  } = options;
  const UTILITY_NODE_NAME = 'UTILITY';

  let nodeCatalog = null;
  let modifierCatalog = null;
  let nodeFilter = '';
  let hideReplaced = false;
  let highlightCallback = initialHighlightNode || (() => {});

  function buildControlMetaFromDefinition(control) {
    if (!control) return null;
    const meta = {};
    let touched = false;
    if (control.kind) { meta.kind = String(control.kind).toLowerCase(); touched = true; }
    if (Number.isFinite(control.labelCount)) { meta.labelCount = Number(control.labelCount); touched = true; }
    if (control.inputControl) { meta.inputControl = control.inputControl; touched = true; }
    if (control.defaultValue != null) { meta.defaultValue = control.defaultValue; touched = true; }
    return touched ? meta : null;
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

  function clearNodeSelection() {
    if (!state.parseResult) return;
    state.parseResult.nodeSelection = new Set();
    updateNodeSelectionButtons();
    parseAndRenderNodes();
  }

  function updateNodeSelectionButtons() {
    try {
      const size = (state.parseResult && state.parseResult.nodeSelection instanceof Set) ? state.parseResult.nodeSelection.size : 0;
      if (publishSelectedBtn) publishSelectedBtn.disabled = size === 0;
      if (clearNodeSelectionBtn) clearNodeSelectionBtn.disabled = size === 0;
    } catch (_) {}
  }

  nodesSearch?.addEventListener('input', (e) => {
    nodeFilter = (e.target.value || '').toLowerCase();
    if (state.parseResult) parseAndRenderNodes();
  });

  hideReplacedEl?.addEventListener('change', (e) => {
    hideReplaced = !!e.target.checked;
    if (state.parseResult) parseAndRenderNodes();
  });

  publishSelectedBtn?.addEventListener('click', handlePublishSelected);
  clearNodeSelectionBtn?.addEventListener('click', () => clearNodeSelection());

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
        const kind = row.dataset.kind;
        const op = row.dataset.sourceOp;
        if (kind === 'group') {
          const chs = (row.dataset.channels || '').split('|').filter(Boolean);
          for (const id of chs) items.push({ sourceOp: op, source: id, name: id });
        } else if (kind === 'control') {
          const id = row.dataset.source;
          const nm = row.dataset.name || id;
          const ctrlKind = row.dataset.controlKind || '';
          const labelCount = row.dataset.labelCount ? parseInt(row.dataset.labelCount, 10) : null;
          const inputControl = row.dataset.inputControl || '';
          const defaultValue = row.dataset.defaultValue || null;
          items.push({
            sourceOp: op,
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
        });
        if (idx != null) idxs.push(idx);
      }
      if (idxs.length) {
        const pos = getInsertionPosUnderSelection();
        state.parseResult.order = insertIndicesAt(state.parseResult.order, idxs, pos);
        try { logDiag(`Batch publish count=${idxs.length} at pos ${pos}`); } catch (_) {}
        renderPublishedList(state.parseResult.entries, state.parseResult.order);
        clearNodeSelection();
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
      let grp = findEnclosingGroupForIndex(state.originalText, state.parseResult.inputs.openIndex);
      if (!grp) {
        try { logDiag('Nodes: no group bounds found; falling back to whole file.'); } catch (_) {}
        grp = { name: state.parseResult.macroName || 'Unknown', groupOpenIndex: 0, groupCloseIndex: state.originalText.length };
      }
      const tools = parseToolsInGroup(state.originalText, grp.groupOpenIndex, grp.groupCloseIndex);
      const modifiers = parseModifiersInGroup(state.originalText, grp.groupOpenIndex, grp.groupCloseIndex);
      const downstream = buildDownstreamMap(tools);
      const modifierBindings = buildModifierBindings(tools);
      let nodes = tools.map(t => {
        const typeStr = String(t.type || '');
        const isMod = !!(modifierCatalog && hasTypeInCatalog(modifierCatalog, typeStr));
        const catalogRef = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
        return {
          name: t.name,
          type: t.type,
          controls: deriveControlsForTool(t, catalogRef),
          isModifier: isMod,
          external: false,
        };
      }).concat(modifiers.map(m => ({
        name: m.name,
        type: m.type,
        controls: deriveControlsForTool(m, (modifierCatalog || nodeCatalog)),
        isModifier: true,
        external: false,
      })));
      const controlledBy = buildControlledByMap(modifierBindings);
      nodes = nodes.map(n => ({
        ...n,
        downstream: downstream.get(n.name) || [],
        bindings: modifierBindings.get(n.name) || [],
        controlledBy: controlledBy.get(n.name) || new Map(),
      }));
      augmentWithReferencedNodes(nodes, tools);
      augmentWithAllTools(nodes);
      const downstreamAll = buildDownstreamForAll(nodes);
      nodes = nodes.map(n => ({
        ...n,
        downstream: downstreamAll.get(n.name) || [],
        bindings: modifierBindings.get(n.name) || (n.bindings || []),
        controlledBy: (controlledBy.get(n.name) || n.controlledBy || new Map()),
      }));

      const macroControls = parseGroupLevelUserControls(
        state.originalText,
        grp.groupOpenIndex,
        grp.groupCloseIndex,
      );
      if (macroControls.length) {
        const macroName = state.parseResult?.macroName || grp.name || 'Macro';
        const macroNode = {
          name: macroName,
          type: state.parseResult?.operatorType || 'GroupOperator',
          controls: groupColorControls(macroName, macroControls),
          isModifier: false,
          external: false,
          isMacroRoot: true,
          downstream: [],
          downstreamAll: [],
          bindings: [],
          controlledBy: new Map(),
        };
        nodes.unshift(macroNode);
      }
      nodes = prioritizeUtilityNode(nodes, UTILITY_NODE_NAME);
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
        map.get(r.modifier).push({ tool: t.name, id: r.id });
      }
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

  function augmentWithReferencedNodes(nodes, tools) {
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
            const isMod = !!(modifierCatalog && hasTypeInCatalog(modifierCatalog, typeStr));
            const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
            nodes.push({ name: t.name, type: t.type, controls: deriveControlsForTool(t, cat), isModifier: isMod, external: true });
            existing.add(op);
          }
        }
      }
    } catch (_) {}
  }

  function augmentWithAllTools(nodes) {
    try {
      const existingNames = new Set(nodes.map(n => n.name));
      const reAll = /(^|\n)\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*\{/g;
      let mAll;
      while ((mAll = reAll.exec(state.originalText)) !== null) {
        const name = mAll[2];
        const type = mAll[3];
        if (!name || existingNames.has(name)) continue;
        const typeStr = String(type || '');
        if (/^(GroupInfo|OperatorInfo)$/i.test(typeStr)) continue;
        const t = findToolByNameAnywhere(state.originalText, name);
        if (!t) continue;
        const isMod = !!(modifierCatalog && hasTypeInCatalog(modifierCatalog, typeStr));
        const cat = isMod ? (modifierCatalog || nodeCatalog) : nodeCatalog;
        nodes.push({ name: t.name, type: t.type, controls: deriveControlsForTool(t, cat), isModifier: isMod, external: true });
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

  function renderNodes(nodes) {
    if (!nodesList) return;
    nodesList.innerHTML = '';
    if (!state.parseResult) state.parseResult = {};
    if (!(state.parseResult.nodesCollapsed instanceof Set)) state.parseResult.nodesCollapsed = new Set();
    if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
    if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
    if (!state.parseResult.nodesViewInitialized) {
      state.parseResult.nodesCollapsed.clear();
      state.parseResult.nodesPublishedOnly.clear();
      nodes.forEach(n => {
        const hasPublished = (n.controls || []).some(c => {
          if (c.type === 'color-group') return (c.channels || []).some(ch => isPublished(n.name, ch.id));
          const key = resolveControlSource(c);
          return key ? isPublished(n.name, key) : false;
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
    for (const n of nodes) {
      const wrapper = document.createElement('div');
      wrapper.className = 'node';
      wrapper.dataset.op = n.name;
      const header = document.createElement('div');
      header.className = 'node-header';
      const nodeTwisty = document.createElement('span');
      const title = document.createElement('div');
      title.className = 'node-title';
      if (n.name && n.name.toUpperCase() === UTILITY_NODE_NAME) {
        title.textContent = n.name;
      } else {
        title.textContent = `${n.name} (${n.type || 'Unknown'})`;
      }
      const flags = document.createElement('div');
      flags.className = 'node-flags';
      (function () {
        const isMod = !!n.isModifier;
        if (isMod) {
          const badge = document.createElement('span');
          badge.className = 'node-badge modifier';
          badge.textContent = 'Modifier';
          flags.appendChild(badge);
        }
      })();
      const nodeControls = ensureDissolveMixControl(n, n.controls || []);
      const hasAnyPublished = (nodeControls || []).some(c => {
        if (c.type === 'color-group') return (c.channels || []).some(ch => isPublished(n.name, ch.id));
        const key = resolveControlSource(c);
        return key ? isPublished(n.name, key) : false;
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
        renderNodes(nodes);
      });
      header.appendChild(nodeTwisty);
      header.appendChild(title);
      if (Array.isArray(n.bindings) && n.bindings.length > 0) {
        const b = n.bindings[0];
        const nextWrap = document.createElement('div');
        nextWrap.className = 'node-next';
        const sep = document.createElement('span'); sep.className = 'sep'; sep.innerHTML = createIcon('chevron-right');
        const link = document.createElement('a');
        link.href = '#'; link.className = 'deep-link';
        link.textContent = `${b.tool}.${b.id}`;
        link.title = 'Go to controlled parameter';
        link.addEventListener('click', (ev) => {
          ev.preventDefault(); ev.stopPropagation();
          try { clearHighlights(); highlightCallback(b.tool, b.id); } catch (_) {}
        });
        nextWrap.appendChild(sep);
        nextWrap.appendChild(link);
        header.appendChild(nextWrap);
      } else if (Array.isArray(n.downstream) && n.downstream.length > 0) {
        const nextWrap = document.createElement('div');
        nextWrap.className = 'node-next';
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
        nextWrap.appendChild(sep);
        nextWrap.appendChild(link);
        header.appendChild(nextWrap);
      }
      const headerTail = document.createElement('div');
      headerTail.className = 'node-header-tail';
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
      const isControlled = (id) => {
        try {
          const m = n && n.controlledBy;
          if (!m) return false;
          if (typeof m.get === 'function') {
            const arr = m.get(id);
            return !!(arr && arr.length);
          }
          const arr = m[id];
          return !!(arr && arr.length);
        } catch (_) { return false; }
      };
        for (const c of (nodeControls || [])) {
          if (c.id === 'SourceOp') continue;
          if (c.id === 'Source' && !allowSourceControl) continue;
          if (c.type === 'color-group') {
          if (hideReplaced) {
            const anyControlled = (c.channels || []).some(ch => isControlled(ch.id));
            if (anyControlled) continue;
          }
          if (showPublishedOnly) {
            const allPublished = (c.channels || []).length > 0 && (c.channels || []).every(ch => isPublished(n.name, ch.id));
            if (!allPublished) continue;
          }
          const row = document.createElement('div'); row.className = 'node-row';
          row.dataset.kind = 'group';
          row.dataset.sourceOp = n.name;
          row.dataset.groupBase = c.base || '';
          row.dataset.channels = (c.channels || []).map(ch => ch.id).join('|');
          try {
            const key0 = `group|${n.name}|${c.base || ''}`;
            if (getNodeSelection().has(key0)) row.classList.add('selected');
          } catch (_) {}
          if (enableNodeDrag) {
            row.draggable = true;
            row.addEventListener('dragstart', (ev) => {
              try {
                const payload = { kind: 'group', sourceOp: n.name, base: c.base, channels: (c.channels || []).map(ch => ch.id) };
                const txt = 'FMR_NODE:' + JSON.stringify(payload);
                ev.dataTransfer && ev.dataTransfer.setData('text/plain', txt);
                ev.dataTransfer && (ev.dataTransfer.effectAllowed = 'copy');
              } catch (_) {}
            });
          }
          const cb = document.createElement('input'); cb.type = 'checkbox'; cb.className = 'node-ctrl group'; cb.title = 'Toggle publish color group';
          cb.dataset.sourceOp = n.name; cb.dataset.groupBase = c.base; cb.dataset.channels = (c.channels || []).map(ch => ch.id).join('|');
          cb.checked = c.channels.every(ch => isPublished(n.name, ch.id));
          cb.addEventListener('change', () => {
            if (cb.checked) {
              const allIdxs = [];
              for (const ch of c.channels) {
                const r = ensurePublished(n.name, ch.id, ch.name, null);
                if (r) allIdxs.push(r.index);
              }
              if (allIdxs.length) {
                const pos = getInsertionPosUnderSelection();
                try { logDiag(`Insert group count=${allIdxs.length} at pos ${pos}`); } catch (_) {}
                state.parseResult.order = insertIndicesAt(state.parseResult.order, allIdxs, pos);
              }
            } else {
              for (const ch of c.channels) removePublished(n.name, ch.id);
            }
            renderPublishedList(state.parseResult.entries, state.parseResult.order);
          });
          const label = document.createElement('span'); label.textContent = c.groupLabel || (c.base + ' (color)');
          row.addEventListener('click', (ev) => {
            const t = ev.target;
            if (t && t.tagName === 'INPUT') return;
            const key = `group|${n.name}|${c.base || ''}`;
            const sel = getNodeSelection();
            if (sel.has(key)) sel.delete(key); else sel.add(key);
            state.parseResult.nodeSelection = sel;
            row.classList.toggle('selected');
            updateNodeSelectionButtons();
          });
          label.addEventListener('click', (ev) => { ev.preventDefault(); ev.stopPropagation(); row.click(); });
          row.appendChild(cb); row.appendChild(label);
          list.appendChild(row);
          continue;
        }
        if (hideReplaced && isControlled(c.id)) continue;
        const srcKey = resolveControlSource(c);
        if (showPublishedOnly && !(srcKey && isPublished(n.name, srcKey))) continue;
        const row = document.createElement('div'); row.className = 'node-row';
        row.dataset.kind = 'control';
        row.dataset.sourceOp = n.name;
        row.dataset.source = srcKey;
        row.dataset.name = c.name || c.id;
        const pendingMeta = typeof getPendingControlMeta === 'function' ? getPendingControlMeta(n.name, srcKey) : null;
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
        try {
          const key0 = `control|${n.name}|${srcKey}`;
          if (getNodeSelection().has(key0)) row.classList.add('selected');
        } catch (_) {}
        const controlMeta = buildControlMetaFromDefinition(c);
        if (enableNodeDrag) {
          row.draggable = true;
          row.addEventListener('dragstart', (ev) => {
            try {
              const payload = {
                kind: 'control',
                sourceOp: n.name,
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
        const cb = document.createElement('input'); cb.type = 'checkbox'; cb.className = 'node-ctrl'; cb.title = 'Toggle publish control';
        cb.dataset.sourceOp = n.name; cb.dataset.source = srcKey;
        cb.checked = isPublished(n.name, srcKey);
        cb.addEventListener('change', () => {
          if (cb.checked) {
            const r = ensurePublished(n.name, srcKey, c.name, controlMeta);
            if (r) {
              const pos = getInsertionPosUnderSelection();
              try { logDiag(`Insert single idx=${r.index} at pos ${pos}`); } catch (_) {}
              state.parseResult.order = insertIndicesAt(state.parseResult.order, [r.index], pos);
            }
            if (typeof consumePendingControlMeta === 'function') {
              consumePendingControlMeta(n.name, srcKey);
            }
          } else {
            removePublished(n.name, srcKey);
          }
          renderPublishedList(state.parseResult.entries, state.parseResult.order);
        });
        const label = document.createElement('span'); label.className = 'ctrl-name'; label.textContent = c.name || c.id;
        row.addEventListener('click', (ev) => {
          const t = ev.target; if (t && t.tagName === 'INPUT') return;
          const key = `control|${n.name}|${srcKey}`;
          const sel = getNodeSelection();
          if (sel.has(key)) sel.delete(key); else sel.add(key);
          state.parseResult.nodeSelection = sel;
          row.classList.toggle('selected');
          updateNodeSelectionButtons();
        });
        label.addEventListener('click', (ev) => { ev.preventDefault(); ev.stopPropagation(); row.click(); });
        let byEl = null;
        try {
          const mods = (n && n.controlledBy) ? (typeof n.controlledBy.get === 'function' ? (n.controlledBy.get(c.id) || []) : (n.controlledBy[c.id] || [])) : [];
          if (mods && mods.length > 0) {
            const by = document.createElement('span'); by.className = 'ctrl-by';
            const modName = mods[0];
            const a = document.createElement('a'); a.href = '#'; a.textContent = modName; a.title = 'Jump to modifier';
            a.addEventListener('click', (ev) => { ev.preventDefault(); ev.stopPropagation(); try { clearHighlights(); highlightCallback(modName, null); } catch (_) {} });
            const icon = document.createElement('span'); icon.className = 'sep'; icon.innerHTML = createIcon('chevron-right', 12); by.appendChild(icon); by.appendChild(a);
            byEl = by;
          }
        } catch (_) {}
        row.appendChild(cb);
        row.appendChild(label);
        if (byEl) row.appendChild(byEl);
        list.appendChild(row);
      }
      wrapper.appendChild(list);
      nodesList.appendChild(wrapper);
    }
  }

  function parseToolsInGroup(text, groupOpen, groupClose) {
    const out = [];
    const toolsPos = text.indexOf('Tools = ordered()', groupOpen);
    if (toolsPos < 0 || toolsPos > groupClose) return out;
    const open = text.indexOf('{', toolsPos);
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
        out.push({ name: toolName, type: toolType, body });
        i = tClose + 1; continue;
      }
      i++;
    }
    return out;
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
        out.push({ name: modName, type: modType, body, isModifier: true });
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
        if (mod) res.push({ id, modifier: mod });
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
          if (mod2) res.push({ id: id2, modifier: mod2 });
        }
      } catch (_) {}
    }
    return res;
  }

  function findToolByNameAnywhere(text, toolName) {
    if (!toolName) return null;
    try {
      const name = String(toolName);
      const re = new RegExp('(^|\\n)\\s*' + name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&') + '\\s*=\\s*([A-Za-z_][A-Za-z0-9_]*)\\s*\\{', 'g');
      let m;
      while ((m = re.exec(text)) !== null) {
        const type = m[2];
        if (/^(Input|InstanceInput|InstanceOutput)$/i.test(type)) continue;
        const openIndex = m.index + m[0].lastIndexOf('{');
        const closeIndex = findMatchingBrace(text, openIndex);
        if (closeIndex < 0) continue;
        const body = text.slice(openIndex + 1, closeIndex);
        return { name, type, body };
      }
      return null;
    } catch (_) { return null; }
  }

  function deriveControlsForTool(tool, catalog) {
    const fromUC = parseUserControls(tool.body);
    const fromInputs = parseToolInputs(tool.body);
    const isUtility = String(tool.name || '').toUpperCase() === UTILITY_NODE_NAME;
    if (isUtility) {
      return groupColorControls(tool.name, fromUC);
    }
    const map = new Map();
    for (const c of fromInputs) map.set(c.id, { id: c.id, name: c.name, kind: c.kind || 'Input' });
    for (const c of fromUC) map.set(c.id, { id: c.id, name: c.name, kind: c.kind || 'UserControl', launchUrl: c.launchUrl || null });
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
      return groupColorControls(tool.name, finalControls, overrides);
    }
    const overrides = getPrimaryColorOverrides(tool, merged);
    const adjusted = applyPrimaryColorOverrides(merged, overrides);
    let finalControls = applyDissolveMixOverrides(tool, adjusted);
    finalControls = applySwitchSourceOverrides(tool, finalControls);
    return groupColorControls(tool.name, finalControls, overrides);
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
      let hasTopLeft = false;
      for (const c of controls || []) {
        if (!c || !c.id) continue;
        if (c.id === 'TopLeftRed') {
          channelNameOverrides.set(c.id, 'Color');
          hasTopLeft = true;
        }
      }
      if (hasTopLeft) {
        groupLabelOverrides.set('TopLeft', 'Color');
      }
      if (!channelNameOverrides.size && !groupLabelOverrides.size) return null;
      return { channelNameOverrides, groupLabelOverrides, primaryGroupBase: hasTopLeft ? 'TopLeft' : null };
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

  function parseUserControls(body) {
    const match = /UserControls\s*=\s*(?:ordered\(\))?\s*\{/m.exec(body);
    if (!match) return [];
    const open = match.index + match[0].lastIndexOf('{');
    if (open < 0) return [];
    const close = findMatchingBrace(body, open);
    if (close < 0) return [];
    const uc = body.slice(open + 1, close);
    const controls = [];
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
        controls.push({ id, name: disp, kind, launchUrl, inputControl, labelCount, defaultValue });
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

  function parseToolInputs(body) {
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
      if (depth === 0 && isIdentStart(ch)) {
        const idStart = i; i++;
        while (i < inner.length && isIdentPart(inner[i])) i++;
        const rawId = inner.slice(idStart, i);
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
        if (!isConnectorInputName(id)) out.push({ id, name: humanizeName(id), kind: 'Input' });
        i = cClose + 1; continue;
      }
      i++;
    }
    return out;
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
    const colorSuf = ['Red','Green','Blue','Alpha'];
    for (const c of controls) {
      const suf = colorSuf.find(s => c.id.endsWith(s));
      if (!suf) { others.push({ ...c, sourceOp }); continue; }
      const base = c.id.slice(0, c.id.length - suf.length);
      if (!groups.has(base)) groups.set(base, []);
      groups.get(base).push({ ...c, base, sourceOp });
    }
    const grouped = [];
    for (const [base, ch] of groups.entries()) {
      const sorted = ['Red','Green','Blue','Alpha'].map(s => ch.find(v => v.id.endsWith(s))).filter(Boolean);
      grouped.push({
        id: base, base, type: 'color-group', groupLabel: groupLabelOverrides.get(base) || `${humanizeName(base)} (RGB)`,
        channels: sorted.map(c => ({ id: c.id, name: c.name })),
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
      try {
        const tool = findToolByNameAnywhere(state.originalText, payload.sourceOp);
        if (tool) {
          const cat = (modifierCatalog && hasTypeInCatalog(modifierCatalog, tool.type || ''))
            ? (modifierCatalog || nodeCatalog)
            : nodeCatalog;
          const ctrls = deriveControlsForTool(tool, cat) || [];
          for (const c of ctrls) {
            if (c && c.id) nameById.set(String(c.id), c.name || c.id);
          }
          if (/perturb/i.test(tool.type || '')) {
            const fb = new Map([['1','Value'],['2','X Scale'],['3','Y Scale'],['6','Strength']]);
            for (const [id, nm] of fb) if (!nameById.has(id)) nameById.set(id, nm);
          }
        }
      } catch (_) {}
      for (const id of payload.channels) {
        items.push({ sourceOp: payload.sourceOp, source: id, displayName: nameById.get(String(id)) || id, base });
      }
      return items;
    }
    if (payload.kind === 'control' && payload.source) {
      const normalizedKind = (payload.controlKind || payload.kind || '').toString().toLowerCase();
      const labelCount = Number.isFinite(payload.labelCount) ? Number(payload.labelCount) : null;
      return [{
        sourceOp: payload.sourceOp,
        source: payload.source,
        displayName: payload.name || payload.source,
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
    const cbs = nodesList.querySelectorAll('input.node-ctrl');
    cbs.forEach(cb => {
      const op = cb.dataset.sourceOp;
      const src = cb.dataset.source;
      if (src) cb.checked = isPublished(op, src);
      else if (cb.dataset.groupBase) {
        const base = cb.dataset.groupBase;
        const listStr = (cb.dataset.channels || '');
        const list = listStr ? listStr.split('|').filter(s => s) : [];
        const ids = list.length ? list : [base + 'Red', base + 'Green', base + 'Blue', base + 'Alpha'];
        cb.checked = ids.length > 0 && ids.every(id => isPublished(op, id));
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
    setNodeCatalog(data) { nodeCatalog = data || null; },
    setModifierCatalog(data) { modifierCatalog = data || null; },
    getNodeCatalog() { return nodeCatalog; },
    getModifierCatalog() { return modifierCatalog; },
    expandNode: (name, mode = 'open') => {
      if (!state.parseResult) return;
      if (!(state.parseResult.nodesCollapsed instanceof Set)) state.parseResult.nodesCollapsed = new Set();
      if (!(state.parseResult.nodesPublishedOnly instanceof Set)) state.parseResult.nodesPublishedOnly = new Set();
      state.parseResult.nodesCollapsed.delete(name);
      state.parseResult.nodesPublishedOnly.delete(name);
      if (mode === 'published') state.parseResult.nodesPublishedOnly.add(name);
    },
    setHighlightHandler(fn) {
      if (typeof fn === 'function') highlightCallback = fn;
    },
  };
}
