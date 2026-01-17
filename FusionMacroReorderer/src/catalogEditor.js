export function createCatalogEditor({
  openBtn,
  closeBtn,
  root,
  datasetSelect,
  searchInput,
  typeList,
  controlList,
  emptyState,
  downloadBtn,
  nodesPane,
  logDiag = () => {},
  info = () => {},
  error = () => {},
}) {
  const state = {
    dataset: 'nodes',
    catalogs: {
      nodes: null,
      modifiers: null,
    },
    dirty: {
      nodes: false,
      modifiers: false,
    },
    selectedType: null,
    loading: false,
  };
  const dragState = { index: null };

  async function ensureCatalogLoaded(kind) {
    if (state.catalogs[kind]) return state.catalogs[kind];
    try {
      state.loading = true;
      updateEmptyState('Loading catalog…');
      let source = null;
      if (kind === 'nodes' && nodesPane?.getNodeCatalog) {
        source = nodesPane.getNodeCatalog();
      } else if (kind === 'modifiers' && nodesPane?.getModifierCatalog) {
        source = nodesPane.getModifierCatalog();
      }
      if (!source) {
        const file = kind === 'nodes' ? 'FusionNodeCatalog.cleaned.json' : 'FusionModifierCatalog.json';
        const resp = await fetch(file);
        if (!resp.ok) throw new Error(`Failed to fetch ${file}`);
        source = await resp.json();
      }
      const clone = JSON.parse(JSON.stringify(source || {}));
      state.catalogs[kind] = clone;
      state.loading = false;
      renderTypes();
      return clone;
    } catch (err) {
      state.loading = false;
      error('Catalog load failed: ' + (err?.message || err));
      updateEmptyState('Failed to load catalog.');
      throw err;
    }
  }

  function open() {
    if (!root) return;
    root.hidden = false;
    ensureCatalogLoaded(state.dataset).catch(() => {});
    renderTypes();
    renderControls();
  }

  function close() {
    if (!root) return;
    root.hidden = true;
  }

  function updateEmptyState(text) {
    if (!emptyState) return;
    emptyState.textContent = text || 'Select a type to edit its default controls.';
  }

  function getActiveCatalog() {
    return state.catalogs[state.dataset] || null;
  }

  function getFilteredTypes() {
    const cat = getActiveCatalog();
    if (!cat) return [];
    const search = (searchInput?.value || '').trim().toLowerCase();
    const names = Object.keys(cat);
    names.sort((a, b) => a.localeCompare(b));
    if (!search) return names;
    return names.filter(name => name.toLowerCase().includes(search));
  }

  function renderTypes() {
    if (!typeList) return;
    const types = getFilteredTypes();
    typeList.innerHTML = '';
    if (!types.length) {
      const li = document.createElement('li');
      li.textContent = state.loading ? 'Loading…' : 'No types found.';
      li.style.opacity = '0.8';
      typeList.appendChild(li);
      return;
    }
    const active = getActiveCatalog();
    types.forEach(name => {
      const entry = active?.[name];
      const li = document.createElement('li');
      const label = document.createElement('span');
      label.textContent = name;
      const count = document.createElement('span');
      count.className = 'count';
      const total = Array.isArray(entry?.controls) ? entry.controls.length : 0;
      count.textContent = total ? `${total}` : '0';
      li.appendChild(label);
      li.appendChild(count);
      if (state.selectedType === name) li.classList.add('selected');
      li.addEventListener('click', () => {
        state.selectedType = name;
        renderTypes();
        renderControls();
      });
      typeList.appendChild(li);
    });
  }

  function renderControls() {
    if (!controlList || !emptyState) return;
    const cat = getActiveCatalog();
    const entry = cat ? cat[state.selectedType] : null;
    const controls = Array.isArray(entry?.controls) ? entry.controls : null;
    controlList.innerHTML = '';
    if (!controls || !controls.length) {
      controlList.hidden = true;
      updateEmptyState(state.selectedType ? 'No controls found for this type.' : 'Select a type to edit its default controls.');
      return;
    }
    controlList.hidden = false;
    updateEmptyState('');
    controls.forEach((ctrl, index) => {
      const li = document.createElement('li');
      li.dataset.index = String(index);
      li.draggable = true;
      const infoBox = document.createElement('div');
      infoBox.className = 'catalog-control-info';
      const labelEl = document.createElement('div');
      labelEl.className = 'catalog-control-id';
      const codeName = ctrl?.id || ctrl?.name || '(missing id)';
      labelEl.textContent = codeName;
      if (ctrl?.name && ctrl?.name !== ctrl?.id) {
        labelEl.title = ctrl.name;
      }
      infoBox.appendChild(labelEl);
      const actions = document.createElement('div');
      actions.className = 'catalog-control-actions';
      const upBtn = document.createElement('button');
      upBtn.type = 'button';
      upBtn.textContent = '↑';
      upBtn.disabled = index === 0;
      upBtn.addEventListener('click', () => moveControl(index, -1));
      const downBtn = document.createElement('button');
      downBtn.type = 'button';
      downBtn.textContent = '↓';
      downBtn.disabled = index === controls.length - 1;
      downBtn.addEventListener('click', () => moveControl(index, +1));
      actions.appendChild(upBtn);
      actions.appendChild(downBtn);
      li.appendChild(infoBox);
      li.appendChild(actions);
      li.addEventListener('dragstart', (ev) => {
        dragState.index = index;
        li.classList.add('dragging');
        ev.dataTransfer?.setData('text/plain', String(index));
        ev.dataTransfer?.setDragImage(li, 12, 12);
      });
      li.addEventListener('dragend', () => {
        dragState.index = null;
        li.classList.remove('dragging');
        li.classList.remove('drag-over');
      });
      li.addEventListener('dragover', (ev) => {
        ev.preventDefault();
        li.classList.add('drag-over');
      });
      li.addEventListener('dragleave', () => {
        li.classList.remove('drag-over');
      });
      li.addEventListener('drop', (ev) => {
        ev.preventDefault();
        li.classList.remove('drag-over');
        const from = dragState.index ?? parseInt(ev.dataTransfer?.getData('text/plain') || '-1', 10);
        const to = parseInt(li.dataset.index || '-1', 10);
        reorderControl(from, to);
      });
      controlList.appendChild(li);
    });
  }

  function moveControl(index, delta) {
    const cat = getActiveCatalog();
    const entry = cat ? cat[state.selectedType] : null;
    const controls = entry?.controls;
    if (!Array.isArray(controls)) return;
    const newIndex = index + delta;
    if (newIndex < 0 || newIndex >= controls.length) return;
    reorderControl(index, newIndex);
  }

  function reorderControl(fromIndex, toIndex) {
    const cat = getActiveCatalog();
    const entry = cat ? cat[state.selectedType] : null;
    const controls = entry?.controls;
    if (!Array.isArray(controls)) return;
    if (fromIndex == null || toIndex == null) return;
    if (fromIndex < 0 || fromIndex >= controls.length) return;
    let target = Math.max(0, Math.min(toIndex, controls.length - 1));
    if (fromIndex === target) return;
    const [item] = controls.splice(fromIndex, 1);
    if (fromIndex < target) target -= 1;
    controls.splice(target, 0, item);
    state.dirty[state.dataset] = true;
    renderControls();
    logDiag(`Catalog "${state.selectedType}" moved index ${fromIndex} -> ${target}`);
    if (state.dataset === 'nodes') {
      nodesPane?.setNodeCatalog?.(state.catalogs.nodes);
    } else {
      nodesPane?.setModifierCatalog?.(state.catalogs.modifiers);
    }
  }

  function downloadCatalog(kind) {
    const catalog = state.catalogs[kind];
    if (!catalog) {
      error('Catalog not loaded yet.');
      return;
    }
    const fileName = kind === 'nodes' ? 'FusionNodeCatalog.cleaned.json' : 'FusionModifierCatalog.json';
    const blob = new Blob([JSON.stringify(catalog, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    info(`Downloaded ${fileName} with updated ordering.`);
  }

  datasetSelect?.addEventListener('change', async () => {
    state.dataset = datasetSelect.value === 'modifiers' ? 'modifiers' : 'nodes';
    state.selectedType = null;
    renderControls();
    renderTypes();
    await ensureCatalogLoaded(state.dataset).catch(() => {});
    renderTypes();
  });

  searchInput?.addEventListener('input', () => renderTypes());
  downloadBtn?.addEventListener('click', () => downloadCatalog(state.dataset));
  openBtn?.addEventListener('click', () => open());
  closeBtn?.addEventListener('click', () => close());
  root?.addEventListener('click', (ev) => {
    if (ev.target === root) close();
  });

  return {
    open,
    close,
  };
}
