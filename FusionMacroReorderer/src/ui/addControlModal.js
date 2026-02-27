export function createAddControlModal(options = {}) {
  const {
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
    onAddControl,
    getTargetNodes = null,
  } = options;

  let pendingNode = null;
  let availableNodes = [];

  const getKnownPageNames = () => {
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
  };

  const updateAddControlPageOptionsList = () => {
    if (!addControlPageOptions) return;
    addControlPageOptions.innerHTML = '';
    const pages = getKnownPageNames();
    pages.forEach((page) => {
      const opt = document.createElement('option');
      opt.value = page;
      addControlPageOptions.appendChild(opt);
    });
  };

  const resolveTargetNodes = () => {
    try {
      const nodes = (typeof getTargetNodes === 'function') ? getTargetNodes() : [];
      if (!Array.isArray(nodes)) return [];
      const seen = new Set();
      const out = [];
      nodes.forEach((name) => {
        const val = (name && String(name).trim()) ? String(name).trim() : '';
        if (!val || seen.has(val)) return;
        seen.add(val);
        out.push(val);
      });
      return out;
    } catch (_) {
      return [];
    }
  };

  const syncTitleWithTarget = () => {
    if (!addControlTitle) return;
    addControlTitle.textContent = pendingNode || 'Node';
  };

  const updateTargetNodeOptions = (preferredNode = '') => {
    availableNodes = resolveTargetNodes();
    let nextTarget = '';
    if (preferredNode && availableNodes.includes(preferredNode)) {
      nextTarget = preferredNode;
    } else if (pendingNode && availableNodes.includes(pendingNode)) {
      nextTarget = pendingNode;
    } else if (availableNodes.length) {
      nextTarget = availableNodes[0];
    }
    pendingNode = nextTarget || null;
    if (addControlTargetSelect) {
      addControlTargetSelect.innerHTML = '';
      if (!availableNodes.length) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = 'No nodes available';
        addControlTargetSelect.appendChild(opt);
      } else {
        availableNodes.forEach((nodeName) => {
          const opt = document.createElement('option');
          opt.value = nodeName;
          opt.textContent = nodeName;
          addControlTargetSelect.appendChild(opt);
        });
      }
      addControlTargetSelect.value = pendingNode || '';
    }
    syncTitleWithTarget();
  };

  const getSuggestedAddControlPage = () => {
    const active = (state.parseResult && state.parseResult.activePage)
      ? String(state.parseResult.activePage).trim()
      : '';
    if (active) return active;
    if (Array.isArray(state.parseResult?.pageOrder) && state.parseResult.pageOrder.length) {
      return state.parseResult.pageOrder[0];
    }
    return 'Controls';
  };

  const updateAddControlTypeVisibility = () => {
    const typeVal = (addControlTypeSelect?.value || 'label').toLowerCase();
    if (addLabelOptions) {
      const showLabel = typeVal === 'label';
      addLabelOptions.hidden = !showLabel;
      addLabelOptions.style.display = showLabel ? '' : 'none';
    }
    if (addComboOptions) {
      const showCombo = typeVal === 'combo';
      addComboOptions.hidden = !showCombo;
      addComboOptions.style.display = showCombo ? 'flex' : 'none';
    }
  };

  const parseComboOptionsInput = (raw) => {
    const text = String(raw || '').replace(/\r\n?/g, '\n');
    const rows = text.includes('\n') ? text.split('\n') : text.split(',');
    return rows
      .map((line) => String(line || '').trim().replace(/^"(.*)"$/, '$1'))
      .filter((line) => line.length > 0);
  };

  const resetAddControlFormFields = () => {
    syncTitleWithTarget();
    if (addControlNameInput) addControlNameInput.value = '';
    if (addControlTypeSelect) addControlTypeSelect.value = 'label';
    if (addControlLabelCountInput) addControlLabelCountInput.value = '0';
    if (addControlLabelDefaultSelect) addControlLabelDefaultSelect.value = 'closed';
    if (addControlComboOptionsInput) addControlComboOptionsInput.value = 'Option 1\nOption 2\nOption 3';
    if (addControlPageInput) {
      const suggested = getSuggestedAddControlPage();
      addControlPageInput.value = suggested && suggested !== 'Controls' ? suggested : '';
    }
    if (addControlError) addControlError.textContent = '';
    updateAddControlTypeVisibility();
  };

  const open = (arg) => {
    if (!addControlModal) return;
    const preferredNode = (typeof arg === 'string')
      ? arg
      : (arg && typeof arg === 'object' && arg.nodeName ? String(arg.nodeName) : '');
    updateAddControlPageOptionsList();
    updateTargetNodeOptions(preferredNode);
    resetAddControlFormFields();
    if (!pendingNode && addControlError) addControlError.textContent = 'No target nodes available.';
    addControlModal.hidden = false;
    setTimeout(() => {
      try { addControlNameInput?.focus(); } catch (_) {}
    }, 0);
  };

  const close = () => {
    if (!addControlModal) return;
    pendingNode = null;
    addControlModal.hidden = true;
  };

  addControlTypeSelect?.addEventListener('change', () => updateAddControlTypeVisibility());
  updateAddControlTypeVisibility();
  addControlTargetSelect?.addEventListener('change', () => {
    const raw = addControlTargetSelect.value || '';
    pendingNode = raw ? String(raw) : null;
    syncTitleWithTarget();
    if (pendingNode && addControlError && /No target nodes/i.test(addControlError.textContent || '')) {
      addControlError.textContent = '';
    }
  });
  addControlCancelBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    close();
  });
  addControlCloseBtn?.addEventListener('click', (ev) => {
    ev.preventDefault();
    close();
  });
  addControlModal?.addEventListener('click', (ev) => {
    if (ev.target === addControlModal) close();
  });
  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && addControlModal && !addControlModal.hidden) {
      close();
    }
  });
  if (addControlForm) {
    addControlForm.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      try {
        if (!pendingNode) {
          if (addControlError) addControlError.textContent = 'Select a target node before adding controls.';
          return;
        }
        const typeRaw = (addControlTypeSelect?.value || 'label').toLowerCase();
        const allowedTypes = new Set(['label', 'separator', 'button', 'slider', 'combo', 'screw']);
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
        } else if (type === 'combo') {
          const comboOptions = parseComboOptionsInput(addControlComboOptionsInput?.value || '');
          if (!comboOptions.length) {
            if (addControlError) addControlError.textContent = 'Combo controls require at least one option.';
            addControlComboOptionsInput?.focus();
            return;
          }
          config.comboOptions = comboOptions;
        }
        if (addControlError) addControlError.textContent = '';
        if (addControlSubmitBtn) addControlSubmitBtn.disabled = true;
        if (typeof onAddControl === 'function') {
          await onAddControl(pendingNode, config);
        }
        close();
      } catch (err) {
        if (addControlError) addControlError.textContent = err?.message || 'Unable to add control.';
      } finally {
        if (addControlSubmitBtn) addControlSubmitBtn.disabled = false;
      }
    });
  }

  return { open, close };
}
