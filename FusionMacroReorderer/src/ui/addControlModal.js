export function createAddControlModal(options = {}) {
  const {
    state,
    addControlModal,
    addControlForm,
    addControlTitle,
    addControlNameInput,
    addControlTypeSelect,
    addLabelOptions,
    addControlLabelCountInput,
    addControlLabelDefaultSelect,
    addControlPageInput,
    addControlPageOptions,
    addControlError,
    addControlCancelBtn,
    addControlCloseBtn,
    addControlSubmitBtn,
    onAddControl,
  } = options;

  let pendingNode = null;

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
    if (!addLabelOptions) return;
    const typeVal = (addControlTypeSelect?.value || 'label').toLowerCase();
    addLabelOptions.hidden = typeVal !== 'label';
  };

  const resetAddControlFormFields = (nodeName) => {
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
  };

  const open = (nodeName) => {
    if (!addControlModal) return;
    pendingNode = nodeName;
    updateAddControlPageOptionsList();
    resetAddControlFormFields(nodeName);
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
