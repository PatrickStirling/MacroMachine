export function createIntroPanelController(options = {}) {
  const {
    introPanel,
    introToggleBtn,
    openExplorerBtn,
    closeExplorerPanelBtn,
    macroExplorerPanel,
    macroExplorerMini,
    dropHint,
    isElectron,
    onClearSelections,
    onHideDetailDrawer,
    onSyncDrawerAnchor,
    onInfo,
  } = options;

  const dropHintDefaultText = dropHint ? dropHint.textContent : '';

  if (!isElectron && macroExplorerMini) {
    macroExplorerMini.hidden = true;
  }

  const setIntroCollapsed = (collapsed) => {
    if (!introPanel) return;
    introPanel.hidden = collapsed;
    if (introToggleBtn) {
      introToggleBtn.setAttribute('aria-pressed', collapsed ? 'true' : 'false');
    }
    if (!collapsed) {
      if (typeof onClearSelections === 'function') onClearSelections();
      if (typeof onHideDetailDrawer === 'function') onHideDetailDrawer();
    }
    if (typeof onSyncDrawerAnchor === 'function') onSyncDrawerAnchor();
  };

  const setIntroToggleVisible = (visible) => {
    if (!introToggleBtn) return;
    introToggleBtn.hidden = !visible;
    if (!visible) {
      introToggleBtn.setAttribute('aria-pressed', 'false');
    }
  };

  const updateDropHint = (text) => {
    try {
      if (!dropHint) return;
      dropHint.textContent = text || dropHintDefaultText;
    } catch (_) {}
  };

  introToggleBtn?.addEventListener('click', () => {
    if (!introPanel) return;
    setIntroCollapsed(!introPanel.hidden);
  });

  openExplorerBtn?.addEventListener('click', () => {
    if (!isElectron) {
      if (typeof onInfo === 'function') onInfo('Macro Explorer is available in the desktop app.');
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

  return {
    setIntroCollapsed,
    setIntroToggleVisible,
    updateDropHint,
  };
}
