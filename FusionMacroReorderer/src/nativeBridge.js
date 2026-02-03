export function detectElectron() {
  try {
    return typeof process !== 'undefined' &&
      !!(process.versions && process.versions.electron);
  } catch (_) {
    return false;
  }
}

export function setupNativeBridge({
  isElectron,
  exportBtn,
  importClipboardBtn,
  exportClipboardBtn,
  setDiagnosticsEnabled,
  handleNativeOpen,
  onNormalizeLegacyNames,
  onImportCsvFile,
  onImportCsvUrl,
  onGenerateFromCsv,
}) {
  if (!isElectron || typeof window === 'undefined' || window.FusionMacroReordererNative) {
    return;
  }

  try {
    const electronRequire = typeof window !== 'undefined' && typeof window.require === 'function'
      ? window.require
      : (typeof require !== 'undefined' ? require : null);
    if (!electronRequire) return null;
    const { clipboard, ipcRenderer } = electronRequire('electron');
    window.FusionMacroReordererNative = {
      isElectron: true,
      saveSettingFile: (payload) => ipcRenderer.invoke('save-setting-file', payload || {}),
      saveDrfxPreset: (payload) => ipcRenderer.invoke('drfx-save-preset', payload || {}),
      openSettingFile: () => ipcRenderer.invoke('open-setting-file'),
      readClipboard: () => {
        try { return clipboard.readText(); } catch (_) { return ''; }
      },
      writeClipboard: (text) => {
        try { clipboard.writeText(String(text || '')); } catch (_) {}
      },
    };

    try {
      ipcRenderer.on('fmr-menu', async (_event, payload = {}) => {
        const action = payload.action;
        try {
          if (action === 'open') {
            if (typeof handleNativeOpen === 'function') await handleNativeOpen();
          } else if (action === 'save') {
            if (exportBtn) exportBtn.click();
          } else if (action === 'importClipboard') {
            if (importClipboardBtn) importClipboardBtn.click();
          } else if (action === 'exportClipboard') {
            if (exportClipboardBtn) exportClipboardBtn.click();
          } else if (action === 'toggleDiagnostics') {
            if (typeof setDiagnosticsEnabled === 'function') setDiagnosticsEnabled();
          } else if (action === 'normalizeLegacy') {
            if (typeof onNormalizeLegacyNames === 'function') onNormalizeLegacyNames();
          } else if (action === 'csvImportFile') {
            if (typeof onImportCsvFile === 'function') onImportCsvFile();
          } else if (action === 'csvImportUrl') {
            if (typeof onImportCsvUrl === 'function') onImportCsvUrl();
          } else if (action === 'csvGenerate') {
            if (typeof onGenerateFromCsv === 'function') onGenerateFromCsv();
          }
        } catch (_) {
          /* ignore menu handler errors */
        }
      });
    } catch (_) {
      /* ignore ipc wiring errors */
    }

    return { ipcRenderer };
  } catch (_) {
    // If require/electron fails, leave the bridge undefined.
    return null;
  }
}
