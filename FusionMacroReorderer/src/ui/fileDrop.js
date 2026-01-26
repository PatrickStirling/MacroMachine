export function createFileDropController(options = {}) {
  const {
    dropZone,
    updateDropHint,
    logDiag,
    handleFiles,
    handleDroppedPath,
    handleDrfxFiles,
    extractFilePathFromData,
    isDrfxPath,
    isElectron,
  } = options;

  const setHint = (text) => {
    if (typeof updateDropHint === 'function') updateDropHint(text);
  };

  const log = (message) => {
    if (!logDiag) return;
    try { logDiag(message); } catch (_) {}
  };

  const hasFiles = (e) => {
    const dt = e?.dataTransfer;
    if (!dt) return false;
    return Array.from(dt.types || []).includes('Files');
  };

  const handleDropData = async (dt, label) => {
    if (!dt) return false;
    const files = dt.files;
    if (files && files.length) {
      log(`[Drop] ${label} drop (files)`);
      await handleFiles(Array.from(files));
      dropZone?.classList.remove('dragover');
      setHint('Drop .setting file here...');
      return true;
    }
    if (isElectron) {
      const uriList = dt.getData('text/uri-list') || dt.getData('text/plain');
      const filePath = extractFilePathFromData ? extractFilePathFromData(uriList) : null;
      if (filePath) {
        log(`[Drop] ${label} drop (path) ${filePath}`);
        if (isDrfxPath && isDrfxPath(filePath)) {
          await handleDrfxFiles([filePath]);
        } else {
          await handleDroppedPath(filePath);
        }
        dropZone?.classList.remove('dragover');
        setHint('Drop .setting file here...');
        return true;
      }
    }
    return false;
  };

  const bindDropZone = () => {
    if (!dropZone) return;
    ['dragenter', 'dragover'].forEach((type) => {
      dropZone.addEventListener(type, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('dragover');
        setHint('Drop .setting file here... (over panel)');
        log(`[Drop] ${type} on dropZone`);
      });
    });
    ['dragleave', 'drop'].forEach((type) => {
      dropZone.addEventListener(type, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('dragover');
        setHint('Drop .setting file here...');
        log(`[Drop] ${type} on dropZone`);
      });
    });
    dropZone.addEventListener('drop', async (e) => {
      await handleDropData(e.dataTransfer, 'dropZone');
    });
  };

  const bindGlobalDropHandlers = () => {
    let dragDepth = 0;
    window.addEventListener('dragenter', (e) => {
      if (!hasFiles(e)) return;
      dragDepth += 1;
      dropZone?.classList.add('dragover');
      setHint('Drop .setting file here... (window dragenter)');
      log('[Drop] window dragenter');
    });
    window.addEventListener('dragover', (e) => {
      if (hasFiles(e)) {
        e.preventDefault();
        setHint('Drop .setting file here... (window dragover)');
        log('[Drop] window dragover');
      }
    });
    window.addEventListener('dragleave', (e) => {
      if (!hasFiles(e)) return;
      dragDepth = Math.max(0, dragDepth - 1);
      if (dragDepth === 0 && dropZone) {
        dropZone.classList.remove('dragover');
        setHint('Drop .setting file here...');
        log('[Drop] window dragleave');
      }
    });
    window.addEventListener('drop', (e) => {
      if (hasFiles(e)) {
        e.preventDefault();
        log('[Drop] window drop');
      }
      dragDepth = 0;
      dropZone?.classList.remove('dragover');
      setHint('Drop .setting file here...');
    });
    document.addEventListener('drop', async (e) => {
      const dt = e.dataTransfer;
      if (!dt) return;
      const files = dt.files;
      if (files && files.length > 0) {
        e.preventDefault();
        e.stopPropagation();
        log('[Drop] document drop (files)');
        await handleFiles(Array.from(files));
        dropZone?.classList.remove('dragover');
        setHint('Drop .setting file here...');
        return;
      }
      await handleDropData(dt, 'document');
    });
  };

  bindDropZone();
  bindGlobalDropHandlers();

  return {
    refreshHint: (text) => setHint(text),
  };
}
