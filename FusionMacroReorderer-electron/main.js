const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

  let mainWindow = null;
  let explorerWindow = null;
  const drfxPresetTempMap = new Map();

  const PROTOCOL_NAME = 'macromachine';

  let protocolQueue = [];
  let rendererReady = false;

  function flushProtocolQueue() {
    if (!mainWindow || mainWindow.isDestroyed() || !rendererReady) return;
    const queued = protocolQueue.slice();
    protocolQueue = [];
    queued.forEach((url) => {
      try {
        mainWindow.webContents.send('fmr-protocol', { url });
      } catch (_) {}
    });
  }

  function handleProtocolUrl(url) {
    if (!url) return;
    try {
      if (!mainWindow || mainWindow.isDestroyed()) {
        protocolQueue.push(url);
        return;
      }
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
      if (!rendererReady) {
        protocolQueue.push(url);
        return;
      }
      mainWindow.webContents.send('fmr-protocol', { url });
    } catch (_) {}
  }

  function findProtocolArg(argv) {
    const prefix = `${PROTOCOL_NAME}://`;
    const arg = (argv || []).find(a => typeof a === 'string' && a.toLowerCase().startsWith(prefix));
    return arg || '';
  }

  const gotLock = app.requestSingleInstanceLock();
  if (!gotLock) {
    app.quit();
  } else {
    app.on('second-instance', (_event, argv) => {
      const url = findProtocolArg(argv);
      if (url) handleProtocolUrl(url);
      if (mainWindow && !mainWindow.isDestroyed()) {
        if (mainWindow.isMinimized()) mainWindow.restore();
        mainWindow.focus();
      }
    });
  }
let updateCheckInFlight = false;
let lastUpdateCheckManual = false;
let autoUpdaterReady = false;

function showUpdateMessage(options) {
  const win = mainWindow && !mainWindow.isDestroyed() ? mainWindow : null;
  return dialog.showMessageBox(win || undefined, options);
}

function showAboutDialog() {
  const version = app.getVersion();
  return showUpdateMessage({
    type: 'info',
    title: 'About Macro Machine',
    message: 'Macro Machine',
    detail: `Version ${version}`,
  });
}

function setupAutoUpdater() {
  if (autoUpdaterReady) return;
  autoUpdaterReady = true;
  if (!autoUpdater) return;
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.on('error', (err) => {
    updateCheckInFlight = false;
    if (!lastUpdateCheckManual) return;
    showUpdateMessage({
      type: 'error',
      title: 'Update Error',
      message: 'Unable to check for updates.',
      detail: err?.message || String(err),
    });
  });
  autoUpdater.on('update-available', (info) => {
    updateCheckInFlight = false;
    showUpdateMessage({
      type: 'info',
      title: 'Update Available',
      message: `Version ${info?.version || ''} is available.`,
      detail: 'Download the update now?',
      buttons: ['Download', 'Later'],
      defaultId: 0,
      cancelId: 1,
    }).then((res) => {
      if (res.response === 0) autoUpdater.downloadUpdate();
    });
  });
  autoUpdater.on('update-not-available', () => {
    updateCheckInFlight = false;
    if (!lastUpdateCheckManual) return;
    showUpdateMessage({
      type: 'info',
      title: 'No Updates',
      message: 'You are already on the latest version.',
    });
  });
  autoUpdater.on('update-downloaded', () => {
    updateCheckInFlight = false;
    showUpdateMessage({
      type: 'info',
      title: 'Update Ready',
      message: 'The update has been downloaded.',
      detail: 'Install and restart now?',
      buttons: ['Install and Restart', 'Later'],
      defaultId: 0,
      cancelId: 1,
    }).then((res) => {
      if (res.response === 0) autoUpdater.quitAndInstall();
    });
  });
}

function runUpdateCheck(source = 'manual') {
  setupAutoUpdater();
  if (!app.isPackaged) {
    if (source === 'manual') {
      showUpdateMessage({
        type: 'info',
        title: 'Updates Unavailable',
        message: 'Updates are available in the packaged app.',
      });
    }
    return;
  }
  if (updateCheckInFlight) return;
  updateCheckInFlight = true;
  lastUpdateCheckManual = source === 'manual';
  Promise.resolve(autoUpdater.checkForUpdates())
    .catch((err) => {
      updateCheckInFlight = false;
      if (!lastUpdateCheckManual) return;
      showUpdateMessage({
        type: 'error',
        title: 'Update Error',
        message: 'Unable to check for updates.',
        detail: err?.message || String(err),
      });
    });
}

function createMainWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      // Simpler model for this local tool: allow Node APIs in the renderer
      // and avoid using a preload script (which has been causing load issues).
      contextIsolation: false,
      nodeIntegration: true,
      nodeIntegrationInSubFrames: true,
    },
  });

  // Load the web app: prefer the sibling FusionMacroReorderer folder in dev,
  // fall back to a bundled copy inside the app for packaged builds.
  const devIndexPath = path.join(__dirname, '..', 'FusionMacroReorderer', 'index.html');
  const prodIndexPath = path.join(__dirname, 'FusionMacroReorderer', 'index.html');
  const indexPath = fs.existsSync(devIndexPath) ? devIndexPath : prodIndexPath;
  win.loadFile(indexPath);
  win.webContents.on('did-finish-load', () => {
    rendererReady = true;
    flushProtocolQueue();
  });
  win.on('closed', () => {
    rendererReady = false;
  });
  mainWindow = win;
}

function createExplorerWindow() {
  if (explorerWindow) {
    explorerWindow.focus();
    return;
  }
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: true,
    },
  });
  const devPath = path.join(__dirname, '..', 'FusionMacroReorderer', 'explorer.html');
  const prodPath = path.join(__dirname, 'FusionMacroReorderer', 'explorer.html');
  const explorerPath = fs.existsSync(devPath) ? devPath : prodPath;
  win.loadFile(explorerPath);
  win.on('closed', () => {
    explorerWindow = null;
  });
  explorerWindow = win;
}

function resolveZipModule(name) {
  try {
    return require(name);
  } catch (_) {}
  try {
    if (process.resourcesPath) {
      const candidate = path.join(process.resourcesPath, 'FusionMacroReorderer', 'node_modules', name);
      return require(candidate);
    }
  } catch (_) {}
  return null;
}

let yauzl = null;
let archiver = null;

function ensureZipLibs() {
  if (!yauzl) {
    try {
      yauzl = require('yauzl');
    } catch (_) {
      yauzl = resolveZipModule('yauzl');
    }
  }
  if (!archiver) {
    try {
      archiver = require('archiver');
    } catch (_) {
      archiver = resolveZipModule('archiver');
    }
  }
  return { yauzl, archiver };
}

function stripDisabledSuffix(value) {
  if (!value) return value;
  const text = String(value);
  if (text.toLowerCase().endsWith('.disabled')) {
    return text.slice(0, -'.disabled'.length);
  }
  return text;
}

function isDisabledPath(value) {
  if (!value) return false;
  return String(value).toLowerCase().endsWith('.disabled');
}

function getDrfxPaths() {
  const base = app.getPath('userData');
  return {
    metaPath: path.join(base, 'macro-explorer', 'metadata.json'),
    vaultDir: path.join(base, 'macro-explorer', 'vault'),
  };
}

function ensureMetadataDir() {
  const { metaPath } = getDrfxPaths();
  fs.mkdirSync(path.dirname(metaPath), { recursive: true });
}

function loadMetadata() {
  const { metaPath } = getDrfxPaths();
  if (!fs.existsSync(metaPath)) return {};
  try {
    return JSON.parse(fs.readFileSync(metaPath, 'utf8'));
  } catch (_) {
    return {};
  }
}

function saveMetadata(data) {
  ensureMetadataDir();
  const { metaPath } = getDrfxPaths();
  fs.writeFileSync(metaPath, JSON.stringify(data, null, 2));
}

function metadataKey(drfxPath) {
  try {
    return path.resolve(stripDisabledSuffix(drfxPath));
  } catch (_) {
    return stripDisabledSuffix(drfxPath);
  }
}

function getDisabledPresets(drfxPath) {
  const data = loadMetadata();
  const presets = data.drfx_presets || {};
  return presets[metadataKey(drfxPath)] || [];
}

function setDisabledPresets(drfxPath, list) {
  const data = loadMetadata();
  const presets = data.drfx_presets || {};
  if (list.length) {
    presets[metadataKey(drfxPath)] = list;
  } else {
    delete presets[metadataKey(drfxPath)];
  }
  data.drfx_presets = presets;
  saveMetadata(data);
}

function vaultFileName(drfxPath) {
  const base = path.basename(stripDisabledSuffix(drfxPath));
  const hash = crypto.createHash('sha1').update(metadataKey(drfxPath)).digest('hex').slice(0, 8);
  return `${base}.${hash}`;
}

function storeInVault(drfxPath, force = false) {
  const { vaultDir } = getDrfxPaths();
  fs.mkdirSync(vaultDir, { recursive: true });
  const vaultPath = path.join(vaultDir, vaultFileName(drfxPath));
  if (force || !fs.existsSync(vaultPath)) {
    fs.copyFileSync(drfxPath, vaultPath);
  }
  return vaultPath;
}

function getVaultPath(drfxPath) {
  const { vaultDir } = getDrfxPaths();
  return path.join(vaultDir, vaultFileName(drfxPath));
}

function buildPresetSkipSet(disabledList) {
  const skip = new Set();
  disabledList.forEach((name) => {
    const lower = String(name || '').toLowerCase();
    skip.add(lower);
    if (lower.endsWith('.setting')) {
      skip.add(lower.replace(/\.setting$/i, '.png'));
    }
  });
  return skip;
}

function readZipEntries(zipPath, skipSet) {
  const libs = ensureZipLibs();
  if (!libs.yauzl) {
    return Promise.reject(new Error('DRFX support unavailable (missing yauzl).'));
  }
  return new Promise((resolve, reject) => {
    const entries = [];
    libs.yauzl.open(zipPath, { lazyEntries: true }, (err, zip) => {
      if (err) return reject(err);
      zip.readEntry();
      zip.on('entry', (entry) => {
        const name = entry.fileName;
        const lower = name.toLowerCase();
        if (name.endsWith('/')) {
          zip.readEntry();
          return;
        }
        if (skipSet && skipSet.has(lower)) {
          zip.readEntry();
          return;
        }
        zip.openReadStream(entry, (streamErr, stream) => {
          if (streamErr || !stream) {
            zip.readEntry();
            return;
          }
          const chunks = [];
          stream.on('data', (chunk) => chunks.push(chunk));
          stream.on('end', () => {
            entries.push({
              name,
              data: Buffer.concat(chunks),
              date: entry.getLastModDate ? entry.getLastModDate() : null,
            });
            zip.readEntry();
          });
          stream.on('error', () => {
            zip.readEntry();
          });
        });
      });
      zip.on('end', () => resolve(entries));
      zip.on('error', reject);
    });
  });
}

function readZipEntry(zipPath, entryName) {
  const libs = ensureZipLibs();
  if (!libs.yauzl) {
    return Promise.reject(new Error('DRFX support unavailable (missing yauzl).'));
  }
  return new Promise((resolve, reject) => {
    let resolved = false;
    const target = String(entryName || '');
    const targetLower = target.toLowerCase();
    libs.yauzl.open(zipPath, { lazyEntries: true }, (err, zip) => {
      if (err) return reject(err);
      zip.readEntry();
      zip.on('entry', (entry) => {
        if (resolved) return;
        const name = entry.fileName;
        const lower = name.toLowerCase();
        if (name.endsWith('/')) {
          zip.readEntry();
          return;
        }
        if (lower !== targetLower) {
          zip.readEntry();
          return;
        }
        zip.openReadStream(entry, (streamErr, stream) => {
          if (streamErr || !stream) {
            zip.readEntry();
            return;
          }
          const chunks = [];
          stream.on('data', (chunk) => chunks.push(chunk));
          stream.on('end', () => {
            resolved = true;
            try {
              zip.close();
            } catch (_) {}
            resolve({ name, data: Buffer.concat(chunks) });
          });
          stream.on('error', (streamError) => {
            reject(streamError);
          });
        });
      });
      zip.on('end', () => {
        if (!resolved) resolve(null);
      });
      zip.on('error', reject);
    });
  });
}

function readFirstPng(zipPath) {
  const libs = ensureZipLibs();
  if (!libs.yauzl) {
    return Promise.reject(new Error('DRFX support unavailable (missing yauzl).'));
  }
  return new Promise((resolve, reject) => {
    let resolved = false;
    libs.yauzl.open(zipPath, { lazyEntries: true }, (err, zip) => {
      if (err) return reject(err);
      zip.readEntry();
      zip.on('entry', (entry) => {
        if (resolved) return;
        const name = entry.fileName;
        const lower = name.toLowerCase();
        if (name.endsWith('/')) {
          zip.readEntry();
          return;
        }
        const base = path.basename(name);
        if (!lower.endsWith('.png') || base.startsWith('._')) {
          zip.readEntry();
          return;
        }
        zip.openReadStream(entry, (streamErr, stream) => {
          if (streamErr || !stream) {
            zip.readEntry();
            return;
          }
          const chunks = [];
          stream.on('data', (chunk) => chunks.push(chunk));
          stream.on('end', () => {
            resolved = true;
            try { zip.close(); } catch (_) {}
            resolve(Buffer.concat(chunks).toString('base64'));
          });
          stream.on('error', (streamError) => {
            reject(streamError);
          });
        });
      });
      zip.on('end', () => {
        if (!resolved) resolve('');
      });
      zip.on('error', reject);
    });
  });
}

function writeZipEntries(destPath, entries) {
  const libs = ensureZipLibs();
  if (!libs.archiver) {
    return Promise.reject(new Error('DRFX support unavailable (missing archiver).'));
  }
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destPath);
    const archive = libs.archiver('zip', { zlib: { level: 9 } });
    output.on('close', () => resolve());
    output.on('error', reject);
    archive.on('error', reject);
    archive.pipe(output);
    entries.forEach((entry) => {
      archive.append(entry.data, { name: entry.name, date: entry.date || undefined });
    });
    archive.finalize();
  });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function replaceFileFromTemp(tempPath, targetPath) {
  const attempts = 3;
  let lastErr = null;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      fs.renameSync(tempPath, targetPath);
      return;
    } catch (err) {
      lastErr = err;
      try {
        if (fs.existsSync(targetPath)) {
          try { fs.chmodSync(targetPath, 0o666); } catch (_) {}
        }
        fs.copyFileSync(tempPath, targetPath);
        fs.unlinkSync(tempPath);
        return;
      } catch (copyErr) {
        lastErr = copyErr;
      }
      const code = lastErr && lastErr.code ? String(lastErr.code) : '';
      if (attempt < attempts && (code === 'EPERM' || code === 'EACCES' || code === 'EBUSY')) {
        await sleep(200 * attempt);
        continue;
      }
    }
  }
  const detail = lastErr?.message || String(lastErr || 'Unable to replace DRFX file.');
  throw new Error(`${detail}. Close DaVinci Resolve (or any app using this DRFX pack) and try again.`);
}

async function rebuildActiveFromVault(drfxPath, vaultPath, disabledList) {
  const tempPath = drfxPath + '.tmp';
  if (!disabledList.length) {
    fs.copyFileSync(vaultPath, tempPath);
    await replaceFileFromTemp(tempPath, drfxPath);
    return;
  }
  const skip = buildPresetSkipSet(disabledList);
  const keptEntries = await readZipEntries(vaultPath, skip);
  await writeZipEntries(tempPath, keptEntries);
  await replaceFileFromTemp(tempPath, drfxPath);
}

function buildPresetList(entries, disabledList) {
  const disabledSet = new Set(disabledList.map((name) => String(name || '').toLowerCase()));
  const pngMap = new Map();
  entries.forEach((entry) => {
    const lower = entry.name.toLowerCase();
    if (lower.endsWith('.png')) {
      pngMap.set(lower, entry.data);
    }
  });
  return entries
    .filter((entry) => entry.name.toLowerCase().endsWith('.setting'))
    .map((entry) => {
      const lower = entry.name.toLowerCase();
      const pngKey = lower.replace(/\.setting$/i, '.png');
      const pngData = pngMap.get(pngKey) || null;
      const baseName = path.basename(entry.name);
      const displayName = baseName.replace(/\.setting$/i, '');
      return {
        entryName: entry.name,
        displayName,
        thumbBase64: pngData ? pngData.toString('base64') : '',
        disabled: disabledSet.has(lower),
      };
  });
}

function presetThumbnailEntry(presetName) {
  const raw = String(presetName || '').replace(/\\/g, '/');
  const ext = path.posix.extname(raw);
  const base = ext ? raw.slice(0, -ext.length) : raw;
  return `${base}.png`;
}

function sanitizeFileSegment(value) {
  return String(value || '')
    .replace(/[<>:"/\\|?*]+/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeDrfxCategoryPath(value) {
  const raw = String(value || '').replace(/\\/g, '/').trim();
  const parts = raw.split('/').map(p => p.trim()).filter(p => p);
  return parts.join('/');
}

function normalizeTempKey(filePath) {
  if (!filePath) return '';
  try {
    return path.resolve(String(filePath));
  } catch (_) {
    return String(filePath);
  }
}

function buildDrfxPresetTempPath(drfxPath, presetName) {
  const basePath = stripDisabledSuffix(drfxPath);
  const drfxBase = path.basename(basePath, path.extname(basePath));
  const presetBase = path.posix.basename(String(presetName || '')).replace(/\.setting$/i, '');
  const hash = crypto
    .createHash('sha1')
    .update(`${metadataKey(drfxPath)}::${presetName}`)
    .digest('hex')
    .slice(0, 8);
  const fileName = `${sanitizeFileSegment(drfxBase)}_${sanitizeFileSegment(presetBase)}_${hash}.setting`;
  const tempDir = path.join(app.getPath('userData'), 'macro-explorer', 'temp');
  fs.mkdirSync(tempDir, { recursive: true });
  return path.join(tempDir, fileName);
}

function buildPresetRemovalSet(presetNames) {
  const presetSet = new Set();
  const skipSet = new Set();
  presetNames.forEach((name) => {
    const raw = String(name || '').replace(/\\/g, '/');
    if (!raw) return;
    const lower = raw.toLowerCase();
    presetSet.add(lower);
    skipSet.add(lower);
    skipSet.add(presetThumbnailEntry(raw).toLowerCase());
  });
  return { presetSet, skipSet };
}

// IPC: save .setting file via native dialog
ipcMain.handle('save-setting-file', async (event, payload = {}) => {
  const browserWindow = BrowserWindow.getFocusedWindow();
  const defaultPath = payload.defaultPath || 'macro.setting';
  const content = String(payload.content || '');

  const result = await dialog.showSaveDialog(browserWindow, {
    defaultPath,
    filters: [
      { name: 'Fusion Setting', extensions: ['setting', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });

  if (result.canceled || !result.filePath) return { canceled: true };

  try {
    fs.writeFileSync(result.filePath, content, 'utf8');
    return { canceled: false, filePath: result.filePath };
  } catch (err) {
    return { canceled: false, filePath: result.filePath, error: String(err && err.message || err) };
  }
});

// IPC: pick a save path without writing content
ipcMain.handle('pick-save-path', async (event, payload = {}) => {
  const browserWindow = BrowserWindow.getFocusedWindow();
  const defaultPath = payload.defaultPath || 'macro.setting';
  const result = await dialog.showSaveDialog(browserWindow, {
    defaultPath,
    filters: [
      { name: 'Fusion Setting', extensions: ['setting', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  return { canceled: false, filePath: result.filePath };
});

// IPC: write a .setting file to a specific path
ipcMain.handle('write-setting-file', async (event, payload = {}) => {
  const filePath = payload.filePath;
  const content = String(payload.content || '');
  if (!filePath) return { canceled: false, error: 'No file path provided.' };
  try {
    fs.writeFileSync(filePath, content, 'utf8');
    return { canceled: false, filePath };
  } catch (err) {
    return { canceled: false, filePath, error: String(err && err.message || err) };
  }
});

// IPC: read a .setting file from a specific path
ipcMain.handle('read-setting-file', async (event, payload = {}) => {
  const filePath = payload.filePath;
  if (!filePath) return { ok: false, error: 'No file path provided.' };
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    return { ok: true, filePath, baseName: path.basename(filePath), content };
  } catch (err) {
    return { ok: false, filePath, error: String(err && err.message || err) };
  }
});

// IPC: open a .setting file via native dialog
ipcMain.handle('open-setting-file', async () => {
  const browserWindow = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(browserWindow, {
    filters: [
      { name: 'Fusion Setting', extensions: ['setting', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
    properties: ['openFile'],
  });

  if (result.canceled || !result.filePaths || !result.filePaths[0]) {
    return { canceled: true };
  }

  const filePath = result.filePaths[0];
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    return {
      canceled: false,
      filePath,
      baseName: path.basename(filePath),
      content,
    };
  } catch (err) {
    return {
      canceled: false,
      filePath,
      error: String((err && err.message) || err),
    };
  }
});

// IPC: pick an image file via native dialog (used by header image UI)
ipcMain.handle('pick-image-file', async (_event, payload = {}) => {
  const browserWindow = BrowserWindow.getFocusedWindow();
  const startPath = payload.defaultPath || app.getPath('pictures');
  const result = await dialog.showOpenDialog(browserWindow, {
    defaultPath: startPath,
    filters: [
      { name: 'Images', extensions: ['png', 'jpg', 'jpeg'] },
      { name: 'All Files', extensions: ['*'] },
    ],
    properties: ['openFile'],
  });

  if (result.canceled || !result.filePaths || !result.filePaths[0]) {
    return { canceled: true };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

// IPC: read an image file and return a data URI for Fusion header labels.
ipcMain.handle('read-image-data-uri', async (_event, payload = {}) => {
  try {
    const filePath = String(payload.filePath || '').trim();
    if (!filePath) return { ok: false, error: 'Missing image path.' };
    const ext = path.extname(filePath).toLowerCase();
    let mime = '';
    if (ext === '.png') mime = 'image/png';
    else if (ext === '.jpg' || ext === '.jpeg') mime = 'image/jpeg';
    else return { ok: false, error: 'Unsupported image format. Use PNG or JPG.' };
    const bytes = fs.readFileSync(filePath);
    const base64 = bytes.toString('base64');
    return { ok: true, dataUri: `data:${mime};base64,${base64}` };
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
});

// IPC: select a folder via native dialog (used by Macro Explorer)
ipcMain.handle('select-folder', async (_event, payload = {}) => {
  const browserWindow = BrowserWindow.getFocusedWindow();
  const defaultPath = payload.defaultPath || app.getPath('home');
  const result = await dialog.showOpenDialog(browserWindow, {
    defaultPath,
    properties: ['openDirectory'],
  });
  if (result.canceled || !result.filePaths || !result.filePaths[0]) {
    return { canceled: true };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

// IPC: open a .setting file path inside the main Macro Machine window
ipcMain.handle('open-in-macro-machine', async (_event, payload = {}) => {
  const filePath = payload.path;
  if (!filePath || !mainWindow || mainWindow.isDestroyed()) {
    return { ok: false, error: 'Macro Machine window not available.' };
  }
  mainWindow.webContents.send('fmr-open-path', { path: filePath });
  return { ok: true };
});

ipcMain.handle('open-explorer-window', async () => {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return { ok: false, error: 'Macro Machine window not available.' };
  }
  createExplorerWindow();
  return { ok: true };
});

ipcMain.handle('drfx-list', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  if (!drfxPath) return { ok: false, error: 'Missing DRFX path.' };
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  try {
    const disabled = isDisabledPath(drfxPath);
    const disabledList = getDisabledPresets(drfxPath);
    const vaultPath = storeInVault(drfxPath, false);
    const entries = await readZipEntries(vaultPath);
    const presets = buildPresetList(entries, disabledList);
    return {
      ok: true,
      disabled,
      disabledList,
      presets,
    };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-toggle-preset', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  const presetName = payload.presetName;
  if (!drfxPath || !presetName) return { ok: false, error: 'Missing preset details.' };
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  if (isDisabledPath(drfxPath)) {
    return { ok: false, error: 'DRFX pack is disabled.' };
  }
  try {
    const disabledList = getDisabledPresets(drfxPath);
    const lower = String(presetName).toLowerCase();
    const index = disabledList.findIndex((name) => String(name).toLowerCase() === lower);
    if (index >= 0) {
      disabledList.splice(index, 1);
    } else {
      disabledList.push(presetName);
    }
    setDisabledPresets(drfxPath, disabledList);
    const vaultPath = storeInVault(drfxPath, false);
    await rebuildActiveFromVault(drfxPath, vaultPath, disabledList);
    return { ok: true, disabledList };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-set-disabled-list', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  const disabledList = Array.isArray(payload.disabledList) ? payload.disabledList : [];
  if (!drfxPath) return { ok: false, error: 'Missing DRFX path.' };
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  if (isDisabledPath(drfxPath)) {
    return { ok: false, error: 'DRFX pack is disabled.' };
  }
  try {
    setDisabledPresets(drfxPath, disabledList);
    const vaultPath = storeInVault(drfxPath, false);
    await rebuildActiveFromVault(drfxPath, vaultPath, disabledList);
    return { ok: true, disabledList };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-thumbnail', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  if (!drfxPath) return { ok: false, error: 'Missing DRFX path.' };
  const libs = ensureZipLibs();
  if (!libs.yauzl) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  try {
    const basePath = stripDisabledSuffix(drfxPath);
    const sourcePath = fs.existsSync(basePath) ? basePath : drfxPath;
    const thumbBase64 = await readFirstPng(sourcePath);
    return { ok: true, thumbBase64: thumbBase64 || '' };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-export-pack', async (_event, payload = {}) => {
  const drfxName = payload.drfxName;
  const categoryPath = payload.categoryPath;
  const presets = Array.isArray(payload.presets) ? payload.presets : [];
  if (!drfxName || !presets.length) {
    return { ok: false, error: 'Missing DRFX export data.' };
  }
  const libs = ensureZipLibs();
  if (!libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  const normalizedCategory = normalizeDrfxCategoryPath(categoryPath);
  if (!normalizedCategory) {
    return { ok: false, error: 'Category path is required.' };
  }
  const browserWindow = BrowserWindow.getFocusedWindow();
  const baseName = sanitizeFileSegment(drfxName) || 'Macro Machine Export';
  const defaultPath = `${baseName}.drfx`;
  const result = await dialog.showSaveDialog(browserWindow, {
    defaultPath,
    filters: [
      { name: 'DRFX Pack', extensions: ['drfx'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });
  if (result.canceled || !result.filePath) return { ok: false, canceled: true };
  let filePath = result.filePath;
  if (!filePath.toLowerCase().endsWith('.drfx')) {
    filePath += '.drfx';
  }
  try {
    const entries = presets.map((preset) => {
      const safeName = sanitizeFileSegment(preset.name) || 'Preset';
      const entryName = path.posix.join(normalizedCategory, `${safeName}.setting`);
      return {
        name: entryName,
        data: Buffer.from(String(preset.content || ''), 'utf8'),
      };
    });
    const tempPath = filePath + '.tmp';
    await writeZipEntries(tempPath, entries);
    await replaceFileFromTemp(tempPath, filePath);
    return { ok: true, filePath };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-extract-preset', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  const presetName = payload.presetName;
  if (!drfxPath || !presetName) {
    return { ok: false, error: 'Missing preset details.' };
  }
  const libs = ensureZipLibs();
  if (!libs.yauzl) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  try {
    const vaultPath = storeInVault(drfxPath, false);
    const entry = await readZipEntry(vaultPath, presetName);
    if (!entry || !entry.data) {
      return { ok: false, error: 'Preset not found in DRFX pack.' };
    }
    const filePath = buildDrfxPresetTempPath(drfxPath, entry.name || presetName);
    fs.writeFileSync(filePath, entry.data);
    const key = normalizeTempKey(filePath);
    if (key) {
      drfxPresetTempMap.set(key, {
        drfxPath,
        presetName: entry.name || presetName,
      });
    }
    return { ok: true, filePath };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-get-link', async (_event, payload = {}) => {
  const sourcePath = payload.path;
  if (!sourcePath) return { ok: true, linked: false };
  const key = normalizeTempKey(sourcePath);
  const mapping = key ? drfxPresetTempMap.get(key) : null;
  if (!mapping) return { ok: true, linked: false };
  return { ok: true, linked: true, drfxPath: mapping.drfxPath, presetName: mapping.presetName };
});

ipcMain.handle('drfx-save-preset', async (_event, payload = {}) => {
  const sourcePath = payload.sourcePath;
  const content = payload.content;
  if (!sourcePath || typeof content !== 'string') {
    return { ok: false, error: 'Missing preset content.' };
  }
  const key = normalizeTempKey(sourcePath);
  const mapping = key ? drfxPresetTempMap.get(key) : null;
  if (!mapping) {
    return { ok: false, code: 'not-linked', error: 'Not a DRFX preset file.' };
  }
  const drfxPath = mapping.drfxPath;
  const presetName = mapping.presetName;
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  if (isDisabledPath(drfxPath)) {
    return { ok: false, code: 'disabled', error: 'DRFX pack is disabled.' };
  }
  try {
    const vaultPath = storeInVault(drfxPath, false);
    const entries = await readZipEntries(vaultPath);
    const targetLower = String(presetName || '').toLowerCase();
    let replaced = false;
    const nextEntries = entries.map((entry) => {
      if (entry.name.toLowerCase() !== targetLower) return entry;
      replaced = true;
      return {
        ...entry,
        data: Buffer.from(content, 'utf8'),
        date: new Date(),
      };
    });
    if (!replaced) {
      nextEntries.push({
        name: presetName,
        data: Buffer.from(content, 'utf8'),
        date: new Date(),
      });
    }
    const tempVault = vaultPath + '.tmp';
    await writeZipEntries(tempVault, nextEntries);
    fs.renameSync(tempVault, vaultPath);
    const disabledList = getDisabledPresets(drfxPath);
    await rebuildActiveFromVault(drfxPath, vaultPath, disabledList);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-set-thumbnail', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  const presetName = payload.presetName;
  const imagePath = payload.imagePath;
  if (!drfxPath || !presetName || !imagePath) {
    return { ok: false, error: 'Missing thumbnail details.' };
  }
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  if (isDisabledPath(drfxPath)) {
    return { ok: false, error: 'DRFX pack is disabled.' };
  }
  if (path.extname(imagePath).toLowerCase() !== '.png') {
    return { ok: false, error: 'Thumbnail must be a .png image.' };
  }
  if (!fs.existsSync(imagePath)) {
    return { ok: false, error: 'Thumbnail file not found.' };
  }
  try {
    const data = fs.readFileSync(imagePath);
    const targetName = presetThumbnailEntry(presetName);
    const targetLower = targetName.toLowerCase();
    const vaultPath = storeInVault(drfxPath, false);
    const entries = await readZipEntries(vaultPath);
    let existingDate = null;
    const nextEntries = entries.filter((entry) => {
      if (entry.name.toLowerCase() === targetLower) {
        existingDate = entry.date || null;
        return false;
      }
      return true;
    });
    nextEntries.push({
      name: targetName,
      data,
      date: existingDate || new Date(),
    });
    const tempVault = vaultPath + '.tmp';
    await writeZipEntries(tempVault, nextEntries);
    fs.renameSync(tempVault, vaultPath);
    const disabledList = getDisabledPresets(drfxPath);
    await rebuildActiveFromVault(drfxPath, vaultPath, disabledList);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-delete-pack', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  if (!drfxPath) return { ok: false, error: 'Missing DRFX path.' };
  if (!fs.existsSync(drfxPath)) {
    return { ok: false, error: 'DRFX pack not found.' };
  }
  try {
    const key = metadataKey(drfxPath);
    const data = loadMetadata();
    const presets = data.drfx_presets || {};
    if (presets[key]) {
      delete presets[key];
      data.drfx_presets = presets;
      saveMetadata(data);
    }
    const vaultPath = getVaultPath(drfxPath);
    if (fs.existsSync(vaultPath)) {
      fs.unlinkSync(vaultPath);
    }
    fs.unlinkSync(drfxPath);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('drfx-delete-presets', async (_event, payload = {}) => {
  const drfxPath = payload.path;
  const presetNames = Array.isArray(payload.presetNames) ? payload.presetNames : [];
  if (!drfxPath || !presetNames.length) {
    return { ok: false, error: 'Missing preset details.' };
  }
  const libs = ensureZipLibs();
  if (!libs.yauzl || !libs.archiver) {
    return { ok: false, error: 'DRFX support unavailable (zip libraries missing).' };
  }
  if (!fs.existsSync(drfxPath)) {
    return { ok: false, error: 'DRFX pack not found.' };
  }
  try {
    const { presetSet, skipSet } = buildPresetRemovalSet(presetNames);
    const disabledList = getDisabledPresets(drfxPath).filter(
      (name) => !presetSet.has(String(name || '').toLowerCase())
    );
    setDisabledPresets(drfxPath, disabledList);
    const vaultPath = storeInVault(drfxPath, false);
    const entries = await readZipEntries(vaultPath);
    const nextEntries = entries.filter((entry) => !skipSet.has(entry.name.toLowerCase()));
    const tempVault = vaultPath + '.tmp';
    await writeZipEntries(tempVault, nextEntries);
    fs.renameSync(tempVault, vaultPath);
    await rebuildActiveFromVault(drfxPath, vaultPath, disabledList);
    return { ok: true, disabledList };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

ipcMain.handle('set-export-folder', async (_event, payload = {}) => {
  const folderPath = payload.path;
  if (!folderPath || !mainWindow || mainWindow.isDestroyed()) {
    return { ok: false, error: 'Macro Machine window not available.' };
  }
  mainWindow.webContents.send('fmr-set-export-folder', { path: folderPath });
  return { ok: true };
});

ipcMain.handle('set-data-menu-state', async (_event, payload = {}) => {
  try {
    const menu = Menu.getApplicationMenu();
    const item = menu?.getMenuItemById('dataReloadCsv');
    if (item) item.enabled = !!payload.reloadEnabled;
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err?.message || String(err) };
  }
});

  ipcMain.handle('get-user-data-path', async () => {
    return { path: app.getPath('userData') };
  });

  app.on('open-url', (event, url) => {
    event.preventDefault();
    handleProtocolUrl(url);
  });

  app.whenReady().then(() => {
    createMainWindow();
    try {
      app.setAsDefaultProtocolClient(PROTOCOL_NAME);
    } catch (_) {}
    const initialUrl = findProtocolArg(process.argv || []);
    if (initialUrl) {
      setTimeout(() => handleProtocolUrl(initialUrl), 300);
    }
    // Application menu with basic file actions wired into the renderer
    const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'Open...',
          accelerator: 'CmdOrCtrl+O',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'open' });
          },
        },
        {
          label: 'Save',
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'save' });
          },
        },
        { type: 'separator' },
        {
          label: 'Import from Clipboard',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'importClipboard' });
          },
        },
        {
          label: 'Export to Clipboard',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'exportClipboard' });
          },
        },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        {
          label: 'Normalize legacy names...',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'normalizeLegacy' });
          },
        },
        {
          label: 'Header Image...',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'headerImage' });
          },
        },
        { type: 'separator' },
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'pasteAndMatchStyle' },
        { role: 'delete' },
        { role: 'selectAll' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'toggledevtools' },
        { type: 'separator' },
        {
          label: 'Toggle Diagnostics',
          accelerator: 'CmdOrCtrl+D',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'toggleDiagnostics' });
          },
        },
      ],
    },
      {
        label: 'Data',
        submenu: [
        {
          label: 'Import CSV (File)...',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'csvImportFile' });
          },
        },
        {
          label: 'Import CSV (URL)...',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'csvImportUrl' });
          },
        },
        {
          label: 'Import Google Sheet (Public URL)...',
          click: () => {
            const win = BrowserWindow.getFocusedWindow();
            if (win) win.webContents.send('fmr-menu', { action: 'csvImportSheet' });
          },
        },
          {
            label: 'Reload Sheet',
            id: 'dataReloadCsv',
            enabled: false,
            click: () => {
              const win = BrowserWindow.getFocusedWindow();
              if (win) win.webContents.send('fmr-menu', { action: 'csvReload' });
            },
          },
          { type: 'separator' },
          {
            label: 'Generate from CSV...',
            click: () => {
              const win = BrowserWindow.getFocusedWindow();
              if (win) win.webContents.send('fmr-menu', { action: 'csvGenerate' });
            },
          },
          { type: 'separator' },
          {
            label: 'Insert Update Data Button',
            click: () => {
              const win = BrowserWindow.getFocusedWindow();
              if (win) win.webContents.send('fmr-menu', { action: 'insertUpdateData' });
            },
          },
        ],
      },
    {
      label: 'Help',
      submenu: [
        {
          label: 'Check for Updates...',
          click: () => runUpdateCheck('manual'),
        },
        {
          label: 'About Macro Machine',
          click: () => showAboutDialog(),
        },
      ],
    },
  ];
  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);

  setTimeout(() => runUpdateCheck('auto'), 2000);

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

