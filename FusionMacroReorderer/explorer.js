const { ipcRenderer, shell } = require('electron');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const yauzl = null;
const archiver = null;

const DISABLED_SUFFIX = '.disabled';

const rootSelect = document.getElementById('rootSelect');
const chooseRootBtn = document.getElementById('chooseRootBtn');
const refreshBtn = document.getElementById('refreshBtn');
const exportFolderBtn = document.getElementById('exportFolderBtn');
const currentPathEl = document.getElementById('currentPath');
const pathBackBtn = document.getElementById('pathBackBtn');
const pathForwardBtn = document.getElementById('pathForwardBtn');
const searchInput = document.getElementById('searchInput');
const clearSearchBtn = document.getElementById('clearSearchBtn');
const entriesBody = document.getElementById('entriesBody');
const statusEl = document.getElementById('status');
const bulkCountEl = document.getElementById('bulkCount');
const createFolderBtn = document.getElementById('createFolderBtn');
const setThumbBtn = document.getElementById('setThumbBtn');
const bulkEnableBtn = document.getElementById('bulkEnableBtn');
const bulkDisableBtn = document.getElementById('bulkDisableBtn');
const bulkDeleteBtn = document.getElementById('bulkDeleteBtn');
const bulkClearBtn = document.getElementById('bulkClearBtn');
const drfxModal = document.getElementById('drfxModal');
const closeDrfxBtn = document.getElementById('closeDrfxBtn');
const drfxPathEl = document.getElementById('drfxPath');
const drfxWarningEl = document.getElementById('drfxWarning');
const drfxStatusEl = document.getElementById('drfxStatus');
const drfxSearchInput = document.getElementById('drfxSearchInput');
const drfxClearSearchBtn = document.getElementById('drfxClearSearchBtn');
const drfxBulkCountEl = document.getElementById('drfxBulkCount');
const drfxSetThumbBtn = document.getElementById('drfxSetThumbBtn');
const drfxBulkEnableBtn = document.getElementById('drfxBulkEnableBtn');
const drfxBulkDisableBtn = document.getElementById('drfxBulkDisableBtn');
const drfxBulkDeleteBtn = document.getElementById('drfxBulkDeleteBtn');
const drfxBulkClearBtn = document.getElementById('drfxBulkClearBtn');
const drfxImportAllBtn = document.getElementById('drfxImportAllBtn');
const drfxListEl = document.getElementById('drfxList');
const textPromptModal = document.getElementById('textPromptModal');
const textPromptTitle = document.getElementById('textPromptTitle');
const textPromptLabel = document.getElementById('textPromptLabel');
const textPromptInput = document.getElementById('textPromptInput');
const textPromptCancelBtn = document.getElementById('textPromptCancelBtn');
const textPromptOkBtn = document.getElementById('textPromptOkBtn');
const textPromptCloseBtn = document.getElementById('textPromptCloseBtn');
const thumbFileInput = document.getElementById('thumbFileInput');
const viewParams = new URLSearchParams(window.location.search);
const isFoldersOnly = viewParams.get('mode') === 'folders';
if (isFoldersOnly && typeof document !== 'undefined') {
  document.body.classList.add('folders-only');
}

const roots = buildDefaultRoots();
let currentRootKey = null;
let currentPath = '';
let userDataPath = '';
let currentDrfxPath = '';
let currentDrfxVaultPath = '';
let currentDrfxPresets = [];
let currentDrfxDisabled = false;
let currentEntries = [];
let pathHistory = [];
let pathHistoryIndex = -1;
const drfxSelected = new Set();
const selectedEntries = new Map();
const drfxThumbCache = new Map();
let pendingThumbContext = null;
let textPromptResolver = null;

function buildDefaultRoots() {
  const base = getFusionBasePath();
  const templates = path.join(base, 'Templates');
  const macros = path.join(base, 'Macros');
  return {
    Templates: templates,
    Macros: macros,
    Custom: null,
  };
}

function getFusionBasePath() {
  if (process.platform === 'win32') {
    const appdata = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
    const support = path.join(appdata, 'Blackmagic Design', 'DaVinci Resolve', 'Support', 'Fusion');
    const legacy = path.join(appdata, 'Blackmagic Design', 'DaVinci Resolve', 'Fusion');
    return fs.existsSync(support) ? support : legacy;
  }
  if (process.platform === 'darwin') {
    return path.join(
      os.homedir(),
      'Library',
      'Application Support',
      'Blackmagic Design',
      'DaVinci Resolve',
      'Fusion'
    );
  }
  return path.join(
    os.homedir(),
    '.local',
    'share',
    'Blackmagic Design',
    'DaVinci Resolve',
    'Fusion'
  );
}

function initRootSelect() {
  rootSelect.innerHTML = '';
  ['Templates', 'Macros', 'Custom'].forEach((key) => {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = key;
    rootSelect.appendChild(opt);
  });
  currentRootKey = roots.Templates ? 'Templates' : 'Macros';
  rootSelect.value = currentRootKey || 'Templates';
}

function setStatusText(el, message, isError = false) {
  if (!el) return;
  el.textContent = message || '';
  el.style.color = isError ? '#ff6b6b' : '';
}

function setStatus(message, isError = false) {
  setStatusText(statusEl, message, isError);
}

function setDrfxStatus(message, isError = false) {
  setStatusText(drfxStatusEl, message, isError);
}

function updateDrfxBulkActions() {
  const count = drfxSelected.size;
  if (drfxBulkCountEl) drfxBulkCountEl.textContent = `${count} selected`;
  const disabled = count === 0 || currentDrfxDisabled;
  const canSetThumb = count === 1 && !currentDrfxDisabled;
  if (drfxSetThumbBtn) drfxSetThumbBtn.disabled = !canSetThumb;
  if (drfxBulkEnableBtn) drfxBulkEnableBtn.disabled = disabled;
  if (drfxBulkDisableBtn) drfxBulkDisableBtn.disabled = disabled;
  if (drfxBulkDeleteBtn) drfxBulkDeleteBtn.disabled = count === 0;
  if (drfxBulkClearBtn) drfxBulkClearBtn.disabled = count === 0;
  if (drfxImportAllBtn) drfxImportAllBtn.disabled = !currentDrfxPath || currentDrfxPresets.length === 0;
}

function updateExportFolderButton() {
  if (!exportFolderBtn) return;
  exportFolderBtn.disabled = !currentPath;
}

function updatePathNavButtons() {
  if (!pathBackBtn || !pathForwardBtn) return;
  pathBackBtn.disabled = pathHistoryIndex <= 0;
  pathForwardBtn.disabled = pathHistoryIndex < 0 || pathHistoryIndex >= pathHistory.length - 1;
}

function pushPathHistory(nextPath) {
  if (!nextPath) return;
  const resolved = path.resolve(nextPath);
  if (pathHistoryIndex >= 0 && pathHistory[pathHistoryIndex] === resolved) {
    updatePathNavButtons();
    return;
  }
  if (pathHistoryIndex < pathHistory.length - 1) {
    pathHistory = pathHistory.slice(0, pathHistoryIndex + 1);
  }
  pathHistory.push(resolved);
  pathHistoryIndex = pathHistory.length - 1;
  updatePathNavButtons();
}

function getSearchTerm() {
  if (!searchInput) return '';
  return String(searchInput.value || '').trim().toLowerCase();
}

function updateSearchControls() {
  if (!clearSearchBtn) return;
  clearSearchBtn.disabled = !getSearchTerm();
}

function getDrfxSearchTerm() {
  if (!drfxSearchInput) return '';
  return String(drfxSearchInput.value || '').trim().toLowerCase();
}

function updateDrfxSearchControls() {
  if (!drfxClearSearchBtn) return;
  drfxClearSearchBtn.disabled = !getDrfxSearchTerm();
}

function filterEntries(entries) {
  const term = getSearchTerm();
  if (!term) return entries;
  return entries.filter((entry) => entry.displayName.toLowerCase().includes(term));
}

function filterDrfxPresets(presets) {
  const term = getDrfxSearchTerm();
  if (!term) return presets;
  return presets.filter((preset) => {
    const name = String(preset.displayName || '').toLowerCase();
    const entry = String(preset.entryName || '').toLowerCase();
    return name.includes(term) || entry.includes(term);
  });
}

function updateListStatus(total, filtered, term) {
  if (!term) {
    setStatus(`${total} item(s) found.`);
    return;
  }
  if (!filtered) {
    setStatus(`No matches for "${term}".`, true);
    return;
  }
  if (filtered === total) {
    setStatus(`${total} item(s) found.`);
    return;
  }
  setStatus(`${filtered} match(es) (of ${total}).`);
}

async function setExportFolder(targetPath) {
  if (!targetPath) return;
  try {
    const res = await ipcRenderer.invoke('set-export-folder', { path: targetPath });
    if (res?.ok) {
      setStatus(`Export folder set to ${targetPath}.`);
    } else {
      setStatus(`Export folder failed: ${res?.error || 'Unable to reach Macro Machine.'}`, true);
    }
  } catch (err) {
    setStatus(`Export folder failed: ${err?.message || err}`, true);
  }
}

function updatePathDisplay() {
  if (currentPathEl) currentPathEl.textContent = currentPath || '';
}

function updateBulkActions() {
  const count = selectedEntries.size;
  if (bulkCountEl) bulkCountEl.textContent = `${count} selected`;
  const disabled = count === 0;
  const singleEntry = count === 1 ? Array.from(selectedEntries.values())[0] : null;
  const canSetThumb = !!singleEntry && singleEntry.kind === 'setting';
  if (setThumbBtn) setThumbBtn.disabled = !canSetThumb;
  if (bulkEnableBtn) bulkEnableBtn.disabled = disabled;
  if (bulkDisableBtn) bulkDisableBtn.disabled = disabled;
  if (bulkDeleteBtn) bulkDeleteBtn.disabled = disabled;
  if (bulkClearBtn) bulkClearBtn.disabled = disabled;
  if (createFolderBtn) createFolderBtn.disabled = !currentPath;
}

function getThumbCacheKey(entryPath) {
  return metadataKey(entryPath);
}

async function loadDrfxThumbnail(entryPath, imgEl) {
  if (!entryPath || !imgEl) return;
  const cacheKey = getThumbCacheKey(entryPath);
  if (drfxThumbCache.has(cacheKey)) {
    const cached = drfxThumbCache.get(cacheKey);
    if (cached) imgEl.src = cached;
    return;
  }
  try {
    const res = await ipcRenderer.invoke('drfx-thumbnail', { path: entryPath });
    if (res && res.ok && res.thumbBase64) {
      const src = `data:image/png;base64,${res.thumbBase64}`;
      drfxThumbCache.set(cacheKey, src);
      imgEl.src = src;
      return;
    }
    drfxThumbCache.set(cacheKey, '');
  } catch (_) {
    drfxThumbCache.set(cacheKey, '');
  }
}

function isDisabledFile(name) {
  return name.toLowerCase().endsWith(DISABLED_SUFFIX);
}

function normalizeName(name) {
  if (isDisabledFile(name)) return name.slice(0, -DISABLED_SUFFIX.length);
  return name;
}

function stripDisabledSuffix(value) {
  if (!value) return value;
  const text = String(value);
  if (text.toLowerCase().endsWith(DISABLED_SUFFIX)) {
    return text.slice(0, -DISABLED_SUFFIX.length);
  }
  return text;
}

function isDisabledPath(value) {
  if (!value) return false;
  return String(value).toLowerCase().endsWith(DISABLED_SUFFIX);
}

function getSettingThumbnailPath(settingPath) {
  if (!settingPath) return null;
  const baseName = stripDisabledSuffix(path.basename(settingPath));
  if (!baseName.toLowerCase().endsWith('.setting')) return null;
  const pngName = baseName.replace(/\.setting$/i, '.png');
  return path.join(path.dirname(settingPath), pngName);
}

function fileUrlForPath(filePath) {
  if (!filePath) return '';
  const resolved = path.resolve(filePath).replace(/\\/g, '/');
  return `file:///${resolved}`;
}

async function ensureUserDataPath() {
  if (userDataPath) return userDataPath;
  const res = await ipcRenderer.invoke('get-user-data-path');
  userDataPath = res?.path || '';
  return userDataPath;
}

function getMetadataPath() {
  const base = userDataPath || '';
  return path.join(base, 'macro-explorer', 'metadata.json');
}

function getVaultDir() {
  const base = userDataPath || '';
  return path.join(base, 'macro-explorer', 'vault');
}

function ensureMetadataDir() {
  const metaPath = getMetadataPath();
  fs.mkdirSync(path.dirname(metaPath), { recursive: true });
}

function loadMetadata() {
  const metaPath = getMetadataPath();
  if (!fs.existsSync(metaPath)) return {};
  try {
    const raw = fs.readFileSync(metaPath, 'utf8');
    return JSON.parse(raw);
  } catch (_) {
    return {};
  }
}

function saveMetadata(data) {
  ensureMetadataDir();
  fs.writeFileSync(getMetadataPath(), JSON.stringify(data, null, 2));
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
  const base = path.basename(drfxPath);
  const hash = crypto.createHash('sha1').update(metadataKey(drfxPath)).digest('hex').slice(0, 8);
  return `${base}.${hash}`;
}

function storeInVault(drfxPath, force = false) {
  const vaultDir = getVaultDir();
  fs.mkdirSync(vaultDir, { recursive: true });
  const vaultPath = path.join(vaultDir, vaultFileName(drfxPath));
  if (force || !fs.existsSync(vaultPath)) {
    fs.copyFileSync(drfxPath, vaultPath);
  }
  return vaultPath;
}

function buildPresetSkipSet(disabledList) {
  const skip = new Set();
  disabledList.forEach((name) => {
    const lower = name.toLowerCase();
    skip.add(lower);
    if (lower.endsWith('.setting')) {
      skip.add(lower.replace(/\.setting$/i, '.png'));
    }
  });
  return skip;
}

function readZipEntries(zipPath, skipSet) {
  if (!yauzl) {
    return Promise.reject(new Error('DRFX support unavailable (missing yauzl).'));
  }
  return new Promise((resolve, reject) => {
    const entries = [];
    yauzl.open(zipPath, { lazyEntries: true }, (err, zip) => {
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

function writeZipEntries(destPath, entries) {
  if (!archiver) {
    return Promise.reject(new Error('DRFX support unavailable (missing archiver).'));
  }
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(destPath);
    const archive = archiver('zip', { zlib: { level: 9 } });
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

async function rebuildActiveFromVault(drfxPath, vaultPath, disabledList) {
  if (!disabledList.length) {
    fs.copyFileSync(vaultPath, drfxPath);
    return;
  }
  const skip = buildPresetSkipSet(disabledList);
  const keptEntries = await readZipEntries(vaultPath, skip);
  const tempPath = drfxPath + '.tmp';
  await writeZipEntries(tempPath, keptEntries);
  fs.renameSync(tempPath, drfxPath);
}

function resetDrfxModal() {
  currentDrfxPath = '';
  currentDrfxVaultPath = '';
  currentDrfxPresets = [];
  currentDrfxDisabled = false;
  drfxSelected.clear();
  if (drfxSearchInput) drfxSearchInput.value = '';
  if (drfxPathEl) drfxPathEl.textContent = '';
  if (drfxWarningEl) {
    drfxWarningEl.textContent = '';
    drfxWarningEl.hidden = true;
  }
  if (drfxListEl) drfxListEl.innerHTML = '';
  setDrfxStatus('');
  updateDrfxSearchControls();
  updateDrfxBulkActions();
}

function normalizeDrfxPresets(presets = []) {
  return presets.map((preset) => {
    const entryName = preset.entryName || preset.name || '';
    const displayName = preset.displayName || path.basename(entryName).replace(/\.setting$/i, '');
    const thumbUrl = preset.thumbBase64 ? `data:image/png;base64,${preset.thumbBase64}` : (preset.thumbUrl || '');
    return {
      entryName,
      displayName,
      thumbUrl,
      disabled: !!preset.disabled,
    };
  });
}

function renderDrfxPresets(presets = currentDrfxPresets) {
  if (!drfxListEl) return;
  drfxListEl.innerHTML = '';
  if (!presets.length) {
    const empty = document.createElement('div');
    empty.className = 'status';
    const term = getDrfxSearchTerm();
    empty.textContent = term ? `No presets match "${term}".` : 'No presets found in this DRFX pack.';
    drfxListEl.appendChild(empty);
    return;
  }
  presets.forEach((preset) => {
    const card = document.createElement('div');
    card.className = `drfx-card${preset.disabled ? ' is-disabled' : ''}`;
    card.dataset.preset = preset.entryName;
    card.title = preset.entryName;

    const selectRow = document.createElement('div');
    selectRow.className = 'select-row';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.dataset.action = 'select-preset';
    checkbox.disabled = currentDrfxDisabled;
    checkbox.checked = drfxSelected.has(preset.entryName);
    selectRow.appendChild(checkbox);
    const label = document.createElement('span');
    label.textContent = 'Select';
    selectRow.appendChild(label);
    card.appendChild(selectRow);

    if (preset.thumbUrl) {
      const img = document.createElement('img');
      img.src = preset.thumbUrl;
      img.alt = '';
      card.appendChild(img);
    } else {
      const placeholder = document.createElement('div');
      placeholder.className = 'thumb-placeholder';
      placeholder.textContent = 'No preview';
      card.appendChild(placeholder);
    }

    const title = document.createElement('div');
    title.className = 'title';
    title.textContent = preset.displayName;
    card.appendChild(title);

    const row = document.createElement('div');
    row.className = 'row';

    const badge = document.createElement('span');
    badge.className = `badge ${preset.disabled ? 'disabled' : 'enabled'}`;
    badge.textContent = preset.disabled ? 'Disabled' : 'Enabled';
    row.appendChild(badge);

    const actions = document.createElement('div');
    actions.className = 'row-actions';
    const openBtn = document.createElement('button');
    openBtn.type = 'button';
    openBtn.dataset.action = 'open-preset';
    openBtn.textContent = 'Open';
    actions.appendChild(openBtn);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.dataset.action = 'toggle-preset';
    btn.disabled = currentDrfxDisabled;
    btn.textContent = preset.disabled ? 'Enable' : 'Disable';
    actions.appendChild(btn);

    row.appendChild(actions);

    card.appendChild(row);
    drfxListEl.appendChild(card);
  });
  updateDrfxBulkActions();
}

function updateDrfxStatus(disabledList, filteredCount = null, term = '') {
  const total = currentDrfxPresets.length;
  const disabledCount = (disabledList || []).length;
  if (!total) {
    setDrfxStatus('No presets found.');
    return;
  }
  if (term) {
    if (!filteredCount) {
      setDrfxStatus(`No matches for "${term}".`, true);
      return;
    }
    const base = `${filteredCount} match(es) (of ${total}).`;
    if (disabledCount) {
      setDrfxStatus(`${base} ${disabledCount} disabled.`);
    } else {
      setDrfxStatus(base);
    }
    return;
  }
  if (!disabledCount) {
    setDrfxStatus(`${total} preset(s) available.`);
    return;
  }
  setDrfxStatus(`${total} preset(s), ${disabledCount} disabled.`);
}

function getCurrentDrfxDisabledList() {
  return currentDrfxPresets
    .filter((preset) => preset.disabled)
    .map((preset) => preset.entryName);
}

async function applyDrfxBulkToggle(shouldDisable) {
  if (!currentDrfxPath) return;
  if (currentDrfxDisabled) {
    setDrfxStatus('Enable the DRFX pack before editing presets.', true);
    return;
  }
  const selected = Array.from(drfxSelected);
  if (!selected.length) return;
  const disabledSet = new Set(getCurrentDrfxDisabledList().map((name) => String(name).toLowerCase()));
  if (shouldDisable) {
    selected.forEach((name) => disabledSet.add(String(name).toLowerCase()));
  } else {
    selected.forEach((name) => disabledSet.delete(String(name).toLowerCase()));
  }
  const nextDisabled = currentDrfxPresets
    .map((preset) => preset.entryName)
    .filter((name) => disabledSet.has(String(name).toLowerCase()));
  try {
    const res = await ipcRenderer.invoke('drfx-set-disabled-list', {
      path: currentDrfxPath,
      disabledList: nextDisabled,
    });
    if (!res || !res.ok) {
      setDrfxStatus(res?.error || 'Failed to update presets.', true);
      return;
    }
    const normalized = new Set((res.disabledList || []).map((name) => String(name).toLowerCase()));
    currentDrfxPresets = currentDrfxPresets.map((preset) => ({
      ...preset,
      disabled: normalized.has(String(preset.entryName).toLowerCase()),
    }));
    drfxSelected.clear();
    const filtered = filterDrfxPresets(currentDrfxPresets);
    renderDrfxPresets(filtered);
    updateDrfxStatus(res.disabledList || [], filtered.length, getDrfxSearchTerm());
  } catch (err) {
    setDrfxStatus(`Failed to update presets: ${err?.message || err}`, true);
  }
}

async function openDrfxModal(drfxPath) {
  if (!drfxModal) return;
  resetDrfxModal();
  currentDrfxPath = drfxPath;
  if (drfxPathEl) drfxPathEl.textContent = drfxPath || '';
  drfxModal.hidden = false;

  try {
    setDrfxStatus('Loading presets...');
    const res = await ipcRenderer.invoke('drfx-list', { path: drfxPath });
    if (!res || !res.ok) {
      setDrfxStatus(res?.error || 'Failed to read DRFX.', true);
      return;
    }
    currentDrfxDisabled = !!res.disabled;
    if (currentDrfxDisabled && drfxWarningEl) {
      drfxWarningEl.textContent = 'This DRFX pack is disabled. Enable it to apply preset changes.';
      drfxWarningEl.hidden = false;
    }
    currentDrfxPresets = normalizeDrfxPresets(res.presets || []);
    const filtered = filterDrfxPresets(currentDrfxPresets);
    renderDrfxPresets(filtered);
    updateDrfxStatus(res.disabledList || [], filtered.length, getDrfxSearchTerm());
    updateDrfxSearchControls();
    updateDrfxBulkActions();
  } catch (err) {
    setDrfxStatus(`Failed to read DRFX: ${err?.message || err}`, true);
  }
}

function closeDrfxModal() {
  if (!drfxModal) return;
  drfxModal.hidden = true;
  resetDrfxModal();
}

async function toggleDrfxPreset(presetName) {
  if (!presetName || !currentDrfxPath) return;
  if (currentDrfxDisabled) {
    setDrfxStatus('Enable the DRFX pack before editing presets.', true);
    return;
  }
  try {
    const res = await ipcRenderer.invoke('drfx-toggle-preset', {
      path: currentDrfxPath,
      presetName,
    });
    if (!res || !res.ok) {
      setDrfxStatus(res?.error || 'Failed to update preset.', true);
      return;
    }
    const disabledList = res.disabledList || [];
    const disabledSet = new Set(disabledList.map((name) => String(name).toLowerCase()));
    currentDrfxPresets = currentDrfxPresets.map((preset) => ({
      ...preset,
      disabled: disabledSet.has(String(preset.entryName).toLowerCase()),
    }));
    const filtered = filterDrfxPresets(currentDrfxPresets);
    renderDrfxPresets(filtered);
    updateDrfxStatus(disabledList, filtered.length, getDrfxSearchTerm());
  } catch (err) {
    setDrfxStatus(`Failed to update preset: ${err?.message || err}`, true);
  }
}

async function openDrfxPreset(presetName) {
  if (!presetName || !currentDrfxPath) return;
  try {
    setDrfxStatus('Preparing preset...');
    const res = await ipcRenderer.invoke('drfx-extract-preset', {
      path: currentDrfxPath,
      presetName,
    });
    if (!res || !res.ok || !res.filePath) {
      setDrfxStatus(res?.error || 'Failed to extract preset.', true);
      return;
    }
    const openRes = await ipcRenderer.invoke('open-in-macro-machine', { path: res.filePath });
    if (!openRes || !openRes.ok) {
      setDrfxStatus(openRes?.error || 'Macro Machine window not available.', true);
      return;
    }
    setDrfxStatus('Sent to Macro Machine.');
  } catch (err) {
    setDrfxStatus(`Failed to open preset: ${err?.message || err}`, true);
  }
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function importAllDrfxPresets() {
  if (!currentDrfxPath) return;
  const presetNames = currentDrfxPresets
    .map(preset => preset.entryName)
    .filter(name => !!name);
  if (!presetNames.length) {
    setDrfxStatus('No presets available to import.', true);
    return;
  }
  if (presetNames.length > 5) {
    const ok = window.confirm(`Import ${presetNames.length} presets into Macro Machine? This will open a tab for each.`);
    if (!ok) return;
  }
  setDrfxStatus(`Importing ${presetNames.length} preset(s) to Macro Machine...`);
  let imported = 0;
  let failed = 0;
  for (const presetName of presetNames) {
    try {
      const res = await ipcRenderer.invoke('drfx-extract-preset', {
        path: currentDrfxPath,
        presetName,
      });
      if (!res || !res.ok || !res.filePath) {
        failed += 1;
        continue;
      }
      const openRes = await ipcRenderer.invoke('open-in-macro-machine', { path: res.filePath });
      if (!openRes || !openRes.ok) {
        failed += 1;
        continue;
      }
      imported += 1;
      setDrfxStatus(`Importing... (${imported}/${presetNames.length})`);
      await sleep(80);
    } catch (_) {
      failed += 1;
    }
  }
  if (!failed) {
    setDrfxStatus(`Imported ${imported} preset(s) to Macro Machine.`);
  } else {
    setDrfxStatus(`Imported ${imported} preset(s), ${failed} failed.`, true);
  }
}

function classifyEntry(dirent) {
  const name = dirent.name;
  const lower = name.toLowerCase();
  if (dirent.isDirectory()) {
    return { kind: 'folder', name, disabled: false };
  }
  if (lower.endsWith('.drfx_b') || lower.endsWith('.drfx_b' + DISABLED_SUFFIX)) {
    return null;
  }
  if (lower.endsWith('.setting' + DISABLED_SUFFIX) || lower.endsWith('.setting')) {
    return { kind: 'setting', name, disabled: lower.endsWith(DISABLED_SUFFIX) };
  }
  if (lower.endsWith('.drfx' + DISABLED_SUFFIX) || lower.endsWith('.drfx')) {
    return { kind: 'drfx', name, disabled: lower.endsWith(DISABLED_SUFFIX) };
  }
  return null;
}

function formatTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleString();
}

function listEntries(dirPath) {
  const items = fs.readdirSync(dirPath, { withFileTypes: true });
  const entries = [];
  items.forEach((dirent) => {
    if (dirent.name.startsWith('.')) return;
    const info = classifyEntry(dirent);
    if (!info) return;
    if (isFoldersOnly && info.kind !== 'folder') return;
    const fullPath = path.join(dirPath, dirent.name);
    let mtime = null;
    try {
      const stat = fs.statSync(fullPath);
      mtime = stat.mtime;
    } catch (_) {}
    entries.push({
      ...info,
      path: fullPath,
      displayName: normalizeName(info.name),
      mtime,
    });
  });
  entries.sort((a, b) => {
    if (a.kind === b.kind) {
      return a.displayName.localeCompare(b.displayName);
    }
    if (a.kind === 'folder') return -1;
    if (b.kind === 'folder') return 1;
    return a.kind.localeCompare(b.kind);
  });
  return entries;
}

function buildEntryNameCell(entry) {
  const cell = document.createElement('td');
  if (entry.kind === 'drfx') {
    const nameWrap = document.createElement('div');
    nameWrap.className = 'name-wrap';
    const thumb = document.createElement('img');
    thumb.className = 'drfx-thumb';
    thumb.alt = '';
    nameWrap.appendChild(thumb);
    const label = document.createElement('span');
    label.textContent = entry.displayName;
    nameWrap.appendChild(label);
    cell.appendChild(nameWrap);
    loadDrfxThumbnail(entry.path, thumb);
    return cell;
  }
  if (entry.kind === 'setting') {
    const nameWrap = document.createElement('div');
    nameWrap.className = 'name-wrap';
    const thumbPath = getSettingThumbnailPath(entry.path);
    if (thumbPath && fs.existsSync(thumbPath)) {
      const thumb = document.createElement('img');
      thumb.className = 'preset-thumb';
      thumb.alt = '';
      thumb.src = `${fileUrlForPath(thumbPath)}?v=${Date.now()}`;
      nameWrap.appendChild(thumb);
    } else {
      const placeholder = document.createElement('div');
      placeholder.className = 'thumb-placeholder';
      placeholder.title = 'No thumbnail';
      nameWrap.appendChild(placeholder);
    }
    const label = document.createElement('span');
    label.textContent = entry.displayName;
    nameWrap.appendChild(label);
    cell.appendChild(nameWrap);
    return cell;
  }
  cell.textContent = entry.displayName;
  return cell;
}

function renderEntries(entries) {
  entriesBody.innerHTML = '';
  if (!entries.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 6;
    const term = getSearchTerm();
    cell.textContent = term ? `No presets match "${term}".` : 'No presets found in this folder.';
    row.appendChild(cell);
    entriesBody.appendChild(row);
    updateBulkActions();
    return;
  }
  entries.forEach((entry) => {
    const row = document.createElement('tr');
    row.dataset.path = entry.path;
    row.dataset.kind = entry.kind;
    if (entry.kind === 'folder') row.classList.add('is-folder');

    const selectCell = document.createElement('td');
    selectCell.className = 'select-col';
    if (entry.kind !== 'folder') {
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.action = 'select';
      checkbox.checked = selectedEntries.has(entry.path);
      selectCell.appendChild(checkbox);
    }
    row.appendChild(selectCell);

    const nameCell = buildEntryNameCell(entry);
    row.appendChild(nameCell);

    const typeCell = document.createElement('td');
    typeCell.textContent = entry.kind === 'setting' ? '.setting' : entry.kind === 'drfx' ? '.drfx' : 'Folder';
    row.appendChild(typeCell);

    const statusCell = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = `badge ${entry.disabled ? 'disabled' : 'enabled'}`;
    badge.textContent = entry.disabled ? 'Disabled' : 'Enabled';
    statusCell.appendChild(badge);
    row.appendChild(statusCell);

    const modCell = document.createElement('td');
    modCell.textContent = formatTime(entry.mtime);
    row.appendChild(modCell);

    const actionsCell = document.createElement('td');
    actionsCell.className = 'actions';
    if (entry.kind === 'folder') {
      if (isFoldersOnly) {
        actionsCell.appendChild(buildActionButton('Select', 'select-export'));
      } else {
        actionsCell.appendChild(buildActionButton('Open', 'open'));
      }
    }
    if (entry.kind === 'setting') {
      actionsCell.appendChild(buildActionButton('Open in Macro Machine', 'open-macro'));
      actionsCell.appendChild(buildActionButton('Set Thumbnail', 'set-thumb'));
      actionsCell.appendChild(buildActionButton(entry.disabled ? 'Enable' : 'Disable', 'toggle'));
      actionsCell.appendChild(buildActionButton('Delete', 'delete'));
    }
    if (entry.kind === 'drfx') {
      actionsCell.appendChild(buildActionButton('Browse Presets', 'browse-drfx'));
      actionsCell.appendChild(buildActionButton(entry.disabled ? 'Enable' : 'Disable', 'toggle'));
      actionsCell.appendChild(buildActionButton('Delete', 'delete'));
    }
    if (!isFoldersOnly) {
      actionsCell.appendChild(buildActionButton('Reveal', 'reveal'));
    }
    row.appendChild(actionsCell);

    entriesBody.appendChild(row);
  });
  updateBulkActions();
}

function buildActionButton(label, action) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.textContent = label;
  btn.dataset.action = action;
  return btn;
}

function refreshList() {
  if (!currentPath) return;
  try {
    currentEntries = listEntries(currentPath);
    const filtered = filterEntries(currentEntries);
    renderEntries(filtered);
    updateListStatus(currentEntries.length, filtered.length, getSearchTerm());
    updateSearchControls();
  } catch (err) {
    setStatus(`Unable to read folder: ${err.message || err}`, true);
    entriesBody.innerHTML = '';
  }
  updatePathDisplay();
}

function setCurrentPath(newPath, options = {}) {
  if (!newPath) return;
  const resolved = path.resolve(newPath);
  if (resolved === currentPath) {
    updatePathDisplay();
    updateExportFolderButton();
    updatePathNavButtons();
    return;
  }
  currentPath = resolved;
  if (options.pushHistory !== false) {
    pushPathHistory(resolved);
  } else {
    updatePathNavButtons();
  }
  clearSelection();
  refreshList();
  updateExportFolderButton();
}

function selectRoot(key) {
  currentRootKey = key;
  const nextPath = roots[key];
  if (nextPath) {
    setCurrentPath(nextPath);
  } else if (currentPath) {
    setCurrentPath(currentPath);
  } else {
    setStatus('Select a folder to begin.', false);
  }
}

async function chooseRootFolder() {
  try {
    const res = await ipcRenderer.invoke('select-folder', { defaultPath: currentPath });
    if (!res || res.canceled || !res.filePath) return;
    roots.Custom = res.filePath;
    rootSelect.value = 'Custom';
    selectRoot('Custom');
  } catch (err) {
    setStatus(`Folder selection failed: ${err.message || err}`, true);
  }
}

function toggleDisabled(entryPath) {
  const baseName = path.basename(entryPath);
  let nextName = null;
  if (isDisabledFile(baseName)) {
    nextName = baseName.slice(0, -DISABLED_SUFFIX.length);
  } else {
    nextName = baseName + DISABLED_SUFFIX;
  }
  const nextPath = path.join(path.dirname(entryPath), nextName);
  fs.renameSync(entryPath, nextPath);
}

function setDisabledState(entryPath, shouldDisable) {
  const baseName = path.basename(entryPath);
  const isDisabled = isDisabledFile(baseName);
  if (shouldDisable === isDisabled) return;
  let nextName = null;
  if (shouldDisable) {
    nextName = baseName + DISABLED_SUFFIX;
  } else {
    nextName = baseName.slice(0, -DISABLED_SUFFIX.length);
  }
  const nextPath = path.join(path.dirname(entryPath), nextName);
  fs.renameSync(entryPath, nextPath);
}

function clearSelection() {
  selectedEntries.clear();
  updateBulkActions();
}

async function openInMacroMachine(entryPath) {
  try {
    await ipcRenderer.invoke('open-in-macro-machine', { path: entryPath });
    setStatus('Sent to Macro Machine.');
  } catch (err) {
    setStatus(`Open failed: ${err.message || err}`, true);
  }
}

function revealInExplorer(entryPath) {
  try {
    shell.showItemInFolder(entryPath);
  } catch (_) {}
}

function closeTextPrompt(value) {
  if (!textPromptModal) return;
  textPromptModal.hidden = true;
  if (textPromptResolver) {
    const resolve = textPromptResolver;
    textPromptResolver = null;
    resolve(value);
  }
}

function openTextPrompt({ title, label, confirmText, initialValue } = {}) {
  return new Promise((resolve) => {
    if (!textPromptModal || !textPromptInput || !textPromptOkBtn) {
      resolve(null);
      return;
    }
    if (textPromptResolver) {
      const prev = textPromptResolver;
      textPromptResolver = null;
      prev(null);
    }
    textPromptResolver = resolve;
    if (textPromptTitle) textPromptTitle.textContent = title || 'Enter value';
    if (textPromptLabel) textPromptLabel.textContent = label || 'Value';
    textPromptInput.value = initialValue || '';
    textPromptOkBtn.textContent = confirmText || 'OK';
    textPromptModal.hidden = false;
    setTimeout(() => {
      try {
        textPromptInput.focus();
        textPromptInput.select();
      } catch (_) {}
    }, 0);
  });
}

async function createFolder() {
  if (!currentPath) return;
  const name = await openTextPrompt({
    title: 'Create folder',
    label: 'Folder name',
    confirmText: 'Create',
    initialValue: '',
  });
  if (!name) return;
  const cleaned = String(name).trim();
  if (!cleaned) return;
  const target = path.join(currentPath, cleaned);
  try {
    if (fs.existsSync(target)) {
      setStatus('Folder already exists.', true);
      return;
    }
    fs.mkdirSync(target, { recursive: true });
    setStatus('Folder created.');
    refreshList();
  } catch (err) {
    setStatus(`Failed to create folder: ${err?.message || err}`, true);
  }
}

function reportThumbnailStatus(contextType, message, isError = false) {
  if (contextType === 'drfx') {
    setDrfxStatus(message, isError);
  } else {
    setStatus(message, isError);
  }
}

function requestThumbnailPick(context) {
  pendingThumbContext = context;
  if (!thumbFileInput) {
    reportThumbnailStatus(context?.type, 'Thumbnail selection is unavailable.', true);
    pendingThumbContext = null;
    return;
  }
  thumbFileInput.value = '';
  thumbFileInput.click();
}

function getEntryDisplayName(entry) {
  if (!entry || !entry.path) return '';
  const baseName = stripDisabledSuffix(path.basename(entry.path));
  return baseName || entry.path;
}

function buildDeletePrompt(label, names) {
  if (!names.length) return '';
  const preview = names.slice(0, 10).join('\n');
  const extra = names.length > 10 ? `\n...and ${names.length - 10} more` : '';
  const noun = names.length === 1 ? label : `${label}s`;
  return `Are you sure you want to permanently delete ${names.length} ${noun}?\n${preview}${extra}`;
}

function deleteSettingFile(entryPath) {
  if (!entryPath || !fs.existsSync(entryPath)) {
    throw new Error('File not found.');
  }
  fs.unlinkSync(entryPath);
  const thumbPath = getSettingThumbnailPath(entryPath);
  if (thumbPath && fs.existsSync(thumbPath)) {
    fs.unlinkSync(thumbPath);
  }
}

async function deleteDrfxPack(entryPath) {
  const res = await ipcRenderer.invoke('drfx-delete-pack', { path: entryPath });
  if (!res || !res.ok) {
    throw new Error(res?.error || 'Failed to delete DRFX pack.');
  }
}

async function deleteEntries(entries) {
  const targets = Array.isArray(entries) ? entries : [];
  if (!targets.length) return;
  const names = targets.map(getEntryDisplayName).filter(Boolean);
  const prompt = buildDeletePrompt('item', names);
  if (prompt && !window.confirm(prompt)) return;
  let failed = 0;
  let changed = 0;
  for (const entry of targets) {
    try {
      if (entry.kind === 'setting') {
        deleteSettingFile(entry.path);
      } else if (entry.kind === 'drfx') {
        await deleteDrfxPack(entry.path);
      } else {
        throw new Error('Unsupported delete type.');
      }
      changed += 1;
    } catch (_) {
      failed += 1;
    }
  }
  clearSelection();
  refreshList();
  if (failed) {
    setStatus(`Deleted ${changed} item(s). ${failed} failed.`, true);
  } else {
    setStatus(`Deleted ${changed} item(s).`);
  }
}

async function deleteDrfxPresets() {
  if (!currentDrfxPath) return;
  const selected = Array.from(drfxSelected);
  if (!selected.length) return;
  const prompt = buildDeletePrompt('preset', selected);
  if (prompt && !window.confirm(prompt)) return;
  try {
    const res = await ipcRenderer.invoke('drfx-delete-presets', {
      path: currentDrfxPath,
      presetNames: selected,
    });
    if (!res || !res.ok) {
      setDrfxStatus(res?.error || 'Failed to delete presets.', true);
      return;
    }
    drfxSelected.clear();
    await openDrfxModal(currentDrfxPath);
    setDrfxStatus('Preset(s) deleted.');
  } catch (err) {
    setDrfxStatus(`Failed to delete presets: ${err?.message || err}`, true);
  }
}

function applyBulkToggle(shouldDisable) {
  const targets = Array.from(selectedEntries.values());
  if (!targets.length) return;
  let changed = 0;
  let failed = 0;
  targets.forEach((entry) => {
    try {
      setDisabledState(entry.path, shouldDisable);
      changed += 1;
    } catch (_) {
      failed += 1;
    }
  });
  clearSelection();
  refreshList();
  if (failed) {
    setStatus(`Updated ${changed} item(s). ${failed} failed.`, true);
  } else if (changed) {
    setStatus(`${shouldDisable ? 'Disabled' : 'Enabled'} ${changed} item(s).`);
  }
}

entriesBody.addEventListener('click', async (event) => {
  const btn = event.target.closest('button');
  if (!btn) return;
  const row = btn.closest('tr');
  if (!row) return;
  const entryPath = row.dataset.path;
  const kind = row.dataset.kind;
  if (!entryPath || !kind) return;

  const action = btn.dataset.action;
  try {
    if (action === 'open' && kind === 'folder') {
      setCurrentPath(entryPath);
    } else if (action === 'select-export' && kind === 'folder') {
      await setExportFolder(entryPath);
    } else if (action === 'toggle') {
      toggleDisabled(entryPath);
      refreshList();
    } else if (action === 'set-thumb' && kind === 'setting') {
      requestThumbnailPick({ type: 'setting', path: entryPath });
    } else if (action === 'open-macro' && kind === 'setting') {
      await openInMacroMachine(entryPath);
    } else if (action === 'browse-drfx' && kind === 'drfx') {
      await openDrfxModal(entryPath);
    } else if (action === 'delete') {
      await deleteEntries([{ path: entryPath, kind }]);
    } else if (action === 'reveal') {
      revealInExplorer(entryPath);
    }
  } catch (err) {
    setStatus(`Action failed: ${err.message || err}`, true);
  }
});

entriesBody.addEventListener('dblclick', async (event) => {
  const row = event.target.closest('tr');
  if (!row) return;
  const entryPath = row.dataset.path;
  const kind = row.dataset.kind;
  if (kind === 'folder' && entryPath) {
    setCurrentPath(entryPath);
  } else if (kind === 'setting' && entryPath) {
    await openInMacroMachine(entryPath);
  } else if (kind === 'drfx' && entryPath) {
    await openDrfxModal(entryPath);
  }
});

entriesBody.addEventListener('change', (event) => {
  const target = event.target;
  if (!target || target.tagName !== 'INPUT' || target.type !== 'checkbox') return;
  if (target.dataset.action !== 'select') return;
  const row = target.closest('tr');
  if (!row) return;
  const entryPath = row.dataset.path;
  const kind = row.dataset.kind;
  if (!entryPath || !kind || kind === 'folder') return;
  if (target.checked) {
    selectedEntries.set(entryPath, { path: entryPath, kind });
  } else {
    selectedEntries.delete(entryPath);
  }
  updateBulkActions();
});

closeDrfxBtn?.addEventListener('click', () => closeDrfxModal());
drfxModal?.addEventListener('click', (event) => {
  if (event.target === drfxModal) closeDrfxModal();
});
textPromptCancelBtn?.addEventListener('click', () => closeTextPrompt(null));
textPromptCloseBtn?.addEventListener('click', () => closeTextPrompt(null));
textPromptOkBtn?.addEventListener('click', () => closeTextPrompt(textPromptInput ? textPromptInput.value : ''));
textPromptModal?.addEventListener('click', (event) => {
  if (event.target === textPromptModal) closeTextPrompt(null);
});
textPromptInput?.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    closeTextPrompt(textPromptInput.value);
  }
});
drfxListEl?.addEventListener('click', async (event) => {
  const btn = event.target.closest('button');
  if (!btn) return;
  const card = btn.closest('.drfx-card');
  if (!card) return;
  const presetName = card.dataset.preset;
  const action = btn.dataset.action;
  if (action === 'toggle-preset') {
    await toggleDrfxPreset(presetName);
  } else if (action === 'open-preset') {
    await openDrfxPreset(presetName);
  }
});
drfxListEl?.addEventListener('change', (event) => {
  const target = event.target;
  if (!target || target.tagName !== 'INPUT' || target.type !== 'checkbox') return;
  if (target.dataset.action !== 'select-preset') return;
  const card = target.closest('.drfx-card');
  if (!card) return;
  const presetName = card.dataset.preset;
  if (!presetName) return;
  if (target.checked) {
    drfxSelected.add(presetName);
  } else {
    drfxSelected.delete(presetName);
  }
  updateDrfxBulkActions();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && textPromptModal && !textPromptModal.hidden) {
    closeTextPrompt(null);
    return;
  }
  if (event.key === 'Escape' && drfxModal && !drfxModal.hidden) {
    closeDrfxModal();
    return;
  }
  if (!(event.ctrlKey || event.metaKey)) return;
  const key = String(event.key || '').toLowerCase();
  if (key === 'r') {
    event.preventDefault();
    refreshList();
    return;
  }
  if (key === 'f') {
    event.preventDefault();
    if (drfxModal && !drfxModal.hidden && drfxSearchInput) {
      drfxSearchInput.focus();
      drfxSearchInput.select();
      return;
    }
    if (searchInput) {
      searchInput.focus();
      searchInput.select();
    }
  }
});

setThumbBtn?.addEventListener('click', () => {
  const entry = selectedEntries.size === 1 ? Array.from(selectedEntries.values())[0] : null;
  if (!entry || entry.kind !== 'setting') {
    setStatus('Select a single .setting file to set a thumbnail.', true);
    return;
  }
  requestThumbnailPick({ type: 'setting', path: entry.path });
});

drfxSetThumbBtn?.addEventListener('click', () => {
  if (currentDrfxDisabled) {
    setDrfxStatus('Enable the DRFX pack before editing presets.', true);
    return;
  }
  const selected = Array.from(drfxSelected);
  if (selected.length !== 1) {
    setDrfxStatus('Select a single preset to set a thumbnail.', true);
    return;
  }
  requestThumbnailPick({ type: 'drfx', path: currentDrfxPath, presetName: selected[0] });
});

thumbFileInput?.addEventListener('change', async () => {
  const file = thumbFileInput.files && thumbFileInput.files[0] ? thumbFileInput.files[0] : null;
  const context = pendingThumbContext;
  pendingThumbContext = null;
  if (!file || !context) return;
  const imagePath = file.path || '';
  thumbFileInput.value = '';
  if (!imagePath) {
    reportThumbnailStatus(context.type, 'Unable to read the selected file path.', true);
    return;
  }
  if (path.extname(imagePath).toLowerCase() !== '.png') {
    reportThumbnailStatus(context.type, 'Thumbnail must be a .png image.', true);
    return;
  }
  try {
    if (context.type === 'setting') {
      const targetPath = getSettingThumbnailPath(context.path);
      if (!targetPath) {
        reportThumbnailStatus(context.type, 'Select a .setting file to set a thumbnail.', true);
        return;
      }
      fs.copyFileSync(imagePath, targetPath);
      reportThumbnailStatus(context.type, `Thumbnail set for ${path.basename(context.path)}.`);
      refreshList();
    } else if (context.type === 'drfx') {
      const res = await ipcRenderer.invoke('drfx-set-thumbnail', {
        path: context.path,
        presetName: context.presetName,
        imagePath,
      });
      if (!res || !res.ok) {
        reportThumbnailStatus(context.type, res?.error || 'Failed to update thumbnail.', true);
        return;
      }
      drfxThumbCache.delete(getThumbCacheKey(context.path));
      await openDrfxModal(context.path);
      reportThumbnailStatus(context.type, 'Thumbnail updated.');
      refreshList();
    }
  } catch (err) {
    reportThumbnailStatus(context.type, `Thumbnail update failed: ${err?.message || err}`, true);
  }
});

rootSelect.addEventListener('change', () => selectRoot(rootSelect.value));
chooseRootBtn.addEventListener('click', () => chooseRootFolder());
refreshBtn.addEventListener('click', () => refreshList());
searchInput?.addEventListener('input', () => {
  clearSelection();
  const filtered = filterEntries(currentEntries);
  renderEntries(filtered);
  updateListStatus(currentEntries.length, filtered.length, getSearchTerm());
  updateSearchControls();
});
clearSearchBtn?.addEventListener('click', () => {
  if (searchInput) searchInput.value = '';
  clearSelection();
  const filtered = filterEntries(currentEntries);
  renderEntries(filtered);
  updateListStatus(currentEntries.length, filtered.length, getSearchTerm());
  updateSearchControls();
});
drfxSearchInput?.addEventListener('input', () => {
  drfxSelected.clear();
  const filtered = filterDrfxPresets(currentDrfxPresets);
  renderDrfxPresets(filtered);
  updateDrfxStatus(getCurrentDrfxDisabledList(), filtered.length, getDrfxSearchTerm());
  updateDrfxSearchControls();
});
drfxClearSearchBtn?.addEventListener('click', () => {
  if (drfxSearchInput) drfxSearchInput.value = '';
  drfxSelected.clear();
  const filtered = filterDrfxPresets(currentDrfxPresets);
  renderDrfxPresets(filtered);
  updateDrfxStatus(getCurrentDrfxDisabledList(), filtered.length, getDrfxSearchTerm());
  updateDrfxSearchControls();
});
pathBackBtn?.addEventListener('click', () => {
  if (pathHistoryIndex <= 0) return;
  pathHistoryIndex -= 1;
  setCurrentPath(pathHistory[pathHistoryIndex], { pushHistory: false });
});
pathForwardBtn?.addEventListener('click', () => {
  if (pathHistoryIndex < 0 || pathHistoryIndex >= pathHistory.length - 1) return;
  pathHistoryIndex += 1;
  setCurrentPath(pathHistory[pathHistoryIndex], { pushHistory: false });
});
exportFolderBtn?.addEventListener('click', () => setExportFolder(currentPath));
createFolderBtn?.addEventListener('click', () => createFolder());
bulkEnableBtn?.addEventListener('click', () => applyBulkToggle(false));
bulkDisableBtn?.addEventListener('click', () => applyBulkToggle(true));
bulkDeleteBtn?.addEventListener('click', () => {
  const entries = Array.from(selectedEntries.values());
  deleteEntries(entries);
});
bulkClearBtn?.addEventListener('click', () => {
  clearSelection();
  refreshList();
});
drfxBulkEnableBtn?.addEventListener('click', () => applyDrfxBulkToggle(false));
drfxBulkDisableBtn?.addEventListener('click', () => applyDrfxBulkToggle(true));
drfxBulkDeleteBtn?.addEventListener('click', () => {
  deleteDrfxPresets();
});
drfxBulkClearBtn?.addEventListener('click', () => {
  drfxSelected.clear();
  const filtered = filterDrfxPresets(currentDrfxPresets);
  renderDrfxPresets(filtered);
  updateDrfxStatus(getCurrentDrfxDisabledList(), filtered.length, getDrfxSearchTerm());
});
drfxImportAllBtn?.addEventListener('click', () => {
  importAllDrfxPresets();
});
initRootSelect();
selectRoot(currentRootKey || 'Templates');
updateExportFolderButton();
