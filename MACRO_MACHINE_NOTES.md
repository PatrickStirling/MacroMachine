# Macro Machine – Project Notes

## Overview

Macro Machine is a desktop + browser tool for reorganizing **Fusion `.setting` macros**:

- Parses a macro’s `GroupOperator` / `MacroOperator` block.
- Shows **Published Controls** and lets you reorder, group, and rename them.
- Shows a **Nodes** pane with tools/modifiers and their controls to publish.
- Exports a cleaned `.setting` with the new published control order and macro name.

There are two main parts:

- `FusionMacroReorderer/` – the web app (pure HTML/CSS/JS + local JSON).
- `FusionMacroReorderer-electron/` – the Electron wrapper and installers.

---

## Structure

### Web app (`FusionMacroReorderer/`)

Key files:

- `index.html`
  - Title/header: “Macro Machine by Stirling Supply Co”.
  - File loader: choose/drop `.setting`, import from clipboard.
  - Controls section:
    - **Published Controls** pane.
    - **Nodes** pane (tools & modifiers).
  - Diagnostics panel (toggle from Electron menu).

- `main.js`
  - Parses `.setting` text:
    - Finds outer `GroupOperator`/`MacroOperator`.
    - Extracts `Inputs = ordered() { ... InstanceInput ... }`.
    - Extracts `Tools = ordered() { ... }`.
  - Tracks state in `parseResult`:
    - `entries` (published controls), `order` (current order).
    - Selection, label groups, color groups.
    - Macro name: `macroName`, `macroNameOriginal`.
  - **Published Controls UI**
    - Editable display name.
    - Label groups (`isLabel` + `labelCount`).
    - Color groups via `controlGroup` (per tool + group id).
    - Undo / redo / reset / remove / validate / fix pages.
  - **Nodes pane**
    - Builds a node list from `Tools` and references.
    - Uses node and modifier catalogs for control metadata.
    - Lets you publish controls from nodes into the Published list.

- `style.css`
  - Dark theme, split panels, scrollbars.
  - Published list rows: grid columns:
    - Checkbox, drag handle, arrow column, text, buttons.
  - **Label groups (green):**
    - Label row: stronger green left border; member rows: weaker green left border.
    - Green arrow in its own column (between handle and text).
  - **Color groups (orange):**
    - Orange vertical bar on the left of the text block.
    - Orange arrow in the same arrow column as the green arrow.
    - Color-group rows can overlap with label groups; both signals show.

- `FusionNodeCatalog.cleaned.json`
- `FusionModifierCatalog.json`
  - Local JSON catalogs of tools/modifiers and their controls.
  - Used to derive control lists in the Nodes pane.
  - In browser: loaded via “Import Node Catalog” / “Import Modifier Catalog”.
  - In Electron: auto-loaded via `fetch` on startup.

---

## Behavior & Features

### Macro name

- Parsed from the `Name = GroupOperator {` / `MacroOperator` line.
- UI shows an editable “Macro:” field.
- On export, the first matching macro definition is rewritten to the edited name.

### Label groups

- Any entry can be turned into a label (`isLabel`) with `labelCount` = number of following rows in the group.
- Visual:
  - Green left border for the group.
  - Green arrow in the arrow column (dedicated grid column).
- Groups move as a block when reordering.

### Color groups (control groups)

- Derived from `ControlGroup = N` for a tool’s controls.
- Contiguous runs of the same group (per `sourceOp`) form a color group.
- Visual:
  - Orange vertical bar at the left of the text block.
  - Orange arrow in the arrow column.
- Combined with label groups, both green (labels) and orange (color groups) cues are visible.

### URL “Insert Link Launcher” button

- For ButtonControls (`INPID_InputControl = "ButtonControl"`) with empty or recognized `BTNCS_Execute`:
  - Shows a **URL** field and **Insert Link Launcher** button in Published controls.
  - Per-control state tracked with:
    - `parseResult.insertUrlMap` (key: `"SourceOp.Source"`).
    - `parseResult.insertClickedKeys` / `buttonExactInsert`.
- On export, inserts a cross-platform launcher:

  ```lua
  local url = "<user URL>"
  local sep = package.config:sub(1,1)
  if sep == "\\" then
    os.execute('start "" "'..url..'"')
  else
    local uname = io.popen("uname"):read("*l")
    if uname == "Darwin" then
      os.execute('open "'..url..'"')
    else
      os.execute('xdg-open "'..url..'"')
    end
  end
  ```

- Inserted both at:
  - Tool-level `UserControls` block.
  - Group-level `UserControls` block (macro-level button), creating it if needed.

### Export / validation

- Rebuilds the Inputs block with the new `InstanceInput` order.
- Ensures each `InstanceInput { ... },` has a trailing comma.
- Validation tools:
  - `Validate` – checks braces, counts `InstanceInput`, missing `Page` / `Name`.
  - `Fix missing pages` – adds `Page = "Controls"` where missing.

---

## Electron wrapper (`FusionMacroReorderer-electron/`)

Key files:

- `package.json`
  - `"main": "main.js"`.
  - Scripts:
    - `"start": "electron ."`
    - `"dist": "electron-builder"`
  - Dev deps:
    - `electron`
    - `electron-builder`
  - `build`:
    - `"appId": "com.stirlingsupply.macromachine"`
    - `"productName": "Macro Machine"`
    - `"files": ["**/*"]`
    - `"extraResources"`: bundles `../FusionMacroReorderer` into app.
    - `"mac": { "target": "dmg", "category": "public.app-category.video" }`
    - `"win": { "target": "nsis" }`

- `main.js`
  - Creates a `BrowserWindow` with:
    - `contextIsolation: false`
    - `nodeIntegration: true`
  - Loads:
    - Dev: `../FusionMacroReorderer/index.html`
    - Prod: `FusionMacroReorderer/index.html` from bundled resources.
  - IPC handlers:
    - `save-setting-file` – native save dialog + `fs.writeFileSync`.
    - `open-setting-file` – native open dialog + `fs.readFileSync`.
  - App menu (sends `fmr-menu` actions to renderer):
    - File → Open…, Save, Import/Export Clipboard, Quit.
    - View → Reload, Toggle DevTools, Toggle Diagnostics.

---

## Electron bridge (renderer side)

In `FusionMacroReorderer/main.js`:

- Detects Electron via `process.versions.electron`.
- `require('electron')` in renderer to get:
  - `clipboard`
  - `ipcRenderer`
- Defines `window.FusionMacroReordererNative`:

  - `isElectron: true`
  - `saveSettingFile(payload)`
  - `openSettingFile()`
  - `readClipboard()`
  - `writeClipboard(text)`

- Listens for `fmr-menu` from main:

  - `action: 'open'` → native open → `loadSettingText`.
  - `action: 'save'` → click Export button.
  - `action: 'importClipboard'` / `exportClipboard` → call existing clipboard handlers.
  - `action: 'toggleDiagnostics'` → toggle diagnostics panel.

- Auto-loads catalogs when `isElectron`:
  - `fetch('FusionNodeCatalog.cleaned.json')`
  - `fetch('FusionModifierCatalog.json')`
  - Populates `nodeCatalog` / `modifierCatalog` and re-parses nodes if a macro is loaded.

---

## Running & Building

### Dev (Electron)

From `FusionMacroReorderer-electron`:

```bash
npm install
npm start
```

- Loads the web app from the sibling `FusionMacroReorderer/` folder.
- Uses auto-loaded catalogs and native file/clipboard actions.

### Dev (Browser only)

- Open `FusionMacroReorderer/index.html` in a browser (Chrome/Edge/etc.).
- Manually import node/modifier catalogs via buttons in the Nodes pane.
- File saving uses browser download; clipboard uses browser APIs.

### Build installers

From `FusionMacroReorderer-electron`:

```bash
npm run dist
```

- **Windows**: NSIS installer (`.exe`) under `dist/`.
- **macOS**: DMG installer under `dist/`.

Note: raw `npm run dist` macOS builds are **unsigned / not notarized** unless you run the signing steps below.

---

## macOS Signing + Notarization (current workflow)

Pre-reqs:
- Developer ID Application cert + private key in **login** keychain.
- `Developer ID Certification Authority` and `Apple Root CA` in **System** keychain.
- Verify TeamIdentifier shows up after signing:
  ```bash
  cp /bin/echo /tmp/mm_echo
  codesign --force --sign "Developer ID Application: Patrick Flynn (GK88Z35UG6)" /tmp/mm_echo
  codesign --display --verbose=4 /tmp/mm_echo | egrep "TeamIdentifier|Authority"
  ```

Build + sign (run from `FusionMacroReorderer-electron`):

```bash
CSC_IDENTITY_AUTO_DISCOVERY=false npm run dist -- --dir

ID="Developer ID Application: Patrick Flynn (GK88Z35UG6)"
ENT="entitlements.plist"
APP="dist/mac-arm64/Macro Machine.app"

# Sign all Mach-O payloads with hardened runtime
find "$APP/Contents/Frameworks" -type f -perm -111 -print0 | while IFS= read -r -d '' f; do
  if /usr/bin/file -b "$f" | /usr/bin/grep -q "Mach-O"; then
    /usr/bin/codesign --force --timestamp --options runtime --sign "$ID" "$f"
  fi
done

# Seal frameworks + helper apps
find "$APP/Contents/Frameworks" -type d -name "*.framework" -print0 | while IFS= read -r -d '' d; do
  /usr/bin/codesign --force --timestamp --options runtime --sign "$ID" "$d"
done
find "$APP/Contents/Frameworks" -type d -name "*.app" -print0 | while IFS= read -r -d '' d; do
  /usr/bin/codesign --force --timestamp --options runtime --entitlements "$ENT" --sign "$ID" "$d"
done

# Sign the main app
/usr/bin/codesign --force --timestamp --options runtime --entitlements "$ENT" --sign "$ID" "$APP"
codesign --verify --deep --strict --verbose=2 "$APP"
```

Notarize + staple the `.app` (optional if you only ship a DMG):

```bash
ZIP="dist/mac-arm64/Macro Machine.zip"
ditto -c -k --keepParent "$APP" "$ZIP"

xcrun notarytool submit "$ZIP" --apple-id "patrick@stirlingsupply.co" --team-id "GK88Z35UG6" --password "<app-specific>" --wait
xcrun stapler staple "$APP"
xcrun stapler validate "$APP"
```

Create a DMG with Applications shortcut + fixed layout, then notarize:

```bash
STAGING="dist/mac-arm64/dmg-staging"
VOLUME_NAME="Macro Machine"
RW_DMG="dist/mac-arm64/Macro Machine-rw.dmg"
DMG="dist/mac-arm64/Macro Machine.dmg"

rm -rf "$STAGING"
mkdir -p "$STAGING"
cp -R "$APP" "$STAGING/"
ln -s /Applications "$STAGING/Applications"

rm -f "$RW_DMG" "$DMG"
hdiutil create -volname "$VOLUME_NAME" -srcfolder "$STAGING" -fs HFS+ -format UDRW "$RW_DMG"
hdiutil attach -readwrite -nobrowse -noverify -noautoopen -mountpoint "/Volumes/$VOLUME_NAME" "$RW_DMG" >/dev/null

osascript <<EOF
tell application "Finder"
  tell disk "$VOLUME_NAME"
    open
    set current view of container window to icon view
    set toolbar visible of container window to false
    set statusbar visible of container window to false
    set the bounds of container window to {200, 200, 780, 520}
    set theViewOptions to the icon view options of container window
    set arrangement of theViewOptions to not arranged
    set icon size of theViewOptions to 128
    set position of item "Macro Machine.app" of container window to {160, 240}
    set position of item "Applications" of container window to {480, 240}
    close
  end tell
end tell
EOF

hdiutil detach "/Volumes/$VOLUME_NAME" >/dev/null
hdiutil convert "$RW_DMG" -format UDZO -imagekey zlib-level=9 -o "$DMG"
rm -f "$RW_DMG"

xcrun notarytool submit "$DMG" --apple-id "patrick@stirlingsupply.co" --team-id "GK88Z35UG6" --password "<app-specific>" --wait
xcrun stapler staple "$DMG"
xcrun stapler validate "$DMG"
```

If `spctl` says “rejected / insufficient context” for a DMG, use `xcrun stapler validate` instead.

---

## Known macOS Quirks

- Unsigned DMG/apps show:
  - “`Macro Machine` is damaged and can’t be opened. You should move it to the Trash.”
- Workarounds for testers:
  - Right-click the app in `Applications` → **Open** → confirm.
  - Or, in Terminal:  
    `xattr -dr com.apple.quarantine "/Applications/Macro Machine.app"`
- Proper fix (not implemented yet):
  - Join Apple Developer Program (paid).
  - Configure code signing + notarization in `electron-builder`.

---

## Future Ideas

- Code signing & notarized macOS builds.
- Custom app icon and About dialog.
- Keyboard shortcuts for Published controls (move up/down, group, etc.).
- “Recent files” menu and small in-app help/legend for the visual cues.
