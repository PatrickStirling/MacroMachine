# Macro Machine (Fusion Macro Reorderer)

Macro Machine is an offline tool that reads DaVinci Resolve / Fusion `.setting` macro files, parses the published controls in the top-level `GroupOperator` / `MacroOperator`, and lets you reorganize, annotate, and augment those controls before exporting a cleaned `.setting`.

There are two entry points:

- `FusionMacroReorderer/` – static HTML/CSS/JS version that runs entirely in the browser.
- `FusionMacroReorderer-electron/` – Electron shell that adds native open/save dialogs, clipboard helpers, and packaged installers.

Everything stays on your machine; no uploads occur.

## Key Features

- **Published Controls pane**
  - Drag-and-drop or click ▲/▼ to reorder `InstanceInput` entries.
  - Inline editing of the display name and macro name.
  - Label groups (green) and color/control groups (orange) with clear visual cues.
  - Remove controls, undo/redo, reset-to-original-order.

- **Nodes pane**
  - Parses tools/modifiers from the macro’s `Tools = ordered()` block.
  - Uses local catalogs (`FusionNodeCatalog.cleaned.json`, `FusionModifierCatalog.json`) to show each node’s publishable controls.
  - Batch select, publish, or drag controls directly into the Published list.
  - “Hide replaced controls” filter to focus on unused inputs.

- **URL launcher helper**
  - Detects ButtonControls with empty or known `BTNCS_Execute` code.
  - Lets you attach a URL that becomes a cross-platform launcher (Windows `start`, macOS `open`, Linux `xdg-open`) when you export.

- **Validation tools**
  - `Validate` counts braces, checks `InstanceInput` totals, and flags missing `Page`/`Name`.
  - `Fix missing pages` adds `Page = "Controls"` to entries that lack one.

- **Diagnostics**
  - Toggleable diagnostics log (hidden by default; enable it via the Electron menu or the new “Show diagnostics” button in the Published pane) that records parser events and errors.

## Using the Browser Version

1. Open `FusionMacroReorderer/index.html` in Chrome, Edge, or another Chromium-based browser.
2. Drag & drop a `.setting` file or click “Select .setting file”.
3. Optional: import the node and modifier catalogs if you want Nodes-pane metadata (Electron auto-loads them).
4. Reorder controls, group them, or publish new controls from the Nodes pane.
5. Click “Export reordered .setting” to download the updated macro, or “Export to clipboard” for copy/paste workflows.

All edits touch only the selected macro’s `Inputs = ordered()` block (plus any optional URL launchers you add). Each `InstanceInput { ... }` block is otherwise preserved verbatim, including comments.

## Electron (Native) Workflow

```
cd FusionMacroReorderer-electron
npm install
npm start
```

- `npm start` loads the sibling `FusionMacroReorderer/index.html` in a BrowserWindow with Node integration enabled.
- Native File menu exposes Open…, Save, Import/Export Clipboard, Diagnostics toggle, and standard reload/devtools items.
- `npm run dist` uses `electron-builder` to package DMG (macOS) and NSIS (Windows) installers. macOS artifacts are unsigned/not notarized, so expect Gatekeeper prompts.

## Troubleshooting

- Make sure you’re opening a macro (`GroupOperator` / `MacroOperator`). Plain tool `.setting` files don’t have the published `Inputs` list this app expects.
- If a macro contains multiple nested groups, the first valid `Inputs = ordered()` block is used. You can strip out other content and try again if parsing fails.
- Keep the node/modifier catalogs close by—they supply human-friendly control names in the Nodes pane; without them, you’ll see only raw IDs.

Sample `.setting` files in this folder (`SSC_*.setting`, etc.) are handy for testing layout and catalog parsing. Feel free to duplicate them and experiment. Once you’re happy with the order, export and drop the result into your Fusion macros directory as usual.
