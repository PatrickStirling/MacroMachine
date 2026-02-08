# Cavalry (motion graphics) — Google Sheets integration (what it appears to do)

Sources (official docs):
- Spreadsheet utility: https://docs.cavalry.scenegroup.co/nodes/utilities/spreadsheet/
- Google Sheets Asset: https://docs.cavalry.scenegroup.co/user-interface/menus/window-menu/assets-window/google-sheets-asset/

## Key behavioral details (from docs)

### Permissions model
- Cavalry requires Google Sheets **Link Sharing = “Anyone with the link”**.
- Quote: “It’s not yet possible to access Sheets with restricted permissions.”
- This strongly implies Cavalry **does not** use OAuth to access private sheets.

### Data model in Cavalry
- A Google Sheet is imported as a **Google Sheets Asset** in Cavalry’s Assets Window.
- That asset can be connected to a **Spreadsheet Utility**, which reads data.
- Spreadsheet Utility outputs **one column at a time** (you create one Spreadsheet Utility per column).

### Update / refresh behavior
- The docs explicitly include a **Reload** option:
  - “Reload — Reload the Google Sheet data. This is required whenever the Google Sheet is changed.”
- That suggests Cavalry does not auto-poll continuously; it fetches data and you manually refresh.

### Selecting a tab (sheet)
- They rely on the URL (including `gid=...`) to reference a specific tab:
  - “To reference a specific sheet/tab, copy the url from the address bar with that sheet/tab open.”

### Parsing & typing
- They expose “Spreadsheet Settings” like delimiter/quote/decimal separator.
- Docs warn about mixed-type columns and suggest setting Google Sheets formatting (e.g., Plain Text) to control parsing.

## What this implies about implementation

Given:
- no restricted sheets,
- requires “anyone with link,”
- tab selection via `gid`,
- manual reload,

…the most likely implementation is **unauthenticated HTTP fetch** of the sheet in a machine-readable form (CSV-like), keyed by `spreadsheetId` + `gid`.

Typical pattern in the wild:
- `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/export?format=csv&gid=<GID>`

Cavalry’s delimiter requirement (“must be comma for Google Sheets Assets”) aligns with CSV export.

## Other Cavalry features worth stealing for Macro Machine

(Quick scan of Cavalry Utilities docs; these are concepts/UX patterns that translate well.)

### Spreadsheet Lookup (key/value design tokens)
Source: https://docs.cavalry.scenegroup.co/nodes/utilities/spreadsheet-lookup/
- Pattern: treat a sheet like a **design token table** (`Key`, `Value`) and reference values by key.
- Macro Machine idea: “Token lookup” values for macros (e.g., `corner-radius`, `primary-color`, `font-size`) so users can reorganize rows without breaking links.

### Column-only Spreadsheet output
Source: https://docs.cavalry.scenegroup.co/nodes/utilities/spreadsheet/
- Cavalry outputs **one column per node**; you add multiple nodes for multiple columns.
- Macro Machine idea: keep mapping simple by default: one mapping per column, then provide higher-level presets that bundle mappings.

### Manual reload/sync
Source: Google Sheets Asset docs (Reload required): https://docs.cavalry.scenegroup.co/user-interface/menus/window-menu/assets-window/google-sheets-asset/
- Clear, explicit “Reload” reduces surprises and avoids background polling.
- Macro Machine idea: Sync button + “last synced” + diff preview.

### Animation Control (remap existing keyframes with a single % slider)
Source: https://docs.cavalry.scenegroup.co/nodes/utilities/animation-control/
- Concept: drive an existing animation curve by a 0–100% control (time scaling / scrubbing).
- Macro Machine idea (Resolve/Fusion): a “progress” control that can drive multiple macro animations consistently.

### Scheduling Group (procedural time offsets / sequencing)
Source: https://docs.cavalry.scenegroup.co/nodes/utilities/scheduling-group/
- Concept: group layers and procedurally offset them in time; optionally auto-sequence with overlap/gaps.
- Macro Machine idea: higher-level “stagger/sequence” tools for macro instances (especially for template systems).

### Asset Array (indexable asset selection)
Source: https://docs.cavalry.scenegroup.co/nodes/utilities/asset-array/
- Concept: an indexable list of assets, with deterministic or randomized selection.
- Macro Machine idea: “asset sets” (e.g., 10 textures/icons) with index/random controls.

---

## Product takeaways for Macro Machine

### Cavalry-style integration (no OAuth)
If you want a Cavalry-like experience without OAuth:
- Restrict to **public link-sharing sheets** (Cavalry requires “Anyone with the link”).
- Accept a **Google Sheet URL** (use the browser address-bar URL to capture the correct `gid` for the tab).
- Fetch data via **CSV export** (likely):
  - `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/export?format=csv&gid=<GID>`
- Parse locally and provide a **manual Sync/Reload** action.

### UX recommendation: column/row mapping (no A1 notation)
Suggested mapping UI:
- Pick **Column** from header names (row 1)
- Pick **Row** by:
  - fixed index (row N), or
  - key match (find row where ID == X)
Then map to Macro Machine variables (e.g., Title, Price, AccentColor).

### Resolve integration note (important)
Fusion macros typically can’t fetch HTTP/CSV directly. So even with a Resolve-side “Sync Now” button, something outside the macro must:
- fetch/parse the sheet, then
- push values into Resolve/macros (e.g., write a local cache file the macro reads, or update macro controls).

### OAuth note
If you want private sheets, you need OAuth; Cavalry’s docs indicate they intentionally avoided that complexity.
