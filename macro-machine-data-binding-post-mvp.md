# Macro Machine — Data Binding (CSV / Google Sheets) Enhancements (Post-Core)

Scope note: This document is **explicitly POST-CORE implementation**.

**Core implementation (assumed already done):**
- User provides a Google Sheet URL (Link Sharing: “Anyone with the link”) or a local CSV.
- App fetches sheet as CSV (or reads CSV locally).
- App parses data into an in-memory table.
- User can map **Column + Row** to macro-exposed variables.
- Manual **Sync/Reload** updates the cached table.

Everything below is **after** that core is stable.

---

## Goals for Post-Core
1) **Safety:** avoid silent breakage and make changes auditable/reversible.
2) **Scalability:** move from “one-off linking” to “data-driven systems.”
3) **Speed:** make syncing and applying changes fast and predictable.

---

## Post-Core Feature Set (Recommended Order)

### 1) Schema + Validation (Type system)
**Problem:** CSV/Sheets data is messy. Without types, you get silent failures.

**Feature:** Let users define (or infer) a schema:
- Column type: `string | number | color | boolean | date | enum`
- Optional: default value, required flag, min/max (numbers), regex (strings)

**On Sync:**
- Validate every row.
- Surface errors/warnings:
  - “Row 12: `Price` expected number, got `$12,xx`”
  - “Row 3: `Color` not valid hex.”

**UX:**
- Errors panel + inline row highlighting.
- “Apply anyway” only if user confirms.

---

### 2) Token Lookup Mode (Key/Value design system)
**Problem:** Row-index mapping breaks when users reorder rows.

**Feature:** Support a second mapping mode:
- Sheet/tab with: `Key | Value | Type (optional)`
- Macro variables reference **keys** (e.g., `primary-color`, `corner-radius`).

**Behavior:**
- On Sync, build a dictionary `{key -> typedValue}`.
- If key missing: show warning and use fallback.

**Why it’s valuable:**
- Links are stable even when rows move.
- Enables sharable “design system” style presets.

---

### 3) Diff Preview + Safe Apply
**Problem:** Users fear Sync because it might break things.

**Feature:** On Sync, show:
- Added keys/rows
- Removed keys/rows
- Changed values (before/after)

**Controls:**
- Apply all changes
- Apply selected changes
- Ignore a change (do not apply)

---

### 4) Versioning / Snapshots / Rollback
**Problem:** Bad data syncs happen.

**Feature:** Store a snapshot per Sync:
- Timestamp
- Source URL / file hash
- Parsed data
- Validation results

**UX:**
- “Revert to last known good”
- “Revert to snapshot…”

---

### 5) Row Selection Strategies (beyond fixed row)
**Problem:** Fixed row is too limiting for real template workflows.

Add options:
- **Active row index**: a single variable the user can change.
- **Key match**: “Find row where `ID == X`”
- **Per-clip / per-instance selection**: index based on clip number, duplicator index, etc.

---

### 6) Asset Binding (images/video/audio)
**Problem:** Text/number binding is great; binding assets makes it a “system.”

Feature:
- Columns can reference:
  - local paths
  - URLs
  - Drive links (later)

Behavior:
- Download/cache assets locally.
- Validate presence and file type.

---

### 7) Batch Render / Batch Export Harness
**Problem:** Once data binding exists, users will want “generate 200 variants.”

Feature:
- “Render rows 1–200”
- Progress tracking
- Failure report by row (“Row 57 missing asset”) and retry.

---

## Implementation Notes / Guardrails
- Default to **manual sync** (explicit Reload button).
- Be strict about “public sheet only” unless/until OAuth is implemented.
- Prefer stability:
  - if source is unreachable, keep last cached data and warn.
  - never silently apply partially-validated changes.

---

## Open Questions (for later)
- Do we ever support private sheets via OAuth?
- Do we support multi-tab datasets and joins?
- How do we store mappings in a way that survives project moves?

---

## TL;DR
After the core CSV/Sheets mapping works, the next big wins are:
1) schema/validation
2) key/value token lookup
3) diff + rollback
4) smarter row selection
5) asset binding
6) batch render
