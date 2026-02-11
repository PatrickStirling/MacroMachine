# Advanced Linking Goals

## Context
The Update Data button needs a reliable way to locate the original `.setting` so Macro Machine can reload and overwrite it after data updates. Absolute paths break across machines; DRFX adds another layer. Below are robustness options with recommended phases.

## Option A: Dual Path + Auto-Resolve (Recommended First)
Store both an absolute path and a relative path.
- `exportPath`: full absolute path (current behavior)
- `exportRel`: path relative to Resolve Templates root

Reload order:
1. Try `exportPath`
2. If missing, resolve `exportRel` against local Templates root

Pros
- Backward-compatible
- Portable across machines when presets live inside Templates

Cons
- Slightly more metadata

## Option B: Dual Path + GUID Search (High Robustness)
Add a stable hidden ID to each macro.
- `FMR_UID` stored in hidden UserControl

Reload order:
1. `exportPath`
2. `exportRel`
3. Search Templates root for `.setting` containing `FMR_UID`

Pros
- Survives rename/move
- Fewer user prompts

Cons
- Heavier search step

## Option C: DRFX-Aware Rehydration (Full Automation)
If macro was exported as DRFX, store:
- `drfxPath`
- `presetName`

Reload flow:
1. Extract preset to temp
2. Apply reload
3. Repack into DRFX

Pros
- Most seamless for DRFX users

Cons
- More moving parts

## Option D: Fallback Prompt (Always Recoverable)
If nothing resolves, prompt user to select save-back location.
- Once selected, store new `exportPath` and `exportRel`

Pros
- Always recoverable
- Simple UX

Cons
- Manual step when needed

## Suggested Phases
Phase 1: Implement Option A + Option D fallback
Phase 2: Add Option B (GUID search)
Phase 3: Consider Option C if DRFX automation is a priority

## Metadata Sketch
Hidden control (existing style): `FMR_FileMeta`
Example JSON:
```
{
  "exportPath": "C:/.../Templates/Edit/Titles/Foo.setting",
  "exportRel": "Edit/Titles/Foo.setting",
  "uid": "fmr_8c7f...",
  "drfxPath": "C:/.../MyPack.drfx",
  "presetName": "Foo"
}
```

## Resolution Order (Proposed)
1. `exportPath`
2. `exportRel` + Templates root
3. `uid` search within Templates root
4. Prompt user to select save-back path

---

## Image Header Feature Checklist

## Decision
- Target rollout: **Option C (Hybrid)**.
- Default mode: `path`.
- Optional mode: `base64` with hard cap.

## MVP (Ship Now)
- Scope: path-based image headers only, no base64 entry in UI.
- Goal: stable import/export behavior with zero UI freezes.

### MVP Tasks
- [ ] Add Header Image editor in details drawer for `LabelControl`.
- [ ] Inputs: enable toggle, path field, alt text field.
- [ ] Validation: supported extensions (`png`, `jpg`, `jpeg`, `webp`, `svg`).
- [ ] Render markup as `<img src=\"...\">` inside `LINKS_Name`.
- [ ] Keep non-image label text fallback when invalid.
- [ ] Preserve existing labels that do not use image markup.
- [ ] Add lightweight diagnostics: image mode (`path`) and validation fallback reason.

### MVP Gates
- [ ] No freeze on import, edit, export.
- [ ] Exported preset loads in Fusion.
- [ ] Reopen/save loop x10 has no drift/corruption.
- [ ] Existing presets unchanged.
- [ ] Test matrix pass: no image, valid path, missing path, malformed `<img>` markup, Resolve roundtrip.

## Later (Post-MVP)

### Phase 2 - Base64 Support (Guarded)
- [ ] Add source mode selector: `path` / `base64`.
- [ ] Add base64 size cap (start: **200 KB** payload).
- [ ] If over cap, block with explicit error and keep prior valid value.
- [ ] Add parser guard for long quoted blobs (treat as opaque text).
- [ ] Skip expensive metadata passes on oversized label payloads.

### Phase 2 Gates
- [ ] Base64 under cap loads and exports.
- [ ] Base64 over cap fails safely with clear message.
- [ ] No parser stalls from large payload strings.

### Phase 3 - Portability Enhancements
- [ ] Add optional export-time asset copy helper.
- [ ] Rewrite image path to project-relative/preset-relative when possible.
- [ ] Add integrity checks for missing image assets on import.

## Diagnostics
- [ ] Log image mode (`path`/`base64`) and payload bytes.
- [ ] Log fallback reason (invalid path, malformed markup, over cap).
- [ ] Add one per-load image summary line in diagnostics.

## Full Test Matrix
- [ ] No image label.
- [ ] Path image (valid local file).
- [ ] Path image (missing file).
- [ ] Base64 image under cap.
- [ ] Base64 image over cap.
- [ ] Malformed `<img>` markup.
- [ ] Resolve roundtrip: export -> timeline -> Fusion copy -> MM import.

## Explicit Non-Goals (for now)
- [ ] No image upload manager.
- [ ] No automatic remote URL downloading.
- [ ] No DRFX embedded asset rewriting in v1.

## Execution Order (Recommended)
1. Implement MVP drawer UI + serializer/parser path handling.
2. Add MVP diagnostics and run MVP gates.
3. Ship MVP.
4. Implement Phase 2 base64 mode behind size cap.
5. Implement Phase 3 portability helpers only after Phase 2 is stable.
