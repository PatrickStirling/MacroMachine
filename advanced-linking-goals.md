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
