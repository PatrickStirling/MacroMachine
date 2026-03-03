# Group-Layer Controls Regression Suite

## Purpose

This is the repeatable structural regression suite for macro/group-level `UserControls`.

Use it after changes to:

- macro-level `UserControls` parsing
- macro-level `UserControls` writing
- export normalization
- hidden macro-layer controls
- launcher helpers
- presets or other group-layer systems

This suite is intentionally structural. It is meant to catch:

- duplicate top-level macro `UserControls` blocks
- duplicate macro-layer control ids
- unexpected macro-layer control type drift
- accidental changes to known reference macros

It is not a UI test and it is not a Fusion runtime test.

## Canonical samples

### Clean samples

- `Published_ALL.setting`
- `Published_ALL_Modifiers.setting`
- `Published_Connect.setting`
- `BG_XF.setting`
- `Cleanup testing\SSC_ScreenPump.setting`

Expected:

- `0` macro-level `UserControls` blocks
- `0` macro-level controls

Note:

- `SSC_ScreenPump.setting` is also a no-op round-trip reference now.
- It contains tool-level `UserControls` and one authored macro-level label control.
- Current expectation for no-op export is byte-identical preservation.

### Known macro-layer reference

- `Macro_ref\_STORE\64_ProtoV3\V1.01\Edit\Generators\Stirling Supply Co\ProtoV3.setting`

Note:

- `ProtoV3.setting` is now also a passing no-op round-trip reference.
- Current expectation for no-op export is byte-identical preservation.

Expected:

- `1` macro-level `UserControls` block
- `3` macro-level controls
- control ids:
  - `Preset`
  - `ScriptTxt`
  - `Walkthrough`
- no duplicate ids
- no unknown macro-layer input control types

### Known broken sentinel

- `Pre_Test18.setting`

Expected:

- `2` macro-level `UserControls` blocks
- duplicate id: `FMR_FileMeta`

This file is intentionally kept in the suite as a sentinel for the failure class we hit during the preset-engine work. It is not a clean sample.

## Current known macro-layer control types in archive

Archive sweep result from `_STORE`:

- `ButtonControl`
- `ComboControl`
- `SliderControl`
- `TextEditControl`

No other macro-level `INPID_InputControl` types were found in the scanned archive at the time this suite was created.

## How to run

From `FusionMacroReorderer`:

```powershell
node .\scripts\check-group-usercontrols-regression.js
```

Verbose listing:

```powershell
node .\scripts\check-group-usercontrols-regression.js --verbose
```

## Pass criteria

The script should:

- print the sample summary table
- exit `0`
- end with:
  - `Group UserControls regression check passed.`

If it exits non-zero, treat that as a structural regression until the expectation set is deliberately updated.

## When to run

Run this after:

- changing `parseGroupUserControls(...)`
- changing `hydrateGroupUserControlsState(...)`
- changing `upsertGroupUserControl(...)`
- changing `removeGroupUserControlById(...)`
- changing `normalizeGroupUserControlsBlocks(...)`
- changing export-side hidden macro-layer control persistence
- changing launcher helpers that can touch group-level controls

## Notes

- This suite is deliberately small. It is meant to be fast and repeatable.
- Keep the clean samples stable.
- Only change the expectation set when a structural change is intentional and understood.
