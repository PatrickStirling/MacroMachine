# Group-Layer Controls Stabilization Checklist

## Goal

Make Macro Machine structurally safe and reliable around macro/group-layer controls before adding more features that depend on them.

This phase is about:

- parsing
- data modeling
- writing/export
- validation
- regression safety

It is not about shipping new preset-engine functionality yet.

## Current status

Completed or materially underway:

- separate parsed macro/group `UserControls` model
- macro-layer control classification (`system` / `known` / `unknown`)
- dedicated macro-layer writer path
- export-side structural validation and diagnostics
- preserve-only mutation guards for unknown macro-layer controls
- repeatable regression suite:
  - `FusionMacroReorderer/scripts/check-group-usercontrols-regression.js`
  - `group-layer-controls-regression-suite.md`
- clean no-op round-trip confirmed on small reference macros:
  - `BG_XF.setting`
  - `Published_Connect.setting`
  - `SSC_ScreenPump.setting`
  - `ProtoV3.setting`

Still to do:

- broader no-op round-trip regression process for larger/advanced macros
- remaining cleanup of direct/raw helper paths
- final UI boundary decisions before resuming major macro-layer feature work

## Phase 1: Define the internal architecture

### 1. Split control systems in the parse result

- Add a dedicated structure for macro/group `UserControls`
- Keep published controls in the existing published entries structure
- Do not model macro-layer controls as fake published entries

Target outcome:

- `published controls` and `group-layer controls` are separate first-class systems in MM

### 2. Define explicit control categories

- `published`
- `group_user_control`
- optionally later `tool_user_control`

Target outcome:

- every control MM touches has an explicit ownership/type category

### 3. Document ownership rules in code comments or internal docs

- published controls belong to macro `Inputs`
- group-layer controls belong to macro `UserControls`
- do not convert one into the other automatically

Target outcome:

- future code changes have a clear boundary to follow

Status:

- in progress
- internal docs added
- main writer boundary comments added in `main.js`

## Phase 2: Parsing foundation

### 4. Build a dedicated parser for macro/group `UserControls`

Parse and preserve:

- control id
- `INPID_InputControl`
- `LINKS_Name`
- `LINKID_DataType`
- `ICS_ControlPage`
- `IC_Visible`
- `INP_Default`
- `INPS_ExecuteOnChange`
- `BTNCS_Execute`
- list props like `CCS_AddString`
- common label/button/combo metadata

Target outcome:

- MM can fully read macro-level controls without treating them as published entries

### 5. Preserve raw text/literal fidelity where needed

For group-layer controls, keep enough original raw/literal data to safely round-trip:

- multiline script text
- escaped strings
- combo option lists
- HTML/image markup

Target outcome:

- no destructive reformatting of advanced group-layer controls

### 6. Distinguish group-level vs tool-level `UserControls`

- parse group-level `UserControls` separately from tool-level `UserControls`
- do not conflate them

Target outcome:

- MM can reason about macro-owned controls independently from node-owned custom controls

## Phase 3: Export/write foundation

### 7. Build a dedicated writer for macro/group `UserControls`

Create one responsible path for:

- locating the real macro-level `UserControls`
- inserting controls
- updating controls
- removing controls
- preserving order where appropriate

Target outcome:

- all macro-layer edits flow through one safe writer

### 8. Keep the published `Inputs` writer separate

- no macro-layer controls inside `rewritePrimaryInputsBlock(...)`
- no published controls written through the group `UserControls` writer

Target outcome:

- fewer cross-system bugs and less accidental corruption

### 9. Enforce one macro-level `UserControls` block

- detect duplicates
- normalize if possible
- refuse export if structure is invalid and cannot be safely repaired

Target outcome:

- MM never emits multiple top-level group `UserControls` blocks

### 10. Add a preserve-only mode for unknown group-layer controls

If MM does not fully understand a group-layer control:

- keep it
- do not reinterpret it
- write it back unchanged

Target outcome:

- advanced hand-authored macros remain safe in MM

## Phase 4: Validation and safety checks

### 11. Add structural export validation

Before write/export, validate:

- balanced braces
- exactly one macro-level `UserControls` block
- no duplicate control ids inside macro-level `UserControls`
- no synthetic group-layer controls leaking into published `Inputs`
- no malformed hidden/helper control blocks

Target outcome:

- MM fails safely instead of exporting broken macros

### 12. Add targeted diagnostics for group-layer controls

Log:

- group-level `UserControls` count
- parsed group-layer control count
- duplicate ids
- unknown control types
- write/normalize operations

Target outcome:

- debugging future issues is fast and concrete

### 13. Add hard guards around generic export

Generic export should not invent advanced scaffolding unless explicitly required.

Do not auto-create:

- preset payload controls
- helper script controls
- file-meta blocks
- utility buttons

unless the macro is explicitly using that system

Target outcome:

- plain macros stay stable

## Phase 5: Regression suite

### 14. Build a small reference test set

Include at least:

- plain macro with no macro-layer controls
- macro with group-level combo + script text
- macro with group-level button
- macro with hidden helper text control
- macro with header image control
- Proto sample

Target outcome:

- every structural change can be checked against known real-world cases

Status:

- done
- automated by `scripts/check-group-usercontrols-regression.js`

### 15. Define no-op round-trip tests

For each reference macro:

- import into MM
- export without functional changes
- re-import
- confirm structure still matches expectations

Target outcome:

- MM can safely round-trip advanced macros

Status:

- in progress
- small clean references now pass:
  - `BG_XF.setting`
  - `Published_Connect.setting`
  - `SSC_ScreenPump.setting`
- larger advanced reference now also passes:
  - `ProtoV3.setting`

### 16. Add "preserve unknowns" regression checks

Use hand-authored macros that contain:

- custom scripts
- helper controls
- unusual combo/button setups

Target outcome:

- MM does not strip or corrupt unfamiliar group-layer constructs

## Phase 6: UI planning before feature work

### 17. Decide UI ownership boundaries

Set explicit product rules for:

- Published Controls UI = published `Inputs`
- future Macro Utilities / System Controls UI = group `UserControls`

Target outcome:

- future feature work does not blur the two systems again

### 18. Decide what should remain hidden/internal

Examples:

- helper script text
- metadata
- preset payload storage
- header carrier controls

Target outcome:

- MM does not accidentally surface internal controls in the wrong UI

### 19. Decide how editable group-layer controls should be

Questions to answer:

- should MM expose all group-layer controls?
- should some be read-only?
- should some be preserve-only?

Target outcome:

- safer editing scope for the first implementation

## Phase 7: Only then revisit features

### 20. Revisit preset engine export only after Phases 1-6

When resumed:

- selector should be treated as a true group-layer control
- target published controls should remain normal macro `Inputs`
- payload/storage should use the dedicated group-layer writer

Target outcome:

- preset engine work happens on a stable foundation

### 21. Revisit header-image and other macro-utility systems on the same foundation

Potential systems:

- header image
- help/tutorial button
- preset selector
- hidden script storage
- metadata/config blocks

Target outcome:

- all macro-layer systems use the same safe infrastructure

## Minimum safe milestone

If this needs to be reduced to the smallest meaningful checkpoint, complete these first:

1. Separate macro/group `UserControls` from published `Inputs` in MM's data model
2. Build a dedicated parser for macro/group `UserControls`
3. Build a dedicated writer for macro/group `UserControls`
4. Enforce exactly one macro-level `UserControls` block
5. Add structural export validation

That is the minimum foundation before more feature work should be added.

## Working principle

No new feature should be allowed to mutate macro/group `UserControls` through the published-control pipeline again.

That rule should remain in place until the separate parser/writer system exists.


