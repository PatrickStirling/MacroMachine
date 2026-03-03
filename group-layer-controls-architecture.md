# Group-Layer Controls Architecture

## Why this matters

Fusion macros support controls that live directly on the macro/group itself, not on any inner node. This is a separate system from normal published controls and needs to be treated that way in Macro Machine.

This is important for:

- preset selectors
- hidden script storage
- helper buttons
- walkthrough links
- metadata fields
- header-only utility controls

## The two control planes

### 1. Published controls: `Inputs = ordered()`

These are macro controls that proxy real controls on tools inside the macro.

They are written as `InstanceInput` blocks and usually include:

- `SourceOp`
- `Source`
- `Name`
- `Page`
- `Default`
- `ControlGroup`

This is the normal publish workflow MM already understands well.

Example:

- `ProtoV3.setting` uses many `InstanceInput` entries in the macro `Inputs` block to expose inner tool controls like `BG_Red1`, `SourcePosition1`, etc.

## 2. Macro-owned controls: `UserControls = ordered()`

These controls belong directly to the macro/group itself.

They are not proxies by default and do not need:

- `SourceOp`
- `Source`

These are true macro-layer controls.

Example from `ProtoV3.setting`:

- `Preset` is a macro-level `ComboControl`
- `ScriptTxt` is a macro-level `TextEditControl`
- `Walkthrough` is a macro-level `ButtonControl`

All of those live in the group `UserControls` block, not the macro `Inputs` block.

## What macro-layer `UserControls` can do

Macro-layer controls can be:

- visible UI controls
- hidden helper/storage controls
- buttons that execute actions
- script launchers
- labels and internal organization controls

Examples:

- visible combo:
  - `Preset`
- hidden script field:
  - `ScriptTxt`
- help button:
  - `Walkthrough`

## What `OnChange` can do at the group layer

A group-layer control can:

- read itself with `tool:GetInput("Preset")`
- run stored script text with `fusion:Execute(tool:GetInput("ScriptTxt"))`
- reach inner tools with `comp:FindTool("ToolName")`
- manipulate published macro inputs if those published inputs exist

This makes macro-layer controls ideal for:

- preset engines
- mode selectors
- reveal/hide systems
- internal config/state

## The critical distinction

These two systems are not interchangeable.

### Use `Inputs` when:

- the control is meant to publish an inner tool input
- the macro control should act like a direct exposed parameter
- ordering/grouping belongs in the published control list

### Use group `UserControls` when:

- the control belongs to the macro itself
- the control is utility/UI logic
- the control stores script or metadata
- the control drives behavior instead of directly exposing a single inner input

## What Proto proves

Proto uses both systems successfully in the same macro:

- `Inputs` for published controls
- group `UserControls` for presets, helper scripts, and buttons

That means the correct architecture is not "pick one system." It is "use the right system for the right job."

## Why MM has been fragile here

MM has mostly modeled controls as published entries with:

- `sourceOp`
- `source`
- published ordering
- instance-input rewriting

That model fits `Inputs`.

It does not fit pure macro-layer `UserControls`.

When MM tries to force macro-layer controls into the published-control pipeline, it causes:

- duplicate controls
- invalid `InstanceInput` generation
- stale `SourceOp` references
- broken export structure
- extra `UserControls` blocks

## Structural rule going forward

Macro Machine should treat these as separate first-class systems:

### Published entries

- backed by macro `Inputs`
- tied to inner nodes/tools

### Group user controls

- backed by macro `UserControls`
- owned directly by the macro
- only tied to inner nodes if their script explicitly does so

## Export safety rules

1. A macro/group should have exactly one group-level `UserControls` block.
2. Macro-layer controls must not be auto-converted into `InstanceInput`s.
3. Hidden metadata/script/helper controls should stay in group `UserControls`.
4. Standard published controls should stay in `Inputs`.

## Best-fit systems for future features

### Good fits for group `UserControls`

- Preset selector
- Preset script text
- Hidden preset payload
- Walkthrough/help button
- Header image carrier
- Hidden metadata/config

### Good fits for `Inputs`

- Sliders
- checkboxes
- text fields
- colors
- grouped color controls
- any direct publish of a real inner tool parameter

## Presets engine implication

If presets are revisited later, the stable architecture is:

- preset selector lives in group `UserControls`
- optional script storage lives in group `UserControls`
- preset payload lives in group `UserControls` or MM-only until export is ready
- the controls being changed remain normal published controls in `Inputs`

That is the key split:

- selector/control logic on the macro itself
- target values still point to normal published inputs

## Recommended MM architecture later

### Parser

Maintain separate structures for:

- published controls
- group-level user controls

### Writer

Maintain separate writers for:

- macro `Inputs`
- macro `UserControls`

### UI

Treat macro-layer controls as a separate category from published controls.

They should not be shoved through the same assumptions as node-backed published entries.

## Bottom line

Fusion clearly supports macro-owned controls at the group layer.

They are powerful and useful, but they are a different system from published controls.

MM should model them separately, write them separately, and use them deliberately for macro-level behavior.
