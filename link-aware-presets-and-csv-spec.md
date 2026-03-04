# Link-Aware Presets and CSV Integration Spec

This is a future-facing design note, not an implementation commitment.

The purpose of this document is to describe how Macro Machine could allow the Presets system and the CSV/Data Linking system to work together instead of flattening one into the other.

## Core Idea

Presets should eventually become **link-aware**, not just value-aware.

That means a preset should be able to store, per control:

- a literal value
- or a live link descriptor
- or a live link descriptor with preset-specific row/key override

This would allow preset switching to preserve live linked behavior instead of replacing all links with resolved snapshot values.

## Why This Matters

Without link-aware presets, the likely behavior is:

- source macros contain linked controls
- preset builder reads those macros
- MM stores only the resolved values
- preset switching restores those literal values
- the original CSV linkage is effectively lost

That may work visually, but it undermines the value of the data-linking system.

If presets become link-aware, then:

- linked controls can remain linked after preset creation
- preset switching can restore link states, not just values
- `Update Data` can still refresh the current preset correctly

## Conceptual Model

There are two different axes of state:

### 1. Preset State

Preset state changes authored behavior and appearance:

- style choices
- enabled features
- look/layout variations
- control defaults and configuration

### 2. Data Link State

Data-link state provides live external content:

- text
- numbers
- colors
- timings
- values resolved from CSV or other linked sources

The system becomes powerful when MM can preserve both.

## What a Preset Should Store

Instead of storing only:

- `Control A = 0.5`
- `Control B = "Hello"`

MM should be able to store:

- `Control A = literal 0.5`
- `Control B = linked to CSV column Name`
- `Control C = linked to CSV column Score using preset-specific row override`

In other words, a preset should store **control state**, not just **resolved values**.

## Suggested Data Model

Each scoped control in a preset could use a model like:

```json
{
  "controlKey": "Text1.StyledText",
  "mode": "linked",
  "link": {
    "provider": "csv",
    "column": "Name",
    "rowMode": "fixed",
    "rowValue": "7"
  },
  "fallbackValue": "Patrick"
}
```

Or:

```json
{
  "controlKey": "XF_Shake.Blend",
  "mode": "literal",
  "value": 0.35
}
```

The important part is that preset storage distinguishes between:

- `literal`
- `linked`

and can optionally include:

- a fallback resolved value for preview/debugging

## Preset Application Behavior

When the user switches presets:

1. Apply all literal controls as values.
2. Restore all linked controls as live links.
3. Restore any preset-specific row/key overrides for those links.
4. Optionally trigger a linked-data refresh.

This is the key difference from a snapshot-only preset system.

## Update Data Behavior

In a link-aware preset system, `Update Data` should operate on the **currently active preset state**.

That means:

- if the active preset defines a control as linked, `Update Data` refreshes it
- if the active preset defines a control as literal, `Update Data` leaves it alone
- if the active preset changes the row/key context for linked controls, `Update Data` resolves using that current preset context

This preserves the value of the update workflow instead of making presets and CSV fight each other.

## Preset Build Workflow With Linked Variants

Ideal workflow:

1. User creates multiple macro variants in Fusion.
2. Some controls are literal.
3. Some controls are CSV-linked.
4. MM imports those variant macros.
5. Preset builder detects both literal state and link state.
6. MM creates a preset pack that preserves both kinds of control state.
7. Exported macro can switch presets while retaining live links.
8. `Update Data` continues to work for the active preset.

## Supported Modes

This feature could be implemented in phases.

### Phase 1: Freeze Values

Store only resolved values.

Pros:

- simplest
- most portable

Cons:

- loses live link behavior

This is the current conceptual baseline and should not be the long-term goal for linked workflows.

### Phase 2: Preserve Live Links

Store whether a control is linked and what it is linked to.

Pros:

- preserves CSV behavior
- keeps `Update Data` meaningful

Cons:

- requires link-aware preset storage and apply logic

This is the most important future milestone.

### Phase 3: Preserve Links With Preset-Specific Context

Allow different presets to change row/key context while keeping the same linked control structure.

Examples:

- Preset A uses row 2
- Preset B uses row 7
- both still link `Name` and `Score`

Pros:

- extremely powerful
- ideal for productized preset/data workflows

Cons:

- most complex

## Important Edge Cases

These need explicit thought before implementation.

### 1. Literal in One Preset, Linked in Another

Same control may be:

- literal in Preset A
- linked in Preset B

System must restore the correct state on switch.

### 2. Same Control, Different Linked Columns

Example:

- Preset A: `StyledText -> Name`
- Preset B: `StyledText -> TeamName`

System must restore the link descriptor, not just the resolved value.

### 3. Same Link, Different Row/Key Context

Example:

- Preset A linked to `row 1`
- Preset B linked to `row 5`

This should be supported once preset-specific link context exists.

### 4. Resolved Preview vs Actual Source of Truth

Builder may see current resolved values while importing variants.

The system must avoid confusing:

- “what was visible when imported”

with:

- “what should be stored as the preset definition”

### 5. Export Stability

Restoring link-aware presets must not destabilize export.

This is especially important given the prior fragility around generated controls and group-layer behavior.

## Recommended Direction

If this feature is ever implemented, the guiding rule should be:

**Presets must store link state, not just resolved values.**

That is the difference between:

- a nice snapshot feature

and:

- a genuinely powerful integrated product system

## Suggested Implementation Order

If this becomes active work later, the safest order is:

1. Define a formal preset control-state model that distinguishes literal vs linked.
2. Teach the preset builder to detect link descriptors in imported variants.
3. Store that information in preset pack data.
4. Update preset application logic to restore links as links.
5. Make `Update Data` operate against active preset link state.
6. Only after that, add preset-specific row/key overrides.

## Current Recommendation

Do not build this immediately.

There is still broad generic testing to do across MM, and this feature touches:

- presets
- linking
- export behavior
- runtime update behavior

It is a high-value future direction, but it should be approached after the current system is more broadly validated.
