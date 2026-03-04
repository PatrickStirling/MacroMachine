# Macro Machine UX Roadmap Notes

This is not a commitment list. It is a structured idea bank for future UX work now that Macro Machine has grown beyond a simple reorder/publish tool.

The goal is to identify improvements that could make MM feel cleaner, faster to scan, and more intentional without losing the power it has accumulated.

## Design Tension

MM currently spans multiple roles:

- Core macro editor
- Nodes-panel publishing browser
- Advanced authoring environment
- Automation/systems layer

That breadth is powerful, but it also creates UX pressure. Features can start to compete for visual attention unless the app gets more deliberate about structure.

## Small Wins

These are modest improvements that could make the app feel better quickly without redesigning core workflows.

### 1. Stronger Structural Distinction

Make it easier to visually distinguish:

- Macro Root
- Normal nodes
- Modifiers
- Dynamic slot sections like `Layer 1` or `Text 2`
- Utility/system nodes

Some of this is already happening with color and badges. More consistency here would help large macros scan faster.

### 2. Better Control Intent in the Published Panel

Help users understand what a published control is for:

- User-facing control
- Utility/helper control
- Layout control like label or separator
- System control like `Preset`

This could be done through subtle UI treatment rather than more text.

### 3. Temporary Focus Tools

Useful lightweight helpers for larger macros:

- Show only controls from selected node
- Highlight recently edited controls
- Pin important controls temporarily
- Filter by page, type, or source node

These would improve day-to-day editing without changing the overall structure.

### 4. Better Empty/Initial States

A few areas still depend too much on user knowledge:

- Presets UI
- Macro-root controls
- Dynamic nodes
- System controls

Small empty-state hints or labels could make those features feel more discoverable.

## Medium Reworks

These would change how sections of MM feel, but still within the current overall app shape.

### 5. Authoring-Focused Nodes View

Add a cleaner authoring mode for the nodes panel that prioritizes:

- Publishable controls
- Relevant grouped controls
- High-value metadata
- Less parser completeness clutter

This would be especially valuable on complex macros.

### 6. Better Large-Macro Navigation

The nodes panel is strong, but dense macros still become difficult to navigate.

Possible improvements:

- Focus or solo one node temporarily
- Better global expand/collapse behavior
- More reliable remembered section state
- Quick jump tools between published controls and source nodes

### 7. Dynamic Node UX Refinement

`MultiMerge`, `MultiText`, and `MultiPoly` are much stronger now, but could feel more purpose-built with:

- Better slot summaries
- Slot counts
- Per-slot publish counts
- Clearer distinction between common controls and repeated slot controls

### 8. Detail Drawer Parity

Continue eliminating differences between control types so everything feels equally editable:

- Macro-root controls
- Labels
- Buttons
- Combos
- Expressions
- Grouped controls

This is less flashy but very high leverage.

## Bigger Redesign Opportunities

These are larger conceptual shifts, not immediate implementation suggestions.

### 9. Clearer Separation Between Editing and Systems

MM now has two identities:

- Macro editor
- Macro systems/automation environment

The interface could eventually reflect that more clearly so advanced systems like presets or data linking feel layered on top of editing, not mixed into the same surface indiscriminately.

### 10. Better Mode Awareness

The app could feel more intentional if it more clearly communicated what the user is doing at a given moment:

- Editing controls
- Browsing source nodes
- Building systems
- Managing automation/runtime features

This does not mean multiple apps. It means stronger context cues and cleaner grouping.

### 11. Workflow-Centered Layout Thinking

Instead of organizing only by data type or feature bucket, more of the app could be framed around tasks:

- Build macro UI
- Refine/organize published controls
- Add system behavior
- Validate/export

That could improve clarity as the feature set continues to grow.

## Suggested Priority Order

If this ever becomes active work, the safest order is probably:

1. Stronger structural distinction
2. Published panel intent/focus tools
3. Large-macro navigation improvements
4. Dynamic node UX refinement
5. Broader editing vs systems separation

## Current Recommendation

Do not redesign yet just because many ideas exist.

The better approach is:

- keep noticing friction during real use
- collect repeated pain points
- promote only the best ideas into implementation

That keeps MM grounded in actual workflow value instead of abstract UX theory.
