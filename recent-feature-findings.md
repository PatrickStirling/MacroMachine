# Recent Feature Findings (March 2026)

## Shipped in `v0.2.52`
- Added `Explorer Here` button in Macro Explorer to open the current folder in OS explorer.
- Added node-to-published focus behavior:
  - Clicking already-published controls in the Nodes panel now focuses the matching Published control entry.
- Improved Published panel drag/drop behavior:
  - Dragging above/below the list now clamps to top/bottom drop positions instead of no-op.

## Saved Research: Path Mask Draw Modes
- Path mask tools (e.g. `PolylineMask`) store mode state with:
  - `DrawMode`
  - `DrawMode2`
- Confirmed real values in reference sample:
  - `ClickAppend`, `Freehand`, `InsertAndModify`, `ModifyOnly`, `Done`
- Product idea (saved):
  - Expose a macro control (combo/multibutton) for draw mode so users can change path edit mode from Edit page inspector (without Fusion viewer toolbar).

## Saved Research: Contours “Versions” System
- `Contours.setting` uses GroupOperator-native snapshot system:
  - `CurrentSettings = N`
  - `CustomData.Settings` with numbered slots (`[1]..[N-1]`) storing full internal snapshots.
- Key behavior:
  - Not scoped; captures full state.
  - Numeric slot-based, limited workflow.
- Takeaway for MM:
  - Useful compatibility target for import/inspection.
  - Different from MM scoped presets engine by design.

## Saved Research: `MultiButtonControl`
- `MultiButtonControl` outputs number-style index values (same family as combo behavior, 0-based style usage).
- Found control attributes in sample:
  - Button list: repeated `{ MBTNC_AddButton = "..." }`
  - Display/layout flags: `MBTNC_ShowBasicButton`, `MBTNC_ShowToolTip`, `MBTNC_ShowName`, `MBTNC_StretchToFit`, `MBTNC_ButtonWidth`, `MBTNC_ButtonHeight`
- Current MM gap:
  - MM has stronger combo-specific handling than multibutton-specific handling, so some expected display differences are limited today.

## Next Candidate Work Items
- High priority: revisit and finish the dropped `Effect Input` work:
  - restore full parser/export support
  - verify publish/edit behavior in MM UI
  - run regression tests on real macros that rely on effect inputs
- Add MVP support for path-mask draw mode controls in Published panel workflow.
- Extend MM control handling parity for `MultiButtonControl`:
  - Parsing/display options
  - Detail drawer editing support
  - Export preservation/update for MBTNC props
- Optional: add “native Versions” read-only/bridge view when `CurrentSettings/CustomData.Settings` are present.
