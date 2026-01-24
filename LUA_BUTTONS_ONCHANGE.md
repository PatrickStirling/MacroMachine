# Fusion Lua for Buttons + On-Change

Quick cheat sheet for Lua snippets used by Button Execute (BTNCS_Execute) and On-Change (INPS_ExecuteOnChange)
in Macro Machine exports.

## Where it runs
- Button Execute: runs once when the button is pressed.
- On-Change: runs when the control changes (and can re-run if your script changes inputs).

## What this does inside Fusion
- Changes happen immediately in the macro: inputs update, expressions reevaluate, and connected tools react in real time.
- Setting an input can change what controls are visible or active (for example, switching a mode can reveal different sliders).
- For viewers, the timeline/render result updates as soon as the tool values change.

## Basics
```lua
-- comment
local x = 1
local s = "text"
```

## If / elseif / else
```lua
if condition then
  -- do something
elseif other_condition then
  -- do something else
else
  -- fallback
end
```
Use this to branch behavior based on tool inputs (e.g., if a switch is on, set a parameter; otherwise reset it).

## Comparisons + logic
```
==  ~=  <  >  <=  >=
and  or  not
```
Comparisons return true/false. `==` equals, `~=` not equal, the others are numeric comparisons.
Combine them with `and` / `or` / `not` to build checks like:
`if source == 0 and enabled == 1 then ... end`

## Find tools and read/write inputs
```lua
local tool = comp:FindTool("Transform1")
if not tool then return end

local blend = tool:GetInput("Blend")
tool:SetInput("Blend", 0.5)
```
`FindTool` grabs a node by its tool ID. `GetInput` reads the current value; `SetInput` updates it.
This directly changes the node in the Fusion comp and updates the viewer output.

You can also write directly to inputs in many cases:
```lua
tool.Blend = 0.5
tool["Blend"] = 0.5
```
`SetInput` is just the most consistent/safe option when IDs are unusual or you want clearer intent.

Likewise, reading inputs can be shortened:
```lua
if tool.Blend > 0.5 then
  -- ...
end
```
`GetInput` is clearer and safer when you’re not sure a field exists or the ID is odd.

Point-type inputs typically use a table:
```lua
tool:SetInput("Center", {0.5, 0.5})
```
Point values are X/Y pairs. `{0.5, 0.5}` centers a point-based control.

## Example: switch-driven behavior
```lua
local sw = comp:FindTool("Switch1")
local xf = comp:FindTool("Transform1")
if not (sw and xf) then return end

local source = sw:GetInput("Source")
if source == 0 then
  xf:SetInput("Blend", 0.2)
elseif source == 1 then
  xf:SetInput("Blend", 1.0)
end
```
This checks a switch's `Source` input and maps it to a blend amount on another node.
Useful for mode selectors: different switch positions drive different parameter presets.

## Example: toggle a checkbox
```lua
local tool = comp:FindTool("Transform1")
if not tool then return end

local current = tool:GetInput("UseSize")
if current == 1 then
  tool:SetInput("UseSize", 0)
else
  tool:SetInput("UseSize", 1)
end
```
This flips a boolean-style input on/off, which is great for a "toggle" button in your macro.

## Notes
- Use tool/control IDs (not display names). Autocomplete and Pick help.
- Write normal Lua here; Macro Machine escapes strings on export.
- Avoid editing the same control in On-Change without a guard (it can retrigger).
