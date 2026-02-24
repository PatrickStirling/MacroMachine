# Macro Machine Optimization Options

## Goal
Reduce lag on large macros and heavy CSV-linked workflows without destabilizing export/import behavior.

## Priority Recommendations

1. Do less work per edit
- Debounce expensive CSV apply/reload operations.
- Avoid full parse + full rebuild on every small UI change.
- Only trigger expensive work when structural text actually changed.

2. Incremental parsing and patching
- Maintain cached boundaries for group/tool/input blocks.
- Patch only touched regions instead of rewriting entire documents.
- Reuse previously computed indexes until invalidated by structural edits.

3. Move heavy work off the UI thread
- Run large parse/transform tasks in a worker/background process.
- Keep main thread focused on interaction and rendering.
- Show progress/working state for long operations.

4. Virtualize large UI lists
- Render only visible rows in Nodes and Published panels.
- Use windowed rendering for long control lists.
- Avoid re-rendering full lists when one row changes.

5. Large payload handling (base64/blob placeholder strategy)
- Detect very large literal payloads (base64/image/html-like blocks).
- Replace in-editor display with lightweight placeholders or previews.
- Keep an out-of-band blob map and re-inject exact originals on export.
- Preserve full round-trip fidelity while protecting editor responsiveness.

## Lower Priority / Higher Complexity

### Dictionary compression for repeated keywords
- Idea: map repeated tokens (`InstanceInput`, `SourceOp`, etc.) to compact internal representation.
- Assessment: likely lower ROI vs. complexity for current bottlenecks.
- Most lag is expected from repeated full parse/rewrite/render cycles, not token storage size.

## Suggested Rollout Order

1. Debounce + reduce unnecessary rebuilds
2. Incremental patch/index cache
3. Worker-thread heavy parse path
4. List virtualization
5. Base64/blob placeholder + reinjection system

## Success Metrics

- Time to import large macro (ms)
- Time to apply CSV updates with mixed row overrides (ms)
- Main-thread stalls > 100 ms (count)
- Export correctness parity (no round-trip regressions)
- Memory use on large macros
