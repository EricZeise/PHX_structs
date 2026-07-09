# PHX_structs — strategy and routines

## Why this exists

PHPP encodes cross-worksheet references as composite `"<id>-<description>"`
strings built by its own dropdown selectors — `WINDOWS.window_rows[].frame_id`/
`glazing_id` point into `COMPONENTS.frames`/`glazings`, and
`AREAS.surface_rows[].assembly_id` points into the R-Values assembly library.
Neither PHX's upstream shape JSON nor PHX_pyxl/PHX_xlwg's field maps or
pydantic models resolve these — they're read and written as opaque scalar
values with no join/lookup step. This repo adds that join, read-only, as a
standalone cross-reference file.

## The missing piece: enumerating R-Values assemblies

Unlike `COMPONENTS.frames`/`glazings` (already read as repeating-row lists
via PHX_pyxl's `resolve_block`), R-Values assemblies repeat as **blocks**
down a single column — the header `"Description of building assembly"`
recurs once per assembly. No existing PHX_pyxl/PHX_xlwg reader strategy
enumerates repeated header occurrences; each one resolves at most one or two
anchor rows. `blocks.py` adds `find_all_rows_in_col()` (a generalization of
PHX_pyxl's `find_row_in_col`, which only returns the first match) plus
`read_repeating_blocks()`, which applies the field map's *existing*
`UVALUES.constructor` section spec — unmodified, no new markdown syntax — at
every discovered header row.

Only `id` and `display_name` are extracted per assembly block — the fields
needed to resolve `assembly_id` references. The section's other declared
offsets (`rsi_row_offset`, `first_layer_row_offset`, etc.) do not resolve
consistently against a real workbook for this identity pair when checked
directly (e.g. `name_row_offset: 2` implies a row where `display_name`
verifiably is not) — and no current PHX_pyxl reader dispatch rule ever
exercises this section's `column_fields` at all (it falls through to
`_read_items_section`, which only echoes literal config back). Resolving
the remaining construction-detail fields (R-values, per-layer conductivity)
is a distinct, unstarted task, not attempted here.

## Verified against a real workbook

Running `phpp-struct-ref build` against `PHX_xlwg/Data/Example_IP.xlsx`
resolves all 64 window rows' `frame_id`/`glazing_id` and all 25 area rows'
`assembly_id` with zero unresolved references, against 3 real (of 19
templated, mostly-unused) R-Values assembly slots — matching the identity
values (`01ud`/`Externl wall`, `02ud`/`Roof`, `03ud`/`Basement ceiling`)
confirmed by direct inspection of the workbook during this tool's design.
