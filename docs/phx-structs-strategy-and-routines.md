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

> Can you scope the process to resolve the remaining construction-detail fields (R-values, per-layer conductivity)?

# Resolve UVALUES.constructor's remaining construction-detail fields

## Context

`PHX_structs/src/phx_structs/blocks.py`'s `read_repeating_blocks()` deliberately only extracts `id`/`display_name` per R-Values assembly block, documented at the time as: the field map's other declared row-offset config (`rsi_row_offset`, `first_layer_row_offset`, etc.) "does not hold up under direct verification... this needs fresh archaeology, not a quick patch" (see `project_phx_structs.md` memory). This plan is that archaeology, done against the real `R-Values` sheet in `PHX_xlwg/Data/Example_IP.xlsx` (two independent assembly blocks checked: "Externl wall" and "Roof"), plus the design for extending `blocks.py` to emit the full per-assembly construction detail (R-values, per-layer conductivity/thickness) once resolved.

**Key finding that simplifies this a lot:** every declared `*_row_offset` is off by exactly the same amount. Checked six independent fields against real data:

| Field                    | Declared offset | Actual offset (confirmed)                                    |
| ------------------------ | --------------- | ------------------------------------------------------------ |
| `name_row_offset`        | 2               | 1 (`display_name`/`id` row)                                  |
| `rsi_row_offset`         | 4               | 3 ("Orientation of building assembly (or Rsi)" row)          |
| `rse_row_offset`         | 5               | 4 ("Adjacent to (or Rsa)" row)                               |
| `first_layer_row_offset` | 7               | 6 (first layer-table row)                                    |
| `last_layer_row_offset`  | 14              | 13 (last *possible* layer-table row, i.e. layer slots span `first-1` through `last-1` inclusive) |
| `result_val_row_offset`  | 19              | 18 ("R-value [...]:" row, column `R`, matches `result_val_col`) |

So `actual_offset = declared_offset - 1`, uniformly. This isn't six unrelated bugs — it's one systematic off-by-one, almost certainly from a fencepost error when these offsets were originally authored. The fix in `blocks.py` is one correction applied consistently, not six separate special cases.

## What each `column_fields` entry actually is

Confirmed by reading every non-empty cell in both "Externl wall" (5 layers) and "Roof" (6 layers) blocks, columns A–S, offsets 0–18:

**Scalar fields** (one value per assembly, each at `header_row + (declared_offset - 1)`):
- `display_name` (col L), `id` (col = `header_locator.col` + `phpp_id_num_col_offset`, same row) — already implemented, unchanged.
- `r_si` (col M, at the corrected `rsi_row_offset` row) and `r_se` (col M, at the corrected `rse_row_offset` row) — confirmed these cells are **dual-purpose**: they hold either a dropdown preset string (e.g. `"2-Wall"`, `"1-Outdoor air"`) or a numeric override, depending on what the PHPP user picked. A non-numeric value here is expected PHPP behavior, not a resolution bug — don't coerce or reject it.
- `result_val` (col `R`, at the corrected `result_val_row_offset` row) — the assembly's overall U/R-value.
- `interior_insulation` and `u_val_supplement` (both declared col `R`): **positionally hypothesized, not yet confirmed by a populated example** — both sample assemblies leave these blank. By analogy (their column-Q labels — "Interior insulation?", "U-value supplement..." — sit at the same corrected `rsi`/`rse` offsets respectively), they're most likely col `R` at those same two rows. Flagged as a verification gap to close during implementation (try a third/fourth assembly across the sample workbooks that actually sets one of these before trusting the position).

**Per-layer fields** (repeat once per populated row across the fixed 8-slot layer-table window, offsets `(first_layer_row_offset-1)` through `(last_layer_row_offset-1)` inclusive — 5 layers used in "Externl wall", 6 in "Roof", confirming the table is variable-length within a fixed-size window, not a fixed count):
- `sec_1_description` (L) / `sec_1_conductivity` (M) — present on every populated layer row.
- `sec_2_description` (N) / `sec_2_conductivity` (O) and `sec_3_description` (P) / `sec_3_conductivity` (Q) — present only on layer rows with a mixed/inhomogeneous material (e.g. stud framing alongside cavity insulation); blank otherwise. This matches PHX's own `PhxLayerDivisionGrid` concept from the earlier HBJSON research.
- `thickness` (col R) — present on every populated layer row.

**Post-layer-table scalar fields** (one fixed row right after the layer window, offset 14 — same row as the "Percentage of sec. 1/2/3:" labels):
- `sec_2_percentage` (col O), `sec_3_percentage` (col Q). (A `sec_1` percentage value is visibly present in col M at this row too, e.g. `0.906`, but isn't a declared field in the field map — not adding it since it's not asked for and can be derived as `1 - sec_2 - sec_3` if ever needed.)

**Likely not per-assembly data at all** — `variants_layer_name`/`variants_conductivity`/`variants_thickness` (cols E/F/G) sit under a `"Selection of types of building layer"` header that reads like a static dropdown-source reference table, not assembly-specific values; confirmed empty in both sample blocks. Recommend **excluding these from the resolved output** rather than guessing a row, and documenting why, unless a workbook is found where they're populated.

## Implementation

- Add a new function in `blocks.py`, `read_assembly_construction_detail()` (keep `read_repeating_blocks()` as-is for callers that only need id/display_name — this is strictly additive, not a breaking change), that for each header row found via the existing `find_all_rows_in_col()`:
  1. Applies the uniform `-1` correction to read the scalar fields (`r_si`, `r_se`, `result_val`, and — once confirmed — `interior_insulation`/`u_val_supplement`) via the existing `resolve_row_offset()` primitive (already used for `id`/`display_name`, no new primitive needed here).
  2. Walks the 8-slot layer window and, for each row with a non-empty `sec_1_description`, builds a layer dict (`sec_1_description`/`conductivity`, `sec_2_*`/`sec_3_*` when present, `thickness`), stopping the *emitted* list at the first genuinely blank slot (rows can be sparse mid-window in theory, but confirmed contiguous in both real samples — treat a single blank as end-of-layers, matching how `COMPONENTS`' own repeating-row reader already treats sparse rows as a stop signal).
  3. Reads `sec_2_percentage`/`sec_3_percentage` at the fixed post-layer offset.
- Wire this into `crossref.py`: `_read_assemblies()` currently calls `read_repeating_blocks()` — switch it to call the new function instead (or merge results), so `assemblies_by_id` records gain the new fields. **No new join/resolution logic is needed** — this is enrichment of data already being resolved by id, not a new cross-reference relationship, so `crossref.py`'s `resolve_reference`/`resolve_ordinal` machinery is untouched.
- Verify the same technique against `EN_10_6_SI.md` + `Example_SI.xlsx` before assuming it generalizes — the SI field map is a separately hand-maintained file (per earlier `phpp-shape-sync` diff-audit memory, ~95.6% leaf-identical to IP but not guaranteed identical in this specific section) and its real workbook's `U-values` sheet was not checked at this offset-by-offset level of detail this session.

## Verification

- Unit tests in `PHX_structs/tests/test_blocks.py`: extend the existing synthetic-sheet fixtures with a populated layer table (mixed sec_1/sec_2 rows, a couple of blank trailing slots) and assert the new function returns the right per-layer list and scalar fields, including a case with `sec_2`/`sec_3` present and absent.
- Real-workbook check against `Example_IP.xlsx`: assembly `01ud` ("Externl wall") should show 5 layers (`Gypsum`, `Cellulose`×3, `Sheathing`), `result_val` ≈ `41.64`; assembly `02ud` ("Roof") should show 6 layers (`Gypsum`, `Cellulose`×3, `Polyiso`, `OSB`), `result_val` ≈ `110.49` — both values already confirmed by direct inspection this session, so this is a strict regression check, not new exploration.
- Repeat against `Example_SI.xlsx` with `EN_10_6_SI.md` once that field map's offsets are separately confirmed (see note above) — do not assume the IP correction transfers without checking.
- Run the full `pytest tests/ -v` suite and the CLI end-to-end (`phpp-struct-ref build ...`) against all five sample workbooks already used for prior verification, confirming `unresolved` counts don't regress.

## Listing assemblies with their referencing areas grouped together

`build_crossref()` keeps `assemblies` and `areas` as two separate flat structures, cross-referenced only by `resolved_assembly_id` — reading "what does assembly `01ud` look like, and what surfaces use it?" out of that output means scanning `areas` by hand for every assembly you care about. `list_assemblies()` in `crossref.py` does that grouping once: for each assembly, it bundles the construction detail already resolved by `read_assembly_construction_detail()` (identity, R/U-value, full layer stack) together with every `AREAS.surface_rows` entry that references it, plus a summed `area_total`. Assemblies with no referencing surfaces are still included (empty `areas` list, `area_total: 0`) — an unused library slot is real information, not noise to hide.

### CLI usage

```bash
# Human-readable summary table (identity, result_val, layer materials, area count/total)
phpp-struct-ref assemblies Data/Example_IP.xlsx --phpp-version EN_10_6_IP
```

```
3 assemblies identified:

  01ud: 'Externl wall'
      result_val=41.64030609362415
      layers=[Gypsum + Cellulose + Cellulose + Cellulose + Sheathing]
      used by 20 area(s), total area=4357.574072157234

  02ud: 'Roof'
      result_val=110.49475855233152
      layers=[Gypsum + Cellulose + Cellulose + Cellulose + Polyiso + OSB]
      used by 4 area(s), total area=9855.613233790264

  03ud: 'Basement ceiling'
      result_val=58.84530490283202
      layers=[XPS + Concrete]
      used by 1 area(s), total area=9056.11117544875
```

```bash
# Full grouped JSON (one record per assembly, "areas" list nested inside each)
phpp-struct-ref assemblies Data/Example_IP.xlsx --phpp-version EN_10_6_IP -o assemblies.json
```

Verified against `Example_IP.xlsx` (area totals above, matching the per-area sums confirmed during construction-detail verification) and `Example_SI.xlsx`. The SI example surfaced a real, if unrelated, data fact worth knowing: every `area_total` there comes back `0` — not a bug, `AREAS.surface_rows[].area` is genuinely `null` for all 6 surfaces in that particular demo workbook (a small hand-built example that never had its area formulas cached), confirmed by reading the raw `read_phpp()` output directly before trusting the summary.

### Reusing `list_assemblies()` outside the CLI

It's a pure function over `build_crossref()`'s output, so it composes directly:

```python
from phx_structs.crossref import build_crossref, list_assemblies

field_map = "/Users/smini/Documents/Coding/PHX_pyxl/phpp-field-mapping/EN_10_6_IP.md"
result = build_crossref("Data/Example_IP.xlsx", field_map)
for assembly in list_assemblies(result):
    print(assembly["id"], assembly["display_name"], assembly["area_total"])
```
