# CLAUDE.md — PHPP Component/Window/Area Cross-Reference Tool

## What this project is

A read-only Python CLI that resolves PHPP's internal cross-worksheet reference conventions and writes the joined result to a single cross-reference JSON file. Two distinct conventions are resolved:

- **Id-string convention** (`"<id>-<description>"`, e.g. `"01ud-Climatop Ultra N"`): `WINDOWS.window_rows[].frame_id`/`glazing_id` → `COMPONENTS.frames`/`glazings`; `AREAS.surface_rows[].assembly_id` → the R-Values assembly library; `ADDNL_VENT.units[].unit_selected` → `COMPONENTS.ventilators`.
- **Ordinal-position convention** (a 1-based integer naming a row in a sibling list, not an id string): `ADDNL_VENT.rooms[].vent_unit_assigned` → `ADDNL_VENT.units[]`; `ADDNL_VENT.ducts[].duct_assign_1..duct_assign_10` (ten flag columns, a duct can serve multiple rooms) → `ADDNL_VENT.rooms[]`.

`AREAS.thermal_bridge_rows` is also surfaced, though it carries no id-style link to any other worksheet at all — its `group_number` field is parsed into `(group_code, group_label)` for grouping, not resolved as a cross-reference (confirmed: thermal-bridge group codes and surface group codes are disjoint families).

### MVP pipeline

```
Filled PHPP (.xlsx)  →  read (via PHX_pyxl) + enumerate R-Values assemblies (new)  →  join by id-string or ordinal position  →  crossref.json
```

### Relationship to PHX_pyxl / PHX_xlwg / phpp-shape-sync

PHX_structs does not read or write PHPP cell values on its own — it imports PHX_pyxl's `map_parser.py`/`locators.py`/`reader.py` directly from the sibling repo (same pattern `phpp-shape-sync` uses for `map_parser.py`), because those modules are pure Python with no Excel dependency. `COMPONENTS`, `WINDOWS`, and `AREAS` are read via PHX_pyxl's existing `read_phpp()` unmodified.

What's missing from PHX_pyxl/PHX_xlwg and added here: neither repo's `reader.py` can enumerate a *repeating* header block — every existing addressing strategy (see `locators.py`'s six strategies) resolves at most one or two anchor rows via `find_row_in_col`, which itself returns only the first match. PHPP's R-Values sheet, however, repeats `"Description of building assembly"` down a single column once per assembly (confirmed: rows 7, 28, 49, 70... in a real sample workbook, each block carrying its own `display_name` and PHPP-generated `id`). This repo adds that missing enumeration as new code — `find_all_rows_in_col()` and `read_repeating_blocks()` — reusing PHX_pyxl's existing `find_row_in_col`/`resolve_row_offset` primitives rather than duplicating them. **Nothing in PHX_pyxl or PHX_xlwg is modified** — the field map's existing `UVALUES.constructor` section spec is read as-is; no new markdown grammar was needed.

This is intentionally out of scope for phpp-shape-sync (which only maps upstream shape JSON to local field-map markdown, and never opens a workbook to read data values) and out of scope for PHX_pyxl/PHX_xlwg (whose job is reading/writing a single building record, not cross-worksheet identity resolution).

### Requirement: no Excel installation

Like PHX_pyxl, this uses openpyxl exclusively — no `xlwings`/Excel dependency.

---

## Architecture

```
PHX_pyxl/src/phpp_tool/{map_parser,locators,reader}.py   (imported directly, sibling path — never modified)
        ↓
   sibling_import.py     (adds PHX_pyxl/src to sys.path on demand)
        ↓
   blocks.py             (NEW: find_all_rows_in_col + read_repeating_blocks — R-Values assembly enumeration)
        ↓
 ┌──────┴──────┐
 read_phpp()   read_repeating_blocks()   (Components/Windows/Areas)   (Assemblies)
 └──────┬──────┘
     crossref.py          (build id-keyed indexes, resolve "<id>-<description>" references, track unresolved)
        ↓
     cli.py                (Click CLI: `build`)
```

---

## Repository structure

```
PHX_structs/
├── CLAUDE.md                    ← this file
├── pyproject.toml               ← deps: openpyxl, click (no pydantic/lxml/xlwings — read-only, no write-back)
├── src/
│   └── phx_structs/
│       ├── __init__.py
│       ├── sibling_import.py    ← PHX_PYXL_SRC path + ensure_phpp_tool_importable()
│       ├── blocks.py            ← repeating-block enumerator (the new reader capability)
│       ├── crossref.py          ← index-building + ID-resolution/join logic
│       └── cli.py               ← Click CLI: build
├── tests/
│   ├── test_blocks.py
│   └── test_crossref.py
└── docs/
    └── phx-structs-strategy-and-routines.md
```

---

## Commands

```bash
# Install in dev mode
pip install -e ".[dev]"

# Run tests (no Excel needed)
pytest tests/ -v

# Build a cross-reference file from a real, populated PHPP workbook
phpp-struct-ref build Data/Example_IP.xlsx --phpp-version EN_10_6_IP -o crossref.json
```

---

## Constraints

- **Read-only** — this tool never writes to a PHPP workbook. If write-back (e.g. reassigning a window's `frame_id`) is ever needed, that's a distinct follow-on scope, not this repo's job.
- **Never modify PHX_pyxl or PHX_xlwg** — all reuse goes through direct sibling-path imports (`sibling_import.py`), exactly like `phpp-shape-sync`'s `normalizer.py` does for `map_parser.py`. If a capability added here (e.g. `find_all_rows_in_col`) later proves broadly useful to the field-map grammar itself, that's a deliberate separate decision, not an automatic upstream.
- **The field map is still the single source of truth for cell locations** — `blocks.py` reads its header locator string, `name_row_offset`, and `phpp_id_num_col_offset` from the *existing* `UVALUES.constructor` section spec parsed by `map_parser.py`; it does not hardcode sheet coordinates.
- **Never silently drop an unresolved reference** — any `frame_id`/`glazing_id`/`assembly_id`/`unit_selected`/`vent_unit_assigned`/`duct_assign_N` that doesn't resolve against its target index/list is recorded in the output's `unresolved` section for human review, not dropped.
- **Don't invent cross-references the data doesn't have** — `AREAS.thermal_bridge_rows` has no id-style link to any other worksheet; resist the temptation to force a join against `surface_rows.group_number` just because both fields share the same `"<code>-<label>"` text shape. Confirmed the two use disjoint code families in a real workbook — parse and group by category, don't resolve as a reference.
- **Two distinct resolution primitives, don't conflate them** — `resolve_reference()` (id-string split) and `resolve_ordinal()` (1-based sibling-list position) solve different problems; `ADDNL_VENT` needs both in the same worksheet (`units[].unit_selected` is id-string, `rooms[].vent_unit_assigned`/`ducts[].duct_assign_N` are ordinal). HP/Boiler worksheets are genuinely empty stubs in the field map — nothing to associate there yet.
