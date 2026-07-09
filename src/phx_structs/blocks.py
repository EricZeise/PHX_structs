"""Enumerates repeating header blocks on a PHPP worksheet.

PHPP's R-Values sheet repeats a "Description of building assembly" header
down column L once per assembly (confirmed empirically against a real
workbook: rows 7, 28, 49, 70... a fixed 21-row stride, though this reader
does not assume a fixed stride -- it just keeps searching). None of
PHX_pyxl's existing locators.py strategies enumerate *repeated* occurrences
of a header string -- every one of them (resolve_block, _read_row_offset_section,
etc.) calls find_row_in_col at most once or twice to get a single anchor row.
This module adds that missing enumeration as new code, reusing PHX_pyxl's
find_row_in_col as its only primitive -- no changes to PHX_pyxl itself.
"""

from __future__ import annotations

from typing import Any

from phx_structs.sibling_import import ensure_phpp_tool_importable

ensure_phpp_tool_importable()

from phpp_tool.locators import WsPair, cell_value, col_to_idx, find_row_in_col  # noqa: E402


def find_all_rows_in_col(
    ws_vals: Any, col: str, needle: str, *, contains: bool = True,
) -> list[int]:
    """Return every row where *col*'s cell text matches *needle*.

    Generalizes locators.find_row_in_col (which returns only the first
    match) by repeating the same call with an advancing start_from -- the
    same two-step pattern resolve_block already uses for its header-then-entry
    lookup, just continued until no more matches are found.
    """
    rows: list[int] = []
    start = 1
    while True:
        row = find_row_in_col(ws_vals, col, needle, contains=contains, start_from=start)
        if row is None:
            break
        rows.append(row)
        start = row + 1
    return rows


def read_repeating_blocks(
    ws_pair: WsPair, sec_spec: dict[str, Any], *, skip_formulas: bool = False,
) -> list[dict[str, Any]]:
    """Enumerate every assembly block in a UVALUES-shaped repeating section.

    Reads the *existing, unmodified* field-map section spec that
    map_parser.py already produces for UVALUES.constructor -- no new
    markdown syntax was needed. Only two fields are extracted here:
    `id` and `display_name`, the pair needed to resolve
    AREAS.surface_rows[].assembly_id references (the cross-referencing use
    case this repo exists for).

    Deliberately narrow scope: the section's other column_fields (r_si,
    r_se, per-layer conductivity/thickness, etc.) use row-offset config
    values (rsi_row_offset, first_layer_row_offset, ...) whose anchor
    convention does not hold up under direct verification against a real
    workbook for this identity pair -- e.g. `name_row_offset: 2` implies
    row 9 for display_name, but display_name is confirmed to actually sit
    at row 8 (header_row + 1) in the real sample workbook, and this
    section's column_fields are never read by any current PHX_pyxl reader
    dispatch rule at all (this shape falls through to `_read_items_section`,
    which only echoes the literal offset config back, never touches
    column_fields) -- so there is no working reference implementation to
    validate those other offsets against. Resolving them is a separate,
    unstarted task, not attempted here.
    """
    ws_vals, _ = ws_pair
    header_locator = sec_spec["header_locator"]
    items = sec_spec.get("items", {})
    column_fields = sec_spec.get("column_fields", {})

    name_col = column_fields.get("display_name", {}).get("column", header_locator["col"])
    id_col_offset = items.get("phpp_id_num_col_offset")
    id_col = None
    if id_col_offset is not None:
        id_col_idx = col_to_idx(header_locator["col"]) + id_col_offset
        id_col = _idx_to_col(id_col_idx)

    # Empirically confirmed: both display_name and the assembly id sit one
    # row below the header row, not at the row math `name_row_offset`
    # implies (see docstring above).
    IDENTITY_ROW_OFFSET = 1

    header_rows = find_all_rows_in_col(
        ws_vals, header_locator["col"], header_locator["string"],
    )

    blocks: list[dict[str, Any]] = []
    for header_row in header_rows:
        data_row = header_row + IDENTITY_ROW_OFFSET
        display_name = cell_value(
            ws_pair, name_col, data_row, skip_formulas=skip_formulas,
        )
        if not display_name:
            continue
        record: dict[str, Any] = {"_row": data_row, "display_name": display_name}
        if id_col is not None:
            record["id"] = cell_value(ws_pair, id_col, data_row, skip_formulas=skip_formulas)
        blocks.append(record)
    return blocks


def _idx_to_col(idx: int) -> str:
    """Convert a 1-based column index back to letters (inverse of col_to_idx)."""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters
