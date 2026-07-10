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

from phpp_tool.locators import WsPair, cell_value, col_to_idx, find_row_in_col, resolve_row_offset  # noqa: E402


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


def read_assembly_construction_detail(
    ws_pair: WsPair, sec_spec: dict[str, Any], *, skip_formulas: bool = False,
) -> list[dict[str, Any]]:
    """Enumerate every assembly block with full construction detail (R-values,
    per-layer conductivity/thickness), not just id/display_name.

    Every declared `*_row_offset` in this section is off by exactly one,
    confirmed against two independent real assembly blocks ("Externl wall",
    "Roof" in Example_IP.xlsx, and reconfirmed in Example_SI.xlsx's SI field
    map) -- a single systematic fencepost error, not per-field drift:

        name_row_offset (2) -> actual 1 (display_name/id row)
        rsi_row_offset (4) -> actual 3 ("...or Rsi" row)
        rse_row_offset (5) -> actual 4 ("...or Rsa" row)
        first_layer_row_offset (7) -> actual 6 (first layer-table row)
        last_layer_row_offset (14) -> actual 13 (last *possible* layer row)
        result_val_row_offset (19) -> actual 18 (the "R-value [...]:" row)

    So `actual_offset = declared_offset - 1`, applied uniformly below. The
    "Percentage of sec. 1/2/3:" row sits right after the layer window, at
    the *uncorrected* last_layer_row_offset value (a coincidence of the
    fencepost error landing exactly on that boundary row).

    r_si/r_se cells are dual-purpose in real PHPP data -- they hold either a
    dropdown preset string (e.g. "2-Wall") or a numeric override, depending
    on what the user picked. A non-numeric value here is expected, not a
    resolution bug.

    interior_insulation/u_val_supplement's exact row is a positional
    hypothesis (same row as r_si/r_se respectively, by analogy with their
    column-Q labels sitting on those rows) -- neither sample assembly checked
    had these set, so this hasn't been confirmed against a populated example.

    variants_layer_name/variants_conductivity/variants_thickness (columns
    E/F/G) are deliberately excluded -- they sit under a "Selection of types
    of building layer" header that reads as a static dropdown-source
    reference table, not per-assembly data, and were confirmed empty in
    every block checked.
    """
    identity_blocks = read_repeating_blocks(ws_pair, sec_spec, skip_formulas=skip_formulas)

    items = sec_spec.get("items", {})
    column_fields = sec_spec.get("column_fields", {})

    def col(field_name: str, default: str) -> str:
        return column_fields.get(field_name, {}).get("column", default)

    rsi_offset = items.get("rsi_row_offset", 4) - 1
    rse_offset = items.get("rse_row_offset", 5) - 1
    first_layer_offset = items.get("first_layer_row_offset", 7) - 1
    last_layer_offset = items.get("last_layer_row_offset", 14) - 1
    post_layer_offset = last_layer_offset + 1
    result_val_offset = items.get("result_val_row_offset", 19) - 1
    result_val_col = items.get("result_val_col", "R")

    r_si_col = col("r_si", "M")
    r_se_col = col("r_se", "M")
    interior_insulation_col = col("interior_insulation", "R")
    u_val_supplement_col = col("u_val_supplement", "R")
    sec1_desc_col = col("sec_1_description", "L")
    sec1_cond_col = col("sec_1_conductivity", "M")
    sec2_desc_col = col("sec_2_description", "N")
    sec2_cond_col = col("sec_2_conductivity", "O")
    sec3_desc_col = col("sec_3_description", "P")
    sec3_cond_col = col("sec_3_conductivity", "Q")
    thickness_col = col("thickness", "R")
    sec2_pct_col = col("sec_2_percentage", "O")
    sec3_pct_col = col("sec_3_percentage", "Q")

    blocks: list[dict[str, Any]] = []
    for identity in identity_blocks:
        # identity["_row"] is header_row + 1 (see read_repeating_blocks) --
        # recover the header row to anchor the rest of this block's offsets.
        header_row = identity["_row"] - 1

        record: dict[str, Any] = dict(identity)
        record["r_si"] = resolve_row_offset(ws_pair, header_row, r_si_col, rsi_offset, skip_formulas=skip_formulas)
        record["r_se"] = resolve_row_offset(ws_pair, header_row, r_se_col, rse_offset, skip_formulas=skip_formulas)
        record["interior_insulation"] = resolve_row_offset(
            ws_pair, header_row, interior_insulation_col, rsi_offset, skip_formulas=skip_formulas)
        record["u_val_supplement"] = resolve_row_offset(
            ws_pair, header_row, u_val_supplement_col, rse_offset, skip_formulas=skip_formulas)
        record["result_val"] = resolve_row_offset(
            ws_pair, header_row, result_val_col, result_val_offset, skip_formulas=skip_formulas)

        layers: list[dict[str, Any]] = []
        for offset in range(first_layer_offset, last_layer_offset + 1):
            row = header_row + offset
            sec1_desc = cell_value(ws_pair, sec1_desc_col, row, skip_formulas=skip_formulas)
            if not sec1_desc:
                break  # confirmed contiguous in real samples -- first blank ends the table
            layer: dict[str, Any] = {
                "_row": row,
                "sec_1_description": sec1_desc,
                "sec_1_conductivity": cell_value(ws_pair, sec1_cond_col, row, skip_formulas=skip_formulas),
                "thickness": cell_value(ws_pair, thickness_col, row, skip_formulas=skip_formulas),
            }
            sec2_desc = cell_value(ws_pair, sec2_desc_col, row, skip_formulas=skip_formulas)
            if sec2_desc:
                layer["sec_2_description"] = sec2_desc
                layer["sec_2_conductivity"] = cell_value(ws_pair, sec2_cond_col, row, skip_formulas=skip_formulas)
            sec3_desc = cell_value(ws_pair, sec3_desc_col, row, skip_formulas=skip_formulas)
            if sec3_desc:
                layer["sec_3_description"] = sec3_desc
                layer["sec_3_conductivity"] = cell_value(ws_pair, sec3_cond_col, row, skip_formulas=skip_formulas)
            layers.append(layer)
        record["layers"] = layers

        post_row = header_row + post_layer_offset
        record["sec_2_percentage"] = cell_value(ws_pair, sec2_pct_col, post_row, skip_formulas=skip_formulas)
        record["sec_3_percentage"] = cell_value(ws_pair, sec3_pct_col, post_row, skip_formulas=skip_formulas)

        blocks.append(record)
    return blocks


def _idx_to_col(idx: int) -> str:
    """Convert a 1-based column index back to letters (inverse of col_to_idx)."""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters
