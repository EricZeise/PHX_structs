import random
import warnings

import openpyxl
import pytest

from phx_structs.blocks import find_all_rows_in_col, read_repeating_blocks

warnings.filterwarnings("ignore", category=UserWarning)

HEADER = "Description of building assembly"


@pytest.fixture
def assembly_sheet(tmp_path):
    """A synthetic R-Values-shaped sheet: two populated blocks, one blank slot."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "R-Values"

    # Block 1: header at row 7, identity data at row 8 (populated).
    ws["L7"] = "Description of building assembly"
    ws["Q7"] = "Assembly no."
    ws["L8"] = "Externl wall"
    ws["Q8"] = "01ud"

    # Block 2: header at row 28, identity data at row 29 (populated).
    ws["L28"] = "Description of building assembly"
    ws["Q28"] = "Assembly no."
    ws["L29"] = "Roof"
    ws["Q29"] = "02ud"

    # Block 3: header at row 49, but no display_name at row 50 (unused slot).
    ws["L49"] = "Description of building assembly"
    ws["Q49"] = "Assembly no."

    path = tmp_path / "assembly_sheet.xlsx"
    wb.save(path)
    wb_vals = openpyxl.load_workbook(path, data_only=True)
    wb_fmls = openpyxl.load_workbook(path, data_only=False)
    yield (wb_vals["R-Values"], wb_fmls["R-Values"])
    wb_vals.close()
    wb_fmls.close()


def test_find_all_rows_in_col_finds_every_occurrence(assembly_sheet):
    ws_vals, _ = assembly_sheet
    rows = find_all_rows_in_col(ws_vals, "L", "Description of building assembly")
    assert rows == [7, 28, 49]


def test_find_all_rows_in_col_no_match_returns_empty(assembly_sheet):
    ws_vals, _ = assembly_sheet
    assert find_all_rows_in_col(ws_vals, "L", "Nonexistent header") == []


def test_read_repeating_blocks_skips_unused_slot(assembly_sheet):
    sec_spec = {
        "header_locator": {"col": "L", "string": "Description of building assembly"},
        "items": {"phpp_id_num_col_offset": 5},
        "column_fields": {"display_name": {"column": "L", "unit": None}},
    }
    blocks = read_repeating_blocks(assembly_sheet, sec_spec)
    assert blocks == [
        {"_row": 8, "display_name": "Externl wall", "id": "01ud"},
        {"_row": 29, "display_name": "Roof", "id": "02ud"},
    ]


def test_handles_many_blocks_at_irregular_strides(tmp_path):
    """Real PHPP workbooks vary widely in how many assemblies are populated
    and how they're spaced -- the real sample used a fixed 21-row stride,
    but nothing in find_all_rows_in_col/read_repeating_blocks assumes a fixed
    stride, so this checks a much larger, deliberately irregular layout
    (15-40 rows between blocks, ~1/3 left as unused/blank template slots).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "R-Values"

    rng = random.Random(0)
    row = 7
    expected: list[dict] = []
    n_blocks = 40
    for i in range(n_blocks):
        ws.cell(row=row, column=12, value=HEADER)
        ws.cell(row=row, column=17, value="Assembly no.")
        if i % 3 != 0:  # every 3rd slot left blank/unused, like a real template
            name, ident = f"Assembly_{i:03d}", f"{i:02d}xx"
            ws.cell(row=row + 1, column=12, value=name)
            ws.cell(row=row + 1, column=17, value=ident)
            expected.append({"_row": row + 1, "display_name": name, "id": ident})
        row += rng.randint(15, 40)

    path = tmp_path / "many_blocks.xlsx"
    wb.save(path)
    wb_vals = openpyxl.load_workbook(path, data_only=True)
    wb_fmls = openpyxl.load_workbook(path, data_only=False)
    try:
        ws_pair = (wb_vals["R-Values"], wb_fmls["R-Values"])
        header_rows = find_all_rows_in_col(wb_vals["R-Values"], "L", HEADER)
        assert len(header_rows) == n_blocks

        sec_spec = {
            "header_locator": {"col": "L", "string": HEADER},
            "items": {"phpp_id_num_col_offset": 5},
            "column_fields": {"display_name": {"column": "L", "unit": None}},
        }
        assert read_repeating_blocks(ws_pair, sec_spec) == expected
    finally:
        wb_vals.close()
        wb_fmls.close()


def test_no_header_occurrences_returns_empty_list(tmp_path):
    """A workbook/sheet with no matching header at all (e.g. wrong sheet,
    or a PHPP version with different wording) must not error -- just
    produce zero blocks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "R-Values"
    ws["L7"] = "Some unrelated label"

    path = tmp_path / "no_header.xlsx"
    wb.save(path)
    wb_vals = openpyxl.load_workbook(path, data_only=True)
    wb_fmls = openpyxl.load_workbook(path, data_only=False)
    try:
        ws_pair = (wb_vals["R-Values"], wb_fmls["R-Values"])
        sec_spec = {
            "header_locator": {"col": "L", "string": HEADER},
            "items": {"phpp_id_num_col_offset": 5},
            "column_fields": {"display_name": {"column": "L", "unit": None}},
        }
        assert read_repeating_blocks(ws_pair, sec_spec) == []
    finally:
        wb_vals.close()
        wb_fmls.close()
