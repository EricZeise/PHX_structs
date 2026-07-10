from phx_structs.crossref import (
    _build_hvac,
    _build_thermal_bridges,
    _index_by_id,
    _split_ref,
    list_assemblies,
    resolve_ordinal,
    resolve_reference,
)


def test_split_ref_splits_on_first_dash():
    assert _split_ref("01ud-Climatop Ultra N") == ("01ud", "Climatop Ultra N")


def test_split_ref_handles_dash_in_description():
    # Only the first "-" is a delimiter; the rest belongs to the description.
    assert _split_ref("01ud-Climatop-Ultra-N") == ("01ud", "Climatop-Ultra-N")


def test_split_ref_none_for_non_string_or_no_dash():
    assert _split_ref(None) is None
    assert _split_ref(42) is None
    assert _split_ref("01ud") is None  # no "-" at all


def test_index_by_id_keys_on_id_field():
    records = [{"id": "01ud", "description": "Window Fixed"}, {"id": "02ud"}, {}]
    index = _index_by_id(records)
    assert set(index) == {"01ud", "02ud"}
    assert index["01ud"]["description"] == "Window Fixed"


def test_resolve_reference_matches_existing_id():
    index = {"01ud": {"description": "Climatop Ultra N"}}
    row = {"glazing_id": "01ud-Climatop Ultra N"}
    assert resolve_reference(row, "glazing_id", index) == "01ud"


def test_resolve_reference_returns_none_for_unresolved_id():
    index = {"01ud": {"description": "Climatop Ultra N"}}
    row = {"glazing_id": "99zz-Nonexistent Glazing"}
    assert resolve_reference(row, "glazing_id", index) is None


def test_resolve_reference_returns_none_for_missing_field():
    assert resolve_reference({}, "glazing_id", {"01ud": {}}) is None


# --- resolve_ordinal (ADDNL_VENT's rooms<->units<->ducts position-based convention) ---

def test_resolve_ordinal_matches_first_position():
    records = [{"display_name": "HRV-1"}, {"display_name": "HRV-2"}]
    assert resolve_ordinal(1, records) == {"display_name": "HRV-1"}


def test_resolve_ordinal_matches_last_position():
    records = [{"display_name": "HRV-1"}, {"display_name": "HRV-2"}]
    assert resolve_ordinal(2, records) == {"display_name": "HRV-2"}


def test_resolve_ordinal_accepts_integer_valued_float():
    # openpyxl often returns whole numbers as float
    records = [{"display_name": "HRV-1"}]
    assert resolve_ordinal(1.0, records) == {"display_name": "HRV-1"}


def test_resolve_ordinal_out_of_range_returns_none():
    records = [{"display_name": "HRV-1"}]
    assert resolve_ordinal(2, records) is None


def test_resolve_ordinal_rejects_non_positive_or_non_integer():
    records = [{"display_name": "HRV-1"}]
    assert resolve_ordinal(0, records) is None
    assert resolve_ordinal(-1, records) is None
    assert resolve_ordinal(1.5, records) is None
    assert resolve_ordinal(None, records) is None
    assert resolve_ordinal("1", records) is None
    assert resolve_ordinal(True, records) is None  # bool is an int subclass -- must not slip through


# --- _build_hvac: the full rooms<->units<->ducts<->ventilators chain ---

def _sample_hvac_data():
    return {
        "COMPONENTS": {"ventilators": [{"id": "01ud", "display_name": "Comfoair 550"}]},
        "ADDNL_VENT": {
            "units": [
                {"_row": 70, "display_name": "HRV-1", "unit_selected": "01ud-Comfoair 550"},
                {"_row": 71, "display_name": "HRV-2", "unit_selected": "99zz-Unknown Unit"},
            ],
            "rooms": [
                {"_row": 31, "display_name": "childrens 1", "vent_unit_assigned": 1},
                {"_row": 32, "display_name": "childrens 2", "vent_unit_assigned": 2},
                {"_row": 33, "display_name": "orphan room", "vent_unit_assigned": 9},
            ],
            "ducts": [
                {"_row": 95, "duct_assign_1": 1, "duct_assign_2": None},
                {"_row": 96, "duct_assign_1": 1, "duct_assign_2": 1},
                {"_row": 97, "duct_assign_5": 1},  # room index 5 doesn't exist -> unresolved
            ],
        },
    }


def test_build_hvac_resolves_unit_to_ventilator():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert hvac["units"][0]["resolved_ventilator_id"] == "01ud"


def test_build_hvac_unresolved_ventilator_selection():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert "resolved_ventilator_id" not in hvac["units"][1]
    assert unresolved["hvac_units"] == [
        {"_row": 71, "field": "unit_selected", "value": "99zz-Unknown Unit"}
    ]


def test_build_hvac_resolves_room_to_unit_by_ordinal():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert hvac["rooms"][0]["resolved_unit_row"] == 70
    assert hvac["rooms"][0]["resolved_unit_display_name"] == "HRV-1"
    assert hvac["rooms"][1]["resolved_unit_display_name"] == "HRV-2"


def test_build_hvac_unresolved_room_out_of_range():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert "resolved_unit_row" not in hvac["rooms"][2]
    assert unresolved["hvac_rooms"] == [
        {"_row": 33, "field": "vent_unit_assigned", "value": 9}
    ]


def test_build_hvac_duct_can_resolve_multiple_rooms():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert hvac["ducts"][1]["resolved_rooms"] == [
        {"_row": 31, "display_name": "childrens 1"},
        {"_row": 32, "display_name": "childrens 2"},
    ]


def test_build_hvac_duct_assignment_out_of_range_is_unresolved():
    data = _sample_hvac_data()
    ventilators_by_id = _index_by_id(data["COMPONENTS"]["ventilators"])
    hvac, unresolved = _build_hvac(data, ventilators_by_id)
    assert "resolved_rooms" not in hvac["ducts"][2]
    assert unresolved["hvac_ducts"] == [
        {"_row": 97, "field": "duct_assign_5", "value": 1}
    ]


# --- _build_thermal_bridges: category parsing, not a cross-reference ---

def test_build_thermal_bridges_parses_group_number():
    data = {
        "AREAS": {
            "thermal_bridge_rows": [
                {"_row": 146, "description": "Estimate Load Bearing Walls",
                 "group_number": "17-Thermal bridges FS/BC", "psi_value": 0.17},
            ]
        }
    }
    result = _build_thermal_bridges(data)
    assert result[0]["group_code"] == "17"
    assert result[0]["group_label"] == "Thermal bridges FS/BC"
    assert result[0]["psi_value"] == 0.17  # original fields preserved


def test_build_thermal_bridges_handles_missing_group_number():
    data = {"AREAS": {"thermal_bridge_rows": [{"_row": 146, "description": "No group"}]}}
    result = _build_thermal_bridges(data)
    assert "group_code" not in result[0]


def test_build_thermal_bridges_empty_when_no_rows():
    assert _build_thermal_bridges({"AREAS": {}}) == []
    assert _build_thermal_bridges({}) == []


# --- list_assemblies: group construction detail + referencing areas ---

def _sample_crossref():
    return {
        "assemblies": {
            "01ud": {"id": "01ud", "display_name": "Externl wall", "result_val": 41.64, "layers": []},
            "02ud": {"id": "02ud", "display_name": "Roof", "result_val": 110.49, "layers": []},
            "03ud": {"id": "03ud", "display_name": "Unused Slot", "result_val": 12.0, "layers": []},
        },
        "areas": [
            {"_row": 41, "description": "Floor_A", "area": 100.0, "resolved_assembly_id": "01ud"},
            {"_row": 42, "description": "Floor_B", "area": 50.0, "resolved_assembly_id": "01ud"},
            {"_row": 43, "description": "Roof_A", "area": 200.0, "resolved_assembly_id": "02ud"},
            {"_row": 44, "description": "Unresolved_area", "area": 10.0},  # no resolved_assembly_id at all
        ],
    }


def test_list_assemblies_returns_one_record_per_assembly():
    grouped = list_assemblies(_sample_crossref())
    assert [a["id"] for a in grouped] == ["01ud", "02ud", "03ud"]


def test_list_assemblies_groups_referencing_areas():
    grouped = list_assemblies(_sample_crossref())
    wall = next(a for a in grouped if a["id"] == "01ud")
    assert [area["description"] for area in wall["areas"]] == ["Floor_A", "Floor_B"]
    assert wall["area_total"] == 150.0


def test_list_assemblies_includes_unused_assembly_with_empty_areas():
    grouped = list_assemblies(_sample_crossref())
    unused = next(a for a in grouped if a["id"] == "03ud")
    assert unused["areas"] == []
    assert unused["area_total"] == 0


def test_list_assemblies_preserves_original_assembly_fields():
    grouped = list_assemblies(_sample_crossref())
    wall = next(a for a in grouped if a["id"] == "01ud")
    assert wall["display_name"] == "Externl wall"
    assert wall["result_val"] == 41.64


def test_list_assemblies_empty_when_no_assemblies():
    assert list_assemblies({"assemblies": {}, "areas": []}) == []
    assert list_assemblies({}) == []
