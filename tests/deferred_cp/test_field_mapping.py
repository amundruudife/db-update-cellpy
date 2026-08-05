import json
from pathlib import Path

import pytest

from src.deferred_cp.field_mapping import TARGET_COLUMNS, build_db_row, mapping_blockers, preserve_by_id


FIXTURES = Path(__file__).parent / "fixtures" / "gate_b"


def test_dictionary_has_exactly_73_columns_and_only_known_ownership_classes():
    assert len(TARGET_COLUMNS) == 73
    allowed = {"formula", "system", "manual", "preserve_existing_blank_new", "unresolved"}
    assert all(entry.ownership in allowed for entry in TARGET_COLUMNS)
    assert not mapping_blockers() == []


def test_build_row_uses_id_keyed_preservation_and_approved_rules():
    source = {
        "key": 101,
        "channel": "C1",
        "cell label": "cell-101",
        "file name": "raw-101",
        "active material working electrode (g)": 0.012,
        "test schedule": "C/10",
    }
    existing = {"id": 101, "b01": "manual", "tester": "T-1", "comment_one": "keep"}

    row = build_db_row(source, existing, allow_blocked=True)

    assert row[0] == "=Slurry!A4"
    assert row[3] == 1
    assert row[5] == "manual"
    assert row[18] == "T-1"
    assert row[48] == "=Slurry!B4"
    assert row[56] == "=Slurry!L4"


def test_unresolved_mapping_cannot_generate_a_row():
    with pytest.raises(ValueError, match="unresolved"):
        build_db_row({"key": 101}, {})


def test_exact_layout_fixture_and_golden_artifact_cover_the_dictionary():
    database = json.loads((FIXTURES / "database_layout.json").read_text())
    golden = json.loads((FIXTURES / "golden_rows.json").read_text())
    assert len(database["columns"]) == 73
    assert database["columns"] == [entry.header for entry in TARGET_COLUMNS]
    assert golden["target_column_count"] == len(TARGET_COLUMNS)


def test_formula_system_manual_and_preservation_rules_are_exact_in_golden_row():
    golden = json.loads((FIXTURES / "golden_rows.json").read_text())
    row = build_db_row(
        {"key": 101},
        {"b01": "keep-b01", "tester": "tester-101", "comment_one": "keep-comment"},
        allow_blocked=True,
    )
    assert row[0] == "=Slurry!A4"
    assert row[3] == golden["approved_system_values"]["D"]
    assert row[5] == "keep-b01"
    assert row[18] == "tester-101"
    assert row[48] == "=Slurry!B4"
    assert row[56] == "=Slurry!L4"
    assert row[14] is None  # conversion is unresolved, never guessed


def test_preservation_is_keyed_by_id_and_new_rows_are_blank():
    existing = {101: {"id": 101, "b01": "old"}, 202: {"id": 202, "b01": "new-old"}}
    assert preserve_by_id((202, 101), existing)[0]["id"] == 202
    new_row = build_db_row({"key": 303}, {}, allow_blocked=True)
    assert new_row[5:12] == [None] * 7
    assert new_row[1] is None


def test_rendering_is_idempotent_for_unchanged_logical_input():
    source = {"key": 101}
    existing = {"b01": "keep", "tester": "T"}
    assert build_db_row(source, existing, allow_blocked=True) == build_db_row(source, existing, allow_blocked=True)
