"""Focused tests for the pure hard-coded source validation contract."""

import pytest

from src.deferred_cp.contracts import (
    BUSINESS_LAST_COLUMN,
    DATA_START_ROW,
    HEADER_ROW,
    KEY_COLUMN,
    LOCAL_SNAPSHOT_PATH,
    SOURCE_COLUMN_COUNT,
    SOURCE_SHEET_NAME,
    SOURCE_WORKBOOK_NAME,
    UNITS_ROW,
)
from src.contracts import DATABASE_SHEET_NAME, MIRROR_SHEET_NAME, PRODUCTION_DATABASE_PATH, SYSTEM_VALUES
from src.deferred_cp.source_validation import (
    validate_first_manifest,
    validate_manifest_continuity,
    validate_source_values,
)


def _row(value=None):
    values = [None] * SOURCE_COLUMN_COUNT
    values[KEY_COLUMN - 1] = value
    return values


def _table(*ids):
    return [
        _row("metadata"),
        _row("key"),
        _row("units"),
        *[_row(value) for value in ids],
    ]


def test_contract_contains_the_fixed_source_and_workbook_geometry():
    assert SOURCE_WORKBOOK_NAME == "Cell_Log.xlsx"
    assert SOURCE_SHEET_NAME == "c&p"
    assert str(LOCAL_SNAPSHOT_PATH).replace("\\", "/") == "source_data/Cell_Log_CP.xlsx"
    assert str(PRODUCTION_DATABASE_PATH).endswith(
        "cellpy_data\\db\\2025_Cell_Analysis_db_001.xlsx"
    )
    assert MIRROR_SHEET_NAME == "Slurry"
    assert DATABASE_SHEET_NAME == "db_table"
    assert SOURCE_COLUMN_COUNT == 19
    assert BUSINESS_LAST_COLUMN == 16
    assert (HEADER_ROW, UNITS_ROW, DATA_START_ROW, KEY_COLUMN) == (2, 3, 4, 1)
    assert SYSTEM_VALUES == {
        "exists": 1,
        "instrument": "arbin_sql_h5",
        "experiment_type": "cycling",
    }


def test_valid_evaluated_values_return_ids_in_source_order():
    result = validate_source_values(_table(11, 12.0, 13))

    assert result.is_valid
    assert result.ids == (11, 12, 13)
    assert result.diagnostics == ()


@pytest.mark.parametrize(
    ("value", "code"),
    [
        (True, "boolean_id"),
        (None, "blank_id"),
        ("", "blank_id"),
        (0, "non_positive_id"),
        (-2, "non_positive_id"),
        (1.25, "non_integral_float_id"),
        ("12", "string_id"),
        ("#REF!", "excel_error_id"),
    ],
)
def test_invalid_ids_are_rejected_with_stable_diagnostics(value, code):
    result = validate_source_values(_table(value))

    assert not result.is_valid
    assert [diagnostic.code for diagnostic in result.diagnostics] == [code]


def test_id_anomalies_are_aggregated_and_duplicates_are_reported_once():
    table = _table(True, "#DIV/0!", 7, 7, None, "8")

    result = validate_source_values(table)

    assert not result.is_valid
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        "boolean_id",
        "excel_error_id",
        "blank_id",
        "string_id",
        "duplicate_ids",
    ]
    assert result.diagnostics[-1].values == ("7",)
    assert result.diagnostics[-1].rows == (6, 7)


def test_shape_and_header_contract_is_validated_before_data_ids():
    table = [_row("metadata"), ["not-key"] + [None] * 18, _row("units")]

    result = validate_source_values(table)

    assert not result.is_valid
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        "header_mismatch",
    ]

    result = validate_source_values([_row("metadata"), _row("key")])

    assert not result.is_valid
    assert result.diagnostics[0].code == "insufficient_metadata_rows"

    result = validate_source_values([_row("metadata"), _row("key"), _row("units")[:-1]])

    assert not result.is_valid
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        "non_rectangular",
        "wrong_column_count",
    ]


def test_manifest_continuity_allows_only_appended_ids_and_preserves_order():
    accepted = validate_manifest_continuity((11, 12), (11, 12, 13, 14))

    assert accepted.is_valid
    assert accepted.ids == (11, 12, 13, 14)

    missing = validate_manifest_continuity((11, 12, 13), (11, 12, 14))
    reordered = validate_manifest_continuity((11, 12, 13), (12, 11, 13))
    inserted = validate_manifest_continuity((11, 12, 13), (11, 12, 14, 13))

    assert [d.code for d in missing.diagnostics] == ["missing_historical_ids"]
    assert missing.diagnostics[0].values == ("13",)
    assert [d.code for d in reordered.diagnostics] == ["reordered_historical_ids"]
    assert [d.code for d in inserted.diagnostics] == ["new_ids_not_appended"]


def test_first_manifest_rejects_any_specified_existing_database_id_absent_from_source():
    result = validate_first_manifest((11, 12), existing_database_ids=(11, 99, 100))

    assert not result.is_valid
    assert [d.code for d in result.diagnostics] == ["missing_existing_database_ids"]
    assert result.diagnostics[0].values == ("99", "100")
