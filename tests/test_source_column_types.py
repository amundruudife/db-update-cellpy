import pytest

from src.source_validation import (
    COLUMN_ERROR_POLICIES,
    SOURCE_COLUMN_CONTRACTS,
    SourceValidationError,
    validate_source_rows,
    validate_values_only,
)


def test_every_mirrored_column_has_an_explicit_value_type_contract():
    assert set(SOURCE_COLUMN_CONTRACTS) == set("ABCDEFGHIJKLMNOPQRS")

    for rule in SOURCE_COLUMN_CONTRACTS.values():
        assert "blank" in rule.allowed_types
        assert "excel_error" in rule.allowed_types

    assert SOURCE_COLUMN_CONTRACTS["A"].allowed_types == {
        "integer",
        "blank",
        "excel_error",
    }
    for letter in "BCDJL":
        assert SOURCE_COLUMN_CONTRACTS[letter].allowed_types == {
            "text",
            "blank",
            "excel_error",
        }
    for letter in "EFGHIKMNOPQRS":
        assert SOURCE_COLUMN_CONTRACTS[letter].allowed_types == {
            "number",
            "blank",
            "excel_error",
        }


def test_error_policy_rejects_mapping_errors_and_preserves_helper_errors():
    assert set(COLUMN_ERROR_POLICIES) == set("ABCDEFGHIJKLMNOPQRS")
    assert all(COLUMN_ERROR_POLICIES[letter] == "reject" for letter in "ABCDEFGHIJKLMNOP")
    assert all(COLUMN_ERROR_POLICIES[letter] == "preserve" for letter in "QRS")

    valid_row = [
        1,
        "channel",
        "cell",
        "file.xlsx",
        1.0,
        2.0,
        3.0,
        4.0,
        5.0,
        "constant",
        6.0,
        "schedule",
        7.0,
        8.0,
        9.0,
        10.0,
        11.0,
        12.0,
        "#N/A",
    ]
    assert validate_source_rows([valid_row]) == [valid_row]

    invalid_row = valid_row.copy()
    invalid_row[0] = "#VALUE!"
    with pytest.raises(SourceValidationError, match="A4"):
        validate_source_rows([invalid_row])


def test_values_only_validation_rejects_formula_text_after_type_validation():
    row = valid_row = [
        1,
        "channel",
        "cell",
        "file.xlsx",
        1.0,
        2.0,
        3.0,
        4.0,
        5.0,
        "constant",
        6.0,
        "schedule",
        7.0,
        8.0,
        9.0,
        10.0,
        11.0,
        12.0,
        "#N/A",
    ]
    row[1] = "=A4"

    with pytest.raises(SourceValidationError, match="formula"):
        validate_values_only(
            [
                ["metadata"] + [None] * 18,
                ["key"] + [None] * 18,
                ["units"] + [None] * 18,
                row,
            ]
        )


def test_values_only_validation_returns_evaluated_values_without_coercion():
    row = [
        1,
        "channel",
        "cell",
        "file.xlsx",
        1.0,
        2.0,
        3.0,
        4.0,
        5.0,
        "constant",
        6.0,
        "schedule",
        7.0,
        8.0,
        9.0,
        10.0,
        11.0,
        12.0,
        "#N/A",
    ]
    values = validate_values_only(
        [
            ["metadata"] + [None] * 18,
            ["key"] + [None] * 18,
            ["units"] + [None] * 18,
            row,
        ]
    )

    assert values[3] == row
    assert isinstance(values[3][0], int)
    assert isinstance(values[3][4], float)
    assert isinstance(values[3][1], str)
    assert values[3][18] == "#N/A"
