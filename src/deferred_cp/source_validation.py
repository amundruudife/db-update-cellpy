"""Pure validation for the hard-coded evaluated-values source mirror."""

import math
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

from .contracts import (
    DATA_START_ROW,
    HEADER_ROW,
    KEY_COLUMN,
    KEY_HEADER,
    SOURCE_COLUMN_COUNT,
    UNITS_ROW,
)


EXCEL_ERROR_VALUES = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#UNKNOWN!",
        "#BUSY!",
        "#CONNECT!",
        "#PYTHON!",
    }
)


@dataclass(frozen=True)
class ColumnContract:
    letter: str
    allowed_types: frozenset[str]


class SourceValidationError(ValueError):
    """Raised when a source row violates the mirror contract."""


@dataclass(frozen=True)
class ValidationDiagnostic:
    """One deterministic validation finding."""

    code: str
    message: str
    rows: Tuple[int, ...] = ()
    values: Tuple[str, ...] = ()
    row: Optional[int] = None
    column: Optional[int] = None


@dataclass(frozen=True)
class SourceValidationResult:
    """Typed result for one pure source-table validation."""

    row_count: int
    column_count: int
    ids: Tuple[int, ...]
    diagnostics: Tuple[ValidationDiagnostic, ...] = ()

    @property
    def is_valid(self) -> bool:
        return not self.diagnostics

    @property
    def valid(self) -> bool:
        return self.is_valid

    @property
    def errors(self) -> Tuple[ValidationDiagnostic, ...]:
        return self.diagnostics


@dataclass(frozen=True)
class ManifestValidationResult:
    """Typed result for append-only manifest checks."""

    ids: Tuple[int, ...]
    diagnostics: Tuple[ValidationDiagnostic, ...] = ()

    @property
    def is_valid(self) -> bool:
        return not self.diagnostics

    @property
    def valid(self) -> bool:
        return self.is_valid

    @property
    def errors(self) -> Tuple[ValidationDiagnostic, ...]:
        return self.diagnostics


def _is_excel_error(value) -> bool:
    return isinstance(value, str) and value.upper() in EXCEL_ERROR_VALUES


def _display(value) -> str:
    """Return a stable, non-repr display form for diagnostic values."""

    return str(value)


def _materialize_table(values) -> Tuple[List[object], ...]:
    try:
        rows = tuple(values)
    except TypeError:
        return ()

    materialized = []
    for row in rows:
        if isinstance(row, (str, bytes)):
            materialized.append([row])
            continue
        try:
            materialized.append(list(row))
        except TypeError:
            materialized.append([row])
    return tuple(materialized)


def _id_anomaly(value) -> Tuple[Optional[int], Optional[str]]:
    """Return ``(normalized_id, diagnostic_code)`` for one key value."""

    if isinstance(value, bool):
        return None, "boolean_id"
    if value is None:
        return None, "blank_id"
    if isinstance(value, str):
        if not value.strip():
            return None, "blank_id"
        if _is_excel_error(value):
            return None, "excel_error_id"
        return None, "string_id"
    if isinstance(value, int):
        normalized = value
    elif isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None, "non_integral_float_id"
        normalized = int(value)
    else:
        return None, "unsupported_id_type"

    if normalized <= 0:
        return None, "non_positive_id"
    return int(normalized), None


def validate_source_values(values) -> SourceValidationResult:
    """Validate an in-memory evaluated-values table.

    The table is never coerced, written, or read from an external source. IDs
    are normalized only after they pass the positive-integer rule; all other
    source values are left untouched for the later mirror validator.
    """

    rows = _materialize_table(values)
    row_count = len(rows)
    first_width = len(rows[0]) if rows else 0
    diagnostics = []

    widths = tuple(len(row) for row in rows)
    if widths and any(width != first_width for width in widths):
        irregular_rows = tuple(
            row_number for row_number, width in enumerate(widths, start=1) if width != first_width
        )
        diagnostics.append(
            ValidationDiagnostic(
                code="non_rectangular",
                message="Source values must be rectangular.",
                rows=irregular_rows,
            )
        )
    if not widths or any(width != SOURCE_COLUMN_COUNT for width in widths):
        diagnostics.append(
            ValidationDiagnostic(
                code="wrong_column_count",
                message="Source values must contain exactly 19 columns.",
                values=tuple(str(width) for width in sorted(set(widths))) or ("0",),
            )
        )
    if row_count < UNITS_ROW:
        diagnostics.append(
            ValidationDiagnostic(
                code="insufficient_metadata_rows",
                message="Source values must contain metadata rows 1 through 3.",
            )
        )

    # A malformed shape cannot safely expose a key column.  Return structural
    # findings before inspecting IDs, keeping diagnostics deterministic.
    if diagnostics:
        return SourceValidationResult(
            row_count=row_count,
            column_count=first_width,
            ids=(),
            diagnostics=tuple(diagnostics),
        )

    if rows[HEADER_ROW - 1][KEY_COLUMN - 1] != KEY_HEADER:
        return SourceValidationResult(
            row_count=row_count,
            column_count=first_width,
            ids=(),
            diagnostics=(
                ValidationDiagnostic(
                    code="header_mismatch",
                    message="The source key header must be 'key' at row 2/A.",
                    row=HEADER_ROW,
                    column=KEY_COLUMN,
                    values=(_display(rows[HEADER_ROW - 1][KEY_COLUMN - 1]),),
                ),
            ),
        )

    ids = []
    id_rows: Dict[int, List[int]] = {}
    anomaly_rows: Dict[str, List[int]] = {}
    anomaly_values: Dict[str, List[str]] = {}
    for offset, row in enumerate(rows[DATA_START_ROW - 1 :], start=DATA_START_ROW):
        normalized, code = _id_anomaly(row[KEY_COLUMN - 1])
        if code is not None:
            anomaly_rows.setdefault(code, []).append(offset)
            anomaly_values.setdefault(code, []).append(_display(row[KEY_COLUMN - 1]))
            continue
        ids.append(normalized)
        id_rows.setdefault(normalized, []).append(offset)

    findings = []
    for code, rows_for_code in anomaly_rows.items():
        findings.append(
            ValidationDiagnostic(
                code=code,
                message="Invalid source IDs were found in the key column.",
                rows=tuple(rows_for_code),
                values=tuple(anomaly_values[code]),
                column=KEY_COLUMN,
            )
        )

    duplicate_ids = [value for value, rows_for_value in id_rows.items() if len(rows_for_value) > 1]
    if duplicate_ids:
        duplicate_rows = tuple(
            row_number
            for value in duplicate_ids
            for row_number in id_rows[value]
        )
        findings.append(
            ValidationDiagnostic(
                code="duplicate_ids",
                message="Source IDs must be unique.",
                rows=duplicate_rows,
                values=tuple(_display(value) for value in duplicate_ids),
                column=KEY_COLUMN,
            )
        )

    return SourceValidationResult(
        row_count=row_count,
        column_count=first_width,
        ids=tuple(ids),
        diagnostics=tuple(findings),
    )


def _duplicate_manifest_diagnostic(ids: Tuple[int, ...]) -> Optional[ValidationDiagnostic]:
    positions: Dict[int, List[int]] = {}
    for position, value in enumerate(ids, start=1):
        positions.setdefault(value, []).append(position)
    duplicate_values = [value for value, positions_for_value in positions.items() if len(positions_for_value) > 1]
    if not duplicate_values:
        return None
    duplicate_positions = tuple(
        position
        for value in duplicate_values
        for position in positions[value]
    )
    return ValidationDiagnostic(
        code="duplicate_manifest_ids",
        message="Manifest IDs must be unique.",
        rows=duplicate_positions,
        values=tuple(_display(value) for value in duplicate_values),
    )


def validate_manifest_continuity(
    historical_ids: Iterable[int], current_ids: Iterable[int]
) -> ManifestValidationResult:
    """Allow a manifest to retain its historical prefix and append new IDs."""

    historical = tuple(historical_ids)
    current = tuple(current_ids)
    diagnostics = []

    duplicate_diagnostic = _duplicate_manifest_diagnostic(current)
    if duplicate_diagnostic is not None:
        diagnostics.append(duplicate_diagnostic)

    missing = tuple(value for value in historical if value not in current)
    if missing:
        diagnostics.append(
            ValidationDiagnostic(
                code="missing_historical_ids",
                message="Historical source IDs cannot disappear.",
                values=tuple(_display(value) for value in missing),
            )
        )

    historical_set = set(historical)
    present_historical = tuple(value for value in current if value in historical_set)
    expected_present = tuple(value for value in historical if value in current)
    if present_historical != expected_present:
        diagnostics.append(
            ValidationDiagnostic(
                code="reordered_historical_ids",
                message="Historical source IDs must retain their source order.",
                values=tuple(_display(value) for value in present_historical),
            )
        )

    new_ids = tuple(value for value in current if value not in historical_set)
    new_ids_are_not_appended = bool(new_ids) and current[-len(new_ids) :] != new_ids
    if new_ids_are_not_appended:
        diagnostics.append(
            ValidationDiagnostic(
                code="new_ids_not_appended",
                message="New source IDs may only be appended after historical IDs.",
                values=tuple(_display(value) for value in new_ids),
            )
        )

    return ManifestValidationResult(ids=current, diagnostics=tuple(diagnostics))


def validate_first_manifest(
    source_ids: Iterable[int], existing_database_ids: Iterable[int] = ()
) -> ManifestValidationResult:
    """Reject any specified existing database ID absent from first source data."""

    source = tuple(source_ids)
    existing = tuple(existing_database_ids)
    diagnostics = []
    duplicate_diagnostic = _duplicate_manifest_diagnostic(source)
    if duplicate_diagnostic is not None:
        diagnostics.append(duplicate_diagnostic)

    missing = tuple(value for value in existing if value not in source)
    if missing:
        diagnostics.append(
            ValidationDiagnostic(
                code="missing_existing_database_ids",
                message="Existing database IDs must be present in the first source manifest.",
                values=tuple(_display(value) for value in missing),
            )
        )
    return ManifestValidationResult(ids=source, diagnostics=tuple(diagnostics))


# Clear aliases keep the later snapshot layer independent from these names.
validate_source_table = validate_source_values
validate_append_only_manifest = validate_manifest_continuity


def _contract(letter: str, value_type: str) -> ColumnContract:
    return ColumnContract(
        letter=letter,
        allowed_types=frozenset({value_type, "blank", "excel_error"}),
    )


SOURCE_COLUMN_CONTRACTS = {
    letter: _contract(
        letter,
        "integer"
        if letter == "A"
        else "text"
        if letter in "BCDJL"
        else "number",
    )
    for letter in "ABCDEFGHIJKLMNOPQRS"
}

COLUMN_ERROR_POLICIES = {
    **{letter: "reject" for letter in "ABCDEFGHIJKLMNOP"},
    **{letter: "preserve" for letter in "QRS"},
}


def _cell_value_type(value, letter: str) -> str:
    if value is None:
        return "blank"
    if isinstance(value, str) and value.upper() in EXCEL_ERROR_VALUES:
        return "excel_error"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "integer" if letter == "A" and float(value).is_integer() else "number"
    if isinstance(value, str):
        return "text"
    return "unknown"


def validate_source_rows(rows, start_row: int = 4):
    """Validate and return values-only A:S rows without coercion."""

    validated = []
    for offset, row in enumerate(rows):
        row_values = list(row)
        excel_row = start_row + offset
        if len(row_values) != len(SOURCE_COLUMN_CONTRACTS):
            raise SourceValidationError(
                f"Row {excel_row} has {len(row_values)} values; expected 19"
            )
        for index, (letter, contract) in enumerate(SOURCE_COLUMN_CONTRACTS.items()):
            kind = _cell_value_type(row_values[index], letter)
            if kind not in contract.allowed_types:
                raise SourceValidationError(
                    f"{letter}{excel_row}: value type {kind!r} is not permitted"
                )
            if kind == "excel_error" and COLUMN_ERROR_POLICIES[letter] == "reject":
                raise SourceValidationError(f"{letter}{excel_row}: Excel error is not allowed")
        validated.append(row_values)
    return validated


def validate_values_only(values):
    """Validate a complete A:S table and reject formula text everywhere."""

    rows = [list(row) for row in values]
    source_result = validate_source_values(rows)
    if not source_result.is_valid:
        raise SourceValidationError(
            "; ".join(diagnostic.code for diagnostic in source_result.diagnostics)
        )
    validate_source_rows(rows[DATA_START_ROW - 1 :], start_row=DATA_START_ROW)
    for row_number, row in enumerate(rows, start=1):
        for column_number, value in enumerate(row, start=1):
            if isinstance(value, str) and value.startswith("="):
                raise SourceValidationError(
                    f"{row_number}:{column_number}: formula text is not permitted"
                )
    return rows
