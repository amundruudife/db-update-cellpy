"""Small candidate-only updater for the retained legacy ``log`` workflow.

This module deliberately has no production replacement capability. It reads a
local legacy Cell Log workbook, keeps the approved project filter, appends only
new IDs to ``Slurry``, and adds minimal ``db_table`` rows. Existing database
rows, including ``b01`` through ``b07``, are never rewritten.
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .contracts import (
    DATA_START_ROW as MIRROR_DATA_START_ROW,
    DATABASE_SHEET_NAME,
    MIRROR_SHEET_NAME,
    SYSTEM_VALUES,
    is_production_database,
)


SOURCE_SHEET_NAME = "log"
SOURCE_DATA_START_ROW = 5
SOURCE_PROJECT_COLUMN = 3
SOURCE_KEY_COLUMN = 1
APPROVED_PROJECTS = frozenset(
    {
        "SIS-Larger",
        "SIS-Large",
        "CellMap",
        "Norse-HV",
        "SUMBAT-SP5",
        "SUMBAT",
        "ASAP",
    }
)
DATABASE_DATA_START_ROW = 3
MANUAL_COLUMNS = tuple(range(6, 13))
SYSTEM_COLUMNS = {4: "exists", 20: "instrument", 24: "experiment_type"}


class CandidateBuildError(ValueError):
    """Raised when a candidate cannot be built without guessing or data loss."""


@dataclass(frozen=True)
class CandidateReport:
    source_path: str
    database_path: str
    candidate_path: str
    source_rows: int
    filtered_rows: int
    existing_slurry_rows: int
    retained_ids: tuple[int, ...]
    new_ids: tuple[int, ...]
    absent_existing_ids: tuple[int, ...]
    existing_duplicate_ids: tuple[int, ...]

    def to_dict(self) -> dict:
        return asdict(self)


def _resolved(path: str | Path) -> Path:
    return Path(path).resolve(strict=False)


def _positive_integer(value, *, context: str) -> int:
    if isinstance(value, bool) or value is None:
        raise CandidateBuildError(f"invalid {context}: {value!r}")
    if isinstance(value, int):
        normalized = value
    elif isinstance(value, float) and value.is_integer():
        normalized = int(value)
    else:
        raise CandidateBuildError(f"invalid {context}: {value!r}")
    if normalized <= 0:
        raise CandidateBuildError(f"invalid {context}: {value!r}")
    return normalized


def _nonempty_rows(rows: Iterable[tuple]) -> Iterable[tuple]:
    for row in rows:
        if any(value is not None for value in row):
            yield row


def _read_filtered_source(
    source_path: Path,
    *,
    max_columns: int | None = None,
) -> tuple[int, list[tuple[int, tuple]]]:
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    try:
        if SOURCE_SHEET_NAME not in workbook.sheetnames:
            raise CandidateBuildError("source workbook must contain the 'log' sheet")
        sheet = workbook[SOURCE_SHEET_NAME]
        source_column_count = max(sheet.max_column, SOURCE_PROJECT_COLUMN)
        column_count = (
            min(source_column_count, max_columns)
            if max_columns is not None
            else source_column_count
        )
        if column_count < SOURCE_PROJECT_COLUMN:
            raise CandidateBuildError("Slurry has no project column")

        source_rows = 0
        filtered = []
        seen = set()
        rows = sheet.iter_rows(
            min_row=SOURCE_DATA_START_ROW,
            max_col=column_count,
            values_only=True,
        )
        for row in _nonempty_rows(rows):
            source_rows += 1
            if row[SOURCE_PROJECT_COLUMN - 1] not in APPROVED_PROJECTS:
                continue
            record_id = _positive_integer(
                row[SOURCE_KEY_COLUMN - 1], context="source ID"
            )
            if record_id in seen:
                raise CandidateBuildError(f"duplicate source ID {record_id}")
            seen.add(record_id)
            filtered.append((record_id, tuple(row)))
        return source_rows, filtered
    finally:
        workbook.close()


def _existing_slurry_id(value, *, context: str) -> int:
    if isinstance(value, str) and value.isascii() and value.isdigit():
        value = int(value)
    return _positive_integer(value, context=context)


def _existing_slurry_ids(sheet) -> tuple[list[int], tuple[int, ...]]:
    ids = []
    for row_number in range(MIRROR_DATA_START_ROW, sheet.max_row + 1):
        value = sheet.cell(row_number, SOURCE_KEY_COLUMN).value
        if value is None:
            continue
        ids.append(
            _existing_slurry_id(value, context=f"existing Slurry ID at A{row_number}")
        )
    duplicate_ids = tuple(
        sorted(record_id for record_id, count in Counter(ids).items() if count > 1)
    )
    return ids, duplicate_ids


def _assert_safe_paths(
    source_path: Path,
    database_path: Path,
    candidate_path: Path,
    report_path: Path | None,
) -> None:
    if candidate_path in {source_path, database_path}:
        raise CandidateBuildError("candidate must use a separate output path")
    if report_path is not None and report_path in {source_path, database_path, candidate_path}:
        raise CandidateBuildError("report must use a separate output path")
    if any(is_production_database(path) for path in (source_path, database_path, candidate_path)):
        raise CandidateBuildError("candidate workflow cannot read from or write to production")
    if not source_path.is_file():
        raise FileNotFoundError(f"source workbook not found: {source_path}")
    if not database_path.is_file():
        raise FileNotFoundError(f"database workbook not found: {database_path}")


def _worksheet_values(sheet, *, max_row: int, max_column: int) -> tuple[tuple, ...]:
    return tuple(
        tuple(row)
        for row in sheet.iter_rows(
            min_row=1,
            max_row=max_row,
            min_col=1,
            max_col=max_column,
            values_only=True,
        )
    )


def _verify_candidate(
    candidate_path: Path,
    *,
    initial_slurry_rows: int,
    initial_slurry_columns: int,
    existing_slurry_values: tuple[tuple, ...],
    initial_database_rows: int,
    initial_database_columns: int,
    existing_database_values: tuple[tuple, ...],
    manual_values: tuple[tuple, ...],
    appended_slurry_rows: tuple[int, ...],
    appended_database_rows: tuple[int, ...],
    new_ids: tuple[int, ...],
) -> None:
    workbook = load_workbook(candidate_path, read_only=True, data_only=False)
    try:
        if MIRROR_SHEET_NAME not in workbook.sheetnames or DATABASE_SHEET_NAME not in workbook.sheetnames:
            raise CandidateBuildError("candidate is missing Slurry or db_table")
        slurry = workbook[MIRROR_SHEET_NAME]
        database = workbook[DATABASE_SHEET_NAME]

        actual_slurry_values = _worksheet_values(
            slurry,
            max_row=initial_slurry_rows,
            max_column=initial_slurry_columns,
        )
        if actual_slurry_values != existing_slurry_values:
            raise CandidateBuildError("candidate changed existing Slurry rows")

        actual_slurry_ids = tuple(
            _positive_integer(
                slurry.cell(row_number, SOURCE_KEY_COLUMN).value,
                context=f"candidate Slurry ID at A{row_number}",
            )
            for row_number in appended_slurry_rows
        )
        if actual_slurry_ids != new_ids:
            raise CandidateBuildError("candidate Slurry append verification failed")

        actual_database_values = _worksheet_values(
            database,
            max_row=initial_database_rows,
            max_column=initial_database_columns,
        )
        if actual_database_values != existing_database_values:
            raise CandidateBuildError("candidate changed existing db_table rows")

        manual_start = MANUAL_COLUMNS[0] - 1
        actual_manual = tuple(
            row[manual_start : manual_start + len(MANUAL_COLUMNS)]
            for row in actual_database_values[DATABASE_DATA_START_ROW - 1 :]
        )
        if actual_manual != manual_values:
            raise CandidateBuildError("candidate changed existing b01:b07 values")

        if len(appended_database_rows) != len(new_ids):
            raise CandidateBuildError("candidate db_table append verification failed")
        for row_number, record_id in zip(appended_database_rows, new_ids):
            if database.cell(row_number, 1).value != record_id:
                raise CandidateBuildError("candidate db_table ID verification failed")
            allowed_columns = {1, *SYSTEM_COLUMNS}
            if any(
                column not in allowed_columns and database.cell(row_number, column).value is not None
                for column in range(1, database.max_column + 1)
            ):
                raise CandidateBuildError("new db_table cells must be blank except approved system values")
            for column, system_name in SYSTEM_COLUMNS.items():
                if database.cell(row_number, column).value != SYSTEM_VALUES[system_name]:
                    raise CandidateBuildError("candidate system-value verification failed")
    finally:
        workbook.close()


def build_candidate(
    source_path: str | Path,
    database_path: str | Path,
    candidate_path: str | Path,
    *,
    report_path: str | Path | None = None,
) -> CandidateReport:
    """Build and verify a candidate without touching source or production."""

    source = _resolved(source_path)
    database = _resolved(database_path)
    candidate = _resolved(candidate_path)
    report = _resolved(report_path) if report_path is not None else None
    _assert_safe_paths(source, database, candidate, report)

    candidate.parent.mkdir(parents=True, exist_ok=True)
    if report is not None:
        report.parent.mkdir(parents=True, exist_ok=True)

    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            suffix=".xlsx",
            prefix=f".{candidate.stem}.",
            dir=candidate.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
        shutil.copy2(database, temporary)

        workbook = load_workbook(temporary, data_only=False)
        try:
            if MIRROR_SHEET_NAME not in workbook.sheetnames or DATABASE_SHEET_NAME not in workbook.sheetnames:
                raise CandidateBuildError("database must contain Slurry and db_table")
            slurry = workbook[MIRROR_SHEET_NAME]
            database_sheet = workbook[DATABASE_SHEET_NAME]

            source_rows, filtered = _read_filtered_source(
                source,
                max_columns=slurry.max_column,
            )
            existing_ids, existing_duplicate_ids = _existing_slurry_ids(slurry)
            existing_set = set(existing_ids)
            filtered_ids = tuple(record_id for record_id, _ in filtered)
            filtered_set = set(filtered_ids)
            retained_ids = tuple(record_id for record_id in filtered_ids if record_id in existing_set)
            new_rows = tuple((record_id, row) for record_id, row in filtered if record_id not in existing_set)
            new_ids = tuple(record_id for record_id, _ in new_rows)
            absent_existing_ids = tuple(sorted(existing_set - filtered_set))

            initial_database_rows = database_sheet.max_row
            initial_database_columns = database_sheet.max_column
            existing_database_values = _worksheet_values(
                database_sheet,
                max_row=initial_database_rows,
                max_column=initial_database_columns,
            )
            manual_values = tuple(
                tuple(database_sheet.cell(row_number, column).value for column in MANUAL_COLUMNS)
                for row_number in range(DATABASE_DATA_START_ROW, initial_database_rows + 1)
            )
            initial_slurry_rows = slurry.max_row
            initial_slurry_columns = slurry.max_column
            existing_slurry_values = _worksheet_values(
                slurry,
                max_row=initial_slurry_rows,
                max_column=initial_slurry_columns,
            )

            appended_slurry_rows = []
            appended_database_rows = []
            for record_id, row in new_rows:
                row_number = max(MIRROR_DATA_START_ROW - 1, slurry.max_row) + 1
                appended_slurry_rows.append(row_number)
                for column, value in enumerate(row, start=1):
                    slurry.cell(row_number, column, value)

                database_row = max(DATABASE_DATA_START_ROW - 1, database_sheet.max_row) + 1
                appended_database_rows.append(database_row)
                database_sheet.cell(database_row, 1, record_id)
                for column, system_name in SYSTEM_COLUMNS.items():
                    database_sheet.cell(database_row, column, SYSTEM_VALUES[system_name])

            if new_rows:
                workbook.save(temporary)
        finally:
            workbook.close()

        _verify_candidate(
            temporary,
            initial_slurry_rows=initial_slurry_rows,
            initial_slurry_columns=initial_slurry_columns,
            existing_slurry_values=existing_slurry_values,
            initial_database_rows=initial_database_rows,
            initial_database_columns=initial_database_columns,
            existing_database_values=existing_database_values,
            manual_values=manual_values,
            appended_slurry_rows=tuple(appended_slurry_rows),
            appended_database_rows=tuple(appended_database_rows),
            new_ids=new_ids,
        )

        result = CandidateReport(
            source_path=str(source),
            database_path=str(database),
            candidate_path=str(candidate),
            source_rows=source_rows,
            filtered_rows=len(filtered),
            existing_slurry_rows=len(existing_ids),
            retained_ids=retained_ids,
            new_ids=new_ids,
            absent_existing_ids=absent_existing_ids,
            existing_duplicate_ids=existing_duplicate_ids,
        )
        os.replace(temporary, candidate)
        temporary = None
        if report is not None:
            report.write_text(
                json.dumps(result.to_dict(), indent=2, sort_keys=True),
                encoding="utf-8",
            )
        return result
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
