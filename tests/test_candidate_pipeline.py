import json
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import src.candidate_pipeline as candidate_pipeline
from src.contracts import PRODUCTION_DATABASE_PATH
from src.candidate_pipeline import CandidateBuildError, build_candidate


PROJECT = "CellMap"


def _source(path: Path, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "log"
    for column in range(1, 88):
        sheet.cell(1, column, f"group-{column}")
        sheet.cell(2, column, f"header-{column}")
        sheet.cell(3, column, f"unit-{column}")
        sheet.cell(4, column, f"field-{column}")
    sheet["A2"] = "key"
    sheet["C2"] = "proj"
    sheet["A4"] = "key"
    sheet["C4"] = "proj"
    for row_number, (record_id, project, marker) in enumerate(rows, start=5):
        sheet.cell(row_number, 1, record_id)
        sheet.cell(row_number, 3, project)
        sheet.cell(row_number, 4, marker)
    workbook.save(path)
    workbook.close()


def _database(path: Path, ids=(101,)):
    workbook = Workbook()
    slurry = workbook.active
    slurry.title = "Slurry"
    for column in range(1, 88):
        slurry.cell(1, column, f"header-{column}")
        slurry.cell(2, column, f"unit-{column}")
        slurry.cell(3, column, f"field-{column}")

    database = workbook.create_sheet("db_table")
    for column in range(1, 74):
        database.cell(1, column, f"db-header-{column}")
        database.cell(2, column, f"db-type-{column}")
    database["A1"] = "id"
    database["D1"] = "exists"
    for offset, record_id in enumerate(ids, start=4):
        slurry.cell(offset, 1, record_id)
        slurry.cell(offset, 4, f"existing-{record_id}")
        db_row = offset - 1
        database.cell(db_row, 1, record_id)
        for column in range(6, 13):
            database.cell(db_row, column, f"manual-{record_id}-{column}")
        database.cell(db_row, 30, f"historical-{record_id}")
    workbook.save(path)
    workbook.close()


def test_candidate_appends_filtered_rows_and_preserves_existing_database_rows(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    report_path = tmp_path / "candidate.report.json"
    _source(
        source,
        [
            (101, PROJECT, "already-present"),
            (202, "Other", "filtered-out"),
            (303, PROJECT, "new-row"),
        ],
    )
    _database(database)
    source_before = source.read_bytes()
    database_before = database.read_bytes()

    report = build_candidate(source, database, candidate, report_path=report_path)

    assert source.read_bytes() == source_before
    assert database.read_bytes() == database_before
    assert report.new_ids == (303,)
    assert report.absent_existing_ids == ()
    assert json.loads(report_path.read_text(encoding="utf-8"))["new_ids"] == [303]

    workbook = load_workbook(candidate, data_only=False)
    try:
        slurry = workbook["Slurry"]
        assert [slurry.cell(row, 1).value for row in range(4, 6)] == [101, 303]
        assert slurry["D5"].value == "new-row"

        database_sheet = workbook["db_table"]
        assert database_sheet["A3"].value == 101
        assert [database_sheet.cell(3, column).value for column in range(6, 13)] == [
            f"manual-101-{column}" for column in range(6, 13)
        ]
        assert database_sheet["AD3"].value == "historical-101"
        assert database_sheet["A4"].value == 303
        assert database_sheet["D4"].value == 1
        assert database_sheet["T4"].value == "arbin_sql_h5"
        assert database_sheet["X4"].value == "cycling"
        assert [database_sheet.cell(4, column).value for column in range(6, 13)] == [None] * 7
        assert not any(
            isinstance(database_sheet.cell(4, column).value, str)
            and database_sheet.cell(4, column).value.startswith("=")
            for column in range(1, database_sheet.max_column + 1)
        )
    finally:
        workbook.close()


def test_candidate_reports_absent_existing_ids_without_removing_them(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(101, PROJECT, "present")])
    _database(database, ids=(101, 404))

    report = build_candidate(source, database, candidate)

    assert report.absent_existing_ids == (404,)
    workbook = load_workbook(candidate, read_only=True, data_only=False)
    try:
        assert [workbook["Slurry"].cell(row, 1).value for row in range(4, 6)] == [101, 404]
        assert workbook["db_table"]["A4"].value == 404
    finally:
        workbook.close()


def test_noop_candidate_is_byte_identical_and_writes_report(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    report_path = tmp_path / "candidate.report.json"
    _source(source, [(101, PROJECT, "present")])
    _database(database, ids=(101, 404))
    database_before = database.read_bytes()
    save_calls = []
    original_load_workbook = candidate_pipeline.load_workbook

    def load_workbook_with_save_spy(*args, **kwargs):
        workbook = original_load_workbook(*args, **kwargs)
        if not kwargs.get("read_only", False):
            original_save = workbook.save

            def save(*save_args, **save_kwargs):
                save_calls.append(save_args)
                return original_save(*save_args, **save_kwargs)

            workbook.save = save
        return workbook

    monkeypatch.setattr(candidate_pipeline, "load_workbook", load_workbook_with_save_spy)

    report = build_candidate(source, database, candidate, report_path=report_path)

    assert save_calls == []
    assert candidate.read_bytes() == database_before
    assert report.new_ids == ()
    assert json.loads(report_path.read_text(encoding="utf-8")) == {
        "absent_existing_ids": [404],
        "candidate_path": str(candidate.resolve()),
        "database_path": str(database.resolve()),
        "existing_duplicate_ids": [],
        "existing_slurry_rows": 2,
        "filtered_rows": 1,
        "new_ids": [],
        "retained_ids": [101],
        "source_path": str(source.resolve()),
        "source_rows": 1,
    }


def test_candidate_accepts_digit_string_ids_in_existing_slurry(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(8231, PROJECT, "existing"), (9000, PROJECT, "new")])
    _database(database, ids=(8231,))
    workbook = load_workbook(database)
    try:
        workbook["Slurry"]["A4"] = "8231"
        workbook.save(database)
    finally:
        workbook.close()

    report = build_candidate(source, database, candidate)

    assert report.retained_ids == (8231,)
    assert report.new_ids == (9000,)
    workbook = load_workbook(candidate, read_only=True, data_only=False)
    try:
        assert workbook["Slurry"]["A4"].value == "8231"
        assert workbook["Slurry"]["A5"].value == 9000
    finally:
        workbook.close()


def test_candidate_appends_slurry_rows_only_to_existing_width(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(9000, PROJECT, "new")])
    workbook = load_workbook(source)
    try:
        workbook["log"].cell(1, 88, "source-only metadata")
        workbook.save(source)
    finally:
        workbook.close()
    _database(database)

    build_candidate(source, database, candidate)

    workbook = load_workbook(candidate, read_only=True, data_only=False)
    try:
        assert workbook["Slurry"].max_column == 87
        assert workbook["Slurry"].cell(5, 88).value is None
    finally:
        workbook.close()


def test_candidate_ignores_invalid_ids_in_filtered_out_projects(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(None, "Other", "ignored"), (0, "Other", "ignored"), (303, PROJECT, "new")])
    _database(database)

    report = build_candidate(source, database, candidate)

    assert report.new_ids == (303,)


def test_candidate_reports_existing_duplicate_ids_without_rewriting_them(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(303, PROJECT, "new")])
    _database(database, ids=(101, 101))

    report = build_candidate(source, database, candidate)

    assert report.existing_duplicate_ids == (101,)
    workbook = load_workbook(candidate, read_only=True, data_only=False)
    try:
        assert [workbook["Slurry"].cell(row, 1).value for row in range(4, 7)] == [101, 101, 303]
        assert [workbook["db_table"].cell(row, 1).value for row in range(3, 6)] == [101, 101, 303]
    finally:
        workbook.close()


@pytest.mark.parametrize("bad_id", [None, 0, -1, 1.5, "12", "not-an-id"])
def test_candidate_rejects_invalid_filtered_ids(tmp_path, bad_id):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(bad_id, PROJECT, "bad")])
    _database(database)

    with pytest.raises(CandidateBuildError, match="invalid source ID"):
        build_candidate(source, database, candidate)

    assert not candidate.exists()


def test_candidate_does_not_require_an_unstated_source_header(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(303, PROJECT, "new-row")])
    workbook = load_workbook(source)
    try:
        workbook["log"]["A2"] = "not-required"
        workbook.save(source)
    finally:
        workbook.close()
    _database(database)

    report = build_candidate(source, database, candidate)

    assert report.new_ids == (303,)


def test_candidate_rejects_duplicate_filtered_ids(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    _source(source, [(303, PROJECT, "first"), (303, PROJECT, "second")])
    _database(database)

    with pytest.raises(CandidateBuildError, match="duplicate source ID 303"):
        build_candidate(source, database, candidate)

    assert not candidate.exists()


def test_candidate_rejects_in_place_or_production_paths(tmp_path):
    source = tmp_path / "source.xlsx"
    database = tmp_path / "database.xlsx"
    _source(source, [(101, PROJECT, "present")])
    _database(database)

    with pytest.raises(CandidateBuildError, match="separate output"):
        build_candidate(source, database, database)


def test_candidate_rejects_exact_production_database_before_access(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    candidate = tmp_path / "candidate.xlsx"
    accesses = []

    def fail_if_accessed(*args, **kwargs):
        accesses.append((args, kwargs))
        raise AssertionError("production workbook must not be accessed")

    monkeypatch.setattr(candidate_pipeline, "load_workbook", fail_if_accessed)
    monkeypatch.setattr(candidate_pipeline.shutil, "copy2", fail_if_accessed)

    with pytest.raises(CandidateBuildError, match="cannot read from or write to production"):
        build_candidate(source, PRODUCTION_DATABASE_PATH, candidate)

    assert accesses == []
