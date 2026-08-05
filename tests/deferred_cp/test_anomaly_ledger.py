from src.deferred_cp.anomaly_ledger import build_anomaly_ledger
from src.deferred_cp.contracts import KEY_COLUMN, SOURCE_COLUMN_COUNT


def _row(key):
    row = [None] * SOURCE_COLUMN_COUNT
    row[KEY_COLUMN - 1] = key
    return row


def _table(*keys):
    return [
        _row("metadata"),
        _row("key"),
        _row("units"),
        *[_row(key) for key in keys],
    ]


def test_duplicate_ids_are_grouped_and_require_source_correction():
    ledger = build_anomaly_ledger(_table(11, 12, 12, 13, 11))

    assert [finding.code for finding in ledger] == ["duplicate_ids"]
    assert ledger[0].values == ("11", "12")
    assert ledger[0].rows == (4, 8, 5, 6)
    assert ledger[0].disposition == "source_correction_required"


def test_invalid_key_rows_are_retained_as_explicit_blocking_findings():
    ledger = build_anomaly_ledger(_table(None, "#REF!", 0))

    assert [finding.code for finding in ledger] == [
        "blank_id",
        "excel_error_id",
        "non_positive_id",
    ]
    assert [finding.rows for finding in ledger] == [(4,), (5,), (6,)]
    assert [finding.values for finding in ledger] == [("None",), ("#REF!",), ("0",)]
    assert all(
        finding.disposition == "source_correction_required" for finding in ledger
    )


def test_missing_existing_database_id_requires_reconciliation_before_first_run():
    ledger = build_anomaly_ledger(_table(11, 12), existing_database_ids=(11, 8206))

    assert len(ledger) == 1
    assert ledger[0].code == "missing_existing_database_ids"
    assert ledger[0].values == ("8206",)
    assert ledger[0].disposition == "reconcile_before_first_run"
