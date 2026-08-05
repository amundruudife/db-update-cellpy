import pytest

from src.source_acquisition import SourceMetadata
from src.contracts import KEY_COLUMN, SOURCE_COLUMN_COUNT
from src.source_manifest import (
    SourceManifest,
    SourceManifestError,
    build_first_manifest,
    validate_manifest_change,
)


def _row(value):
    row = [None] * SOURCE_COLUMN_COUNT
    row[KEY_COLUMN - 1] = value
    return row


def _manifest(*, ids, etag="e1", row_count=None, drive_id="drive", item_id="item", used_range=None):
    total_rows = row_count if row_count is not None else len(ids) + 3
    return SourceManifest(
        drive_id=drive_id,
        item_id=item_id,
        workbook_name="Cell_Log.xlsx",
        sheet_name="c&p",
        used_range=used_range or f"c&p!A1:S{total_rows}",
        etag=etag,
        last_modified="2026-08-04T10:00:00Z",
        content_hash="hash-1",
        ids=tuple(ids),
        row_count=total_rows,
    )


def test_manifest_change_allows_only_metadata_version_change_and_appended_ids():
    previous = _manifest(ids=(11, 12))
    current = _manifest(ids=(11, 12, 13), etag="e2", row_count=6)

    result = validate_manifest_change(previous, current)

    assert result.is_valid
    assert result.ids == (11, 12, 13)


def test_manifest_change_rejects_source_identity_changes():
    previous = _manifest(ids=(11, 12))
    current = _manifest(ids=(11, 12), drive_id="other-drive")

    result = validate_manifest_change(previous, current)

    assert not result.is_valid
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        "source_identity_changed"
    ]


def test_manifest_change_rejects_shrunk_or_malformed_used_ranges():
    previous = _manifest(ids=(11, 12))
    shrunk = _manifest(ids=(11, 12), used_range="c&p!A1:S4", row_count=4)
    malformed = _manifest(ids=(11, 12), used_range="log!A1:S5")

    shrunk_result = validate_manifest_change(previous, shrunk)
    malformed_result = validate_manifest_change(previous, malformed)

    assert not shrunk_result.is_valid
    assert "used_range_shrank" in [
        diagnostic.code for diagnostic in shrunk_result.diagnostics
    ]
    assert not malformed_result.is_valid
    assert "invalid_used_range" in [
        diagnostic.code for diagnostic in malformed_result.diagnostics
    ]


def test_first_manifest_contains_accepted_ids_and_full_source_row_count():
    metadata = SourceMetadata(
        drive_id="drive",
        item_id="item",
        etag="e1",
        last_modified="2026-08-04T10:00:00Z",
        workbook_name="Cell_Log.xlsx",
        sheet_name="c&p",
        used_range="c&p!A1:S5",
    )

    manifest = build_first_manifest(
        [
            _row("metadata"),
            _row("key"),
            _row("units"),
            _row(11),
            _row(12),
        ],
        metadata=metadata,
        content_hash="hash-1",
    )

    assert manifest.ids == (11, 12)
    assert manifest.row_count == 5
    assert manifest.used_range == "c&p!A1:S5"


def test_first_manifest_refuses_anomalous_source_instead_of_filtering_it():
    metadata = SourceMetadata(
        drive_id="drive",
        item_id="item",
        etag="e1",
        last_modified="2026-08-04T10:00:00Z",
        workbook_name="Cell_Log.xlsx",
        sheet_name="c&p",
        used_range="c&p!A1:S5",
    )

    with pytest.raises(SourceManifestError, match="duplicate_ids"):
        build_first_manifest(
            [_row("metadata"), _row("key"), _row("units"), _row(11), _row(11)],
            metadata=metadata,
            content_hash="hash-1",
        )


def test_first_manifest_rejects_metadata_row_count_mismatch():
    metadata = SourceMetadata(
        drive_id="drive",
        item_id="item",
        etag="e1",
        last_modified="2026-08-04T10:00:00Z",
        workbook_name="Cell_Log.xlsx",
        sheet_name="c&p",
        used_range="c&p!A1:S4",
    )

    with pytest.raises(SourceManifestError, match="row_count_mismatch"):
        build_first_manifest(
            [_row("metadata"), _row("key"), _row("units"), _row(11), _row(12)],
            metadata=metadata,
            content_hash="hash-1",
        )


@pytest.mark.parametrize(
    ("current_ids", "expected_code"),
    [
        ((11, 12, 12, 13), "duplicate_manifest_ids"),
        ((11, 12), "missing_historical_ids"),
        ((12, 11, 13), "reordered_historical_ids"),
        ((11, 12, 14, 13), "new_ids_not_appended"),
    ],
)
def test_manifest_change_rejects_duplicate_invalid_or_non_append_only_ids(
    current_ids, expected_code
):
    previous = _manifest(ids=(11, 12, 13))
    current = _manifest(ids=current_ids, row_count=len(current_ids) + 3)

    result = validate_manifest_change(previous, current)

    assert not result.is_valid
    assert expected_code in [diagnostic.code for diagnostic in result.diagnostics]
