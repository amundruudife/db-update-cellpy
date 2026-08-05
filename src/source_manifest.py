"""Immutable source-manifest data and append-only change checks."""

import re
from dataclasses import dataclass
from typing import Tuple

from .source_validation import (
    ManifestValidationResult,
    ValidationDiagnostic,
    validate_first_manifest,
    validate_manifest_continuity,
    validate_source_values,
)


@dataclass(frozen=True)
class SourceManifest:
    """The source identity, range, version, and accepted ID order."""

    drive_id: str
    item_id: str
    workbook_name: str
    sheet_name: str
    used_range: str
    etag: str
    last_modified: str
    content_hash: str
    ids: Tuple[int, ...]
    row_count: int


class SourceManifestError(ValueError):
    """Raised when a first-run manifest cannot be accepted."""

    def __init__(self, diagnostics):
        self.diagnostics = tuple(diagnostics)
        super().__init__(
            "; ".join(diagnostic.code for diagnostic in self.diagnostics)
            or "source manifest is invalid"
        )


_USED_RANGE = re.compile(
    r"^(?:c&p|'c&p')!A1:S(?P<last_row>[1-9][0-9]*)$",
)


def _used_range_last_row(address: str):
    match = _USED_RANGE.fullmatch(str(address).strip())
    return int(match.group("last_row")) if match else None


def validate_manifest_change(
    previous: SourceManifest, current: SourceManifest
) -> ManifestValidationResult:
    """Allow only source-version changes and appended IDs in source order."""

    diagnostics = list(validate_manifest_continuity(previous.ids, current.ids).diagnostics)

    immutable_fields = ("drive_id", "item_id", "workbook_name", "sheet_name")
    changed = tuple(
        field
        for field in immutable_fields
        if getattr(previous, field) != getattr(current, field)
    )
    if changed:
        diagnostics.append(
            ValidationDiagnostic(
                code="source_identity_changed",
                message="Source drive item and workbook identity must not change.",
                values=changed,
            )
        )

    previous_last_row = _used_range_last_row(previous.used_range)
    current_last_row = _used_range_last_row(current.used_range)
    if previous_last_row is None or current_last_row is None:
        diagnostics.append(
            ValidationDiagnostic(
                code="invalid_used_range",
                message="Source manifests must use an A1:S<last-row> used range.",
                values=(previous.used_range, current.used_range),
            )
        )
    elif current_last_row < previous_last_row:
        diagnostics.append(
            ValidationDiagnostic(
                code="used_range_shrank",
                message="The accepted source used range may not shrink.",
                values=(str(previous_last_row), str(current_last_row)),
            )
        )

    if current_last_row is not None and current.row_count != current_last_row:
        diagnostics.append(
            ValidationDiagnostic(
                code="row_count_mismatch",
                message="Manifest row count must equal the used-range row count.",
                values=(str(current.row_count), str(current_last_row)),
            )
        )

    return ManifestValidationResult(ids=current.ids, diagnostics=tuple(diagnostics))


def build_first_manifest(
    values,
    *,
    metadata,
    content_hash: str,
    existing_database_ids=(),
) -> SourceManifest:
    """Build an accepted first-run manifest or raise for any anomaly."""

    source_result = validate_source_values(values)
    diagnostics = list(source_result.diagnostics)
    diagnostics.extend(
        validate_first_manifest(
            source_result.ids,
            existing_database_ids=existing_database_ids,
        ).diagnostics
    )
    metadata_last_row = _used_range_last_row(metadata.used_range)
    if metadata_last_row is None:
        diagnostics.append(
            ValidationDiagnostic(
                code="invalid_used_range",
                message="Source metadata must use an A1:S<last-row> used range.",
                values=(metadata.used_range,),
            )
        )
    elif metadata_last_row != source_result.row_count:
        diagnostics.append(
            ValidationDiagnostic(
                code="row_count_mismatch",
                message="Manifest row count must equal the used-range row count.",
                values=(str(source_result.row_count), str(metadata_last_row)),
            )
        )
    if diagnostics:
        raise SourceManifestError(diagnostics)

    return SourceManifest(
        drive_id=metadata.drive_id,
        item_id=metadata.item_id,
        workbook_name=metadata.workbook_name,
        sheet_name=metadata.sheet_name,
        used_range=metadata.used_range,
        etag=metadata.etag,
        last_modified=metadata.last_modified,
        content_hash=content_hash,
        ids=source_result.ids,
        row_count=source_result.row_count,
    )
