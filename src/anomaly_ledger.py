"""Structured Gate A findings for source and first-run identity review."""

from dataclasses import dataclass
from typing import Iterable, Tuple

from .source_validation import (
    ValidationDiagnostic,
    validate_first_manifest,
    validate_source_values,
)


@dataclass(frozen=True)
class AnomalyFinding:
    """One source or first-run anomaly and its required disposition."""

    code: str
    message: str
    rows: Tuple[int, ...] = ()
    values: Tuple[str, ...] = ()
    disposition: str = "source_correction_required"


def _finding(diagnostic: ValidationDiagnostic, disposition: str) -> AnomalyFinding:
    return AnomalyFinding(
        code=diagnostic.code,
        message=diagnostic.message,
        rows=diagnostic.rows,
        values=diagnostic.values,
        disposition=disposition,
    )


def build_anomaly_ledger(
    values, existing_database_ids: Iterable[int] = ()
) -> Tuple[AnomalyFinding, ...]:
    """Return deterministic findings without repairing or selecting source rows."""

    source_result = validate_source_values(values)
    findings = []
    for diagnostic in source_result.diagnostics:
        findings.append(_finding(diagnostic, "source_correction_required"))

    first_manifest = validate_first_manifest(
        source_result.ids,
        existing_database_ids=existing_database_ids,
    )
    for diagnostic in first_manifest.diagnostics:
        if diagnostic.code == "missing_existing_database_ids":
            findings.append(_finding(diagnostic, "reconcile_before_first_run"))
        elif diagnostic.code != "duplicate_manifest_ids":
            findings.append(_finding(diagnostic, "source_correction_required"))
    return tuple(findings)
