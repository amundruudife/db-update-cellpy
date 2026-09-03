#!/usr/bin/env python3
"""Safe CLI shell for the replacement Cell Log updater.

Stage 0 exposes candidate-only legacy local log input while production writes
remain disabled. The validation command becomes functional only after the
source and mapping contracts have passed their gates.
"""

import argparse
import sys
from pathlib import Path

from src.candidate_pipeline import CandidateBuildError, CandidateReport, build_candidate
from src.deferred_cp.credentials import CredentialStoreError, KeyringCredentialStore
from src.deferred_cp.source_acquisition import (
    GraphHttpTransport,
    GraphRequestError,
    GraphWorkbookProvider,
    NonPersistentWorkbookSession,
    ReadOnlyGraphResolver,
    acquire_and_stage_snapshot,
)


EXIT_NOT_IMPLEMENTED = 3


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Validate the Cell Log to Cellpy update contract or build a "
            "candidate from a legacy local log workbook; production writes remain disabled"
        )
    )
    commands = parser.add_subparsers(dest="command")
    commands.add_parser(
        "validate",
        help="Run validation-only checks (currently disabled pending Gates A and B)",
    )
    acquire = commands.add_parser(
        "acquire",
        help="Copy the validated SharePoint c&p values into the local source snapshot",
    )
    acquire.add_argument("--account", required=True, help="OS keyring account containing the Graph token")
    acquire.add_argument(
        "--root",
        type=Path,
        default=Path("."),
        help="Workspace root for source_data/Cell_Log_CP.xlsx",
    )
    candidate = commands.add_parser(
        "candidate",
        help="Build a non-production candidate workbook",
    )
    candidate.add_argument("--source", required=True, help="Source log workbook")
    candidate.add_argument("--database", required=True, help="Non-production database workbook")
    candidate.add_argument("--output", required=True, help="Candidate workbook output path")
    candidate.add_argument("--report", help="Optional JSON report output path")
    candidate.add_argument(
        "--neware-source",
        help="Optional Neware workbook containing the test_log sheet",
    )
    candidate.add_argument(
        "--neware-manifest",
        "--neware-manifest-path",
        dest="neware_manifest",
        help="Optional Neware natural-key ID manifest path",
    )
    candidate.add_argument(
        "--cellpy-ready",
        action="store_true",
        help="Add example db_table formulas and recalculate through Windows Excel",
    )
    parser.set_defaults(command="validate")
    return parser


def parse_arguments(argv=None):
    return build_parser().parse_args(argv)


def main(argv=None) -> int:
    arguments = parse_arguments(argv)
    if arguments.command == "acquire":
        try:
            token = KeyringCredentialStore().get_token(arguments.account)
            transport = GraphHttpTransport()
            resolver = ReadOnlyGraphResolver(transport, access_token=token)
            provider = GraphWorkbookProvider(transport, access_token=token)

            def session_factory(metadata):
                provider.bind_workbook(metadata.drive_id, metadata.item_id)
                return NonPersistentWorkbookSession(
                    provider,
                    drive_id=metadata.drive_id,
                    item_id=metadata.item_id,
                )

            metadata, snapshot = acquire_and_stage_snapshot(
                resolver,
                session_factory,
                root=arguments.root,
            )
        except (CredentialStoreError, GraphRequestError, OSError, RuntimeError, ValueError) as exc:
            print(f"Source acquisition failed: {exc}", file=sys.stderr)
            return 1

        print(
            f"Source staged: {snapshot}; "
            f"drive_item={metadata.drive_id}/{metadata.item_id}; "
            f"used_range={metadata.used_range}"
        )
        return 0

    if arguments.command == "candidate":
        try:
            candidate_kwargs = {
                "report_path": arguments.report,
                "cellpy_ready": arguments.cellpy_ready,
            }
            if arguments.neware_source is not None:
                candidate_kwargs["neware_source_path"] = arguments.neware_source
            if arguments.neware_manifest is not None:
                candidate_kwargs["neware_manifest_path"] = arguments.neware_manifest
            report: CandidateReport = build_candidate(
                arguments.source,
                arguments.database,
                arguments.output,
                **candidate_kwargs,
            )
        except (CandidateBuildError, FileNotFoundError, OSError, ValueError) as exc:
            print(f"Candidate build failed: {exc}", file=sys.stderr)
            return 1

        print(
            f"Candidate built: {report.candidate_path}; "
            f"filtered={report.filtered_rows}; "
            f"retained={len(report.retained_ids)}; "
            f"new={len(report.new_ids)}; "
            f"absent={len(report.absent_existing_ids)}; "
            f"existing duplicates={len(report.existing_duplicate_ids)}; "
            f"cellpy-ready={report.cellpy_ready}; "
            f"recalculated={report.recalculated}"
        )
        if report.neware_source_path is not None:
            print(
                f"Neware: usable={report.neware_usable_rows}; "
                f"new={len(report.neware_new_ids)}; "
                f"retained={len(report.neware_retained_ids)}; "
                f"placeholders={len(report.neware_placeholder_rows)}"
            )
        return 0

    print(
        "Validation-only workflow is not implemented yet. "
        "Legacy acquisition and production update paths are disabled."
    )
    return EXIT_NOT_IMPLEMENTED


if __name__ == "__main__":
    raise SystemExit(main())
