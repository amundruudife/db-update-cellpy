#!/usr/bin/env python3
"""Safe CLI shell for the replacement Cell Log updater.

Stage 0 exposes candidate-only legacy local log input while production writes
remain disabled. The validation command becomes functional only after the
source and mapping contracts have passed their gates.
"""

import argparse
import sys

from src.candidate_pipeline import CandidateBuildError, build_candidate


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
    candidate = commands.add_parser(
        "candidate",
        help="Build a non-production candidate workbook",
    )
    candidate.add_argument("--source", required=True, help="Source log workbook")
    candidate.add_argument("--database", required=True, help="Non-production database workbook")
    candidate.add_argument("--output", required=True, help="Candidate workbook output path")
    candidate.add_argument("--report", help="Optional JSON report output path")
    parser.set_defaults(command="validate")
    return parser


def parse_arguments(argv=None):
    return build_parser().parse_args(argv)


def main(argv=None) -> int:
    arguments = parse_arguments(argv)
    if arguments.command == "candidate":
        try:
            report = build_candidate(
                arguments.source,
                arguments.database,
                arguments.output,
                report_path=arguments.report,
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
            f"existing duplicates={len(report.existing_duplicate_ids)}"
        )
        return 0

    print(
        "Validation-only workflow is not implemented yet. "
        "Legacy acquisition and production update paths are disabled."
    )
    return EXIT_NOT_IMPLEMENTED


if __name__ == "__main__":
    raise SystemExit(main())
