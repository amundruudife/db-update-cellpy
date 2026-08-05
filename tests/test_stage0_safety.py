"""Stage 0 tests: legacy code must not reach production or Downloads."""

from pathlib import Path

import pytest

import main
from src.contracts import PRODUCTION_DATABASE_PATH
from src.copy_sharepoint_file import copy_cell_log_to_source_data
from src.database import dry_run_full_pipeline, update_slurry
from src.exceptions import LegacyWorkflowDisabledError, ProductionWriteBlockedError


@pytest.mark.parametrize(
    "legacy_args",
    [
        ["--apply"],
        ["--config", "other.json"],
        ["--skip-sharepoint"],
        ["--stage-only"],
        ["--maintenance"],
    ],
)
def test_legacy_cli_flags_are_rejected(legacy_args):
    with pytest.raises(SystemExit) as exc_info:
        main.parse_arguments(legacy_args)

    assert exc_info.value.code == 2


def test_cli_exposes_only_safe_validation_placeholder():
    args = main.parse_arguments([])

    assert args.command == "validate"
    assert main.main([]) == main.EXIT_NOT_IMPLEMENTED


def test_legacy_downloads_acquisition_is_disabled():
    with pytest.raises(LegacyWorkflowDisabledError):
        copy_cell_log_to_source_data()


def test_legacy_pipeline_is_disabled():
    with pytest.raises(LegacyWorkflowDisabledError):
        dry_run_full_pipeline({})


def test_legacy_append_refuses_the_hard_coded_production_database():
    with pytest.raises(ProductionWriteBlockedError):
        update_slurry(
            [{"id": 1}],
            Path(PRODUCTION_DATABASE_PATH),
            "Slurry",
            dry_run=True,
        )
