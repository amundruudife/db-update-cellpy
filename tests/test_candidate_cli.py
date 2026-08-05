from types import SimpleNamespace

import pytest

import main
from src.candidate_pipeline import CandidateBuildError


def _candidate_args(*extra):
    return [
        "candidate",
        "--source",
        "source.xlsx",
        "--database",
        "database.xlsx",
        "--output",
        "candidate.xlsx",
        *extra,
    ]


def test_candidate_forwards_arguments_and_prints_summary(monkeypatch, capsys):
    calls = []
    report = SimpleNamespace(
        candidate_path="candidate.xlsx",
        filtered_rows=5,
        retained_ids=(101, 102),
        new_ids=(103, 104, 105),
        absent_existing_ids=(99,),
        existing_duplicate_ids=(101, 101),
    )

    def fake_build_candidate(*args, **kwargs):
        calls.append((args, kwargs))
        return report

    monkeypatch.setattr(main, "build_candidate", fake_build_candidate, raising=False)

    result = main.main(_candidate_args("--report", "candidate.json"))

    captured = capsys.readouterr()
    assert result == 0
    assert calls == [
        (
            ("source.xlsx", "database.xlsx", "candidate.xlsx"),
            {"report_path": "candidate.json"},
        )
    ]
    assert "candidate.xlsx" in captured.out
    assert "filtered=5" in captured.out
    assert "retained=2" in captured.out
    assert "new=3" in captured.out
    assert "absent=1" in captured.out
    assert "existing duplicates=2" in captured.out
    assert captured.err == ""


def test_candidate_requires_source_database_and_output():
    for omitted in ("--source", "--database", "--output"):
        arguments = _candidate_args()
        index = arguments.index(omitted)
        del arguments[index : index + 2]

        with pytest.raises(SystemExit) as exc_info:
            main.parse_arguments(arguments)

        assert exc_info.value.code == 2


def test_candidate_reports_build_failure_to_stderr(monkeypatch, capsys):
    calls = []

    def fake_build_candidate(*args, **kwargs):
        calls.append((args, kwargs))
        raise CandidateBuildError("unsafe workbook")

    monkeypatch.setattr(main, "build_candidate", fake_build_candidate, raising=False)

    result = main.main(_candidate_args())

    captured = capsys.readouterr()
    assert result == 1
    assert calls == [(tuple(_candidate_args()[2::2]), {"report_path": None})]
    assert captured.out == ""
    assert captured.err == "Candidate build failed: unsafe workbook\n"


@pytest.mark.parametrize("arguments", [[], ["validate"]])
def test_default_and_validate_remain_disabled(monkeypatch, capsys, arguments):
    def fail_if_called(*args, **kwargs):
        raise AssertionError("candidate builder must not run for validation")

    monkeypatch.setattr(main, "build_candidate", fail_if_called, raising=False)

    result = main.main(arguments)

    captured = capsys.readouterr()
    assert result == main.EXIT_NOT_IMPLEMENTED
    assert "Validation-only workflow is not implemented yet." in captured.out
    assert captured.err == ""


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
def test_legacy_flags_remain_rejected(legacy_args):
    with pytest.raises(SystemExit) as exc_info:
        main.parse_arguments(legacy_args)

    assert exc_info.value.code == 2
