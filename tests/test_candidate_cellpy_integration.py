import importlib.util
import os
import sys
from pathlib import Path

import pytest

from src.candidate_pipeline import build_candidate


@pytest.mark.skipif(
    sys.platform != "win32" or importlib.util.find_spec("cellpy") is None,
    reason="Cellpy integration requires Windows and an installed Cellpy runtime",
)
def test_cellpy_ready_candidate_is_readable_by_database_reader(tmp_path):
    try:
        from win32com.client import DispatchEx
    except Exception as exc:  # pragma: no cover - depends on the operator runtime
        pytest.skip(f"Excel COM is unavailable: {exc}")

    excel = None
    try:
        excel = DispatchEx("Excel.Application")
    except Exception as exc:  # pragma: no cover - depends on the operator runtime
        pytest.skip(f"Excel is unavailable: {exc}")
    finally:
        if excel is not None:
            excel.Quit()

    repository_root = Path(__file__).resolve().parents[1]
    source = Path(
        os.environ.get(
            "CELLLOG_CANONICAL_SOURCE",
            str(repository_root / "source_data" / "Cell_Log.xlsx"),
        )
    )
    database = repository_root / "example" / "2025_Cell_Analysis_db_001.xlsx"
    if not source.is_file() or not database.is_file():
        pytest.skip("canonical Cell Log and example database are not available")

    candidate = tmp_path / "cellpy-ready-candidate.xlsx"
    report = build_candidate(source, database, candidate, cellpy_ready=True)
    assert report.cellpy_ready is True
    assert report.recalculated is True

    from cellpy.readers import dbreader

    reader = dbreader.Reader(db_file=str(candidate))
    selected_ids = reader.select_batch("norse", "b02")
    assert len(selected_ids) > 0
    assert all(isinstance(record_id, int) and record_id > 0 for record_id in selected_ids)
