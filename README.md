# Cell Log to Cellpy Database Updater

This repository currently provides verified candidate-only legacy V1. It reads a local `log` workbook, filters the approved projects, appends new IDs to a separate candidate copy, preserves existing workbook data, verifies the saved workbook can be reopened, and can write a JSON report. It never writes to the source workbook or production database.

## Use

```powershell
python main.py candidate --source ... --database ... --output ... [--report ...]
```

`--source` must be a local workbook containing the `log` sheet. `--database` must be a non-production workbook containing `Slurry` and `db_table`. `--output` is a separate candidate workbook path. `--report` is optional JSON output.

The candidate keeps source rows for the approved projects (`SIS-Larger`, `SIS-Large`, `CellMap`, `Norse-HV`, `SUMBAT-SP5`, `SUMBAT`, and `ASAP`). New `Slurry` rows and minimal `db_table` rows are appended; existing rows, including manual `b01:b07` cells, are preserved. New system values are limited to `exists=1`, `instrument=arbin_sql_h5`, and `experiment_type=cycling`; no formulas are added.

The exact production workbook is rejected, and source/database/output paths must be distinct. Invalid or duplicate included-source IDs fail the build. Existing duplicate IDs and existing IDs absent from the filtered source are retained and reported.

## Disabled commands and scope

`python main.py` and `python main.py validate` are intentionally disabled and return exit code 3. Retired flags such as `--apply`, `--config`, `--skip-sharepoint`, `--stage-only`, and `--maintenance` are rejected.

This is non-production output only. SharePoint `c&p` acquisition, Cellpy qualification, Excel recalculation or automation, complete `db_table` refinement, and production replacement/apply are deferred. See [PROJECT_SCOPE.md](PROJECT_SCOPE.md) and the [deferred SharePoint `c&p` archive](docs/deferred_cp/README.md).

## Install and verify

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Runtime requires `openpyxl>=3.1.5`. Development verification adds `pytest` and the repository Graphify tooling.
