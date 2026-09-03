# Cell Log to Cellpy Database Updater

This repository provides a gated source-acquisition step and verified candidate-only legacy V1. The acquisition step reads the approved SharePoint `c&p` range through a non-persistent session and stages validated values in `source_data/Cell_Log_CP.xlsx`. The candidate workflow reads a local workbook, filters the approved projects, appends new IDs to a separate candidate copy, preserves existing workbook data, verifies the saved workbook can be reopened, and can write a JSON report. It never writes to the SharePoint source or production database.

## Use

```powershell
python main.py acquire --account ife-user
python main.py candidate --source ... --database ... --output ... [--report ...] [--cellpy-ready]
python main.py candidate --source ... --database ... --output ... --neware-source Neware_log.xlsx [--neware-manifest ...]
```

`acquire` retrieves the Graph access token from the OS keyring account, resolves the approved site and default-drive item, reads only the fixed workbook and `c&p!A1:S<last-row>` range, and replaces the local snapshot only when the source identity is stable and the returned row array exactly matches the terminal row. Install `keyring` and store the token without putting it in a command line or file:

```powershell
python -m keyring set ife-cell-log-updater ife-user
```

`--source`, `--database`, and `--output` must be separate `.xlsx` workbooks. `.xlsm` and OOXML workbooks containing `xl/vbaProject.bin` are rejected without modification. `--report` is optional JSON output. `--cellpy-ready` is an opt-in Windows-only path that uses Excel COM (pywin32>=312) to recalculate the temporary candidate before publishing it.

The candidate keeps source rows for the approved projects (`SIS-Larger`, `SIS-Large`, `CellMap`, `Norse-HV`, `SUMBAT-SP5`, `SUMBAT`, and `ASAP`). New `Slurry` rows and minimal `db_table` rows are appended; existing rows, formulas, and manual `b01:b07` cells are preserved. The default path writes only `exists=1`, `instrument=arbin_sql_h5`, and `experiment_type=cycling` for new database rows. With `--cellpy-ready`, new rows additionally receive the example's 12 `db_table` formulas and are accepted only after Excel recalculates and cached-value checks pass.

The exact production workbook is rejected, and source/database/output paths must be distinct. Invalid or duplicate included-source IDs fail the build. Resolved `db_table` IDs must reconcile exactly to every `Slurry` ID plus validated manifest-backed Neware IDs; missing, orphan, duplicate, malformed, or overlapping IDs fail closed.

An optional `--neware-source` reads only the `test_log` sheet, excludes and reports its two exact template rows, and appends the usable Neware rows directly to `db_table`. Neware records use unique `Cell test label` values and a persisted `source_data/neware_id_manifest.json` ID map; see [docs/NEWARE_MAPPING.md](docs/NEWARE_MAPPING.md). Neware input cannot be combined with `--cellpy-ready`.

## Disabled commands and scope

`python main.py` and `python main.py validate` are intentionally disabled and return exit code 3. Retired flags such as `--apply`, `--config`, `--skip-sharepoint`, `--stage-only`, and `--maintenance` are rejected.

This is non-production output only. Publication is ordered and restart-recoverable, not a same-transaction or multi-file/power-loss atomic operation; a failed interruption may leave old-or-new artifacts and a rerun completes the staged set. The opt-in Excel recalculation path is candidate-only and fails closed when Excel/COM is unavailable; it does not claim full semantic Cellpy qualification. Live source identity/anomaly approval, delegated `Sites.Read.All` permission, old-link revocation, production replacement/apply, and complete `db_table` refinement remain deferred. See [PROJECT_SCOPE.md](PROJECT_SCOPE.md) and the [deferred SharePoint `c&p` archive](docs/deferred_cp/README.md).

## Install and verify

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Runtime requires `openpyxl>=3.1.5`; Windows `--cellpy-ready` additionally requires `pywin32>=312` and a local Excel installation. Development verification adds `pytest` and the repository Graphify tooling.
