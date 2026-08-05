# Cell Log to Cellpy Database Updater

This repository currently provides a deliberately small interim candidate-only slice for the legacy local `log` workflow. It reads a local legacy log workbook, keeps the fixed project filter, appends new IDs to existing `Slurry`, appends minimal `db_table` rows, preserves all existing `db_table` cells including `b01:b07`, and writes a separate candidate workbook and optional report.

## Current status

The usable candidate command is:

```powershell
python main.py candidate --source <source.xlsx> --database <non-production-copy.xlsx> --output <candidate.xlsx> [--report <candidate.json>]
```

Included source projects are fixed in code: `SIS-Larger`, `SIS-Large`, `CellMap`, `Norse-HV`, `SUMBAT-SP5`, `SUMBAT`, and `ASAP`.

For each new ID, the candidate appends the source row to `Slurry` and a minimal `db_table` row containing only: column `A` = the literal ID, `D` (`exists`) = `1`, `T` (`instrument`) = `arbin_sql_h5`, and `X` (`experiment_type`) = `cycling`. All other new `db_table` cells are blank; no formulas are added. Existing rows and cells are preserved, including `b01:b07`.

The source workbook and starting database copy are untouched. The exact production workbook (`C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx`) is rejected as an input or output, in-place output is rejected, and production replacement is unavailable. Legacy flags such as `--apply`, `--config`, `--skip-sharepoint`, `--stage-only`, or `--maintenance` are rejected.

Duplicate or invalid IDs in included projects fail the build. Invalid IDs in filtered-out projects are ignored. Existing duplicate `Slurry` IDs and existing IDs absent from the filtered source are retained and reported in the result/report.

The separate validation-only command remains intentionally unavailable:

```powershell
python main.py validate
```

It exits without accessing SharePoint, the local snapshot, or the production database.

This is not the production updater. SharePoint `c&p` acquisition, full `db_table` mapping/formulas, Excel recalculation, Cellpy qualification, and production apply remain deferred. No representative live success is claimed.

## Governing documents

- `PROJECT_SCOPE.md` - authoritative scope and safety boundaries.
- `IMPLEMENTATION_PLAN.md` - Stage 0 plus Gates A-E.
- `SOURCE_CONTRACT.md` - exact source and values-only mirror contract.
- `DB_FIELD_MAPPING.md` - complete 73-column ownership inventory.
- `MIGRATION_DECISIONS.md` - user-approved migration decisions.
- `IMPLEMENTATION_PLAN_ADVERSARIAL_REVIEW.md` - production-risk review.
- `STAGE0_SAFETY_EVIDENCE.md` - evidence that legacy write paths are disabled.

## Verification

```powershell
python -m pytest -q
```

Production replacement will be introduced last, only after the source, mapping, candidate, Excel/Cellpy, transaction, backup, and rollback gates pass.
