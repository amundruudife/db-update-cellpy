# Future Cleanup (Out of Current Gate Scope)

This is a deferred cleanup register only. None of the items below is required
to close Gate A or Gate B, and none is implemented by the current change.

## Observed redundant or legacy structure

- Legacy workflow modules remain alongside the replacement-oriented modules:
  `src/config.py`, `src/data_processing.py`, `src/database.py`,
  `src/file_operations.py`, `src/common_utils.py`, and
  `src/copy_sharepoint_file.py`. Stage 0 tests disable the dangerous paths, but
  the files are still present.
- `src/cleanup_old_files.py` contains separate cleanup behavior for old output
  and source copies. It is not part of the current gate workflow and should not
  be made production-reachable without a separate safety review.
- Legacy/project documentation is duplicated across
  `IMPLEMENTATION_CHECKPOINTS.md`, `GATE_A_NEXT_10_CHECKPOINTS.md`, and
  `update_log.txt`. These representations can drift.
- Mapping and workbook semantics are represented in more than one place:
  `PROJECT_SCOPE.md`, `DB_FIELD_MAPPING.md`, the field-mapping tests/fixtures,
  and the read-only Gate B JSON artifacts.

## Observed artifact and data duplication

- `source_data` contains many dated `*_cellog_*.xlsx` copies in addition to
  `Cell_Log.xlsx`.
- `output` contains numerous `dryrun_*.xlsx` and `backup_*.bak.xlsx` files.
- `update_log.txt` is approximately 244 KB; the Gate B inventory and
  first-migration JSON artifacts are approximately 1.07 MB and 761 KB.
- `graphify-out` is generated navigation/audit output and duplicates some
  repository knowledge in generated form.

Possible later cleanup includes consolidating documentation authority,
retiring unreachable legacy modules after a dependency audit, defining a
retention policy for workbook copies, and reducing or archiving oversized
generated logs/artifacts. Those actions require separate scope, ownership, and
data-retention decisions.
