# Gate A: next-ten implementation checkpoints

Short red-green notes for the first ten unchecked items in
`../../IMPLEMENTATION_PLAN.md`. A checkpoint is marked blocked when repository code
is complete but an external source or approval is still required.

## Task 1 — duplicate positive IDs

- Added `src.anomaly_ledger.build_anomaly_ledger`; duplicate IDs are grouped by
  source row and receive `source_correction_required` without selecting a row.
- Recorded the 17 provisional duplicate findings in
  `SOURCE_ANOMALY_LEDGER.md`.
- Red: `ModuleNotFoundError` before the ledger module existed.
- Green: `python -m pytest tests/test_anomaly_ledger.py -q` — 1 passed.
- Gate status: **blocked** until the live source owner corrects the rows or
  approves a different stable key.

## Task 2 — zero, blank, and formula-error key rows

- Reused the existing strict key validator through
  `build_anomaly_ledger`; no coercion or fallback is allowed.
- Added a regression test proving blank, `#REF!`, and zero keys retain their
  exact source rows/values and all require source correction.
- Test: `python -m pytest tests/test_anomaly_ledger.py -q` — 2 passed.
- Gate status: **blocked** pending correction of the five provisional invalid
  key rows recorded in `SOURCE_ANOMALY_LEDGER.md`.

## Task 3 — existing database ID `8206`

- Extended the anomaly-ledger test to require the distinct
  `reconcile_before_first_run` disposition for an existing ID absent from the
  accepted source set.
- Test: `python -m pytest tests/test_anomaly_ledger.py -q` — 3 passed.
- The local production workbook contains `8206`, while the provisional local
  source does not. The live SharePoint comparison could not run because no
  browser/connector session is available, so this reconciliation remains
  **blocked** and production writes stay disabled.

## Task 4 — allowed changes between source manifests

- Added immutable `SourceManifest` and `validate_manifest_change` in
  `src/source_manifest.py`.
- Allowed changes are a changed source version (`eTag`, timestamp, content
  hash) and an appended ID suffix that preserves historical order. Drive item,
  workbook/sheet identity, and a shrinking or malformed A:S used range are
  rejected.
- Red: `ModuleNotFoundError` before the manifest module existed.
- Green: `python -m pytest tests/test_source_manifest.py -q` — 2 passed.

## Task 5 — expected first-run ID manifest and row count

- Added `SourceManifest` and `build_first_manifest` in
  `src/source_manifest.py`; it records the accepted source-order IDs and full
  used-range row count, and raises on any duplicate/invalid/missing historical
  ID instead of filtering it.
- Red: missing `SourceManifestError`/builder, followed by a fixture-helper
  failure; both were corrected before the green run.
- Green: `python -m pytest tests/test_source_manifest.py -q` — 4 passed.
- Gate status: **blocked** because the current provisional source has 17
  duplicate IDs, five invalid key rows, and the unresolved database ID `8206`;
  no approved first-run manifest is emitted.
- Follow-up hardening rejects metadata used-range row-count mismatches; latest
  focused result: `python -m pytest tests/test_source_manifest.py -q` — 10
  passed.

## Task 6 — exact workbook and sheet identity

- Added a strict `c&p!A1:S<last-row>` used-range check in
  `resolve_source_metadata`, alongside the existing exact workbook and sheet
  name checks.
- Red: the wrong-sheet-address regression initially did not raise.
- Green: `python -m pytest tests/test_source_acquisition.py -q` — 8 passed.

## Task 7 — evaluated values and value types without formulas

- Added `validate_values_only` and made `write_snapshot` use it as the single
  pre-write boundary.
- Formula text is rejected across metadata and data cells; valid integer,
  float, text, blank, and preserved error values are returned without coercion.
- Red: the new validator import failed before implementation.
- Green: `python -m pytest tests/test_source_column_types.py tests/test_snapshot.py -q` — 6 passed.

## Task 8 — duplicate, invalid, missing, reordered, and removed IDs

- Added a manifest-change regression matrix covering duplicate IDs, historical
  removal, reordering, and insertion before the append boundary; existing source
  validation tests cover blank/zero/error/non-numeric keys.
- All such changes are rejected before a manifest can be accepted.
- Test: `python -m pytest tests/test_source_manifest.py tests/test_source_validation.py -q` — 22 passed.

## Task 9 — stale or changed eTag behavior

- Added `SourceChangedError` and a before/after metadata comparison around the
  source read. eTag, last-modified time, drive item, workbook/sheet identity,
  and used range must all remain unchanged.
- Red: importing the new exception failed before implementation.
- Green: `python -m pytest tests/test_source_acquisition.py -q` — 10 passed,
  including changed-eTag rejection and unchanged-source success.

## Task 10 — atomic snapshot write and reopen

- Added failure-injection tests for validation failure and `os.replace` failure;
  both preserve the previously accepted snapshot and clean the temporary file.
- The replace test verifies the temporary and destination paths share the
  snapshot directory, then reopens the accepted workbook.
- Test: `python -m pytest tests/test_snapshot.py -q` — 4 passed.

## Task 11 — authentication failure preserves accepted state

- Added `acquire_and_stage_snapshot` in `src/source_acquisition.py`; it
  completes source acquisition before entering the atomic snapshot-write
  boundary and has no fallback or production path.
- The regression test seeds an accepted snapshot and production sentinel,
  injects an authentication failure, and verifies both byte-for-byte files are
  unchanged.
- Red: test collection failed with
  `ImportError: cannot import name 'acquire_and_stage_snapshot'`.
- Green: `python -m pytest tests/test_source_acquisition.py -q` — 11 passed;
  Gate A focused suite — 46 passed.
- Gate status remains **blocked** by live SharePoint comparison, unresolved
  duplicate/invalid IDs and `8206`, and approval of the expected first-run
  manifest; no accepted first manifest or production write is claimed.
