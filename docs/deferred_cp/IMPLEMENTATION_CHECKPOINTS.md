# Implementation checkpoints

Brief, test-backed notes for the first ten implementation tasks in
`../../IMPLEMENTATION_PLAN.md`.

## Task 1 — Resolve the hard-coded source identity

- Added `src.source_acquisition.resolve_source_identity`.
- It sends only `SHAREPOINT_WORKBOOK_URL` to an injected resolver and requires
  both the SharePoint drive ID and item ID.
- Tests: `python -m pytest tests/test_source_acquisition.py -q` — 2 passed.

### Adversarial open points

- [ ] Resolve the approved live URL with the real Graph response and retain the drive/item identity; abort if a later resolution points at a different item.
- [ ] Add negative tests for non-object, blank, and whitespace-only identity fields and prove no session opens after any identity failure.

## Task 2 — Record source metadata

- Added immutable `SourceMetadata` and `write_source_manifest` support.
- Metadata requires the expected workbook and exact `c&p` sheet, then records
  drive ID, item ID, eTag, last-modified time, workbook name, sheet name, and
  used-range address.
- Tests: `python -m pytest tests/test_source_acquisition.py -q` — 3 passed.

### Adversarial open points

- [ ] Validate the live used-range geometry and returned values, not only an address string: exact `c&p!A1:S<last-row>` shape, row/column counts, headers, and units must agree.
- [ ] Extend the accepted manifest with retrieval time, content hash, row count, and ID statistics, then write and reopen it atomically while preserving the last accepted manifest on failure.
- [ ] Compare source identity/eTag before and after the read and abort if the workbook changes during acquisition.

## Task 3 — Stop on authentication or network failure

- Added straight-line `acquire_source` orchestration with no local/download
  fallback and no recovery that could continue with stale data.
- Resolution failures propagate before a workbook session is opened.
- Tests: `python -m pytest tests/test_source_acquisition.py -q` — 4 passed.

### Adversarial open points

- [ ] Inject failures at drive resolution, used-range retrieval, session creation, value retrieval, and session close; prove the prior snapshot, manifest, and production workbook remain unchanged.
- [ ] Wire the straight-line failure contract through the eventual CLI and verify that no retry, stale-source fallback, or partial local artifact can continue the run.

## Task 4 — Keep source acquisition read-only

- Added `ReadOnlyGraphResolver` backed by an injected GET-only transport.
- It can resolve the share URL and inspect the `c&p` used range, but exposes
  no POST, PUT, PATCH, DELETE, or source-write operation.
- Tests: `python -m pytest tests/test_source_acquisition.py -q` — 5 passed.

### Adversarial open points

- [ ] Prove the values-reading transport and workbook provider used by the real acquisition path are GET/read-only as well; the current resolver test does not constrain an injected `session_factory`.
- [ ] Retain live before/after source non-mutation evidence and the approved permission scope; `GET` calls alone do not prove least privilege.

## Task 5 — Use a non-persistent workbook session

- Added `NonPersistentWorkbookSession`, which always creates sessions with
  `persist_changes=False`, delegates reads only, and closes in `__exit__`.
- Tests: `python -m pytest tests/test_source_acquisition.py -q` — 6 passed.

### Adversarial open points

- [ ] Verify against the actual provider that `persist_changes=False` is sent on session creation and reject providers that cannot prove that setting.
- [ ] Restrict session reads to the approved `c&p` range and test cleanup when value retrieval or session close fails; the current wrapper accepts arbitrary sheet/range arguments.

## Task 6 — Store credentials in the OS/user credential mechanism

- Added lazy `KeyringCredentialStore` using the platform keyring selected by
  the `keyring` package.
- Missing credentials or the optional package fails explicitly; there is no
  file, environment, or repository fallback.
- Tests: `python -m pytest tests/test_credentials.py -q` — 3 passed.

### Adversarial open points

- [ ] Wire live acquisition to retrieve its token through `KeyringCredentialStore`; prove no CLI, config, environment, raw-token, or logging path can supply or expose credentials.
- [ ] Verify the operator machine uses an approved OS-backed keyring rather than an insecure fallback, and document token account, scopes, expiry, and refresh behavior.

## Task 7 — Define permitted value types for A:S

- Added `SOURCE_COLUMN_CONTRACTS` covering every mirrored column.
- A is integer-shaped; textual source fields are explicit; measurement and
  range fields are numeric; blanks and Excel errors are represented explicitly
  for all columns.
- Tests: `python -m pytest tests/test_source_validation.py -q` — 1 passed.

### Adversarial open points

- [ ] Approve a per-column field/type/unit/null dictionary from live source evidence; generic `text` and `number` categories do not establish semantic correctness.
- [ ] Reject non-finite and unsupported values, and add parameterized tests for every column plus metadata/header/units validation before any mirror is accepted.

## Task 8 — Define error handling for every A:S column

- Added explicit `COLUMN_ERROR_POLICIES`: A:P reject mapping-critical Excel
  errors; Q:S preserve helper/range errors in the values-only mirror.
- Added focused row validation without coercion, plus pure typed source/manifest
  diagnostics retained by the repository’s broader validation suite.
- Tests: `python -m pytest tests/test_source_column_types.py tests/test_source_validation.py -q` — 16 passed.

### Adversarial open points

- [ ] Add parameterized coverage for every Excel error token in every column: A:P must reject, while Q:S must preserve the exact value and source location without coercion.
- [ ] Record error counts in validation/manifest evidence and prove preserved Q:S errors can never enter a database mapping without an explicit approved rule.

## Task 9 — Confirm the local snapshot contains only `c&p`

- Added `src.snapshot.write_snapshot`, which writes a validated values-only
  workbook with exactly one `c&p` sheet and reopens it before replacement.
- Formula text and extra sheets are rejected; temporary files are closed and
  cleaned up safely on Windows.
- Tests: `python -m pytest tests/test_snapshot.py -q` — 1 passed.

### Adversarial open points

- [ ] Reopen the snapshot with values-only and compare dimensions, metadata, every value, and content hash with the accepted source manifest; sheet count and formula absence are insufficient.
- [ ] Inject save, reopen, replacement, and concurrent-writer failures and prove the previous accepted snapshot remains intact with no orphan temporary files.

## Task 10 — Confirm the fixed local snapshot path

- Added `snapshot_path`, which resolves only to the hard-coded
  `source_data/Cell_Log_CP.xlsx` path under the workspace root.
- Tests: `python -m pytest tests/test_snapshot.py -q` — 2 passed.

### Adversarial open points

- [ ] Resolve the fixed path from the workspace root rather than process CWD and reject runtime path/root overrides; the current test seam accepts an arbitrary `root`.
- [ ] Add a CLI-level negative test proving no argument or environment setting can redirect the snapshot to another workbook, sheet, or directory.

## Boundary note

The implementations use injected test doubles and do not claim a live
SharePoint/Graph resolution, credential retrieval, or source comparison. Gate A
remains blocked on that live evidence and the still-open anomaly decisions in
`../../IMPLEMENTATION_PLAN.md`.

The adversarial open points above are also not
production-readiness evidence until their artifacts and tests exist.

## Evidence hygiene

- [ ] Refresh the historical per-task pass counts after implementation changes; record current focused and full-suite results with the verification date.
