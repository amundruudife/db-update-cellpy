# Implementation Plan: Hard-Coded One-Way Cell Log Update

Status: simplified legacy candidate slice implemented; production remains blocked at Gates A and B  
Authority: `PROJECT_SCOPE.md`  
Adversarial review: `IMPLEMENTATION_PLAN_ADVERSARIAL_REVIEW.md`  
Decision artifacts: `SOURCE_CONTRACT.md`, `DB_FIELD_MAPPING.md`, `MIGRATION_DECISIONS.md`  
Last revised: 2026-08-05

## Simplified interim slice

The immediate implementation is intentionally smaller than the production
design below. It exists to prove the basic two-sheet append workflow against
local legacy copies before refining the `db_table` mapping or implementing the
SharePoint `c&p` migration.

Implemented scope:

- read an explicitly selected local workbook's legacy `log` sheet as evaluated values;
- retain the existing hard-coded project filter;
- append only valid, unique, filtered IDs that are absent from `Slurry`;
- leave every existing `Slurry` and `db_table` cell unchanged;
- append one minimal `db_table` row per new ID with only `id`, `exists`,
  `instrument`, and `experiment_type` populated;
- keep `b01:b07` blank for new rows and preserved for existing rows;
- retain and report existing duplicate or currently absent IDs without deleting them;
- write and reopen a separate candidate plus an optional JSON report; and
- reject in-place paths and the exact production workbook as either input or output.

Deferred from this interim slice:

- SharePoint access and the 19-column `c&p` source contract;
- full replacement of `Slurry`;
- the complete `db_table` formula and field mapping;
- Excel recalculation and Cellpy qualification; and
- backup, rollback, and production replacement.

The production gate design below remains relevant only for that later
production-capable workflow. It is not the implementation checklist for the
interim candidate command.

## Current Gate A/B evidence status

Repository evidence is integrated below; this status does not declare either
gate accepted. The current verification command is `python -m pytest -q` → 70
passed. Stage 0 remains complete and production-write capability remains
disabled.

Gate A repository-complete items include strict source validation, anomaly-ledger
recording, manifest append-only checks, read-only acquisition boundaries,
non-persistent session behavior, atomic snapshot preservation, and injected
authentication-failure preservation. Gate A remains blocked by live SharePoint
identity/schema comparison, source-owner correction or approval for duplicate
and invalid IDs, ID `8206` reconciliation, and first-run manifest approval.

Gate B repository-complete items include the read-only production inventory,
topology manifest, first-migration cell-change artifact, exact-layout fixtures,
golden mapping tests, and ID-keyed preservation tests. Gate B remains blocked by
unresolved mappings, formula/source-range and calculated-value approval,
representative Cellpy reconciliation, and approval of the first-migration diff
and topology allowlist.

## Production verdict

**NO-GO:** production-write capability must remain mechanically disabled until Gates A through E have passed and every Critical and High adversarial finding is closed with recorded evidence.

The target steady-state operation is:

1. read evaluated values from the hard-coded SharePoint workbook and exact `c&p` sheet;
2. atomically replace the hard-coded local snapshot;
3. validate the source schema and IDs;
4. build and recalculate a candidate copy of the hard-coded Cellpy database;
5. validate formulas, manual data, workbook structure, and Cellpy behavior;
6. create and verify a rollback backup; and
7. atomically replace production only when every gate succeeds.

No run may write to the SharePoint source. No run may use Downloads, an arbitrary path, or a different sheet as a fallback.

## Execution model

```mermaid
flowchart LR
    S0[Stage 0: disable legacy writes] --> A[Gate A: source and identity]
    A --> B[Gate B: database mapping]
    B --> C[Gate C: offline candidate]
    C --> D[Gate D: Excel and Cellpy]
    D --> E[Gate E: production transaction]
    E --> P[Production enabled]
```

Each gate requires:

- all listed tasks checked;
- its exit artifacts committed or otherwise retained;
- validation commands and results recorded;
- unresolved deviations recorded as blockers; and
- explicit approval where the gate says approval is required.

Detailed task completion is not sufficient if the gate-level acceptance criteria fail.

## Implementation strategy: controlled core rewrite

Keep the existing repository, packaging, history, and any small utility that passes the new contract. Replace the legacy workflow itself rather than incrementally adapting its append/filter pipeline.

Retain only after review and tests:

- logging setup and exception reporting;
- narrow path/hash/directory helpers that do not accept production overrides;
- packaging metadata and useful test-runner configuration; and
- general workbook-inspection helpers that preserve the new safety boundaries.

Replace or remove:

- configuration-driven source, target, sheet, and mapping selection;
- Downloads acquisition and fallback behavior;
- legacy `log`-sheet copying;
- project filtering that depends on the old `log` schema;
- skip-existing-key deduplication;
- positional append-only `update_slurry` behavior;
- optional backup and direct `shutil.copy2` production apply; and
- cleanup behavior that could remove accepted snapshots or rollback backups.

Build the new workflow as small contract-oriented modules:

1. `contracts` — hard-coded identities, rows, columns, and system values;
2. `source_acquisition` — non-persistent Graph extraction of evaluated `c&p` values;
3. `source_validation` — schema, error, ID, and append-only checks;
4. `snapshot` — atomic values-only local mirror and manifest;
5. `field_mapping` — reviewed 73-column ownership/formula definitions;
6. `candidate_builder` — `Slurry` replacement and keyed `db_table` reconstruction;
7. `excel_calculation` — bounded desktop Excel recalculation;
8. `candidate_validation` — topology, values, formulas, preservation, and Cellpy checks;
9. `production_transaction` — lock, fingerprint, backup, journal, atomic replace, and guarded rollback; and
10. `cli` — validation-only first, production enablement last.

Use red-green tests for each module. Do not provide a compatibility shim that keeps legacy CLI flags or legacy production behavior reachable.

## Evidence already established

- [x] Source workbook inspected at `C:\Users\IFE13213\Downloads\Cell_Log.xlsx`.
- [x] Exact source sheet is lowercase `c&p`.
- [x] `c&p` has 4,590 rows and 19 columns, with headers on row 2, units on row 3, and data/formulas from row 4.
- [x] `c&p` depends on formulas in other source sheets, so the local snapshot must contain evaluated values.
- [x] Source column A `key` corresponds to `db_table` column A `id`.
- [x] The production target is `C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx`.
- [x] The production sheets are `Slurry` and `db_table`.
- [x] `b01:b07` in columns F:L are the only user-input columns.
- [x] Other non-formula cells are preservation-only for existing IDs and blank for new IDs unless `DB_FIELD_MAPPING.md` explicitly reclassifies them.
- [x] Current `db_table` contains 361 populated IDs; current `Slurry` contains 367 populated IDs.
- [x] The inspected source contains 4,582 positive numeric key rows, 17 duplicated positive IDs, two zero-key rows, and formula-error key values.
- [x] Existing database ID `8206` is absent from the inspected evaluated source.
- [x] Source order differs from current database order, so preservation must be keyed by ID.
- [x] `c&p` has no `experiment_type` field.
- [x] Current formula-owned columns are A, C, O, P, Q, R, S, U, V, Z, AA, and AG.
- [x] Legacy formulas address the old 87-column `Slurry` layout and cannot be reused with the new 19-column mirror.

## Stage 0: Remove the existing production hazard

Purpose: make it impossible for legacy or partially implemented code to overwrite production.

- [x] Add focused failing tests proving the old `--apply`, `--config`, `--skip-sharepoint`, arbitrary path, and arbitrary sheet behaviors are rejected.
- [x] Remove or disable the current `shutil.copy2(..., config["db_path"])` production path.
- [x] Remove optional-backup behavior from every production-capable path; no production-capable path remains in Stage 0.
- [x] Remove the Downloads fallback from runnable code.
- [x] Prevent direct calls to the legacy append operation from targeting the hard-coded production workbook.
- [x] Retain one validation-only entry point, currently an explicit not-implemented shell.
- [x] Run focused safety tests and the retained full suite: `python -m pytest -q` → 11 passed on 2026-08-04.

**Exit evidence:** `STAGE0_SAFETY_EVIDENCE.md`; tests demonstrate that no current CLI or legacy function can modify the production workbook through the retired paths.

## Gate A: Source and identity contract

Purpose: establish exactly what records and cells constitute the authoritative source.

Primary artifact: `SOURCE_CONTRACT.md`  
Decision record: `MIGRATION_DECISIONS.md`

### A1. Live source identity and access

- [x] Approve delegated Microsoft Graph for the simplest working implementation, conditional on non-persistence and no source update requests.
- [ ] Resolve the hard-coded URL to the immutable drive item identity (repository resolver/tests use injected doubles; no live Graph response is retained).
- [ ] Record drive item ID, eTag, last-modified timestamp, workbook name, sheet name, and used-range address from the live source.
- [x] Prove that authentication or network failure stops without fallback.
- [x] Prove that acquisition code contains no source update operation.
- [x] Use a non-persistent Excel workbook session if Microsoft Graph is approved.
- [x] Record upstream values-only export as the fallback design only if the approved Graph spike cannot work without modifying source.
- [ ] Store credentials only in an operator-approved OS/user credential mechanism (repository keyring abstraction and tests exist; machine approval is not retained).

### A2. Exact mirror contract

- [x] Copy the values-only mirror as A:S.
- [ ] Copy rows 1 through 3 exactly as evaluated/displayed values and metadata rows (the current snapshot writer proves values-only structure, not full source metadata fidelity).
- [x] Define permitted value types for each included column.
- [x] Define error handling for every included column.
- [x] Treat the mirror as an exact evaluated A:S copy; restrict database mappings to approved A:P fields.
- [x] Confirm that only `c&p` appears in the local snapshot.
- [x] Confirm local snapshot path `source_data/Cell_Log_CP.xlsx`.

### A3. Record identity and population

- [x] Include every unique positive integer ID; v1 has no legacy project/type filter because `c&p` has no authoritative filter field.
- [ ] Resolve all 17 duplicate positive IDs or establish a different key.
- [ ] Resolve zero, blank, and formula-error key rows.
- [ ] Reconcile existing database ID `8206` against the live source.
- [x] Abort on an absent existing/historical ID until explicitly reconciled.
- [x] Preserve accepted source order in `Slurry` and use the same order in `db_table`.
- [x] Define allowed changes between successive source manifests.
- [ ] Generate the expected first-run ID manifest and row count.

### A4. Acquisition and snapshot tests

- [x] Test exact workbook and sheet identity.
- [x] Test values and value types without source formulas.
- [x] Test duplicate, invalid, missing, reordered, and unexpectedly removed IDs.
- [x] Test stale or changed eTag behavior.
- [x] Test atomic snapshot write and reopen.
- [x] Test that source authentication failures leave the accepted snapshot and production unchanged.
- [ ] Compare the first live acquisition with the inspected Downloads workbook and reopen decisions if materially different.

**Gate A acceptance criteria**

- [ ] `SOURCE_CONTRACT.md` contains no unresolved blocking decision.
- [ ] Source access has been approved by the responsible owner.
- [ ] The anomaly ledger contains a disposition for every duplicate/invalid ID and ID `8206`.
- [ ] The expected first-run ID set and row count are approved.

**Gate A blockers:** the repository records provisional local anomalies and
strict failure behavior, but it does not contain live Graph identity/schema
evidence, source-owner dispositions for 17 duplicate IDs and five invalid key
rows, reconciliation for ID `8206`, or an approved first-run manifest.

**Exit artifacts:** approved source contract, anomaly ledger, source manifest, and expected first-run ID manifest.

## Gate B: Database ownership and migration contract

Purpose: define every target cell before implementing row generation.

Primary artifact: `DB_FIELD_MAPPING.md`  
Decision record: `MIGRATION_DECISIONS.md`

### B1. Complete field dictionary

- [ ] Assign one approved ownership class to each of the 73 `db_table` columns: unresolved entries remain explicit blockers.
- [ ] For every formula column, approve source field, formula template, units, conversion, type, and null policy.
- [x] For every currently approved system column, record the literal or deterministic rule; unresolved formula/mapping decisions remain open.
- [x] Confirm F:L are manual and never contain formulas or automated values.
- [x] Confirm preservation-only cells are copied by ID for existing records and blank for new records.
- [x] Map `c&p.B channel` to AW `channel`; preserve S `tester` for existing IDs and blank it for new IDs.
- [ ] Resolve `loading_active_material` versus `assumed capacity working electrode`.
- [ ] Resolve `batch`, `nominal_capacity`, `cell_type`, `group`, `project`, and `comment_slurry` under the new source schema.
- [x] Set X `experiment_type` to the approved system constant `cycling`.
- [x] Map BE `schedule` from `c&p` column L, subject to representative Cellpy verification.
- [x] Set D `exists = 1` and T `instrument = "arbin_sql_h5"` as approved system rules.

### B2. Existing-data and workbook-topology inventory

- [x] Inventory all populated non-formula cells outside F:L by column and value pattern.
- [x] Produce a cell-level retain/recreate/delete report for the first migration; approval remains open.
- [x] Capture sheet names/order, dimensions, styles, tables, names, validations, comments, merged cells, hidden state, filters, panes, widths/heights, print settings, external links, macros, and calculation properties as applicable.
- [x] Define an allowlist of workbook-package changes; approval remains open.
- [ ] Approve the source-to-target row offset used by formula templates.

### B3. Golden examples and tests

- [x] Create a small exact-layout source fixture.
- [x] Create a database fixture containing every ownership class and nonblank manual/preservation cells.
- [x] Create golden rows covering all 73 target columns.
- [ ] Record expected formula strings and calculated values.
- [ ] Test unit conversions and null behavior.
- [x] Test preservation by ID when source order changes.
- [x] Test that new IDs receive blank manual and preservation-only cells.
- [x] Test idempotence: unchanged logical input produces no row churn or ownership changes.

**Gate B acceptance criteria**

- [ ] Every row in `DB_FIELD_MAPPING.md` is approved and contains no `TBD`.
- [ ] The first-migration cell-level change report is approved.
- [ ] Golden examples reconcile with representative production rows and Cellpy semantics.
- [ ] The workbook-topology change allowlist is approved.

**Gate B blockers:** eight mapping entries remain unresolved; the baseline has
an offset-7 formula outlier; calculated values and representative Cellpy
semantics are not verified; and the first-migration diff and topology allowlist
are evidence artifacts, not approvals.

**Exit artifacts:** approved 73-column field dictionary, golden mapping fixture, existing-data inventory, and topology manifest/allowlist.

## Gate C: Offline candidate builder

Purpose: implement the transformation without production-write capability.

### C1. Hard-coded contract

- [ ] Add a single constants/contract module.
- [ ] Hard-code the SharePoint workbook identity, `c&p`, local snapshot path, production path, `Slurry`, `db_table`, row boundaries, columns, and ID fields.
- [ ] Encode the approved field dictionary in code.
- [ ] Require code and test changes for any future path, year, sheet, or schema change.
- [ ] Reject runtime overrides.

### C2. Validated snapshot

- [ ] Retrieve evaluated values and value types using the approved acquisition mechanism.
- [ ] Write only the approved range and sheet into a temporary workbook.
- [ ] Validate schema, headers, units, dimensions, keys, value types, and errors.
- [ ] Generate a manifest containing source identity, eTag, range, retrieval time, row count, ID statistics, and content hash.
- [ ] Atomically replace the accepted local snapshot only after reopen and validation.

### C3. Candidate construction

- [ ] Fingerprint production before reading it.
- [ ] Copy production to a unique candidate on the same volume.
- [ ] Reject duplicate or invalid existing database IDs.
- [ ] Snapshot manual and preservation-only values by ID.
- [ ] Clear and rewrite the existing `Slurry` sheet object without deleting or renaming it.
- [ ] Rebuild one `db_table` data row per approved source ID in the approved order.
- [ ] Generate formula cells only from approved templates.
- [ ] Populate system-owned cells only from approved rules.
- [ ] Restore manual and preservation-only values to existing IDs.
- [ ] Leave manual and preservation-only cells blank for new IDs.
- [ ] Reapply approved formatting and workbook structures.
- [ ] Save and reopen only the candidate.

### C4. Offline verification

- [ ] Compare candidate IDs one-to-one with the approved source ID manifest.
- [ ] Compare every preserved manual and preservation-only cell by ID.
- [ ] Verify formula and system ownership patterns.
- [ ] Verify the workbook topology diff matches the allowlist.
- [ ] Generate machine-readable and human-readable migration reports.
- [ ] Run focused tests, fixture integration tests, and a sanitized full-structure rehearsal.

**Gate C acceptance criteria**

- [ ] Production-write functionality remains disabled.
- [ ] Candidate generation succeeds on a full-structure copy.
- [ ] All negative and idempotence tests pass.
- [ ] Candidate validation reports contain no unexplained differences.

**Exit artifacts:** candidate workbook, candidate manifest, topology diff, migration diff, and test report.

## Gate D: Excel and Cellpy qualification

Purpose: prove that formulas are calculated and Cellpy reads the intended values.

### D1. Environment contract

- [ ] Record supported Windows, Python, Excel, Cellpy, and `pywin32` versions.
- [ ] Add preflight checks for Excel availability, calculation mode, Cellpy import/configuration, disk space, and candidate/backup directories.
- [ ] Confirm Cellpy validation is read-only or isolate any output it creates.
- [ ] Define timeouts and operator-visible failure messages.

### D2. Dedicated Excel calculation

- [ ] Launch a dedicated hidden Excel application instance owned by the run.
- [ ] Disable alerts, events, link updates, and macros where supported.
- [ ] Open only the candidate.
- [ ] force a full calculation rebuild and wait for completion with a bounded timeout.
- [ ] Save, close, and quit cleanly.
- [ ] Guarantee cleanup after success, error, prompt, calculation timeout, and save failure.
- [ ] Record Excel version and calculation state.

### D3. Post-calculation validation

- [ ] Reopen formulas and verify approved patterns.
- [ ] Reopen cached values and reject mapping-critical Excel errors or missing caches.
- [ ] Load the candidate through Cellpy without modifying it.
- [ ] Verify representative IDs and every mapped Cellpy field.
- [ ] Run a full-size candidate rehearsal on the operator machine.
- [ ] Inject Excel unavailable, timeout, orphan-process, calculation-error, and Cellpy-failure cases.

**Gate D acceptance criteria**

- [ ] The qualified environment passes normal and injected-failure runs.
- [ ] Formula strings and cached values are correct.
- [ ] Cellpy reads the expected values from the candidate.
- [ ] No unexpected topology or source/production mutation occurs.

**Exit artifact:** environment-qualified Excel and Cellpy rehearsal report.

## Gate E: Transaction and production authorization

Purpose: add production replacement as the last, separately reviewed capability.

### E1. Transaction controls

- [ ] Introduce an exclusive updater lock covering final fingerprint, backup, replace, post-check, and rollback.
- [ ] Refuse a locked/open target or concurrent updater.
- [ ] Re-fingerprint production immediately before backup and abort on change.
- [ ] Stage candidate and replacement file on the same volume as production.
- [ ] Create a transaction journal containing source/candidate/production identities and state transitions.
- [ ] Create, hash, reopen, and validate a timestamped backup.
- [ ] Close every Excel/workbook handle before replacement.
- [ ] Replace production with a same-volume atomic rename.

### E2. Post-replacement checks and guarded rollback

- [ ] Reopen production and rerun structural, topology, formula-cache, manual/preservation, ID, and Cellpy smoke checks.
- [ ] Before rollback, verify production still has the exact hash written by the current transaction.
- [ ] Never restore over an unrecognized or concurrently changed hash.
- [ ] Restore the verified backup automatically only while the transaction lock is held.
- [ ] Verify the restored hash equals the original production hash.
- [ ] Preserve the transaction journal and validation reports after failure.

### E3. Operational authorization

- [ ] Define backup location, retention, disk-space threshold, and recovery instructions.
- [ ] Define log redaction and retention.
- [ ] Complete an intentional backup-and-rollback drill.
- [ ] Produce the first-run dry report listing retained, added, excluded, duplicate, invalid, absent, and changed IDs/cells.
- [x] Require manual approval of the first-run dry report; actual approval remains a runtime Gate E artifact.
- [ ] Enable production replacement in a separate reviewed code change.
- [ ] Keep a validation-only command that can never replace production.

**Gate E acceptance criteria**

- [ ] Transaction, crash, concurrency, and rollback tests pass.
- [ ] Backup restoration reproduces the original production hash.
- [ ] The first-run report and production enablement change are approved.
- [ ] The operator checklist and recovery procedure are complete.

**Exit artifacts:** signed go-live record, verified rollback evidence, first-run report, and production enablement change.

## Mandatory stop conditions

Abort without modifying production if:

- source authentication, identity, schema, range, eTag, or freshness is unexpected;
- a fallback source or runtime override is requested;
- an ID is invalid, duplicated, unexpectedly removed, or outside the approved population;
- any field mapping or ownership rule is unresolved;
- production is open, locked, changed, or cannot be exclusively controlled;
- a candidate save, topology check, formula check, Excel calculation, cached-value check, manual/preservation comparison, or Cellpy check fails;
- backup creation or verification fails;
- rollback would overwrite a hash not written by the current transaction; or
- any stage-gate artifact or required approval is missing.

## Adversarial-review closure matrix

| Finding | Closed by |
|---|---|
| AR-01 migration population | Gate A3 and first-run ID manifest |
| AR-02 duplicate IDs | Gate A3 anomaly ledger |
| AR-03 ID 8206 / append-only contradiction | Gate A3 absent-ID policy |
| AR-04 incomplete 73-column contract | Gate B1 field dictionary |
| AR-05 source permission boundary | Gate A1 access approval |
| AR-06 ambiguous mirror contract | Gate A2 exact range/error policy |
| AR-07 incomplete transaction | Gate E1/E2 lock, journal, guarded rollback |
| AR-08 reachable legacy production path | Stage 0 and Gate C |
| AR-09 unexplained literals | Gate B2 existing-data inventory |
| AR-10 workbook topology loss | Gate B2 and Gate C4 |
| AR-11 Excel automation failures | Gate D |
| AR-12 unrealistic fixtures | Gate C4 and Gate D3 |
| AR-13 stale Downloads evidence | Gate A4 live comparison |
| AR-14 row ordering | Gate A3 and idempotence tests |
| AR-15 checklist false confidence | Gate-level artifacts and acceptance criteria |
| AR-16 incomplete operations contract | Gate D1 and Gate E3 |

## Definition of done

- [ ] All five gates have passed with retained evidence.
- [ ] Every Critical and High adversarial finding is closed.
- [ ] Source acquisition is approved and proven non-mutating.
- [ ] Every target column has one approved hard-coded ownership rule.
- [ ] Manual and preservation-only values are proven unchanged by ID.
- [ ] The candidate is recalculated by qualified desktop Excel and read correctly by Cellpy.
- [ ] Workbook topology changes match the approved allowlist.
- [ ] Backup, lock, journal, replacement, and guarded rollback are tested.
- [ ] Production is enabled only by the separately approved final change.
- [ ] Scope, decision artifacts, implementation plan, code, tests, and operator documentation agree.
