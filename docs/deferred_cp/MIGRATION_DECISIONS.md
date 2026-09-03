# Migration Decisions and Required Approvals

Status: user decisions recorded; technical evidence still blocks Gates A and B
Last revised: 2026-08-04

This file separates decisions already given by the user from recommendations that still require confirmation.

## Confirmed decisions

| ID | Decision | Status |
|---|---|---|
| D01 | The workflow is one-way from Cell Log to the local database. | Confirmed |
| D02 | The updater must never edit the Cell Log source. | Confirmed |
| D03 | Source workbook, source sheet, local snapshot, production workbook, and target sheets are hard-coded. | Confirmed |
| D04 | The exact source sheet is lowercase `c&p`. | Confirmed |
| D05 | The local source snapshot contains evaluated values, not source formulas. | Confirmed |
| D06 | `Slurry` is completely replaced on every accepted run. | Confirmed |
| D07 | Source column A `key` and database column A `id` are the record identity. | Confirmed, subject to duplicate resolution |
| D08 | Source is expected to be append-only. | Confirmed, subject to ID `8206` reconciliation |
| D09 | `db_table` must contain one row for every unique positive integer ID in `c&p`; no legacy project/type filter in v1. | Confirmed |
| D10 | `b01:b07` are manual and preserved by ID. | Confirmed |
| D11 | Formulas and source extractions are hard-coded by target column. | Confirmed |
| D12 | Formula calculation is performed by desktop Excel. | Confirmed |
| D13 | Production replacement is automatic after every gate passes. | Confirmed |
| D14 | Existing non-formula cells outside F:L are preservation-only for existing IDs and blank for new IDs unless explicitly reclassified. | Confirmed in scope |

## Recorded answers

### Q1 — Which source records enter `db_table`?

**Recommended:** every unique positive integer ID in `c&p` after invalid/duplicate source rows have been corrected. This likely expands the database from 361 rows to roughly 4,500 rows.

Alternative: define a hard-coded filter with an authoritative source column and allowed values. No such filter is currently known.

**Decision:** Approved for v1. Mirror and process every unique positive integer ID. Do not retain the legacy project filter in v1 because `c&p` has no authoritative project or type column. A later filter requires an explicit `c&p` field, code change, and matching tests.

### Q2 — What exactly is copied into the values-only mirror?

**Recommended:** copy the complete evaluated A:S used range, including rows 1–3, preserving source order and errors. Permit database mappings only from A:P. Reject the run if an error occurs in column A or any mapped source cell.

Alternative: copy only A:P, making `Slurry` a normalized business extract rather than a complete `c&p` mirror.

**Decision:** Approved. Copy the complete evaluated A:S used range, including rows 1–3. Only A:P may feed database mappings unless the contract is deliberately revised.

### Q3 — How are duplicate or invalid IDs handled?

**Recommended:** abort and require correction in Cell Log. Never choose first/last, merge, renumber, or maintain a silent local exception.

**Decision:** Approved. Abort with a clear validation error. IDs must be valid and unique; no automatic correction is permitted.

### Q4 — What happens when an existing database ID is absent from source?

This currently applies to ID `8206`.

**Recommended:** abort until the discrepancy is explained. After reconciliation, use a strict append-only rule: an existing accepted ID may not disappear.

Alternatives after investigation: retain the row with an explicit system status, or remove it with one-time written approval.

**Decision:** Approved. Abort on an unexpectedly absent historical ID. ID `8206` must be reconciled before the first accepted migration.

### Q5 — Which ordering should the rebuilt database use?

**Recommended:** preserve accepted `c&p` source order in both `Slurry` and `db_table`. Manual/preservation cells still join by ID. This keeps formula row offsets simple and makes unchanged input idempotent.

Alternative: ascending numeric ID order.

**Decision:** Approved. Preserve source order in `Slurry` and propagate the same accepted-ID order into `db_table`. Ordering is not a business key; all preservation remains ID-based.

### Q6 — Are the proposed system values authoritative?

**Recommended:**

- `exists = 1`
- `instrument = "arbin_sql_h5"`
- `experiment_type = "cycling"`

The source contains no `experiment_type` column, so `cycling` cannot be extracted directly from `c&p`.

**Decision:** Approved system rules: `exists = 1`, `instrument = "arbin_sql_h5"`, and `experiment_type = "cycling"`. These are updater-owned values, not manual fields and not inferred from unrelated `c&p` columns.

### Q7 — Where does source `channel` belong?

**Recommended:** map `c&p.B channel` to `db_table.AW channel`. Preserve existing `db_table.S tester` for existing IDs and leave it blank for new IDs unless Cellpy requires a tester rule.

Alternative: map source channel to `tester`, or populate both under a documented Cellpy convention.

**Decision:** Approved recommendation. Map `c&p.B channel` to `db_table.AW channel`. Preserve existing `tester` by ID and leave it blank for new IDs unless later Cellpy evidence establishes a deterministic rule.

### Q8 — Is delegated Microsoft Graph `Sites.Read.All` acceptable?

**Recommended:** only if the responsible SharePoint/security owner explicitly approves the delegated read-only permission, the code uses a non-persistent workbook session, contains no update request, and verifies source identity before and after retrieval.

If not acceptable, require an approved upstream process to publish the standalone values-only `c&p` snapshot.

**Decision:** Approved for the simplest working implementation, conditional on `Sites.Read.All`, a non-persistent session, no update requests, fixed source identity, and before/after source metadata checks. The updater must not modify Cell Log; revoke the superseded old sharing link before live acceptance.

## Technical questions to resolve from evidence

These do not require guessing by the user, but they must be investigated and presented for approval:

- [ ] Explain all 17 duplicate positive IDs.
- [ ] Explain the zero and formula-error keys.
- [ ] Reconcile ID `8206` against the live workbook.
- [ ] Record exact headers for AJ:AV and BN:BT.
- [ ] Derive or reject mappings for C, P, Q, R, S, U, V, and AG.
- [ ] Confirm `mass_active_material = c&p.E × 1000` with units and representative rows.
- [ ] Confirm `cell = c&p.C`, `file_name_indicator = c&p.D`, and `schedule = c&p.L`.
- [ ] Inventory all non-formula values outside F:L and produce the first-migration loss/change report.
- [ ] Compare the live SharePoint source with the inspected Downloads copy.

## First production run

Proposed rule:

- Candidate-only rehearsals may be automated.
- The first production migration requires manual approval of the exact ID and cell-change report.
- Later runs may replace production automatically when all manifests, append-only checks, mappings, Excel checks, Cellpy checks, topology checks, lock checks, and backup checks pass.
- Any schema, mapping, target year/path, permission, or source-identity change returns the workflow to NO-GO.

**Decision:** Approved. The first production migration requires manual approval; later unchanged-contract runs may be automatic after all gates pass.

## Approved system-field policy

- `exists = 1` because an accepted source row exists in the mirrored source population.
- `instrument = "arbin_sql_h5"` as the hard-coded Cellpy instrument value for this workflow.
- `experiment_type = "cycling"` as the hard-coded experiment type represented by the `c&p` workflow.
- These fields are not added to the manual-input contract.
- A future change requires an explicit code, test, scope, and mapping-contract revision.
