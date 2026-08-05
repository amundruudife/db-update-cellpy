# Adversarial Review: Hard-Coded One-Way Cell Log Update Plan

Review date: 2026-08-04
Reviewed artifact: `../../IMPLEMENTATION_PLAN.md`
Verdict: **NO-GO for implementation that can write production**

The plan has the right safety direction, but it is not yet an executable contract. It still leaves decisions open that determine record identity, row population, field meaning, source access, and rollback safety. Those are design inputs, not details that can be resolved during implementation.

Offline spikes, fixtures, mapping analysis, and candidate-only code may proceed. Production-write capability must remain absent or mechanically disabled until every Critical and High finding below is closed with evidence.

## What the plan gets right

- Production is not used as a working file.
- `Slurry` replacement and `db_table` reconciliation are treated separately.
- Manual `b01:b07` values are preserved by ID, not by row position.
- Formula recalculation is assigned to desktop Excel.
- Source acquisition, mapping, validation, backup, and replacement are explicit gates.
- Known anomalies are recorded instead of silently normalized.

## Findings

### AR-01 — Critical: the migration population is undefined

**Failure scenario:** the first run expands `db_table` from 361 populated IDs to roughly 4,500 source IDs, even though it has not been established that every positive `c&p` key belongs in the Cellpy database.

The target outcome already says to create one database row per accepted source ID, while “accepted” is undefined. This makes the largest data migration decision look like a validation detail.

**Required correction:** approve and test a deterministic inclusion contract that produces an exact expected ID set and row count. The first-run dry report must list added, retained, excluded, duplicated, and removed IDs. A human must approve that report before the first production migration. Steady-state automation can be considered only afterward.

### AR-02 — Critical: column A is not currently a proven unique key

**Failure scenario:** two source rows share an ID but contain different data. Rejecting all duplicates prevents every run; selecting first or last corrupts one record; merging values invents data.

The observed source contains 17 duplicated positive IDs. This contradicts the assumed uniqueness needed to preserve manual data and generate one `db_table` row per ID.

**Required correction:** classify every duplicate as either a source defect or evidence that the real key is composite. Record the resolution in a reviewed anomaly ledger. Do not implement a first/last/merge policy. If a composite key is required, revisit the manual-data preservation contract before coding.

### AR-03 — Critical: the append-only assumption is contradicted by ID 8206

**Failure scenario:** the updater interprets a missing historical ID as a deletion and drops a valid database row, or preserves an obsolete row indefinitely.

Existing database ID `8206` is absent from the evaluated source. That can mean deletion, formula failure, stale source evidence, or an incorrect population rule. Any of those invalidates automatic historical-deletion checks until explained.

**Required correction:** reconcile ID `8206` against the live SharePoint workbook and its upstream source rows. Define an explicit absent-ID policy: abort, retain with status, or remove. Default must be abort.

### AR-04 — Critical: the 73-column database contract is not defined

**Failure scenario:** the script writes numerically plausible values into semantically wrong Cellpy fields. The known `loading_active_material` versus `assumed capacity` and `tester` versus `channel` ambiguities are examples.

A candidate mapping list is not enough to rebuild the authoritative database sheet. “Formula”, “system”, “manual”, and “blank” ownership also does not specify formulas, units, types, null behavior, or allowed constants.

**Required correction:** create a signed field dictionary for all 73 columns containing source field, ownership, units, conversion, formula template, null policy, expected type, and rationale. Validate it on representative existing rows through Cellpy. No database-row generator should be written before this artifact is approved.

### AR-05 — Critical: Graph access may violate the source access boundary

**Failure scenario:** a delegated token capable of writing the source workbook is used on a workflow whose stated security requirement is read-only access. A non-persistent Excel session reduces accidental persistence but does not remove token authority.

The plan correctly notes that the relevant workbook endpoint documents delegated `Files.ReadWrite`. An unchanged eTag and absence of update calls are useful evidence, but neither proves least privilege.

**Required correction:** obtain explicit approval for the permission model or use an upstream export mechanism whose identity has read-only access to the produced snapshot. State the threat model precisely: requesting only `c&p` is not the same as the Excel service accessing only that sheet, because calculation depends on other sheets.

### AR-06 — Critical: the source mirror contract is internally ambiguous

**Failure scenario:** one implementation copies A:S exactly, another copies business columns A:P, and both claim compliance. Formula errors or helper values are silently removed in one version and retained in another.

The plan says “complete agreed used range” but calls A:P the business contract. It also calls for a complete `Slurry` replacement while leaving noncritical errors open to blanking or rejection.

**Required correction:** define the exact snapshot and `Slurry` rectangle, including rows 1–3 and columns Q:S. Decide whether the mirror is an exact evaluated representation or a normalized business extract. Define error handling per included column. Name the sheet by its purpose in code; keep `Slurry` only because the production workbook requires that sheet name.

### AR-07 — Critical: the replacement is not an end-to-end transaction

**Failure scenario:** after the pre-replacement fingerprint, another user or process modifies the workbook; a failed post-check then restores the backup over that newer work.

`os.replace` can make one same-volume rename atomic. The complete operation—fingerprint, backup, replacement, Excel/Cellpy checks, and possible restoration—is not atomic.

**Required correction:** hold an exclusive transaction lock across the final fingerprint, backup, replace, post-check, and any rollback. Stage on the same volume. Before rollback, verify that production still has the exact hash written by this run. Add a small recovery journal so a crash can be diagnosed without guessing. Never restore over an unrecognized hash.

### AR-08 — High: automatic production mode is insufficiently isolated

**Failure scenario:** partially implemented code exposes `--apply`, or a legacy generic path still reaches production. The current code already accepts `--config`, `--skip-sharepoint`, and `--apply`; backup is optional and production is overwritten with `shutil.copy2` in `main.py:24-48` and `main.py:119-129`.

**Required correction:** remove the legacy apply path before adding new behavior. Initially ship only candidate generation and validation. Add production replacement in a later change guarded by a code-level release constant that defaults off. Add negative tests proving legacy flags, arbitrary targets, Downloads fallback, and direct legacy functions cannot update production.

### AR-09 — High: existing non-manual values may be destroyed without an inventory

**Failure scenario:** treating only `b01:b07` as manual is interpreted as permission to blank all other unexplained literals, including historical system values or corrected exceptions that remain operationally required.

**Required correction:** inventory every populated, non-formula cell outside `b01:b07`, grouped by column and value pattern. For the first migration, produce a cell-level loss/change report. Each such value must be recreated by an approved rule or explicitly approved for deletion.

### AR-10 — High: workbook topology preservation is under-specified

**Failure scenario:** cell values and formulas are correct, but Excel tables, names, validations, comments, styles, hidden state, filters, print settings, calculation properties, or other Open XML parts are damaged.

**Required correction:** capture a baseline workbook-topology manifest and compare the candidate package against it with an allowlist of intended changes. Include workbook/sheet properties, defined names, tables, validations, comments, merged cells, hidden rows/columns/sheets, panes, filters, dimensions, styles, VBA/external-link presence, and calculation settings as applicable.

### AR-11 — High: Excel automation needs a stricter failure contract

**Failure scenario:** Excel opens external links, displays a hidden prompt, uses manual calculation mode, loads an add-in, times out while still calculating, or leaves an orphan process holding the candidate.

**Required correction:** use a dedicated Excel application instance owned by the run; record Excel version and calculation mode; disable events, alerts, link updates, and macros where supported; bound open, calculation, save, and quit; confirm calculation completion; verify cached values after close; and prove cleanup under injected failures. Test on the actual operator machine and Excel version.

### AR-12 — High: fixture tests will not prove a real workbook survives

**Failure scenario:** small synthetic workbooks pass, while the 4 MB source and production workbook fail because of real styles, formulas, package parts, size, or Excel behavior.

**Required correction:** keep focused fixtures, but add a sanitized full-structure clone and a full-size candidate-only rehearsal. Acceptance must include Open XML topology diff, desktop Excel recalculation, reopen with formulas and cached values, and read-only Cellpy loading.

### AR-13 — High: the downloaded workbook is evidence, not production truth

**Failure scenario:** mappings and anomaly counts are frozen from a stale Downloads copy and differ on the first live SharePoint acquisition.

**Required correction:** bind the live source to drive item ID, eTag, and used-range address. On the first authorized acquisition, compare its schema, ID anomalies, and representative evaluated values with the inspected local workbook. Reopen the planning decisions if they differ.

### AR-14 — High: row ordering is unspecified

**Failure scenario:** the rebuilt table is semantically equivalent by ID but disrupts Excel users, Cellpy assumptions, or future diffs because its order changes on every run.

**Required correction:** define a stable order—normally ascending ID or approved source order—and test idempotence. A second run against unchanged inputs must produce the same logical workbook and no row churn.

### AR-15 — Medium: the checklist can create false confidence

The plan contains many implementation-level checkboxes but does not attach each gate to an artifact, approver, command, and evidence. Boxes can be completed independently while the system remains unsafe as a whole.

**Required correction:** reorganize execution around the five stage gates below. Keep detailed tasks underneath them, but do not use task count as readiness evidence.

### AR-16 — Medium: operational boundaries need completion

The plan does not yet fix supported Python, Excel, Cellpy, `pywin32`, and authentication versions; backup retention and disk-space requirements; log retention/redaction; or whether Cellpy validation itself creates side effects.

**Required correction:** add an environment preflight and an operations contract. Cellpy validation must target the candidate and be demonstrated read-only. Distinguish rollback backup from disaster recovery.

## Required stage gates

### Gate A — Source and identity contract

- [ ] Live source identity and freshness are proven.
- [ ] Exact copied range and error policy are approved.
- [ ] All duplicate IDs and ID `8206` are reconciled.
- [ ] Inclusion, exclusion, deletion, and row-order rules are approved.
- [ ] The source permission model is approved.

**Exit artifact:** source contract, anomaly ledger, and expected first-run ID manifest.

### Gate B — Database ownership and migration contract

- [ ] All 73 columns have reviewed ownership and transformation rules.
- [ ] Existing non-formula values have a retain/recreate/delete decision.
- [ ] The first-run cell-level change report is approved.
- [ ] Workbook topology allowlist is approved.

**Exit artifact:** field dictionary, golden mapping examples, and approved migration diff.

### Gate C — Offline candidate builder

- [ ] Production-write functionality does not exist or is mechanically disabled.
- [ ] Red-green tests cover keys, mappings, manual preservation, and idempotence.
- [ ] Candidate construction passes fixture and sanitized-clone checks.
- [ ] Legacy generic apply and fallback paths are unreachable.

**Exit artifact:** candidate workbook plus machine-readable validation report.

### Gate D — Excel and Cellpy qualification

- [ ] Dedicated Excel automation passes normal and injected-failure tests.
- [ ] Formula strings and cached values pass validation.
- [ ] Cellpy reads representative fields from a full-size candidate.
- [ ] Workbook topology changes match the allowlist.

**Exit artifact:** environment-qualified rehearsal report.

### Gate E — Transaction and production authorization

- [ ] Exclusive locking, same-volume staging, fingerprint checks, and recovery journal pass tests.
- [ ] Backup and guarded rollback drill reproduce the original hash.
- [ ] First-run dry report is manually approved.
- [ ] A separate reviewed change enables production replacement.

**Exit artifact:** signed go-live record and verified rollback evidence.

## Mandatory stop conditions

Abort without modifying production if any of these occurs:

- live source identity, schema, range, or eTag is unexpected;
- any ID is invalid, duplicated, unexpectedly removed, or outside the approved population;
- the production fingerprint changes during the run;
- the target is open, locked, or cannot be exclusively controlled;
- any mapping rule, formula cache, manual-field comparison, topology check, Excel calculation, or Cellpy check fails;
- backup verification fails;
- rollback would overwrite a production hash not written by the current transaction;
- the run uses a fallback source, arbitrary path, arbitrary sheet, or unapproved credentials.

## Recommended plan decision

Do not start with SharePoint or production replacement. First close Gates A and B using the downloaded workbook as provisional evidence, then build a candidate-only transformer. Prove it on a structurally faithful copy. Qualify live acquisition and Excel/Cellpy behavior next. Add production replacement last, as a separately reviewed and deliberately enabled capability.

Until then, the safe status is: **analysis and candidate-only development allowed; production update prohibited**.
