# Source Contract: Cell Log `c&p`

Status: source rules approved; Gate A remains blocked on live-source and anomaly evidence
Authority: `../../PROJECT_SCOPE.md`
Last revised: 2026-08-04

## Purpose

Define the only permitted source and the exact values-only local mirror used by the database updater. The source is never a write target.

## Fixed identities

| Item | Contract |
|---|---|
| SharePoint identity | host `ifecloud.sharepoint.com`; site path `sites/UsersofIFEBatteryLab`; default-drive item `General/00_Logs/Cell_Log.xlsx` |
| Workbook | `Cell_Log.xlsx` |
| Source sheet | `c&p` (case-sensitive) |
| Local snapshot | `source_data/Cell_Log_CP.xlsx` |
| Snapshot sheet | `c&p` |
| Source key | Column A, header `key` |
| Target key | `db_table` column A, header `id` |
| Production database | `C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx` |
| Machine-owned mirror | `Slurry` |

None of these values may be overridden at runtime.

## Observed source shape

The inspected local workbook is provisional evidence until checked against the live drive item.

| Property | Observed value |
|---|---|
| Sheet dimensions | 4,590 rows × 19 columns (A:S) |
| Header row | 2 |
| Units row | 3 |
| First data/formula row | 4 |
| Business columns | A:P |
| Helper/range columns | Q:S, currently unlabeled |
| Positive numeric key rows | 4,582 |
| Duplicated positive IDs | 17 IDs |
| Zero keys | 2 |
| Formula-error values | Present |
| Formula dependencies | Other workbook sheets, including `log` and `lookup` |

## Source columns

| Column | Header | Use |
|---|---|---|
| A | key | Stable record identity; must be a positive integer and unique |
| B | channel | Mapping candidate; target semantics unresolved |
| C | cell label | Mapping candidate for `cell` |
| D | file name | Mapping candidate for `file_name_indicator` |
| E | active material working electrode (g) | Mapping candidate for `mass_active_material` after approved unit conversion |
| F | active material counter electrode (g) | No approved target yet |
| G | assumed capacity working electrode (Ah/g) | No approved target yet; must not be mislabeled as loading |
| H | assumed capacity counter electrode (Ah/g) | No approved target yet |
| I | max C rate | No approved target yet |
| J | counter electrode mode | No approved target yet |
| K | max current (mA) | No approved target yet |
| L | test schedule | Mapping candidate for `schedule` |
| M | R1 | No approved target yet |
| N | R2/HR | No approved target yet |
| O | R3/MR | No approved target yet |
| P | R4/LR | No approved target yet |
| Q:S | unlabeled helper/range values | Mirror inclusion pending; never mapped without explicit approval |

## Approved mirror rule

- Copy evaluated values for rows 1 through the reported used-range last row and columns A:S.
- Preserve source order and typed values.
- Copy no formulas, links, macros, credentials, or other sheets.
- Treat A:P as the only business mapping surface.
- Preserve Q:S in the mirror for fidelity, but prohibit database mappings from them.
- Preserve non-key Excel error values in the exact mirror.
- Reject the snapshot if an Excel error appears in column A or in any column used by an approved database mapping.

This proposal separates exact source mirroring from the stricter database transformation contract.

## Approved identity and population rules

- Accepted IDs are every unique positive integer in `c&p`; v1 applies no project/type filter.
- IDs must be unique. The updater never selects first, selects last, or merges duplicate records.
- Blank, zero, nonnumeric, and Excel-error IDs are invalid.
- Source order is preserved in `Slurry`; `db_table` uses the same accepted-ID order.
- Manual and preservation-only database values are joined by ID, never by row.
- An accepted historical ID disappearing from a later source manifest is an abort condition.
- An existing database ID missing from the first accepted manifest is an abort condition until explicitly reconciled; this currently blocks ID `8206`.

The legacy project filter is not retained because the approved source surface `c&p` contains no authoritative project or experiment-type field. Introducing a filter later requires an explicit source field and a code-and-test contract change.

## Allowed successive-manifest changes

- `eTag`, last-modified timestamp, and content hash may change when the source
  workbook changes.
- New accepted IDs may be appended after the historical ID sequence; existing
  IDs must retain their source order.
- The drive item, workbook name, sheet name, and A:S range shape are fixed. A
  shrinking, malformed, or differently shaped range is an abort condition.
- Duplicate, invalid, removed, or reordered IDs remain abort conditions.

## Acquisition boundary

- The updater requests only the hard-coded workbook and `c&p` range.
- The Excel service may still evaluate dependencies in other sheets; “request one sheet” does not mean the service reads only one sheet internally.
- No source update request is permitted in code.
- Authentication or retrieval failure aborts. Downloads and cached-source fallbacks are prohibited.
- Each retrieval records immutable drive item identity, eTag, last-modified timestamp, used-range address, retrieval time, and content hash.
- Microsoft Graph delegated `Sites.Read.All` access is required. Resolution performs the exact site lookup followed by the default-drive item lookup with authorization headers; sharing-link `/shares` fallback is prohibited. Operationally revoke the superseded old sharing link before live acceptance.

## Snapshot acceptance

A new snapshot is accepted only when:

- [ ] live drive item identity matches the approved identity;
- [ ] sheet name, dimensions, headers, units, and used range match this contract;
- [ ] all accepted IDs satisfy the approved identity/population rules;
- [ ] mapping-critical columns contain no Excel errors;
- [ ] source manifest changes satisfy the append-only policy;
- [ ] the temporary snapshot reopens successfully and matches its manifest; and
- [ ] ordered replacement of the prior local snapshot succeeds (restart-recoverable; not a multi-file/power-loss atomic claim).

Failure preserves the last accepted snapshot and leaves production unchanged.

## Required Gate A evidence

- [ ] Live source identity record.
- [ ] Approved Graph or upstream-export permission decision.
- [ ] Duplicate/invalid-ID anomaly ledger.
- [ ] Reconciliation record for database ID `8206`.
- [ ] Approved first-run ID manifest and row count.
- [ ] Live-versus-Downloads comparison report.
- [ ] Source non-mutation evidence.
