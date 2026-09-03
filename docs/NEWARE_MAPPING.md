# Neware `test_log` candidate mapping

This is a candidate-only input contract. It does not enable production writes
and does not create a Neware sheet or a legacy `Slurry` row.

## Source contract

The reader opens only the `test_log` sheet. Row 1 contains headers, row 2 is
the Neware type-description row, and records start at row 3. Other sheets and
all unlisted columns are ignored.

The two template rows whose `Cell ID` is exactly `Cell ID or Cell serial
number` (the supplied workbook includes a leading space) are excluded and
reported by source row number. The remaining rows must have a nonblank,
unique `Cell test label`. `Test No` is retained only in the manifest payload
for audit continuity; it is not the identity key and may repeat.

## Mapping

| Neware `test_log` field | `db_table` field | Column |
| --- | --- | --- |
| Cell ID | label | W |
| Cell batch | batch | C |
| Project | project | V |
| Cell type | cell_type | R |
| Cell test label | cell | Z |
| Cell capacity (Ah) | nominal_capacity | P |
| Test Schedule | schedule | BE |
| Test temp (°C) | temperature | BJ |
| Comment | comment_general | BG |

Values are copied directly with their workbook types. Blank cells become
blank database cells. `exists=1`, `instrument=arbin_sql_h5`, and
`experiment_type=cycling` retain the existing system-owned values. All other
Neware columns remain blank in a new row; no legacy formulas are used.

## IDs and repeat runs

The default manifest is `source_data/neware_id_manifest.json`; the CLI accepts
`--neware-manifest` to select another path. Entries are keyed by exact
`Cell test label` and contain the assigned positive integer ID plus the full
approved source payload, including the audit-only `Test No`.

New IDs are allocated above every existing `Slurry` ID, database ID, source ID,
and prior manifest assignment. Existing assignments are reused. The build
fails closed
on duplicate labels, manifest ID collisions, changed existing payloads,
changed mapped database values, missing/orphan/duplicate/malformed IDs, or an
ID that cannot be reconciled with `Slurry` or the validated manifest.
Candidate, JSON report, and changed manifest are staged first and published in
order (manifest, candidate, report); this is restart-recoverable, not a
same-transaction or multi-file/power-loss atomic claim. The source/database
inputs are never modified.

`--cellpy-ready` cannot be combined with `--neware-source`; a separate reviewed
formula/Cellpy mapping is required before that boundary changes.
