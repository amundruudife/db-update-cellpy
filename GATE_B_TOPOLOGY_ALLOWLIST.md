# Gate B2 Workbook Topology Allowlist

The baseline is `gate_b_artifacts/production_inventory.json`. Candidate package changes are allowed only when they are explicitly listed here.

The baseline evidence is read-only and records `Slurry` followed by `db_table`,
dimensions `A1:CI370` and `A1:BU363`, respectively. The allowlist is not an
approval to write production and remains subject to Gate B review.

## Allowed

1. Replace evaluated data rows in `Slurry` while retaining its required sheet name and workbook package structure.
2. Rebuild `db_table` data rows for the approved ID set, including approved formula strings and updater-owned system values.
3. Preserve existing non-formula F:L values by ID and preserve existing non-formula values outside F:L by ID unless a later approved mapping changes their ownership.
4. Change `db_table` row count only as required by the approved source ID set; row order follows the approved source order.
5. Update calculation metadata only when required to force the approved desktop Excel recalculation, and record the exact change.

## Prohibited without a new reviewed decision

Changes to sheet order/names, workbook dimensions outside the approved data rectangles, styles, tables, defined names, data validations, comments, merged cells, hidden sheets/rows/columns, filters, panes, column widths, row heights, print settings, external links, VBA/macros, or unrelated package parts. Adding a table, link, macro, name, validation, comment, merge, hidden state, or filter is not allowed.

Any observed baseline topology difference fails the comparison rather than being silently normalized. The allowlist is intentionally conservative because the production workbook is also a Cellpy-facing artifact.

The baseline also contains a formula row-offset outlier. Formula alignment and
calculated-value approval therefore remain blockers even though the package
topology itself has been inventoried.
