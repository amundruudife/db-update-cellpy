# Gate B2 Existing-Data Inventory

Status: evidence captured read-only from the hard-coded production workbook on 2026-08-05.

Source: `C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx`.
Machine-readable evidence is in `gate_b_artifacts/production_inventory.json`.

## Observed production facts

- Sheet order is `Slurry`, then `db_table`.
- `Slurry` is `A1:CI370` (370 rows, 87 columns); `db_table` is `A1:BU363` (363 rows, 73 columns).
- `db_table` has 4,332 formulas. The populated non-formula cells outside F:L are 1,083 cells, in D, T, and X (361 each).
- The observed direct mapping formula at `db_table!A3` is `=Slurry!A4`; the intended normal source-to-target row offset is source row = target row + 1.
- Offset scanning found an outlier at `db_table!A247 = Slurry!A254`, so the offset is not globally consistent and must be investigated before any production rewrite.
- Neither sheet has tables, defined names, comments, validations, external links, or macros in this package. Styles, merges, dimensions, hidden state, filters, panes, widths/heights, print settings, and calculation properties are captured in the JSON artifact.
- The baseline `db_table` headers are 73 columns A:BU. The observed AJ:AV
  headers are `F1`, `F2`, `F3`, `SP`, `F5`, `RATE`, `LC`, `A1`, `A2`, `A3`,
  `A4`, `A5`, `A6`; BN:BT are `material_class`, `material_label`,
  `material_group_label`, `material_sub_label`, `material_solvent`,
  `material_pre_processing`, and `material_surface_processing`. The generic
  fixture labels remain draft representations until the mapping implementation
  is deliberately synchronized with these observed production headers.

## Cell ownership evidence

Every populated non-formula cell outside F:L is classified `retain_existing` for existing IDs. Formula cells are separated and classified `recreate_formula`. A first-migration delete set cannot be inferred from one workbook alone; deletion is therefore conservative and empty until an explicit before/after candidate comparison is supplied. `compare_cell_inventories` emits cell-level `retain_existing`, `recreate`, and `delete` records when both inventories exist.

The inventory is read-only: it opens the workbook for inspection and closes it without saving.

## Remaining blockers

- The row-offset artifact is not globally consistent: most direct references
  use offset 1, but an offset-7 outlier exists at `db_table!A247`.
- The inventory does not approve the semantic field dictionary, first-migration
  diff, or topology allowlist; those remain Gate B approval items.
