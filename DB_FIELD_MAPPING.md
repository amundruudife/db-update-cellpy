# Database Field Mapping Contract

Status: executable draft; Gate B remains blocked by the blockers below  
Target: `C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx`, sheet `db_table`  
Mirror: sheet `Slurry`, values-only `c&p`  
Last revised: 2026-08-05

The executable dictionary is in `src/field_mapping.py`; exact-layout fixtures
and golden expectations are in `tests/fixtures/gate_b/`. `build_db_row` refuses
to generate a row while a blocker remains. Its `allow_blocked=True` option is
for fixture inspection only and is not production enablement.

## Ownership classes

- `formula`: Excel formula generated from a reviewed template, normally referencing `Slurry`.
- `system`: deterministic value or rule owned by the updater.
- `manual`: user input; preserved by ID for existing records and blank for new records.
- `preserve_existing_blank_new`: preserved by ID for existing records and blank for new records.
- `unresolved`: implementation blocker; no row generator may be enabled while present.

Manual and preservation-only values are never inferred from neighboring rows.

## Formula alignment

Proposed alignment, pending source-range approval:

- `Slurry` retains source rows 1–3.
- First source data row is `Slurry!4`.
- First `db_table` data row is row 3.
- Formula templates therefore use `source_row = db_row + 1`.
- `db_table` is rebuilt in the same accepted-ID order as `Slurry`.

Golden templates currently encoded in code and fixtures are `A =Slurry!A{source_row}`, `Z =Slurry!C{source_row}`, `AA =Slurry!D{source_row}`, `AW =Slurry!B{source_row}`, and `BE =Slurry!L{source_row}`. Source values are text except `A`, which is a positive integer key; a blank source value yields Excel's blank formula result. Final source-range and Cellpy approval remain blockers.

## Complete 73-column inventory

| Col | Target header | Current behavior | Proposed ownership | Proposed rule / blocker |
|---|---|---|---|---|
| A | id | Formula | formula | `Slurry.A key`; direct reference |
| B | batch_number | Literal/blank | preserve_existing_blank_new | No approved source |
| C | batch | Formula | unresolved | Legacy source no longer exists; approve new rule or preservation |
| D | exists | System literal | system | Approved constant `1` |
| E | exists_txt | Literal/blank | preserve_existing_blank_new | No approved source |
| F | b01 | Manual | manual | Preserve by ID; blank for new ID |
| G | b02 | Manual | manual | Preserve by ID; blank for new ID |
| H | b03 | Manual | manual | Preserve by ID; blank for new ID |
| I | b04 | Manual | manual | Preserve by ID; blank for new ID |
| J | b05 | Manual | manual | Preserve by ID; blank for new ID |
| K | b06 | Manual | manual | Preserve by ID; blank for new ID |
| L | b07 | Manual | manual | Preserve by ID; blank for new ID |
| M | nom_cap_specifics | Literal/blank | preserve_existing_blank_new | No approved source |
| N | area | Literal/blank | preserve_existing_blank_new | No approved source |
| O | mass_active_material | Formula | unresolved | Candidate `c&p.E × 1000` (g to mg) lacks approved type, rounding, and null policy |
| P | nominal_capacity | Formula | unresolved | Derivation must be defined and unit-checked |
| Q | loading_active_material | Formula | unresolved | Must not copy assumed capacity solely because values resemble legacy data |
| R | cell_type | Formula | unresolved | No approved `c&p` source |
| S | tester | Formula | preserve_existing_blank_new | Preserve existing tester by ID; blank for new IDs unless a later Cellpy rule is approved |
| T | instrument | System literal | system | Approved constant `arbin_sql_h5` |
| U | group | Formula | unresolved | No approved `c&p` source |
| V | project | Formula | unresolved | No approved `c&p` source |
| W | label | Literal/blank | preserve_existing_blank_new | No approved source |
| X | experiment_type | System literal | system | Approved constant `cycling` |
| Y | selected | Literal/blank | preserve_existing_blank_new | No approved source |
| Z | cell | Formula | formula | `Slurry.C cell label`; direct reference |
| AA | file_name_indicator | Formula | formula | `Slurry.D file name`; direct reference |
| AB | materials_used | Literal/blank | preserve_existing_blank_new | No approved source |
| AC | slurry_name | Literal/blank | preserve_existing_blank_new | No approved source |
| AD | comment_one | Literal/blank | preserve_existing_blank_new | No approved source |
| AE | comment_two | Literal/blank | preserve_existing_blank_new | No approved source |
| AF | comment_three | Literal/blank | preserve_existing_blank_new | No approved source |
| AG | comment_slurry | Formula | unresolved | Legacy source no longer exists; approve new rule or preservation |
| AH | finished | Literal/blank | preserve_existing_blank_new | No approved source |
| AI | freeze | Literal/blank | preserve_existing_blank_new | No approved source |
| AJ | boolean flag 01 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AK | boolean flag 02 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AL | boolean flag 03 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AM | boolean flag 04 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AN | boolean flag 05 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AO | boolean flag 06 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AP | boolean flag 07 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AQ | boolean flag 08 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AR | boolean flag 09 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AS | boolean flag 10 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AT | boolean flag 11 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AU | boolean flag 12 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AV | boolean flag 13 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| AW | channel | Literal/blank | formula | `Slurry.B channel`; direct reference |
| AX | cell_design | Literal/blank | preserve_existing_blank_new | No approved source |
| AY | separator | Literal/blank | preserve_existing_blank_new | No approved source |
| AZ | electrolyte | Literal/blank | preserve_existing_blank_new | No approved source |
| BA | mass_total | Literal/blank | preserve_existing_blank_new | No approved source |
| BB | active_material_mass_fraction | Literal/blank | preserve_existing_blank_new | No approved source |
| BC | pasting_thickness | Literal/blank | preserve_existing_blank_new | No approved source |
| BD | solvent_solid_ratio | Literal/blank | preserve_existing_blank_new | No approved source |
| BE | schedule | Literal/blank | formula | `Slurry.L test schedule`; direct reference, subject to representative Cellpy verification |
| BF | comment_cell | Literal/blank | preserve_existing_blank_new | No approved source |
| BG | comment_general | Literal/blank | preserve_existing_blank_new | No approved source |
| BH | missing_raw | Literal/blank | preserve_existing_blank_new | No approved source |
| BI | inactive_additive_mass | Literal/blank | preserve_existing_blank_new | No approved source |
| BJ | temperature | Literal/blank | preserve_existing_blank_new | No approved source |
| BK | cellpy_file_name | Literal/blank | preserve_existing_blank_new | No approved source |
| BL | raw_file_names | Literal/blank | preserve_existing_blank_new | No approved source |
| BM | formation | Literal/blank | preserve_existing_blank_new | No approved source |
| BN | material field 01 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BO | material field 02 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BP | material field 03 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BQ | material field 04 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BR | material field 05 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BS | material field 06 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BT | material field 07 — exact header pending inventory | Literal/blank | preserve_existing_blank_new | Exact production header must be recorded |
| BU | db_comment | Literal/blank | preserve_existing_blank_new | No approved source |

## Blocking mapping decisions

- [x] Exact headers for AJ:AV and BN:BT are recorded from the read-only production inventory; semantic ownership approval remains separate.
- [x] Golden formula templates are recorded for A, Z, AA, AW, and BE.
- [ ] Confirm formula alignment/source range and representative Cellpy behavior.
- [ ] Ownership and formulas are resolved for C, O, P, Q, R, U, V, and AG.
- [x] AW `channel` receives `Slurry.B`; S `tester` is preserved for existing IDs and blank for new IDs.
- [x] BE `schedule` receives `Slurry.L`, subject to representative Cellpy verification.
- [x] D `exists = 1`, T `instrument = arbin_sql_h5`, and X `experiment_type = cycling` are approved system rules.
- [ ] Existing values in every preservation-only column are inventoried and included in the first-migration diff.
- [ ] Every numeric formula documents source units, target units, conversion, type, rounding, and null behavior.
- [ ] No `unresolved` or “header pending” entry remains.

The retained Gate B production inventory confirms the hard-coded baseline is
`db_table!A1:BU363` with 73 columns. Earlier 74-column observations came from a
non-baseline dry-run workbook and must not be used for mapping. The baseline
also records a row-offset outlier and legacy formulas; those observations do
not approve reuse of the old mapping. Exact baseline headers for AJ:AV and
BN:BT are captured in the read-only workbook evidence, but unresolved semantic
ownership and formula approvals still block acceptance.

## Required verification

- [ ] Golden fixture covers all 73 columns.
- [ ] Manual F:L cells contain no formulas and survive row reordering by ID.
- [ ] Preservation-only values survive by ID; new IDs are blank.
- [ ] Formula patterns match exact approved templates.
- [ ] Formula cached values contain no mapping-critical Excel errors.
- [ ] Representative values are read through Cellpy with the intended semantics.
