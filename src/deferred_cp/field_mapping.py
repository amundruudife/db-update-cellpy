"""Gate B field dictionary and non-production golden-row renderer.

The dictionary is deliberately stricter than the legacy workbook writer.  Any
unresolved entry is a hard blocker; callers cannot accidentally turn a draft
mapping into a production row generator.
"""

from dataclasses import dataclass
from typing import Any, Mapping, Optional, Tuple


ALLOWED_OWNERSHIP = frozenset(
    {"formula", "system", "manual", "preserve_existing_blank_new", "unresolved"}
)
SOURCE_ROW_OFFSET = 1  # observed: db_table row 3 corresponds to Slurry row 4


@dataclass(frozen=True)
class FieldMapping:
    column: str
    header: str
    ownership: str
    source_field: Optional[str] = None
    formula_template: Optional[str] = None
    literal: Any = None
    source_units: Optional[str] = None
    target_units: Optional[str] = None
    conversion: Optional[str] = None
    value_type: Optional[str] = None
    null_policy: Optional[str] = None
    blocker: Optional[str] = None


def _preserve(column: str, header: str) -> FieldMapping:
    return FieldMapping(
        column, header, "preserve_existing_blank_new", value_type="cell", null_policy="blank_new"
    )


def _unresolved(column: str, header: str, reason: str, **kwargs: Any) -> FieldMapping:
    return FieldMapping(column, header, "unresolved", blocker=reason, **kwargs)


_HEADERS = (
    ("A", "id"), ("B", "batch_number"), ("C", "batch"), ("D", "exists"),
    ("E", "exists_txt"), ("F", "b01"), ("G", "b02"), ("H", "b03"),
    ("I", "b04"), ("J", "b05"), ("K", "b06"), ("L", "b07"),
    ("M", "nom_cap_specifics"), ("N", "area"), ("O", "mass_active_material"),
    ("P", "nominal_capacity"), ("Q", "loading_active_material"), ("R", "cell_type"),
    ("S", "tester"), ("T", "instrument"), ("U", "group"), ("V", "project"),
    ("W", "label"), ("X", "experiment_type"), ("Y", "selected"), ("Z", "cell"),
    ("AA", "file_name_indicator"), ("AB", "materials_used"), ("AC", "slurry_name"),
    ("AD", "comment_one"), ("AE", "comment_two"), ("AF", "comment_three"),
    ("AG", "comment_slurry"), ("AH", "finished"), ("AI", "freeze"),
    ("AJ", "boolean flag 01"), ("AK", "boolean flag 02"), ("AL", "boolean flag 03"),
    ("AM", "boolean flag 04"), ("AN", "boolean flag 05"), ("AO", "boolean flag 06"),
    ("AP", "boolean flag 07"), ("AQ", "boolean flag 08"), ("AR", "boolean flag 09"),
    ("AS", "boolean flag 10"), ("AT", "boolean flag 11"), ("AU", "boolean flag 12"),
    ("AV", "boolean flag 13"), ("AW", "channel"), ("AX", "cell_design"),
    ("AY", "separator"), ("AZ", "electrolyte"), ("BA", "mass_total"),
    ("BB", "active_material_mass_fraction"), ("BC", "pasting_thickness"),
    ("BD", "solvent_solid_ratio"), ("BE", "schedule"), ("BF", "comment_cell"),
    ("BG", "comment_general"), ("BH", "missing_raw"), ("BI", "inactive_additive_mass"),
    ("BJ", "temperature"), ("BK", "cellpy_file_name"), ("BL", "raw_file_names"),
    ("BM", "formation"), ("BN", "material field 01"), ("BO", "material field 02"),
    ("BP", "material field 03"), ("BQ", "material field 04"), ("BR", "material field 05"),
    ("BS", "material field 06"), ("BT", "material field 07"), ("BU", "db_comment"),
)


def _build_dictionary() -> Tuple[FieldMapping, ...]:
    entries = {column: _preserve(column, header) for column, header in _HEADERS}

    entries["A"] = FieldMapping("A", "id", "formula", "key", "=Slurry!A{source_row}", value_type="int", null_policy="reject")
    entries["D"] = FieldMapping("D", "exists", "system", literal=1, value_type="int", null_policy="never_null")
    for column, header in _HEADERS[5:12]:
        entries[column] = FieldMapping(column, header, "manual", value_type="cell", null_policy="preserve_or_blank")
    entries["T"] = FieldMapping("T", "instrument", "system", literal="arbin_sql_h5", value_type="str", null_policy="never_null")
    entries["X"] = FieldMapping("X", "experiment_type", "system", literal="cycling", value_type="str", null_policy="never_null")
    entries["S"] = _preserve("S", "tester")

    for column, header, source, template in (
        ("Z", "cell", "cell label", "=Slurry!C{source_row}"),
        ("AA", "file_name_indicator", "file name", "=Slurry!D{source_row}"),
        ("AW", "channel", "channel", "=Slurry!B{source_row}"),
        ("BE", "schedule", "test schedule", "=Slurry!L{source_row}"),
    ):
        entries[column] = FieldMapping(column, header, "formula", source, template, value_type="str", null_policy="formula_blank")

    entries["O"] = _unresolved(
        "O", "mass_active_material",
        "candidate c&p.E grams-to-milligrams conversion lacks final approval, type, rounding, and null policy",
        source_field="active material working electrode (g)", source_units="g", target_units="mg",
    )
    for column, header, reason in (
        ("C", "batch", "no approved c&p source or preservation reclassification"),
        ("P", "nominal_capacity", "c&p.G is assumed capacity, not an approved target mapping"),
        ("Q", "loading_active_material", "loading versus assumed capacity is unresolved"),
        ("R", "cell_type", "no approved c&p source"),
        ("U", "group", "no approved c&p source"),
        ("V", "project", "no approved c&p source"),
        ("AG", "comment_slurry", "no approved c&p source or preservation reclassification"),
    ):
        entries[column] = _unresolved(column, header, reason)
    return tuple(entries[column] for column, _ in _HEADERS)


TARGET_COLUMNS = _build_dictionary()


def mapping_blockers() -> Tuple[str, ...]:
    """Return blockers, including the observed 74-column production mismatch."""
    blockers = [f"{entry.column} {entry.header}: {entry.blocker}" for entry in TARGET_COLUMNS if entry.blocker]
    blockers.append("accessible production workbook has 74 db_table columns, while the Gate B contract defines 73")
    blockers.append("exact production headers for AJ:AV and BN:BT remain unapproved")
    return tuple(blockers)


def _by_header(entries: Mapping[str, Any]) -> Mapping[str, Any]:
    return entries


def build_db_row(source: Mapping[str, Any], existing: Mapping[str, Any], *, db_row: int = 3, allow_blocked: bool = False) -> list:
    """Render a golden row; production callers must not pass ``allow_blocked``."""
    if mapping_blockers() and not allow_blocked:
        raise ValueError("unresolved Gate B field mapping; row generation is blocked")
    source_row = db_row + SOURCE_ROW_OFFSET
    result = []
    for entry in TARGET_COLUMNS:
        if entry.ownership == "formula":
            result.append(entry.formula_template.format(source_row=source_row))
        elif entry.ownership == "system":
            result.append(entry.literal)
        elif entry.ownership in {"manual", "preserve_existing_blank_new"}:
            result.append(existing.get(entry.header))
        else:
            result.append(None)
    return result


def preserve_by_id(source_order: Tuple[int, ...], existing_rows: Mapping[int, Mapping[str, Any]]) -> Tuple[Mapping[str, Any], ...]:
    """Return existing rows in source order without positional rebinding."""
    return tuple(existing_rows[record_id] for record_id in source_order if record_id in existing_rows)
