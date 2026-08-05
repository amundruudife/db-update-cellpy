"""Active candidate safety contracts."""

from pathlib import Path


PRODUCTION_DATABASE_PATH = Path(
    r"C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx"
)
MIRROR_SHEET_NAME = "Slurry"
DATABASE_SHEET_NAME = "db_table"
SLURRY_DATA_START_ROW = 4

SYSTEM_VALUES = {
    "exists": 1,
    "instrument": "arbin_sql_h5",
    "experiment_type": "cycling",
}


def is_production_database(path) -> bool:
    """Return True only for the exact hard-coded production workbook."""

    candidate = Path(path).resolve(strict=False)
    production = PRODUCTION_DATABASE_PATH.resolve(strict=False)
    return candidate == production
