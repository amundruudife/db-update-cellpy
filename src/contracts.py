"""Hard-coded identities that define the updater's safety boundary."""

from pathlib import Path


SHAREPOINT_WORKBOOK_URL = (
    "https://ifecloud.sharepoint.com/:x:/r/sites/UsersofIFEBatteryLab/"
    "_layouts/15/Doc.aspx?"
    "sourcedoc=%7BEED439B5-B14D-42A0-B992-AE5F08CC1F02%7D"
    "&file=Cell_Log.xlsx&action=default&mobileredirect=true"
    "&DefaultItemOpen=1"
)
SOURCE_WORKBOOK_URL = SHAREPOINT_WORKBOOK_URL
SOURCE_WORKBOOK_NAME = "Cell_Log.xlsx"
SOURCE_SHEET_NAME = "c&p"
SNAPSHOT_SHEET_NAME = SOURCE_SHEET_NAME
LOCAL_SNAPSHOT_PATH = Path("source_data") / "Cell_Log_CP.xlsx"
PRODUCTION_DATABASE_PATH = Path(
    r"C:\Users\IFE13213\cellpy_data\db\2025_Cell_Analysis_db_001.xlsx"
)
MIRROR_SHEET_NAME = "Slurry"
DATABASE_SHEET_NAME = "db_table"

# The source mirror is an evaluated-values copy of the fixed A:S range.  The
# first three rows are metadata; source records begin at row 4.
SOURCE_FIRST_COLUMN = 1
SOURCE_LAST_COLUMN = 19
SOURCE_COLUMN_COUNT = SOURCE_LAST_COLUMN - SOURCE_FIRST_COLUMN + 1
SOURCE_RANGE = "A:S"
SOURCE_RANGE_ADDRESS = SOURCE_RANGE
BUSINESS_FIRST_COLUMN = 1
BUSINESS_LAST_COLUMN = 16
BUSINESS_COLUMN_COUNT = BUSINESS_LAST_COLUMN - BUSINESS_FIRST_COLUMN + 1
BUSINESS_RANGE = "A:P"
BUSINESS_BOUNDARY = BUSINESS_RANGE
METADATA_ROWS = (1, 2, 3)
SOURCE_METADATA_ROWS = METADATA_ROWS
HEADER_ROW = 2
SOURCE_HEADER_ROW = HEADER_ROW
UNITS_ROW = 3
SOURCE_UNITS_ROW = UNITS_ROW
DATA_START_ROW = 4
SOURCE_DATA_START_ROW = DATA_START_ROW
KEY_COLUMN = 1
SOURCE_KEY_COLUMN = "A"
SOURCE_KEY_COLUMN_LETTER = "A"
KEY_HEADER = "key"
DATABASE_KEY_COLUMN = 1
DATABASE_KEY_HEADER = "id"
PRODUCTION_SHEET_NAMES = (MIRROR_SHEET_NAME, DATABASE_SHEET_NAME)

SYSTEM_VALUES = {
    "exists": 1,
    "instrument": "arbin_sql_h5",
    "experiment_type": "cycling",
}
APPROVED_SYSTEM_VALUES = SYSTEM_VALUES


def is_production_database(path) -> bool:
    """Return True only for the exact hard-coded production workbook."""

    candidate = Path(path).resolve(strict=False)
    production = PRODUCTION_DATABASE_PATH.resolve(strict=False)
    return candidate == production
