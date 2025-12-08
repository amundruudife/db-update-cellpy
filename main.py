#!/usr/bin/env python3
"""
Minimal pipeline: copy fresh Cell_Log.xlsx from Downloads, filter by project,
drop duplicates already in Slurry, and append the new rows.
"""

import sys
import argparse
from pathlib import Path

from src.config import load_config
from src.logging_utils import setup_logging, get_logger
from src.data_processing import filter_by_projects, check_duplicates
from src.database import update_slurry
import src.copy_sharepoint_file as copy_sharepoint_file


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Copy relevant rows from Cell_Log.xlsx into Slurry"
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to configuration file (default: config.json)",
    )
    parser.add_argument(
        "--skip-sharepoint",
        action="store_true",
        help="Skip fetching newest Cell_Log from Downloads",
    )
    return parser.parse_args()


def copy_source_file(skip_sharepoint: bool):
    logger = get_logger()
    if skip_sharepoint:
        logger.info("Skipping SharePoint fetch (per flag)")
        return True

    logger.info("Fetching latest Cell_Log.xlsx from Downloads...")
    success = copy_sharepoint_file.copy_cell_log_to_source_data()
    if not success:
        logger.error("Could not copy Cell_Log.xlsx from Downloads; aborting.")
    return success


def main():
    args = parse_arguments()
    try:
        config = load_config(args.config)
        logger = setup_logging(config["work_dir"], config["logging_format"])
        logger.info("=" * 60)
        logger.info("Starting minimal Cell_Log -> Slurry update")
        logger.info(f"Projects: {', '.join(config['projects'])}")

        # Step 0: ensure fresh source copy
        if not copy_source_file(args.skip_sharepoint):
            sys.exit(1)

        # Step 1: filter source rows by project
        if not Path(config["source_path"]).exists():
            logger.error(f"Source file missing after copy: {config['source_path']}")
            sys.exit(1)

        filtered_df = filter_by_projects(
            config["source_path"],
            config["sheet_to_copy"],
            config["projects"],
        )
        logger.info(f"Filtered rows: {len(filtered_df)}")

        # Step 2: drop duplicates already in Slurry (by column A)
        new_rows_df, duplicate_keys = check_duplicates(
            filtered_df,
            config["db_path"],
            config["target_sheet"],
            config["unique_key_col"],
        )
        logger.info(f"Duplicate keys skipped: {len(duplicate_keys)}")

        if len(new_rows_df) == 0:
            logger.info("No new rows to append; done.")
            sys.exit(0)

        # Step 3: append to Slurry
        rows_added = update_slurry(
            new_rows_df,
            config["db_path"],
            config["target_sheet"],
            dry_run=False,
        )
        logger.info(f"Appended rows: {rows_added}")
        logger.info("✅ Completed.")
        sys.exit(0)

    except Exception as e:
        if "logger" in locals():
            logger.error(f"Fatal error: {e}")
        else:
            print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()