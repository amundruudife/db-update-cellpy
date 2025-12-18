#!/usr/bin/env python3
"""
Minimal pipeline: copy fresh Cell_Log.xlsx from Downloads, filter by project,
drop duplicates already in Slurry, and append the new rows.
"""

import sys
import argparse
import shutil
from pathlib import Path

from src.config import load_config
from src.logging_utils import setup_logging, get_logger
from src.database import dry_run_full_pipeline
from src.file_operations import backup_db
import src.copy_sharepoint_file as copy_sharepoint_file
from src.cleanup_old_files import (
    cleanup_old_output_files,
    cleanup_python_cache,
    cleanup_source_data_copies,
)


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
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--stage-only",
        action="store_true",
        help="Stage changes to output/ only (default behavior)",
    )
    mode.add_argument(
        "--apply",
        action="store_true",
        help="Apply staged output to production database after staging",
    )
    parser.add_argument(
        "--maintenance",
        action="store_true",
        help="Run cleanup tasks (output, source copies, caches) and exit",
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


def run_maintenance():
    """Execute workspace cleanup tasks."""
    logger = get_logger()
    logger.info("Starting maintenance: cleaning output, source copies, and caches")
    cleanup_old_output_files()
    cleanup_source_data_copies()
    cleanup_python_cache()
    logger.info("Maintenance complete")


def main():
    args = parse_arguments()
    stage_only = args.stage_only or not args.apply
    try:
        config = load_config(args.config)
        logger = setup_logging(config["work_dir"], config["logging_format"])
        logger.info("=" * 60)
        logger.info("Starting Cell_Log -> Slurry update (staged workflow)")
        logger.info(f"Projects: {', '.join(config['projects'])}")

        if args.maintenance:
            run_maintenance()
            sys.exit(0)

        # Step 0: ensure fresh source copy
        if not copy_source_file(args.skip_sharepoint):
            sys.exit(1)

        # Stage pipeline to output/
        results = dry_run_full_pipeline(config)

        logger.info(
            "Row counts - source: %s, filtered: %s, duplicates: %s, new: %s | source max key: %s",
            results.get("source_rows", 0),
            results.get("filtered_rows", 0),
            results.get("duplicate_rows", 0),
            results.get("appended_rows", 0),
            results.get("source_max_key"),
        )

        staged_name = results.get("output_database")
        staged_path = Path(config["work_dir"]) / "output" / staged_name if staged_name else None

        if stage_only:
            logger.info("Stage-only mode: production database unchanged.")
            if staged_path:
                logger.info(f"Staged file ready at: {staged_path}")
            sys.exit(0)

        # Apply staged output to production database
        if not staged_path or not staged_path.exists():
            logger.error("Staged output database not found; cannot apply.")
            sys.exit(1)

        logger.info("Applying staged output to production database...")
        if config.get("auto_backup", False):
            backup_db(config["db_path"], auto_backup=True, work_dir=config["work_dir"])

        shutil.copy2(staged_path, config["db_path"])
        logger.info(f"✅ Applied staged output to production database from {staged_path.name}")
        sys.exit(0)

    except Exception as e:
        if "logger" in locals():
            logger.error(f"Fatal error: {e}")
        else:
            print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()