#!/usr/bin/env python3
"""
Cell Analysis Database Auto-Update System - Main Entry Point

This is the main script that orchestrates the complete database update pipeline.
It integrates SharePoint file copying and database update operations.

Usage:
    python main.py          # Run with current config.json settings
    python main.py --dry-run # Force dry-run mode
    python main.py --live    # Force live mode (override config)

Author: Auto-generated (Refactored)
Date: 2025-01-14
Version: 2.0
"""

import sys
import argparse
from pathlib import Path

# Import our refactored modules from src/
from src.config import load_config
from src.logging_utils import setup_logging, get_logger
from src.file_operations import copy_log_sheet, backup_db
from src.data_processing import filter_by_projects, check_duplicates, prepare_update_data
from src.database import update_slurry, dry_run_full_pipeline
import src.copy_sharepoint_file as copy_sharepoint_file

def get_fresh_source_data():
    """
    Attempt to get fresh source data from SharePoint downloads.
    
    Returns:
        bool: True if successful, False if failed or skipped
    """
    logger = get_logger()
    
    try:
        logger.info("Checking for fresh SharePoint data...")
        success = copy_sharepoint_file.copy_cell_log_to_source_data()
        
        if success:
            logger.info("✅ Fresh source data obtained from SharePoint")
            return True
        else:
            logger.warning("⚠️ Could not get fresh SharePoint data - using existing source file")
            return False
            
    except Exception as e:
        logger.warning(f"SharePoint copy failed: {e}")
        logger.info("Proceeding with existing source data...")
        return False

def run_live_pipeline(config):
    """
    Execute the live database update pipeline.
    
    Args:
        config (dict): Configuration dictionary
    
    Returns:
        bool: True if successful, False otherwise
    """
    logger = get_logger()
    
    try:
        logger.info("=" * 50)
        logger.info("LIVE MODE - PRODUCTION DATABASE UPDATE")
        logger.info("=" * 50)
        
        # Prepare data using the centralized function
        prep_results = prepare_update_data(config)
        new_rows_df = prep_results['new_rows_df']
        filtered_df = prep_results['filtered_df']
        duplicate_keys = prep_results['duplicate_keys']

        # Step 4: Create backup (if enabled)
        if config['auto_backup']:
            logger.info("Step 4: Creating database backup...")
            backup_path = backup_db(config['db_path'], config['auto_backup'], config['work_dir'])
        
        # Step 5: Update database
        logger.info("Step 5: Updating database...")
        rows_added = update_slurry(
            new_rows_df,
            config['db_path'],
            config['target_sheet'],
            dry_run=False
        )
        
        # Final summary
        project_summary = new_rows_df.iloc[:, 2].value_counts().to_dict() if len(new_rows_df) > 0 else {}
        summary_text = ", ".join([f"{proj}:{count}" for proj, count in project_summary.items()])
        
        logger.info("=" * 50)
        logger.info("DATABASE UPDATE COMPLETED SUCCESSFULLY")
        logger.info(f"Total rows processed: {len(filtered_df)}")
        logger.info(f"Rows appended: {rows_added} ({summary_text})")
        logger.info(f"Duplicate rows skipped: {len(duplicate_keys)}")
        logger.info("=" * 50)
        
        return True
        
    except Exception as e:
        logger.error(f"Live pipeline failed: {e}")
        return False

def run_dry_run_pipeline(config):
    """
    Execute the dry-run database update pipeline.
    
    Args:
        config (dict): Configuration dictionary
    
    Returns:
        bool: True if successful, False otherwise
    """
    logger = get_logger()
    
    try:
        logger.info("Dry-run mode: Executing full pipeline with output to folder...")
        results = dry_run_full_pipeline(config)
        
        if results['errors']:
            logger.error("Dry run completed with errors")
            return False
        else:
            logger.info("Dry run completed successfully - no production files were modified")
            return True
            
    except Exception as e:
        logger.error(f"Dry run pipeline failed: {e}")
        return False

def parse_arguments():
    """
    Parse command line arguments.
    
    Returns:
        argparse.Namespace: Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Cell Analysis Database Auto-Update System",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py                 # Use config.json settings
  python main.py --dry-run       # Force dry-run mode
  python main.py --live          # Force live mode
  python main.py --get-sharepoint # Get fresh data from SharePoint first
        """
    )
    
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        '--dry-run', 
        action='store_true',
        help='Force dry-run mode (override config)'
    )
    mode_group.add_argument(
        '--live', 
        action='store_true',
        help='Force live mode (override config)'
    )
    
    parser.add_argument(
        '--get-sharepoint',
        action='store_true',
        help='Attempt to get fresh data from SharePoint first'
    )
    
    parser.add_argument(
        '--config',
        default='config.json',
        help='Path to configuration file (default: config.json)'
    )
    
    return parser.parse_args()

def main():
    """
    Main execution function.
    """
    # Parse command line arguments
    args = parse_arguments()
    
    try:
        # Load configuration
        config = load_config(args.config)
        
        # Setup logging
        logger = setup_logging(config['work_dir'], config['logging_format'])
        
        logger.info("=" * 60)
        logger.info("Cell Analysis Database Auto-Update System v2.0 - Starting")
        logger.info("=" * 60)
        
        # Handle mode overrides from command line
        if args.dry_run:
            config['dry_run'] = True
            logger.info("Command line override: DRY-RUN mode enabled")
        elif args.live:
            config['dry_run'] = False
            logger.info("Command line override: LIVE mode enabled")
        
        logger.info(f"Operating mode: {'DRY-RUN' if config['dry_run'] else 'LIVE'}")
        logger.info(f"Projects: {', '.join(config['projects'])}")
        
        # Get fresh SharePoint data if requested
        if args.get_sharepoint:
            fresh_data_success = get_fresh_source_data()
            if not fresh_data_success:
                logger.warning("Continuing with existing source data...")
        
        # Execute appropriate pipeline
        if config['dry_run']:
            success = run_dry_run_pipeline(config)
        else:
            success = run_live_pipeline(config)
        
        # Final status
        if success:
            logger.info("✅ Operation completed successfully")
            sys.exit(0)
        else:
            logger.error("❌ Operation failed")
            sys.exit(1)
        
    except KeyboardInterrupt:
        print("\n🛑 Operation cancelled by user")
        sys.exit(1)
        
    except Exception as e:
        if 'logger' in locals():
            logger.error(f"Fatal error: {e}")
            logger.error("System update failed")
        else:
            print(f"Fatal error during initialization: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 