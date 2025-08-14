#!/usr/bin/env python3
"""
File operations for Cell Analysis Database Auto-Update System.

This module handles file copying, versioning, and backup operations.
"""

import shutil
from datetime import datetime
from pathlib import Path
from .logging_utils import get_logger

# File operation constants
BACKUP_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"

def copy_log_sheet(source_path, work_dir):
    """
    Copy the source log file with proper versioning to source_data folder.
    
    Args:
        source_path (str): Path to source Cell_Log.xlsx
        work_dir (str): Working directory (used to find source_data folder)
    
    Returns:
        str: Path to the copied file
    
    Raises:
        FileNotFoundError: If source file doesn't exist
        PermissionError: If unable to copy file
    """
    logger = get_logger()
    
    try:
        source_file = Path(source_path)
        # Save versioned copies to source_data/ instead of root
        source_data_dir = Path(work_dir) / "source_data"
        source_data_dir.mkdir(exist_ok=True)
        
        if not source_file.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")
        
        # Generate versioned filename: YYMMDD_cellog_n.xlsx
        today = datetime.now()
        date_str = today.strftime("%y%m%d")  # YYMMDD format
        
        # Find next available version number
        version = 1
        while True:
            filename = f"{date_str}_cellog_{version}.xlsx"
            dest_path = source_data_dir / filename
            if not dest_path.exists():
                break
            version += 1
        
        # Copy the file
        shutil.copy2(source_file, dest_path)
        
        logger.info(f"Source file copied to source_data/: {filename}")
        return str(dest_path)
        
    except Exception as e:
        logger.error(f"Error copying log sheet: {e}")
        raise

def backup_db(db_path, auto_backup=True, work_dir=None):
    """
    Create a timestamped backup of the database file in output/ folder.
    
    Args:
        db_path (str): Path to database file
        auto_backup (bool): Whether to create backup
        work_dir (str): Working directory (for output/ folder location)
    
    Returns:
        str: Path to backup file, or None if no backup created
    
    Raises:
        PermissionError: If unable to create backup
    """
    logger = get_logger()
    
    if not auto_backup:
        logger.info("Auto-backup disabled")
        return None
        
    try:
        db_file = Path(db_path)
        if not db_file.exists():
            raise FileNotFoundError(f"Database file not found: {db_path}")
        
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime(BACKUP_TIMESTAMP_FORMAT)
        backup_name = f"backup_{db_file.stem}_{timestamp}.bak{db_file.suffix}"
        
        # Save to output/ folder if work_dir provided, otherwise use current directory
        if work_dir:
            output_dir = Path(work_dir) / "output"
            output_dir.mkdir(exist_ok=True)
            backup_path = output_dir / backup_name
        else:
            backup_path = Path.cwd() / "output" / backup_name
            backup_path.parent.mkdir(exist_ok=True)
        
        # Create backup
        shutil.copy2(db_file, backup_path)
        
        logger.info(f"Database backup created: {backup_path}")
        return str(backup_path)
        
    except Exception as e:
        logger.error(f"Error creating backup: {e}")
        raise 