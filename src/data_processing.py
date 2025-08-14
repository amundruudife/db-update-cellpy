#!/usr/bin/env python3
"""
Data processing for Cell Analysis Database Auto-Update System.

This module handles data filtering, duplicate checking, and validation.
"""

import pandas as pd
from pathlib import Path
from .logging_utils import get_logger, log_exceptions
from .file_operations import copy_log_sheet
from .common_utils import read_excel_with_cellpy_format, read_excel_raw, validate_file_path

@log_exceptions("Error filtering by projects: {e}")
def filter_by_projects(copied_file_path, sheet_name, projects, project_col_index=2):
    """
    Filter data from copied file by project list.
    
    Args:
        copied_file_path (str): Path to copied log file
        sheet_name (str): Name of sheet to read
        projects (list): List of project names to filter by
        project_col_index (int): Index of project column (default: 2 for column C)
    
    Returns:
        pd.DataFrame: Filtered data
    
    Raises:
        FileNotFoundError: If copied file doesn't exist
        ValueError: If sheet doesn't exist
    """
    logger = get_logger()
    
    # Read the Excel file - skip the first 4 header rows for cellpy format
    df = read_excel_with_cellpy_format(copied_file_path, sheet_name)
    
    logger.info(f"Read {len(df)} rows from {sheet_name} sheet (after skipping headers)")
    
    # Filter by projects (case-sensitive exact match)
    project_col_name = df.columns[project_col_index] if project_col_index < len(df.columns) else None
    
    if project_col_name is None:
        raise ValueError(f"Project column index {project_col_index} is out of range")
    
    filtered_df = df[df[project_col_name].isin(projects)]
    
    # Log filtering results
    project_counts = filtered_df[project_col_name].value_counts()
    logger.info(f"Filtered to {len(filtered_df)} rows:")
    for project, count in project_counts.items():
        logger.info(f"  {project}: {count} rows")
    
    return filtered_df

@log_exceptions("Error checking duplicates: {e}")
def check_duplicates(filtered_df, db_path, target_sheet, unique_key_col, key_col_index=0):
    """
    Check for duplicate keys between filtered data and existing database.
    
    Args:
        filtered_df (pd.DataFrame): Filtered data from source
        db_path (str): Path to database file
        target_sheet (str): Target sheet name
        unique_key_col (str): Unique key column identifier
        key_col_index (int): Index of key column (default: 0 for column A)
    
    Returns:
        tuple: (new_rows_df, duplicate_keys_list)
    
    Raises:
        FileNotFoundError: If database file doesn't exist
        ValueError: If target sheet doesn't exist
    """
    logger = get_logger()
    
    # Read existing data from database (handle complex cellpy format)
    existing_df = read_excel_raw(db_path, target_sheet)
    
    # Get existing keys from column A (index 0)
    existing_keys = set()
    if len(existing_df) > 0 and len(existing_df.columns) > key_col_index:
        # Bug fix: Don't convert to string, preserve original data types
        # This ensures accurate comparison, especially for numeric keys
        existing_keys_series = existing_df.iloc[:, key_col_index].dropna()
        existing_keys = set(existing_keys_series.tolist())
    
    logger.info(f"Found {len(existing_keys)} existing keys in database")
    
    # Get source key column name
    source_key_col_name = filtered_df.columns[key_col_index] if key_col_index < len(filtered_df.columns) else None
    
    if source_key_col_name is None:
        raise ValueError(f"Key column index {key_col_index} is out of range in source data")
    
    # Check for duplicates - preserve data types for accurate comparison
    source_keys = filtered_df[source_key_col_name]
    
    # Bug fix: Handle potential type mismatches between source and database
    # Convert both to the same type only if necessary
    duplicate_mask = pd.Series([False] * len(source_keys), index=source_keys.index)
    
    for idx, key in source_keys.items():
        if pd.notna(key):  # Skip NaN values
            # Check if key exists, handling potential type differences
            if key in existing_keys:
                duplicate_mask[idx] = True
            elif isinstance(key, (int, float)):
                # Check for numeric equivalence (e.g., 1.0 == 1)
                for existing_key in existing_keys:
                    if isinstance(existing_key, (int, float)) and key == existing_key:
                        duplicate_mask[idx] = True
                        break
    
    new_rows_df = filtered_df[~duplicate_mask].copy()
    duplicate_keys = source_keys[duplicate_mask].tolist()
    
    # Log results
    logger.info(f"Duplicate check results:")
    logger.info(f"  New rows to add: {len(new_rows_df)}")
    logger.info(f"  Duplicate rows skipped: {len(duplicate_keys)}")
    
    if duplicate_keys:
        logger.warning(f"Skipping duplicate keys: {duplicate_keys[:10]}{'...' if len(duplicate_keys) > 10 else ''}")
    
    return new_rows_df, duplicate_keys

def prepare_update_data(config):
    """
    Orchestrate the data preparation steps: copy, filter, and check duplicates.
    
    Args:
        config (dict): Configuration dictionary
    
    Returns:
        dict: A dictionary containing the results of the preparation steps.
    """
    logger = get_logger()

    # Step 1: Copy log sheet with versioning
    logger.info("Step 1: Copying source file with versioning...")
    copied_file_path = copy_log_sheet(config['source_path'], config['work_dir'])
    
    # Step 2: Filter data by projects
    logger.info("Step 2: Filtering data by projects...")
    filtered_df = filter_by_projects(
        copied_file_path, 
        config['sheet_to_copy'], 
        config['projects']
    )
    
    # Step 3: Check for duplicates
    logger.info("Step 3: Checking for duplicate keys...")
    new_rows_df, duplicate_keys = check_duplicates(
        filtered_df,
        config['db_path'],
        config['target_sheet'],
        config['unique_key_col']
    )

    return {
        'copied_file_path': copied_file_path,
        'filtered_df': filtered_df,
        'new_rows_df': new_rows_df,
        'duplicate_keys': duplicate_keys,
    } 