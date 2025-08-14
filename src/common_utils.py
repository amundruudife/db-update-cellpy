#!/usr/bin/env python3
"""
Common utilities for Cell Analysis Database Auto-Update System.

This module contains shared utility functions to reduce code duplication.
"""

from pathlib import Path
import pandas as pd
from .logging_utils import get_logger

def ensure_directory(directory_path):
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        directory_path (str or Path): Path to the directory
    
    Returns:
        Path: The directory path object
    """
    dir_path = Path(directory_path)
    dir_path.mkdir(parents=True, exist_ok=True)
    return dir_path

def read_excel_with_cellpy_format(file_path, sheet_name, header_rows_to_skip=4):
    """
    Read Excel file with cellpy format (skipping header rows).
    
    Args:
        file_path (str or Path): Path to the Excel file
        sheet_name (str): Name of the sheet to read
        header_rows_to_skip (int): Number of header rows to skip (default: 4)
    
    Returns:
        pd.DataFrame: The data from the Excel file
    
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If sheet doesn't exist
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    return pd.read_excel(file_path, sheet_name=sheet_name, skiprows=header_rows_to_skip)

def read_excel_raw(file_path, sheet_name, header_rows_to_skip=4):
    """
    Read Excel file without headers (for database format).
    
    Args:
        file_path (str or Path): Path to the Excel file
        sheet_name (str): Name of the sheet to read
        header_rows_to_skip (int): Number of header rows to skip (default: 4)
    
    Returns:
        pd.DataFrame: The data from the Excel file
    
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If sheet doesn't exist
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    return pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=header_rows_to_skip)

def generate_project_summary(df, project_col_index=2):
    """
    Generate a summary of project counts from a DataFrame.
    
    Args:
        df (pd.DataFrame): DataFrame containing the data
        project_col_index (int): Column index containing project names (default: 2)
    
    Returns:
        tuple: (project_dict, summary_text)
            - project_dict: Dictionary of project counts
            - summary_text: Formatted string of project counts
    """
    if len(df) == 0:
        return {}, ""
    
    project_summary = df.iloc[:, project_col_index].value_counts().to_dict()
    summary_text = ", ".join([f"{proj}:{count}" for proj, count in project_summary.items()])
    
    return project_summary, summary_text

def validate_file_path(file_path, must_exist=True):
    """
    Validate a file path and optionally check if it exists.
    
    Args:
        file_path (str or Path): Path to validate
        must_exist (bool): Whether the file must exist (default: True)
    
    Returns:
        Path: The validated path object
    
    Raises:
        FileNotFoundError: If must_exist=True and file doesn't exist
    """
    path = Path(file_path)
    if must_exist and not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return path