#!/usr/bin/env python3
"""
Configuration management for Cell Analysis Database Auto-Update System.

This module handles loading and validating configuration from JSON files.
"""

import json
import os
from pathlib import Path

# Configuration constants
CONFIG_FILE = "config.json"

def load_config(config_path=CONFIG_FILE):
    """
    Load and validate configuration from JSON file.
    
    Args:
        config_path (str): Path to configuration file
    
    Returns:
        dict: Configuration dictionary
    
    Raises:
        FileNotFoundError: If config file doesn't exist
        json.JSONDecodeError: If config file has invalid JSON
        ValueError: If required config fields are missing
    """
    try:
        # Security fix: Prevent path traversal attacks
        config_file = Path(config_path).resolve()
        
        # Ensure the config file is within the current working directory or its subdirectories
        cwd = Path.cwd().resolve()
        try:
            config_file.relative_to(cwd)
        except ValueError:
            raise ValueError(f"Configuration file must be within the workspace directory: {config_path}")
        
        # Also check for suspicious patterns
        if ".." in str(config_path) or config_path.startswith("/"):
            raise ValueError(f"Invalid configuration path: {config_path}")
        
        if not config_file.exists():
            raise FileNotFoundError(f"Configuration file not found: {config_path}")
        
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        # Validate required fields
        required_fields = [
            'projects', 'source_path', 'work_dir', 'db_path', 
            'sheet_to_copy', 'target_sheet', 'unique_key_col',
            'logging_format', 'dry_run', 'auto_backup'
        ]
        
        missing_fields = [field for field in required_fields if field not in config]
        if missing_fields:
            raise ValueError(f"Missing required configuration fields: {missing_fields}")
        
        # Validate paths exist
        if not Path(config['work_dir']).exists():
            raise ValueError(f"Work directory does not exist: {config['work_dir']}")
        
        if not Path(config['db_path']).exists():
            raise ValueError(f"Database file does not exist: {config['db_path']}")
        
        return config
        
    except Exception as e:
        print(f"Error loading configuration: {e}")
        raise 