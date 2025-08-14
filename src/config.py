#!/usr/bin/env python3
"""
Configuration management for Cell Analysis Database Auto-Update System.

This module handles loading and validating configuration from JSON files.
"""

import json
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
        config_file = Path(config_path)
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
        
        # Always auto-fetch fresh SharePoint data (restore old behavior)
        print("Fetching fresh data from SharePoint...")
        
        try:
            # Import here to avoid circular imports
            from . import copy_sharepoint_file
            success = copy_sharepoint_file.copy_cell_log_to_source_data()
            
            if success:
                print("✅ Successfully fetched fresh data from SharePoint")
            else:
                # If SharePoint fetch fails, check if we have an existing file to fall back to
                if not Path(config['source_path']).exists():
                    raise ValueError(f"SharePoint auto-fetch failed and no existing source file found: {config['source_path']}")
                else:
                    print("⚠️ SharePoint fetch failed, using existing source file")
        except Exception as e:
            # If SharePoint fetch fails, check if we have an existing file to fall back to
            if not Path(config['source_path']).exists():
                raise ValueError(f"SharePoint auto-fetch failed and no existing source file found: {config['source_path']} (Error: {e})")
            else:
                print(f"⚠️ SharePoint fetch failed, using existing source file (Error: {e})")
        
        if not Path(config['db_path']).exists():
            raise ValueError(f"Database file does not exist: {config['db_path']}")
        
        return config
        
    except Exception as e:
        print(f"Error loading configuration: {e}")
        raise 