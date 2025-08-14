#!/usr/bin/env python3
"""
Cell Analysis Database Auto-Update System - Package Initialization

This package provides automated SharePoint integration and database updates
for cell analysis data management.

Main modules:
- config: Configuration management
- logging_utils: Logging setup and utilities
- file_operations: File copying, versioning, and backup
- data_processing: Data filtering and duplicate checking
- database: Database update operations
- copy_sharepoint_file: SharePoint file extraction
- common_utils: Common utility functions
"""

# Version info
__version__ = "2.0.0"
__author__ = "Auto-generated (Refactored)"
__date__ = "2025-01-14"

# Import main functions for easier access
from .config import load_config
from .logging_utils import setup_logging, get_logger
from .database import update_slurry, dry_run_full_pipeline
from .common_utils import ensure_directory, generate_project_summary 