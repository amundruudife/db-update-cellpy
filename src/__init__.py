"""
Cell Analysis Database Auto-Update System - Refactored Version

A modular system for automating cell analysis database updates with SharePoint integration.

Modules:
- config: Configuration management
- logging_utils: Logging setup and utilities
- file_operations: File copying, versioning, and backup operations
- data_processing: Data filtering and duplicate checking
- database: Database update operations
- copy_sharepoint_file: SharePoint file extraction utility

Version: 2.0
"""

__version__ = "2.0"
__author__ = "Open Source Community"

# Import main modules for easy access
from . import config
from . import logging_utils
from . import file_operations
from . import data_processing
from . import database
from . import copy_sharepoint_file 