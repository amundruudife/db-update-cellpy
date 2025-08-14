#!/usr/bin/env python3
"""
Logging utilities for Cell Analysis Database Auto-Update System.

This module handles logging setup and configuration.
"""

import logging
from pathlib import Path
from functools import wraps

# Logging constants
LOG_FILE = "update_log.txt"
TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M:%S"

# Global logger
logger = None

def setup_logging(work_dir, log_format):
    """
    Setup logging configuration for both console and file output.
    
    Args:
        work_dir (str): Working directory path
        log_format (str): Log message format template
    
    Returns:
        logging.Logger: Configured logger instance
    """
    global logger
    logger = logging.getLogger('db_updater')
    logger.setLevel(logging.INFO)
    
    # Clear any existing handlers
    logger.handlers = []
    
    # Create formatters
    log_format_final = log_format.replace('{timestamp}', '%(asctime)s').replace('{message}', '%(message)s')
    formatter = logging.Formatter(log_format_final, datefmt=TIMESTAMP_FORMAT)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler
    log_path = Path(work_dir) / LOG_FILE
    file_handler = logging.FileHandler(log_path, mode='a', encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger

def get_logger():
    """
    Get the current logger instance.
    
    Returns:
        logging.Logger: Current logger instance
    """
    global logger
    return logger

def log_exceptions(error_message_template):
    """
    Decorator to handle common exception logging pattern.
    
    Args:
        error_message_template (str): Template for error message, can include {e} for exception
    
    Returns:
        Decorator function
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger = get_logger()
                if logger:
                    logger.error(error_message_template.format(e=e))
                raise
        return wrapper
    return decorator 