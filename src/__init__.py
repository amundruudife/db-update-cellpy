#!/usr/bin/env python3
"""
Cell Analysis Database Auto-Update System.

Legacy workflow functions are intentionally not re-exported. The replacement
pipeline is introduced behind validation gates and has no production-write
entry point during Stage 0.
"""

# Version info
__version__ = "2.0.0"
__author__ = "Auto-generated (Refactored)"
__date__ = "2025-01-14"

from .logging_utils import setup_logging, get_logger

__all__ = ["setup_logging", "get_logger"]
