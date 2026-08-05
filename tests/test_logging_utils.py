"""Tests for the narrow logging utility retained by the replacement workflow."""

import logging

from src.logging_utils import get_logger, setup_logging


def test_setup_logging_creates_workspace_log(tmp_path):
    logger = setup_logging(str(tmp_path), "[{timestamp}] {message}")

    logger.info("stage-zero-test")
    for handler in logger.handlers:
        handler.flush()

    assert logger.name == "db_updater"
    assert (tmp_path / "update_log.txt").exists()


def test_get_logger_returns_configured_logger(tmp_path):
    configured = setup_logging(str(tmp_path), "[{timestamp}] {message}")

    assert get_logger() is configured
    assert isinstance(configured, logging.Logger)
