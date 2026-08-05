"""Explicit safety failures for retired and production-write workflows."""


class LegacyWorkflowDisabledError(RuntimeError):
    """Raised when code attempts to invoke a retired legacy workflow."""


class ProductionWriteBlockedError(PermissionError):
    """Raised when non-transactional code targets the production workbook."""
