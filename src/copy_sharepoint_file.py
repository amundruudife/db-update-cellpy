"""Retired Downloads-based source acquisition.

The replacement implementation will acquire evaluated values from the
hard-coded SharePoint workbook. This module remains only to fail old callers
explicitly during the transition.
"""

from .exceptions import LegacyWorkflowDisabledError


def copy_cell_log_to_source_data():
    raise LegacyWorkflowDisabledError(
        "Downloads/log-sheet acquisition is retired. "
        "Use the approved hard-coded c&p acquisition workflow when implemented."
    )


def main() -> int:
    try:
        copy_cell_log_to_source_data()
    except LegacyWorkflowDisabledError as exc:
        print(exc)
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
