"""Text that stays text in the files ea_irms writes (1.4.10).

openpyxl stores a string that begins with "=" as a formula, and one that reads like an error code
(such as "#N/A") as an error, so a sample named "=100+1" would reach Excel or LibreOffice as a
formula they compute. Every cell that carries text from an upload (an export, a sample sheet, a
configuration, or the run information) goes through ``as_text`` or ``put``. A CSV file has no cell
types, so ``csv_text`` puts an apostrophe before text that a spreadsheet would read as a formula,
the rule OWASP gives for CSV injection; numbers are left as they are.
"""
from __future__ import annotations

# The first characters that make a spreadsheet read a CSV cell as a formula.
FORMULA_START = ("=", "+", "-", "@", "\t", "\r")


def as_text(cell):
    """The cell, with a text value stored as text rather than as a formula or an error."""
    if isinstance(cell.value, str) and cell.data_type in ("f", "e"):
        cell.data_type = "s"
    return cell


def put(ws, row: int, col: int, value):
    """Write a value that came from an upload into ws at (row, col); returns the cell."""
    cell = ws.cell(row, col)
    cell.value = value
    return as_text(cell)


def csv_text(value):
    """A CSV cell: text that begins like a formula gets a leading apostrophe; anything else is unchanged."""
    if isinstance(value, str) and value.startswith(FORMULA_START):
        return "'" + value
    return value
