"""Shared test settings.

The engine tests read the reference workbook from ``EA_TEMPLATE``. Its Raw sheet
is the 2026-09-15 test run (standards, blanks and one position named Random),
the one real run this public repository carries; ``tests/data/20260915_raw.csv``
is the same export as a CSV file.
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "tests" / "data"
sys.path.insert(0, str(ROOT))
os.environ.setdefault("EA_TEMPLATE", str(DATA / "EA-IRMS_Correction_Template_v2.xlsx"))
