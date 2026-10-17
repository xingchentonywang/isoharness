"""Helpers for the tests of the page's logic (web/logic.js), which run it under node (tests/page_harness.js).

run_page(bodies, code) makes a batch of the /api/process answers in bodies, as the page makes it, and returns
what the JavaScript in code returns, with L the page, engine the preview port, and input the payload. renamed()
writes a copy of the 2026-09-15 test run with the identifiers of some positions changed.
"""
import csv
import io
import json
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RUN = ROOT / "tests" / "data" / "20260915_raw.csv"
NODE = shutil.which("node")


def run_page(bodies, code, **data):
    payload = json.dumps(dict({"bodies": bodies, "code": code}, **data))
    out = subprocess.run([NODE, str(ROOT / "tests" / "page_harness.js")], input=payload, capture_output=True,
                         text=True, cwd=ROOT, timeout=300)
    if out.returncode != 0:
        raise AssertionError("the page's logic failed under node:\n" + out.stderr[-3000:])
    return json.loads(out.stdout)


def renamed(folder, names, file_name, source=RUN):
    """The test run with the identifier of each position in names ({row: name}) changed, as folder/file_name."""
    rows = list(csv.reader(io.StringIO(Path(source).read_text(encoding="utf-8"))))
    h = rows[0]
    row, ident = h.index("Row"), h.index("Identifier 1")
    for r in rows[1:]:
        if r and r[row].strip().isdigit() and int(r[row]) in names:
            r[ident] = names[int(r[row])]
    path = Path(folder) / file_name
    with open(path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh, lineterminator="\n").writerows(rows)
    return path
