"""End-to-end check of the page in Chromium, against a service started for it.

    python scripts/e2e_web.py [--shots DIR] [--keep DIR] [--url URL]

Starts ``python -m uvicorn isoharness.app:app`` on a free port from the repository
root (LibreOffice on the PATH computes the workbook), or with --url uses a service
already running there (the container image, say), opens the page in Chromium
(Python Playwright, with PLAYWRIGHT_BROWSERS_PATH set) and:

0. opens the home page, checks that it lists the three instruments with two of them
   planned, and follows its link to the EA-IRMS page;
1. uploads two files that are not exports and checks that the page shows the
   service's message for each, word for word;
2. uploads tests/data/20260915_raw.csv and checks the upload screen (the twelve
   columns found, the positions, the page's agreement with the engine);
3. downloads the Excel template of the lab configuration and checks that it reads
   as the package defaults; loads a lab configuration after the upload, as JSON and
   then as the Excel file the service writes for it, checks that the run is processed
   again with it and the decisions start from it, then goes back to the defaults;
4. visits all ten step screens, checking the instrument screen's charts and its
   settling line and the quality-check summary's links back to it, and changing
   decisions on the way: the blank switch,
   a drift monitor left out of the d15N fit, the drift model, a monitor left out of
   the elemental correction, and the d15N repeatability (the drift monitors' SD);
5. checks that every value the results screen shows (mass, d15N, 1σ, d13C, 1σ, %N,
   %C, C:N, and the flags) equals, to the shown precision, the results table the
   service returns for the same decisions, built here from the processed
   configuration without the page;
6. downloads the workbook, the results table and the configuration, checks their
   names (from Content-Disposition), the configuration's fields (the Excel file read
   back by the service), the table, and the workbook's computed Results sheet, and
   posts the downloaded configuration back to /api/results.csv to check that it
   reproduces the page's values;
7. adds two more runs at once (isoharness 0.2.0), copies of the test run with each
   sample named by its position, the second with each sample's N2 peak 0.5 permil
   higher; checks the batch list, that the first run kept its decisions when shown
   again, the samples measured in more than one run, and the combined results and
   table, each run's rows against the service's table for that run; removes a run;
8. saves a screenshot of every screen at 1440 by 1000 (full page) and of the upload
   and results screens at 400 px wide;
9. loads the page afresh, checks the landing page at 400 px and 1440 px wide, follows
   the brand back to the home page and checks it at 400 px, and reads the example run
   through its button.

Any console error, page error, failed request, or HTTP error the test did not
cause on purpose fails the run. Prints one result line; exits 1 on failure.
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import uuid
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import expect, sync_playwright

ROOT = Path(__file__).resolve().parents[1]
RUN = ROOT / "tests" / "data" / "20260915_raw.csv"
SCREENS = ["inst", "setup", "blank", "lin", "drift", "cal", "elem", "unc", "qc", "results"]
TITLES = {"inst": "Instrument performance", "setup": "Standards and roles", "blank": "Blank", "lin": "Linearity", "drift": "Drift",
          "cal": "Calibration", "elem": "Elemental content", "unc": "Uncertainty",
          "qc": "Quality checks", "results": "Results"}
# The step names the Back and Continue buttons use, where they differ from the headings.
BUTTONS = dict(TITLES, inst="Instrument", setup="Setup")
# Results screen column -> results table column, and the decimals the page shows.
SHOWN = {"amt": ("Mass (mg)", 3), "dN": ("d15N (permil)", 2), "UN": ("1sigma(d15N)", 2),
         "dC": ("d13C (permil)", 2), "UC": ("1sigma(d13C)", 2), "pN": ("%N", 2), "pC": ("%C", 2),
         "cn": ("C:N molar", 1)}
# Workbook Results sheet column -> results table column (as scripts/smoke.py reads it).
BOOK = {5: "%C", 6: "1sigma(%C)", 7: "%N", 8: "1sigma(%N)", 9: "C:N molar", 10: "1sigma(C:N molar)",
        11: "C:N mass", 12: "1sigma(C:N mass)", 13: "d13C (permil)", 14: "1sigma(d13C)", 15: "d15N (permil)",
        16: "1sigma(d15N)", 21: "Elemental drift factor N", 22: "Elemental drift factor C", 23: "%C as measured",
        24: "%N as measured"}
# The decisions the test makes, and the drift monitors it leaves out.
EXCL_N, EXCL_E = 18, 55


class Failed(Exception):
    pass


def check(cond, what):
    if not cond:
        raise Failed(what)


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def post(base, path, file: Path, fields=None, name=None):
    """POST a multipart form; returns (status, headers, body bytes)."""
    boundary = uuid.uuid4().hex
    parts = []
    for k, v in (fields or {}).items():
        parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
    parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="{name or file.name}"\r\n'
                 "Content-Type: application/octet-stream\r\n\r\n".encode() + file.read_bytes() + b"\r\n")
    parts.append(f"--{boundary}--\r\n".encode())
    req = urllib.request.Request(base + path, data=b"".join(parts), method="POST",
                                 headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()


def get(base, path):
    """GET; returns (status, headers, body bytes)."""
    try:
        with urllib.request.urlopen(base + path, timeout=60) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()


def post_json(base, path, obj):
    """POST a JSON body; returns (status, headers, body bytes)."""
    req = urllib.request.Request(base + path, data=json.dumps(obj).encode(), method="POST",
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()


def table_diff(a_text, b_text, rel=1e-12):
    """Cells of two results tables that differ, text exactly and numbers by more than rel, and the
    largest relative difference between numbers. An Excel file keeps 16 significant digits, so a
    configuration read back from one may carry a value computed to more digits a few 1e-16 apart."""
    a, b = list(csv.reader(io.StringIO(a_text))), list(csv.reader(io.StringIO(b_text)))
    if len(a) != len(b):
        return [f"{len(a)} rows against {len(b)}"], 0.0
    out, worst = [], 0.0
    for i, (ra, rb) in enumerate(zip(a, b)):
        if len(ra) != len(rb):
            out.append(f"row {i}: {len(ra)} cells against {len(rb)}")
            continue
        for j, (x, y) in enumerate(zip(ra, rb)):
            if x == y:
                continue
            try:
                fx, fy = float(x), float(y)
            except ValueError:
                out.append(f"row {i}, column {j}: {x!r} against {y!r}")
                continue
            d = abs(fx - fy) / max(abs(fx), 1e-300)
            worst = max(worst, d)
            if abs(fx - fy) > rel * max(1.0, abs(fx)):
                out.append(f"row {i}, column {j}: {x} against {y}")
    return out, worst


def rows_of(text):
    return {int(r["Row"]): r for r in csv.DictReader(io.StringIO(text))}


def variant(path: Path, samples: set, shift=0.0) -> bytes:
    """The test run with each sample (its positions in samples) named S<position> and its N2 peak's
    d15N moved by shift, as tests/test_batch.py builds it: the sample peak is the position's last
    mass-28 peak."""
    rows = list(csv.reader(io.StringIO(path.read_text(encoding="utf-8"))))
    h = rows[0]
    col = {k: h.index(k) for k in ("Row", "Identifier 1", "Ampl  28", "Is Ref _", "d 15N/14N")}
    last28 = {r[col["Row"]].strip(): i for i, r in enumerate(rows[1:], start=1)
              if len(r) > col["Ampl  28"] and r[col["Ampl  28"]].strip()}
    for i, r in enumerate(rows[1:], start=1):
        row = r[col["Row"]].strip() if r else ""
        if row in samples:
            r[col["Identifier 1"]] = f"S{row}"
            if shift and last28.get(row) == i and r[col["Is Ref _"]].strip() != "1":
                r[col["d 15N/14N"]] = repr(float(r[col["d 15N/14N"]]) + shift)
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    return buf.getvalue().encode("utf-8")


def sheet_diff(rows, text, rel=1e-15):
    """Rows of an Excel sheet (values) against a results table's rows (CSV text, header left out):
    the cells that differ, and the largest relative difference between numbers."""
    table = list(csv.reader(io.StringIO(text)))[1:]
    if len(rows) != len(table):
        return [f"{len(rows)} rows against {len(table)}"], 0.0
    out, worst = [], 0.0
    for i, (ra, rb) in enumerate(zip(rows, table)):
        for j, (a, b) in enumerate(zip(ra, rb)):
            if b == "":
                if a not in (None, ""):
                    out.append(f"row {i}, column {j}: {a!r} against empty")
                continue
            if isinstance(a, str) or a is None:
                if a != b:
                    out.append(f"row {i}, column {j}: {a!r} against {b!r}")
                continue
            d = abs(a - float(b))
            worst = max(worst, d / max(abs(float(b)), 1e-300))
            if d > rel * max(1.0, abs(float(b))):
                out.append(f"row {i}, column {j}: {a!r} against {b}")
    return out, worst


def ftext(flag):
    """The page's wording of an engine flag (logic.js ftext)."""
    return re.sub(r"(\d)x the blank", r"\1× the blank", flag.replace("d15N", "δ¹⁵N").replace("d13C", "δ¹³C"), count=1)


def js_fixed(value, d):
    """value.toFixed(d) as the page writes it, with its minus sign."""
    q = Decimal(abs(value)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    s = f"{q:.{d}f}"
    return ("−" if value < 0 and float(s) != 0 else "") + s


def agree(shown, value, d):
    """Whether a shown value equals the table's to the shown precision; and whether exactly as formatted."""
    if value in ("", None):
        return shown == "", shown == ""
    if shown == "":
        return False, False
    v = float(value)
    close = abs(float(shown.replace("−", "-")) - v) <= 0.5 * 10 ** -d + 1e-9
    return close, shown == js_fixed(v, d)


def compare(page_rows, table, label):
    """Every shown value of every sample against a results table; returns counts."""
    n = values = exact = 0
    for r in page_rows:
        row = table.get(int(r["r"]))
        check(row is not None, f"{label}: position {r['r']} is not in the results table")
        check(row["Type"] == "Sample", f"{label}: position {r['r']} is a {row['Type']}, not a sample")
        for key, (col, d) in SHOWN.items():
            ok, same = agree(r[key], row[col], d)
            check(ok, f"{label}: position {r['r']} {col}: page shows {r[key]!r}, table has {row[col]!r}")
            values += 1
            exact += same
        want = [ftext(f) for f in row["Flags"].split("; ") if f]
        check(r["flags"] == want, f"{label}: position {r['r']} flags: page {r['flags']}, table {want}")
        n += 1
    return n, values, exact


def sd(values):
    m = sum(values) / len(values)
    return math.sqrt(sum((v - m) ** 2 for v in values) / (len(values) - 1))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--shots", default=str(ROOT / "build" / "e2e_web" / "shots"), help="folder for the screenshots")
    ap.add_argument("--keep", help="folder to keep the downloads in (default: a temporary folder)")
    ap.add_argument("--url", help="a running service to check instead of starting one")
    args = ap.parse_args(argv)
    shots = Path(args.shots)
    shots.mkdir(parents=True, exist_ok=True)
    work = Path(args.keep) if args.keep else Path(tempfile.mkdtemp(prefix="e2e_web-"))
    work.mkdir(parents=True, exist_ok=True)

    log = open(work / "server.log", "w")
    if args.url:
        base, server = args.url.rstrip("/"), None
    else:
        port = free_port()
        base = f"http://127.0.0.1:{port}"
        server = subprocess.Popen([sys.executable, "-m", "uvicorn", "isoharness.app:app", "--host", "127.0.0.1",
                                   "--port", str(port)], cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    noise = []          # console errors, page errors, failed requests and HTTP errors
    expected = []      # the ones this test causes on purpose
    state = {"expect": None}
    stats = {}
    try:
        for _ in range(150):
            try:
                with urllib.request.urlopen(base + "/api/health", timeout=2) as r:
                    if json.loads(r.read()) == {"status": "ok"}:
                        break
            except (urllib.error.URLError, ConnectionError, TimeoutError, json.JSONDecodeError):
                time.sleep(0.2)
        else:
            raise Failed("the service did not start; see " + str(work / "server.log"))
        with urllib.request.urlopen(base + "/api/version", timeout=10) as r:
            version = json.loads(r.read())
        check(version["libreoffice"], "LibreOffice is not on the PATH, so the workbook would carry formulas only")
        ver = version["ea_irms"]

        def note(kind, text, url=""):
            want = state["expect"]
            if want and ((kind == "console" and re.search(r"status of (415|422)", text))
                         or (kind == "http" and url.endswith("/api/process") and text in ("415", "422"))):
                expected.append((kind, text))
                return
            noise.append(f"{kind}: {text} {url}".strip())

        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            context = browser.new_context(viewport={"width": 1440, "height": 1000}, accept_downloads=True)
            page = context.new_page()
            page.on("console", lambda m: note("console", m.text) if m.type == "error" else None)
            page.on("pageerror", lambda e: note("pageerror", str(e)))
            page.on("requestfailed", lambda r: note("requestfailed", r.failure or "", r.url))
            page.on("response", lambda r: note("http", str(r.status), r.url) if r.status >= 400 else None)

            def shot(name, full=True):
                page.wait_for_timeout(250)
                page.screenshot(path=str(shots / name), full_page=full)

            def no_prototype():
                check("Prototype" not in page.inner_text("body"), "the page says Prototype")

            def heading(scr):
                expect(page.get_by_role("heading", level=1, name=TITLES[scr], exact=True)).to_be_visible()

            # ---- the home page, which lists the instruments
            page.goto(base + "/")
            expect(page.get_by_role("heading", level=1, name="Stable isotope data reduction")).to_be_visible()
            names = page.locator("main li h3").all_text_contents()
            check(names == ["EA-IRMS", "GasBench N₂O", "GasBench carbonate"], f"the home page lists {names}")
            status = page.locator("main li .status").all_text_contents()
            check(status == ["Available", "Planned", "Planned"], f"the home page marks them {status}")
            shot("1440-home.png")
            page.get_by_role("link", name="Open EA-IRMS").click()
            page.wait_for_url(re.compile(r"/ea-irms$"))

            # ---- the landing page
            expect(page.get_by_text("Drop exports here, or choose files")).to_be_visible()
            check("deletes each file before it sends its reply" in page.inner_text("body"),
                  "the upload screen does not say that the file is deleted")
            no_prototype()
            shot("1440-01-upload-empty.png")

            # ---- files that are not exports: the service's message, as written
            pdf = work / "notes.pdf"
            pdf.write_bytes(b"%PDF-1.4\n% not an export\n")
            bad_csv = work / "not_an_export.csv"
            bad_csv.write_text("sample,mass\nA,0.4\n")
            state["expect"] = True
            for bad in (pdf, bad_csv):
                status, _, body = post(base, "/api/process", bad)
                detail = json.loads(body)["detail"]
                check(status in (415, 422), f"{bad.name}: the service answered {status}")
                page.set_input_files("#dc-export", str(bad))
                alert = page.locator("[role=alert]")
                expect(alert).to_contain_text(f"{bad.name} was not processed")
                expect(page.locator('[data-dc="error-detail"]')).to_have_text(detail)
            state["expect"] = None
            check(len([e for e in expected if e[0] == "http"]) == 2, f"expected two refused uploads, saw {expected}")
            shot("1440-02-upload-refused.png")

            # ---- the test run
            page.set_input_files("#dc-export", str(RUN))
            expect(page.get_by_text("Read without errors")).to_be_visible(timeout=30000)
            check(page.locator("[role=alert]").count() == 0, "the refusal is still shown after a good upload")
            cols = page.locator('[aria-label="Column mapping"] > [role=row]')
            check(cols.count() == 13, f"the column table has {cols.count() - 1} rows, not 12")
            found = page.locator('[aria-label="Column mapping"] [role=cell]:has-text("Found")').count()
            check(found == 12, f"{found} of 12 columns marked found")
            body_text = page.inner_text("body")
            check("56 positions" in body_text, "the upload screen does not count 56 positions")
            check("match the server’s at all 56 positions" in body_text,
                  "the page's own check against the engine did not pass")
            check("Run date 15 Sep 2026, read from the file name" in body_text, "the run date was not read from the file name")
            shot("1440-03-upload.png")

            # ---- the Excel template of the lab configuration (ea_irms 1.4.6), as the package defaults
            with page.expect_download(timeout=30000) as dl:
                page.get_by_role("button", name="Download the template").click()
            template = work / dl.value.suggested_filename
            dl.value.save_as(str(template))
            check(template.name == f"lab_configuration_ea_irms_v{ver}.xlsx", f"template named {template.name}")
            status, _, body = post(base, "/api/config/read", template)
            check(status == 200, f"/api/config/read answered {status} for the template")
            status, _, dflt = get(base, "/api/defaults")
            check(json.loads(body)["config"] == json.loads(dflt)["config"], "the template is not the package defaults")

            # ---- a lab configuration loaded after the upload is processed on the server again,
            # first as JSON, then as the Excel file the service writes for the same settings
            lab = work / "lab-config.json"
            lab.write_text(json.dumps({"blank_correction": "On", "coverage_k": 3.0}))
            with page.expect_request(lambda r: r.url.endswith("/api/process") and r.method == "POST"):
                page.set_input_files("#dc-config", str(lab))
            expect(page.get_by_text("The run was processed again with lab-config.json")).to_be_visible(timeout=30000)
            status, _, body = post_json(base, "/api/config.xlsx", {"config": json.loads(lab.read_text()), "name": "lab"})
            check(status == 200, f"/api/config.xlsx answered {status}")
            lab = work / "lab-config.xlsx"
            lab.write_bytes(body)
            # (k = 3 on the uncertainty step can only come from the server, through the port block's settings)
            with page.expect_request(lambda r: r.url.endswith("/api/process") and r.method == "POST"):
                page.set_input_files("#dc-config", str(lab))
            expect(page.get_by_text("The run was processed again with lab-config.xlsx")).to_be_visible(timeout=30000)
            page.get_by_role("button", name="Start the corrections").click()
            heading("inst")
            page.locator("nav[aria-label='Correction steps'] button", has_text="Setup").click()
            heading("setup")
            expect(page.get_by_text("From lab-config.xlsx.")).to_be_visible()
            page.locator("nav[aria-label='Correction steps'] button", has_text="Blank").click()
            heading("blank")
            expect(page.get_by_role("group", name="Blank correction").get_by_role("button", name="On")).to_have_attribute("aria-pressed", "true")
            page.locator("nav[aria-label='Correction steps'] button", has_text="Uncertainty").click()
            heading("unc")
            expect(page.get_by_text("by more than their uncertainty, it uses 3σ")).to_be_visible()
            page.locator("nav[aria-label='Correction steps'] button", has_text="Results").click()
            heading("results")
            expect(page.get_by_text("with 1σ, %N, %C, and C:N")).to_be_visible()
            page.get_by_role("navigation", name="Main").get_by_role("button", name="EA-IRMS").click()
            with page.expect_request(lambda r: r.url.endswith("/api/process") and r.method == "POST"):
                page.get_by_role("button", name="Use the package defaults").click()
            expect(page.get_by_text("processed again with the package defaults")).to_be_visible(timeout=30000)

            # ---- the steps, with decisions on the way
            page.get_by_role("button", name="Start the corrections").click()
            for scr in SCREENS:
                heading(scr)
                no_prototype()
                if scr == "inst":
                    for title in ("N2 reference pulses", "CO2 reference pulses", "Reference peak size", "Background at m/z 28"):
                        expect(page.locator("figure", has_text=title)).to_have_count(1)
                    expect(page.get_by_text("26 positions ran before the background settled at position 27")).to_be_visible()
                    expect(page.get_by_text("under the 0.1‰ limit the FIRMS guide")).to_have_count(2)
                    check(page.get_by_role("button", name=re.compile(r"^Step 1, ")).count() == 0,
                          "the instrument screen links to itself")
                if scr == "setup":
                    expect(page.get_by_text("Package defaults, ea_irms " + ver + ".")).to_be_visible()
                if scr == "blank":
                    group = page.get_by_role("group", name="Blank correction")
                    expect(group.get_by_role("button", name="Auto")).to_have_attribute("aria-pressed", "true")
                    group.get_by_role("button", name="On").click()
                    expect(group.get_by_role("button", name="On")).to_have_attribute("aria-pressed", "true")
                    expect(page.get_by_text("δ¹⁵N: subtracted")).to_be_visible()
                if scr == "drift":
                    chart = page.locator("figure", has_text="δ¹⁵N of the drift monitors")
                    point = chart.get_by_role("button", name=re.compile(rf", position {EXCL_N}: "))
                    point.click()
                    expect(chart.get_by_role("button", name=re.compile(rf", position {EXCL_N}: .*Left out"))).to_have_count(1)
                    expect(chart).to_contain_text("12 of 13 in the fit")
                    page.get_by_role("group", name="Model").get_by_role("button", name="Background").click()
                    expect(chart).to_contain_text("against background")
                if scr == "elem":
                    chart = page.locator("figure", has_text="%N of the drift monitors against certified")
                    chart.get_by_role("button", name=re.compile(rf", position {EXCL_E}: ")).click()
                    expect(chart).to_contain_text("12 of 13 monitors used")
                if scr == "qc":
                    back = page.get_by_role("button", name="Step 1, Instrument")
                    expect(back).to_have_count(5)
                    back.first.click()
                    heading("inst")
                    page.locator("nav[aria-label='Correction steps'] button", has_text="Quality checks").click()
                    heading("qc")
                if scr == "unc":
                    group = page.get_by_role("group", name="δ¹⁵N repeatability")
                    group.get_by_role("button", name="Drift monitors").click()
                    expect(group.get_by_role("button", name="Drift monitors")).to_have_attribute("aria-pressed", "true")
                    expect(page.get_by_text("SD of the drift monitors in the fit")).to_be_visible()
                shot(f"1440-{SCREENS.index(scr) + 4:02d}-{scr}.png")
                if scr != "results":
                    page.get_by_role("button", name="Continue to " + BUTTONS[SCREENS[SCREENS.index(scr) + 1]]).first.click()

            # ---- the results screen against the service, decisions built here without the page
            page_rows = page.eval_on_selector_all(
                'section[aria-label="Samples"] [role="table"] > button',
                """els => els.map(b => { const c = [...b.children].map(x => x.textContent);
                     return { r: c[0], id: c[1], amt: c[2], aN: c[3], dN: c[4], UN: c[5], dC: c[6], UC: c[7],
                              pN: c[8], pC: c[9], cn: c[10],
                              flags: [...b.children[11].children].map(x => x.getAttribute('title')) }; })""")
            check(len(page_rows) >= 5, f"the results screen lists {len(page_rows)} samples")
            status, _, body = post(base, "/api/process", RUN)
            processed = json.loads(body)["config"]
            mine = dict(processed, blank_correction="On", exclude_drift_n=[EXCL_N], drift_model="Background",
                        exclude_elemental=[EXCL_E])
            status, _, body = post(base, "/api/process", RUN, {"config": json.dumps(mine)})
            check(status == 200, f"/api/process answered {status} for the test's own configuration")
            positions = json.loads(body)["positions"]
            mons = [p["d15n_drift"] for p in positions
                    if p["role"] == "Drift" and p["row"] != EXCL_N and p["d15n_drift"] is not None]
            mine["u_sample_n"] = sd(mons)
            status, _, body = post(base, "/api/results.csv", RUN, {"config": json.dumps(mine)})
            check(status == 200, f"/api/results.csv answered {status}")
            independent = rows_of(body.decode())
            n, values, exact = compare(page_rows, independent, "results screen against the service")
            stats["compared"] = (n, values, exact)
            status, _, body = post(base, "/api/results.csv", RUN)
            defaults = rows_of(body.decode())
            moved = sum(1 for r in page_rows if not agree(r["dN"], defaults[int(r["r"])]["d15N (permil)"], 2)[0])
            check(moved > 0, "the decisions changed no shown d15N")
            stats["moved"] = moved

            # ---- downloads
            files = page.locator('section[aria-label="Files and methods"]')
            got = {}
            for label, timeout in (("Workbook", 180000), ("Results table", 60000), ("Configuration", 30000)):
                with page.expect_download(timeout=timeout) as dl:
                    files.get_by_role("button", name=re.compile("^" + label)).click()
                d = dl.value
                path = work / d.suggested_filename
                d.save_as(str(path))
                got[label] = path
                expect(files.get_by_role("button", name=re.compile("^" + label))).to_be_enabled(timeout=timeout)
            check(got["Workbook"].name == f"20260915_raw_ea_irms_v{ver}.xlsx", f"workbook named {got['Workbook'].name}")
            check(got["Results table"].name == f"20260915_raw_results_ea_irms_v{ver}.csv", f"table named {got['Results table'].name}")
            check(got["Configuration"].name == "20260915_raw_config.xlsx", f"configuration named {got['Configuration'].name}")
            check(page.locator("[role=status]").count() == 0, "a download left a message: " + (
                page.locator("[role=status]").first.inner_text() if page.locator("[role=status]").count() else ""))

            # the configuration, read back by the service: the decisions in their fields, every other field
            # as processed
            status, _, body = post(base, "/api/config/read", got["Configuration"])
            check(status == 200, f"/api/config/read answered {status} for the downloaded configuration")
            check(json.loads(body)["notes"] == [], f"reading the configuration left notes: {json.loads(body)['notes']}")
            cfg = json.loads(body)["config"]
            want = dict(processed, blank_correction="On", linearity_correction="Auto", linearity_correction_c="Auto",
                        drift_correction="Auto", drift_model="Background", exclude_drift_n=[EXCL_N], exclude_drift_c=[],
                        elemental_correction="Auto", exclude_elemental=[EXCL_E], u_sample_c=None)
            check(set(cfg) == set(want), f"configuration fields differ: {sorted(set(cfg) ^ set(want))}")
            for k, v in want.items():
                if k != "u_sample_n":
                    check(cfg[k] == v, f"configuration {k} is {cfg[k]!r}, not {v!r}")
            check(abs(cfg["u_sample_n"] - mine["u_sample_n"]) <= 1e-12 * mine["u_sample_n"],
                  f"u_sample_n is {cfg['u_sample_n']}, the monitors' SD is {mine['u_sample_n']}")

            # the downloaded configuration posted back reproduces the page and the downloaded table
            status, _, body = post(base, "/api/results.csv", RUN, {"config": json.dumps(cfg)})
            check(status == 200, f"/api/results.csv answered {status} for the downloaded configuration")
            diff, worst = table_diff(body.decode(), got["Results table"].read_text())
            check(not diff, "the downloaded table differs from the one its configuration gives: " + "; ".join(diff[:3]))
            stats["round trip"] = worst
            round_trip = rows_of(body.decode())
            compare(page_rows, round_trip, "results screen against the downloaded configuration")
            compare(page_rows, rows_of(got["Results table"].read_text()), "results screen against the downloaded table")

            # the workbook: computed by LibreOffice, and holding the same values
            book = openpyxl.load_workbook(got["Workbook"], data_only=True)["Results"]
            cells = 0
            for row, rec in round_trip.items():
                for c, col in BOOK.items():
                    mine_v = rec[col]
                    theirs = book.cell(4 + row, c).value
                    theirs = theirs if isinstance(theirs, (int, float)) and not isinstance(theirs, bool) else None
                    check((mine_v == "") == (theirs is None), f"workbook row {row} {col}: {theirs!r} against {mine_v!r}")
                    if theirs is not None:
                        cells += 1
                        check(abs(float(mine_v) - theirs) <= 1e-9 * max(1.0, abs(theirs)),
                              f"workbook row {row} {col}: {theirs!r} against {mine_v!r}")
            check(cells > 100, f"only {cells} computed values in the workbook")
            stats["book"] = cells

            # the name comes from Content-Disposition: a stand-in answer names the file differently
            page.route("**/api/results.csv", lambda route: route.fulfill(
                status=200, content_type="text/csv", body="Row\n",
                headers={"Content-Disposition": 'attachment; filename="e2e-named-by-the-service.csv"'}))
            with page.expect_download(timeout=60000) as dl:
                files.get_by_role("button", name=re.compile("^Results table")).click()
            check(dl.value.suggested_filename == "e2e-named-by-the-service.csv",
                  f"the page named the table {dl.value.suggested_filename}, not as the service did")
            page.unroute("**/api/results.csv")
            expect(files.get_by_role("button", name=re.compile("^Results table"))).to_be_enabled(timeout=60000)

            # ---- several runs at once (isoharness 0.2.0)
            before = page.locator('section[aria-label="Samples"]').inner_text()
            run_a, run_b = work / "20260916_run_a.csv", work / "20260917_run_b.csv"
            status, _, body = post(base, "/api/process", RUN)
            samples = {str(q["row"]) for q in json.loads(body)["positions"] if q["role"] == "Sample"}
            run_a.write_bytes(variant(RUN, samples))
            run_b.write_bytes(variant(RUN, samples, 0.5))
            page.get_by_role("navigation", name="Main").get_by_role("button", name="EA-IRMS").click()
            page.set_input_files("#dc-export", [str(run_a), str(run_b)])
            expect(page.get_by_text("Batch of 3 runs")).to_be_visible(timeout=60000)
            expect(page.get_by_role("heading", level=1, name="Correct 3 runs")).to_be_visible()
            listed = page.locator('section[aria-label="Runs in this batch"] li')
            check(listed.count() == 3, f"the batch lists {listed.count()} runs, not 3")
            dated = (("20260915_raw.csv", "15 Sep 2026"), ("20260916_run_a.csv", "16 Sep 2026"),
                     ("20260917_run_b.csv", "17 Sep 2026"))
            for k, (name, date) in enumerate(dated):
                text = listed.nth(k).inner_text()
                check(name in text and "run of " + date in text and "Read without errors" in text,
                      f"batch row {k + 1} reads {text!r}")
            check(listed.nth(2).get_attribute("aria-current") == "true", "the last run added is not the one shown")
            shot("1440-batch-upload.png")
            page.get_by_role("button", name="Show 20260915_raw.csv").click()
            page.get_by_role("button", name="Start the corrections").click()
            heading("results")
            expect(page.get_by_text("Run 1 of 3")).to_be_visible()
            check(page.locator('section[aria-label="Samples"]').inner_text() == before,
                  "the first run's results changed when two runs were added and it was shown again")
            reps = page.locator('[aria-label="Samples measured in more than one run"] [role=row]')
            status, _, body = post(base, "/api/process", run_a)
            slope = json.loads(body)["fits"]["cal_n"]["slope"]
            check(reps.count() - 1 == 26, f"{reps.count() - 1} samples measured in more than one run, not 26")
            ranges = {reps.nth(i).locator("[role=cell]").nth(5).inner_text() for i in range(1, reps.count())}
            want = js_fixed(0.5 * slope, 2)
            check(ranges == {want}, f"the d15N ranges read {ranges}, not {want}")
            shot("1440-batch-results.png")
            # the combined results: each run's rows equal the service's table for that run, run 1 with its decisions
            files = page.locator('section[aria-label="Files and methods"]')
            combined = {}
            for label, key in (("Combined results, all 3 runs", "xlsx"), ("Combined results table", "csv")):
                with page.expect_download(timeout=120000) as dl:
                    files.get_by_role("button", name=re.compile("^" + label)).click()
                d = dl.value
                combined[key] = work / d.suggested_filename
                d.save_as(str(combined[key]))
                expect(files.get_by_role("button", name=re.compile("^" + label))).to_be_enabled(timeout=120000)
            for key in ("xlsx", "csv"):
                check(combined[key].name == f"combined_3_runs_results_ea_irms_v{ver}.{key}",
                      f"the combined {key} file is named {combined[key].name}")
            book = openpyxl.load_workbook(combined["xlsx"])
            check(book.sheetnames == ["README", "Results", "Runs", "Repeats"], f"combined sheets {book.sheetnames}")
            rows = list(book["Results"].iter_rows(min_row=2, values_only=True))
            check(len(rows) == 168, f"the combined Results sheet has {len(rows)} rows, not 168")
            lines = combined["csv"].read_text().splitlines()
            check(len(lines) == 169 and lines[0].startswith("Run,Row,Identifier"),
                  f"the combined table has {len(lines)} lines, the first {lines[0][:40]!r}")
            tables = [got["Results table"].read_text()]
            for f in (run_a, run_b):
                status, _, body = post(base, "/api/results.csv", f)
                check(status == 200, f"/api/results.csv answered {status} for {f.name}")
                tables.append(body.decode())
            worst = 0.0
            for k, (label, table) in enumerate(zip(("20260915_raw", "20260916_run_a", "20260917_run_b"), tables)):
                part = rows[56 * k:56 * (k + 1)]
                check({r[0] for r in part} == {label},
                      f"combined rows {56 * k + 1} to {56 * (k + 1)} are not all {label}")
                diff, w = sheet_diff([r[1:] for r in part], table)
                check(not diff, f"combined results, {label}: " + "; ".join(diff[:3]))
                worst = max(worst, w)
                own = [ln.split(",", 1)[1] for ln in lines[1 + 56 * k:1 + 56 * (k + 1)]]
                check(own == table.splitlines()[1:], f"combined table, {label}: not the run's own table")
            runs = list(book["Runs"].iter_rows(min_row=2, values_only=True))
            check([r[2] for r in runs] == ["2026-09-15", "2026-09-16", "2026-09-17"],
                  f"run dates {[r[2] for r in runs]}")
            status, _, body = post(base, "/api/process", RUN, {"config": json.dumps(mine)})
            fits = json.loads(body)["fits"]
            n_on, c_on = fits["blank_n"]["in_force"], fits["blank_c"]["in_force"]
            blank = "N and C" if n_on and c_on else "N only" if n_on else "C only" if c_on else "no"
            check(runs[0][7] == blank and blank != "no",
                  f"the first run's blank reads {runs[0][7]!r} on the Runs sheet, not {blank!r}")
            rep = list(book["Repeats"].iter_rows(values_only=True))
            ix = {h: i for i, h in enumerate(rep[0])}
            off = [r for r in rep[1:] if r[ix["d15N minus mean"]] is not None
                   and abs(abs(r[ix["d15N minus mean"]]) - 0.25 * slope) > 1e-9]
            check(len(rep) - 1 == 52 and not off, f"Repeats: {len(rep) - 1} rows, {len(off)} off by more than 1e-9")
            stats["batch"] = (len(rows), worst, len(rep) - 1)
            page.get_by_role("navigation", name="Main").get_by_role("button", name="EA-IRMS").click()
            page.get_by_role("button", name="Remove 20260917_run_b.csv").click()
            expect(page.get_by_text("Batch of 2 runs")).to_be_visible()
            page.get_by_role("button", name="Start the corrections").click()
            heading("results")

            # ---- 400 px wide
            page.set_viewport_size({"width": 400, "height": 1000})
            shot("400-results.png")
            page.get_by_role("navigation", name="Main").get_by_role("button", name="EA-IRMS").click()
            expect(page.get_by_text("Read without errors").first).to_be_visible()
            shot("400-upload.png")
            wide = page.evaluate("document.documentElement.scrollWidth")
            check(wide <= 400, f"the upload screen is {wide} px wide at a 400 px viewport")

            # ---- the landing page, fresh, at both widths; the home page at 400 px; the example run
            page.goto(base + "/ea-irms")
            expect(page.get_by_role("heading", level=1, name="EA-IRMS data reduction", exact=True)).to_be_visible()
            shot("400-landing.png")
            wide = page.evaluate("document.documentElement.scrollWidth")
            check(wide <= 400, f"the landing page is {wide} px wide at a 400 px viewport")
            page.get_by_role("link", name="isoharness, all instruments").click()
            expect(page.get_by_role("heading", level=1, name="Stable isotope data reduction")).to_be_visible()
            shot("400-home.png")
            wide = page.evaluate("document.documentElement.scrollWidth")
            check(wide <= 400, f"the home page is {wide} px wide at a 400 px viewport")
            page.set_viewport_size({"width": 1440, "height": 1000})
            page.goto(base + "/ea-irms")
            shot("1440-00-landing.png")
            with page.expect_request(lambda r: r.url.endswith("/api/example") and r.method == "GET"):
                page.get_by_role("button", name="Try the 15 Sep 2026 test run").click()
            expect(page.get_by_text("Read without errors")).to_be_visible(timeout=30000)
            check("56 positions" in page.inner_text("body"), "the example run does not count 56 positions")
            check(page.get_by_role("heading", level=1, name="Correct a run").count() == 1,
                  "the upload screen has no heading once a run is read")

            context.close()
            browser.close()
        check(not noise, "unexpected console errors or failed requests:\n  " + "\n  ".join(noise))
    except (Failed, PlaywrightError) as exc:
        print(f"e2e_web: FAILED: {exc}")
        if noise:
            print("  also seen:\n    " + "\n    ".join(noise))
        return 1
    finally:
        if server is not None:
            server.terminate()
            try:
                server.wait(timeout=20)
            except subprocess.TimeoutExpired:
                server.kill()
        log.close()
    n, values, exact = stats["compared"]
    print(f"e2e_web: passed. {len(SCREENS) + 1} screens and the home page; {n} samples on the results screen equal the service's table for the "
          f"same decisions ({values} values and {n} flag lists, {exact} formatted identically; {stats['moved']} "
          f"d15N values moved by the decisions); workbook, table and configuration downloaded, {stats['book']} "
          f"workbook values checked, configuration round trip within {stats['round trip']:.1e} relative; "
          f"2 refused uploads shown as written; a batch of 3 runs, {stats['batch'][0]} combined rows equal to each "
          f"run's own table within {stats['batch'][1]:.1e} relative, {stats['batch'][2]} repeat analyses; "
          f"0 console errors, 0 failed requests; screenshots in {shots}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
