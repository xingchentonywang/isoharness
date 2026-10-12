"""The engine check before a release: on every reference run, the API's values
must equal its workbook, and the workbook must equal the one delivered.

    python scripts/reference_check.py MANIFEST.json [--url http://localhost:8000]

The reference runs are lab data and stay out of this public repository. The
manifest names them, with paths relative to the manifest's own folder:

    {"runs": [{"name": "PanamaSids2",
               "export": "runs/PanamaSids2_09232026_RAW.xls",
               "config": "configs/cfg_PanamaSids2.json",          (optional)
               "delivered": "delivered/09232026_PanamaSids2.xlsx"} (optional)]}

For each run it asks the API (at --url, or in this process when --url is left
out) for the corrected values and for the workbook, with the run name and date
read from the delivered workbook's Setup sheet, and checks three things:

1. Values. Every position value, flag list and fit result /api/process returns
   equals the workbook /api/workbook returns, after LibreOffice has computed it
   (1e-9 relative), with no formula error and no stray row.
2. Engine. /api/process returns what ea_irms returns when called directly here,
   value for value (only when ea_irms imports).
3. Delivered. The workbook equals the delivered one in every cell: formulas and
   text exactly, computed numbers to 1e-9 relative.

Exits 1 on any mismatch.
"""
from __future__ import annotations

import argparse
import io
import json
import math
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

TOL = 1e-9
# Work sheet columns by letter, and the 1.4.2 columns found by their row-5 header
WORK = {"d15n_final": "AA", "d13c_final": "AF", "pct_n_measured": "AI", "pct_c_measured": "AJ",
        "cn_molar": "AK", "cn_mass": "AL", "ref_bgd_n": "BI", "settling_fraction": "BJ",
        "ref_pulse_d15n": "BT", "ref_pulse_d13c": "BW", "t_drift_n": "CB", "t_drift_c": "CC",
        "u_d15n": "CL", "U_d15n": "CM", "u_d13c": "CN", "U_d13c": "CO",
        "area_n": "L", "d15n": "O", "area_c": "P", "d13c": "R", "bgd_n": "N",
        "d15n_blank": "W", "d15n_lin": "X", "d15n_drift": "Z", "d13c_blank": "AB", "d13c_lin": "AC",
        "d13c_drift": "AE", "ug_n": "AG", "ug_c": "AH"}
WORK_BY_HEADER = {"elem_bias_n": "elem bias N", "elem_bias_c": "elem bias C",
                  "elem_factor_n": "elem factor N", "elem_factor_c": "elem factor C",
                  "pct_n_reported": "%N corrected", "pct_c_reported": "%C corrected",
                  "cn_molar_reported": "C:N molar corrected", "cn_mass_reported": "C:N mass corrected",
                  "pulse_ampl_c": "pulse ampl", "ampl_c": "Ampl 44"}
RESULTS = {"pct_c_reported": 5, "pct_n_reported": 6, "cn_molar_reported": 7, "cn_mass_reported": 8,
           "elem_factor_n": 16, "elem_factor_c": 17, "pct_c_measured": 18, "pct_n_measured": 19}
FITS = {"B20": ("lin_n", "applied"), "B95": ("lin_n", "fit", "p"), "B96": ("lin_n", "in_force"),
        "B28": ("lin_c", "applied"), "B102": ("lin_c", "fit", "p"), "B103": ("lin_c", "in_force"),
        "B39": ("drift_n", "applied"), "B114": ("drift_n", "fit", "p"), "B115": ("drift_n", "in_force"),
        "B51": ("drift_c", "applied"), "B127": ("drift_c", "fit", "p"), "B128": ("drift_c", "in_force"),
        "B56": ("cal_n", "slope"), "B57": ("cal_n", "intercept"), "B63": ("cal_c", "slope"),
        "B64": ("cal_c", "intercept"), "B70": ("k_n", "slope"), "B71": ("k_n", "intercept"),
        "B76": ("k_c", "slope"), "B77": ("k_c", "intercept"),
        "B135": ("unc_n", "pooled_sd"), "B155": ("unc_c", "pooled_sd"),
        "B138": ("cal_n", "acc_min"), "B159": ("cal_c", "acc_max"),
        "B201": ("ref_gas_n", "n"), "B205": ("ref_gas_n", "fit", "slope"), "B206": ("ref_gas_n", "fit", "p"),
        "B208": ("ref_gas_c", "n"), "B213": ("ref_gas_c", "fit", "p"),
        "B217": ("checks", "n"), "B97": ("size_limits", "max_n"),
        "B85": ("settling", "excess"), "B86": ("settling", "settled_from"),
        "B229": ("carryover", "oversized"), "B230": ("carryover", "followed"),
        "B234": ("settling", "level"), "B235": ("settling", "excess_level"),
        "B237": ("ref_pulse_check", "n_designated"), "B238": ("ref_pulse_check", "mean_ampl"),
        "B239": ("ref_pulse_check", "mean_d13c"), "B240": ("ref_pulse_check", "active"),
        "B244": ("elem_n", "n"), "B245": ("elem_n", "max_bias_pct"), "B246": ("elem_n", "in_force"),
        "B247": ("elem_c", "n"), "B248": ("elem_c", "max_bias_pct"), "B249": ("elem_c", "in_force")}


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def near(a, b):
    return abs(a - b) <= TOL * max(1.0, abs(a))


def dig(node, path):
    for k in path:
        node = node.get(k) if isinstance(node, dict) else None
        if node is None:
            return None
    return node


class Client:
    """POST to a running service, or to the app in this process."""

    def __init__(self, url):
        if url:
            import httpx
            self.http = httpx.Client(base_url=url.rstrip("/"), timeout=600)
        else:
            from fastapi.testclient import TestClient
            from isoharness.app import app
            self.http = TestClient(app)

    def post(self, path, export: Path, **form):
        r = self.http.post(path, files={"file": (export.name, export.read_bytes())},
                           data={k: v for k, v in form.items() if v is not None})
        if r.status_code != 200:
            raise RuntimeError(f"{path} returned {r.status_code}: {r.text[:300]}")
        return r


def check_values(body, wb):
    """1. The API's values against the workbook it returned."""
    positions, fits = body["positions"], json.loads(json.dumps(body["fits"]))
    for el in ("elem_n", "elem_c"):
        mb = fits[el]["max_bias"]
        fits[el]["max_bias_pct"] = None if mb is None else mb * 100
    wk, ft, rs = wb["Work"], wb["Fits"], wb["Results"]
    head = {str(wk.cell(5, c).value): get_column_letter(c) for c in range(1, wk.max_column + 1)}
    cols = dict(WORK)
    cols.update({attr: head[name] for attr, name in WORK_BY_HEADER.items()})
    feeds = {c for c in range(1, wk.max_column + 1) if str(wk.cell(5, c).value or "").startswith("ch ")}
    bad, n = [], 0
    errors = [(s.title, c.coordinate, c.value) for s in wb for row in s.iter_rows() for c in row
              if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper()
              and not (s.title == "Work" and c.column in feeds and c.value == "#N/A")]
    for p in positions:
        for attr, col in cols.items():
            mine, theirs = p.get(attr), num(wk[f"{col}{5 + p['row']}"].value)
            n += 1
            if (mine is None) != (theirs is None) or (mine is not None and not near(mine, theirs)):
                bad.append(("Work", attr, p["row"], mine, theirs))
        for attr, col in RESULTS.items():
            mine, theirs = p.get(attr), num(rs.cell(4 + p["row"], col).value)
            n += 1
            if (mine is None) != (theirs is None) or (mine is not None and not near(mine, theirs)):
                bad.append(("Results", attr, p["row"], mine, theirs))
        text = rs[f"K{4 + p['row']}"].value or ""
        n += 1
        if {t.strip() for t in text.split(";") if t.strip()} != set(p["flags"]):
            bad.append(("flags", p["row"], text, p["flags"]))
    for cell, path in FITS.items():
        mine, theirs = dig(fits, path), ft[cell].value
        n += 1
        if isinstance(mine, bool):
            mine = "Yes" if mine else "No"
        if isinstance(mine, float):
            if num(theirs) is None or not near(mine, theirs):
                bad.append(("Fits", cell, path, mine, theirs))
        elif mine != theirs:
            bad.append(("Fits", cell, path, mine, theirs))
    extra = [r for r in range(6 + len(positions), 126) if wk[f"B{r}"].value not in (None, "")]
    if extra:
        bad.append(("extra Work rows", extra[:5]))
    return n, bad, errors


def check_engine(body, export: Path, config_text):
    """2. The API's positions against ea_irms called directly."""
    from ea_irms.parse import parse_export
    from ea_irms.pipeline import Config, config_from_dict, process
    from isoharness.serialize import position_record
    cfg = config_from_dict(json.loads(config_text)) if config_text else Config()
    positions, _ = process(parse_export(export), cfg)
    want = [json.loads(json.dumps(position_record(p))) for p in positions]
    got = body["positions"]
    n, bad = 0, []
    if len(want) != len(got):
        return 1, [("position count", len(got), len(want))]
    for g, w in zip(got, want):
        for k in w:
            n += 1
            a, b = g.get(k), w[k]
            if a != b and not (isinstance(a, float) and isinstance(b, float) and math.isnan(a) and math.isnan(b)):
                bad.append((w["row"], k, a, b))
    return n, bad


def check_delivered(blob, delivered: Path):
    """3. Every cell of the API's workbook against the delivered workbook."""
    n, bad = 0, []
    for data_only in (False, True):
        mine = openpyxl.load_workbook(io.BytesIO(blob), data_only=data_only)
        theirs = openpyxl.load_workbook(delivered, data_only=data_only)
        if mine.sheetnames != theirs.sheetnames:
            return 1, [("sheets", mine.sheetnames, theirs.sheetnames)]
        for name in mine.sheetnames:
            a, b = mine[name], theirs[name]
            rows, cols = max(a.max_row, b.max_row), max(a.max_column, b.max_column)
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    x, y = a.cell(r, c).value, b.cell(r, c).value
                    n += 1
                    if data_only and num(x) is not None and num(y) is not None:
                        if not near(x, y):
                            bad.append((name, a.cell(r, c).coordinate, "value", x, y))
                    elif x != y:
                        bad.append((name, a.cell(r, c).coordinate, "value" if data_only else "formula", x, y))
    return n, bad


def main(argv=None):
    ap = argparse.ArgumentParser(description="engine check over the reference runs")
    ap.add_argument("manifest")
    ap.add_argument("--url", help="a running service; default: the app in this process")
    args = ap.parse_args(argv)
    manifest = Path(args.manifest).resolve()
    runs = json.loads(manifest.read_text())["runs"]
    client = Client(args.url)
    failed = 0
    totals = [0, 0, 0]
    for run in runs:
        export = manifest.parent / run["export"]
        config_text = (manifest.parent / run["config"]).read_text() if run.get("config") else None
        delivered = manifest.parent / run["delivered"] if run.get("delivered") else None
        name = date = None
        if delivered:
            setup = openpyxl.load_workbook(delivered)["Setup"]
            name, date = setup["B5"].value, setup["B6"].value
        body = client.post("/api/process", export, config=config_text).json()
        r = client.post("/api/workbook", export, config=config_text, run_name=name, run_date=date)
        blob = r.content
        computed = r.headers.get("x-isoharness-values")
        wb = openpyxl.load_workbook(io.BytesIO(blob), data_only=True)
        n1, bad1, errors = check_values(body, wb)
        try:
            n2, bad2 = check_engine(body, export, config_text)
        except ImportError:
            n2, bad2 = 0, []
        n3, bad3 = check_delivered(blob, delivered) if delivered else (0, [])
        totals = [totals[0] + n1, totals[1] + n2, totals[2] + n3]
        ok = computed == "computed" and not (bad1 or bad2 or bad3 or errors)
        failed += not ok
        print(f"{run['name']}: {len(body['positions'])} positions, values {computed}; "
              f"values {n1} compared, {len(bad1)} mismatches, {len(errors)} formula errors; "
              f"engine {n2} compared, {len(bad2)} mismatches; "
              f"delivered {n3} cells compared, {len(bad3)} differ" + ("" if delivered else " (none given)"))
        for item in (bad1 + bad2 + bad3 + errors)[:8]:
            print("   ", item)
    print(f"{len(runs)} runs: values {totals[0]}, engine {totals[1]}, delivered cells {totals[2]} compared; "
          + ("ALL OK" if not failed else f"{failed} run(s) FAILED"))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
