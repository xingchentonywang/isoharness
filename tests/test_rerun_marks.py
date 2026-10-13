"""ea_irms 1.4.4: rerun marks, in the engine and in the workbook.

A sample is marked only for what a rerun can fix: a peak outside the linearity
range or under the blank multiple, carryover from an oversized sample, or the
reference pulse at its position off; and for a check when it ran before the
background settled or lacks one element's peak. The engine's marks are checked
against those rules on the 2026-09-15 test run and on copies of it with masses
and shifted reference pulses, and the workbook against the engine under six
configurations, after LibreOffice has computed the formulas (skipped without it).
"""
import csv
import io
import json
import math
import shutil
import statistics
import subprocess
from pathlib import Path

import openpyxl
import pytest

from ea_irms.parse import parse_export
from ea_irms.pipeline import ROBUST_SD, Config, _mass_window, config_from_dict, config_to_dict, process
from ea_irms.workbook import build_workbook
from isoharness.serialize import results_csv

DATA = Path(__file__).resolve().parent / "data"
RUN = DATA / "20260915_raw.csv"
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def variant(folder, name, masses=False, shifts=()):
    """A copy of the test run, with masses for its samples and reference pulses moved."""
    rows = list(csv.DictReader(open(RUN, newline="")))
    for r in rows:
        if masses and r["Identifier 1"].strip() == "Random":
            r["Amount"] = f"{0.150 + 0.025 * (int(r['Row']) % 17):.3f}"
        for row, peak, col, delta in shifts:
            if r["Row"] == str(row) and r["Peak Nr"] == str(peak):
                r[col] = f"{float(r[col]) + delta:.3f}"
    out = Path(folder) / f"{name}.csv"
    with open(out, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    return out


def check_config(offset):
    data = json.loads((DATA / "config_check_standard.json").read_text())
    data["check_offset_permil"] = offset
    return config_from_dict(data)


def test_marks_on_the_test_run():
    positions, F = process(parse_export(RUN), Config())
    marks = {p.row: p.rerun for p in positions if p.rerun}
    assert marks == {20: "Check", 21: "Check", 23: "N and C", 24: "Check", 25: "Check", 29: "C",
                     33: "N and C", 40: "N and C", 41: "N and C", 43: "N and C", 44: "N and C",
                     45: "N and C", 47: "N and C", 53: "N"}
    assert all(p.rerun is None for p in positions if p.role != "Sample")
    p53 = next(p for p in positions if p.row == 53)
    assert p53.rerun_why == "N below linearity range; no C peak; N peak only 0.7x the blank"
    assert all(p.rerun_mass is None for p in positions)          # its samples carry no mass
    assert F["rerun"]["counts"] == {"N": 1, "C": 1, "N and C": 8, "Check": 4}
    assert F["rerun"]["run_level_text"] == "none found"
    # δ outside the anchor range is a flag, never a mark on its own
    outside = [p for p in positions if p.role == "Sample" and p.flags == ["d15N outside anchor range"]]
    assert outside and all(p.rerun is None for p in outside)


def test_a_reference_pulse_off_marks_its_element(tmp_path):
    shifted = variant(tmp_path, "shift", shifts=((20, 1, "d 15N/14N", 1.0), (24, 6, "d 13C/12C", -1.0)))
    positions, F = process(parse_export(shifted), Config())
    ys = [p.ref_pulse_d15n for p in positions if p.ref_pulse_d15n is not None]
    med = statistics.median(ys)
    rsd = ROBUST_SD * statistics.median(abs(y - med) for y in ys)
    assert math.isclose(F["pulse_off"]["n"]["median"], med)
    assert math.isclose(F["pulse_off"]["n"]["limit"], 4.0 * rsd)
    p20, p24 = (next(p for p in positions if p.row == r) for r in (20, 24))
    assert p20.rerun == "N" and p20.rerun_why.startswith("ran before the background settled; N2 reference pulse off by +")
    assert p24.rerun == "C" and "CO2 reference pulse off by -" in p24.rerun_why
    quiet, _ = process(parse_export(shifted), config_from_dict({"ref_pulse_off_sd": None}))
    assert not any("reference pulse off" in f for p in quiet for f in p.flags)
    assert next(p for p in quiet if p.row == 20).rerun == "Check"


def test_mass_window_by_hand():
    class P:
        area_n, amount = 20.0, 0.5
    assert _mass_window(P, "n", 10.0, 50.0) == (0.25, 1.25)
    assert _mass_window(P, "n", 60.0, 50.0) is None              # no mass reaches the range
    P.amount = None
    assert _mass_window(P, "n", 10.0, 50.0) is None


def test_suggested_mass_from_the_linearity_range(tmp_path):
    weighed = variant(tmp_path, "mass", masses=True)
    positions, F = process(parse_export(weighed), Config())
    L, mult = F["size_limits"], Config().min_blank_multiple

    def window(p, el):
        per_mg = getattr(p, f"area_{el}") / p.amount
        lo = max(L[f"min_{el}"], mult * F[f"blank_{el}"]["area"])
        return lo / per_mg, L[f"max_{el}"] / per_mg

    p29 = next(p for p in positions if p.row == 29)                 # C above the range
    (ln, hn), (lc, hc) = window(p29, "n"), window(p29, "c")
    assert p29.rerun == "C" and p29.rerun_mass == f"{max(ln, lc):.3f} to {min(hn, hc):.3f}"
    p23 = next(p for p in positions if p.row == 23)                 # no one mass suits both
    (ln, hn), (lc, hc) = window(p23, "n"), window(p23, "c")
    assert max(ln, lc) > min(hn, hc)
    assert p23.rerun_mass == f"N {ln:.3f} to {hn:.3f}; C {lc:.3f} to {hc:.3f} (two capsules)"
    assert all(p.rerun_mass is None for p in positions if p.rerun == "Check")


def test_problems_of_the_whole_run():
    _, F = process(parse_export(RUN), check_config(0.05))
    assert F["rerun"]["run_level"][0] == "check standard off by more than 0.05 permil at positions 2, 3, 19"
    _, F = process(parse_export(RUN), config_from_dict({"coverage_k": 0.5, "ref_gas_sd_limit": 0.05}))
    items = F["rerun"]["run_level"]
    assert any(t.startswith("anchors outside U, %N: ") for t in items)
    assert any(t.startswith("N2 reference pulses SD ") and t.endswith("at or above 0.05") for t in items)
    _, F = process(parse_export(RUN), config_from_dict({"ref_gas_sd_limit": None}))
    assert F["rerun"]["run_level_text"] == "none found"


def test_results_table_carries_the_mark_after_flags():
    positions, _ = process(parse_export(RUN), Config())
    text = results_csv(positions, Config())
    header = next(csv.reader(io.StringIO(text)))
    i = header.index("Flags")
    assert header[i + 1:i + 4] == ["Rerun", "Why", "Suggested mass (mg)"]
    rows = {r["Row"]: r for r in csv.DictReader(io.StringIO(text))}
    assert rows["53"]["Rerun"] == "N" and rows["53"]["Why"].startswith("N below linearity range")
    assert rows["11"]["Rerun"] == "" and rows["11"]["Suggested mass (mg)"] == ""


def test_configuration_round_trip_keeps_the_new_settings():
    assert Config().ref_pulse_off_sd == 4.0 and Config().ref_gas_sd_limit == 0.1
    cfg = config_from_dict({"ref_pulse_off_sd": 3.0, "ref_gas_sd_limit": None})
    back = config_from_dict(config_to_dict(cfg))
    assert back.ref_pulse_off_sd == 3.0 and back.ref_gas_sd_limit is None


CASES = ("default", "masses", "pulses", "off", "run_level", "check")


@pytest.fixture(scope="module")
def computed(tmp_path_factory):
    if not SOFFICE:
        pytest.skip("needs LibreOffice")
    folder = tmp_path_factory.mktemp("rerun")
    pulses = ((20, 1, "d 15N/14N", 1.0), (24, 6, "d 13C/12C", -1.0), (30, 1, "d 15N/14N", -0.8))
    cases = {
        "default": (RUN, Config()),
        "masses": (variant(folder, "masses", masses=True), Config()),
        "pulses": (variant(folder, "pulses", masses=True, shifts=pulses), Config()),
        "off": (variant(folder, "off", masses=True, shifts=pulses[:1]),
                config_from_dict({"ref_pulse_off_sd": None, "min_blank_multiple": None, "ref_gas_sd_limit": None})),
        "run_level": (variant(folder, "run_level", masses=True),
                      config_from_dict({"coverage_k": 0.5, "ref_gas_sd_limit": 0.05, "ref_pulse_off_sd": 2.0})),
        "check": (RUN, check_config(0.05)),
    }
    files = []
    for name, (export, cfg) in cases.items():
        out = folder / f"{name}.xlsx"
        build_workbook(export, out, cfg)
        files.append(out)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(folder / "calc")]
                   + [str(f) for f in files], check=True, capture_output=True, timeout=900)
    return folder, cases


def _close(mine, theirs):
    theirs = theirs if isinstance(theirs, (int, float)) and not isinstance(theirs, bool) else None
    if mine is None:
        return theirs is None
    return theirs is not None and abs(mine - theirs) <= 1e-9 * max(1.0, abs(mine))


@pytest.mark.parametrize("name", CASES)
def test_workbook_agrees_with_engine(computed, name):
    folder, cases = computed
    export, cfg = cases[name]
    positions, F = process(parse_export(export), cfg)
    wb = openpyxl.load_workbook(folder / "calc" / f"{name}.xlsx", data_only=True)
    rs, ft, qc, rr = wb["Results"], wb["Fits"], wb["QC"], wb["Rerun"]
    head = {str(rs.cell(4, c).value).split(",")[0]: c for c in range(1, rs.max_column + 1)}
    assert [rs.cell(4, c).value for c in range(17, 21)] == ["Flags", "Rerun", "Why", "Suggested mass (mg)"]
    for p in positions:
        i = 4 + p.row
        flags = {t.strip() for t in str(rs.cell(i, head["Flags"]).value or "").split(";") if t.strip()}
        assert flags == set(p.flags), (name, p.row, flags ^ set(p.flags))
        for h, attr in (("Rerun", "rerun"), ("Why", "rerun_why"), ("Suggested mass (mg)", "rerun_mass")):
            assert (getattr(p, attr) or "") == (rs.cell(i, head[h]).value or ""), (name, p.row, h)
        for h, attr in (("U(%N)", "U_pct_n"), ("u(d15N)", "u_d15n"), ("%N as measured", "pct_n_measured")):
            assert _close(getattr(p, attr), rs.cell(i, head[h]).value), (name, p.row, h)
    marked = [p for p in positions if p.rerun]
    assert [rr.cell(8 + k, 1).value for k in range(len(marked))] == [p.row for p in marked]
    assert rr.cell(8 + len(marked), 1).value in (None, "")
    for k, p in enumerate(marked):
        assert [rr.cell(8 + k, j).value or "" for j in (5, 6, 7)] == [p.rerun, p.rerun_why, p.rerun_mass or ""]
    for base, el in ((339, "n"), (343, "c")):
        po = F["pulse_off"][el]
        for off, key in ((0, "n"), (1, "median"), (2, "robust_sd"), (3, "limit")):
            mine, theirs = po[key], ft.cell(base + off, 2).value
            assert (theirs in (None, "")) if mine is None else _close(mine, theirs), (name, key)
    for row, key, el in ((347, "lo_area", "n"), (348, "hi_area", "n"), (349, "lo_area", "c"), (350, "hi_area", "c")):
        assert _close(F["rerun"][key][el], ft.cell(row, 2).value), (name, row)
    c = F["rerun"]["counts"]
    assert (qc["B168"].value, qc["B169"].value) == (c["N"] + c["C"] + c["N and C"], c["Check"])
    assert ft["B356"].value == rr["A5"].value == F["rerun"]["run_level_text"]
    errors = [x.coordinate for s in (rs, ft, qc, rr) for row in s.iter_rows() for x in row
              if isinstance(x.value, str) and x.value.startswith("#") and x.value[1:4].isupper()]
    assert not errors, errors[:5]


def test_rows_are_shaded_by_their_mark(computed):
    folder, _ = computed
    wb = openpyxl.load_workbook(folder / "default.xlsx")
    for sheet, col in (("Results", "$R5"), ("Rerun", "$E8")):
        rules = [r for rng in wb[sheet].conditional_formatting for r in rng.rules]
        assert {r.formula[0] for r in rules} == {f'{col}="Check"', f'AND({col}<>"",{col}<>"Check")'}
