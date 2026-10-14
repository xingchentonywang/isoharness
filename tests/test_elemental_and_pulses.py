"""ea_irms 1.4.2: elemental drift from the drift monitors, reference pulses read as
sample CO2 in exports without reference peaks, settling against the run's level,
the Acetanilide default rows, the five-monitor floor on the outlier test, and the
standards-table limit. The engine is checked against known inputs, and the
workbook, computed by LibreOffice, against the engine.
"""
import csv
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from ea_irms.parse import parse_export                                                # noqa: E402
from ea_irms.pipeline import (Config, DEFAULT_STANDARDS, MAX_STANDARDS, config_from_dict,  # noqa: E402
                              config_to_dict, process)
from ea_irms.workbook import build_workbook                                           # noqa: E402
from test_statistics import HEADER, _sequence                                         # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def _ramped(tmp_path, ramp=0.10, exclude_bad=False, bad_last=False, acetanilide=False):
    """The synthetic sequence with a sensitivity that falls linearly through the run.

    Every peak area and amplitude at position r is scaled by 1 - ramp * (r - 1) / 40,
    so the monitors read low in %N and %C by position, and the samples with them.
    """
    base = _sequence(tmp_path, extra=(("sample A", 5.0, -20.0, 2000, 33.0, 6000, 140.0),
                                      ("sample B", 8.0, -18.0, 2200, 36.0, 6500, 150.0)))
    rows = list(csv.reader(open(base, newline="")))
    out = [rows[0]]
    for rec in rows[1:]:
        if rec[1] == "sample A":
            rec[0] = "15"                                # between Drift1 (12) and Drift2 (18)
        r = int(rec[0])
        f = 1 - ramp * (r - 1) / 40
        if bad_last and rec[1] == "Ace Drift6":
            f *= 1.25                                   # a capsule twice as heavy as entered
        for j in (5, 7, 9, 10):
            if rec[j] not in ("", None):
                rec[j] = f"{float(rec[j]) * f:.4f}"
        if acetanilide:
            rec[1] = rec[1].replace("Ace ", "Acetanilide ")
        out.append(rec)
    path = tmp_path / "ramped.csv"
    with open(path, "w", newline="") as fh:
        csv.writer(fh).writerows(out)
    return path


def _by(positions):
    return {p.identifier: p for p in positions}


def test_correction_recovers_the_monitors_and_moves_the_samples(tmp_path):
    positions, fits = process(parse_export(_ramped(tmp_path)), Config(elemental_correction="On"))
    by = _by(positions)
    assert fits["elem_n"]["in_force"] and fits["elem_c"]["in_force"]
    assert fits["elem_n"]["n"] == 6 and fits["elem_n"]["max_bias"] > 0.02
    for k in range(1, 7):
        m = by[f"Ace Drift{k}"]
        assert abs(m.pct_n_reported - 10.36) < 1e-9 and abs(m.pct_c_reported - 71.09) < 1e-9
        assert abs(m.pct_n_measured / m.elem_factor_n - m.pct_n_reported) < 1e-12
    # between two monitors the factor is interpolated in position; beyond the last it is held
    rows = [p.row for p in positions if p.role == "Drift"]
    a, b = by["sample A"], by["sample B"]
    assert rows[0] < a.row < rows[1] and b.row > rows[-1]
    assert b.elem_factor_n == by["Ace Drift6"].elem_factor_n
    f0, f1 = by["Ace Drift1"].elem_factor_n, by["Ace Drift2"].elem_factor_n
    expect = f0 + (a.row - rows[0]) / (rows[1] - rows[0]) * (f1 - f0)
    assert abs(a.elem_factor_n - expect) < 1e-12 and min(f0, f1) < a.elem_factor_n < max(f0, f1)
    assert abs(a.cn_molar_reported - a.cn_molar * a.elem_factor_n / a.elem_factor_c) < 1e-12


def test_off_and_auto_switches(tmp_path):
    off, f_off = process(parse_export(_ramped(tmp_path)), Config(elemental_correction="Off"))
    assert f_off["elem_n"]["in_force"] is False
    assert all(p.elem_factor_n == 1.0 and p.elem_factor_c == 1.0 for p in off)
    assert all(p.pct_n_reported == p.pct_n_measured for p in off if p.pct_n_measured is not None)
    _, f_auto = process(parse_export(_ramped(tmp_path, ramp=0.02)), Config())
    assert f_auto["elem_n"]["max_bias"] * 100 < 5.0 and f_auto["elem_n"]["in_force"] is False
    _, f_auto = process(parse_export(_ramped(tmp_path, ramp=0.10)), Config())
    assert f_auto["elem_n"]["max_bias"] * 100 > 5.0 and f_auto["elem_n"]["in_force"] is True
    _, f_thr = process(parse_export(_ramped(tmp_path, ramp=0.10)), Config(elemental_threshold_pct=50.0))
    assert f_thr["elem_n"]["in_force"] is False


def test_excluded_monitor_is_left_out(tmp_path):
    seq = _ramped(tmp_path, bad_last=True)
    with_bad, f1 = process(parse_export(seq), Config(elemental_correction="On"))
    last_row = _by(with_bad)["Ace Drift6"].row
    without, f2 = process(parse_export(seq), Config(elemental_correction="On", exclude_elemental=[last_row]))
    assert f1["elem_n"]["n"] == 6 and f2["elem_n"]["n"] == 5 and f2["elem_n"]["excluded"] == [last_row]
    bad = _by(without)["Ace Drift6"]
    assert "excluded from elemental correction" in bad.flags
    assert bad.elem_bias_n > 0.08                      # still reported, not used (the k-factor fit absorbs part of the 25%)
    b = _by(without)["sample B"]                       # beyond the last monitor used, Drift5
    assert abs(b.elem_factor_n - _by(without)["Ace Drift5"].elem_factor_n) < 1e-12
    assert _by(with_bad)["sample B"].pct_n_reported < b.pct_n_reported


def test_pulse_read_as_sample_is_rejected(tmp_path):
    """A filtered export: at positions with no sample gas the two CO2 pulses are exported, the
    second flagged; at a position with N2 but no CO2 the first pulse is exported unflagged."""
    rows = [list(r) for r in csv.reader(open(_sequence(tmp_path), newline=""))]
    nxt = max(int(r[0]) for r in rows[1:]) + 1
    def pulse(row, ident, pk, ampl, flagged, d13):
        return [row, ident, "", pk, 0, "", "", "", 1 if flagged else 0, ampl, round(ampl * 0.0192, 3), "", "", "", d13, ""]
    rows.append(pulse(nxt, "Blank Capsule", 3, 1240, False, -24.98))
    rows.append(pulse(nxt, "Blank Capsule", 4, 1251, True, -25.22))
    rows.append(pulse(nxt + 1, "Blank Capsule", 3, 1238, False, -25.05))
    rows.append(pulse(nxt + 1, "Blank Capsule", 4, 1249, True, -25.22))
    rows.append([nxt + 2, "no CO2 sample", "", 3, 1.0, 2000, "23.0", 33.0, 0, "", "", "", "", 5.0, "", ""])
    rows.append(pulse(nxt + 2, "no CO2 sample", 4, 1245, False, -25.01))
    rows.append([nxt + 3, "small sample", "", 3, 1.0, 2000, "23.0", 33.0, 0, "", "", "", "", 5.0, "", ""])
    rows.append([nxt + 3, "small sample", "", 4, 1.0, "", "", "", 0, 1246, 24.0, "", "", "", -2.5, ""])
    path = tmp_path / "pulses.csv"
    with open(path, "w", newline="") as fh:
        csv.writer(fh).writerows(rows)
    positions, fits = process(parse_export(path), Config())
    by = {p.row: p for p in positions}
    check = fits["ref_pulse_check"]
    assert check["filtered"] and check["active"] and check["n_designated"] == 2
    assert check["rejected"] == [nxt, nxt + 1, nxt + 2]
    assert by[nxt].area_c is None and by[nxt].d13c is None and by[nxt].c_rejected
    assert "CO2 peak rejected as a reference pulse" in by[nxt + 2].flags and "no C peak" in by[nxt + 2].flags
    assert by[nxt + 2].d15n_final is not None                 # its nitrogen is untouched
    assert by[nxt + 3].d13c_final is not None and not by[nxt + 3].c_rejected   # same size, sample d13C
    assert fits["blank_c"]["n"] == 0                            # so no carbon blank correction from pulses
    off, f_off = process(parse_export(path), Config(ref_pulse_ampl_tol=None))
    assert f_off["ref_pulse_check"]["active"] is False and {p.row: p for p in off}[nxt].d13c is not None


def test_settled_when_the_start_sits_at_the_run_level(tmp_path):
    """A full export whose background is flat from the start but ends lower: before 1.4.2 the floor
    came from the run's last readings and the run counted as never settled."""
    rows = [list(r) for r in csv.reader(open(_sequence(tmp_path), newline=""))]
    out = [rows[0]]
    seen = set()
    for rec in rows[1:]:
        r = int(rec[0])
        if r not in seen:
            seen.add(r)
            bgd = 22.0 if r < 18 else 20.8                       # flat, then a lower tail
            out.append([r, rec[1], "", 1, rec[4], 4800, f"{bgd}", 95.0, 1, "", "", "", "", 0.0, "", ""])
        out.append(rec)
    path = tmp_path / "level.csv"
    with open(path, "w", newline="") as fh:
        csv.writer(fh).writerows(out)
    _, fits = process(parse_export(path), Config())
    s = fits["settling"]
    assert s["excess"] > 1.0 and abs(s["excess_level"]) < 1e-9 and s["settled_from"] == 1


def test_acetanilide_rows_match_by_default(tmp_path):
    positions, fits = process(parse_export(_ramped(tmp_path, acetanilide=True)), Config())
    roles = {p.identifier: p.role for p in positions}
    assert roles["Acetanilide Lin1"] == "Linearity" and roles["Acetanilide Drift3"] == "Drift"
    assert fits["lin_c"]["n"] == 5 and fits["drift_n"]["n"] == 6


def test_outlier_test_needs_five_monitors(tmp_path):
    four = [list(r) for r in csv.reader(open(_sequence(tmp_path, drift_c_last=-14.85), newline=""))
            if r[1] not in ("Ace Drift5",)]
    path = tmp_path / "four.csv"
    with open(path, "w", newline="") as fh:
        csv.writer(fh).writerows(four)
    positions, fits = process(parse_export(path), Config())
    assert fits["drift_c"]["n"] == 5 and fits["drift_c"]["outliers"] == [[p.row for p in positions if p.identifier == "Ace Drift6"][0]]
    five = [r for r in four if r[1] not in ("Ace Drift4",)]
    with open(path, "w", newline="") as fh:
        csv.writer(fh).writerows(five)
    positions, fits = process(parse_export(path), Config())
    assert fits["drift_c"]["n"] == 4 and fits["drift_c"]["outliers"] == []
    assert all(p.t_drift_c is None for p in positions if p.role == "Drift")


def test_standards_table_limit(tmp_path):
    assert len(DEFAULT_STANDARDS) <= MAX_STANDARDS
    cfg = Config(standards=list(DEFAULT_STANDARDS) + [(f"extra {k}", "Check", -10.0, 5.0, 40.0, 9.0, "x", None, None)
                                                       for k in range(MAX_STANDARDS - len(DEFAULT_STANDARDS) + 1)])
    with pytest.raises(ValueError, match="standards table"):
        build_workbook(_sequence(tmp_path), tmp_path / "too_many.xlsx", cfg)


def test_new_fields_survive_the_config_file():
    cfg = Config(elemental_correction="On", elemental_threshold_pct=7.5, exclude_elemental=[11, 58],
                 ref_pulse_ampl_tol=0.03, ref_pulse_d13c_window=0.5)
    back = config_from_dict(json.loads(json.dumps(config_to_dict(cfg))))
    assert back.elemental_correction == "On" and back.elemental_threshold_pct == 7.5
    assert back.exclude_elemental == [11, 58] and back.ref_pulse_ampl_tol == 0.03 and back.ref_pulse_d13c_window == 0.5
    assert config_from_dict({"exclude_elemental": "11, 58"}).exclude_elemental == [11, 58]
    old = config_from_dict({"linearity_correction": "Auto"})            # a 1.4.1 file
    assert old.elemental_correction == "Auto" and old.exclude_elemental == [] and old.ref_pulse_ampl_tol == 0.02


@pytest.mark.skipif(not SOFFICE, reason="needs LibreOffice")
def test_workbook_agrees_with_engine_on_the_new_columns(tmp_path):
    seq = _ramped(tmp_path, bad_last=True)
    last_row = [p.row for p in parse_export(seq) if p.identifier == "Ace Drift6"][0]
    cfg = Config(elemental_correction="On", exclude_elemental=[last_row])
    out = tmp_path / "ramped.xlsx"
    build_workbook(seq, out, cfg)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp_path / "calc"),
                    str(out)], check=True, capture_output=True, timeout=600)
    wb = openpyxl.load_workbook(tmp_path / "calc" / "ramped.xlsx", data_only=True)
    wk, ft, rs = wb["Work"], wb["Fits"], wb["Results"]
    head = {str(wk.cell(5, c).value): c for c in range(1, wk.max_column + 1)}
    rhead = {str(rs.cell(4, c).value): c for c in range(1, rs.max_column + 1)}
    positions, fits = process(parse_export(seq), cfg)
    for attr, name in (("elem_factor_n", "elem factor N"), ("elem_factor_c", "elem factor C"),
                       ("pct_n_reported", "%N corrected"), ("pct_c_reported", "%C corrected"),
                       ("cn_molar_reported", "C:N molar corrected"), ("elem_bias_n", "elem bias N")):
        for p in positions:
            mine = getattr(p, attr)
            theirs = wk.cell(5 + p.row, head[name]).value
            theirs = theirs if isinstance(theirs, (int, float)) else None
            assert (mine is None) == (theirs is None), (attr, p.row, mine, theirs)
            if mine is not None:
                assert abs(mine - theirs) < 1e-9 * max(1.0, abs(mine)), (attr, p.row, mine, theirs)
    assert ft["B246"].value == "Yes" and ft["B249"].value == "Yes" and ft["B244"].value == 5
    assert abs(ft["B245"].value - fits["elem_n"]["max_bias"] * 100) < 1e-9
    for p in positions:
        for col, attr in ((rhead["%N"], "pct_n_reported"), (rhead["%C"], "pct_c_reported"),
                          (rhead["Elemental drift factor N"], "elem_factor_n"), (rhead["%N as measured"], "pct_n_measured"),
                          (rhead["1σ(%N)"], "u_pct_n"), (rhead["1σ(C:N molar)"], "u_cn_molar")):
            mine, theirs = getattr(p, attr), rs.cell(4 + p.row, col).value
            theirs = theirs if isinstance(theirs, (int, float)) else None
            assert (mine is None) == (theirs is None) and (mine is None or abs(mine - theirs) < 1e-9 * max(1.0, abs(mine)))
        text = rs.cell(4 + p.row, rhead["Flags"]).value or ""
        assert p.flags == [t.strip() for t in text.split(";") if t.strip()], p.row
