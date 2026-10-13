"""The statistics behind the correction decisions and the uncertainty budget.

The p-values are checked against the numbers Excel's Data Analysis tool printed
in the lab's own correction sheet for the 2026-09-22 run, and the closed-form
uncertainty weights against a numerical perturbation of the calibration fit,
which is what a Kragten spreadsheet does by hand.
"""
import csv
import math
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ea_irms.parse import parse_export                       # noqa: E402
from ea_irms.pipeline import Config, _fit, process, t_p_two_sided  # noqa: E402

HEADER = ["Row", "Identifier 1", "Identifier 2", "Peak Nr", "Amount", "Ampl  28",
          "BGD 28", "Area 28", "Is Ref _", "Ampl  44", "Area 44",
          "Peak Nr  Flash TCD", "rArea Flash TCD", "d 15N/14N", "d 13C/12C",
          "BGD Flash TCD"]


def _sequence(tmp_path, drift_c_last=-13.40, extra=()):
    """A small filtered-export sequence: anchors, a size series, six drift monitors."""
    rows = []
    def add(row, ident, amount, a28, area28, d15, a44, area44, d13):
        rows.append([row, ident, "", 3, amount, a28, "23.0", area28, 0, "", "", "", "", d15, "", ""])
        rows.append([row, ident, "", 4, amount, "", "", "", 0, a44, area44, "", "", "", d13, ""])
    lin = [(0.125, 457, 7.9, -0.20, 1679, 37.8, -13.36), (0.245, 959, 16.4, -0.28, 3430, 78.4, -13.39),
           (0.621, 2525, 42.2, -0.26, 8401, 202.5, -13.31), (0.760, 3113, 51.7, -0.07, 10001, 247.4, -13.27),
           (0.970, 4038, 66.4, 0.02, 12327, 316.6, -13.25)]
    r = 1
    for k, (amt, a, ar, d, ac, arc, dc) in enumerate(lin, start=1):
        add(r, f"Ace Lin{k}", amt, a, ar, d, ac, arc, dc); r += 1
    for d15, d13 in ((1.29, -23.77), (1.35, -23.71), (1.16, -23.79)):
        add(r, "USGS 64", 0.41, 3000, 51.0, d15, 2700, 61.0, d13); r += 1
    for d15, d13 in ((20.02, -3.33), (19.87, -3.38), (19.96, -3.34)):
        add(r, "USGS 65", 0.41, 3100, 52.0, d15, 2750, 63.0, d13); r += 1
    drift = [(-0.10, -13.38), (-0.16, -13.36), (-0.09, -13.39), (-0.09, -13.37), (-0.36, -13.43),
             (-0.33, drift_c_last)]
    for k, (d15, d13) in enumerate(drift, start=1):
        add(r, f"Ace Drift{k}", 0.42, 1650, 28.0, d15, 5700, 132.0, d13); r += 6
    for ident, d15, d13, a28, ar28, a44, ar44 in extra:
        add(r, ident, 5.0, a28, ar28, d15, a44, ar44, d13); r += 1
    path = tmp_path / "seq.csv"
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh); w.writerow(HEADER); w.writerows(rows)
    return path


def test_p_value_matches_excel():
    # Excel TDIST(|t|, 4, 2) as printed in the lab's regression output
    assert abs(t_p_two_sided(-2.1788387694199747, 4) - 0.09486756775340509) < 1e-12
    assert abs(t_p_two_sided(-1.8895774100595444, 4) - 0.13181484263871995) < 1e-12
    assert abs(t_p_two_sided(0.0, 10) - 1.0) < 1e-12
    assert t_p_two_sided(50.0, 3) < 1e-4


def test_fit_carries_slope_uncertainty():
    f = _fit([1, 2, 3, 4, 5], [2.0, 4.1, 5.9, 8.2, 9.8])
    assert f["df"] == 3 and f["se"] > 0 and 0 < f["p"] < 0.001
    assert abs(f["se"] - f["sd"] / math.sqrt(f["sxx"])) < 1e-12


def test_auto_applies_only_a_significant_slope(tmp_path):
    positions, fits = process(parse_export(_sequence(tmp_path)), Config())
    assert fits["lin_c"]["in_force"] is True          # a clear size effect in carbon
    assert fits["lin_n"]["in_force"] is False         # nitrogen series too scattered
    assert fits["lin_n"]["applied"] == 0.0
    positions, fits = process(parse_export(_sequence(tmp_path)),
                              Config(linearity_correction="On"))
    assert fits["lin_n"]["in_force"] is True and fits["lin_n"]["applied"] != 0.0


def test_one_bad_monitor_is_flagged_not_fitted(tmp_path):
    positions, fits = process(parse_export(_sequence(tmp_path, drift_c_last=-14.85)), Config())
    last = [p for p in positions if p.identifier == "Ace Drift6"][0]
    assert abs(last.t_drift_c) > 10
    assert "C drift monitor outlier" in last.flags
    assert fits["drift_c"]["outliers"] == [last.row]
    positions, fits = process(parse_export(_sequence(tmp_path, drift_c_last=-14.85)),
                              Config(exclude_drift_c=[last.row], drift_correction="On"))
    last = [p for p in positions if p.identifier == "Ace Drift6"][0]
    assert "excluded from C drift fit" in last.flags
    assert fits["drift_c"]["n"] == 5 and abs(fits["drift_c"]["applied"]) < 0.01
    assert fits["drift_n"]["n"] == 6                  # the other element keeps it


def test_uncertainty_matches_numerical_perturbation(tmp_path):
    extra = [("S1", 5.0, -18.0, 3000, 50.0, 6000, 140.0), ("S2", 12.0, -12.0, 3000, 50.0, 6000, 140.0)]
    positions, fits = process(parse_export(_sequence(tmp_path, extra=extra)), Config())
    anchors = [p for p in positions if p.role == "Anchor"]
    for xattr, accattr, uattr, dst, key in (("d15n_drift", "acc15", "u_acc15", "u_d15n", "unc_n"),
                                            ("d13c_drift", "acc13", "u_acc13", "u_d13c", "unc_c")):
        xs = [getattr(p, xattr) for p in anchors]
        ys = [getattr(p, accattr) for p in anchors]
        base = _fit(xs, ys)
        u_s = fits[key]["u_sample"]
        for s in positions:
            x = getattr(s, xattr)
            if x is None:
                continue
            y0 = base["slope"] * x + base["intercept"]
            terms = []
            for m in {p.std_index for p in anchors}:
                eps = 1e-6
                u_acc = [getattr(p, uattr) for p in anchors if p.std_index == m][0] or 0.0
                ys2 = [y + (eps if p.std_index == m else 0) for y, p in zip(ys, anchors)]
                f2 = _fit(xs, ys2)
                terms.append((f2["slope"] * x + f2["intercept"] - y0) / eps * u_acc)
                g = [v for v, p in zip(xs, anchors) if p.std_index == m]
                sd = math.sqrt(sum((v - sum(g) / len(g)) ** 2 for v in g) / (len(g) - 1))
                xs2 = [v + (eps if p.std_index == m else 0) for v, p in zip(xs, anchors)]
                f3 = _fit(xs2, ys)
                terms.append((f3["slope"] * x + f3["intercept"] - y0) / eps * sd / math.sqrt(len(g)))
            terms.append(base["slope"] * u_s)
            assert abs(math.sqrt(sum(t * t for t in terms)) - getattr(s, dst)) < 1e-8
        assert fits["coverage_k"] == 2.0
        assert abs(positions[0].U_d15n - 2.0 * positions[0].u_d15n) < 1e-12


def test_range_flags_and_missing_check(tmp_path):
    extra = [("Big", 6.0, 5.0, 13000, 220.0, 24000, 775.0), ("Tiny", 1.0, -30.0, 134, 3.2, 500, 11.0)]
    positions, fits = process(parse_export(_sequence(tmp_path, extra=extra)), Config())
    big = [p for p in positions if p.identifier == "Big"][0]
    tiny = [p for p in positions if p.identifier == "Tiny"][0]
    assert "N above linearity range" in big.flags and "C above linearity range" in big.flags
    assert "d13C outside anchor range" in big.flags          # raw +5 calibrates heavier than USGS65
    assert "N below linearity range" in tiny.flags
    assert "d15N outside anchor range" in tiny.flags         # below USGS64
    assert fits["checks"]["n"] == 0                          # nothing independent in this run
    assert not any("outside anchor range" in f for p in positions if p.role == "Drift" for f in p.flags)


def test_reference_pulses_absent_from_a_filtered_export(tmp_path):
    positions, fits = process(parse_export(_sequence(tmp_path)), Config())
    assert fits["ref_gas_n"]["n"] == 0 and fits["ref_gas_n"]["trend"] is False
    assert all(p.ref_pulse_d15n is None for p in positions)


def test_old_boolean_switches_still_load():
    from ea_irms.pipeline import config_from_dict, config_to_dict
    c = config_from_dict({"linearity_correction": True, "drift_correction": False,
                          "standards": [["USGS 64", "Anchor", -40.81, 1.76, 32, 18.66, "glycine"]],
                          "exclude_drift_c": "59, 61"})
    assert c.linearity_correction == "On" and c.drift_correction == "Off"
    assert len(c.standards[0]) == 11 and c.standards[0][7] is None and c.standards[0][10] is None
    assert c.exclude_drift_c == [59, 61]
    assert config_to_dict(c)["exclude_drift_c"] == [59, 61]
