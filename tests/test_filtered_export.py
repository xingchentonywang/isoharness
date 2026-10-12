"""Exports filtered to Peak Nr 3 and 4 carry no reference peaks.

The lab's correction workbooks paste an export already cut down to the sample
peaks. The workbook formulas read carbon from it; before 1.2.1 the parser did
not, so the two code paths disagreed on every carbon value. Values below are
taken from positions of a real filtered export (2026-09-22).
"""
import csv

from ea_irms.parse import parse_export

HEADER = ["Row", "Identifier 1", "Identifier 2", "Peak Nr", "Amount", "Ampl  28",
          "BGD 28", "Area 28", "Is Ref _", "Ampl  44", "Area 44",
          "Peak Nr  Flash TCD", "rArea Flash TCD", "d 15N/14N", "d 13C/12C",
          "BGD Flash TCD"]
ROWS = [
    # blank: peak 3 is the flagged CO2 reference, peak 4 the next pulse
    [4, "Blank Capsule - small", "0.000", 3, 0, "", "", "", 1, 4840, 91.31, "", "", "", -25.22, ""],
    [4, "Blank Capsule - small", "0.000", 4, 0, "", "", "", 0, 4826, 91.422, "", "", "", -25.503, ""],
    # ordinary sample: N on peak 3, C on peak 4, neither flagged
    [6, "Ace Lin1", "0.125", 3, 0.125, 457, "23.1", 7.91, 0, "", "", "", "", -0.196, "", ""],
    [6, "Ace Lin1", "0.125", 4, 0.125, "", "", "", 0, 1679, 37.778, "", "", "", -13.364, ""],
    # carbon with no nitrogen: CO2 on peak 3, reference pulse on peak 4
    [9, "carbonate", "1.0", 3, 1.0, "", "", "", 0, 3000, 70.0, "", "", "", 1.5, ""],
    [9, "carbonate", "1.0", 4, 1.0, "", "", "", 0, 4850, 91.0, "", "", "", -25.4, ""],
]


def _write(tmp_path):
    path = tmp_path / "filtered.csv"
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(HEADER)
        w.writerows(ROWS)
    return path


def test_sample_carbon_is_read(tmp_path):
    p = {q.row: q for q in parse_export(_write(tmp_path))}[6]
    assert p.d15n == -0.196 and p.area_n == 7.91
    assert p.d13c == -13.364 and p.area_c == 37.778


def test_blank_reference_pulse_is_not_taken_as_sample(tmp_path):
    p = {q.row: q for q in parse_export(_write(tmp_path))}[4]
    assert p.d13c is None and p.d15n is None


def test_first_co2_peak_wins_when_nitrogen_is_absent(tmp_path):
    p = {q.row: q for q in parse_export(_write(tmp_path))}[9]
    assert p.d15n is None
    assert p.d13c == 1.5 and p.area_c == 70.0
