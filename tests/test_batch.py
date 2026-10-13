"""isoharness 0.2.0: several runs at once, one combined results table.

Two runs are made from the 2026-09-15 test run. In both, each sample is named by its
position (S20, S21, ...), since the test run calls every sample Random; in the second,
each sample's N2 peak reads 0.5 permil higher. Every correction is then the same in the
two runs, so each sample's two final d15N values differ by 0.5 times the calibration
slope and its d13C values are equal, which is what the Repeats sheet must show.
"""
import csv
import io
import json
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from ea_irms.parse import parse_export
from ea_irms.pipeline import Config, config_to_dict, process
from ea_irms.workbook import VERSION
from isoharness.app import XLSX, app
from isoharness.batch import BatchRun, combined_csv, label_runs, repeats
from isoharness.serialize import csv_cells, results_values

RUN = Path(__file__).resolve().parent / "data" / "20260915_raw.csv"
client = TestClient(app)
SHIFT = 0.5


def variant(shift=0.0):
    """The test run with each sample named S<position> and its N2 peak's d15N moved by shift."""
    text = RUN.read_text(encoding="utf-8")
    rows = list(csv.reader(io.StringIO(text)))
    h = rows[0]
    col = {k: h.index(k) for k in ("Row", "Identifier 1", "Ampl  28", "Is Ref _", "d 15N/14N")}
    positions, _ = process(parse_export(RUN), Config())
    samples = {str(p.row) for p in positions if p.role == "Sample"}
    last28 = {}
    for i, r in enumerate(rows[1:], start=1):
        if len(r) > col["Ampl  28"] and r[col["Ampl  28"]].strip():
            last28[r[col["Row"]].strip()] = i
    for i, r in enumerate(rows[1:], start=1):
        row = r[col["Row"]].strip() if r else ""
        if row in samples:
            r[col["Identifier 1"]] = f"S{row}"
            if shift and last28.get(row) == i and r[col["Is Ref _"]].strip() != "1":
                r[col["d 15N/14N"]] = repr(float(r[col["d 15N/14N"]]) + shift)
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    return buf.getvalue().encode("utf-8")


@pytest.fixture(scope="module")
def pair(tmp_path_factory):
    d = tmp_path_factory.mktemp("batch")
    a, b = d / "run_a.csv", d / "run_b.csv"
    a.write_bytes(variant())
    b.write_bytes(variant(SHIFT))
    return a, b


def send(path, files, **form):
    return client.post(path, files=[("files", (Path(f).name, Path(f).read_bytes())) for f in files], data=form)


def test_combined_results_hold_every_position_of_every_run(pair):
    r = send("/api/batch/results.xlsx", pair, run_dates=json.dumps(["2026-09-15", None]))
    assert r.status_code == 200 and r.headers["content-type"] == XLSX
    assert r.headers["content-disposition"] == f'attachment; filename="combined_2_runs_results_ea_irms_v{VERSION}.xlsx"'
    wb = openpyxl.load_workbook(io.BytesIO(r.content))
    assert wb.sheetnames == ["README", "Results", "Runs", "Repeats"]
    rs = wb["Results"]
    header = [c.value for c in rs[1]]
    assert header[:4] == ["Run", "Row", "Identifier", "Type"] and "1σ(d15N)" in header and len(header) == 25
    got = [list(row) for row in rs.iter_rows(min_row=2, values_only=True)]
    want = []
    for f in pair:
        positions, _ = process(parse_export(f), Config())
        want += [[f.stem] + [v if v != "" else None for v in results_values(p)] for p in positions]
    assert len(got) == len(want) == 112
    for g, w in zip(got, want):
        for x, y in zip(g, w):
            same = x == y or (x in (None, "") and y in (None, ""))
            assert same or abs(x - y) <= 1e-12 * max(1, abs(y)), (g[:3], x, y)
    runs = list(wb["Runs"].iter_rows(values_only=True))
    assert [r[:4] for r in runs[1:]] == [("run_a", "run_a.csv", "2026-09-15", 56), ("run_b", "run_b.csv", None, 56)]


def test_repeats_show_each_sample_across_the_runs(pair):
    wb = openpyxl.load_workbook(io.BytesIO(send("/api/batch/results.xlsx", pair).content))
    rows = list(wb["Repeats"].iter_rows(values_only=True))
    head, body = rows[0], rows[1:]
    ix = {h: i for i, h in enumerate(head)}
    positions, fits = process(parse_export(pair[0]), Config())
    samples = [p for p in positions if p.role == "Sample"]
    assert len(body) == 2 * len(samples)
    half = SHIFT * fits["cal_n"]["slope"] / 2
    for row in body:
        assert row[ix["Runs"]] == 2 and row[ix["Analyses"]] == 2 and row[ix["Name"]] == row[ix["Identifier"]]
        if row[ix["d15N (permil)"]] is None:
            continue
        sign = -1 if row[ix["Run"]] == "run_a" else 1
        assert abs(row[ix["d15N minus mean"]] - sign * half) < 1e-9
        if row[ix["d13C minus mean"]] is not None:
            assert abs(row[ix["d13C minus mean"]]) < 1e-9
        assert abs(row[ix["SD d15N"]] - SHIFT * fits["cal_n"]["slope"] / 2 ** 0.5) < 1e-9


def test_the_combined_csv_is_the_results_table_with_a_run_column(pair):
    r = send("/api/batch/results.csv", pair)
    assert r.status_code == 200 and r.headers["content-type"].startswith("text/csv")
    lines = r.text.splitlines()
    assert lines[0].startswith("Run,Row,Identifier,Type,Mass (mg),%C,1sigma(%C)") and lines[0].isascii()
    assert len(lines) == 1 + 112
    one = client.post("/api/results.csv", files={"file": (pair[0].name, pair[0].read_bytes())}).text.splitlines()
    assert lines[1:57] == ["run_a," + line for line in one[1:]]


def test_each_run_takes_its_own_configuration(pair):
    off = config_to_dict(Config())
    on = dict(off, linearity_correction="On")
    r = send("/api/batch/results.xlsx", pair, configs=json.dumps([off, on]), config_name="lab.xlsx")
    runs = list(openpyxl.load_workbook(io.BytesIO(r.content))["Runs"].iter_rows(values_only=True))
    col = runs[0].index("Linearity applied")
    assert runs[1][col] == "C only" and runs[2][col] == "N and C"
    same = send("/api/batch/results.xlsx", pair, configs=json.dumps(on))
    runs = list(openpyxl.load_workbook(io.BytesIO(same.content))["Runs"].iter_rows(values_only=True))
    assert runs[1][col] == runs[2][col] == "N and C"
    readme = [c.value for c in openpyxl.load_workbook(io.BytesIO(r.content))["README"]["B"] if c.value]
    assert any("lab.xlsx" in t for t in readme)


def test_two_files_of_one_name_are_told_apart():
    runs = [BatchRun(file="20260915_raw.csv", positions=[], fits={}, config=Config()) for _ in range(3)]
    assert [r.label for r in label_runs(runs)] == ["20260915_raw", "20260915_raw (2)", "20260915_raw (3)"]
    assert [r.index for r in runs] == [1, 2, 3]


def test_a_sample_in_one_run_only_is_not_a_repeat(pair):
    positions, fits = process(parse_export(pair[0]), Config())
    runs = label_runs([BatchRun(file="a.csv", positions=positions, fits=fits, config=Config())])
    assert repeats(runs) == []
    text = combined_csv(runs)
    assert next(csv.reader(io.StringIO(text.splitlines()[1]))) == [str(v) for v in
                                                                   csv_cells(["a"] + results_values(positions[0]))]
    wb = openpyxl.load_workbook(io.BytesIO(send("/api/batch/results.xlsx", [RUN]).content))
    assert "No sample or check standard" in wb["Repeats"]["A1"].value


def test_refusals(pair, monkeypatch):
    r = client.post("/api/batch/results.xlsx", data={"configs": "{}"})
    assert r.status_code == 422
    r = send("/api/batch/results.xlsx", pair, configs=json.dumps([{}]))
    assert r.status_code == 422 and "one entry for each of the 2 files" in r.json()["detail"]
    r = send("/api/batch/results.xlsx", pair, configs="{not json")
    assert r.status_code == 422 and "configs could not be read" in r.json()["detail"]
    r = send("/api/batch/results.xlsx", pair, configs=json.dumps([{}, {"significance_level": "abc"}]))
    assert r.status_code == 422 and "run_b.csv" in r.json()["detail"]
    r = send("/api/batch/results.xlsx", pair, configs=json.dumps([{}, ["not", "an", "object"]]))
    assert r.status_code == 422 and r.json()["detail"] == "the configuration for run_b.csv is not a JSON object"
    r = client.post("/api/batch/results.xlsx", files=[("files", ("notes.pdf", b"%PDF"))])
    assert r.status_code == 415
    monkeypatch.setenv("ISOHARNESS_MAX_BATCH_RUNS", "1")
    r = send("/api/batch/results.xlsx", pair)
    assert r.status_code == 422 and "at most 1 runs" in r.json()["detail"]
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "0.03")
    monkeypatch.setenv("ISOHARNESS_MAX_BATCH_RUNS", "20")
    r = send("/api/batch/results.csv", pair)
    assert r.status_code == 413 and "together" in r.json()["detail"]


def test_batch_requests_count_as_uploads(pair, monkeypatch):
    from isoharness.app import limiter
    monkeypatch.setenv("ISOHARNESS_READS_PER_HOUR", "1")
    limiter.reset()
    try:
        assert send("/api/batch/results.csv", pair).status_code == 200
        assert send("/api/batch/results.xlsx", pair).status_code == 429
    finally:
        limiter.reset()
