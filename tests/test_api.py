"""The API against the engine it wraps, on the 2026-09-15 test run.

Every number the API returns must be the engine's own, unchanged; the workbook it
returns must hold the same values once LibreOffice has computed it; and a bad
upload must be refused with a message a person can act on.
"""
import csv
import io
import json
import math
import shutil
from pathlib import Path

import openpyxl
import pytest
from fastapi.testclient import TestClient

from ea_irms.parse import parse_export
from ea_irms.pipeline import Config, config_from_dict, config_to_dict, process
from ea_irms.workbook import VERSION
from isoharness import __version__
from isoharness.app import app
from isoharness.serialize import position_record

DATA = Path(__file__).resolve().parent / "data"
RUN = DATA / "20260915_raw.csv"
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
client = TestClient(app)


def post(path, name=RUN.name, content=None, **form):
    content = RUN.read_bytes() if content is None else content
    return client.post(path, files={"file": (name, content)}, data=form)


def same(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return a == b or (math.isnan(a) and math.isnan(b))
    return a == b


@pytest.fixture(scope="module")
def engine():
    return process(parse_export(RUN), Config())


def test_health_version_defaults():
    assert client.get("/api/health").json() == {"status": "ok"}
    v = client.get("/api/version").json()
    assert v["isoharness"] == __version__ and v["ea_irms"] == VERSION
    d = client.get("/api/defaults").json()
    assert d["config"] == json.loads(json.dumps(config_to_dict(Config())))


def test_home_page_and_the_ea_irms_page():
    home = client.get("/")
    assert home.status_code == 200 and home.headers["content-type"].startswith("text/html")
    assert "Stable isotope data reduction" in home.text
    assert 'href="ea-irms"' in home.text
    assert "GasBench N₂O" in home.text and "GasBench carbonate" in home.text
    assert home.text.count(">Planned<") == 2
    page = client.get("/ea-irms")
    assert page.status_code == 200 and page.headers["content-type"].startswith("text/html")
    assert page.content == (Path(__file__).resolve().parents[1] / "web" / "ea-irms.html").read_bytes()
    assert page.headers["cache-control"] == "no-cache"
    assert "script-src 'self'" in page.headers["content-security-policy"]
    assert client.head("/ea-irms").status_code == 200
    r = client.get("/ea-irms/", follow_redirects=False)
    assert r.status_code == 308 and r.headers["location"] == "/ea-irms"


def test_example_is_the_test_run():
    r = client.get("/api/example")
    assert r.status_code == 200
    assert r.headers["content-type"].startswith("text/csv")
    assert 'filename="20260915_raw.csv"' in r.headers["content-disposition"]
    assert r.content == RUN.read_bytes()


def test_process_returns_the_engine_values(engine):
    positions, fits = engine
    r = post("/api/process")
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["file"]["name"] == RUN.name and body["file"]["bytes"] == RUN.stat().st_size
    assert len(body["positions"]) == len(positions)
    compared = 0
    for got, p in zip(body["positions"], positions):
        want = json.loads(json.dumps(position_record(p)))
        assert got.keys() == want.keys()
        for k in want:
            assert same(got[k], want[k]), (p.row, k, got[k], want[k])
            compared += 1
    assert compared > 10 * len(positions)
    assert body["fits"]["cal_n"]["slope"] == fits["cal_n"]["slope"]
    assert body["fits"]["cal_c"]["intercept"] == fits["cal_c"]["intercept"]
    port = body["port"]
    assert [q["r"] for q in port["pos"]] == [p.row for p in positions]
    assert port["std"] == [list(s) for s in Config().standards]


def test_process_uses_the_configuration_sent():
    cfg = Config(drift_correction="Off", linearity_correction="On", u_sample_n=0.2)
    positions, _ = process(parse_export(RUN), cfg)
    r = post("/api/process", config=json.dumps(config_to_dict(cfg)))
    assert r.status_code == 200, r.text
    got = {q["row"]: q for q in r.json()["positions"]}
    for p in positions:
        assert same(got[p.row]["d15n_final"], p.d15n_final)
        assert same(got[p.row]["U_d15n"], p.U_d15n)
    assert config_from_dict(r.json()["config"]) == cfg


def test_results_csv_matches_the_engine(engine):
    positions, _ = engine
    r = post("/api/results.csv")
    assert r.status_code == 200 and r.headers["content-type"].startswith("text/csv")
    assert f"_results_ea_irms_v{VERSION}.csv" in r.headers["content-disposition"]
    rows = list(csv.DictReader(io.StringIO(r.text)))
    assert len(rows) == len(positions)
    for row, p in zip(rows, positions):
        assert int(row["Row"]) == p.row and row["Identifier"] == p.identifier
        for col, attr in (("d15N (permil)", "d15n_final"), ("d13C (permil)", "d13c_final"),
                          ("U(d15N), k=2.0", "U_d15n"), ("%N", "pct_n_reported"),
                          ("U(%N), k=2.0", "U_pct_n"), ("U(%C), k=2.0", "U_pct_c"),
                          ("U(C:N molar), k=2.0", "U_cn_molar"), ("u(C:N mass)", "u_cn_mass")):
            v = getattr(p, attr)
            assert (row[col] == "") == (v is None)
            if v is not None:
                assert float(row[col]) == v
        assert row["Flags"] == "; ".join(p.flags)


@pytest.mark.skipif(not SOFFICE, reason="needs LibreOffice")
def test_workbook_values_equal_the_engine(engine):
    positions, _ = engine
    r = post("/api/workbook")
    assert r.status_code == 200, r.text
    assert r.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert r.headers["x-isoharness-values"] == "computed"
    assert f'filename="20260915_raw_ea_irms_v{VERSION}.xlsx"' in r.headers["content-disposition"]
    rs = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)["Results"]
    head = {str(rs.cell(4, c).value): c for c in range(1, rs.max_column + 1)}
    compared = 0
    for p in positions:
        for col, attr in ((head["d13C (permil)"], "d13c_final"), (head["d15N (permil)"], "d15n_final"),
                          (head["U(d13C), k=2.0"], "U_d13c"), (head["U(d15N), k=2.0"], "U_d15n"),
                          (head["%C"], "pct_c_reported"), (head["%N"], "pct_n_reported"),
                          (head["U(%N), k=2.0"], "U_pct_n"), (head["U(C:N molar), k=2.0"], "U_cn_molar")):
            mine, theirs = getattr(p, attr), rs.cell(4 + p.row, col).value
            theirs = theirs if isinstance(theirs, (int, float)) else None
            assert (mine is None) == (theirs is None), (p.row, attr, mine, theirs)
            if mine is not None:
                assert abs(mine - theirs) <= 1e-9 * max(1.0, abs(mine)), (p.row, attr, mine, theirs)
                compared += 1
        flags = {t.strip() for t in (rs.cell(4 + p.row, head["Flags"]).value or "").split(";") if t.strip()}
        assert flags == set(p.flags), p.row
    assert compared > 3 * len(positions)


def test_workbook_without_values_and_with_a_run_name():
    r = post("/api/workbook", values="false", run_name="Test run", run_date="2026-09-15")
    assert r.status_code == 200 and r.headers["x-isoharness-values"] == "formulas-only"
    setup = openpyxl.load_workbook(io.BytesIO(r.content))["Setup"]
    assert setup["B5"].value == "Test run" and setup["B6"].value == "2026-09-15"


def test_refuses_an_unknown_file_type():
    r = post("/api/process", name="run.pdf")
    assert r.status_code == 415 and ".xls" in r.json()["detail"]


def test_refuses_an_empty_file():
    r = post("/api/process", content=b"")
    assert r.status_code == 422 and "empty" in r.json()["detail"]


def test_refuses_an_upload_over_the_limit(monkeypatch):
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "0.01")
    r = post("/api/process")
    assert r.status_code == 413 and "MB" in r.json()["detail"]


def test_names_the_missing_columns():
    text = RUN.read_text(encoding="utf-8", errors="replace").splitlines()
    head = text[0].split(",")
    keep = [i for i, h in enumerate(head) if h.strip() != "Area 28"]
    trimmed = "\n".join(",".join(c for i, c in enumerate(line.split(",")) if i in keep) for line in text)
    r = post("/api/process", content=trimmed.encode())
    assert r.status_code == 422 and "Area 28" in r.json()["detail"]


def test_refuses_an_export_with_no_positions():
    header = RUN.read_text(encoding="utf-8", errors="replace").splitlines()[0]
    for path in ("/api/process", "/api/workbook", "/api/results.csv"):
        r = post(path, content=(header + "\n").encode())
        assert r.status_code == 422 and "no positions" in r.json()["detail"], path


def test_refuses_a_file_that_is_not_what_its_name_says():
    r = post("/api/process", name="run.xlsx", content=b"not a spreadsheet")
    assert r.status_code == 422 and "could not be opened" in r.json()["detail"]


def test_refuses_a_bad_configuration():
    assert post("/api/process", config="{not json").status_code == 422
    assert post("/api/process", config="[1, 2]").status_code == 422


def test_refuses_more_standards_than_the_workbook_holds():
    cfg = config_to_dict(Config())
    cfg["standards"] = cfg["standards"] + cfg["standards"][:1] * 5
    r = post("/api/workbook", config=json.dumps(cfg), values="false")
    assert r.status_code == 422 and "standards" in r.json()["detail"]


def test_port_block_carries_the_settings_the_port_reads():
    cfg = config_to_dict(Config(significance_level=0.1, monitor_outlier_t=3.0, check_offset_permil=0.2,
                                min_blank_multiple=None, coverage_k=3.0, elemental_threshold_pct=2.0,
                                first_drift_position=20))
    r = post("/api/process", config=json.dumps(cfg))
    assert r.status_code == 200, r.text
    port = r.json()["port"]
    assert port["settings"] == {"alpha": 0.1, "outlierT": 3.0, "checkOffset": 0.2, "blankMultiple": None,
                                "k": 3.0, "elemGate": 2.0, "firstDrift": 20}
    # A full export: the reference-pulse check sets no CO2 peak aside.
    assert [q["rej"] for q in port["pos"]] == [False] * len(port["pos"])


def test_preview_port_is_the_checked_file():
    r = client.get("/engine.js")
    assert r.status_code == 200
    assert r.content == (Path(__file__).resolve().parents[1] / "port" / "engine.js").read_bytes()
    assert r.headers["content-type"].startswith("text/javascript")
    assert "script-src 'self'" in r.headers["content-security-policy"]
    assert r.headers["x-content-type-options"] == "nosniff"
    assert "content-security-policy" not in client.get("/api/health").headers


def test_upload_limit_per_address(monkeypatch):
    from isoharness.app import limiter
    monkeypatch.setenv("ISOHARNESS_UPLOADS_PER_HOUR", "2")
    limiter.reset()
    try:
        assert post("/api/results.csv").status_code == 200
        assert post("/api/process").status_code == 200
        r = post("/api/process")
        assert r.status_code == 429
        assert "2 uploads in the last hour" in r.json()["detail"]
        assert int(r.headers["retry-after"]) > 0
        assert client.get("/api/health").status_code == 200      # only uploads count
        assert limiter.wait_seconds("another address", 2) is None
    finally:
        limiter.reset()


def test_upload_limiter_window():
    from isoharness.app import UploadLimiter
    lim = UploadLimiter()
    assert lim.wait_seconds("a", 2, now=0.0) is None
    assert lim.wait_seconds("a", 2, now=10.0) is None
    assert lim.wait_seconds("a", 2, now=20.0) == 3581      # the first hit leaves the hour at 3600
    assert lim.wait_seconds("a", 2, now=3600.0) is None     # it has left
    assert lim.wait_seconds("a", 2, now=3601.0) == 10       # the one at 10.0 leaves at 3610

