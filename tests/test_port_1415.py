"""ea_irms 1.4.15 in the preview port (port/engine.js), held to the engine on the 2026-09-15 test run.

The port finds the floor the size series sets from its own blank step, puts the Kragten terms of the corrections in
force and the drift monitors' SD into the 1 sigma, holds every acetanilide analysis against that 1 sigma, and holds
each drift monitor against an accepted value of the standards table. scripts/port_cases.py runs each under the
settings variants that switch them (and with accepted values on the Drift rows), and two of its ways show what the
test run as exported cannot: --no-series gives the Linearity rows of the standards table the role Exclude, so the
run has no size series and its peaks are held to the lab minimum, and --long-series renames ten samples to linearity
standards, so the walk up the series that sets the floor takes many steps. On each the port must equal the engine
in every value, 1 sigma, Kragten term, flag list, and fit, and the cases must reach what they are there for.
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from ea_irms.parse import parse_export                       # noqa: E402
from ea_irms.pipeline import Config, process                 # noqa: E402
from isoharness.serialize import port_input                  # noqa: E402

RUN = ROOT / "tests" / "data" / "20260915_raw.csv"
NODE = shutil.which("node")


def cases_of(tmp_path, *args):
    out = tmp_path / "port"
    gen = subprocess.run([sys.executable, str(ROOT / "scripts" / "port_cases.py"), str(RUN), "--out", str(out),
                          "--random", "40", "--per-variant", "3", *args], capture_output=True, text=True, cwd=ROOT)
    assert gen.returncode == 0, gen.stderr
    run = subprocess.run([NODE, str(ROOT / "port" / "test_port.js"), str(out / "port_input.json"),
                          str(out / "cases.json")], capture_output=True, text=True, cwd=ROOT)
    assert run.returncode == 0, run.stdout[-2000:] + run.stderr[-2000:]
    assert "value mismatches 0, flag-list mismatches 0, fit mismatches 0" in run.stdout, run.stdout[-500:]
    return json.loads((out / "cases.json").read_text())


def flags(cases, keep=lambda c: True):
    return [f for c in cases if keep(c) for r in c["rows"] for f in r[-1]]


def test_the_block_carries_what_the_port_needs():
    cfg = Config(floor_offset_permil=None, floor_sd_multiple=2.0, min_area_n=10.0, no_series_min_area_c=30.0,
                 kragten_corrections="Off", monitor_repeatability="Off", ace_scatter_check="Off")
    positions, fits = process(parse_export(RUN), cfg)
    s = port_input(positions, fits, cfg, RUN.name, 1)["settings"]
    assert (s["floorOffset"], s["floorSdMultiple"], s["minAreaN"], s["minAreaC"]) == (None, 2.0, 10.0, None)
    assert (s["noSeriesMinN"], s["noSeriesMinC"]) == (None, 30.0)
    assert (s["kragten"], s["monitorRep"], s["aceCheck"]) == ("Off", "Off", "Off")


@pytest.mark.skipif(not NODE, reason="needs node for the preview port")
def test_the_port_agrees_on_the_run_as_exported(tmp_path):
    cases = cases_of(tmp_path)
    fits = [c["fits"] for c in cases]
    # the walk moves the nitrogen floor to the second standard, and the flags name it
    assert {f["floorN"][1] for f in fits} >= {"series floor", "smallest standard", "configuration"}
    assert "N below series floor" in flags(cases)
    # the drift monitors' SD in use and not, the acetanilide check on and off, each Kragten term, and off
    assert {f["monN"][6] for f in fits} == {True, False} and {f["aceN"][5] for f in fits} == {True, False}
    assert any(f["aceN"][4] is None for f, c in zip(fits, cases) if c["case"].get("settings", {}).get("aceCheck") == "Off")
    assert {k for f in fits for k in f["corrN"][3]} == {"blank", "linearity", "drift"}
    assert all(f["corrN"][3] == [] for f, c in zip(fits, cases) if c["case"].get("settings", {}).get("kragten") == "Off")
    # the drift monitors against accepted values given to the Drift rows, which the run as exported has none of
    with_acc = flags(cases, lambda c: "std" in c["case"])
    assert any(f.startswith("monitor d15N off by ") for f in with_acc)
    assert any(f.startswith("monitor d13C off by ") for f in with_acc)
    assert not any(f.startswith("monitor ") for f in flags(cases, lambda c: "std" not in c["case"]))
    # every sample with a final value is marked when the acetanilide scatters
    assert "acetanilide d15N scatter above 1σ" in flags(cases)


@pytest.mark.skipif(not NODE, reason="needs node for the preview port")
def test_the_port_agrees_without_a_size_series(tmp_path):
    cases = cases_of(tmp_path, "--no-series")
    assert {c["fits"]["floorN"][1] for c in cases} == {None, "lab minimum", "configuration"}
    assert "N below the minimum peak" in flags(cases) and "C below the minimum peak" in flags(cases)
    assert not any("series floor" in f or "linearity range" in f for f in flags(
        cases, lambda c: not {"minAreaN", "minAreaC"} & set(c["case"].get("settings", {}))))


@pytest.mark.skipif(not NODE, reason="needs node for the preview port")
def test_the_port_agrees_on_a_long_size_series(tmp_path):
    cases = cases_of(tmp_path, "--long-series")
    steps = [len(c["fits"][k][5]) for c in cases for k in ("floorN", "floorC")]
    assert max(steps) >= 10 and 1 in steps
    assert "N below series floor" in flags(cases) and "C below series floor" in flags(cases)
