"""The page's own logic (web/logic.js) under node, on the service's answers for the 2026-09-15 test run.

Findings of the page that scripts/e2e_web.py cannot reach on its own: the Methods text of a run that lacks an
input (W-14), the reference-gas SD limit from the configuration (R-22), the two checks whose data the engine
sends (the elemental correction at the anchors and the anchors' material), which the page shows when the data
is there and leaves out when it is not, and the configuration each run of a batch names (W-11); with the
labels of R-29 and R-30 and the sample sheet's summary of R-34. From ea_irms 1.4.15: the blank below detection,
the floor the size series sets with the walk that set it (or the configuration's floor, or the lab's minimum peak
of a run without a series), what the 1σ holds (the corrections' terms, the drift monitors' SD in place of the
repeatability), and the two new checks, the acetanilide against the samples' 1σ and the drift monitors against an
accepted value, each with its button back to its step. Each test runs node and skips without it.
"""
import json
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from ea_irms.parse import parse_export
from ea_irms.pipeline import DEFAULT_STANDARDS, Config, process
from isoharness.app import app

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pagelogic import NODE, RUN, renamed, run_page  # noqa: E402

pytestmark = pytest.mark.skipif(not NODE, reason="needs node for the page's logic")
client = TestClient(app)
DRIFT = [10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 55]
ANCHORS, LINEARITY, BLANKS = [11, 12, 13, 15, 16, 17], [6, 7, 8, 9], [4, 5, 56]
STATE = "const d = L.state.dec, M = L.model(d);"
H = "const H = { pick() {}, toggle() {}, set() {}, filter() {}, go: () => () => null };"


def body_of(path):
    r = client.post("/api/process", files={"file": (Path(path).name, Path(path).read_bytes())})
    assert r.status_code == 200, r.text
    return r.json()


@pytest.fixture(scope="module")
def body():
    return body_of(RUN)


# ---------------------------------------------------------------- W-14: no zero from a missing input
def methods(path):
    return run_page([body_of(path)], STATE + " return L.methods(M, d);")


def test_w14_a_run_without_drift_monitors_says_so(tmp_path):
    text = methods(renamed(tmp_path, {r: f"M{r}" for r in DRIFT}, "no_monitors.csv"))
    assert "; the run had no drift monitors, so the elemental drift was not checked." in text
    assert "No drift correction was applied; the run has no drift monitors." in text
    assert "0.0%" not in text and "largest" not in text


def test_w14_a_run_without_anchors_blank_or_size_series_says_so(tmp_path):
    text = methods(renamed(tmp_path, {r: f"A{r}" for r in ANCHORS}, "no_anchors.csv"))
    assert "The run could not be put on the AIR and VPDB scales, since it has no anchors." in text
    assert "No uncertainty is reported, since the run has no anchors." in text
    assert "slope" not in text.split("No peak-size")[0]
    assert "No blank was subtracted; the run has no blank capsules." in methods(
        renamed(tmp_path, {r: f"B{r}" for r in BLANKS}, "no_blanks.csv"))
    assert "No peak-size correction was applied; the run has no size standards." in methods(
        renamed(tmp_path, {r: f"L{r}" for r in LINEARITY}, "no_series.csv"))


def test_w14_one_gas_without_a_blank_and_every_monitor_left_out(body):
    out = run_page([body], STATE + """
        const all = Object.assign({}, d, { exclE: L.state.run.port.pos.filter((q) => q.role === 'D').map((q) => q.r) });
        return [L.methods(M, d), L.methods(L.model(all), all)];""")
    # ea_irms 1.4.15: blank capsules that gave no peak put the blank below detection
    assert "The carbon blank was below detection (no blank capsule gave a CO2 peak), so none was subtracted." in out[0]
    assert "largest bias" in out[0] and "0.0%" not in out[0]
    assert "; every drift monitor was left out of the elemental correction, so the elemental drift was not checked." in out[1]


# ---------------------------------------------------------------- R-22: the reference-gas SD limit
def test_r22_the_reference_gas_limit_is_the_configurations(body):
    out = run_page([body], """
        const texts = () => { L._mcs.clear(); return L.instChecks(L.model(L.state.dec)).filter((c) => /^Reference gas/.test(c.title)).map((c) => [c.s, c.t]); };
        const out = { absent: texts() }, real = engine.runModel.bind(engine);
        for (const [key, v] of [['strict', 0.001], ['off', null]]) {
          engine.runModel = (D, e) => { const M = real(D, e); M.settings.refSdLimit = v; return M; };
          out[key] = texts();
        }
        engine.runModel = real;
        L.state.run.port.settings.refSdLimit = 0.15;      // the data block alone carries it
        out.block = texts();
        return out;""")
    firms = "the 0.1‰ limit the FIRMS guide (Dunn and Carter, 2018, section 4.3.2) sets for working-gas pulses"
    assert len(out["absent"]) == 2 and all("under " + firms in t for _, t in out["absent"])
    assert all(s == "note" and "at or above the configuration’s 0.001‰ limit (the FIRMS guide, Dunn and Carter, 2018, "
               "section 4.3.2, sets 0.1‰ for working-gas pulses)" in t for s, t in out["strict"])
    for (s0, t0), (s, t) in zip(out["absent"], out["off"]):
        assert "(the configuration sets no limit for this SD); " in t and "limit the FIRMS guide" not in t
        assert s == s0         # the trend alone sets the status, as it did with the SD under the limit
    assert all("under the configuration’s 0.15‰ limit" in t for _, t in out["block"])


# ---------------------------------------------------------------- the two checks the engine adds
ELEM = [["note", {"n": 6, "rmsMeas": 0.0115, "rmsRep": 0.0162, "added": 0.0047, "level": "Note"},
         {"n": 6, "rmsMeas": 0.0100, "rmsRep": 0.0080, "added": -0.0020, "level": "OK"}],
        ["caution", {"n": 6, "rmsMeas": 0.02, "rmsRep": 0.08, "added": 0.06, "level": "Caution"}, None],
        ["few", {"n": 1, "rmsMeas": None, "rmsRep": None, "added": None, "level": None}, None]]
MATERIAL = {"level": "Caution", "limits": {"ratio": 1.5, "spread": 5.0}, "materials": [
    {"name": "USGS 64", "si": 0, "n": 15, "ratio": [0.51, 2.07], "ratioX": 4.1, "d15n": [-0.42, 1.06],
     "d13c": [-24.05, -13.54], "positions": [11, 12, 13], "over": ["ratio"], "level": "Caution"},
    {"name": "USGS 65", "si": 1, "n": 3, "ratio": [0.95, 1.02], "ratioX": 1.07, "d15n": [20.5, 20.8], "d13c": None,
     "positions": [15, 16, 17], "over": [], "level": "OK"},
    {"name": "Ace Bypass", "si": 7, "n": 1, "ratio": None, "ratioX": None, "d15n": None, "d13c": None,
     "positions": [19], "over": [], "level": None}]}
NEW = ("Elemental correction at the anchors", "Anchors read as one material each")


def test_the_two_engine_checks_show_their_data(body):
    out = run_page([body], """
        const pick = () => { L._mcs.clear(); const d = L.state.dec; return L.qcChecks(L.model(d), d).filter((c) => input.titles.includes(c.title)).map((c) => ({ title: c.title, step: c.step, s: c.s, label: c.label, t: c.t })); };
        const real = engine.runModel.bind(engine), out = { real: pick() };
        // with the engine's data taken away, neither check shows
        engine.runModel = (D, e) => { const M = real(D, e); M.elem.n.check = null; M.elem.c.check = null; return M; };
        delete L.state.run.port.anchorMaterial;
        out.none = pick();
        for (const [key, n, c] of input.elem) {
          engine.runModel = (D, e) => { const M = real(D, e); M.elem.n.check = n; M.elem.c.check = c; return M; };
          out[key] = pick();
        }
        // the material cases below with the elemental check taken away, so each case shows one check
        engine.runModel = (D, e) => { const M = real(D, e); M.elem.n.check = null; M.elem.c.check = null; return M; };
        L.state.run.port.anchorMaterial = input.material;
        out.material = pick();
        L._mcs.clear();
        out.go = L.vQC(L.model(L.state.dec), L.state.dec, null, { go: () => null }).checks.filter((c) => input.titles.includes(c.title)).map((c) => c.goLabel);
        L.state.run.port.anchorMaterial = Object.assign({}, input.material, { level: 'OK', materials: input.material.materials.map((m) => Object.assign({}, m, { level: m.level && 'OK' })) });
        out.ok = pick();
        L.state.run.port.anchorMaterial = Object.assign({}, input.material, { level: null, materials: [input.material.materials[2]] });
        out.unjudged = pick();
        return out;""", elem=ELEM, material=MATERIAL, titles=list(NEW))
    assert out["none"] == [] and out["unjudged"] == []
    # the test run as the engine and the port give it (isoharness 0.5.0): %N corrected, both checks OK
    assert sorted((c["title"], c["step"], c["s"]) for c in out["real"]) == \
        [("Anchors read as one material each", "cal", "ok"), ("Elemental correction at the anchors", "elem", "ok")]
    rule = " A Note when the correction adds more than 2.5 points, a Caution beyond 5, the elemental gate (5%)."
    assert out["note"] == [{"title": NEW[0], "step": "elem", "s": "note", "label": "Note", "t":
        "%N: the anchors read 1.15% from their certified values as measured and 1.62% corrected (RMS of 6 analyses), "
        "0.47 points more. %C: the anchors read 1.00% from their certified values as measured and 0.80% corrected (RMS "
        "of 6 analyses), 0.20 points less." + rule}]
    assert out["caution"] == [{"title": NEW[0], "step": "elem", "s": "caution", "label": "Caution", "t":
        "%N: the anchors read 2.00% from their certified values as measured and 8.00% corrected (RMS of 6 analyses), "
        "6.00 points more." + rule + " Corrected, the anchors read further from their certified values than the gate "
        "the correction itself uses; check the drift monitors it relies on, leave out one that is off, or switch the "
        "correction Off."}]
    assert out["few"][0]["s"] == "note" and out["few"][0]["t"].startswith(
        "%N: fewer than two anchor analyses have a certified %N, so the correction is not judged at the anchors.")
    assert out["material"] == [{"title": NEW[1], "step": "cal", "s": "caution", "label": "Caution", "t":
        "USGS 64, 15 analyses: C/N peak-area ratio 0.51 to 2.07 (4.10 times); raw δ¹⁵N −0.42 to 1.06‰; raw δ¹³C −24.05 "
        "to −13.54‰. USGS 65, 3 analyses: C/N peak-area ratio 0.95 to 1.02 (1.07 times); raw δ¹⁵N 20.50 to 20.80‰. "
        "Analyses of one material agree in their C/N peak-area ratio and their raw values. A ratio more than 1.5 times "
        "another, or raw values more than 5‰ apart, means some capsules hold another material or carry the wrong name, "
        "or that the instrument did not measure the anchors consistently; check the labels, then the calibration "
        "line, before using the calibration."}]
    assert out["go"] == ["Go to Step 6, Calibration, for Anchors read as one material each"]
    assert out["ok"][0]["s"] == "ok" and out["ok"][0]["t"].endswith("raw δ¹⁵N 20.50 to 20.80‰.")


# ---------------------------------------------------------------- W-11: the configuration of each run
def test_w11_each_run_names_its_configuration(tmp_path, body):
    copy = renamed(tmp_path, {}, "20260916_copy.csv")
    out = run_page([body, body_of(copy)], """
        const out = { none: L.batchCfgName() }, view = () => L.upView(L.state, L.state.run, L.model(L.state.dec), '');
        L.setState({ cfg: { name: 'lab.xlsx', data: {}, unknown: [], notes: [] } });
        out.batch = L.batchCfgName();
        out.batchNote = view().cfgNote;
        L.state.runs[1].own = { name: 'own.json', data: {}, unknown: [], notes: [] };
        out.own = L.batchCfgName();
        const up = view();
        out.rows = up.runs.map((r) => [r.cfgText, r.hasOwn, r.cfgChooseT, r.cfgChooseLabel, r.cfgClearLabel, r.cfgId]);
        out.note = up.cfgNote;
        L.state.runs[1].own.name = 'x'.repeat(190) + '.json';
        out.long = L.batchCfgName();
        L.setState({ cfg: null });
        L.state.runs[0].own = { name: 'own.json' };
        out.all = L.batchCfgName();
        out.allNote = view().cfgNote;
        return out;""")
    assert out["none"] is None and out["batch"] == "lab.xlsx"
    assert out["batchNote"].endswith(" Choosing another processes all 2 runs again.")
    assert out["own"] == "lab.xlsx for run 1; own.json for run 2"
    assert out["rows"] == [
        ["Configuration: the batch’s, lab.xlsx", False, "Choose a configuration", "Choose a configuration for 20260915_raw.csv",
         "Use the batch’s configuration for 20260915_raw.csv", "dc-run-config-1"],
        ["Configuration: its own, own.json", True, "Choose another configuration",
         "Choose another configuration for 20260916_copy.csv", "Use the batch’s configuration for 20260916_copy.csv",
         "dc-run-config-2"]]
    assert out["note"].endswith(" Choosing another processes again the run that takes it; 20260916_copy.csv keeps its own.")
    assert out["long"] == "lab.xlsx for run 1; configurations of their own for run 2"
    assert out["all"] == "configurations of their own for runs 1 and 2"
    assert out["allNote"].endswith(" Every run has a configuration of its own, so choosing another here changes no run; it "
                                   "goes with exports added later and with a run set back to the batch’s.")


# ---------------------------------------------------------------- R-29 and R-30: names and labels
def test_r29_r30_names_and_labels(body):
    out = run_page([body], H + """
        const d = Object.assign({}, L.state.dec, { exclN: [18] }), M = L.model(d);
        const insp = L.inspect(M, 20, 'drift'), drift = L.vDrift(M, d, 18, H), figs = L.numbered(drift.charts);
        return { facts: insp.facts.map((f) => f.k), tail: insp.tail.map((f) => f.k), chain: insp.chain.map((c) => [c.stage, c.dn, c.dc]),
          left: drift.stats.rows.find((r) => r.k === 'Left out'), alt: figs[0].alt, points: figs[0].pointsLabel,
          cur: figs[0].hits.filter((h) => h.cur === 'true').length, ribbon: L.ribbon(M, 'drift', 18, H).filter((b) => b.cur === 'true').length,
          go: L.vQC(M, d, null, H).checks.map((c) => c.goLabel), cols: L.upView(L.state, L.state.run, M, '').cols.map((c) => c.field),
          filters: L.vResults(M, d, 20, H, 'all').filters.map((f) => [f.t, f.pressed]), rail: L.railItems(M, d, 'drift', H).map((s) => s.cur) };""")
    assert "Mass (mg)" in out["facts"] and "Amount" not in out["facts"] and "C:N (molar)" in out["tail"]
    assert out["chain"][-1] == ["Calibrated", "to the scale", "to the scale"]
    assert out["left"] == {"k": "Left out", "a": "position 18", "b": "none"}
    assert out["alt"].startswith("Fig. 1, δ¹⁵N of the drift monitors: δ¹⁵N after linearity (‰) against position in the "
                                 "run, 13 points; the selected position is Ace Drift3, position 18: ") and out["alt"].endswith(", left out")
    assert out["points"] == "Fig. 1, δ¹⁵N of the drift monitors, 13 points; the arrow keys move between them"
    assert out["cur"] == 1 and out["ribbon"] == 1 and out["rail"].count("step") == 1 and out["rail"][4] == "step"
    assert len(set(out["go"])) == len(out["go"]) and all(g.startswith("Go to Step ") and ", for " in g for g in out["go"])
    assert "Mass" in out["cols"] and "Weight" not in out["cols"]
    assert out["filters"][0] == ["All 26", "true"] and [p for _, p in out["filters"][1:]] == ["false"] * 3


# ---------------------------------------------------------------- R-34: replicates in the sheet's summary
def sheet_info(bodies, text):
    runs = [{"name": b["file"]["name"], "samples": [[p["row"], p["identifier"]] for p in b["positions"] if p["role"] == "Sample"]}
            for b in bodies]
    r = client.post("/api/samplesheet/read", files={"file": ("sheet.csv", text.encode())}, data={"runs": json.dumps(runs)})
    assert r.status_code == 200, r.text
    return r.json()


def summary(bodies, text):
    return run_page(bodies, """L.setState({ sheet: { file: null, name: 'sheet.csv', idColumn: input.info.sheet.id_column, info: input.info } });
        return L.sheetView(L.state, false).summary;""", info=sheet_info(bodies, text))


def test_r34_the_summary_says_when_no_group_of_replicates_formed(tmp_path, body):
    named = body_of(renamed(tmp_path, {20: "366 - H+S", 21: "366 - H+S Dup"}, "366.csv"))
    assert summary([named], "Sample ID,Replicate of\n366 - H+S,\n") == (
        "2 of 26 samples matched (1 by the exact name, 1 by an ID the name begins with); 24 not in the sheet. No group "
        "of replicates formed in this run. Names ending in Dup, such as 366 - H+S Dup, are compared only when the sheet "
        "has a row for that name whose Replicate of names the analysis it repeats.")
    sheet = "Sample ID,Replicate of\n366 - H+S,\n366 - H+S Dup,366 - H+S\n"
    assert summary([named], sheet) == ("2 of 26 samples matched (2 by the exact name); 24 not in the sheet. 1 group of "
                                       "replicates in this run will be compared.")
    assert summary([named, body], sheet).endswith(" 1 group of replicates within a run will be compared; none formed in "
                                                  "20260915_raw.csv.")
    assert summary([named, body], "Sample ID\nnothing\n").endswith(" No group of replicates formed in any run. Names "
                                                                   "ending in Dup, such as 366 - H+S Dup, are compared only "
                                                                   "when the sheet has a row for that name whose Replicate "
                                                                   "of names the analysis it repeats.")


# ---------------------------------------------------------------- ea_irms 1.4.15
def by_row(positions):
    return {p.row: p for p in positions}


def test_1415_a_blank_without_a_peak_is_below_detection(body):
    out = run_page([body], H + STATE + """
        const b = L.vBlank(M, d, null, H);
        return { lines: b.ctrl[0].lines.map((l) => l.t), empty: b.charts[1].emptyText, rows: b.stats.rows,
          after: L.explainView(M, d, 'blank', null).after.map((l) => l.t), port: [M.blank.c.belowDetection, M.blank.n.belowDetection] };""")
    assert out["port"] == [True, False]
    assert out["lines"][1] == "δ¹³C: below detection, since none of the 3 blank capsules gave a CO2 peak"
    assert out["empty"] == ("None of the 3 blank capsules gave a CO2 peak, so the carbon blank is below detection. It was "
                            "measured and is too small to see, so there is nothing to subtract.")
    rows = {r["k"]: (r["a"], r["b"]) for r in out["rows"]}
    assert rows["Capsules with a peak"] == ("2 of 3", "0 of 3") and rows["Status"] == ("Not subtracted", "Below detection")
    assert any(t.startswith("No blank capsule gave a CO2 peak, so the carbon blank is below detection.") for t in out["after"])


def test_1415_the_floor_the_size_series_sets(body):
    out = run_page([body], H + STATE + """
        const l = L.vLin(M, d, null, H), fig = l.charts[0];
        const on = Object.assign({}, d, { blank: 'On' }), Mo = L.model(on);
        const flagged = (m) => m.P.filter((q) => q.role === 'S' && q.flags.some((f) => /^N below/.test(f))).map((q) => [q.r, q.flags.filter((f) => /^N below/.test(f))]);
        return { floor: l.floor, limits: M.limits, sub: fig.sub, legend: fig.legend.map((x) => x.label), band: fig.s1.d !== '', hollow: fig.s9.d !== '',
          stats: l.stats.rows.filter((r) => /Floor|below the floor/.test(r.k)),
          size: L.qcChecks(M, d).find((c) => c.title === 'Sample size against the size series').t,
          flagged: flagged(M), blankOn: [Mo.limits.n.min, Mo.limits.n.src, flagged(Mo)],
          results: L.vResults(M, d, null, H, 'all').filters.map((f) => f.t),
          chips: input.flags.map((f) => [L.chip(f), L.ftext(f)]) };""",
                   flags=["N below series floor", "C below the minimum peak", "acetanilide d15N scatter above 1σ",
                          "monitor d15N off by -0.73"])
    f = out["floor"]
    assert f["title"] == "Floor of the size series" and f["hasSteps"]
    assert f["lines"][2]["t"].startswith("δ¹⁵N: the floor is 12.06 Vs, the peak of Ace Lin2 (7), since the smallest standard "
                                         "lies off the line through the larger ones. 8 samples have a smaller N2 peak: Random (21), ")
    assert f["lines"][3]["t"].startswith("δ¹³C: the floor is the smallest standard, Ace Lin1 (6), at 37.90 Vs, which lies within "
                                         "the limit of the line through the larger ones. 7 samples have a smaller CO2 peak: ")
    assert [(s["iso"], s["std"], s["area"], s["result"]) for s in f["steps"]] == [
        ("δ¹⁵N", "Ace Lin1 (6)", "8.89", "Beyond the limit, so the floor moves up to Ace Lin2 (7)"),
        ("δ¹³C", "Ace Lin1 (6)", "37.90", "Within the limit, so this standard is the floor")]
    assert (out["limits"]["n"]["src"], out["limits"]["c"]["src"]) == ("series floor", "smallest standard")
    assert "; floor 12.1 Vs; " in out["sub"] and out["band"] and out["hollow"]
    assert {"Off the line, under the floor", "Below the floor"} <= set(out["legend"])
    assert out["stats"] == [{"k": "Floor (Vs)", "a": "12.06, set by the series", "b": "37.90, the smallest standard"},
                            {"k": "Samples above the series, below the floor", "a": "2, 8", "b": "2, 7"}]
    assert out["size"].startswith("N2: 8 samples below the floor (12.06 Vs) and 2 above the series. CO2: 7 samples below the "
                                  "floor (37.90 Vs) and 2 above the series.")
    # the floor and its flags as the engine gives them, for the decisions shown and with the blank subtracted
    for decisions, (lim, src, flagged) in ((Config(), (out["limits"]["n"]["min"], out["limits"]["n"]["src"], out["flagged"])),
                                           (Config(blank_correction="On"), out["blankOn"])):
        ps, F = process(parse_export(RUN), decisions)
        assert (F["size_limits"]["min_n"], F["size_limits"]["min_n_source"]) == (lim, src)
        assert [[p.row, [f for f in p.flags if f.startswith("N below")]] for p in ps if p.role == "Sample"
                and any(f.startswith("N below") for f in p.flags)] == flagged
    assert any(fl == ["N below series floor"] for _, fl in out["flagged"])
    assert out["results"] == ["All 26", "Peak size accepted 15", "Peak size flagged 11", "Carryover 0"]
    assert out["chips"] == [["N2 under the floor", "N below series floor"], ["CO2 under the minimum", "C below the minimum peak"],
                            ["δ¹⁵N scatter high", "acetanilide δ¹⁵N scatter above 1σ"], ["monitor δ¹⁵N off by −0.73", "monitor δ¹⁵N off by −0.73"]]


def test_1415_a_floor_from_the_configuration_or_the_lab_minimum(tmp_path, body):
    code = H + """
        const s = L.state.run.port.settings;
        Object.assign(s, input.settings);
        L._mcs.clear();
        const d = L.state.dec, M = L.model(d), f = L.floorView(M);
        return { lines: f.lines.map((l) => l.t), steps: f.steps.length, src: [M.limits.n.src, M.limits.c.src],
          size: L.qcChecks(M, d).find((c) => c.title === 'Sample size against the size series').t,
          flags: M.P.filter((q) => q.role === 'S').map((q) => [q.r, q.flags]) };"""
    own = run_page([body], code, settings={"minAreaN": 15.0})
    assert own["src"] == ["configuration", "smallest standard"] and own["steps"] == 1
    assert own["lines"][2].startswith("δ¹⁵N: the configuration sets the floor at 15.00 Vs, so the size series does not set it. ")
    ps, _ = process(parse_export(RUN), Config(min_area_n=15.0))
    assert own["flags"] == [[p.row, p.flags] for p in ps if p.role == "Sample"]
    # a run without a size series: the four Ace Lin capsules read as Ace Bypass, an Exclude row
    series = body_of(renamed(tmp_path, {r: "Ace Bypass" for r in LINEARITY}, "no_series.csv"))
    none = run_page([series], code, settings={})
    assert none["src"] == [None, None] and none["steps"] == 0 and len(none["lines"]) == 3
    assert none["lines"][0] == "A sample whose peak is smaller than the floor is flagged and marked for a rerun."
    assert none["lines"][1] == ("δ¹⁵N: this run has no size series, and the configuration sets no minimum peak, so no N2 peak "
                                "is too small.")
    assert none["size"].startswith("N2: no size series and no minimum peak, so no peak is too small or too large. ")
    lab = run_page([series], code, settings={"noSeriesMinN": 20.0, "noSeriesMinC": 60.0})
    assert lab["src"] == ["lab minimum", "lab minimum"]
    assert lab["lines"][1].startswith("δ¹⁵N: this run has no size series, so the floor is the lab’s minimum peak, 20.00 Vs. ")
    assert lab["size"].startswith("N2: ") and "below the lab’s minimum peak (20.00 Vs)" in lab["size"]
    ps, _ = process(parse_export(tmp_path / "no_series.csv"), Config(no_series_min_area_n=20.0, no_series_min_area_c=60.0))
    assert lab["flags"] == [[p.row, p.flags] for p in ps if p.role == "Sample"]
    assert any("N below the minimum peak" in fl for _, fl in lab["flags"])


def test_1415_what_the_1sigma_holds(body):
    code = H + """
        Object.assign(L.state.run.port.settings, input.settings);
        L._mcs.clear();
        const d = L.state.dec, M = L.model(d), u = L.vUnc(M, d, 20, H), ex = L.explainView(M, d, 'unc', 20), q = M.P.find((p) => p.r === 20);
        return { u: [q.n.u, q.c.u], terms: [q.n.terms, q.c.terms], rows: ex.rows, lines: ex.lines.map((l) => l.t), formulas: ex.formulas.map((f) => f.t),
          budget: u.budget.rows.map((r) => [r.k, r.n, r.c]), stats: u.stats.rows.map((r) => [r.k, r.a, r.b]), ctrl: u.ctrl.map((c) => c.lines.map((l) => l.t)),
          effect: u.effect.lines.map((l) => l.t), checks: u.checks.map((c) => c.title), header: L.renderVals().text.uncTerms,
          figure: u.charts.map((c) => [c.sub, c.legend.map((x) => x.label)]), methods: L.methods(M, d),
          mon: [M.unc.n.monitors.used, M.unc.c.monitors.used] };"""
    out = run_page([body], code, settings={})
    ps, F = process(parse_export(RUN), Config())
    p20 = by_row(ps)[20]
    assert out["u"] == pytest.approx([p20.u_d15n, p20.u_d13c], rel=1e-12)
    assert set(out["terms"][0]) == {"drift"} and set(out["terms"][1]) == {"linearity", "drift"}
    rows = {r["k"]: (r["a"], r["b"]) for r in out["rows"]}
    # the worked example adds up: the sum of its variances is the square of the 1σ shown
    for i, u in enumerate(out["u"]):
        assert abs(float(rows["Sum (‰²)"][i]) - u * u) < 2e-6
    assert rows["Linearity term² (‰²)"][0] == "not in force" and rows["Blank term² (‰²)"] == ("not in force", "not in force")
    assert float(rows["Drift term² (‰²)"][0]) == pytest.approx(p20.u_terms_n["drift"] ** 2, abs=1e-6)
    assert out["formulas"][0] == "1σ = √[(s × m)² + certified term + measured-means term + blank term² + linearity term² + drift term²]"
    assert out["lines"][1].startswith("Each correction in force adds a term of its own, as a Kragten spreadsheet does.")
    assert [k for k, _, _ in out["budget"]] == ["Single-analysis repeatability", "Anchors’ certified values", "Anchors’ measured means",
                                                "Linearity correction", "Drift correction"]
    assert out["budget"][3][1] == "" and out["budget"][3][2] != ""          # no linearity term for δ¹⁵N, one for δ¹³C
    stats = {k: (a, b) for k, a, b in out["stats"]}
    assert stats["Largest drift term (‰)"][0] == f"{F['unc_n']['corrections']['drift']:.3f}"
    assert stats["Monitors’ SD in place of the repeatability"] == ("no, p = 0.744", "no, p = 0.083")
    assert out["ctrl"][1][1].endswith("(F test, p = 0.083); larger, but not at p < 0.05, so the repeatability stays")
    assert any(t.startswith("In δ¹⁵N, the drift correction in force adds a term of up to ") for t in out["effect"])
    assert out["checks"] == ["Acetanilide scatter against 1σ"]
    assert out["header"].startswith("It also holds a term for each correction in force")
    assert "(a one-sided F test, which compares the two variances, at p < 0.05)" in out["header"]
    assert "the corrections’ terms lift a sample above the line" in out["figure"][0][0]
    assert out["figure"][0][1][:2] == ["Sample, with the corrections’ terms", "1σ along the line, without them"]
    assert out["methods"].endswith("propagated through the calibration, with a Kragten term for the peak-size and drift corrections in force.")
    # the configuration's switches the other way: no terms, the monitors' SD wherever larger
    off = run_page([body], code, settings={"kragten": "Off", "monitorRep": "On"})
    ps, F = process(parse_export(RUN), Config(kragten_corrections="Off", monitor_repeatability="On"))
    p20 = by_row(ps)[20]
    assert off["u"] == pytest.approx([p20.u_d15n, p20.u_d13c], rel=1e-12) and off["terms"] == [{}, {}]
    assert off["mon"] == [F["unc_n"]["monitors"]["used"], F["unc_c"]["monitors"]["used"]] == [False, True]
    assert off["header"] == ("The configuration leaves out the terms of the blank, linearity, and drift corrections. Where the drift "
                             "monitors scatter more than the repeatability allows, their SD takes its place.")
    assert "Drift monitors’ SD² in its place (‰²)" in [r["k"] for r in off["rows"]]
    assert [k for k, _, _ in off["budget"]][0] == "Repeatability (the drift monitors’ SD for δ¹³C)"
    assert off["ctrl"][1][1].endswith("; their SD is used in its place")
    assert "the drift monitors’ SD in place of s × m" in off["figure"][1][0] and "s × m" not in off["figure"][0][0]
    assert off["figure"][0][1][0] == "Sample"            # without the corrections' terms nothing lifts a sample
    assert off["methods"].endswith("they do not include the blank, peak-size, or drift corrections. For δ¹³C, the drift monitors’ SD "
                                   f"({F['unc_c']['monitors']['sd']:.3f}‰) took the place of the repeatability, being larger "
                                   f"(F test, p = {F['unc_c']['monitors']['p']:.3f}).")


ACCEPTED = [tuple(s[:2]) + (-26.6, -5.2) + tuple(s[4:]) if s[1] == "Drift" else s for s in DEFAULT_STANDARDS]


def test_1415_the_acetanilide_and_the_monitors_against_an_accepted_value(body):
    out = run_page([body], H + """
        const pick = () => { L._mcs.clear(); const d = L.state.dec, M = L.model(d);
          return L.vQC(M, d, null, H).checks.filter((c) => /^Acetanilide|accepted value$/.test(c.title)).map((c) => [c.step, c.title, c.s, c.goLabel, c.t]); };
        const out = { plain: pick() };
        L.state.run.port.std = input.std;
        out.accepted = pick();
        L.state.run.port.settings.checkOffset = 0.15;
        out.tight = pick();
        const d = L.state.dec, M = L.model(d);
        out.steps = [L.vDrift(M, d, null, H).checks.map((c) => c.title), L.vUnc(M, d, null, H).checks.map((c) => c.title)];
        out.flags = M.P.filter((q) => q.role === 'D').map((q) => [q.r, q.flags]);
        L.state.run.port.settings.aceCheck = 'Off';
        out.off = pick();
        return out;""", std=[list(s) for s in ACCEPTED])
    ace = ["unc", "Acetanilide scatter against 1σ", "caution", "Go to Step 8, Uncertainty, for Acetanilide scatter against 1σ"]
    assert len(out["plain"]) == 1 and out["plain"][0][:4] == ace
    assert out["plain"][0][4].startswith("δ¹⁵N: the 4 standards of the size series and 13 drift monitors have an SD of 0.759‰ against "
                                         "the samples’ repeatability of 0.173‰ (F test, p = 0.006), more than the repeatability "
                                         "allows, so every sample’s δ¹⁵N is marked for a check. δ¹³C: the 4 standards of the size "
                                         "series and 13 drift monitors have an SD of 0.043‰ against the samples’ repeatability of "
                                         "0.021‰ (F test, p = 0.086), larger, but not at p < 0.05. ")
    ps, F = process(parse_export(RUN), Config())
    assert (F["ace_n"]["over"], F["ace_c"]["over"]) == (True, False) and round(F["ace_n"]["sd"], 3) == 0.759
    mon = ["drift", "Drift monitors against their accepted value"]
    assert [c[:4] for c in out["accepted"]] == [mon + ["ok", "Go to Step 5, Drift, for Drift monitors against their accepted value"], ace]
    assert out["accepted"][0][4] == ("13 drift monitors in this run against the accepted value the standards table gives them "
                                     "(δ¹⁵N −5.20‰ and δ¹³C −26.60‰); each is within 0.5‰ of it.")
    # with a check offset of 0.15 permil, three monitors lie off it in d15N, as the engine and its workbook find
    ps, _ = process(parse_export(RUN), Config(standards=ACCEPTED, check_offset_permil=0.15))
    assert out["flags"] == [[p.row, p.flags] for p in ps if p.role == "Drift"]
    off = [p for p in ps if p.role == "Drift" and any(f.startswith("monitor d15N off by ") for f in p.flags)]
    assert [p.row for p in off] == [10, 30, 38]
    assert out["tight"][0][2] == "note" and out["tight"][0][4].startswith(
        "13 drift monitors in this run against the accepted value the standards table gives them (δ¹⁵N −5.20‰ and δ¹³C −26.60‰); 3 "
        "differ from it by more than 0.15‰: Ace Drift1 (10), δ¹⁵N off by ")
    assert out["steps"] == [["Effect of the drift choice", "Drift monitors agree", "Drift monitors against their accepted value"],
                            ["Acetanilide scatter against 1σ"]]
    assert out["off"][-1][:3] == ["unc", "Acetanilide scatter against 1σ", "note"] and out["off"][-1][4] == "The configuration turns this check off."
