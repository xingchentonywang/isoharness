"""Engine outputs for the preview port to be checked against.

Runs ea_irms on one export under many decision sets (all 243 combinations of the
blank, linearity and drift switches and the drift model, then random exclusions,
elemental settings and sample repeatabilities), then under random decision sets
with non-default settings (the significance level, the monitor outlier limit,
the check-standard offset, the blank multiple, k, the elemental gate and the
first drift position, and from ea_irms 1.4.15 the floor's offset and SD multiple,
a floor set outright, the lab minimum of a run without a size series, and the
switches kragten_corrections, monitor_repeatability and ace_scatter_check) and
with accepted values on the standards table's Drift rows (a case with "std"), and
writes two files:

    port_input.json   the "port" block of POST /api/process for this export
    cases.json        each decision set with the engine's values, flags and fits

    python3 scripts/port_cases.py tests/data/20260915_raw.csv --out build/port
    node port/test_port.js build/port/port_input.json build/port/cases.json

With --as-filtered the export is first rewritten as a filtered one (written to
the --out folder): no mass-28 reference peaks, and where a position made no
sample CO2, an undesignated CO2 pulse ahead of the designated one, as in the
2025 exports. The reference-pulse check then sets those pulses aside, which the
port must flag as the engine does.

With --no-ref-background the background of the N2 reference peak is first
blanked at one sample position and one drift monitor (no_background_variant),
so that under the Background drift model those two are left without a drift
correction and flagged, which the port must do as the engine does (finding R-23
of the 2026-10-14 review).

With --no-series the Linearity rows of the standards table take the role Exclude
first, so the run has no size series and its peaks are held to the lab minimum
(no_series_min_area_n and _c), and with --long-series ten samples are renamed to
linearity standards first (long_series_variant), so the walk up the size series
that sets the floor takes many steps (ea_irms 1.4.15).

Decisions start from --config (or the package defaults). The port reads those
settings from the block (port/engine.js, settings()); a case with "settings"
overrides them on both sides, and a case with "std" replaces the standards
table, which it does only by giving the Drift rows accepted values, so the roles
the server assigned stand. The other columns of the standards, the reference-pulse
windows, the largest peak of the size range, settling and carryover act before
the block is built, on the server, so this check does not vary them.
"""
from __future__ import annotations

import argparse
import csv
import json
import random
import sys
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ea_irms.parse import parse_export                                # noqa: E402
from ea_irms.pipeline import Config, config_from_dict, process        # noqa: E402
from isoharness.serialize import port_input                           # noqa: E402

SWITCH = ["Auto", "On", "Off"]

# The port's names for the settings it reads, and the Config fields behind them.
SETTING_FIELDS = {"alpha": "significance_level", "outlierT": "monitor_outlier_t",
                  "checkOffset": "check_offset_permil", "blankMultiple": "min_blank_multiple",
                  "k": "coverage_k", "elemGate": "elemental_threshold_pct",
                  "firstDrift": "first_drift_position",
                  # ea_irms 1.4.15
                  "floorOffset": "floor_offset_permil", "floorSdMultiple": "floor_sd_multiple",
                  "minAreaN": "min_area_n", "minAreaC": "min_area_c",
                  "noSeriesMinN": "no_series_min_area_n", "noSeriesMinC": "no_series_min_area_c",
                  "kragten": "kragten_corrections", "monitorRep": "monitor_repeatability",
                  "aceCheck": "ace_scatter_check"}
# Each variant is run under --per-variant random decision sets.
VARIANTS = [{"alpha": 0.10}, {"alpha": 0.01}, {"k": 1.0}, {"k": 3.0},
            {"elemGate": 2.0}, {"elemGate": 10.0}, {"outlierT": 2.5}, {"checkOffset": 0.2},
            {"blankMultiple": None}, {"blankMultiple": 10.0}, {"firstDrift": 20},
            {"alpha": 0.2, "k": 2.5, "elemGate": 1.0, "outlierT": 3.0, "checkOffset": 0.1,
             "blankMultiple": 20.0},
            # ea_irms 1.4.15: each switch of the 1 sigma the other way (On, Auto and the floor's 0.3 and 3 are the
            # defaults), the floor off and at other values, and the switches together
            {"kragten": "Off"}, {"monitorRep": "On"}, {"monitorRep": "Off"}, {"aceCheck": "Off"},
            {"floorOffset": None}, {"floorOffset": 0.05}, {"floorOffset": 1.0}, {"floorSdMultiple": None},
            {"floorSdMultiple": 1.0}, {"floorOffset": 0.0, "floorSdMultiple": None},
            {"kragten": "Off", "monitorRep": "On", "aceCheck": "Off", "floorOffset": 0.1, "floorSdMultiple": None},
            {"monitorRep": "On", "alpha": 0.2, "checkOffset": 0.1, "floorSdMultiple": 0.5}]
# Accepted values for the Drift rows of the standards table (ea_irms 1.4.15), as offsets in d15N and d13C from the
# mean the drift monitors read (finals under the configuration's own decisions), so that some monitors lie further
# than the check-standard offset from them and some do not; with the checkOffset given, too.
DRIFT_ACCEPTED = [((0.4, -0.45), {}), ((-0.25, 0.6), {}), ((0.1, 0.1), {"checkOffset": 0.15})]


def filtered_variant(src: Path, dst: Path) -> Path:
    """Write ``src`` as a filtered export: each position keeps its last mass-28 peak
    unless it is a flagged reference (so no N2 reference peak is left) and either
    its sample CO2 peak or, where the
    first CO2 peak is the designated pulse, the next pulse ahead of it. Peaks are
    renumbered from 3, as the filtered exports number them; the parser never reads
    Peak Nr for roles."""
    with open(src, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    head, body = rows[0], rows[1:]
    i = {k: n for n, k in enumerate(head)}
    out, by_pos = [head], {}
    for r in body:
        by_pos.setdefault(r[i["Row"]], []).append(r)
    for pos_rows in by_pos.values():
        plain = [r for r in pos_rows if not r[i["Peak Nr"]].strip()]
        m28 = [r for r in pos_rows if r[i["Peak Nr"]].strip() and r[i["Ampl  28"]].strip()]
        m44 = [r for r in pos_rows if r[i["Peak Nr"]].strip() and r[i["Ampl  44"]].strip()]
        keep = [r for r in m28[-1:] if r[i["Is Ref _"]].strip() != "1"]
        if m44 and m44[0][i["Is Ref _"]].strip() == "1" and len(m44) > 1:
            keep += [m44[1], m44[0]]
        else:
            keep += m44[:1]
        for n, r in enumerate(keep, start=3):
            r = list(r)
            r[i["Peak Nr"]] = str(n)
            plain.append(r)
        out.extend(plain)
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(out)
    return dst


# A constructed CO2 peak for each blank capsule (ea_irms 1.4.10): illustrative values, not measured
# ones, chosen so that the blank correction for carbon is in force on Auto (three or more blanks whose
# d13C agree to within 2 permil); the k-th blank capsule gets these plus k times the step.
CARBON_BLANK = {"Ampl  44": (140.0, 10.0), "Area 44": (3.0, 0.2), "d 13C/12C": (-27.0, 0.3)}


def carbon_blank_variant(src: Path, dst: Path, standards=None) -> Path:
    """Write ``src`` with a sample CO2 peak in each blank capsule, which no run in the repository has, so
    that the carbon blank correction is exercised (finding R-20 of the 2026-10-14 review). The peak goes
    before the capsule's first designated CO2 peak, as a sample's does, and the peaks after it are
    renumbered one up. Blank capsules are the positions the standards table gives the Blank role."""
    from ea_irms.pipeline import DEFAULT_STANDARDS, Config, assign_roles
    with open(src, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    head, body = rows[0], rows[1:]
    i = {k: n for n, k in enumerate(head)}
    positions = {p.row: p for p in assign_roles(parse_export(src), Config(standards=standards or list(DEFAULT_STANDARDS)))}
    by_pos = {}
    for r in body:
        by_pos.setdefault(r[i["Row"]], []).append(r)
    out, k = [head], 0
    for pos, pos_rows in by_pos.items():
        p = positions.get(int(float(pos)))
        m44 = [n for n, r in enumerate(pos_rows) if r[i["Peak Nr"]].strip() and r[i["Ampl  44"]].strip()]
        refs = [n for n in m44 if pos_rows[n][i["Is Ref _"]].strip() == "1"]
        if p is None or p.role != "Blank" or not refs or (m44 and m44[0] != refs[0]):
            out.extend(pos_rows)            # not a blank capsule, or it already has a sample CO2 peak
            continue
        first = refs[0]
        nr = int(float(pos_rows[first][i["Peak Nr"]]))
        peak = list(pos_rows[first])
        for key, (base, step) in CARBON_BLANK.items():
            peak[i[key]] = f"{base + step * k:.3f}"
        peak[i["Is Ref _"]] = "0"
        k += 1
        for n, r in enumerate(pos_rows):
            r = list(r)
            if n == first:
                out.append(peak)
            if r[i["Peak Nr"]].strip() and int(float(r[i["Peak Nr"]])) >= nr:
                r[i["Peak Nr"]] = str(int(float(r[i["Peak Nr"]])) + 1)
            out.append(r)
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(out)
    return dst


def no_background_variant(src: Path, dst: Path, standards=None) -> Path:
    """Write ``src`` with the background of the N2 reference peak left empty at one sample position and one
    drift monitor, the first of each in the second half of the run (finding R-23 of the 2026-10-14 review).
    The cell emptied is BGD 28 of the position's last designated mass-28 peak in Peak Nr order, the one the
    parser reads as ref_bgd_n; no_background_rows names the two positions in the file written."""
    from ea_irms.pipeline import DEFAULT_STANDARDS, Config, assign_roles
    with open(src, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    head, body = rows[0], rows[1:]
    i = {k: n for n, k in enumerate(head)}
    positions = sorted(assign_roles(parse_export(src), Config(standards=standards or list(DEFAULT_STANDARDS))),
                       key=lambda p: p.row)
    later = positions[len(positions) // 2:]
    picked = [next((p.row for p in later if p.role == role and p.ref_bgd_n is not None), None)
              for role in ("Sample", "Drift")]
    if None in picked:
        raise ValueError(f"{Path(src).name} has no sample or no drift monitor with a reference background in the "
                         "second half of the run")
    last = {}                                       # position -> (Peak Nr, record) of its last designated N2 peak
    for n, r in enumerate(body):
        try:
            pos, nr = int(float(r[i["Row"]])), float(r[i["Peak Nr"]])
        except ValueError:
            continue                                # a TCD record, without a Peak Nr
        if pos in picked and r[i["Ampl  28"]].strip() and r[i["Is Ref _"]].strip() == "1" \
                and (pos not in last or nr >= last[pos][0]):
            last[pos] = (nr, n)
    for _, n in last.values():
        body[n] = list(body[n])
        body[n][i["BGD 28"]] = ""
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows([head] + body)
    return dst


def no_background_rows(export: Path) -> list:
    """The positions of ``export`` whose designated N2 peak carries no background (no_background_variant)."""
    return sorted(p.row for p in parse_export(export) if p.ref_area_n is not None and p.ref_bgd_n is None)


def long_series_variant(src: Path, dst: Path, standards=None, count: int = 10) -> Path:
    """Write ``src`` with ``count`` of its samples renamed to linearity standards, every other sample with an N2
    peak in position order, each to the text of the standards table's first Linearity row and a number, so that the
    size series is long and the walk up it that sets the floor takes many steps (ea_irms 1.4.15). The samples left
    keep their names."""
    from ea_irms.pipeline import DEFAULT_STANDARDS, Config, assign_roles
    standards = standards or list(DEFAULT_STANDARDS)
    text = next((s[0] for s in standards if s[1] == "Linearity" and s[0]), None)
    if text is None:
        raise ValueError("the standards table has no Linearity row to name the samples after")
    with open(src, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    head, body = rows[0], rows[1:]
    i = {k: n for n, k in enumerate(head)}
    positions = sorted(assign_roles(parse_export(src), Config(standards=standards)), key=lambda p: p.row)
    samples = [p.row for p in positions if p.role == "Sample" and p.area_n is not None]
    names = {r: f"{text} x{k}" for k, r in enumerate(samples[::2][:count])}
    out = [head]
    for r in body:
        try:
            pos = int(float(r[i["Row"]]))
        except ValueError:
            pos = None
        if pos in names:
            r = list(r)
            r[i["Identifier 1"]] = names[pos]
        out.append(r)
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(out)
    return dst


def without_series(standards) -> list:
    """The standards table with its Linearity rows given the role Exclude, so that a run has no size series
    (--no-series, ea_irms 1.4.15)."""
    return [(s[0], "Exclude") + tuple(s[2:]) if s[1] == "Linearity" else tuple(s) for s in standards]


def run_variants(positions, standards) -> list:
    """The variants whose values come from the run (ea_irms 1.4.15), as (settings, standards or None): a floor the
    configuration sets outright at the peak a third of the way up the samples', and one of 0; the lab minimum of a
    run without a size series at that peak (it acts only in such a run, --no-series), alone and under a floor set
    outright; and the standards table with accepted values on its Drift rows (DRIFT_ACCEPTED), the offsets taken from
    the mean of the drift monitors' final values under the configuration's own decisions."""
    def third(attr):
        vals = sorted(getattr(p, attr) for p in positions if p.role == "Sample" and getattr(p, attr) is not None)
        return round(vals[len(vals) // 3], 3) if vals else None
    out = []
    n, c = third("area_n"), third("area_c")
    if n is not None or c is not None:
        out += [({"minAreaN": n, "minAreaC": c}, None), ({"minAreaN": 0.0, "minAreaC": 0.0}, None),
                ({"noSeriesMinN": n, "noSeriesMinC": c}, None),
                ({"noSeriesMinN": n, "noSeriesMinC": c, "minAreaC": 0.0, "floorOffset": None}, None)]

    def mean(attr):
        vals = [getattr(p, attr) for p in positions if p.role == "Drift" and getattr(p, attr) is not None]
        return sum(vals) / len(vals) if vals else None
    m15, m13 = mean("d15n_final"), mean("d13c_final")
    if m15 is not None or m13 is not None:
        for (d15, d13), settings in DRIFT_ACCEPTED:
            std = []
            for s in standards:
                s = list(s) + [None] * (4 - len(s))
                if s[1] == "Drift":
                    s[2] = None if m13 is None else round(m13 + d13, 2)
                    s[3] = None if m15 is None else round(m15 + d15, 2)
                std.append(s)
            out.append((dict(settings), std))
    return out


def random_decisions(rng, monitors):
    return {"blank": rng.choice(SWITCH), "linN": rng.choice(SWITCH), "linC": rng.choice(SWITCH),
            "drift": rng.choice(SWITCH), "model": rng.choice(["Row", "Background", "Auto"]),
            "exclN": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 1, 2, 3])))),
            "exclC": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 2])))),
            "elem": rng.choice(SWITCH),
            "exclE": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 1, 2, 2, 3, 5])))),
            "uN": rng.choice([None, None, 0.181, 0.1]), "uC": rng.choice([None, None, 0.12])}


def decision_sets(monitors, seed, n_random, per_variant=0, extra=()):
    cases = []
    for b in SWITCH:
        for ln in SWITCH:
            for lc in SWITCH:
                for d in SWITCH:
                    for m in ("Row", "Background", "Auto"):
                        cases.append({"blank": b, "linN": ln, "linC": lc, "drift": d, "model": m,
                                      "exclN": [], "exclC": [], "elem": "Auto", "exclE": [],
                                      "uN": None, "uC": None})
    rng = random.Random(seed)
    for _ in range(n_random):
        cases.append(random_decisions(rng, monitors))
    # A separate generator, so the default-settings cases stay the same as before.
    rng2 = random.Random(seed + 1)
    for variant in VARIANTS:
        for _ in range(per_variant):
            c = random_decisions(rng2, monitors)
            c["settings"] = dict(variant)
            cases.append(c)
    # ea_irms 1.4.15: the variants whose values come from the run (run_variants), from a generator of their own
    rng3 = random.Random(seed + 2)
    for settings, std in extra:
        for _ in range(per_variant):
            c = random_decisions(rng3, monitors)
            if settings:
                c["settings"] = dict(settings)
            if std is not None:
                c["std"] = std
            cases.append(c)
    return cases


# ea_irms 1.4.15: each position's Kragten terms of the corrections in force, N then C, before its flags (which stay
# last in a case's row)
TERMS = ("blank", "linearity", "drift")


def engine_case(export, base: Config, c: dict) -> dict:
    cfg = replace(base, blank_correction=c["blank"], linearity_correction=c["linN"],
                  linearity_correction_c=c["linC"], drift_correction=c["drift"], drift_model=c["model"],
                  exclude_drift_n=c["exclN"], exclude_drift_c=c["exclC"], elemental_correction=c["elem"],
                  exclude_elemental=c["exclE"], u_sample_n=c["uN"], u_sample_c=c["uC"],
                  **{SETTING_FIELDS[k]: v for k, v in c.get("settings", {}).items()})
    if "std" in c:
        cfg = replace(cfg, standards=[tuple(s) for s in c["std"]])
    ps, f = process(parse_export(export), cfg)
    rows = [[p.row, p.d15n_blank, p.d15n_lin, p.d15n_drift, p.d15n_final, p.u_d15n, p.U_d15n,
             p.d13c_blank, p.d13c_lin, p.d13c_drift, p.d13c_final, p.u_d13c, p.U_d13c,
             p.pct_n_reported, p.pct_c_reported, p.cn_molar_reported, p.elem_factor_n,
             p.elem_factor_c, p.t_drift_n, p.t_drift_c]
            + [getattr(p, f"u_terms_{el}").get(k) for el in ("n", "c") for k in TERMS] + [p.flags] for p in ps]
    fit = lambda key, k: f[key]["fit"][k] if f[key]["fit"] else None  # noqa: E731
    fits = {"blankN": f["blank_n"]["in_force"], "blankC": f["blank_c"]["in_force"],
            "linN": [f["lin_n"]["in_force"], fit("lin_n", "p")],
            "linC": [f["lin_c"]["in_force"], fit("lin_c", "p")],
            "driftN": [f["drift_n"]["in_force"], f["drift_n"]["model"], fit("drift_n", "p"),
                       fit("drift_n", "slope"), f["drift_n"]["outliers"]],
            "driftC": [f["drift_c"]["in_force"], f["drift_c"]["model"], fit("drift_c", "p"),
                       fit("drift_c", "slope"), f["drift_c"]["outliers"]],
            "calN": [f["cal_n"]["slope"], f["cal_n"]["intercept"], f["cal_n"].get("sd")],
            "calC": [f["cal_c"]["slope"], f["cal_c"]["intercept"], f["cal_c"].get("sd")],
            "uncN": [f["unc_n"]["pooled_sd"], f["unc_n"]["u_sample"]],
            "uncC": [f["unc_c"]["pooled_sd"], f["unc_c"]["u_sample"]],
            "elemN": [f["elem_n"]["in_force"], f["elem_n"]["max_bias"], f["elem_n"]["n"]],
            "elemC": [f["elem_c"]["in_force"], f["elem_c"]["max_bias"], f["elem_c"]["n"]]}
    # ea_irms 1.4.13: the elemental drift correction against the anchors, None where it is not in force
    for el, key in (("n", "elemCheckN"), ("c", "elemCheckC")):
        chk = f["elem_check"][el]
        fits[key] = None if chk is None else [chk["level"], chk["rms_measured"], chk["rms_reported"],
                                              chk["added"], chk["n"]]
    # ea_irms 1.4.15: the floor and where it comes from, with each step of the walk up the size series; the drift
    # monitors' SD against the repeatability in use; the largest Kragten term of each correction in force; the
    # acetanilide's whole-run check; and blank capsules that gave no peak
    for el, key in (("n", "N"), ("c", "C")):
        fl, sl = f[f"floor_{el}"], f["size_limits"]
        fits["floor" + key] = [sl[f"min_{el}"], sl[f"min_{el}_source"], fl["floor"], fl["moved"], fl["smallest"],
                               [[s["row"], s["resid"], s["limit"], s["moved"]] for s in fl["steps"]]]
        mon = f[f"unc_{el}"]["monitors"]
        fits["mon" + key] = [mon["n"], mon["sd"], mon["repeatability"], mon["df"], mon["p"], mon["larger"],
                             mon["used"]]
        fits["corr" + key] = [f[f"unc_{el}"]["corrections"].get(k) for k in TERMS] + \
            [sorted(f[f"unc_{el}"]["corrections"])]
        ace = f[f"ace_{el}"]
        fits["ace" + key] = [ace["n"], ace["sd"], ace["basis"], ace["df"], ace["p"], ace["over"]]
    fits["blankBelow"] = [f["blank_n"]["below_detection"], f["blank_c"]["below_detection"], f["blank_n"]["capsules"]]
    return {"case": c, "rows": rows, "fits": fits}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("export")
    ap.add_argument("--config", help="lab configuration JSON (default: package defaults)")
    ap.add_argument("--out", required=True, help="folder for port_input.json and cases.json")
    ap.add_argument("--seed", type=int, default=20261011)
    ap.add_argument("--random", type=int, default=400, help="random decision sets after the 243")
    ap.add_argument("--per-variant", type=int, default=30,
                    help="random decision sets for each non-default settings variant")
    ap.add_argument("--as-filtered", action="store_true",
                    help="rewrite the export as a filtered one first (see above)")
    ap.add_argument("--with-carbon-blank", action="store_true",
                    help="give each blank capsule a sample CO2 peak first (carbon_blank_variant, 1.4.10)")
    ap.add_argument("--no-ref-background", action="store_true",
                    help="empty the N2 reference background at one sample and one drift monitor first "
                         "(no_background_variant, finding R-23)")
    ap.add_argument("--no-series", action="store_true",
                    help="give the Linearity rows of the standards table the role Exclude first, so the run has no "
                         "size series and takes the lab minimum (ea_irms 1.4.15)")
    ap.add_argument("--long-series", action="store_true",
                    help="rename ten samples to linearity standards first (long_series_variant, ea_irms 1.4.15)")
    args = ap.parse_args(argv)

    export = Path(args.export)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    base = config_from_dict(json.loads(Path(args.config).read_text())) if args.config else Config()
    if args.no_series:
        base = replace(base, standards=without_series(base.standards))
    if args.with_carbon_blank:
        export = carbon_blank_variant(export, out / f"{export.stem}_carbon_blank.csv", base.standards)
    if args.no_ref_background:
        export = no_background_variant(export, out / f"{export.stem}_no_ref_background.csv", base.standards)
        print(f"no reference background at positions {', '.join(map(str, no_background_rows(export)))}")
    if args.long_series:
        export = long_series_variant(export, out / f"{export.stem}_long_series.csv", base.standards)
    if args.as_filtered:
        export = filtered_variant(export, out / f"{export.stem}_filtered.csv")
    positions, fits = process(parse_export(export), base)
    block = port_input(positions, fits, base, export.name, export.stat().st_size)
    monitors = [q["r"] for q in block["pos"] if q["role"] == "D"]
    extra = run_variants(positions, base.standards)
    cases = [engine_case(export, base, c)
             for c in decision_sets(monitors, args.seed, args.random, args.per_variant, extra)]

    (out / "port_input.json").write_text(json.dumps(block, separators=(",", ":")))
    (out / "cases.json").write_text(json.dumps(cases))
    n_set = sum(1 for c in cases if "settings" in c["case"] or "std" in c["case"])
    n_std = sum(1 for c in cases if "std" in c["case"])
    if fits["ref_pulse_check"]["active"]:
        print(f"reference-pulse check active, {len(fits['ref_pulse_check']['rejected'])} CO2 peaks set aside")
    if args.no_series or args.long_series:
        print(f"size series of {sum(1 for p in positions if p.role == 'Linearity')} standards")
    print(f"{export.name}: {len(positions)} positions, {len(monitors)} drift monitors, "
          f"{len(cases)} decision sets ({len(cases) - n_set} with the block's settings, "
          f"{n_set} under {len(VARIANTS) + len(extra)} non-default settings variants, {n_std} of them with accepted "
          f"values for the drift monitors) -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
