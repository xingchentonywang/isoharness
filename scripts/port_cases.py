"""Engine outputs for the preview port to be checked against.

Runs ea_irms on one export under many decision sets (all 243 combinations of the
blank, linearity and drift switches and the drift model, then random exclusions,
elemental settings and sample repeatabilities), then under random decision sets
with non-default settings (the significance level, the monitor outlier limit,
the check-standard offset, the blank multiple, k, the elemental gate and the
first drift position), and writes two files:

    port_input.json   the "port" block of POST /api/process for this export
    cases.json        each decision set with the engine's values, flags and fits

    python3 scripts/port_cases.py tests/data/20260915_raw.csv --out build/port
    node port/test_port.js build/port/port_input.json build/port/cases.json

With --as-filtered the export is first rewritten as a filtered one (written to
the --out folder): no mass-28 reference peaks, and where a position made no
sample CO2, an undesignated CO2 pulse ahead of the designated one, as in the
2025 exports. The reference-pulse check then sets those pulses aside, which the
port must flag as the engine does.

Decisions start from --config (or the package defaults). The port reads those seven
settings from the block (port/engine.js, settings()); a case with "settings"
overrides them on both sides. The standards, the reference-pulse windows, the
size limits, settling and carryover act before the block is built, on the
server, so this check does not vary them.
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
                  "firstDrift": "first_drift_position"}
# Each variant is run under --per-variant random decision sets.
VARIANTS = [{"alpha": 0.10}, {"alpha": 0.01}, {"k": 1.0}, {"k": 3.0},
            {"elemGate": 2.0}, {"elemGate": 10.0}, {"outlierT": 2.5}, {"checkOffset": 0.2},
            {"blankMultiple": None}, {"blankMultiple": 10.0}, {"firstDrift": 20},
            {"alpha": 0.2, "k": 2.5, "elemGate": 1.0, "outlierT": 3.0, "checkOffset": 0.1,
             "blankMultiple": 20.0}]


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


def random_decisions(rng, monitors):
    return {"blank": rng.choice(SWITCH), "linN": rng.choice(SWITCH), "linC": rng.choice(SWITCH),
            "drift": rng.choice(SWITCH), "model": rng.choice(["Row", "Background", "Auto"]),
            "exclN": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 1, 2, 3])))),
            "exclC": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 2])))),
            "elem": rng.choice(SWITCH),
            "exclE": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 1, 2, 2, 3, 5])))),
            "uN": rng.choice([None, None, 0.181, 0.1]), "uC": rng.choice([None, None, 0.12])}


def decision_sets(monitors, seed, n_random, per_variant=0):
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
    return cases


def engine_case(export, base: Config, c: dict) -> dict:
    cfg = replace(base, blank_correction=c["blank"], linearity_correction=c["linN"],
                  linearity_correction_c=c["linC"], drift_correction=c["drift"], drift_model=c["model"],
                  exclude_drift_n=c["exclN"], exclude_drift_c=c["exclC"], elemental_correction=c["elem"],
                  exclude_elemental=c["exclE"], u_sample_n=c["uN"], u_sample_c=c["uC"],
                  **{SETTING_FIELDS[k]: v for k, v in c.get("settings", {}).items()})
    ps, f = process(parse_export(export), cfg)
    rows = [[p.row, p.d15n_blank, p.d15n_lin, p.d15n_drift, p.d15n_final, p.u_d15n, p.U_d15n,
             p.d13c_blank, p.d13c_lin, p.d13c_drift, p.d13c_final, p.u_d13c, p.U_d13c,
             p.pct_n_reported, p.pct_c_reported, p.cn_molar_reported, p.elem_factor_n,
             p.elem_factor_c, p.t_drift_n, p.t_drift_c, p.flags] for p in ps]
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
    args = ap.parse_args(argv)

    export = Path(args.export)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    base = config_from_dict(json.loads(Path(args.config).read_text())) if args.config else Config()
    if args.with_carbon_blank:
        export = carbon_blank_variant(export, out / f"{export.stem}_carbon_blank.csv", base.standards)
    if args.as_filtered:
        export = filtered_variant(export, out / f"{export.stem}_filtered.csv")
    positions, fits = process(parse_export(export), base)
    block = port_input(positions, fits, base, export.name, export.stat().st_size)
    monitors = [q["r"] for q in block["pos"] if q["role"] == "D"]
    cases = [engine_case(export, base, c)
             for c in decision_sets(monitors, args.seed, args.random, args.per_variant)]

    (out / "port_input.json").write_text(json.dumps(block, separators=(",", ":")))
    (out / "cases.json").write_text(json.dumps(cases))
    n_set = sum(1 for c in cases if "settings" in c["case"])
    if fits["ref_pulse_check"]["active"]:
        print(f"reference-pulse check active, {len(fits['ref_pulse_check']['rejected'])} CO2 peaks set aside")
    print(f"{export.name}: {len(positions)} positions, {len(monitors)} drift monitors, "
          f"{len(cases)} decision sets ({len(cases) - n_set} with the block's settings, "
          f"{n_set} under {len(VARIANTS)} non-default settings variants) -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
