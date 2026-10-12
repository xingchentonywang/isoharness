"""Engine outputs for the preview port to be checked against.

Runs ea_irms on one export under many decision sets (all 243 combinations of the
blank, linearity and drift switches and the drift model, then random exclusions,
elemental settings and sample repeatabilities) and writes two files:

    port_input.json   the "port" block of POST /api/process for this export
    cases.json        each decision set with the engine's values, flags and fits

    python3 scripts/port_cases.py tests/data/20260915_raw.csv --out build/port
    node port/test_port.js build/port/port_input.json build/port/cases.json

Decisions start from --config (or the package defaults). The port fixes every
other setting at its package default, so a configuration that changes those is
outside what this check covers.
"""
from __future__ import annotations

import argparse
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


def decision_sets(monitors, seed, n_random):
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
        cases.append({"blank": rng.choice(SWITCH), "linN": rng.choice(SWITCH), "linC": rng.choice(SWITCH),
                      "drift": rng.choice(SWITCH), "model": rng.choice(["Row", "Background", "Auto"]),
                      "exclN": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 1, 2, 3])))),
                      "exclC": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 0, 1, 2])))),
                      "elem": rng.choice(SWITCH),
                      "exclE": sorted(rng.sample(monitors, min(len(monitors), rng.choice([0, 1, 2, 2, 3, 5])))),
                      "uN": rng.choice([None, None, 0.181, 0.1]), "uC": rng.choice([None, None, 0.12])})
    return cases


def engine_case(export, base: Config, c: dict) -> dict:
    cfg = replace(base, blank_correction=c["blank"], linearity_correction=c["linN"],
                  linearity_correction_c=c["linC"], drift_correction=c["drift"], drift_model=c["model"],
                  exclude_drift_n=c["exclN"], exclude_drift_c=c["exclC"], elemental_correction=c["elem"],
                  exclude_elemental=c["exclE"], u_sample_n=c["uN"], u_sample_c=c["uC"])
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
    args = ap.parse_args(argv)

    export = Path(args.export)
    base = config_from_dict(json.loads(Path(args.config).read_text())) if args.config else Config()
    positions, fits = process(parse_export(export), base)
    block = port_input(positions, fits, base, export.name, export.stat().st_size)
    monitors = [q["r"] for q in block["pos"] if q["role"] == "D"]
    cases = [engine_case(export, base, c) for c in decision_sets(monitors, args.seed, args.random)]

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    (out / "port_input.json").write_text(json.dumps(block, separators=(",", ":")))
    (out / "cases.json").write_text(json.dumps(cases))
    print(f"{export.name}: {len(positions)} positions, {len(monitors)} drift monitors, "
          f"{len(cases)} decision sets -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
