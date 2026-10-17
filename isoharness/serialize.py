"""Turn ea_irms results into JSON and CSV for the API.

``port_input`` builds the block the browser's preview port (``port/engine.js``)
recomputes from: each position's inputs at full precision, the results that
no user choice changes, and the configuration's settings that act inside the
port's steps. It follows the prototype's extract script field for field, so the
port sees the same numbers it was checked against.
"""
from __future__ import annotations

import csv
import io
import math
from dataclasses import fields

from ea_irms.cells import csv_text
from ea_irms.metadata import output_columns, output_header, output_values
from ea_irms.parse import Position
from ea_irms.pipeline import ATOMIC_C, ATOMIC_N, Config
from ea_irms.workbook import VERSION

ROLE_CODES = {"Sample": "S", "Anchor": "A", "Check": "K", "Drift": "D", "Linearity": "L",
              "Blank": "B", "Exclude": "X"}


def clean(value):
    """A JSON-safe copy: NaN and infinities become null, tuples and sets lists."""
    if value is None or isinstance(value, (bool, str, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, dict):
        return {str(k): clean(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [clean(v) for v in value]
    if isinstance(value, (set, frozenset)):
        items = [clean(v) for v in value]
        try:
            return sorted(items)
        except TypeError:
            return items
    return str(value)


def _r(v, d=6):
    return None if v is None else round(v, d)


def _fit(f, keys=("slope", "intercept", "r2", "sd", "n", "se", "t", "p", "df", "mean_x", "sxx", "sse")):
    if f is None:
        return None
    return {k: (_r(f[k], 10) if isinstance(f.get(k), float) else f.get(k)) for k in keys if k in f}


def position_record(p: Position) -> dict:
    """Every attribute the engine set on a position, inputs first."""
    names = [f.name for f in fields(p)]
    names += [k for k in vars(p) if k not in names]
    return {k: clean(getattr(p, k)) for k in names}


def port_input(positions, fits, config: Config, file_name: str, size: int,
               run_date: str | None = None) -> dict:
    """The block the browser port recomputes from (the prototype's ``demo_data`` form)."""
    rows = []
    for p in positions:
        rows.append({
            "r": p.row, "id": p.identifier, "role": ROLE_CODES[p.role], "si": p.std_index,
            "amt": p.amount,
            "aN": p.area_n, "dN": p.d15n, "aC": p.area_c, "dC": p.d13c,
            "amplN": p.ampl_n, "amplC": p.ampl_c,
            "bgd": p.bgd_n, "rbgd": p.ref_bgd_n,
            "pN": p.ref_pulse_d15n, "pC": p.ref_pulse_d13c,
            # ea_irms 1.4.10: every pulse Isodat did not designate at the position, which the checks read
            "pNs": list(p.ref_pulses_d15n), "pCs": list(p.ref_pulses_d13c),
            "raN": p.ref_area_n, "raC": p.ref_area_c,
            "ugN": p.ug_n, "ugC": p.ug_c,
            "carry": bool(p.carry_flag), "rise": _r(p.carry_rise, 2), "over": bool(p.oversized),
            "set": _r(p.settling_fraction, 4),
            "rej": bool(getattr(p, "c_rejected", False)),
            # 1.4.4: the pulse's distance from the run's median, only where it flags the position
            "poN": p.pulse_dev_n if any(f.startswith("N2 reference pulse off by ") for f in p.flags) else None,
            "poC": p.pulse_dev_c if any(f.startswith("CO2 reference pulse off by ") for f in p.flags) else None,
            # 1.4.12 (finding W-09): two or more CO2 peaks came before the designated pulse, so Flags says so
            "cpre": p.c_before if (p.c_before >= 2 and p.area_c is not None) else None,
        })
    F = fits
    data = {
        "version": VERSION,
        "file": file_name, "bytes": size, "runDate": run_date,
        "std": [list(s) for s in config.standards],
        "pos": rows,
        "k": {"n": _fit(F["k_n"], ("slope", "intercept", "r2", "sd", "n")),
              "c": _fit(F["k_c"], ("slope", "intercept", "r2", "sd", "n"))},
        "limits": {k: _r(v, 4) for k, v in F["size_limits"].items()},
        "carry": {"threshold": F["carryover"]["threshold"], "oversized": F["carryover"]["oversized"],
                  "flagged": F["carryover"]["flagged"], "maxRise": _r(F["carryover"]["max_rise"], 2)},
        "settle": {k: _r(F["settling"][k], 4)
                   for k in ("floor", "excess", "settled_from", "level", "excess_level")},
        "refgas": {el: {"n": F[f"ref_gas_{el}"]["n"], "positions": F[f"ref_gas_{el}"]["positions"],
                        "mean": _r(F[f"ref_gas_{el}"]["mean"], 5),
                        "sd": _r(F[f"ref_gas_{el}"]["sd"], 5), "range": _r(F[f"ref_gas_{el}"]["range"], 4),
                        "fit": _fit(F[f"ref_gas_{el}"]["fit"], ("slope", "intercept", "p", "n")),
                        "trend": F[f"ref_gas_{el}"]["trend"]} for el in ("n", "c")},
        "monBgd": {"n": F["monitor_bgd"]["n"], "min": F["monitor_bgd"]["min"],
                   "max": F["monitor_bgd"]["max"],
                   "fit": _fit(F["monitor_bgd"]["fit"], ("slope", "intercept", "p", "n")),
                   "trend": F["monitor_bgd"]["trend"]},
        "pulseCheck": {"active": F["ref_pulse_check"]["active"],
                       "filtered": F["ref_pulse_check"]["filtered"]},
        # ea_irms 1.4.13 (finding W-13): whether the anchors of each standards-table row read as one material,
        # from the raw peaks, so no decision on the page changes it
        "anchorMaterial": _anchor_material(F.get("anchor_material")),
        "atomic": {"c": ATOMIC_C, "n": ATOMIC_N},
        # Read by port/engine.js (settings()); the names are the port's.
        "settings": {"alpha": config.significance_level, "outlierT": config.monitor_outlier_t,
                     "checkOffset": config.check_offset_permil,
                     "blankMultiple": config.min_blank_multiple, "k": config.coverage_k,
                     "elemGate": config.elemental_threshold_pct,
                     "firstDrift": config.first_drift_position,
                     # 1.4.12 (finding W-02): the calibration line's limits, read by the page's checks
                     "calSlopeNote": config.cal_slope_note, "calSlopeCaution": config.cal_slope_caution,
                     "calSdNote": config.cal_sd_note, "calSdCaution": config.cal_sd_caution,
                     # 1.4.13 (finding R-22): the reference-gas SD at or above which the run has a problem
                     "refSdLimit": config.ref_gas_sd_limit},
    }
    return clean(data)


def _anchor_material(block) -> dict | None:
    """ea_irms 1.4.13: F["anchor_material"] in the port block's names."""
    if block is None:
        return None
    pair = lambda m, key: (None if m[f"{key}_min"] is None else [m[f"{key}_min"], m[f"{key}_max"]])  # noqa: E731
    return {"level": block["level"],
            "limits": {"ratio": block["limits"]["ratio"], "spread": block["limits"]["spread"]},
            "materials": [{"name": m["name"], "si": m["index"], "n": m["n"], "ratio": pair(m, "ratio"),
                           "ratioX": m["ratio_x"], "d15n": pair(m, "d15n"), "d13c": pair(m, "d13c"),
                           "positions": list(m["positions"]), "over": list(m["over"]), "level": m["level"]}
                          for m in block["materials"]]}


# The workbook's Results sheet, column for column (ea_irms 1.4.5): each value followed by its 1 sigma.
# The CSV heads those columns 1sigma(...), since Excel misreads a Greek letter in a CSV file that has
# no byte-order mark; an Excel file heads them 1σ(...), as the workbook does (results_header).
RESULTS_HEADER = ["Row", "Identifier", "Type", "Mass (mg)", "%C", "1sigma(%C)", "%N", "1sigma(%N)",
                  "C:N molar", "1sigma(C:N molar)", "C:N mass", "1sigma(C:N mass)",
                  "d13C (permil)", "1sigma(d13C)", "d15N (permil)", "1sigma(d15N)", "Flags",
                  "Rerun", "Why", "Suggested mass (mg)",
                  "Elemental drift factor N", "Elemental drift factor C", "%C as measured", "%N as measured"]
RESULTS_ATTRS = ["row", "identifier", "role", "amount", "pct_c_reported", "u_pct_c", "pct_n_reported", "u_pct_n",
                 "cn_molar_reported", "u_cn_molar", "cn_mass_reported", "u_cn_mass",
                 "d13c_final", "u_d13c", "d15n_final", "u_d15n", None,
                 "rerun", "rerun_why", "rerun_mass",
                 "elem_factor_n", "elem_factor_c", "pct_c_measured", "pct_n_measured"]


def results_header(sigma: str = "1sigma") -> list[str]:
    """The Results columns' names, with 1sigma(...) or, for an Excel file, 1σ(...)."""
    return [h.replace("1sigma(", sigma + "(") for h in RESULTS_HEADER]


def results_values(p: Position) -> list:
    """One position's Results row: numbers as the engine left them, None where empty, the flags joined."""
    row = []
    for a in RESULTS_ATTRS:
        if a is None:
            row.append("; ".join(p.flags))
        else:
            v = clean(getattr(p, a, None))
            row.append(v)
    return row


def results_csv(positions, config: Config, sheet=None) -> str:
    """The workbook's Results sheet as CSV, at full precision, same columns and order.

    Each value is followed by its 1 sigma (ea_irms 1.4.5). The workbook heads those columns
    1σ(...); here they read 1sigma(...), since Excel misreads a Greek letter in a CSV file
    that has no byte-order mark. With a sample sheet (isoharness 0.3.0) the sheet's columns
    follow, as on the workbook's Results sheet: Sample ID (sheet), Matched by, and each column
    with a value for a sample of the run.
    """
    cols = output_columns(sheet, [positions])
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(csv_cells(RESULTS_HEADER + output_header(sheet, cols)))
    for p in positions:
        w.writerow(csv_cells(results_values(p) + output_values(p, sheet, cols)))
    return buf.getvalue()


def sample_records(positions) -> list[dict]:
    """How each Sample position matched the sample sheet (isoharness 0.3.0)."""
    return [{"row": p.row, "identifier": p.identifier, "sample_id": getattr(p, "sample_id", None),
             "matched_by": getattr(p, "matched_by", None), "group": getattr(p, "replicate_group", None)}
            for p in positions if p.role == "Sample"]


def csv_cells(values) -> list:
    """Values as the results table writes them: floats in full (repr), empty for None, and text that
    begins like a formula after an apostrophe (ea_irms 1.4.10, ea_irms.cells.csv_text)."""
    return ["" if v is None else (repr(v) if isinstance(v, float) else csv_text(v)) for v in values]
