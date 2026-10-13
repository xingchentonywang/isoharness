"""ea_irms 1.4.3: the uncertainty of %N, %C and C:N, in the engine and in the workbook.

The helpers are checked against hand computations, the engine's budget against
its own terms recomputed here, and the workbook against the engine under five
configurations, after LibreOffice has computed the formulas (skipped without it).
"""
import math
import shutil
import subprocess
from pathlib import Path

import openpyxl
import pytest

from ea_irms.parse import parse_export
from ea_irms.pipeline import (DEFAULT_STANDARDS, Config, _interp_multiplier, _kfit_rel, _mssd,
                              config_from_dict, config_to_dict, process)
from ea_irms.workbook import build_workbook

RUN = Path(__file__).resolve().parent / "data" / "20260915_raw.csv"
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def _standards_with_certificates():
    rows = [list(r) for r in DEFAULT_STANDARDS]
    for r in rows:
        if r[6] == "glycine":
            r[9], r[10] = 0.10, 0.06
        elif r[6] and r[6].startswith("acetanilide"):
            r[9], r[10] = 0.15, 0.03
    return rows


def test_successive_differences_by_hand():
    assert _mssd([1.0, 2.0]) is None
    assert math.isclose(_mssd([1.0, 2.0, 4.0]), math.sqrt((1 + 4) / 4))
    # a straight ramp has scatter only in its steps, not in its spread
    ramp = [1.0 + 0.01 * i for i in range(10)]
    assert math.isclose(_mssd(ramp), 0.01 / math.sqrt(2))


def test_interpolation_multiplier_by_hand():
    assert _interp_multiplier(5, [10, 20]) == 1.0          # before the first monitor
    assert _interp_multiplier(10, [10, 20]) == 1.0         # at a monitor
    assert math.isclose(_interp_multiplier(15, [10, 20]), 0.5)
    assert math.isclose(_interp_multiplier(12, [10, 20]), 0.8 ** 2 + 0.2 ** 2)
    assert _interp_multiplier(25, [10, 20]) == 1.0         # after the last
    assert _interp_multiplier(15, []) == 1.0


def test_kfit_term_by_hand():
    fit = {"slope": 2.0, "intercept": 1.0, "sd": 0.5, "n": 4, "mean_x": 10.0, "sxx": 50.0}
    assert math.isclose(_kfit_rel(fit, 10.0), (0.5 / 2.0) * math.sqrt(1 / 4) / 10.0)
    assert math.isclose(_kfit_rel(fit, 20.0), (0.5 / 2.0) * math.sqrt(1 / 4 + 100 / 50) / 20.0)
    assert _kfit_rel(fit, 0.0) is None and _kfit_rel(dict(fit, sd=None), 10.0) is None


def test_budget_terms_recomputed_from_the_positions():
    cfg = config_from_dict({"balance_u_mg": 0.001, "standards": _standards_with_certificates()})
    positions, F = process(parse_export(RUN), cfg)
    u = F["unc_elem"]
    mons = sorted((p for p in positions if p.role == "Drift" and p.pct_n_measured is not None),
                  key=lambda p: p.row)
    r = [p.pct_n_measured / p.pct_n for p in mons]
    s_rep = math.sqrt(sum((b - a) ** 2 for a, b in zip(r, r[1:])) / (2 * (len(r) - 1)))
    w_mon = sum((0.001 / p.amount) ** 2 for p in mons) / len(mons)
    assert math.isclose(u["n"]["s_rep"], s_rep) and math.isclose(u["n"]["s_inst"], math.sqrt(s_rep ** 2 - w_mon))
    assert u["n"]["in_force"] and not u["c"]["in_force"]
    assert math.isclose(u["n"]["cert_scale"], 0.03 / 10.36)                 # the monitors set the N scale
    anchors = [p for p in positions if p.role == "Anchor" and p.pct_n_reported is not None]
    rms = math.sqrt(sum((p.pct_n_reported / p.pct_n - 1) ** 2 for p in anchors) / len(anchors))
    assert math.isclose(u["n"]["trueness"], math.sqrt(rms ** 2 + (0.06 / 18.66) ** 2))
    p = next(q for q in positions if q.row == 27)                         # a sample between monitors 26 and 30
    p11 = next(q for q in positions if q.row == 11)                        # an anchor between 10 and 14
    mult = (1 - 0.25) ** 2 + 0.25 ** 2
    rel = math.sqrt(u["n"]["s_inst"] ** 2 + (0.001 / p11.amount) ** 2 + mult * s_rep ** 2
                    + u["n"]["cert_scale"] ** 2 + u["n"]["trueness"] ** 2)
    assert math.isclose(p11.u_pct_n, p11.pct_n_reported * rel) and math.isclose(p11.U_pct_n, 2 * p11.u_pct_n)
    assert p.u_pct_n is None and p.u_cn_molar is not None                  # the samples carry no mass
    assert math.isclose(p.U_cn_mass / p.cn_mass_reported, p.U_cn_molar / p.cn_molar_reported)


def test_inputs_left_empty_count_as_zero_and_light_samples_cost_more():
    base, F0 = process(parse_export(RUN), Config())
    assert F0["unc_elem"]["n"]["weighing_monitors"] == 0 and F0["unc_elem"]["n"]["cert_scale"] == 0
    weighed, _ = process(parse_export(RUN), config_from_dict({"balance_u_mg": 0.002}))
    lin = {p.row: (p, q) for p, q in zip(base, weighed) if p.role == "Linearity" and p.u_pct_n is not None}
    small = min(lin.values(), key=lambda t: t[0].amount)
    large = max(lin.values(), key=lambda t: t[0].amount)
    grow = lambda t: t[1].u_pct_n / t[1].pct_n_reported - t[0].u_pct_n / t[0].pct_n_reported   # noqa: E731
    # The monitors' own weighing comes out of the repeatability, so a sample heavier than the monitors
    # can end with a smaller u than before; a lighter one always gains.
    assert grow(small) > 0 and grow(small) > grow(large)


def test_too_few_anchors_or_monitors():
    std = [list(r) for r in DEFAULT_STANDARDS]
    for r in std:
        if r[1] == "Anchor":
            r[1] = "Check"                                  # no anchors: no scale, no trueness
    _, F = process(parse_export(RUN), config_from_dict({"standards": std}))
    assert F["unc_elem"]["n"]["trueness"] is None and F["unc_elem"]["n"]["n_anchors"] == 0
    rows = [p.row for p in parse_export(RUN)]
    positions, F = process(parse_export(RUN), Config(exclude_elemental=[r for r in rows if r > 14]))
    assert F["unc_elem"]["n"]["n_monitors"] == 2 and F["unc_elem"]["n"]["s_rep"] is None
    assert all(p.u_pct_n is None for p in positions)


def test_configuration_round_trip_keeps_the_new_fields():
    cfg = config_from_dict({"balance_u_mg": 0.0015, "standards": _standards_with_certificates()})
    back = config_from_dict(config_to_dict(cfg))
    assert back.balance_u_mg == 0.0015 and back.standards == cfg.standards
    old = config_from_dict({"standards": [list(r)[:9] for r in DEFAULT_STANDARDS]})   # written by 1.4.2
    assert all(len(r) == 11 and r[9] is None and r[10] is None for r in old.standards)


CONFIGS = {
    "default": Config(),
    "certs": config_from_dict({"balance_u_mg": 0.002, "standards": _standards_with_certificates()}),
    "both_in_force": config_from_dict({"elemental_correction": "On", "balance_u_mg": 0.001,
                                       "standards": _standards_with_certificates()}),
    "off": config_from_dict({"elemental_correction": "Off", "balance_u_mg": 0.001}),
    "k3_excluded": config_from_dict({"coverage_k": 3.0, "exclude_elemental": [18, 55]}),
}
RESULTS = {"1σ(%C)": "u_pct_c", "1σ(%N)": "u_pct_n", "1σ(C:N molar)": "u_cn_molar", "1σ(C:N mass)": "u_cn_mass",
           "1σ(d13C)": "u_d13c", "1σ(d15N)": "u_d15n", "%N": "pct_n_reported", "C:N mass": "cn_mass_reported"}
FITS = {"n": range(303, 312), "c": range(316, 325)}
TERMS = ("n_monitors", "s_rep", "weighing_monitors", "s_inst", "cert_scale", "n_anchors", "anchor_rms",
         "anchor_cert", "trueness")
CN_ROWS = {329: "n_monitors", 330: "s_rep", 331: "cert_scale", 332: "n_anchors", 333: "anchor_rms",
           334: "anchor_cert", 335: "trueness"}


@pytest.fixture(scope="module")
def computed(tmp_path_factory):
    if not SOFFICE:
        pytest.skip("needs LibreOffice")
    folder = tmp_path_factory.mktemp("unc")
    files = []
    for name, cfg in CONFIGS.items():
        out = folder / f"{name}.xlsx"
        build_workbook(RUN, out, cfg)
        files.append(out)
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(folder / "calc")]
                   + [str(f) for f in files], check=True, capture_output=True, timeout=600)
    return folder / "calc"


def _close(mine, theirs):
    theirs = theirs if isinstance(theirs, (int, float)) and not isinstance(theirs, bool) else None
    if mine is None:
        return theirs is None
    return theirs is not None and abs(mine - theirs) <= 1e-9 * max(1.0, abs(mine))


@pytest.mark.parametrize("name", list(CONFIGS))
def test_workbook_agrees_with_engine(computed, name):
    positions, F = process(parse_export(RUN), CONFIGS[name])
    wb = openpyxl.load_workbook(computed / f"{name}.xlsx", data_only=True)
    rs, ft, qc = wb["Results"], wb["Fits"], wb["QC"]
    head = {str(rs.cell(4, c).value).split(",")[0]: c for c in range(1, rs.max_column + 1)}
    k = CONFIGS[name].coverage_k
    order = [str(rs.cell(4, c).value).split(",")[0] for c in range(5, 17)]
    assert order == ["%C", "1σ(%C)", "%N", "1σ(%N)", "C:N molar", "1σ(C:N molar)", "C:N mass", "1σ(C:N mass)",
                     "d13C (permil)", "1σ(d13C)", "d15N (permil)", "1σ(d15N)"]
    assert [rs.cell(4, c).value for c in range(21, 25)] == ["Elemental drift factor N", "Elemental drift factor C",
                                                            "%C as measured", "%N as measured"]
    assert rs.cell(4, 25).value is None                       # 1.4.5: no block of u columns after these
    compared = 0
    for p in positions:
        for h, attr in RESULTS.items():
            mine = getattr(p, attr)
            assert _close(mine, rs.cell(4 + p.row, head[h]).value), (name, p.row, h, mine)
            compared += mine is not None
    assert compared > 200
    u = F["unc_elem"]
    for el, rows in FITS.items():
        for key, r in zip(TERMS, rows):
            mine, theirs = u[el][key], ft.cell(r, 2).value
            assert (theirs in (None, "")) if mine is None else _close(mine, theirs), (name, el, key, mine, theirs)
    for r, key in CN_ROWS.items():
        mine, theirs = u["cn"][key], ft.cell(r, 2).value
        assert (theirs in (None, "")) if mine is None else _close(mine, theirs), (name, "cn", key, mine, theirs)
    within = sum(1 for p in positions if p.role == "Anchor" and p.U_pct_n is not None
                 and abs(p.pct_n_reported - p.pct_n) <= p.U_pct_n)
    total = sum(1 for p in positions if p.role == "Anchor" and p.U_pct_n is not None)
    assert qc["B155"].value == f"{within} of {total}"
    assert qc["A155"].value == f"Anchors within {k:g}σ, %N"
    errors = [c.coordinate for s in (rs, ft, qc) for row in s.iter_rows() for c in row
              if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper()]
    assert not errors, errors[:5]
