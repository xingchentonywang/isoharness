"""Corrections applied to one EA-IRMS sequence.

The order is blank, linearity, drift, calibration. Every coefficient is fitted
from the sequence's own standards. Calibration is fitted as accepted regressed
on measured and applied directly; it is never fitted the other way and inverted,
which is the arithmetic that fails silently.

From 1.3.0 a correction is applied only when the sequence supports it. On
``Auto`` a drift or linearity slope is applied when it differs from zero at the
significance level in the configuration, which follows the FIRMS guide's advice
to establish that a correction is necessary before applying it (Dunn and
Carter, 2018, section 6.4). Each reported value also carries a standard
uncertainty from the scale calibration and the sample's own repeatability
(section 6.5), computed the way a Kragten spreadsheet would, by the effect of
each input's uncertainty on the result.
"""
from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal

from .metadata import critical_range_factor, match_positions, match_summary
from .parse import Position

ATOMIC_C = 12.011
ATOMIC_N = 14.007

# Settling is found at the start of the run (1.4.0). The scale is the excess of
# the first SETTLE_START reference-peak backgrounds over the floor; a run whose
# scale is under SETTLE_MIN_EXCESS mV is settled from the start; the settled
# position is the first of SETTLE_RUN consecutive readings at or under the
# threshold. A position within CARRY_REACH of an oversized sample is "exposed":
# its reading is skipped in that search, and it is flagged for possible
# carryover when its sample-peak background rose by more than Config.carryover_mv.
SETTLE_START = 5
SETTLE_MIN_EXCESS = 1.0
SETTLE_RUN = 3
CARRY_REACH = 2
# 1.4.2: settling is a start above the run's own level, so the start is also
# measured against the median reference-peak background; a run whose first
# readings sit within SETTLE_MIN_EXCESS of that median was settled from the
# start, whatever its floor (the 2026-09-22 full export was flat from position
# 5 and still counted as never settled, because its floor came from the run's
# last readings).
MIN_OUTLIER_MONITORS = 5
# Excel and LibreOffice treat two numbers within about 1e-15 of each other as
# equal when comparing; the engine allows the same slack so the two agree at a
# threshold.
_SLACK = 1e-9


def _above(a, b):
    """a > b, allowing the rounding slack a spreadsheet comparison allows."""
    return a - b > _SLACK * max(1.0, abs(b))

DEFAULT_STANDARDS = [
    # match text, role, accepted d13C, accepted d15N, %C, %N, material, u(d13C), u(d15N),
    # and from 1.4.3 u(%C), u(%N), the 1-sigma certificate uncertainties of the elemental
    # content, left empty for the lab to fill from its certificates (empty counts as 0).
    # Uncertainties are the 1-sigma values listed by the producer (Indiana
    # University reference-material list, Schimmelmann group); USGS40 and USGS41a
    # are left for the lab to fill from its certificates.
    ("USGS 64", "Anchor", -40.81, 1.76, 32.00, 18.66, "glycine", 0.04, 0.06, None, None),
    ("USGS 65", "Anchor", -20.29, 20.68, 32.00, 18.66, "glycine", 0.04, 0.06, None, None),
    ("USGS 66", "Check", -0.67, 40.83, 32.00, 18.66, "glycine", 0.04, 0.06, None, None),
    ("USGS 40", "Check", -26.39, -4.52, 40.82, 9.52, "L-glutamic acid", None, None, None, None),
    ("USGS 41a", "Check", 36.55, 47.55, 40.82, 9.52, "L-glutamic acid", None, None, None, None),
    ("Ace Lin", "Linearity", None, None, 71.09, 10.36, "acetanilide, size series", None, None, None, None),
    ("Ace Drift", "Drift", None, None, 71.09, 10.36, "acetanilide, drift monitor", None, None, None, None),
    ("Ace Bypass", "Exclude", None, None, 71.09, 10.36, "bypass test", None, None, None, None),
    ("Blank Capsule", "Blank", None, None, None, None, "tin capsule, nothing in it", None, None, None, None),
    ("Blank Hole", "Exclude", None, None, None, None, "empty carousel position", None, None, None, None),
    # 1.4.2: sequences that spell the acetanilide out in full (the Bocas runs of
    # 2025 and 2026) match without a Setup edit.
    ("Acetanilide Lin", "Linearity", None, None, 71.09, 10.36, "acetanilide, size series", None, None, None, None),
    ("Acetanilide Drift", "Drift", None, None, 71.09, 10.36, "acetanilide, drift monitor", None, None, None, None),
    ("Acetanilide Bypass", "Exclude", None, None, 71.09, 10.36, "bypass test", None, None, None, None),
]
STANDARD_FIELDS = 11
# The workbook's standards table has this many rows (1.4.2: 16, from 13); a
# configuration with more cannot be written as a workbook, and build_workbook
# refuses it instead of dropping the extra rows silently.
MAX_STANDARDS = 16


@dataclass
class Config:
    standards: list = field(default_factory=lambda: list(DEFAULT_STANDARDS))
    blank_correction: str = "Auto"      # On | Off | Auto
    linearity_correction: str = "Auto"  # On | Off | Auto; nitrogen, and carbon unless set below
    linearity_correction_c: str | None = None  # On | Off | Auto for carbon; None = as above (1.4.1)
    drift_correction: str = "Auto"      # On | Off | Auto
    drift_model: str = "Row"            # Row | Background | Auto
    first_drift_position: int | None = None
    exclude_drift_n: list = field(default_factory=list)   # positions left out of the N drift fit
    exclude_drift_c: list = field(default_factory=list)   # positions left out of the C drift fit
    significance_level: float = 0.05
    monitor_outlier_t: float = 4.0      # studentized residual that flags a drift monitor
    check_offset_permil: float = 0.5
    min_area_n: float | None = None
    min_area_c: float | None = None
    max_area_n: float | None = None
    max_area_c: float | None = None
    min_blank_multiple: float | None = 3.0
    settled_threshold: float | None = 0.10
    carryover_mv: float | None = 2.0    # background rise that flags possible carryover; None = off
    u_sample_n: float | None = None     # single-analysis repeatability, 1 SD; None = pooled anchors
    u_sample_c: float | None = None
    # 1.4.10: where u_sample_n and u_sample_c came from. "Drift monitors" means the value is the SD of the
    # drift monitors of the run named by run_file, so it belongs to that run (scope_to_run); None means the
    # value, if any, is the lab's own and holds for every run.
    u_sample_n_source: str | None = None
    u_sample_c_source: str | None = None
    coverage_k: float = 2.0             # the checks compare against k x 1 sigma (1.4.5)
    # 1.4.3: standard uncertainty of one weighing, mg; None = not given, counted as 0
    balance_u_mg: float | None = None
    # 1.4.2: %N and %C drift, read from the drift monitors' own elemental content
    elemental_correction: str = "Auto"  # On | Off | Auto
    elemental_threshold_pct: float = 5.0   # Auto applies it when a monitor's relative bias exceeds this
    exclude_elemental: list = field(default_factory=list)   # positions left out of the elemental correction
    # 1.4.2: an unflagged CO2 peak in an export without reference peaks that
    # matches the designated reference pulses is a pulse, not the sample
    ref_pulse_ampl_tol: float | None = 0.02   # fraction of the designated pulses' amplitude at that position; None = off
    ref_pulse_d13c_window: float = 0.75       # permil around the designated pulses' mean d13C
    # 1.4.4: rerun marks. A position's undesignated reference pulse is off when it lies more than
    # this many robust SDs from the run's median (None = off), and a reference-gas SD at or above
    # ref_gas_sd_limit (permil; the FIRMS guide's zero-enrichment limit, None = off) is a run-level
    # problem on the Rerun sheet.
    ref_pulse_off_sd: float | None = 4.0
    ref_gas_sd_limit: float | None = 0.1
    # 1.4.7 (Tony, 2026-10-13): a peak above the linearity range is marked for a rerun only when its
    # amplitude reaches the instrument's largest, in mV (None: every peak above the range is marked,
    # as before 1.4.7); a peak at or above it is marked whatever the range. A sample whose final d15N
    # or d13C has a 1 sigma above rerun_max_1sigma (permil) is marked for a rerun of that element
    # (None = off).
    max_amplitude_mv: float | None = 50000.0
    rerun_max_1sigma: float | None = 0.5
    # 1.4.8: replicates of a sample in one run, named by the sample sheet, whose range is above the
    # critical range of ISO 5725-6 (f(n) x the repeatability of one analysis) are marked for a rerun of
    # that element. On or Off; without a sample sheet nothing is compared.
    rerun_replicates: str = "On"
    # 1.4.10: the export the first drift position, the positions left out, and a repeatability taken from
    # the drift monitors were chosen for. They are applied only to a run of that name; with no name they
    # are applied to any run, with a note (scope_to_run; finding R-08 of the 2026-10-14 review).
    run_file: str | None = None

    def __post_init__(self):
        _normalize(self)


MONITORS = "Drift monitors"       # the one source of a repeatability that belongs to one run (1.4.10)

# The switches and the words each takes. From 1.4.10 they are read in any case, so "off" is Off in the
# engine as it already was in the workbook, where Excel compares text without regard to case; a word
# that is none of these is refused rather than read as Auto (finding R-07 of the 2026-10-14 review).
CHOICES = {"blank_correction": ("On", "Off", "Auto"), "linearity_correction": ("On", "Off", "Auto"),
           "linearity_correction_c": ("On", "Off", "Auto"), "drift_correction": ("On", "Off", "Auto"),
           "drift_model": ("Row", "Background", "Auto"), "elemental_correction": ("On", "Off", "Auto"),
           "rerun_replicates": ("On", "Off"), "u_sample_n_source": (MONITORS,), "u_sample_c_source": (MONITORS,)}
OPTIONAL_CHOICES = ("linearity_correction_c", "u_sample_n_source", "u_sample_c_source")   # empty means none
MAX_RUN_FILE = 255


def _normalize(config):
    """Put each switch in its own spelling; older configurations stored some as booleans."""
    for name, allowed in CHOICES.items():
        v = getattr(config, name)
        if v is True and "On" in allowed:
            v = "On"
        elif v is False and "Off" in allowed:
            v = "Off"
        if name in OPTIONAL_CHOICES and (v is None or (isinstance(v, str) and not v.strip())):
            setattr(config, name, None)
            continue
        hit = next((a for a in allowed if isinstance(v, str) and v.strip().casefold() == a.casefold()), None)
        if hit is None:
            raise ValueError(f"{name} is {v!r}; it needs one of {', '.join(allowed)}"
                             + (" or nothing" if name in OPTIONAL_CHOICES else ""))
        setattr(config, name, hit)
    run = config.run_file
    if run is not None and (isinstance(run, bool) or not isinstance(run, (str, int, float))):
        raise ValueError(f"run_file is {run!r}; it needs the export's file name, or nothing")
    run = str(run).strip() if run is not None else ""
    if len(run) > MAX_RUN_FILE:
        raise ValueError(f"run_file holds {len(run)} characters; a file name has at most {MAX_RUN_FILE}")
    config.run_file = run or None


# ------------------------------------------------- values that belong to one run (1.4.10)
RUN_ONLY = ("first_drift_position", "exclude_drift_n", "exclude_drift_c", "exclude_elemental")
EXPORT_SUFFIXES = (".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".tsv")      # the exports ea_irms reads


def run_key(name) -> str:
    """A run's file name as a configuration compares it: without folders or the export's extension, in any
    case and in Unicode's composed form, so that the same run exported as .xls or as .csv is the same run.
    Two different runs exported under one name are one run to this comparison."""
    base = re.split(r"[\\/]", unicodedata.normalize("NFC", str(name or "")).strip())[-1].strip()
    low = base.casefold()
    for suffix in EXPORT_SUFFIXES:
        if low.endswith(suffix) and len(low) > len(suffix):
            base = base[:-len(suffix)]
            break
    return base.strip().casefold()


def _permil(value) -> str:
    try:
        return f"{_xl(float(value), 3)} permil"
    except (TypeError, ValueError):
        return f"{value} permil"


def _run_only(config) -> list[str]:
    """What the configuration sets for one run only, each as a phrase for a note; empty when nothing."""
    out = []
    if config.first_drift_position:
        out.append(f"the first drift position ({config.first_drift_position})")
    for key, what in (("exclude_drift_n", "the nitrogen drift fit"), ("exclude_drift_c", "the carbon drift fit"),
                      ("exclude_elemental", "the elemental correction")):
        rows = getattr(config, key)
        if rows:
            out.append(f"the position{'s' if len(rows) > 1 else ''} left out of {what} ({', '.join(map(str, rows))})")
    for el, iso in (("n", "d15N"), ("c", "d13C")):
        value = getattr(config, f"u_sample_{el}")
        if value is not None and getattr(config, f"u_sample_{el}_source") == MONITORS:
            out.append(f"the {iso} repeatability from the drift monitors ({_permil(value)})")
    return out


def _and(items: list[str]) -> str:
    return items[0] if len(items) == 1 else " and ".join(items) if len(items) == 2 \
        else ", ".join(items[:-1]) + ", and " + items[-1]


def _cleared(config):
    """The configuration without its run-only values; a repeatability's source is kept, its value not."""
    from dataclasses import replace
    scoped = replace(config, first_drift_position=None, exclude_drift_n=[], exclude_drift_c=[],
                     exclude_elemental=[], standards=list(config.standards))
    for el in ("n", "c"):
        if getattr(config, f"u_sample_{el}_source") == MONITORS:
            setattr(scoped, f"u_sample_{el}", None)
    return scoped


def scope_to_run(config, file_name, shared: bool = False) -> tuple:
    """The configuration as it applies to the run in ``file_name``, and notes on what was not applied.

    The first drift position, the positions left out, and a repeatability taken from the drift monitors are
    decisions about one run's positions. A configuration that names its run (run_file, written by
    isoharness.org from 1.4.10) applies them to that run only; for any other run they are cleared, and the
    repeatability's source is kept, so the page takes that run's own monitors. A configuration that names no
    run applies them to the run it is used with, as before 1.4.10, with a note asking whether they belong to
    it; given to several runs at once (``shared``), it applies them to none (finding R-08 of the 2026-10-14
    review: one run's exclusions had been carried to every run of a batch). A repeatability saved before
    1.4.10 carries no source, so it is taken as the lab's own figure."""
    held = _run_only(config)
    if not held:
        return config, []
    name = str(file_name or "").strip() or "this run"
    many = len(held) > 1 or held[0].startswith("the positions")
    items = _and(held)
    if not config.run_file:
        if shared:
            return _cleared(config), [f"no run is named for {items}, and the configuration was given to several "
                                      f"runs at once, so {'they were' if many else 'it was'} applied to none of them"]
        return config, [f"no run is named for {items}, so {'they were' if many else 'it was'} applied to "
                        f"{name}; check that {'they belong' if many else 'it belongs'} to this run"]
    if run_key(config.run_file) == run_key(file_name):
        return config, []
    return _cleared(config), [f"{items} {'were' if many else 'was'} chosen for {config.run_file}, so "
                              f"{'they were' if many else 'it was'} not applied to {name}"]


# ----------------------------------------------------------------- statistics
def _betacf(a, b, x):
    """Continued fraction for the incomplete beta function (Lentz's method)."""
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < tiny:
        d = tiny
    d = 1.0 / d
    h = d
    for m in range(1, 300):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        de = d * c
        h *= de
        if abs(de - 1.0) < 3e-16:
            break
    return h


def _betainc(a, b, x):
    """Regularised incomplete beta function I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lbeta = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
    front = math.exp(lbeta + a * math.log(x) + b * math.log(1.0 - x))
    if x < (a + 1.0) / (a + b + 2.0):
        return front * _betacf(a, b, x) / a
    return 1.0 - front * _betacf(b, a, 1.0 - x) / b


def t_p_two_sided(t, df):
    """Two-sided p-value of Student's t, the same number Excel's TDIST(|t|, df, 2) gives."""
    if df is None or df < 1 or t is None:
        return None
    x = df / (df + t * t)
    return _betainc(df / 2.0, 0.5, x)


def _fit(x, y):
    """Least-squares line with the statistics a correction decision needs.

    Returns slope, intercept, r2, residual SD, n, and also the slope's standard
    error, its t statistic and two-sided p-value, the mean of x, Sxx and the
    residual sum of squares, which the outlier and uncertainty calculations use.
    """
    pairs = [(a, b) for a, b in zip(x, y) if a is not None and b is not None]
    n = len(pairs)
    if n < 2:
        return None
    mx = sum(a for a, _ in pairs) / n
    my = sum(b for _, b in pairs) / n
    sxx = sum((a - mx) ** 2 for a, _ in pairs)
    if sxx == 0:
        return None
    sxy = sum((a - mx) * (b - my) for a, b in pairs)
    slope = sxy / sxx
    inter = my - slope * mx
    syy = sum((b - my) ** 2 for _, b in pairs)
    res = sum((b - (slope * a + inter)) ** 2 for a, b in pairs)
    r2 = 1 - res / syy if syy > 0 else None
    sd = math.sqrt(res / (n - 2)) if n > 2 else None
    se = sd / math.sqrt(sxx) if sd is not None else None
    t = slope / se if se else None
    p = t_p_two_sided(t, n - 2) if t is not None else None
    return {"slope": slope, "intercept": inter, "r2": r2, "sd": sd, "n": n,
            "se": se, "t": t, "p": p, "df": n - 2, "mean_x": mx, "sxx": sxx, "sse": res}


def _sd(values):
    vals = [v for v in values if v is not None]
    if len(vals) < 2:
        return None
    m = sum(vals) / len(vals)
    return math.sqrt(sum((v - m) ** 2 for v in vals) / (len(vals) - 1))


def _apply(mode, fit, n_used, alpha, minimum=3):
    """Whether a fitted slope is applied under the On / Off / Auto switch."""
    if mode == "Off" or fit is None or n_used < minimum:
        return False
    if mode == "On":
        return True
    return fit.get("p") is not None and fit["p"] < alpha


def assign_roles(positions, config):
    """Attach role and certified values, by case-insensitive substring match."""
    for p in positions:
        ident = p.identifier.lower()
        hit = None
        for entry in config.standards:
            text = (entry[0] or "").strip()
            if text and text.lower() in ident:
                hit = entry
        if hit is None:
            p.role, p.acc13, p.acc15, p.pct_c, p.pct_n = "Sample", None, None, None, None
            p.u_acc13 = p.u_acc15 = None
            p.u_pct_c_cert = p.u_pct_n_cert = None
            p.std_index = None
        else:
            entry = tuple(hit) + (None,) * (STANDARD_FIELDS - len(hit))
            p.role, p.acc13, p.acc15, p.pct_c, p.pct_n = entry[1], entry[2], entry[3], entry[4], entry[5]
            p.u_acc13, p.u_acc15 = entry[7], entry[8]
            p.u_pct_c_cert, p.u_pct_n_cert = entry[9], entry[10]
            p.std_index = config.standards.index(hit)
    return positions


def _studentized(fit, x, y, included):
    """Externally studentized residual of one drift monitor against the others.

    For a monitor in the fit this is the leave-one-out residual divided by the
    residual SD of the other monitors, computed from the full fit in closed form.
    For an excluded monitor it is its distance from the fitted line in units of
    the prediction standard deviation. With fewer than MIN_OUTLIER_MONITORS in
    the fit the leave-one-out test rests on one or two degrees of freedom and is
    not run (1.4.2; with four monitors it flagged residuals of 0.007 permil).
    """
    if fit is None or x is None or y is None:
        return None
    n = fit["n"]
    e = y - (fit["slope"] * x + fit["intercept"])
    h = 1.0 / n + (x - fit["mean_x"]) ** 2 / fit["sxx"]
    if included:
        if n < MIN_OUTLIER_MONITORS or h >= 1.0:
            return None
        s2 = (fit["sse"] - e * e / (1.0 - h)) / (n - 3)
        if s2 <= 0:
            return None
        return e / (math.sqrt(s2) * math.sqrt(1.0 - h))
    if fit["sd"] is None or fit["sd"] <= 0:
        return None
    return e / (fit["sd"] * math.sqrt(1.0 + h))


def _mssd(values):
    """Single-analysis scatter from successive differences, sqrt(sum d^2 / (2 (n - 1))).

    Differences between neighbours in position leave out a drift slow next to the
    spacing of the monitors, which an SD about the mean would count as scatter.
    None below three values.
    """
    if len(values) < 3:
        return None
    return math.sqrt(sum((b - a) ** 2 for a, b in zip(values, values[1:])) / (2 * (len(values) - 1)))


def _interp_multiplier(row, rows):
    """Variance multiplier of a factor interpolated in position between two monitors.

    (1 - w)^2 + w^2 with w the fraction of the way from the monitor before to the one
    after; 1 at a monitor and beyond the first and last, where one monitor sets it.
    """
    if not rows or row <= rows[0] or row >= rows[-1]:
        return 1.0
    for x0, x1 in zip(rows, rows[1:]):
        if x0 <= row <= x1:
            w = (row - x0) / (x1 - x0)
            return (1 - w) ** 2 + w ** 2
    return 1.0


def _kfit_rel(fit, ug):
    """Relative standard uncertainty of micrograms read off the k-factor line at ``ug``.

    The calibration part of an inverse prediction, (s / |b|) sqrt(1/n + (x0 - mean x)^2 / Sxx),
    over x0; the single analysis's own scatter is the repeatability term.
    """
    if fit is None or fit.get("sd") is None or not fit.get("sxx") or not fit.get("slope") or not ug or ug <= 0:
        return None
    u = (fit["sd"] / abs(fit["slope"])) * math.sqrt(1.0 / fit["n"] + (ug - fit["mean_x"]) ** 2 / fit["sxx"])
    return u / ug


def _rms(values):
    vals = [v for v in values if v is not None]
    return math.sqrt(sum(v * v for v in vals) / len(vals)) if vals else None


def _elemental_uncertainty(positions, F, config, excl_e):
    """u and U of every %N, %C, and C:N, from five terms combined in quadrature.

    For %N and %C, as fractions of the value: the instrument's repeatability, the
    drift monitors' measured over certified ratios by successive differences in
    position, less the part their own weighing explains; this position's weighing,
    the balance's u over its mass; the calibration, the interpolated elemental
    drift factor when that correction is in force and the k-factor line at this
    position's micrograms when it is not; the certificate uncertainty of the
    material that sets the scale, the monitors' when the correction is in force and
    the weighed standards' otherwise; and trueness, the anchors' RMS deviation from
    their certified values with their certificate uncertainty (the Nordtest TR 537
    approach). C:N takes the same terms without the weighing, which cancels in a
    ratio, its repeatability from the monitors' own C:N. U = k x u.
    """
    u_m = config.balance_u_mg or 0.0
    k_cov = config.coverage_k
    cert_rel = {}
    for el in ("n", "c"):
        def rel(p, el=el):
            u, v = getattr(p, f"u_pct_{el}_cert", None), getattr(p, f"pct_{el}")
            return (u / v) if (u is not None and v) else 0.0
        cert_rel[el] = rel
    out = {"balance_u_mg": config.balance_u_mg, "k": k_cov}
    for el in ("n", "c"):
        meas, cert = f"pct_{el}_measured", f"pct_{el}"
        mons = sorted((p for p in positions if p.role == "Drift" and p.row not in excl_e
                       and getattr(p, meas) is not None and getattr(p, cert)), key=lambda p: p.row)
        ratios = [getattr(p, meas) / getattr(p, cert) for p in mons]
        s_rep = _mssd(ratios)
        w_mon = (sum((u_m / p.amount) ** 2 for p in mons) / len(mons)) if mons else 0.0
        s_inst = None if s_rep is None else math.sqrt(max(0.0, s_rep ** 2 - w_mon))
        in_force = F[f"elem_{el}"]["in_force"]
        if in_force:
            scale = _rms([cert_rel[el](p) for p in mons]) or 0.0
        else:
            kstd = [p for p in positions if p.role not in ("Blank", "Exclude", "Sample") and p.amount
                    and getattr(p, cert) is not None and getattr(p, f"area_{el}") is not None]
            scale = _rms([cert_rel[el](p) for p in kstd]) or 0.0
        anchors = [p for p in positions if p.role == "Anchor" and getattr(p, f"pct_{el}_reported") is not None
                   and getattr(p, cert)]
        devs = [getattr(p, f"pct_{el}_reported") / getattr(p, cert) - 1.0 for p in anchors]
        rms_dev = _rms(devs) if len(devs) >= 2 else None
        cref = _rms([cert_rel[el](p) for p in anchors]) if anchors else None
        trueness = None if rms_dev is None else math.sqrt(rms_dev ** 2 + (cref or 0.0) ** 2)
        out[el] = {"n_monitors": len(mons), "monitor_rows": [p.row for p in mons], "s_rep": s_rep,
                   "weighing_monitors": math.sqrt(w_mon), "s_inst": s_inst, "in_force": in_force,
                   "cert_scale": scale, "n_anchors": len(anchors), "anchor_rms": rms_dev,
                   "anchor_cert": cref, "trueness": trueness}

    # C:N, from the monitors' own measured C:N over the certified one
    def cn_cert(p):
        return (p.pct_c / ATOMIC_C) / (p.pct_n / ATOMIC_N)
    cn_mons = sorted((p for p in positions if p.role == "Drift" and p.row not in excl_e
                      and p.cn_molar is not None and p.pct_c and p.pct_n), key=lambda p: p.row)
    cn_rep = _mssd([p.cn_molar / cn_cert(p) for p in cn_mons])
    cn_anchors = [p for p in positions if p.role == "Anchor" and p.cn_molar_reported is not None
                  and p.pct_c and p.pct_n]
    cn_devs = [p.cn_molar_reported / cn_cert(p) - 1.0 for p in cn_anchors]
    cn_rms = _rms(cn_devs) if len(cn_devs) >= 2 else None
    cn_cref = (_rms([math.sqrt(cert_rel["n"](p) ** 2 + cert_rel["c"](p) ** 2) for p in cn_anchors])
               if cn_anchors else None)
    out["cn"] = {"n_monitors": len(cn_mons), "monitor_rows": [p.row for p in cn_mons], "s_rep": cn_rep,
                 "cert_scale": math.sqrt(out["n"]["cert_scale"] ** 2 + out["c"]["cert_scale"] ** 2),
                 "n_anchors": len(cn_anchors), "anchor_rms": cn_rms, "anchor_cert": cn_cref,
                 "trueness": None if cn_rms is None else math.sqrt(cn_rms ** 2 + (cn_cref or 0.0) ** 2)}
    F["unc_elem"] = out

    def cal2(p, el):
        """Squared relative calibration term of element el at p, or None."""
        if out[el]["in_force"]:
            if out[el]["s_rep"] is None:
                return None
            return _interp_multiplier(p.row, out[el]["monitor_rows"]) * out[el]["s_rep"] ** 2
        r = _kfit_rel(F[f"k_{el}"], getattr(p, f"ug_{el}"))
        return None if r is None else r * r

    for p in positions:
        for el in ("n", "c"):
            val, o = getattr(p, f"pct_{el}_reported"), out[el]
            u = None
            if val is not None and o["s_inst"] is not None and p.amount:
                c2 = cal2(p, el)
                if c2 is not None:
                    rel2 = (o["s_inst"] ** 2 + (u_m / p.amount) ** 2 + c2 + o["cert_scale"] ** 2
                            + (o["trueness"] or 0.0) ** 2)
                    u = val * math.sqrt(rel2)
            setattr(p, f"u_pct_{el}", u)
            setattr(p, f"U_pct_{el}", None if u is None else k_cov * u)
        rel = None
        if p.cn_molar_reported is not None and cn_rep is not None:
            if out["n"]["in_force"] and out["c"]["in_force"]:
                c2 = _interp_multiplier(p.row, out["cn"]["monitor_rows"]) * cn_rep ** 2
            else:
                parts = [cal2(p, el) for el in ("n", "c")]
                c2 = None if None in parts else sum(parts)
            if c2 is not None:
                rel = math.sqrt(cn_rep ** 2 + c2 + out["cn"]["cert_scale"] ** 2 + (out["cn"]["trueness"] or 0.0) ** 2)
        p.u_cn_molar = None if rel is None else p.cn_molar_reported * rel
        p.U_cn_molar = None if rel is None else k_cov * p.u_cn_molar
        p.u_cn_mass = None if rel is None else p.cn_mass_reported * rel
        p.U_cn_mass = None if rel is None else k_cov * p.u_cn_mass


# ---- rerun marks (1.4.4) ----------------------------------------------------
# 1 / the 0.75 quantile of the standard normal distribution: the median absolute
# deviation times this estimates the SD of normally distributed values.
ROBUST_SD = 1.482602218505602
_SIZE = {"n": ("N below linearity range", "N above linearity range"),
         "c": ("C below linearity range", "C above linearity range")}
_PREFIX = {"n": ("N peak only ", "N2 reference pulse off by "),
           "c": ("C peak only ", "CO2 reference pulse off by ")}
_PLAIN = ("no N peak", "no C peak", "CO2 peak rejected as a reference pulse",
          "ran before the background settled", "possible carryover from an oversized sample")


def _xl(x, places: int, sign: bool = False) -> str:
    """x as the workbook's TEXT(x, "0.00") writes it (1.4.10): to 15 significant digits first, as a
    spreadsheet keeps a number, and then to ``places`` decimals with a half rounded away from zero.
    Python's own format rounds the binary value, so -1.3749999999999998 came out -1.37 where the
    workbook reads -1.375 and writes -1.38. ``sign`` puts a + before a value that is not negative."""
    d = Decimal(f"{x if x != 0 else 0.0:.15g}").quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    return format(d, f"{'+' if sign else ''}.{places}f")


def _k_sigma(k):
    """k x 1 sigma as the checks name it: '2σ' for the default k = 2 (1.4.5)."""
    return f"{k:.15g}σ"


def _median(values):
    vals = sorted(values)
    m = len(vals) // 2
    return vals[m] if len(vals) % 2 else (vals[m - 1] + vals[m]) / 2


def _reason(flag, config=None):
    """Is this flag one that the rerun marks read? From 1.4.7 a peak above the linearity range is
    not, when the configuration gives the instrument's largest amplitude (max_amplitude_mv)."""
    if config is not None and config.max_amplitude_mv is not None and flag in (_SIZE["n"][1], _SIZE["c"][1]):
        return False
    return (flag in _PLAIN or flag in _SIZE["n"] or flag in _SIZE["c"]
            or flag.startswith(_PREFIX["n"]) or flag.startswith(_PREFIX["c"]))


def _extra_reasons(p, config, replicates=None):
    """The rerun reasons that are not flags, as (element, text): a peak at or above the instrument's
    largest amplitude and a final delta value whose 1 sigma is above the limit (1.4.7), and replicates
    whose range is above the critical range (1.4.8). In the order the workbook writes them: amplitude
    N, amplitude C, 1 sigma N, 1 sigma C, replicates N, replicates C."""
    out = []
    lim = config.max_amplitude_mv
    if lim is not None:
        for el, area, ampl, gas in (("n", p.area_n, p.ampl_n, "N2"), ("c", p.area_c, p.ampl_c, "CO2")):
            if area is not None and ampl is not None and not _above(lim, ampl):
                out.append((el, f"{gas} peak at or above {_xl(lim, 0)} mV"))
    usig = config.rerun_max_1sigma
    if usig is not None:
        for el, u, name in (("n", p.u_d15n, "d15N"), ("c", p.u_d13c, "d13C")):
            if u is not None and _above(u, usig):
                out.append((el, f"1σ({name}) {_xl(u, 2)} above {_xl(usig, 2)}"))
    g = (replicates or {}).get(getattr(p, "replicate_group", None))
    if g is not None:
        for el, name in (("n", "d15N"), ("c", "d13C")):
            if g[el]["over"]:
                out.append((el, f"replicates at positions {', '.join(str(r) for r in g['rows'])} differ by "
                                f"{_xl(g[el]['range'], 2)} in {name}, above {_xl(g[el]['limit'], 2)}"))
    return out


def _replicates(positions, F, config, groups):
    """Replicates of a sample in one run (1.4.8), from the sample sheet: their range against the
    critical range of ISO 5725-6, section 5.2, f(n) x s, with s the repeatability of one analysis in
    use (the configuration's, or the pooled anchor replicates') on the calibrated scale, so times the
    calibration slope. n counts the analyses with a value of that element."""
    on = str(config.rerun_replicates or "").strip().casefold() == "on"
    s = {}
    for el in ("n", "c"):
        u, m = F[f"unc_{el}"].get("u_sample"), F[f"cal_{el}"].get("slope")
        s[el] = None if u is None or m is None else u * abs(m)
    by_row = {p.row: p for p in positions}
    out = []
    for g in groups:
        e = dict(g)
        for el, attr in (("n", "d15n_final"), ("c", "d13c_final")):
            vals = [getattr(by_row[r], attr) for r in g["rows"] if getattr(by_row[r], attr) is not None]
            n = len(vals)
            rng = (max(vals) - min(vals)) if n >= 2 else None
            f = critical_range_factor(n)
            lim = f * s[el] if (f is not None and s[el] is not None) else None
            e[el] = {"n": n, "range": rng, "factor": f, "limit": lim,
                     "over": bool(on and rng is not None and lim is not None and _above(rng, lim))}
        out.append(e)
    F["replicates"] = {"on": on, "s": s, "groups": out}


def _mass_window(p, el, lo_area, hi_area, blank_area=0.0):
    """Masses (mg) that put this position's peak of el between lo_area and hi_area.

    Read from this analysis's own peak area per mg above the blank (1.4.10): a capsule
    of m mg gives the blank's area plus m times (area - blank) / mass, so the masses are
    (limit - blank) / that yield. None without a mass, a peak above the blank or both
    limits above the blank, or when no mass can reach the range (lo_area above hi_area).
    """
    area = getattr(p, f"area_{el}")
    ab = blank_area if blank_area and blank_area > 0 else 0.0
    if area is None or not p.amount or p.amount <= 0 or area - ab <= 0:
        return None
    if lo_area - ab <= 0 or hi_area - ab <= 0 or lo_area > hi_area:
        return None
    per_mg = (area - ab) / p.amount
    return (lo_area - ab) / per_mg, (hi_area - ab) / per_mg


def _rerun_marks(positions, F, config):
    """Mark each sample for a rerun of N, C, or both, or for a check, from its flags.

    Only what a rerun can fix marks a sample: the peak's size against the linearity
    standards and the blank, carryover from an oversized sample, and the reference
    pulse at that position. A sample run before the background settled, or missing
    the peak of one element, is marked for a check. A rerun's suggested mass puts
    each peak inside the linearity range (and above the blank multiple), read from
    this analysis's area per mg. Problems of the whole run are listed once, not on
    every sample.
    """
    L = F["size_limits"]
    mult = config.min_blank_multiple
    lo_area = {el: max(L[f"min_{el}"] or 0,
                       (mult * F[f"blank_{el}"]["area"]) if (mult and F[f"blank_{el}"]["area"] > 0) else 0)
               for el in ("n", "c")}
    hi_area = {el: L[f"max_{el}"] or 0 for el in ("n", "c")}
    blank_area = {el: F[f"blank_{el}"]["area"] for el in ("n", "c")}
    # 1.4.10: a suggested mass above every capsule weighed in this run says so
    masses = [p.amount for p in positions if p.amount and p.amount > 0]
    heaviest = max(masses) if masses else None
    counts = {"N": 0, "C": 0, "N and C": 0, "Check": 0}
    reps = {g["group"]: g for g in F.get("replicates", {}).get("groups", [])}
    for p in positions:
        p.rerun = p.rerun_why = p.rerun_mass = None
        p.rerun_mass_lo = p.rerun_mass_hi = None
        if p.role != "Sample":
            continue
        fl = p.flags
        no_n, no_c = "no N peak" in fl, "no C peak" in fl
        extra = _extra_reasons(p, config, reps)
        need = {el: any(_reason(f, config) and (f in _SIZE[el] or f.startswith(_PREFIX[el])) for f in fl)
                or any(e == el for e, _ in extra) for el in ("n", "c")}
        both = "possible carryover from an oversized sample" in fl or (no_n and no_c)
        check = "ran before the background settled" in fl or no_n != no_c
        if both or (need["n"] and need["c"]):
            mark = "N and C"
        elif need["n"]:
            mark = "N"
        elif need["c"]:
            mark = "C"
        elif check:
            mark = "Check"
        else:
            continue
        p.rerun = mark
        p.rerun_why = "; ".join([f for f in fl if _reason(f, config)] + [t for _, t in extra])
        counts[mark] += 1
        if mark == "Check":
            continue
        win = {el: w for el in ("n", "c")
               if (w := _mass_window(p, el, lo_area[el], hi_area[el], blank_area[el])) is not None}
        if not win:
            continue
        lo, hi = max(w[0] for w in win.values()), min(w[1] for w in win.values())
        beyond = (" (more than the heaviest capsule in this run)"
                  if heaviest is not None and _above(lo, heaviest) else "")
        if lo <= hi:
            p.rerun_mass_lo, p.rerun_mass_hi = lo, hi
            p.rerun_mass = f"{_xl(lo, 3)} to {_xl(hi, 3)}" + beyond
        elif len(win) == 2:
            p.rerun_mass = (f"N {_xl(win['n'][0], 3)} to {_xl(win['n'][1], 3)}; "
                            f"C {_xl(win['c'][0], 3)} to {_xl(win['c'][1], 3)} (two capsules)" + beyond)
        else:
            p.rerun_mass = "none"

    # Problems of the whole run, for the top of the Rerun sheet.
    items = []
    # 1.4.10: an element the anchors cannot calibrate (see the calibration step)
    for el, name, drifted in (("n", "d15N", "d15n_drift"), ("c", "d13C", "d13c_drift")):
        f = F[f"cal_{el}"]
        if f.get("calibrated"):
            continue
        if f.get("values") == 1:
            items.append(f"no {name} calibrated: every anchor carries the same accepted {name}, and the scale "
                         "needs two materials with different values")
        elif any(p.role == "Sample" and getattr(p, drifted) is not None for p in positions):
            items.append(f"no {name} calibrated: no anchor has a measured {name}")
    off = [p.row for p in positions if p.role == "Check"
           and any(f.startswith(("check d15N off by", "check d13C off by")) for f in p.flags)]
    if off:
        items.append(f"check standard off by more than {_xl(config.check_offset_permil, 2)} permil at "
                     f"{'position' if len(off) == 1 else 'positions'} {', '.join(str(r) for r in off)}")
    for name, val, cert, big_u in (("%N", "pct_n_reported", lambda p: p.pct_n, "U_pct_n"),
                                   ("%C", "pct_c_reported", lambda p: p.pct_c, "U_pct_c"),
                                   ("C:N", "cn_molar_reported",
                                    lambda p: ((p.pct_c / ATOMIC_C) / (p.pct_n / ATOMIC_N))
                                    if (p.pct_c and p.pct_n and p.pct_c > 0 and p.pct_n > 0) else None,
                                    "U_cn_molar")):
        used = [p for p in positions if p.role == "Anchor" and getattr(p, big_u) is not None
                and cert(p) is not None]
        within = sum(1 for p in used if abs(getattr(p, val) - cert(p)) <= getattr(p, big_u))
        if used and within < len(used):
            items.append(f"anchors outside {_k_sigma(config.coverage_k)}, {name}: {len(used) - within} of {len(used)}")
    if config.ref_gas_sd_limit is not None:
        for el, name in (("n", "N2"), ("c", "CO2")):
            g = F[f"ref_gas_{el}"]
            if g["n"] >= 3 and g["sd"] is not None and g["sd"] >= config.ref_gas_sd_limit:
                items.append(f"{name} reference pulses SD {_xl(g['sd'], 2)} permil, at or above "
                             f"{_xl(config.ref_gas_sd_limit, 2)}")
    F["rerun"] = {"counts": counts, "run_level": items,
                  "run_level_text": "; ".join(items) if items else "none found",
                  "lo_area": lo_area, "hi_area": hi_area}


def process(positions, config=None, samples=None):
    """Apply every correction. Returns (positions, fits).

    ``samples`` is a sample sheet (ea_irms.metadata.SampleSheet, 1.4.8): each Sample position is
    matched to it, and replicates of a sample in the run are compared for the rerun marks."""
    config = config or Config()
    alpha = config.significance_level
    positions = sorted(positions, key=lambda p: p.row)
    assign_roles(positions, config)
    F: dict = {}
    rep_groups = match_positions(positions, samples)
    F["metadata"] = {"sheet": samples.describe() if samples is not None else None,
                     "match": match_summary(positions, samples)}

    role = lambda r: [p for p in positions if p.role == r]      # noqa: E731

    # ---- CO2 peaks that are reference pulses (1.4.2) --------------------------
    # An export cut down to the sample peaks carries no reference peaks to count
    # from, so where a position made no sample CO2 the first exported CO2 peak is
    # a reference pulse that Isodat did not flag. It matches the designated
    # pulses the export does carry (at positions with no sample gas) in
    # amplitude, and reads the working-gas value in d13C, which no sample does.
    # The pulse amplitude moves through a run (2025-12-18: 1241 to 1177 mV), so
    # the yardstick is interpolated in position between the designated pulses.
    filtered = all(p.ref_bgd_n is None for p in positions)
    desig = [(a, d) for p in positions for a, d in p.desig_c if a is not None and d is not None]
    pairs = sorted((p.row, p.desig_c[0][0]) for p in positions if p.desig_c and p.desig_c[0][0] is not None)
    check = {"filtered": filtered, "n_designated": len(desig),
             "mean_ampl": sum(a for a, _ in desig) / len(desig) if desig else None,
             "mean_d13c": sum(d for _, d in desig) / len(desig) if desig else None,
             "ampl_tol": config.ref_pulse_ampl_tol, "d13c_window": config.ref_pulse_d13c_window,
             "active": bool(filtered and desig and pairs and config.ref_pulse_ampl_tol is not None),
             "rejected": []}

    def pulse_ampl_at(row):
        if row <= pairs[0][0]:
            return pairs[0][1]
        if row >= pairs[-1][0]:
            return pairs[-1][1]
        for (x0, a0), (x1, a1) in zip(pairs, pairs[1:]):
            if x0 <= row <= x1:
                return a0 + (row - x0) / (x1 - x0) * (a1 - a0)
    for p in positions:
        p.c_rejected = False
        p.rejected_ampl_c = p.rejected_d13c = None
        p.pulse_ampl_c = pulse_ampl_at(p.row) if pairs else None
    if check["active"]:
        for p in positions:
            if (p.ampl_c is not None and p.d13c is not None
                    and not _above(abs(p.ampl_c - p.pulse_ampl_c), config.ref_pulse_ampl_tol * p.pulse_ampl_c)
                    and not _above(abs(p.d13c - check["mean_d13c"]), config.ref_pulse_d13c_window)):
                p.c_rejected = True
                p.rejected_ampl_c, p.rejected_d13c = p.ampl_c, p.d13c
                p.area_c = p.ampl_c = p.d13c = None
                p.notes.append("CO2 peak rejected: it matches the designated reference pulses")
                check["rejected"].append(p.row)
    F["ref_pulse_check"] = check

    # ---- blank -----------------------------------------------------------
    bn = [p for p in role("Blank") if p.area_n is not None]
    bc = [p for p in role("Blank") if p.area_c is not None]
    F["blank_n"] = {"n": len(bn),
                    "area": sum(p.area_n for p in bn) / len(bn) if bn else 0.0,
                    "d": sum(p.d15n for p in bn) / len(bn) if bn else 0.0,
                    "sd": _sd([p.d15n for p in bn])}
    F["blank_c"] = {"n": len(bc),
                    "area": sum(p.area_c for p in bc) / len(bc) if bc else 0.0,
                    "d": sum(p.d13c for p in bc) / len(bc) if bc else 0.0,
                    "sd": _sd([p.d13c for p in bc])}

    def in_force(block):
        if config.blank_correction == "Off" or block["n"] < 1:
            return False
        if config.blank_correction == "On":
            return True
        return block["n"] >= 3 and block["sd"] is not None and block["sd"] <= 2.0

    F["blank_n"]["in_force"] = in_force(F["blank_n"])
    F["blank_c"]["in_force"] = in_force(F["blank_c"])
    # Blank capsules a lab set to Exclude (1.4.2), so QC can say so instead of
    # "no blank capsules found"; a "Blank Hole" is an empty position, not a capsule.
    F["blank_n"]["excluded"] = F["blank_c"]["excluded"] = sum(
        1 for p in positions if p.role == "Exclude" and "blank" in p.identifier.lower()
        and "hole" not in p.identifier.lower())

    for p in positions:
        p.d15n_blank = p.d15n
        if (p.d15n is not None and F["blank_n"]["in_force"]
                and p.role != "Blank" and p.area_n > F["blank_n"]["area"]):
            ab, db = F["blank_n"]["area"], F["blank_n"]["d"]
            p.d15n_blank = (p.area_n * p.d15n - ab * db) / (p.area_n - ab)
        p.d13c_blank = p.d13c
        if (p.d13c is not None and F["blank_c"]["in_force"]
                and p.role != "Blank" and p.area_c > F["blank_c"]["area"]):
            ab, db = F["blank_c"]["area"], F["blank_c"]["d"]
            p.d13c_blank = (p.area_c * p.d13c - ab * db) / (p.area_c - ab)

    # ---- linearity -------------------------------------------------------
    lin = role("Linearity")
    for el, xattr, yattr, out in (("n", "area_n", "d15n_blank", "lin_n"),
                                  ("c", "area_c", "d13c_blank", "lin_c")):
        xs = [getattr(p, xattr) for p in lin]
        ys = [getattr(p, yattr) for p in lin]
        fit = _fit(xs, ys)
        good = [v for v in xs if v is not None]
        F[out] = {"fit": fit, "pivot": sum(good) / len(good) if good else 0.0,
                  "n": len(good), "min": min(good) if good else None,
                  "max": max(good) if good else None}
        mode = config.linearity_correction
        if el == "c" and config.linearity_correction_c:
            mode = config.linearity_correction_c      # carbon switch set separately (1.4.1)
        F[out]["mode"] = mode
        F[out]["in_force"] = _apply(mode, fit, len(good), alpha)
        F[out]["applied"] = fit["slope"] if F[out]["in_force"] else 0.0
    for p in positions:
        p.d15n_lin = (None if p.d15n_blank is None else
                      p.d15n_blank - F["lin_n"]["applied"] * (p.area_n - F["lin_n"]["pivot"]))
        p.d13c_lin = (None if p.d13c_blank is None else
                      p.d13c_blank - F["lin_c"]["applied"] * (p.area_c - F["lin_c"]["pivot"]))

    # ---- drift -----------------------------------------------------------
    first = config.first_drift_position or 0
    monitors = [p for p in role("Drift") if p.row >= first]
    excluded = {"n": set(int(v) for v in config.exclude_drift_n or []),
                "c": set(int(v) for v in config.exclude_drift_c or [])}
    for el, yattr, out in (("n", "d15n_lin", "drift_n"), ("c", "d13c_lin", "drift_c")):
        used = [p for p in monitors if p.row not in excluded[el]]
        ys = [getattr(p, yattr) for p in used]
        by_row = _fit([float(p.row) for p in used], ys)
        by_bgd = _fit([p.ref_bgd_n for p in used], ys)
        if config.drift_model in ("Row", "Background"):
            model = config.drift_model
        elif (by_row and by_bgd and by_row["sd"] is not None
              and by_bgd["sd"] is not None and by_bgd["sd"] < by_row["sd"]):
            model = "Background"
        else:
            model = "Row"
        chosen = by_row if model == "Row" else by_bgd
        xs = ([float(p.row) for p in used] if model == "Row"
              else [p.ref_bgd_n for p in used])
        pairs = [v for v, y in zip(xs, ys) if v is not None and y is not None]
        F[out] = {"row": by_row, "background": by_bgd, "model": model,
                  "pivot": sum(pairs) / len(pairs) if pairs else 0.0,
                  "n": len(pairs), "excluded": sorted(excluded[el]),
                  "fit": chosen}
        F[out]["in_force"] = _apply(config.drift_correction, chosen, len(pairs), alpha)
        F[out]["applied"] = chosen["slope"] if F[out]["in_force"] else 0.0
    for p in positions:
        for out, src, dst in (("drift_n", "d15n_lin", "d15n_drift"),
                              ("drift_c", "d13c_lin", "d13c_drift")):
            v = getattr(p, src)
            if v is None:
                setattr(p, dst, None)
                continue
            x = float(p.row) if F[out]["model"] == "Row" or p.ref_bgd_n is None else p.ref_bgd_n
            setattr(p, dst, v - F[out]["applied"] * (x - F[out]["pivot"]))

    # Drift monitors that disagree with the others, in units of the others' scatter.
    for p in positions:
        p.t_drift_n = p.t_drift_c = None
    for el, yattr, out, dst in (("n", "d15n_lin", "drift_n", "t_drift_n"),
                                ("c", "d13c_lin", "drift_c", "t_drift_c")):
        fit = F[out]["fit"]
        for p in monitors:
            x = float(p.row) if F[out]["model"] == "Row" else p.ref_bgd_n
            setattr(p, dst, _studentized(fit, x, getattr(p, yattr),
                                         p.row not in excluded[el]))
        F[out]["outliers"] = sorted(p.row for p in monitors
                                    if getattr(p, dst) is not None
                                    and abs(getattr(p, dst)) > config.monitor_outlier_t
                                    and p.row not in excluded[el])

    # ---- calibration -----------------------------------------------------
    # 1.4.10: a scale needs anchors of two materials with different accepted values. With
    # one, the line through them is flat (slope 0) and would give every sample that
    # material's value, so nothing is calibrated and the run says why.
    anchors = role("Anchor")
    for out, xattr, accattr, dst in (("cal_n", "d15n_drift", "acc15", "d15n"),
                                     ("cal_c", "d13c_drift", "acc13", "d13c")):
        xs = [getattr(p, xattr) for p in anchors]
        ys = [getattr(p, accattr) for p in anchors]
        acc = [a for x, a in zip(xs, ys) if x is not None and a is not None]
        two = bool(acc) and max(acc) > min(acc)
        fit = (_fit(xs, ys) if two else None) or {"slope": None, "intercept": None, "r2": None,
                                                  "sd": None, "n": 0}
        fit["acc_min"] = min(acc) if acc else None
        fit["acc_max"] = max(acc) if acc else None
        fit["analyses"] = len(acc)
        fit["values"] = len(set(acc))
        fit["calibrated"] = two and fit["n"] >= 2
        F[out] = fit
    for p in positions:
        f = F["cal_n"]
        p.d15n_final = (None if p.d15n_drift is None or f["n"] < 2
                        else f["slope"] * p.d15n_drift + f["intercept"])
        f = F["cal_c"]
        p.d13c_final = (None if p.d13c_drift is None or f["n"] < 2
                        else f["slope"] * p.d13c_drift + f["intercept"])

    # ---- uncertainty of the calibrated values -----------------------------
    # Inputs: each anchor material's accepted value and its measured mean, and
    # the sample's own value. Each input's standard uncertainty is propagated
    # through the regression by its exact first-order effect on the result,
    # which is what a Kragten spreadsheet computes by perturbation.
    for el, out, xattr, accattr, uattr, dst in (
            ("n", "cal_n", "d15n_drift", "acc15", "u_acc15", "u_d15n"),
            ("c", "cal_c", "d13c_drift", "acc13", "u_acc13", "u_d13c")):
        f = F[out]
        groups: dict = {}
        for p in anchors:
            x, a = getattr(p, xattr), getattr(p, accattr)
            if x is None or a is None:
                continue
            groups.setdefault(p.std_index, {"x": [], "u_acc": getattr(p, uattr)})
            groups[p.std_index]["x"].append(x)
        budget = {"materials": [], "pooled_sd": None, "u_sample": None,
                  "u_sample_source": None}
        if f["n"] >= 3 and f.get("sxx"):
            m, mx, sxx, n = f["slope"], f["mean_x"], f["sxx"], f["n"]
            ss, dfree = 0.0, 0
            for idx, g in groups.items():
                k = len(g["x"])
                sd = _sd(g["x"])
                g["n"] = k
                g["mean"] = sum(g["x"]) / k
                g["sd"] = sd
                g["u_mean"] = sd / math.sqrt(k) if sd is not None else 0.0
                g["s1"] = sum(x - mx for x in g["x"])
                # residual of each replicate against the line, minus m*(x - mean_x)
                acc = [getattr(p, accattr) for p in anchors if p.std_index == idx
                       and getattr(p, xattr) is not None][0]
                g["s2"] = sum((acc - (m * x + f["intercept"])) - m * (x - mx) for x in g["x"])
                if sd is not None:
                    ss += (k - 1) * sd * sd
                    dfree += k - 1
                budget["materials"].append({"index": idx, "n": k, "mean": g["mean"],
                                            "sd": sd, "u_mean": g["u_mean"],
                                            "u_acc": g["u_acc"]})
            pooled = math.sqrt(ss / dfree) if dfree > 0 else None
            budget["pooled_sd"] = pooled
            chosen_u = getattr(config, f"u_sample_{el}")
            if chosen_u is not None:
                budget["u_sample"], budget["u_sample_source"] = chosen_u, "configuration"
            else:
                budget["u_sample"], budget["u_sample_source"] = pooled, "pooled anchor replicates"
            u_s = budget["u_sample"]
            for p in positions:
                x = getattr(p, xattr)
                if x is None or u_s is None:
                    setattr(p, dst, None)
                    continue
                total = (u_s * m) ** 2
                for g in groups.values():
                    w_acc = g["n"] / n + g["s1"] * (x - mx) / sxx
                    w_meas = (x - mx) * g["s2"] / sxx - g["n"] * m / n
                    total += ((g["u_acc"] or 0.0) * w_acc) ** 2 + (g["u_mean"] * w_meas) ** 2
                setattr(p, dst, math.sqrt(total))
        else:
            for p in positions:
                setattr(p, dst, None)
        F[f"unc_{el}"] = budget
    F["coverage_k"] = config.coverage_k
    for p in positions:
        p.U_d15n = None if p.u_d15n is None else config.coverage_k * p.u_d15n
        p.U_d13c = None if p.u_d13c is None else config.coverage_k * p.u_d13c

    # ---- elemental content ----------------------------------------------
    for out, pct, area, mass_key in (("k_n", "pct_n", "area_n", "ug_n"),
                                     ("k_c", "pct_c", "area_c", "ug_c")):
        xs, ys = [], []
        for p in positions:
            if p.role in ("Blank", "Exclude", "Sample") or not p.amount:
                continue
            if getattr(p, pct) is None or getattr(p, area) is None:
                continue
            xs.append(p.amount * 1000 * getattr(p, pct) / 100)
            ys.append(getattr(p, area))
        F[out] = _fit(xs, ys) or {"slope": None, "intercept": None, "r2": None,
                                  "sd": None, "n": 0}
    # 1.4.10: a peak at or below the k-factor line's intercept, the area it gives for none of the
    # element, holds no detectable N or C; %N, %C and C:N are then not reported, and the flag says so.
    for p in positions:
        for out, area, dst in (("k_n", "area_n", "ug_n"), ("k_c", "area_c", "ug_c")):
            f, a = F[out], getattr(p, area)
            setattr(p, dst, None if a is None or f["n"] < 3
                    else (a - f["intercept"]) / f["slope"])
        mass = p.amount is not None and p.amount > 0
        p.pct_n_measured = (None if p.ug_n is None or not mass or p.ug_n <= 0
                            else p.ug_n / (p.amount * 1000) * 100)
        p.pct_c_measured = (None if p.ug_c is None or not mass or p.ug_c <= 0
                            else p.ug_c / (p.amount * 1000) * 100)
        p.below_detection_n = bool(mass and p.ug_n is not None and p.ug_n <= 0)
        p.below_detection_c = bool(mass and p.ug_c is not None and p.ug_c <= 0)
        if p.ug_n and p.ug_c and p.ug_n > 0 and p.ug_c > 0:
            p.cn_molar = (p.ug_c / ATOMIC_C) / (p.ug_n / ATOMIC_N)
            p.cn_mass = p.ug_c / p.ug_n
        else:
            p.cn_molar = p.cn_mass = None

    # ---- elemental drift, from the drift monitors (1.4.2) --------------------
    # The k-factor is one line for the whole run, so a sensitivity that moves
    # through the run leaves the drift monitors reading high or low in %N and
    # %C by position (2026-09-23: 9.0 to 12.9 %N against 10.36). Each monitor's
    # relative bias, measured over certified, gives a factor at its position;
    # between monitors the factor is interpolated in position, and beyond the
    # first and last it is held. Every position's micrograms are divided by the
    # factor at its position. Monitors listed in the configuration are left
    # out, for a capsule whose weight is wrong rather than the instrument.
    excl_e = set(int(v) for v in config.exclude_elemental or [])
    for el, pct_meas, pct_cert, out in (("n", "pct_n_measured", "pct_n", "elem_n"),
                                        ("c", "pct_c_measured", "pct_c", "elem_c")):
        for p in positions:
            setattr(p, f"elem_bias_{el}", None)
        mons = []
        for p in role("Drift"):
            meas, cert = getattr(p, pct_meas), getattr(p, pct_cert)
            if meas is None or not cert:
                continue
            bias = meas / cert - 1.0
            setattr(p, f"elem_bias_{el}", bias)
            if p.row not in excl_e:
                mons.append((p.row, 1.0 + bias))
        mons.sort()
        n_used = len(mons)
        max_bias = max(abs(f - 1.0) for _, f in mons) if mons else None
        mode = config.elemental_correction
        if mode == "Off" or n_used < 2:
            in_force = False
        elif mode == "On":
            in_force = True
        else:
            in_force = _above(max_bias * 100.0, config.elemental_threshold_pct)
        F[out] = {"mode": mode, "n": n_used, "max_bias": max_bias, "in_force": in_force,
                  "excluded": sorted(excl_e), "monitors": mons,
                  "threshold_pct": config.elemental_threshold_pct}

        def factor_at(row, mons=mons, in_force=in_force):
            if not in_force:
                return 1.0
            if row <= mons[0][0]:
                return mons[0][1]
            if row >= mons[-1][0]:
                return mons[-1][1]
            for (x0, f0), (x1, f1) in zip(mons, mons[1:]):
                if x0 <= row <= x1:
                    return f0 + (row - x0) / (x1 - x0) * (f1 - f0)
            return 1.0
        for p in positions:
            setattr(p, f"elem_factor_{el}", factor_at(p.row))
    for p in positions:
        p.pct_n_reported = None if p.pct_n_measured is None else p.pct_n_measured / p.elem_factor_n
        p.pct_c_reported = None if p.pct_c_measured is None else p.pct_c_measured / p.elem_factor_c
        if p.cn_molar is None:
            p.cn_molar_reported = p.cn_mass_reported = None
        else:
            p.cn_molar_reported = p.cn_molar * p.elem_factor_n / p.elem_factor_c
            p.cn_mass_reported = p.cn_mass * p.elem_factor_n / p.elem_factor_c

    # ---- uncertainty of %N, %C, and C:N (1.4.3) -------------------------------
    _elemental_uncertainty(positions, F, config, excl_e)

    # ---- size limits -----------------------------------------------------
    min_n = config.min_area_n if config.min_area_n is not None else (F["lin_n"]["min"] or 0)
    min_c = config.min_area_c if config.min_area_c is not None else (F["lin_c"]["min"] or 0)
    max_n = config.max_area_n if config.max_area_n is not None else (F["lin_n"]["max"] or 0)
    max_c = config.max_area_c if config.max_area_c is not None else (F["lin_c"]["max"] or 0)
    F["size_limits"] = {"min_n": min_n, "min_c": min_c, "max_n": max_n, "max_c": max_c}

    # ---- carryover after oversized samples ---------------------------------
    # A large sample leaves something in the EA that raises the background of
    # the next one or two analyses (2026-09-22: 9 to 14 mV after the three
    # largest N peaks, while the largest itself read the run median). A position
    # is exposed when one of the CARRY_REACH positions before it is above the
    # linearity range; its rise is measured against the last reading that was
    # not exposed, so a slow change across the run does not count.
    over = {p.row: bool((max_n and p.area_n is not None and p.area_n > max_n)
                        or (max_c and p.area_c is not None and p.area_c > max_c))
            for p in positions}
    ref = None
    for p in positions:
        p.oversized = over[p.row]
        p.carry_exposed = any(over.get(p.row - k, False) for k in range(1, CARRY_REACH + 1))
        if not p.carry_exposed and p.bgd_n is not None:
            ref = p.bgd_n
        p.carry_ref = ref
        p.carry_rise = (p.bgd_n - ref if p.carry_exposed and p.bgd_n is not None
                        and ref is not None else None)
        p.carry_flag = (config.carryover_mv is not None and p.carry_rise is not None
                        and _above(p.carry_rise, config.carryover_mv))
    rises = [p.carry_rise for p in positions if p.carry_rise is not None]
    F["carryover"] = {
        "threshold": config.carryover_mv,
        "oversized": sum(1 for p in positions if p.oversized),
        "followed": sum(1 for p in positions if over.get(p.row - 1, False)
                        and p.role in ("Sample", "Check")),
        "flagged": [p.row for p in positions if p.carry_flag],
        "max_rise": max(rises) if rises else None,
    }

    # ---- settling --------------------------------------------------------
    # Settling happens at the start of a run, so it is looked for there. The
    # scale comes from the first readings, not the run's highest, so a spike
    # later in the run neither stretches it nor moves the settled position, and
    # a run that is flat from the start is not mistaken for one that never
    # settled, which the rule before 1.4.0 did because its scale was then noise.
    reads = [p for p in positions if p.ref_bgd_n is not None]
    bgs = sorted(p.ref_bgd_n for p in reads)
    floor = sum(bgs[:5]) / len(bgs[:5]) if bgs else 0.0
    excess = (max(p.ref_bgd_n for p in reads[:SETTLE_START]) - floor) if reads else 0.0
    # 1.4.2: the start must also stand above the run's own level, the median
    # reference-peak background, by the same margin.
    if bgs:
        k = len(bgs)
        level = bgs[k // 2] if k % 2 else (bgs[k // 2 - 1] + bgs[k // 2]) / 2.0
    else:
        level = 0.0
    excess_level = (max(p.ref_bgd_n for p in reads[:SETTLE_START]) - level) if reads else 0.0
    settled = 1
    if (config.settled_threshold and excess > 0
            and not excess < SETTLE_MIN_EXCESS - _SLACK
            and not excess_level < SETTLE_MIN_EXCESS - _SLACK):
        settled, streak, start = None, 0, None
        for p in reads:
            if p.carry_exposed:
                continue
            if not _above((p.ref_bgd_n - floor) / excess, config.settled_threshold):
                if streak == 0:
                    start = p.row
                streak += 1
                if streak >= SETTLE_RUN:
                    settled = start
                    break
            else:
                streak = 0
        if settled is None:                      # never settled
            settled = positions[-1].row + 1
    for p in positions:
        p.settling_fraction = (None if p.ref_bgd_n is None or excess <= 0
                               else (p.ref_bgd_n - floor) / excess)
    F["settling"] = {"floor": floor, "excess": excess, "settled_from": settled,
                     "level": level, "excess_level": excess_level,
                     "early": _fit([float(p.row) for p in role("Drift") if p.row < settled],
                                   [p.d15n_lin for p in role("Drift") if p.row < settled]),
                     "late": _fit([float(p.row) for p in role("Drift") if p.row >= settled],
                                  [p.d15n_lin for p in role("Drift") if p.row >= settled]),
                     "linearity_before": sum(1 for p in lin if p.row < settled),
                     "anchors_before": sum(1 for p in anchors if p.row < settled)}

    # The background under the drift monitors' own peaks: in every export, and
    # comparable across monitors because they are all one size.
    mon = [(float(p.row), p.bgd_n) for p in role("Drift") if p.bgd_n is not None]
    mfit = _fit([r for r, _ in mon], [b for _, b in mon])
    F["monitor_bgd"] = {"n": len(mon), "min": min(b for _, b in mon) if mon else None,
                        "max": max(b for _, b in mon) if mon else None, "fit": mfit,
                        "trend": (mfit is not None and mfit["p"] is not None and mfit["p"] < alpha)}

    # ---- working gas, from the reference pulses Isodat did not use --------
    # The pulse Isodat designates as the reference carries the assigned value
    # by definition, so only the other pulses can show whether the reference
    # measurement held steady across the run.
    # 1.4.10: every such pulse counts, not only the first at each position (the full exports
    # carry two CO2 pulses at each position, and the SD of the first alone understated it).
    for el, attr, out in (("n", "ref_pulses_d15n", "ref_gas_n"),
                          ("c", "ref_pulses_d13c", "ref_gas_c")):
        vals = [(float(p.row), v) for p in positions for v in (getattr(p, attr, None) or [])]
        ys = [v for _, v in vals]
        fit = _fit([r for r, _ in vals], ys)
        F[out] = {"n": len(vals), "positions": len({r for r, _ in vals}),
                  "mean": sum(ys) / len(ys) if ys else None,
                  "sd": _sd(ys), "range": (max(ys) - min(ys)) if ys else None,
                  "fit": fit,
                  "trend": (fit is not None and fit["p"] is not None and fit["p"] < alpha)}

    # ---- run-level checks --------------------------------------------------
    checks = [p for p in role("Check") if p.d15n_final is not None or p.d13c_final is not None]
    F["checks"] = {"n": len(checks), "materials": sorted({p.identifier for p in checks})}

    # ---- the reference pulse at each position (1.4.4) ----------------------
    # Against the run's own scatter: the median and a robust SD, so that the
    # positions being judged cannot widen the limit they are judged by.
    # 1.4.10: over every pulse; a position is judged by its pulse furthest from the median.
    F["pulse_off"] = {}
    for el, attr in (("n", "ref_pulses_d15n"), ("c", "ref_pulses_d13c")):
        ys = [v for p in positions for v in (getattr(p, attr, None) or [])]
        med = _median(ys) if len(ys) >= 5 else None
        rsd = None if med is None else ROBUST_SD * _median([abs(y - med) for y in ys])
        off_sd = config.ref_pulse_off_sd
        lim = (off_sd * rsd) if (off_sd is not None and off_sd > 0 and rsd) else None
        F["pulse_off"][el] = {"n": len(ys), "median": med, "robust_sd": rsd, "limit": lim}
        for p in positions:
            devs = [v - med for v in (getattr(p, attr, None) or [])] if med is not None else []
            # the largest distance, and of two equal ones the one above the median
            setattr(p, f"pulse_dev_{el}", max(devs, key=lambda d: (abs(d), d)) if devs else None)

    # ---- flags -----------------------------------------------------------
    for p in positions:
        flags = []
        if p.area_n is None:
            flags.append("no N peak")
        elif min_n and p.area_n < min_n:
            flags.append("N below linearity range")
        elif max_n and p.area_n > max_n:
            flags.append("N above linearity range")
        if p.area_c is None:
            flags.append("no C peak")
            if getattr(p, "c_rejected", False):
                flags.append("CO2 peak rejected as a reference pulse")
        elif min_c and p.area_c < min_c:
            flags.append("C below linearity range")
        elif max_c and p.area_c > max_c:
            flags.append("C above linearity range")
        for value, fit, name in ((p.d15n_final, F["cal_n"], "d15N"),
                                 (p.d13c_final, F["cal_c"], "d13C")):
            if (value is not None and p.role in ("Sample", "Check")
                    and fit.get("acc_min") is not None
                    and not fit["acc_min"] <= value <= fit["acc_max"]):
                flags.append(f"{name} outside anchor range")
        if p.role == "Check":
            for got, acc, name in ((p.d15n_final, p.acc15, "d15N"),
                                   (p.d13c_final, p.acc13, "d13C")):
                if got is not None and acc is not None and abs(got - acc) > config.check_offset_permil:
                    flags.append(f"check {name} off by {_xl(got - acc, 2, sign=True)}")
        if p.role == "Drift":
            for el, name in (("n", "N"), ("c", "C")):
                if p.row in excluded[el]:
                    flags.append(f"excluded from {name} drift fit")
                elif p.row in F[f"drift_{el}"]["outliers"]:
                    flags.append(f"{name} drift monitor outlier")
            if p.row in excl_e:
                flags.append("excluded from elemental correction")
        if config.min_blank_multiple and p.role != "Blank":
            for area, ab, name in ((p.area_n, F["blank_n"]["area"], "N"),
                                   (p.area_c, F["blank_c"]["area"], "C")):
                if area is not None and ab > 0 and area < ab * config.min_blank_multiple:
                    flags.append(f"{name} peak only {_xl(area / ab, 1)}x the blank")
        if p.row < F["settling"]["settled_from"]:
            flags.append("ran before the background settled")
        if p.carry_flag:
            flags.append("possible carryover from an oversized sample")
        for el, name in (("n", "N2"), ("c", "CO2")):
            dev, lim = getattr(p, f"pulse_dev_{el}"), F["pulse_off"][el]["limit"]
            if dev is not None and lim is not None and _above(abs(dev), lim):
                flags.append(f"{name} reference pulse off by {_xl(dev, 2, sign=True)}")
        # 1.4.10: no %N or %C where the peak is at or below the k-factor line's intercept
        if p.below_detection_n:
            flags.append("%N below detection")
        if p.below_detection_c:
            flags.append("%C below detection")
        p.flags = flags
    _replicates(positions, F, config, rep_groups)
    _rerun_marks(positions, F, config)
    return positions, F


def _clean(value):
    """NaN, empty strings and whitespace all mean 'not given'."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:      # NaN
        return None
    if isinstance(value, str):
        text = value.strip()
        return text or None
    return value


def standards_from_rows(rows) -> list[tuple]:
    """Normalise a table of standards, from JSON or from an edited grid.

    A float column holding no value comes back as NaN rather than None, and a
    NaN accepted value would propagate into every corrected number without
    raising anything, so it is cleaned here rather than at each use. Rows from
    configurations written before the uncertainty columns existed are padded.
    """
    out = []
    for row in rows or []:
        cells = [_clean(v) for v in list(row)[:STANDARD_FIELDS]]
        cells += [None] * (STANDARD_FIELDS - len(cells))
        if cells[0] is None:                             # no match text, no rule
            continue
        out.append(tuple(cells))
    return out


def _positions_list(value) -> list:
    """Accept a list, a number, or text such as '59, 61', '59; 61' or '59 61'. From 1.4.10 anything that is
    not a whole position is refused rather than dropped, so a list cannot quietly leave out nothing."""
    if value is None or value == "":
        return []
    if isinstance(value, bool):
        raise ValueError(f"{value!r} is not a list of positions")
    if isinstance(value, (int, float)):
        items = [value]
    elif isinstance(value, str):
        items = [t for t in re.split(r"[,;\s]+", value) if t]
    else:
        items = list(value)
    out = []
    for v in items:
        if isinstance(v, bool):
            raise ValueError(f"{value!r} is not a list of positions")
        try:
            f = float(v.strip()) if isinstance(v, str) else float(v)
        except (TypeError, ValueError):
            raise ValueError(f"{value!r} is not a list of whole positions") from None
        if not math.isfinite(f) or f != int(f) or f < 1:
            raise ValueError(f"{value!r} is not a list of whole positions from 1")
        out.append(int(f))
    return out


def config_to_dict(config: Config) -> dict:
    """Serialise a configuration so a lab can keep and reuse it."""
    return {
        "standards": [list(r) for r in config.standards],
        "blank_correction": config.blank_correction,
        "linearity_correction": config.linearity_correction,
        "linearity_correction_c": config.linearity_correction_c,
        "drift_correction": config.drift_correction,
        "drift_model": config.drift_model,
        "first_drift_position": config.first_drift_position,
        "exclude_drift_n": list(config.exclude_drift_n),
        "exclude_drift_c": list(config.exclude_drift_c),
        "significance_level": config.significance_level,
        "monitor_outlier_t": config.monitor_outlier_t,
        "check_offset_permil": config.check_offset_permil,
        "min_area_n": config.min_area_n,
        "min_area_c": config.min_area_c,
        "max_area_n": config.max_area_n,
        "max_area_c": config.max_area_c,
        "min_blank_multiple": config.min_blank_multiple,
        "settled_threshold": config.settled_threshold,
        "carryover_mv": config.carryover_mv,
        "u_sample_n": config.u_sample_n,
        "u_sample_c": config.u_sample_c,
        "u_sample_n_source": config.u_sample_n_source,
        "u_sample_c_source": config.u_sample_c_source,
        "coverage_k": config.coverage_k,
        "balance_u_mg": config.balance_u_mg,
        "elemental_correction": config.elemental_correction,
        "elemental_threshold_pct": config.elemental_threshold_pct,
        "exclude_elemental": list(config.exclude_elemental),
        "ref_pulse_ampl_tol": config.ref_pulse_ampl_tol,
        "ref_pulse_d13c_window": config.ref_pulse_d13c_window,
        "ref_pulse_off_sd": config.ref_pulse_off_sd,
        "ref_gas_sd_limit": config.ref_gas_sd_limit,
        "max_amplitude_mv": config.max_amplitude_mv,
        "rerun_max_1sigma": config.rerun_max_1sigma,
        "rerun_replicates": config.rerun_replicates,
        "run_file": config.run_file,
    }


def config_from_dict(data: dict) -> Config:
    """Rebuild a configuration, filling anything absent from the defaults. The switches are read in any
    case and positions as ea_irms.labconfig reads them (1.4.10); labconfig.read_config_json also checks
    every number and names the keys it does not read."""
    config = Config()
    for key, value in (data or {}).items():
        if key == "standards":
            config.standards = standards_from_rows(value)
        elif key in ("exclude_drift_n", "exclude_drift_c", "exclude_elemental"):
            setattr(config, key, _positions_list(value))
        elif hasattr(config, key):
            setattr(config, key, value)
    _normalize(config)
    return config
