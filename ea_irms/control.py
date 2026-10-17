"""Control charts for an in-house standard, run after run.

Follows section 7.1 of the FIRMS Good Practice Guide (Dunn and Carter, 2018):
a mean-value chart with warning limits at 2 SD and control limits at 3 SD, the
limits set from a prior period of about 20 results on at least 6 different days,
and the Westgard rules of Thompson and Wood (1995) to say when the instrument
has gone out of control. Nothing here reads a database; the app hands in the
history and draws what comes back.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date


@dataclass
class Limits:
    mean: float
    sd: float
    n: int
    days: int
    warning: tuple           # (lower, upper) at 2 SD
    control: tuple           # (lower, upper) at 3 SD
    source: str              # what the limits were set from


@dataclass
class Point:
    when: date | None
    value: float
    z: float | None = None
    rules: list = field(default_factory=list)   # Westgard rules this point breaks

    @property
    def out_of_control(self) -> bool:
        return bool(self.rules)


def set_limits(values, dates, min_results: int = 20, min_days: int = 6,
               source: str = "prior period"):
    """Limits from a prior period, or None with the reason they cannot be set yet.

    The guide asks for about 20 results on at least 6 different days, so a run
    of replicates on one afternoon does not pass for a history.
    """
    vals = [v for v in values if v is not None]
    days = len({d for d in dates if d is not None})
    if len(vals) < min_results:
        return None, f"{len(vals)} results so far; limits need about {min_results}"
    if days < min_days:
        return None, f"results from {days} days so far; limits need at least {min_days}"
    mean = sum(vals) / len(vals)
    sd = math.sqrt(sum((v - mean) ** 2 for v in vals) / (len(vals) - 1))
    if sd == 0:
        return None, "no scatter in the prior period"
    return Limits(mean=mean, sd=sd, n=len(vals), days=days,
                  warning=(mean - 2 * sd, mean + 2 * sd),
                  control=(mean - 3 * sd, mean + 3 * sd), source=source), None


def westgard(values, mean: float, sd: float) -> list[list[str]]:
    """Which Westgard rules each observation breaks, in sequence order.

    1_3s   one observation beyond 3 SD
    2_2s   two consecutive observations beyond 2 SD on the same side
    R_4s   two consecutive observations more than 4 SD apart
    4_1s   four consecutive observations beyond 1 SD on the same side
    10_x   ten consecutive observations on the same side of the mean
    These are the rules in figure 18 of the guide. "Beyond" a number of SD allows the slack of every other
    limit in ea_irms (pipeline._above and _below, finding R-24 of the 2026-10-14 review), so an observation on
    a limit is not beyond it; the side of the mean is exact.
    """
    from .pipeline import _above, _below
    z = [(v - mean) / sd for v in values]
    out: list[list[str]] = [[] for _ in z]
    for i, zi in enumerate(z):
        if _above(abs(zi), 3):
            out[i].append("1_3s")
        if i >= 1:
            if _above(abs(zi), 2) and _above(abs(z[i - 1]), 2) and (zi > 0) == (z[i - 1] > 0):
                out[i].append("2_2s")
            if _above(abs(zi - z[i - 1]), 4):
                out[i].append("R_4s")
        if i >= 3:
            window = z[i - 3:i + 1]
            if all(_above(w, 1) for w in window) or all(_below(w, -1) for w in window):
                out[i].append("4_1s")
        if i >= 9:
            window = z[i - 9:i + 1]
            if all(w > 0 for w in window) or all(w < 0 for w in window):
                out[i].append("10_x")
    return out


def evaluate(history, limits: Limits) -> list[Point]:
    """Score a history of (date, value) pairs against fixed limits."""
    pts = [Point(when=d, value=v) for d, v in history if v is not None]
    rules = westgard([p.value for p in pts], limits.mean, limits.sd)
    for p, r in zip(pts, rules):
        p.z = (p.value - limits.mean) / limits.sd
        p.rules = r
    return pts


def chart(history, limits: Limits | None, title: str, ylabel: str, target: float | None = None):
    """A matplotlib figure of the history with the limits drawn, or without them.

    Returns the figure; the caller shows or saves it. Imported lazily so the
    engine does not need matplotlib unless a chart is drawn.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    pts = [(d, v) for d, v in history if v is not None]
    fig, ax = plt.subplots(figsize=(8, 3.6), dpi=150)
    xs = list(range(1, len(pts) + 1))
    ys = [v for _, v in pts]
    if limits:
        for y, style, label in ((limits.mean, "-", "mean"),
                                (limits.warning[0], "--", "warning, 2 SD"), (limits.warning[1], "--", None),
                                (limits.control[0], ":", "control, 3 SD"), (limits.control[1], ":", None)):
            ax.axhline(y, color="#52514e", lw=1, ls=style, label=label)
        scored = evaluate(pts, limits)
        bad = [i + 1 for i, p in enumerate(scored) if p.out_of_control]
        ax.scatter(xs, ys, s=28, color="#2a78d6", zorder=3, label="results")
        if bad:
            ax.scatter(bad, [ys[i - 1] for i in bad], s=60, facecolor="none", edgecolor="#d03b3b",
                       lw=1.6, zorder=4, label="Westgard rule broken")
    else:
        ax.scatter(xs, ys, s=28, color="#2a78d6", zorder=3, label="results")
    if target is not None:
        ax.axhline(target, color="#eb6834", lw=1, label="assigned value")
    ax.plot(xs, ys, color="#2a78d6", lw=0.8, alpha=0.5)
    ax.set_xlabel("Result, in order of run date")
    ax.set_ylabel(ylabel)
    ax.set_title(title, loc="left", fontsize=10, fontweight="bold")
    ax.legend(frameon=False, fontsize=8, ncol=3, loc="upper left", bbox_to_anchor=(0, -0.22))
    fig.tight_layout()
    return fig
