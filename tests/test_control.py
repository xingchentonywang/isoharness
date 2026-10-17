"""Control-chart limits and the Westgard rules, per FIRMS guide section 7.1."""
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ea_irms.control import chart, evaluate, set_limits, westgard   # noqa: E402


def _history(n, days, base=-29.5, spread=0.05):
    out = []
    for i in range(n):
        out.append((date(2026, 1, 1) + timedelta(days=i % days), base + spread * ((-1) ** i) * ((i % 3) / 2)))
    return out


def test_limits_need_twenty_results_on_six_days():
    h = _history(19, 8)
    lim, why = set_limits([v for _, v in h], [d for d, _ in h])
    assert lim is None and "19 results" in why
    h = _history(25, 3)
    lim, why = set_limits([v for _, v in h], [d for d, _ in h])
    assert lim is None and "3 days" in why
    h = _history(25, 8)
    lim, why = set_limits([v for _, v in h], [d for d, _ in h])
    assert why is None and lim.n == 25 and lim.days == 8
    assert abs(lim.warning[1] - (lim.mean + 2 * lim.sd)) < 1e-12
    assert abs(lim.control[0] - (lim.mean - 3 * lim.sd)) < 1e-12


def test_westgard_rules_fire_where_they_should():
    mean, sd = 0.0, 1.0
    assert westgard([0.0, 3.5], mean, sd)[1] == ["1_3s"]
    assert "2_2s" in westgard([2.5, 2.5], mean, sd)[1]
    assert "2_2s" not in westgard([2.5, -2.5], mean, sd)[1]       # opposite sides
    assert "R_4s" in westgard([2.5, -2.5], mean, sd)[1]
    assert "4_1s" in westgard([1.2, 1.3, 1.1, 1.4], mean, sd)[3]
    assert "4_1s" not in westgard([1.2, 1.3, -1.1, 1.4], mean, sd)[3]
    ten = [0.3] * 10
    assert "10_x" in westgard(ten, mean, sd)[9]
    assert "10_x" not in westgard(ten, mean, sd)[8]
    assert westgard([0.5, -0.5, 1.5, -1.5], mean, sd) == [[], [], [], []]


def test_an_observation_on_a_limit_is_not_beyond_it():
    # -29.7 is 3 SD above a mean of -30.0 with an SD of 0.1 in decimal arithmetic and 3.000000000000007 SD in
    # floats; as at every other limit in ea_irms (finding R-24 of the 2026-10-14 review) it is not beyond 3 SD
    assert (-29.7 - -30.0) / 0.1 > 3 and westgard([-29.7], -30.0, 0.1) == [[]]
    assert westgard([-29.6999], -30.0, 0.1) == [["1_3s"]]
    assert (-29.4 - -29.6) / 0.1 > 2 and westgard([-29.4, -29.4], -29.6, 0.1) == [[], []]   # on the 2_2s limit


def test_evaluate_marks_the_offending_points():
    h = _history(25, 8)
    lim, _ = set_limits([v for _, v in h], [d for d, _ in h])
    later = h + [(date(2026, 3, 1), lim.mean + 3.5 * lim.sd)]
    pts = evaluate(later, lim)
    assert pts[-1].out_of_control and "1_3s" in pts[-1].rules
    assert not any(p.out_of_control for p in pts[:-1])


def test_chart_draws_with_and_without_limits(tmp_path):
    h = _history(25, 8)
    lim, _ = set_limits([v for _, v in h], [d for d, _ in h])
    fig = chart(h, lim, "acetanilide d13C", "d13C (permil)", target=-29.53)
    fig.savefig(tmp_path / "with.png")
    fig = chart(h[:5], None, "acetanilide d13C", "d13C (permil)")
    fig.savefig(tmp_path / "without.png")
    assert (tmp_path / "with.png").stat().st_size > 1000
