"""Nitrogen and carbon linearity can be switched apart (1.4.1).

Setup B11 (`linearity_correction`) sets both elements unless the carbon switch,
Setup B18 (`linearity_correction_c`), is filled. Left empty it must change
nothing, so a 1.4.0 configuration gives the same numbers.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from ea_irms.parse import parse_export                                   # noqa: E402
from ea_irms.pipeline import Config, config_from_dict, config_to_dict, process  # noqa: E402
from test_statistics import _sequence                                    # noqa: E402


def _run(tmp_path, **kw):
    return process(parse_export(_sequence(tmp_path)), Config(**kw))


def test_empty_carbon_switch_follows_b11(tmp_path):
    for mode in ("On", "Off", "Auto"):
        _, a = _run(tmp_path, linearity_correction=mode)
        _, b = _run(tmp_path, linearity_correction=mode, linearity_correction_c="")
        for el in ("lin_n", "lin_c"):
            assert a[el]["in_force"] == b[el]["in_force"]
            assert a[el]["applied"] == b[el]["applied"]
    _, f = _run(tmp_path, linearity_correction="Off")
    assert f["lin_n"]["in_force"] is False and f["lin_c"]["in_force"] is False


def test_nitrogen_off_keeps_carbon(tmp_path):
    split, fs = _run(tmp_path, linearity_correction="Off", linearity_correction_c="On")
    both_on, _ = _run(tmp_path, linearity_correction="On")
    both_off, _ = _run(tmp_path, linearity_correction="Off")
    assert fs["lin_n"]["in_force"] is False and fs["lin_n"]["applied"] == 0.0
    assert fs["lin_c"]["in_force"] is True and fs["lin_c"]["mode"] == "On"
    for p, on, off in zip(split, both_on, both_off):
        assert p.d13c_final == on.d13c_final       # carbon as with linearity on
        assert p.d15n_final == off.d15n_final      # nitrogen as with linearity off


def test_carbon_off_keeps_nitrogen(tmp_path):
    split, fs = _run(tmp_path, linearity_correction="On", linearity_correction_c="Off")
    assert fs["lin_n"]["in_force"] is True and fs["lin_c"]["in_force"] is False
    both_on, _ = _run(tmp_path, linearity_correction="On")
    for p, on in zip(split, both_on):
        assert p.d15n_final == on.d15n_final


def test_carbon_switch_survives_the_config_file():
    cfg = Config(linearity_correction="Off", linearity_correction_c="Auto")
    back = config_from_dict(json.loads(json.dumps(config_to_dict(cfg))))
    assert back.linearity_correction == "Off" and back.linearity_correction_c == "Auto"
    assert config_from_dict({"linearity_correction_c": False}).linearity_correction_c == "Off"
    assert config_from_dict({}).linearity_correction_c is None      # a 1.4.0 file
