"""A lab's configuration has to survive being saved, shared and reloaded.

It is the one thing another lab edits by hand, and a silent loss there would
reassign roles without anyone noticing.
"""
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ea_irms.pipeline import (Config, DEFAULT_STANDARDS, assign_roles,   # noqa: E402
                              config_from_dict, config_to_dict, standards_from_rows)
from ea_irms.parse import Position                                        # noqa: E402


def test_json_roundtrip_is_lossless():
    original = Config(blank_correction="Auto", drift_model="Background",
                      first_drift_position=27, min_area_n=8.889)
    restored = config_from_dict(json.loads(json.dumps(config_to_dict(original))))
    assert config_to_dict(restored) == config_to_dict(original)


def test_dataframe_roundtrip_keeps_every_field():
    """The app edits the standards in a grid, which is where columns get lost."""
    columns = ["match", "role", "d13c", "d15n", "pct_c", "pct_n", "material", "u13", "u15", "upc", "upn"]
    frame = pd.DataFrame([list(r) for r in DEFAULT_STANDARDS], columns=columns)
    back = standards_from_rows(frame.values.tolist())
    assert len(back) == len(DEFAULT_STANDARDS)
    for a, b in zip(back, DEFAULT_STANDARDS):
        assert len(a) == 11
        for x, y in zip(a, b):
            assert (x == y) or (x is None and y is None)


def test_blank_accepted_values_do_not_become_nan():
    """A float column with no value comes back as NaN, which would poison every
    corrected number silently."""
    columns = ["match", "role", "d13c", "d15n", "pct_c", "pct_n", "material", "u13", "u15", "upc", "upn"]
    frame = pd.DataFrame([list(r) for r in DEFAULT_STANDARDS], columns=columns)
    for row in standards_from_rows(frame.values.tolist()):
        for cell in row:
            assert not (isinstance(cell, float) and cell != cell)


def test_rows_without_match_text_are_dropped():
    assert standards_from_rows([[None, "Anchor", 1, 2, 3, 4, "x"], ["  ", "Check"]]) == []


def test_short_rows_are_padded():
    cfg = config_from_dict({"standards": [["My Std", "Anchor", -20.0, 5.0]]})
    assert cfg.standards[0] == ("My Std", "Anchor", -20.0, 5.0, None, None, None, None, None, None, None)


def test_roles_match_case_insensitively_and_most_specific_wins():
    cfg = Config()
    positions = [Position(row=1, identifier="USGS 64"),
                 Position(row=2, identifier="ace drift11"),
                 Position(row=3, identifier="blank capsule"),
                 Position(row=4, identifier="Coral BZ-14")]
    assign_roles(positions, cfg)
    assert [p.role for p in positions] == ["Anchor", "Drift", "Blank", "Sample"]


def test_a_lab_with_different_names_still_works():
    cfg = config_from_dict({"standards": [
        ["IAEA-600", "Anchor", -27.77, 1.0, 40.8, 9.5, "caffeine"],
        ["ACET", "Drift", None, None, 71.09, 10.36, "acetanilide"],
    ]})
    positions = [Position(row=1, identifier="IAEA-600 a"), Position(row=2, identifier="ACET-3")]
    assign_roles(positions, cfg)
    assert [p.role for p in positions] == ["Anchor", "Drift"]
    assert positions[0].acc13 == -27.77


def test_cli_writes_a_config_without_an_export(tmp_path):
    from ea_irms.cli import main
    target = tmp_path / "lab.json"
    assert main(["--write-config", str(target)]) == 0
    assert json.loads(target.read_text())["drift_model"] == "Row"


def test_cli_reports_a_broken_config_instead_of_raising(tmp_path, capsys):
    from ea_irms.cli import main
    bad = tmp_path / "bad.json"
    bad.write_text("{not json")
    assert main(["whatever.xls", "--config", str(bad)]) == 2
    assert "cannot read the configuration" in capsys.readouterr().err
