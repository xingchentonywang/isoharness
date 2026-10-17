"""Saving a real run and reading it back, against a real Postgres.

Uses the 2026-09-15 sequence rather than invented numbers, so the round trip is
checked against values the rest of the suite already holds to the workbook.
"""
import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import psycopg
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ea_irms.parse import parse_export                              # noqa: E402
from ea_irms.pipeline import Config, config_to_dict, process        # noqa: E402
from ea_irms import store                                           # noqa: E402

DSN_SUPER = os.environ.get("EA_TEST_SUPER_DSN",
                           "postgresql://postgres@/ea_test?host=/tmp&port=5433")
DSN_APP = os.environ.get("EA_TEST_APP_DSN",
                         "postgresql://ea_login@/ea_test?host=/tmp&port=5433")
MIGRATIONS = sorted((Path(__file__).resolve().parents[1] / "db").glob("*.sql"))
# The 2026-09-15 export is the Raw sheet of the reference template; see
# test_matches_workbook.py for where the template is looked for.
TEMPLATE = Path(os.environ.get(
    "EA_TEMPLATE",
    Path(__file__).resolve().parent / "data" / "EA-IRMS_Correction_Template_v1.xlsx"))
if not TEMPLATE.exists():
    pytest.skip(f"reference template not found at {TEMPLATE}; set EA_TEMPLATE",
                allow_module_level=True)


def _export_from_template(folder):
    import csv
    import openpyxl
    path = folder / "20260915_raw.csv"
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        for record in openpyxl.load_workbook(TEMPLATE)["Raw"].iter_rows(
                max_col=16, values_only=True):
            if record[0] is not None:
                w.writerow(["" if v is None else v for v in record])
    return path

ALICE = store.Identity("google|alice", "alice@lab-a.edu", "Alice")
BOB = store.Identity("google|bob", "bob@lab-b.edu", "Bob")
CAROL = store.Identity("google|carol", "carol@lab-c.edu", "Carol")
DAVE = store.Identity("google|dave", "dave@lab-c.edu", "Dave")
TABLES = ("labs", "lab_members", "instruments", "lab_configs",
          "runs", "positions", "check_results")


@pytest.fixture(scope="module")
def export(tmp_path_factory):
    return _export_from_template(tmp_path_factory.mktemp("export"))


@pytest.fixture(scope="module")
def saved(export):
    subprocess.run(["psql", DSN_SUPER, "-q", "-c",
                    "DROP SCHEMA IF EXISTS public CASCADE; "
                    "DROP SCHEMA IF EXISTS app CASCADE; CREATE SCHEMA public;"],
                   check=True, capture_output=True)
    for migration in MIGRATIONS:
        subprocess.run(["psql", DSN_SUPER, "-q", "-v", "ON_ERROR_STOP=1", "-f", str(migration)],
                       check=True, capture_output=True)
    with psycopg.connect(DSN_SUPER, autocommit=True) as conn:
        conn.execute("DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_roles "
                     "WHERE rolname='ea_login') THEN CREATE ROLE ea_login LOGIN "
                     "NOBYPASSRLS; END IF; END $$;")
        conn.execute("GRANT ea_app TO ea_login")

    store.register_user(DSN_SUPER, ALICE)
    store.register_user(DSN_SUPER, BOB)
    lab = store.create_lab(DSN_SUPER, ALICE, "Wang Lab", "Boston College")
    store.create_lab(DSN_SUPER, BOB, "Other Lab")

    positions, fits = process(parse_export(export), Config())
    raw = export.read_bytes()
    with store.connect(DSN_APP, ALICE) as conn:
        instrument = store.get_or_create_instrument(conn, lab, "Delta V", "Thermo")
        run_id = store.save_run(
            conn, lab_id=lab, instrument_id=instrument, identity=ALICE,
            positions=positions, fits=fits, config=config_to_dict(Config()),
            tool_version="1.1.0", run_name="20260915 TestRun", run_date="2026-09-15",
            raw_bytes=raw, raw_filename=export.name)
    return {"lab": lab, "run": run_id, "instrument": instrument,
            "positions": positions, "fits": fits, "raw": raw}


def test_the_run_is_stored_with_its_fitted_coefficients(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        runs = store.lab_history(conn, saved["lab"])
    assert len(runs) == 1
    run = runs[0]
    fits = saved["fits"]
    assert run["n_positions"] == 56
    assert run["instrument"] == "Delta V"
    assert run["cal_slope_n"] == pytest.approx(fits["cal_n"]["slope"], abs=1e-12)
    assert run["cal_int_n"] == pytest.approx(fits["cal_n"]["intercept"], abs=1e-12)
    assert run["lin_slope_n"] == pytest.approx(fits["lin_n"]["fit"]["slope"], abs=1e-12)
    assert run["drift_model_n"] == fits["drift_n"]["model"]
    assert run["settled_from"] == fits["settling"]["settled_from"]
    assert run["bgd_floor"] == pytest.approx(fits["settling"]["floor"], abs=1e-12)


def test_every_position_survives_the_round_trip(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        rows = store.run_positions(conn, saved["run"])
    assert len(rows) == len(saved["positions"]) == 56
    by_position = {r["position"]: r for r in rows}
    worst = 0.0
    for p in saved["positions"]:
        row = by_position[p.row]
        assert row["identifier"] == p.identifier
        assert row["role"] == p.role
        # cn_molar carries the elemental drift correction from 1.4.2 (the reported value)
        for column, attribute in (("d15n_final", "d15n_final"), ("d13c_final", "d13c_final"),
                                  ("cn_molar", "cn_molar_reported"), ("pct_n", "pct_n_reported"),
                                  ("ref_bgd_n", "ref_bgd_n")):
            mine, theirs = getattr(p, attribute), row[column]
            assert (mine is None) == (theirs is None), f"position {p.row} {column}"
            if mine is not None:
                worst = max(worst, abs(mine - theirs))
    assert worst < 1e-12, f"worst round-trip difference {worst:.3e}"


def test_flags_are_kept(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        rows = {r["position"]: r for r in store.run_positions(conn, saved["run"])}
    assert "ran before the background settled" in (rows[1]["flags"] or "")
    assert rows[20]["flags"] is None or "no N peak" not in rows[20]["flags"]


def test_check_standards_are_recorded_with_their_offsets(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        rows = conn.execute("SELECT * FROM check_results WHERE run_id = %s",
                            (saved["run"],)).fetchall()
    checks = [p for p in saved["positions"] if p.role == "Check"]
    assert len(rows) == len(checks)


def test_the_raw_export_comes_back_byte_for_byte(saved, export):
    """Provenance depends on this: any run must be recorrectable years later."""
    with store.connect(DSN_APP, ALICE) as conn:
        name, data = store.raw_export(conn, saved["run"])
    assert name == export.name
    assert data == saved["raw"]


def test_the_raw_export_is_stored_compressed(saved):
    with psycopg.connect(DSN_SUPER) as conn:
        stored = conn.execute("SELECT length(raw_export) AS n FROM runs WHERE id = %s",
                              (saved["run"],)).fetchone()[0]
    assert stored < len(saved["raw"]) / 3, "the raw export is not being compressed"


def test_another_lab_cannot_see_the_run(saved):
    """The isolation suite proves the policies. This proves this code uses them."""
    with store.connect(DSN_APP, BOB) as conn:
        assert store.lab_history(conn, saved["lab"]) == []
        assert store.run_positions(conn, saved["run"]) == []
        assert store.raw_export(conn, saved["run"]) == (None, None)


def test_a_connection_without_an_identity_sees_nothing(saved):
    with store.connect(DSN_APP, None) as conn:
        assert store.lab_history(conn, saved["lab"]) == []
        assert store.labs_for(conn) == []


@pytest.fixture
def third_lab(saved):
    """A lab of its own, so that deleting it cannot disturb the other tests.

    Carol owns it and Dave is a member who is not an owner. It holds a copy of
    the saved run, one instrument, one lab configuration, and one check result
    (the run has no check standard under the default configuration), so every
    table has a row in it. A new one is made for every test that asks for it.
    """
    store.register_user(DSN_SUPER, CAROL)
    store.register_user(DSN_SUPER, DAVE)
    lab = store.create_lab(DSN_SUPER, CAROL, "Third Lab")
    with store.connect(DSN_APP, CAROL) as conn:
        instrument = store.get_or_create_instrument(conn, lab, "Delta V")
        run = store.save_run(conn, lab_id=lab, instrument_id=instrument, identity=CAROL,
                             positions=saved["positions"], fits=saved["fits"],
                             config={}, tool_version="1.1.0", run_name="copy",
                             raw_bytes=saved["raw"], raw_filename="copy.xls")
        conn.execute("INSERT INTO check_results (lab_id, run_id, material, n, measured_d15n, "
                     "accepted_d15n, offset_d15n) VALUES (%s, %s, 'USGS 66', 1, 40.9, 40.83, 0.07)",
                     (lab, run))
        conn.execute("INSERT INTO lab_configs (lab_id, name, config, is_default) "
                     "VALUES (%s, 'Corals', %s, true)",
                     (lab, json.dumps({"drift_model": "row"})))
        conn.execute("INSERT INTO lab_members (lab_id, user_id) VALUES (%s, %s)",
                     (lab, DAVE.user_id))
    return lab


def _rows_of(lab):
    """How many rows the lab has in each table, as the migration role sees them."""
    with psycopg.connect(DSN_SUPER) as conn:
        return {table: conn.execute(
                    f"SELECT count(*) FROM {table} "
                    f"WHERE {'id' if table == 'labs' else 'lab_id'} = %s", (lab,)).fetchone()[0]
                for table in TABLES}


def test_export_and_delete_return_a_lab_its_data_and_then_remove_it(saved, third_lab):
    """Both buttons have to work before other labs are asked to trust this."""
    lab = third_lab
    with store.connect(DSN_APP, CAROL) as conn:
        bundle = store.export_lab(conn, lab)
    assert bundle["lab"]["name"] == "Third Lab"
    assert len(bundle["runs"]) == 1
    assert len(bundle["positions"]) == 56
    assert len(bundle["check_results"]) == 1 + sum(p.role == "Check" for p in saved["positions"])
    assert [i["name"] for i in bundle["instruments"]] == ["Delta V"]
    assert [(c["name"], c["config"], c["is_default"]) for c in bundle["lab_configs"]] == \
        [("Corals", {"drift_model": "row"}, True)]
    assert {m["user_id"]: m["role"] for m in bundle["lab_members"]} == \
        {CAROL.user_id: "owner", DAVE.user_id: "member"}
    assert all(set(m) == {"lab_id", "user_id", "role", "added_at"}
               for m in bundle["lab_members"])
    # The schema it names is the newest migration, so store.SCHEMA keeps up with db/.
    assert bundle["schema"] == MIGRATIONS[-1].name[:3]
    exported = datetime.fromisoformat(bundle["exported_at"])
    assert exported.utcoffset() == timedelta(0)
    assert abs(datetime.now(timezone.utc) - exported) < timedelta(minutes=5)

    with store.connect(DSN_APP, CAROL) as conn:
        store.delete_lab_data(conn, lab)
    assert _rows_of(lab) == dict.fromkeys(TABLES, 0), "the lab or something under it is left"

    with store.connect(DSN_APP, CAROL) as conn:
        assert store.lab_history(conn, lab) == []
        assert store.export_lab(conn, lab)["lab"] is None
    with store.connect(DSN_APP, ALICE) as conn:
        assert len(store.lab_history(conn, saved["lab"])) == 1, \
            "deleting one lab's data touched another lab"


def test_any_member_can_export_the_lab(third_lab):
    with store.connect(DSN_APP, CAROL) as conn:
        by_owner = store.export_lab(conn, third_lab)
    with store.connect(DSN_APP, DAVE) as conn:
        by_member = store.export_lab(conn, third_lab)
    with store.connect(DSN_APP, BOB) as conn:
        by_stranger = store.export_lab(conn, third_lab)
    for bundle in (by_owner, by_member, by_stranger):
        del bundle["exported_at"]
    assert by_member == by_owner
    assert by_stranger["lab"] is None
    assert all(by_stranger[t] == [] for t in TABLES if t != "labs"), \
        "someone outside the lab exported some of it"


def test_labs_for_gives_each_lab_once_with_the_callers_own_role(third_lab):
    """isoharness 0.5.0: the lab list joined every member's row, so a lab with two members came back twice."""
    with store.connect(DSN_APP, CAROL) as conn:
        mine = [(str(r["id"]), r["role"]) for r in store.labs_for(conn) if str(r["id"]) == third_lab]
    with store.connect(DSN_APP, DAVE) as conn:
        theirs = [(str(r["id"]), r["role"]) for r in store.labs_for(conn) if str(r["id"]) == third_lab]
    with store.connect(DSN_APP, BOB) as conn:
        stranger = [r for r in store.labs_for(conn) if str(r["id"]) == third_lab]
    assert mine == [(third_lab, "owner")]
    assert theirs == [(third_lab, "member")]
    assert stranger == []


def test_only_an_owner_can_delete_the_lab(third_lab):
    """A member who is not an owner, or someone from another lab, is told so and
    nothing is removed."""
    before = _rows_of(third_lab)
    assert all(before.values())
    for who in (DAVE, BOB):
        with pytest.raises(PermissionError, match="only an owner"):
            with store.connect(DSN_APP, who) as conn:
                store.delete_lab_data(conn, third_lab)
        assert _rows_of(third_lab) == before, f"{who.name} removed something"


def test_deleting_a_lab_another_owner_has_just_deleted_raises(third_lab):
    """Two owners pressing delete at once. The second DELETE waits for the first,
    then finds no row, and must say so rather than return as if it had worked."""
    with store.connect(DSN_APP, CAROL) as conn:
        conn.execute("UPDATE lab_members SET role = 'owner' WHERE lab_id = %s AND user_id = %s",
                     (third_lab, DAVE.user_id))
    outcome = {}

    def carol_deletes():
        try:
            with store.connect(DSN_APP, CAROL) as conn:
                store.delete_lab_data(conn, third_lab)
            outcome["carol"] = "deleted"
        except Exception as exc:          # reported by the assertion below
            outcome["carol"] = exc

    thread = threading.Thread(target=carol_deletes)
    with store.connect(DSN_APP, DAVE) as conn:
        assert conn.execute("DELETE FROM labs WHERE id = %s", (third_lab,)).rowcount == 1
        dave = conn.info.backend_pid
        thread.start()
        with psycopg.connect(DSN_SUPER, autocommit=True) as admin:
            deadline = time.monotonic() + 10
            while thread.is_alive() and time.monotonic() < deadline and not admin.execute(
                    "SELECT count(*) FROM pg_stat_activity WHERE %s = ANY(pg_blocking_pids(pid))",
                    (dave,)).fetchone()[0]:
                time.sleep(0.02)
    thread.join(timeout=30)
    assert isinstance(outcome.get("carol"), LookupError), outcome
    assert _rows_of(third_lab) == dict.fromkeys(TABLES, 0)


def test_identity_does_not_outlive_the_transaction(saved, monkeypatch):
    """A connection handed on after the block must carry no identity.

    Every isolation test opens its own connection, so none of them can see what a
    reused connection carries. Here the connection is kept open past the block,
    as a pool would keep it, and inspected afterwards.
    """
    kept = []
    real_connect = psycopg.connect

    class Pooled:
        def __init__(self, conn):
            self.conn = conn

        def __enter__(self):
            return self.conn

        def __exit__(self, *exc):
            return False          # a pool returns the connection instead of closing it

    def pooled_connect(*a, **kw):
        conn = real_connect(*a, **kw)
        kept.append(conn)
        return Pooled(conn)

    monkeypatch.setattr(store.psycopg, "connect", pooled_connect)
    with store.connect(DSN_APP, ALICE) as conn:
        assert len(store.lab_history(conn, saved["lab"])) == 1
    monkeypatch.undo()

    conn = kept[0]
    try:
        leftover = conn.execute("SELECT current_setting('app.user_id', true)").fetchone()
        assert leftover[next(iter(leftover))] in (None, ""), \
            f"identity survived the transaction: {leftover!r}"
        assert conn.execute("SELECT count(*) AS n FROM runs").fetchone()["n"] == 0
    finally:
        conn.close()


def test_decisions_and_uncertainties_are_stored(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        run = store.lab_history(conn, saved["lab"])[0]
        rows = store.run_positions(conn, saved["run"])
    fits = saved["fits"]
    assert run["lin_in_force_c"] is fits["lin_c"]["in_force"]
    assert run["drift_p_n"] == pytest.approx(fits["drift_n"]["fit"]["p"], abs=1e-12)
    assert run["ref_pulse_n_n"] == fits["ref_gas_n"]["n"] == 56
    assert run["u_sample_n"] == pytest.approx(fits["unc_n"]["u_sample"], abs=1e-12)
    stored = {r["position"]: r for r in rows}
    for p in saved["positions"]:
        if p.u_d15n is not None:
            assert stored[p.row]["u_d15n"] == pytest.approx(p.u_d15n, abs=1e-12)


def test_check_history_reads_back_in_run_order(saved):
    with store.connect(DSN_APP, ALICE) as conn:
        conn.execute(
            "INSERT INTO check_results (lab_id, run_id, material, n, measured_d15n, accepted_d15n, "
            "offset_d15n, measured_d13c, accepted_d13c, offset_d13c) "
            "VALUES (%s, %s, 'USGS 66', 1, 40.9, 40.83, 0.07, -0.7, -0.67, -0.03)",
            (saved["lab"], saved["run"]))
    with store.connect(DSN_APP, ALICE) as conn:
        rows = store.check_history(conn, saved["lab"], "USGS 66")
        nothing = store.check_history(conn, saved["lab"], "USGS 40")
    assert len(rows) == 1 and rows[0]["offset_d15n"] == pytest.approx(0.07)
    assert rows[0]["instrument"] == "Delta V" and rows[0]["run_day"] is not None
    assert nothing == []
    with store.connect(DSN_APP, BOB) as conn:
        assert store.check_history(conn, saved["lab"]) == [], "another lab's history leaked"
