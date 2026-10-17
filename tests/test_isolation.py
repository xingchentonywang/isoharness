"""One lab must never see another lab's rows.

This is the test the whole design hangs on. It runs the real migration against a
real Postgres and connects as the same unprivileged role the application uses, so
a missing policy, a forgotten FORCE, or a table added later without one all make
it fail. Asserting isolation in Python rather than in the database would only
prove that the test and the code share the same mistake.
"""
import glob
import os
import subprocess
import threading
import time
import uuid

import psycopg
import pytest

DSN_SUPER = os.environ.get("EA_TEST_SUPER_DSN",
                           "postgresql://postgres@/ea_test?host=/tmp&port=5433")
DSN_APP = os.environ.get("EA_TEST_APP_DSN",
                         "postgresql://ea_login@/ea_test?host=/tmp&port=5433")
MIGRATIONS = sorted(glob.glob(os.path.join(os.path.dirname(__file__), "..", "db", "*.sql")))

TABLES = ["labs", "lab_members", "instruments", "lab_configs",
          "runs", "positions", "check_results"]


@pytest.fixture(scope="module")
def seeded():
    """A clean database holding two labs that must never see each other."""
    subprocess.run(["psql", DSN_SUPER, "-q", "-c",
                    "DROP SCHEMA IF EXISTS public CASCADE; "
                    "DROP SCHEMA IF EXISTS app CASCADE; CREATE SCHEMA public;"],
                   check=True, capture_output=True)
    for migration in MIGRATIONS:
        subprocess.run(["psql", DSN_SUPER, "-q", "-v", "ON_ERROR_STOP=1", "-f", migration],
                       check=True, capture_output=True)
    ids = {}
    with psycopg.connect(DSN_SUPER, autocommit=True) as conn:
        conn.execute("DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_roles "
                     "WHERE rolname='ea_login') THEN CREATE ROLE ea_login LOGIN "
                     "NOBYPASSRLS; END IF; END $$;")
        conn.execute("GRANT ea_app TO ea_login")
        for tag, user, email in (("a", "google|alice", "alice@lab-a.edu"),
                                 ("b", "google|bob", "bob@lab-b.edu")):
            lab = uuid.uuid4()
            conn.execute("INSERT INTO app_users (id, email) VALUES (%s, %s)", (user, email))
            conn.execute("INSERT INTO labs (id, name) VALUES (%s, %s)", (lab, f"Lab {tag}"))
            conn.execute("INSERT INTO lab_members (lab_id, user_id, role) "
                         "VALUES (%s, %s, 'owner')", (lab, user))
            inst = uuid.uuid4()
            conn.execute("INSERT INTO instruments (id, lab_id, name) VALUES (%s, %s, %s)",
                         (inst, lab, f"EA {tag}"))
            run = uuid.uuid4()
            conn.execute("INSERT INTO runs (id, lab_id, instrument_id, uploaded_by, "
                         "tool_version, config, run_name) "
                         "VALUES (%s, %s, %s, %s, '1.1.0', '{}'::jsonb, %s)",
                         (run, lab, inst, user, f"run {tag}"))
            conn.execute("INSERT INTO positions (lab_id, run_id, position, identifier) "
                         "VALUES (%s, %s, 1, %s)", (lab, run, f"sample {tag}"))
            conn.execute("INSERT INTO check_results (lab_id, run_id, material) "
                         "VALUES (%s, %s, 'USGS 66')", (lab, run))
            conn.execute("INSERT INTO lab_configs (lab_id, name, config) "
                         "VALUES (%s, 'default', '{}'::jsonb)", (lab,))
            ids[tag] = {"lab": lab, "user": user, "instrument": inst, "run": run}
    yield ids


def as_user(user_id):
    """A connection that can see only what that user is allowed to see."""
    conn = psycopg.connect(DSN_APP)
    conn.execute("SELECT set_config('app.user_id', %s, false)", (user_id,))
    return conn


@pytest.mark.parametrize("table", TABLES)
def test_each_user_sees_only_their_own_lab(seeded, table):
    for tag, other in (("a", "b"), ("b", "a")):
        with as_user(seeded[tag]["user"]) as conn:
            column = "id" if table == "labs" else "lab_id"
            rows = conn.execute(f"SELECT {column} FROM {table}").fetchall()
            assert rows, f"{tag} sees nothing in {table}, the policy is too strict"
            mine, theirs = seeded[tag]["lab"], seeded[other]["lab"]
            assert all(r[0] == mine for r in rows), f"{tag} can read {other}'s {table}"
            assert not any(r[0] == theirs for r in rows)


@pytest.mark.parametrize("table", TABLES)
def test_nothing_is_visible_without_an_identity(seeded, table):
    """A bug that forgets to set the identity must show nothing, not everything."""
    with psycopg.connect(DSN_APP) as conn:
        assert conn.execute(f"SELECT count(*) FROM {table}").fetchone()[0] == 0


def test_cannot_write_into_another_lab(seeded):
    """Reading is only half of it. WITH CHECK stops a forged lab_id on insert."""
    with as_user(seeded["a"]["user"]) as conn:
        with pytest.raises(psycopg.errors.InsufficientPrivilege):
            conn.execute("INSERT INTO runs (lab_id, instrument_id, tool_version, config) "
                         "VALUES (%s, %s, '1.1.0', '{}'::jsonb)",
                         (seeded["b"]["lab"], seeded["b"]["instrument"]))


def test_cannot_delete_another_labs_run(seeded):
    with as_user(seeded["a"]["user"]) as conn:
        conn.execute("DELETE FROM runs WHERE id = %s", (seeded["b"]["run"],))
        assert conn.execute("SELECT count(*) FROM runs").rowcount is not None
        conn.rollback()
    with as_user(seeded["b"]["user"]) as conn:
        still_there = conn.execute("SELECT count(*) FROM runs WHERE id = %s",
                                   (seeded["b"]["run"],)).fetchone()[0]
    assert still_there == 1, "lab A deleted lab B's run"


def test_cannot_join_a_lab_by_inserting_a_membership(seeded):
    """The obvious escalation: grant yourself membership of someone else's lab."""
    with as_user(seeded["a"]["user"]) as conn:
        with pytest.raises(psycopg.errors.InsufficientPrivilege):
            conn.execute("INSERT INTO lab_members (lab_id, user_id) VALUES (%s, %s)",
                         (seeded["b"]["lab"], seeded["a"]["user"]))


def test_the_application_role_cannot_bypass_policies(seeded):
    with psycopg.connect(DSN_APP) as conn:
        row = conn.execute("SELECT rolsuper, rolbypassrls FROM pg_roles "
                           "WHERE rolname = current_user").fetchone()
    assert row == (False, False), "the app role can bypass row level security"


def test_every_table_is_protected(seeded):
    """Catches a table added later without a policy, which is how this design rots."""
    with psycopg.connect(DSN_SUPER) as conn:
        unprotected = conn.execute("""
            SELECT c.relname FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
            WHERE n.nspname = 'public' AND c.relkind = 'r'
              AND (NOT c.relrowsecurity OR NOT c.relforcerowsecurity
                   OR (SELECT count(*) FROM pg_policy p WHERE p.polrelid = c.oid) = 0)
        """).fetchall()
    assert unprotected == [], f"tables without forced RLS and a policy: {unprotected}"


def _set_rls(table, enabled):
    with psycopg.connect(DSN_SUPER, autocommit=True) as admin:
        if enabled:
            admin.execute(f"ALTER TABLE {table} ENABLE ROW LEVEL SECURITY")
            admin.execute(f"ALTER TABLE {table} FORCE ROW LEVEL SECURITY")
        else:
            admin.execute(f"ALTER TABLE {table} DISABLE ROW LEVEL SECURITY")


def _lab_b_positions_visible_to_a(seeded):
    with as_user(seeded["a"]["user"]) as conn:
        return conn.execute("SELECT count(*) FROM positions WHERE lab_id = %s",
                            (seeded["b"]["lab"],)).fetchone()[0]


def test_the_protection_is_what_is_doing_the_work(seeded):
    """A negative control, because a suite that would pass anyway proves nothing.

    Take the mechanism away and the isolation must break. Without this, every
    test above would still pass if the policies were doing nothing at all and the
    two labs simply never had overlapping rows to confuse them.

    Disabling row level security is what opens the table. Dropping the policy
    alone would close it completely, since a table with RLS on and no policy
    shows no rows to anyone.
    """
    assert _lab_b_positions_visible_to_a(seeded) == 0, "isolation is already broken"
    _set_rls("positions", False)
    try:
        leaked = _lab_b_positions_visible_to_a(seeded)
    finally:
        _set_rls("positions", True)
    assert leaked == 1, ("switching off row level security changed nothing, so the "
                         "tests above are not evidence that anything is enforced")
    assert _lab_b_positions_visible_to_a(seeded) == 0, "the control failed to restore"


# ---------------------------------------------------------------- owners (db/003)
# From 003 the role in lab_members means something. A member who is not an owner
# still reads everything and adds runs, but cannot change who belongs to the lab,
# change or remove a saved record, or delete the lab. Most of these tests delete
# something, so each one gets a lab of its own.

@pytest.fixture
def lab(seeded):
    """A new lab: one owner, two members who are not owners, and a row in every table.

    Made with the migration role, like `seeded`, and made again for every test,
    so a test can delete or demote anything without touching labs a and b or
    another test's lab. `newcomer` has an account but belongs to no lab.
    """
    tag = uuid.uuid4().hex[:8]
    ids = {"lab": uuid.uuid4(), "instrument": uuid.uuid4(), "config": uuid.uuid4(),
           "run": uuid.uuid4()}
    with psycopg.connect(DSN_SUPER, autocommit=True) as conn:
        for key in ("owner", "member", "other", "newcomer"):
            ids[key] = f"google|{key}-{tag}"
            conn.execute("INSERT INTO app_users (id, email) VALUES (%s, %s)",
                         (ids[key], f"{key}-{tag}@lab-c.edu"))
        conn.execute("INSERT INTO labs (id, name) VALUES (%(lab)s, 'Lab c')", ids)
        for key, role in (("owner", "owner"), ("member", "member"), ("other", "member")):
            conn.execute("INSERT INTO lab_members (lab_id, user_id, role) VALUES (%s, %s, %s)",
                         (ids["lab"], ids[key], role))
        conn.execute("INSERT INTO instruments (id, lab_id, name) "
                     "VALUES (%(instrument)s, %(lab)s, 'EA c')", ids)
        conn.execute("INSERT INTO lab_configs (id, lab_id, name, config) "
                     "VALUES (%(config)s, %(lab)s, 'default', '{}'::jsonb)", ids)
        conn.execute("INSERT INTO runs (id, lab_id, instrument_id, uploaded_by, tool_version, "
                     "config, run_name) VALUES (%(run)s, %(lab)s, %(instrument)s, %(member)s, "
                     "'1.4.12', '{}'::jsonb, 'run c')", ids)
        conn.execute("INSERT INTO positions (lab_id, run_id, position, identifier) "
                     "VALUES (%(lab)s, %(run)s, 1, 'sample c')", ids)
        conn.execute("INSERT INTO check_results (lab_id, run_id, material) "
                     "VALUES (%(lab)s, %(run)s, 'USGS 66')", ids)
    return ids


def rows_of(lab_id):
    """Every row the lab has, table by table, as the migration role sees them."""
    with psycopg.connect(DSN_SUPER) as conn:
        return {table: conn.execute(
                    f"SELECT t::text FROM {table} t "
                    f"WHERE {'id' if table == 'labs' else 'lab_id'} = %s ORDER BY 1",
                    (lab_id,)).fetchall()
                for table in TABLES}


def owners_of(lab_id):
    with psycopg.connect(DSN_SUPER) as conn:
        return sorted(r[0] for r in conn.execute(
            "SELECT user_id FROM lab_members WHERE lab_id = %s AND role = 'owner'",
            (lab_id,)).fetchall())


# What only an owner may do. Run by the owner, each statement changes one row of
# the lab fixture's lab. Run by a member who is not an owner, it must change
# nothing; row level security does not raise on an UPDATE or DELETE, it hides
# the rows, so the statement affects none and the tests count rows.
OWNER_ONLY = {
    "remove another member":
        "DELETE FROM lab_members WHERE lab_id = %(lab)s AND user_id = %(other)s",
    "update another member's row":
        "UPDATE lab_members SET added_at = now() WHERE lab_id = %(lab)s AND user_id = %(other)s",
    "change a role":
        "UPDATE lab_members SET role = 'owner' WHERE lab_id = %(lab)s AND user_id = %(other)s",
    "delete a run": "DELETE FROM runs WHERE id = %(run)s",
    "update a run": "UPDATE runs SET run_name = 'changed' WHERE id = %(run)s",
    "delete a position": "DELETE FROM positions WHERE run_id = %(run)s",
    "update a position": "UPDATE positions SET d15n_final = 0 WHERE run_id = %(run)s",
    "delete a check result": "DELETE FROM check_results WHERE run_id = %(run)s",
    "update a check result": "UPDATE check_results SET offset_d15n = 0 WHERE run_id = %(run)s",
    "delete an instrument": "DELETE FROM instruments WHERE id = %(instrument)s",
    "delete a configuration": "DELETE FROM lab_configs WHERE id = %(config)s",
    "rename the lab": "UPDATE labs SET name = 'changed' WHERE id = %(lab)s",
    "delete the lab": "DELETE FROM labs WHERE id = %(lab)s",
}


@pytest.mark.parametrize("action", OWNER_ONLY)
def test_a_member_who_is_not_an_owner_cannot(lab, action):
    before = rows_of(lab["lab"])
    with as_user(lab["member"]) as conn:
        assert conn.execute(OWNER_ONLY[action], lab).rowcount == 0, f"a member could {action}"
    assert rows_of(lab["lab"]) == before


@pytest.mark.parametrize("action", OWNER_ONLY)
def test_an_owner_can(lab, action):
    """The other half: without it, a policy that refused everyone would pass the test above."""
    before = rows_of(lab["lab"])
    with as_user(lab["owner"]) as conn:
        assert conn.execute(OWNER_ONLY[action], lab).rowcount == 1, f"the owner could not {action}"
    assert rows_of(lab["lab"]) != before


def test_only_an_owner_can_add_a_member(lab):
    """An INSERT that no policy admits is an error, unlike a hidden UPDATE or DELETE."""
    add = "INSERT INTO lab_members (lab_id, user_id) VALUES (%(lab)s, %(newcomer)s)"
    with as_user(lab["member"]) as conn:
        with pytest.raises(psycopg.errors.InsufficientPrivilege):
            conn.execute(add, lab)
    with as_user(lab["newcomer"]) as conn:
        assert conn.execute("SELECT count(*) FROM runs").fetchone()[0] == 0
    with as_user(lab["owner"]) as conn:
        assert conn.execute(add, lab).rowcount == 1
    with as_user(lab["newcomer"]) as conn:
        assert conn.execute("SELECT count(*) FROM runs").fetchone()[0] == 1


def test_a_member_cannot_make_themselves_an_owner(lab):
    with as_user(lab["member"]) as conn:
        assert conn.execute("UPDATE lab_members SET role = 'owner' "
                            "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab).rowcount == 0
    assert owners_of(lab["lab"]) == [lab["owner"]]


def test_a_member_who_is_not_an_owner_still_reads_and_adds(lab):
    """The rules must not be stricter than meant: members keep the lab running."""
    with as_user(lab["member"]) as conn:
        for table in TABLES:
            assert conn.execute(f"SELECT count(*) FROM {table}").fetchone()[0] >= 1, table
        conn.execute("INSERT INTO instruments (lab_id, name) VALUES (%(lab)s, 'EA d')", lab)
        assert conn.execute("UPDATE instruments SET model = 'Flash' "
                            "WHERE id = %(instrument)s", lab).rowcount == 1
        conn.execute("INSERT INTO lab_configs (lab_id, name, config) "
                     "VALUES (%(lab)s, 'mine', '{}'::jsonb)", lab)
        assert conn.execute("UPDATE lab_configs SET is_default = true "
                            "WHERE id = %(config)s", lab).rowcount == 1
        run = conn.execute("INSERT INTO runs (lab_id, instrument_id, tool_version, config) "
                           "VALUES (%(lab)s, %(instrument)s, '1.4.12', '{}'::jsonb) "
                           "RETURNING id", lab).fetchone()[0]
        conn.execute("INSERT INTO positions (lab_id, run_id, position) VALUES (%s, %s, 1)",
                     (lab["lab"], run))
        conn.execute("INSERT INTO check_results (lab_id, run_id, material) "
                     "VALUES (%s, %s, 'USGS 66')", (lab["lab"], run))
    after = rows_of(lab["lab"])
    assert [len(after[t]) for t in ("instruments", "lab_configs", "runs", "positions",
                                    "check_results")] == [2, 2, 2, 2, 2]


def test_a_member_can_leave(lab):
    with as_user(lab["member"]) as conn:
        assert conn.execute("DELETE FROM lab_members "
                            "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab).rowcount == 1
    with as_user(lab["member"]) as conn:
        for table in TABLES:
            assert conn.execute(f"SELECT count(*) FROM {table}").fetchone()[0] == 0, \
                f"{table} is still visible after leaving the lab"
    assert len(rows_of(lab["lab"])["lab_members"]) == 2


def test_the_last_owner_cannot_be_removed_or_demoted(lab):
    remove = "DELETE FROM lab_members WHERE lab_id = %(lab)s AND user_id = %(owner)s"
    demote = "UPDATE lab_members SET role = 'member' WHERE lab_id = %(lab)s AND user_id = %(owner)s"
    for statement in (remove, demote):
        with as_user(lab["owner"]) as conn:
            with pytest.raises(psycopg.errors.CheckViolation, match="the last owner of lab"):
                conn.execute(statement, lab)
    # The rule is a trigger, not a policy, so the migration role meets it too.
    with psycopg.connect(DSN_SUPER) as conn:
        with pytest.raises(psycopg.errors.CheckViolation, match="the last owner of lab"):
            conn.execute(remove, lab)
    # So does the cascade from deleting the owner's account.
    with as_user(lab["owner"]) as conn:
        with pytest.raises(psycopg.errors.CheckViolation, match="the last owner of lab"):
            conn.execute("DELETE FROM app_users WHERE id = %(owner)s", lab)
    assert owners_of(lab["lab"]) == [lab["owner"]]


def test_with_a_second_owner_one_may_go_but_not_both(lab):
    promote = ("UPDATE lab_members SET role = 'owner' "
               "WHERE lab_id = %(lab)s AND user_id = %(member)s")
    with as_user(lab["owner"]) as conn:
        assert conn.execute(promote, lab).rowcount == 1
    with as_user(lab["owner"]) as conn:
        assert conn.execute("DELETE FROM lab_members "
                            "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab).rowcount == 1
    assert owners_of(lab["lab"]) == [lab["owner"]]

    # One statement that takes both owners away is refused as a whole.
    with as_user(lab["owner"]) as conn:
        conn.execute("INSERT INTO lab_members (lab_id, user_id, role) "
                     "VALUES (%(lab)s, %(member)s, 'owner')", lab)
    for statement in ("DELETE FROM lab_members WHERE lab_id = %(lab)s AND role = 'owner'",
                      "UPDATE lab_members SET role = 'member' WHERE lab_id = %(lab)s"):
        with as_user(lab["owner"]) as conn:
            with pytest.raises(psycopg.errors.CheckViolation, match="the last owner of lab"):
                conn.execute(statement, lab)
    assert owners_of(lab["lab"]) == sorted([lab["owner"], lab["member"]])

    # An owner may step down, or leave, while another owner stays.
    with as_user(lab["owner"]) as conn:
        assert conn.execute("UPDATE lab_members SET role = 'member' "
                            "WHERE lab_id = %(lab)s AND user_id = %(owner)s", lab).rowcount == 1
    assert owners_of(lab["lab"]) == [lab["member"]]
    with as_user(lab["member"]) as conn:
        with pytest.raises(psycopg.errors.CheckViolation, match="the last owner of lab"):
            conn.execute("DELETE FROM lab_members "
                         "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab)
    with as_user(lab["owner"]) as conn:
        assert conn.execute("DELETE FROM lab_members "
                            "WHERE lab_id = %(lab)s AND user_id = %(owner)s", lab).rowcount == 1
    assert owners_of(lab["lab"]) == [lab["member"]]


def _wait_until_waiting_for_a_lock(pid, thread, seconds=10.0):
    """True once backend `pid` waits for a lock; False if its thread ends first."""
    deadline = time.monotonic() + seconds
    with psycopg.connect(DSN_SUPER, autocommit=True) as admin:
        while thread.is_alive() and time.monotonic() < deadline:
            row = admin.execute("SELECT wait_event_type FROM pg_stat_activity WHERE pid = %s",
                                (pid,)).fetchone()
            if row and row[0] == "Lock":
                return True
            time.sleep(0.02)
    return False


def test_two_owners_leaving_at_once_cannot_both_go(lab):
    """Each would see the other stay, and the lab would be left with neither.

    The trigger takes a lock for the lab, so the second waits until the first
    has committed and then sees that it is the last owner.
    """
    with psycopg.connect(DSN_SUPER, autocommit=True) as admin:
        admin.execute("UPDATE lab_members SET role = 'owner' "
                      "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab)
    leave = "DELETE FROM lab_members WHERE lab_id = %(lab)s AND user_id = %(me)s"
    first, second = as_user(lab["owner"]), as_user(lab["member"])
    second.execute("SET lock_timeout = '20s'")      # never hang the suite
    second_pid = second.info.backend_pid
    outcome = {}

    def second_leaves():
        try:
            second.execute(leave, {**lab, "me": lab["member"]})
            second.commit()
            outcome["second"] = "left"
        except psycopg.Error as exc:
            outcome["second"] = exc

    try:
        assert first.execute(leave, {**lab, "me": lab["owner"]}).rowcount == 1
        thread = threading.Thread(target=second_leaves)
        thread.start()
        waited = _wait_until_waiting_for_a_lock(second_pid, thread)
        first.commit()
        thread.join(timeout=30)
    finally:
        first.close()
        second.close()
    assert waited, "the second owner did not wait for the first"
    assert isinstance(outcome.get("second"), psycopg.errors.CheckViolation), outcome
    assert owners_of(lab["lab"]) == [lab["member"]]


def test_an_owner_deleting_the_lab_removes_every_row_it_had(seeded, lab):
    """The cascade from labs passes the trigger, for every owner row it removes."""
    with psycopg.connect(DSN_SUPER, autocommit=True) as admin:
        admin.execute("UPDATE lab_members SET role = 'owner' "
                      "WHERE lab_id = %(lab)s AND user_id = %(member)s", lab)
    assert all(rows_of(lab["lab"]).values()), "the lab must start with a row in every table"
    with as_user(lab["owner"]) as conn:
        assert conn.execute("DELETE FROM labs WHERE id = %(lab)s", lab).rowcount == 1
    assert rows_of(lab["lab"]) == {table: [] for table in TABLES}
    for tag in ("a", "b"):
        assert all(rows_of(seeded[tag]["lab"]).values()), f"deleting a lab touched lab {tag}"
