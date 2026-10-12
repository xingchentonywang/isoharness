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
