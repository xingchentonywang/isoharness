-- Owner-only rules.
--
-- In 001 every policy said one thing: you may read and write a row if you are a
-- member of the lab that owns it. That left the owner role in lab_members as a
-- label nothing enforced, so any member could add or remove members, make
-- themselves an owner, delete the lab's runs, or delete the lab. This file gives
-- each table one policy per command it allows. Members still read everything,
-- add runs, and keep the instruments and configurations up to date; only an
-- owner decides who belongs to the lab, changes or removes a saved run, deletes
-- an instrument or a configuration, or renames or deletes the lab.
--
-- Run once, after 001 and 002, as the role that ran them:
--     psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -f db/003_owner_rules.sql
-- It is one transaction, so it applies whole or not at all.

BEGIN;

-- The functions here, like is_member in 001, run as the role that creates them
-- and read lab_members past its policies. Only a role that row level security
-- does not filter can do that: a superuser, or a role with BYPASSRLS. Under any
-- other role the policies call is_member, which reads lab_members, whose policy
-- calls is_member again, until Postgres stops with "stack depth limit exceeded"
-- on the first query the application makes. So this file stops at once if it is
-- run as any other role.
DO $$ BEGIN
    IF NOT (SELECT rolsuper OR rolbypassrls FROM pg_roles WHERE rolname = current_user) THEN
        RAISE EXCEPTION 'run 003 as a superuser or a role with BYPASSRLS, not as %',
            current_user;
    END IF;
END $$;

-- ---------------------------------------------------------------- identity
-- The owner question, asked the same way as membership: SECURITY DEFINER, so a
-- policy on lab_members can ask it without recursing into its own policy.
CREATE OR REPLACE FUNCTION app.is_owner(target uuid) RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public, pg_temp AS $$
    SELECT EXISTS (
        SELECT 1 FROM lab_members
        WHERE lab_id = target AND user_id = app.current_user_id() AND role = 'owner'
    )
$$;
GRANT EXECUTE ON FUNCTION app.is_owner(uuid) TO ea_app;

-- ---------------------------------------------------------------- isolation
-- The seven policies of 001 each covered every command. They go, and each table
-- gets one policy per command it allows, so the list below is the whole of the
-- rules. A command with no policy is refused: an UPDATE or DELETE then matches
-- no rows and affects none, without an error, and an INSERT is rejected with
-- one. Row level security stays on and forced as 001 set it, so these rules bind
-- the table owner as well.
DROP POLICY IF EXISTS lab_access ON labs;
DROP POLICY IF EXISTS lab_members_access ON lab_members;
DROP POLICY IF EXISTS instruments_access ON instruments;
DROP POLICY IF EXISTS lab_configs_access ON lab_configs;
DROP POLICY IF EXISTS runs_access ON runs;
DROP POLICY IF EXISTS positions_access ON positions;
DROP POLICY IF EXISTS check_results_access ON check_results;

-- labs. Every member reads the lab; only an owner renames it or deletes it.
-- There is no INSERT policy: store.create_lab makes a lab with the migration
-- role, in the same transaction that makes its creator the first owner, because
-- nobody can be a member of a lab that does not exist yet.
CREATE POLICY labs_select ON labs FOR SELECT
    USING (app.is_member(id));
CREATE POLICY labs_update ON labs FOR UPDATE
    USING (app.is_owner(id)) WITH CHECK (app.is_owner(id));
CREATE POLICY labs_delete ON labs FOR DELETE
    USING (app.is_owner(id));

-- lab_members. Every member sees who else belongs. Adding a member, changing a
-- role, and removing someone each decide who may see the lab's data and what
-- they may do with it, so only an owner does them. Any member may remove their
-- own row, which is how a member leaves a lab. Whether the last owner may go is
-- the trigger's question, below, since a policy looks at one row and cannot
-- count the owners that would remain.
CREATE POLICY lab_members_select ON lab_members FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY lab_members_insert ON lab_members FOR INSERT
    WITH CHECK (app.is_owner(lab_id));
CREATE POLICY lab_members_update ON lab_members FOR UPDATE
    USING (app.is_owner(lab_id)) WITH CHECK (app.is_owner(lab_id));
CREATE POLICY lab_members_delete ON lab_members FOR DELETE
    USING (app.is_owner(lab_id));
CREATE POLICY lab_members_leave ON lab_members FOR DELETE
    USING (user_id = app.current_user_id());

-- instruments and lab_configs. These are the lab's working setup, which every
-- member keeps up to date: adding an instrument, saving a configuration, and
-- editing either. Only an owner deletes one, because deleting an instrument
-- also deletes every run measured on it (runs.instrument_id cascades), and a
-- configuration may be the one the rest of the lab works from.
CREATE POLICY instruments_select ON instruments FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY instruments_insert ON instruments FOR INSERT
    WITH CHECK (app.is_member(lab_id));
CREATE POLICY instruments_update ON instruments FOR UPDATE
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY instruments_delete ON instruments FOR DELETE
    USING (app.is_owner(lab_id));

CREATE POLICY lab_configs_select ON lab_configs FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY lab_configs_insert ON lab_configs FOR INSERT
    WITH CHECK (app.is_member(lab_id));
CREATE POLICY lab_configs_update ON lab_configs FOR UPDATE
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY lab_configs_delete ON lab_configs FOR DELETE
    USING (app.is_owner(lab_id));

-- runs, positions, and check_results. A saved run is a record of what was
-- measured and how it was corrected. Members add runs; only an owner changes or
-- removes one, so a record stays as it was saved unless an owner of the lab
-- decides otherwise.
CREATE POLICY runs_select ON runs FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY runs_insert ON runs FOR INSERT
    WITH CHECK (app.is_member(lab_id));
CREATE POLICY runs_update ON runs FOR UPDATE
    USING (app.is_owner(lab_id)) WITH CHECK (app.is_owner(lab_id));
CREATE POLICY runs_delete ON runs FOR DELETE
    USING (app.is_owner(lab_id));

CREATE POLICY positions_select ON positions FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY positions_insert ON positions FOR INSERT
    WITH CHECK (app.is_member(lab_id));
CREATE POLICY positions_update ON positions FOR UPDATE
    USING (app.is_owner(lab_id)) WITH CHECK (app.is_owner(lab_id));
CREATE POLICY positions_delete ON positions FOR DELETE
    USING (app.is_owner(lab_id));

CREATE POLICY check_results_select ON check_results FOR SELECT
    USING (app.is_member(lab_id));
CREATE POLICY check_results_insert ON check_results FOR INSERT
    WITH CHECK (app.is_member(lab_id));
CREATE POLICY check_results_update ON check_results FOR UPDATE
    USING (app.is_owner(lab_id)) WITH CHECK (app.is_owner(lab_id));
CREATE POLICY check_results_delete ON check_results FOR DELETE
    USING (app.is_owner(lab_id));

-- ---------------------------------------------------------------- owners
-- In a lab without an owner nobody could add a member, change a role, or delete
-- the lab again, except by hand with the migration role. So an UPDATE or DELETE
-- on lab_members that would leave a lab with no owner is refused, whoever runs
-- it, the migration role included.
--
-- The one exception is the cascade from DELETE FROM labs: it reaches the member
-- rows after the lab row is gone, and a lab that no longer exists needs no
-- owner. Deleting the account of a lab's only owner (app_users cascades to
-- lab_members) is refused like any other removal, so that owner has to hand the
-- lab on, or delete it, first.
--
-- The function is SECURITY DEFINER with a fixed search_path, like is_member, so
-- it runs as the role that ran this file. Row level security does not filter
-- that role, so the count sees every owner of the lab, whoever the caller is.
--
-- Two owners leaving at the same moment would each see the other one stay, and
-- the lab would be left with neither. The advisory lock, one per lab and held to
-- the end of the transaction, makes the second wait until the first has
-- committed; in READ COMMITTED, the default and what the application uses, its
-- next statement then sees the first one gone and refuses. The lock is taken
-- only while the lab still exists, so the cascade from DELETE FROM labs never
-- waits on it.
CREATE OR REPLACE FUNCTION app.keep_an_owner() RETURNS trigger
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, pg_temp AS $$
DECLARE
    change text := 'remove';
BEGIN
    -- The trigger fires only for an owner's row (its WHEN clause). An update
    -- that leaves that row an owner of the same lab takes no owner away.
    IF TG_OP = 'UPDATE' THEN
        IF NEW.role = 'owner' AND NEW.lab_id = OLD.lab_id THEN
            RETURN NEW;
        END IF;
        change := CASE WHEN NEW.lab_id = OLD.lab_id THEN 'demote' ELSE 'move' END;
    END IF;

    IF EXISTS (SELECT 1 FROM labs WHERE id = OLD.lab_id) THEN
        PERFORM pg_advisory_xact_lock(hashtextextended(OLD.lab_id::text, 0));
        IF NOT EXISTS (
            SELECT 1 FROM lab_members
            WHERE lab_id = OLD.lab_id AND role = 'owner' AND user_id <> OLD.user_id
        ) THEN
            RAISE EXCEPTION 'cannot % the last owner of lab %', change, OLD.lab_id
                USING ERRCODE = 'check_violation',
                      DETAIL = format('%s is the only owner of this lab.', OLD.user_id),
                      HINT = 'Make another member an owner first, or delete the lab.';
        END IF;
    END IF;

    IF TG_OP = 'DELETE' THEN
        RETURN OLD;
    END IF;
    RETURN NEW;
END
$$;

CREATE TRIGGER keep_an_owner
    BEFORE UPDATE OR DELETE ON lab_members
    FOR EACH ROW WHEN (OLD.role = 'owner')
    EXECUTE FUNCTION app.keep_an_owner();

COMMIT;
