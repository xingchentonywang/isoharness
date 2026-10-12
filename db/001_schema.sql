-- ea_irms schema, with isolation enforced by Postgres rather than by application code.
--
-- Every table carries lab_id and every policy is the same sentence: you may see a
-- row if you are a member of the lab that owns it. Denormalising lab_id onto the
-- child tables costs sixteen bytes a row and makes the security auditable by
-- reading eight identical policies instead of following joins.
--
-- Run once against a new Neon project:
--     psql "$DATABASE_URL" -f db/001_schema.sql

BEGIN;

CREATE SCHEMA IF NOT EXISTS app;
CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- ---------------------------------------------------------------- identity
-- One place decides who the caller is. Locally and on plain Postgres this reads
-- a session setting that the connection helper sets once per checkout. On Neon
-- RLS, replace the body with `auth.user_id()` from pg_session_jwt and nothing
-- else in this file changes.
CREATE OR REPLACE FUNCTION app.current_user_id() RETURNS text
LANGUAGE sql STABLE AS $$
    SELECT nullif(current_setting('app.user_id', true), '')
$$;

CREATE TABLE app_users (
    id          text PRIMARY KEY,                     -- the OIDC subject claim
    email       text NOT NULL,
    name        text,
    created_at  timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE labs (
    id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    name        text NOT NULL,
    institution text,
    created_at  timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE lab_members (
    lab_id   uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    user_id  text NOT NULL REFERENCES app_users(id) ON DELETE CASCADE,
    role     text NOT NULL DEFAULT 'member' CHECK (role IN ('owner', 'member')),
    added_at timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (lab_id, user_id)
);
CREATE INDEX lab_members_user_idx ON lab_members(user_id);

-- Membership is asked on every query, so it is a SECURITY DEFINER function to
-- avoid recursing into lab_members' own policy.
CREATE OR REPLACE FUNCTION app.is_member(target uuid) RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER SET search_path = public, pg_temp AS $$
    SELECT EXISTS (
        SELECT 1 FROM lab_members
        WHERE lab_id = target AND user_id = app.current_user_id()
    )
$$;

-- ---------------------------------------------------------------- the science
CREATE TABLE instruments (
    id       uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    lab_id   uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    name     text NOT NULL,
    model    text,
    notes    text,
    UNIQUE (lab_id, name)
);

CREATE TABLE lab_configs (
    id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    lab_id      uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    name        text NOT NULL,
    config      jsonb NOT NULL,
    is_default  boolean NOT NULL DEFAULT false,
    updated_at  timestamptz NOT NULL DEFAULT now(),
    UNIQUE (lab_id, name)
);

CREATE TABLE runs (
    id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    lab_id        uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    instrument_id uuid NOT NULL REFERENCES instruments(id) ON DELETE CASCADE,
    uploaded_by   text REFERENCES app_users(id) ON DELETE SET NULL,
    uploaded_at   timestamptz NOT NULL DEFAULT now(),
    run_name      text,
    run_date      date,
    tool_version  text NOT NULL,
    config        jsonb NOT NULL,          -- exactly what produced these numbers
    raw_export    bytea,                   -- the Isodat file, gzipped
    raw_filename  text,
    n_positions   integer,
    n_flagged     integer,
    -- settling
    bgd_floor        double precision,
    bgd_excess       double precision,
    settled_from     integer,
    settle_slope_early double precision,
    settle_slope_late  double precision,
    -- blank
    blank_n_n     integer, blank_area_n double precision,
    blank_d15n    double precision, blank_sd_n double precision,
    blank_in_force_n boolean,
    -- linearity
    lin_slope_n double precision, lin_r2_n double precision, lin_n_n integer,
    lin_slope_c double precision, lin_r2_c double precision, lin_n_c integer,
    -- drift
    drift_model_n    text,
    drift_slope_row_n double precision, drift_sd_row_n double precision,
    drift_slope_bgd_n double precision, drift_sd_bgd_n double precision,
    drift_model_c    text,
    drift_slope_row_c double precision, drift_sd_row_c double precision,
    -- calibration
    cal_slope_n double precision, cal_int_n double precision,
    cal_r2_n double precision, cal_sd_n double precision, cal_n_n integer,
    cal_slope_c double precision, cal_int_c double precision,
    cal_r2_c double precision, cal_sd_c double precision, cal_n_c integer,
    -- elemental k-factors
    kfac_slope_n double precision, kfac_int_n double precision, kfac_r2_n double precision,
    kfac_slope_c double precision, kfac_int_c double precision, kfac_r2_c double precision,
    -- precision at each stage, drift standards
    sd_raw double precision, sd_after_blank double precision,
    sd_after_lin double precision, sd_after_drift double precision,
    sd_after_cal double precision
);
CREATE INDEX runs_lab_date_idx ON runs(lab_id, run_date DESC);
CREATE INDEX runs_instrument_idx ON runs(instrument_id, uploaded_at DESC);

CREATE TABLE positions (
    id          bigserial PRIMARY KEY,
    lab_id      uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    run_id      uuid NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    position    integer NOT NULL,
    identifier  text,
    role        text,
    amount_mg   double precision,
    area_n double precision, ampl_n double precision,
    bgd_n double precision, ref_bgd_n double precision,
    d15n_raw double precision, d15n_final double precision,
    area_c double precision, ampl_c double precision,
    d13c_raw double precision, d13c_final double precision,
    pct_n double precision, pct_c double precision,
    cn_molar double precision, cn_mass double precision,
    flags text,
    UNIQUE (run_id, position)
);
CREATE INDEX positions_run_idx ON positions(run_id);

CREATE TABLE check_results (
    id            bigserial PRIMARY KEY,
    lab_id        uuid NOT NULL REFERENCES labs(id) ON DELETE CASCADE,
    run_id        uuid NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    material      text NOT NULL,
    n             integer,
    measured_d15n double precision, accepted_d15n double precision, offset_d15n double precision,
    measured_d13c double precision, accepted_d13c double precision, offset_d13c double precision
);
CREATE INDEX check_results_run_idx ON check_results(run_id);

-- ---------------------------------------------------------------- isolation
-- FORCE is what makes this real. Without it the table owner silently bypasses
-- every policy, and a migration run as the owner would look like it worked.
DO $$
DECLARE t text;
BEGIN
    FOREACH t IN ARRAY ARRAY['labs','lab_members','instruments','lab_configs',
                             'runs','positions','check_results'] LOOP
        EXECUTE format('ALTER TABLE %I ENABLE ROW LEVEL SECURITY', t);
        EXECUTE format('ALTER TABLE %I FORCE ROW LEVEL SECURITY', t);
    END LOOP;
END $$;

CREATE POLICY lab_access ON labs
    USING (app.is_member(id)) WITH CHECK (app.is_member(id));
CREATE POLICY lab_members_access ON lab_members
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY instruments_access ON instruments
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY lab_configs_access ON lab_configs
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY runs_access ON runs
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY positions_access ON positions
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));
CREATE POLICY check_results_access ON check_results
    USING (app.is_member(lab_id)) WITH CHECK (app.is_member(lab_id));

-- A user reads and writes only their own row in app_users.
ALTER TABLE app_users ENABLE ROW LEVEL SECURITY;
ALTER TABLE app_users FORCE ROW LEVEL SECURITY;
CREATE POLICY app_users_self ON app_users
    USING (id = app.current_user_id()) WITH CHECK (id = app.current_user_id());

-- ---------------------------------------------------------------- app role
-- The application never connects as the owner, so it can never bypass a policy.
DO $$ BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'ea_app') THEN
        CREATE ROLE ea_app NOLOGIN NOBYPASSRLS;
    END IF;
END $$;
GRANT USAGE ON SCHEMA public, app TO ea_app;
GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA public TO ea_app;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO ea_app;
GRANT EXECUTE ON FUNCTION app.current_user_id(), app.is_member(uuid) TO ea_app;

COMMIT;
