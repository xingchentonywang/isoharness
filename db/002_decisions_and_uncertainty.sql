-- ea_irms 1.3.0: what each run decided about its corrections, the reference
-- pulses, and the standard uncertainty of every position. Columns only; the
-- row-level security of 001 covers them unchanged.
ALTER TABLE runs
    ADD COLUMN lin_p_n            double precision,
    ADD COLUMN lin_in_force_n     boolean,
    ADD COLUMN lin_p_c            double precision,
    ADD COLUMN lin_in_force_c     boolean,
    ADD COLUMN drift_p_n          double precision,
    ADD COLUMN drift_in_force_n   boolean,
    ADD COLUMN drift_p_c          double precision,
    ADD COLUMN drift_in_force_c   boolean,
    ADD COLUMN ref_pulse_n_n      integer,
    ADD COLUMN ref_pulse_sd_n     double precision,
    ADD COLUMN ref_pulse_trend_p_n double precision,
    ADD COLUMN ref_pulse_n_c      integer,
    ADD COLUMN ref_pulse_sd_c     double precision,
    ADD COLUMN ref_pulse_trend_p_c double precision,
    ADD COLUMN n_checks           integer,
    ADD COLUMN u_sample_n         double precision,
    ADD COLUMN u_sample_c         double precision;

ALTER TABLE positions
    ADD COLUMN u_d15n double precision,
    ADD COLUMN u_d13c double precision;
