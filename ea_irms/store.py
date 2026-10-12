"""Saving runs, and reading a lab's history back.

Identity is set in exactly one place, :func:`connect`. Every other function here
takes an open connection and writes plain SQL with no filtering of its own,
because the filtering is the database's job. If a query here forgot a condition
it would return nothing rather than someone else's data.
"""
from __future__ import annotations

import gzip
import json
from dataclasses import dataclass
from contextlib import contextmanager
from typing import Any, Iterable, Iterator

import psycopg
from psycopg.rows import dict_row


@dataclass
class Identity:
    user_id: str          # the OIDC subject claim
    email: str
    name: str | None = None


@contextmanager
def connect(dsn: str, identity: Identity | None) -> Iterator[psycopg.Connection]:
    """Open a connection that can see only what ``identity`` may see.

    Everything inside the ``with`` block is one transaction: it commits together
    on a clean exit and rolls back together on an exception, so a half-saved run
    cannot happen. The identity is set with ``set_config(..., true)``, which is
    local to that transaction, so it ends with the block and never survives on a
    connection a pooler hands to someone else.

    Passing None opens a connection with no identity, which every policy reads as
    no access. That is the safe direction for a bug to fail in.
    """
    with psycopg.connect(dsn, row_factory=dict_row) as conn:
        with conn.transaction():
            conn.execute("SELECT set_config('app.user_id', %s, true)",
                         (identity.user_id if identity else "",))
            yield conn


def register_user(admin_dsn: str, identity: Identity) -> None:
    """Record a user on first sign-in.

    This runs with the migration role rather than the application role, because a
    user cannot insert their own row before any policy can recognise them.
    """
    with psycopg.connect(admin_dsn, autocommit=True) as conn:
        conn.execute(
            "INSERT INTO app_users (id, email, name) VALUES (%s, %s, %s) "
            "ON CONFLICT (id) DO UPDATE SET email = EXCLUDED.email, name = EXCLUDED.name",
            (identity.user_id, identity.email, identity.name))


def create_lab(admin_dsn: str, identity: Identity, name: str,
               institution: str | None = None) -> str:
    """Create a lab and make the caller its owner, in one transaction.

    Also runs as the migration role: the caller is not yet a member of a lab that
    does not exist, so no policy could admit the insert.
    """
    with psycopg.connect(admin_dsn, row_factory=dict_row) as conn:
        with conn.transaction():
            lab = conn.execute(
                "INSERT INTO labs (name, institution) VALUES (%s, %s) RETURNING id",
                (name, institution)).fetchone()["id"]
            conn.execute("INSERT INTO lab_members (lab_id, user_id, role) "
                         "VALUES (%s, %s, 'owner')", (lab, identity.user_id))
    return str(lab)


def labs_for(conn) -> list[dict]:
    return conn.execute(
        "SELECT l.id, l.name, l.institution, m.role FROM labs l "
        "JOIN lab_members m ON m.lab_id = l.id ORDER BY l.name").fetchall()


def get_or_create_instrument(conn, lab_id: str, name: str,
                             model: str | None = None) -> str:
    row = conn.execute("SELECT id FROM instruments WHERE lab_id = %s AND name = %s",
                       (lab_id, name)).fetchone()
    if row:
        return str(row["id"])
    row = conn.execute(
        "INSERT INTO instruments (lab_id, name, model) VALUES (%s, %s, %s) RETURNING id",
        (lab_id, name, model)).fetchone()
    return str(row["id"])


def _fit(fits: dict, *path, default=None):
    """Dig a value out of the fits structure, tolerating a missing branch."""
    node: Any = fits
    for key in path:
        if not isinstance(node, dict) or node.get(key) is None:
            return default
        node = node[key]
    return node


RUN_COLUMNS = {
    "bgd_floor":          ("settling", "floor"),
    "bgd_excess":         ("settling", "excess"),
    "settled_from":       ("settling", "settled_from"),
    "settle_slope_early": ("settling", "early", "slope"),
    "settle_slope_late":  ("settling", "late", "slope"),
    "blank_n_n":          ("blank_n", "n"),
    "blank_area_n":       ("blank_n", "area"),
    "blank_d15n":         ("blank_n", "d"),
    "blank_sd_n":         ("blank_n", "sd"),
    "blank_in_force_n":   ("blank_n", "in_force"),
    "lin_slope_n":        ("lin_n", "fit", "slope"),
    "lin_r2_n":           ("lin_n", "fit", "r2"),
    "lin_n_n":            ("lin_n", "n"),
    "lin_slope_c":        ("lin_c", "fit", "slope"),
    "lin_r2_c":           ("lin_c", "fit", "r2"),
    "lin_n_c":            ("lin_c", "n"),
    "drift_model_n":      ("drift_n", "model"),
    "drift_slope_row_n":  ("drift_n", "row", "slope"),
    "drift_sd_row_n":     ("drift_n", "row", "sd"),
    "drift_slope_bgd_n":  ("drift_n", "background", "slope"),
    "drift_sd_bgd_n":     ("drift_n", "background", "sd"),
    "drift_model_c":      ("drift_c", "model"),
    "drift_slope_row_c":  ("drift_c", "row", "slope"),
    "drift_sd_row_c":     ("drift_c", "row", "sd"),
    "cal_slope_n":        ("cal_n", "slope"),
    "cal_int_n":          ("cal_n", "intercept"),
    "cal_r2_n":           ("cal_n", "r2"),
    "cal_sd_n":           ("cal_n", "sd"),
    "cal_n_n":            ("cal_n", "n"),
    "cal_slope_c":        ("cal_c", "slope"),
    "cal_int_c":          ("cal_c", "intercept"),
    "cal_r2_c":           ("cal_c", "r2"),
    "cal_sd_c":           ("cal_c", "sd"),
    "cal_n_c":            ("cal_c", "n"),
    "kfac_slope_n":       ("k_n", "slope"),
    "kfac_int_n":         ("k_n", "intercept"),
    "kfac_r2_n":          ("k_n", "r2"),
    "kfac_slope_c":       ("k_c", "slope"),
    "kfac_int_c":         ("k_c", "intercept"),
    "kfac_r2_c":          ("k_c", "r2"),
    # 1.3.0: decisions, reference pulses, repeatability in use
    "lin_p_n":            ("lin_n", "fit", "p"),
    "lin_in_force_n":     ("lin_n", "in_force"),
    "lin_p_c":            ("lin_c", "fit", "p"),
    "lin_in_force_c":     ("lin_c", "in_force"),
    "drift_p_n":          ("drift_n", "fit", "p"),
    "drift_in_force_n":   ("drift_n", "in_force"),
    "drift_p_c":          ("drift_c", "fit", "p"),
    "drift_in_force_c":   ("drift_c", "in_force"),
    "ref_pulse_n_n":      ("ref_gas_n", "n"),
    "ref_pulse_sd_n":     ("ref_gas_n", "sd"),
    "ref_pulse_trend_p_n": ("ref_gas_n", "fit", "p"),
    "ref_pulse_n_c":      ("ref_gas_c", "n"),
    "ref_pulse_sd_c":     ("ref_gas_c", "sd"),
    "ref_pulse_trend_p_c": ("ref_gas_c", "fit", "p"),
    "n_checks":           ("checks", "n"),
    "u_sample_n":         ("unc_n", "u_sample"),
    "u_sample_c":         ("unc_c", "u_sample"),
}

POSITION_COLUMNS = [
    ("position", "row"), ("identifier", "identifier"), ("role", "role"),
    ("amount_mg", "amount"), ("area_n", "area_n"), ("ampl_n", "ampl_n"),
    ("bgd_n", "bgd_n"), ("ref_bgd_n", "ref_bgd_n"), ("d15n_raw", "d15n"),
    ("d15n_final", "d15n_final"), ("area_c", "area_c"), ("ampl_c", "ampl_c"),
    ("d13c_raw", "d13c"), ("d13c_final", "d13c_final"),
    # 1.4.2: the reported elemental content carries the elemental drift correction
    # when it is in force, and equals the measured value otherwise
    ("pct_n", "pct_n_reported"), ("pct_c", "pct_c_reported"),
    ("cn_molar", "cn_molar_reported"), ("cn_mass", "cn_mass_reported"),
    ("u_d15n", "u_d15n"), ("u_d13c", "u_d13c"),
]


def _stage_sds(positions) -> dict:
    """Scatter of the drift standards after each correction, the honest diagnostic."""
    import statistics
    out = {}
    for column, attribute in (("sd_raw", "d15n"), ("sd_after_blank", "d15n_blank"),
                              ("sd_after_lin", "d15n_lin"), ("sd_after_drift", "d15n_drift"),
                              ("sd_after_cal", "d15n_final")):
        values = [getattr(p, attribute, None) for p in positions
                  if getattr(p, "role", None) == "Drift"
                  and "excluded from N drift fit" not in (getattr(p, "flags", None) or [])]
        values = [v for v in values if v is not None]
        out[column] = statistics.stdev(values) if len(values) > 1 else None
    return out


def save_run(conn, *, lab_id: str, instrument_id: str, identity: Identity,
             positions: Iterable, fits: dict, config: dict, tool_version: str,
             run_name: str | None = None, run_date=None,
             raw_bytes: bytes | None = None, raw_filename: str | None = None) -> str:
    """Store one sequence: the run's fits, every position, and the check standards.

    The raw export is kept gzipped, at roughly a sixth of its size, and the
    generated workbook is not kept at all because it can be rebuilt from the raw
    export, the configuration and the tool version recorded here.
    """
    positions = list(positions)
    values = {name: _fit(fits, *path) for name, path in RUN_COLUMNS.items()}
    values.update(_stage_sds(positions))
    values.update(
        lab_id=lab_id, instrument_id=instrument_id, uploaded_by=identity.user_id,
        run_name=run_name, run_date=run_date, tool_version=tool_version,
        config=json.dumps(config),
        raw_export=gzip.compress(raw_bytes, 9) if raw_bytes else None,
        raw_filename=raw_filename,
        n_positions=len(positions),
        n_flagged=sum(1 for p in positions if getattr(p, "flags", None)),
    )
    columns = list(values)
    with conn.transaction():
        run_id = conn.execute(
            f"INSERT INTO runs ({', '.join(columns)}) "
            f"VALUES ({', '.join('%s' for _ in columns)}) RETURNING id",
            [values[c] for c in columns]).fetchone()["id"]

        names = ["lab_id", "run_id"] + [c for c, _ in POSITION_COLUMNS] + ["flags"]
        rows = [[lab_id, run_id]
                + [getattr(p, attr, None) for _, attr in POSITION_COLUMNS]
                + ["; ".join(getattr(p, "flags", []) or []) or None]
                for p in positions]
        with conn.cursor() as cur:
            cur.executemany(
                f"INSERT INTO positions ({', '.join(names)}) "
                f"VALUES ({', '.join('%s' for _ in names)})", rows)

            checks = [p for p in positions if getattr(p, "role", None) == "Check"]
            if checks:
                cur.executemany(
                    "INSERT INTO check_results (lab_id, run_id, material, n, "
                    "measured_d15n, accepted_d15n, offset_d15n, "
                    "measured_d13c, accepted_d13c, offset_d13c) "
                    "VALUES (%s, %s, %s, 1, %s, %s, %s, %s, %s, %s)",
                    [[lab_id, run_id, p.identifier,
                      p.d15n_final, p.acc15,
                      None if p.d15n_final is None or p.acc15 is None else p.d15n_final - p.acc15,
                      p.d13c_final, p.acc13,
                      None if p.d13c_final is None or p.acc13 is None else p.d13c_final - p.acc13]
                     for p in checks])
    return str(run_id)


def lab_history(conn, lab_id: str, instrument_id: str | None = None) -> list[dict]:
    """Every run's QC numbers, oldest first, which is what the trend plots read."""
    sql = ("SELECT r.*, i.name AS instrument FROM runs r "
           "JOIN instruments i ON i.id = r.instrument_id WHERE r.lab_id = %s")
    params: list[Any] = [lab_id]
    if instrument_id:
        sql += " AND r.instrument_id = %s"
        params.append(instrument_id)
    return conn.execute(sql + " ORDER BY COALESCE(r.run_date, r.uploaded_at::date), "
                              "r.uploaded_at", params).fetchall()


def check_history(conn, lab_id: str, material: str | None = None) -> list[dict]:
    """Every check-standard result the lab has saved, oldest first.

    This is the series a control chart is drawn from (guide section 7.1). The
    run date stands in for the day the limits count, and the instrument name
    lets the chart be drawn per instrument.
    """
    sql = ("SELECT c.material, c.measured_d15n, c.accepted_d15n, c.offset_d15n, "
           "c.measured_d13c, c.accepted_d13c, c.offset_d13c, "
           "COALESCE(r.run_date, r.uploaded_at::date) AS run_day, r.run_name, "
           "i.name AS instrument, r.id AS run_id "
           "FROM check_results c JOIN runs r ON r.id = c.run_id "
           "JOIN instruments i ON i.id = r.instrument_id WHERE c.lab_id = %s")
    params: list[Any] = [lab_id]
    if material:
        sql += " AND c.material = %s"
        params.append(material)
    return conn.execute(sql + " ORDER BY run_day, r.uploaded_at, c.id", params).fetchall()


def run_positions(conn, run_id: str) -> list[dict]:
    return conn.execute("SELECT * FROM positions WHERE run_id = %s ORDER BY position",
                        (run_id,)).fetchall()


def raw_export(conn, run_id: str) -> tuple[str | None, bytes | None]:
    """The original Isodat file, so any run can be corrected again later."""
    row = conn.execute("SELECT raw_filename, raw_export FROM runs WHERE id = %s",
                       (run_id,)).fetchone()
    if not row or row["raw_export"] is None:
        return (row["raw_filename"] if row else None), None
    return row["raw_filename"], gzip.decompress(bytes(row["raw_export"]))


def export_lab(conn, lab_id: str) -> dict:
    """Everything a lab owns, for the download-my-data button."""
    return {
        "lab": conn.execute("SELECT * FROM labs WHERE id = %s", (lab_id,)).fetchone(),
        "instruments": conn.execute("SELECT * FROM instruments WHERE lab_id = %s",
                                    (lab_id,)).fetchall(),
        "runs": conn.execute("SELECT * FROM runs WHERE lab_id = %s", (lab_id,)).fetchall(),
        "positions": conn.execute("SELECT * FROM positions WHERE lab_id = %s",
                                  (lab_id,)).fetchall(),
        "check_results": conn.execute("SELECT * FROM check_results WHERE lab_id = %s",
                                      (lab_id,)).fetchall(),
    }


def delete_lab_data(conn, lab_id: str) -> None:
    """The delete-my-data button. Cascades take the children."""
    with conn.transaction():
        conn.execute("DELETE FROM runs WHERE lab_id = %s", (lab_id,))
