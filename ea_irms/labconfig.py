"""A lab configuration as an Excel file (ea_irms 1.4.6).

``config_to_xlsx`` writes it: a README sheet, a Settings sheet with one row per
setting, and a Standards sheet in the column order of the workbook's Setup table.
``config_from_xlsx`` reads one back and returns the configuration, or raises
``LabConfigError`` naming every cell it could not read. JSON configurations
(``config_to_dict`` and ``config_from_dict``) work as before.

A row whose key is missing (a file written by an older version) takes the
package default. An empty cell means "off" or "not given" where the setting has
such a state, as on the workbook's Setup sheet; where it has none, an empty cell
is an error rather than a silent default.
"""
from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .pipeline import (MAX_STANDARDS, STANDARD_FIELDS, Config, config_from_dict, config_to_dict,
                       standards_from_rows)

SWITCH = ("On", "Off", "Auto")
ONOFF = ("On", "Off")
MODELS = ("Row", "Background", "Auto")
ROLES = ("Anchor", "Check", "Drift", "Linearity", "Blank", "Exclude", "Sample")
STANDARD_HEADERS = ["Match text in Identifier 1", "Role", "Accepted d13C", "Accepted d15N", "%C", "%N",
                    "Material", "1σ(d13C)", "1σ(d15N)", "1σ(%C)", "1σ(%N)"]
SETTINGS_HEADERS = ["Setting", "Value", "What it means", "Key"]

# (key, label, kind, note). A row without a key is a heading. Kinds: switch (On, Off or
# Auto), switch_or_empty, onoff (On or Off, 1.4.8), model, number, number_or_empty,
# whole_or_empty, positions.
SETTINGS = [
    (None, "Corrections", None, None),
    ("blank_correction", "Blank correction", "switch",
     "On, Off or Auto. Auto applies it only when at least three blanks agree to within 2 permil."),
    ("linearity_correction", "Linearity correction", "switch",
     "On, Off or Auto. Auto applies the fitted slope only when it differs from zero at the significance "
     "level below. Sets nitrogen, and carbon too unless the next row is filled."),
    ("linearity_correction_c", "Linearity correction, carbon only", "switch_or_empty",
     "Empty: as the row above. On, Off or Auto, for runs where nitrogen and carbon need different answers."),
    ("drift_correction", "Drift correction", "switch",
     "On, Off or Auto. Auto applies the fitted slope only when it differs from zero at the significance "
     "level below."),
    ("drift_model", "Drift model", "model",
     "Row, Background or Auto. Auto takes the lower residual scatter, which is not an independent test."),
    ("elemental_correction", "Elemental drift correction", "switch",
     "On, Off or Auto. Divides %N and %C by the factor the drift monitors show at each position."),
    ("elemental_threshold_pct", "Elemental correction on Auto when a monitor is off by more than (percent)",
     "number", "The largest bias of the monitors used, in percent of the certified value."),
    ("significance_level", "Significance level for Auto", "number",
     "A slope is applied on Auto when its two-sided p-value is below this."),
    (None, "Flag thresholds", None, None),
    ("check_offset_permil", "Check standard offset (permil)", "number",
     "Flag a check standard further than this from its accepted value."),
    ("min_area_n", "Minimum Area 28 (Vs)", "number_or_empty",
     "Empty: the smallest linearity standard in the run."),
    ("min_area_c", "Minimum Area 44 (Vs)", "number_or_empty",
     "Empty: the smallest linearity standard in the run."),
    ("max_area_n", "Maximum Area 28 (Vs)", "number_or_empty",
     "Empty: the largest linearity standard in the run."),
    ("max_area_c", "Maximum Area 44 (Vs)", "number_or_empty",
     "Empty: the largest linearity standard in the run."),
    ("min_blank_multiple", "Minimum peak area, x the blank", "number_or_empty",
     "A peak under this multiple of the blank is flagged, since there the blank correction carries the blank's "
     "scatter into the value. Empty: no check."),
    ("settled_threshold", "Settled when background excess is under (fraction)", "number_or_empty",
     "Fraction of the run's starting excess left in the reference-peak background. Empty: no check."),
    ("monitor_outlier_t", "Drift monitor outlier, studentized residual", "number",
     "A drift monitor further than this from the line through the others, in units of their scatter, is "
     "flagged."),
    ("carryover_mv", "Carryover: background rise after an oversized sample (mV)", "number_or_empty",
     "Flags a position when one of the two before it is above the linearity range and its background rose "
     "by more than this. Empty: no check."),
    (None, "Reference pulses", None, None),
    ("ref_pulse_ampl_tol", "Reference-pulse check, amplitude tolerance (fraction)", "number_or_empty",
     "Used only in an export without N2 reference peaks, where an unflagged CO2 peak this close to the "
     "designated pulses is set aside as a reference pulse. Empty: no check."),
    ("ref_pulse_d13c_window", "Reference-pulse check, d13C window (permil)", "number",
     "Around the designated pulses' mean d13C."),
    ("ref_pulse_off_sd", "Reference pulse off at a position, robust SDs from the run's median",
     "number_or_empty", "Marks that sample for a rerun of that element. Empty: no check."),
    ("ref_gas_sd_limit", "Reference-gas SD that is a problem for the whole run (permil)", "number_or_empty",
     "The FIRMS guide gives 0.1 permil for its zero-enrichment test (section 4.3.2). Empty: no check."),
    (None, "Rerun marks (1.4.7 and 1.4.8)", None, None),
    ("max_amplitude_mv", "Largest peak amplitude the instrument measures (mV)", "number_or_empty",
     "A peak above the linearity range is marked for a rerun only at or above this amplitude, and a peak at or "
     "above it is marked whatever the range. Empty: every peak above the linearity range is marked, as before "
     "1.4.7."),
    ("rerun_max_1sigma", "Largest 1σ of the final d15N or d13C before a rerun (permil)", "number_or_empty",
     "A sample whose final d15N or d13C has a 1σ above this is marked for a rerun of that element. The 1σ leaves "
     "out the blank, linearity, and drift corrections. Empty: no check."),
    ("rerun_replicates", "Compare replicates of a sample in one run", "onoff",
     "On or Off. With a sample sheet, analyses of one sample in a run (the same name, or the same Replicate of) "
     "whose range is above the critical range of ISO 5725-6 are marked for a rerun of that element (1.4.8)."),
    (None, "Uncertainty", None, None),
    ("u_sample_n", "Sample repeatability, d15N (1σ, permil)", "number_or_empty",
     "Empty: the pooled scatter of the anchor replicates in each run."),
    ("u_sample_c", "Sample repeatability, d13C (1σ, permil)", "number_or_empty",
     "Empty: the pooled scatter of the anchor replicates in each run."),
    ("coverage_k", "Multiple of 1σ for the checks (k)", "number",
     "The checks compare a standard's deviation with k x 1σ. k = 2 covers about 95 percent when the "
     "replicates behind 1σ are many."),
    ("balance_u_mg", "Balance, 1σ of one weighing (mg)", "number_or_empty",
     "Empty: the weighing term is left out of the uncertainty of %N and %C."),
    (None, "One run only (leave these empty in a lab configuration)", None, None),
    ("first_drift_position", "First position to use in the drift fit", "whole_or_empty",
     "Empty: every position."),
    ("exclude_drift_n", "Positions left out of the nitrogen drift fit", "positions",
     "Comma-separated autosampler positions."),
    ("exclude_drift_c", "Positions left out of the carbon drift fit", "positions",
     "Comma-separated autosampler positions."),
    ("exclude_elemental", "Positions left out of the elemental correction", "positions",
     "Comma-separated autosampler positions."),
]
KEYS = [s[0] for s in SETTINGS if s[0]]


class LabConfigError(ValueError):
    """The file is not a lab configuration this version reads; ``problems`` names each cell."""

    def __init__(self, problems):
        self.problems = list(problems)
        super().__init__("; ".join(self.problems))


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "On" if value else "Off"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _cell_value(key, value):
    """The configuration's value as it is written into the Value cell."""
    if key in ("exclude_drift_n", "exclude_drift_c", "exclude_elemental"):
        return ", ".join(str(v) for v in value or []) or None
    return value


def config_to_xlsx(config: Config | None = None, path: str | Path | None = None, *,
                   title: str = "Lab configuration") -> bytes:
    """The configuration as an Excel file; also written to ``path`` when one is given."""
    from .workbook import VERSION
    config = config or Config()
    data = config_to_dict(config)
    arial = "Arial"
    f_title = Font(name=arial, size=14, bold=True)
    f_h1 = Font(name=arial, size=11, bold=True)
    f_hdr = Font(name=arial, size=9, bold=True, color="FFFFFF")
    f_sm = Font(name=arial, size=9)
    f_in = Font(name=arial, size=9, color="0000FF")
    f_note = Font(name=arial, size=9, italic=True, color="595959")
    f_key = Font(name=arial, size=8, color="808080")
    fill_hdr = PatternFill("solid", fgColor="1F4E79")
    fill_grp = PatternFill("solid", fgColor="DDEBF7")
    fill_in = PatternFill("solid", fgColor="FFF2CC")
    fill_key = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 100
    lines = [
        (title, f_title),
        (f"ea_irms {VERSION}. A lab configuration for isoharness.org and for the ea_irms workbook.", f_note),
        ("", f_sm),
        ("How to use it", f_h1),
        ("1. Fill in the yellow cells on Settings and Standards. Leave every other cell as it is, including "
         "the gray Key column, which is how the file is read.", f_sm),
        ("2. On isoharness.org, choose this file as the lab configuration before or after you upload an "
         "export. The same file can be used for every run of the lab.", f_sm),
        ("3. An empty cell means off or not given where the setting allows that, as the notes say. "
         "Settings that need a value say so when the file is read.", f_sm),
        ("", f_sm),
        ("Standards", f_h1),
        ("Each row of Standards matches the identifiers of the export that contain its match text, and gives "
         "them a role and, for anchors, check standards, drift monitors, and linearity standards, their "
         "accepted values. When two rows match the same identifier, the lower row wins, so put the more "
         f"specific text lower down. The table holds at most {MAX_STANDARDS} rows. The accepted values that "
         "come filled in are the published USGS values as of 2026-09-16; check them against your "
         "certificates. 1σ is the uncertainty the producer lists, one standard deviation.", f_sm),
        ("", f_sm),
        ("One run only", f_h1),
        ("The last rows of Settings (the first drift position and the positions left out) belong to one run. "
         "A configuration downloaded after a run carries them; leave them empty in a configuration meant for "
         "every run of the lab.", f_sm),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(i, 2, text)
        c.font = font
        c.alignment = wrap

    st = wb.create_sheet("Settings")
    st.sheet_view.showGridLines = False
    for col, width in zip("ABCD", (66, 14, 76, 24)):
        st.column_dimensions[col].width = width
    st.cell(1, 1, "Settings").font = f_title
    st.cell(2, 1, "Yellow cells are yours. The Key column is how the file is read; do not change it.").font = f_note
    for j, h in enumerate(SETTINGS_HEADERS, start=1):
        c = st.cell(4, j, h)
        c.font = f_hdr
        c.fill = fill_hdr
    dv_sw = DataValidation(type="list", formula1='"On,Off,Auto"', allow_blank=False)
    dv_swe = DataValidation(type="list", formula1='"On,Off,Auto"', allow_blank=True)
    dv_dm = DataValidation(type="list", formula1='"Row,Background,Auto"', allow_blank=False)
    dv_oo = DataValidation(type="list", formula1='"On,Off"', allow_blank=False)
    for dv in (dv_sw, dv_swe, dv_dm, dv_oo):
        st.add_data_validation(dv)
    r = 5
    for key, label, kind, note in SETTINGS:
        if key is None:
            c = st.cell(r, 1, label)
            c.font = f_h1
            for j in range(1, 5):
                st.cell(r, j).fill = fill_grp
            r += 1
            continue
        st.cell(r, 1, label).font = f_sm
        v = st.cell(r, 2, _cell_value(key, data[key]))
        v.font = f_in
        v.fill = fill_in
        v.border = box
        if kind == "positions":
            v.number_format = "@"
        n = st.cell(r, 3, note)
        n.font = f_note
        n.alignment = Alignment(wrap_text=True, vertical="top")
        k = st.cell(r, 4, key)
        k.font = f_key
        k.fill = fill_key
        dv = {"switch": dv_sw, "switch_or_empty": dv_swe, "model": dv_dm, "onoff": dv_oo}.get(kind)
        if dv is not None:
            dv.add(f"B{r}")
        r += 1
    st.freeze_panes = "A5"

    sd = wb.create_sheet("Standards")
    sd.sheet_view.showGridLines = False
    for j, width in enumerate((26, 11, 11, 11, 8, 8, 18, 9, 9, 9, 9), start=1):
        sd.column_dimensions[chr(64 + j)].width = width
    sd.cell(1, 1, "Standards").font = f_title
    sd.cell(2, 1, f"Yellow cells are yours, at most {MAX_STANDARDS} rows. Roles: {', '.join(ROLES)}. "
                  "The lower of two matching rows wins.").font = f_note
    for j, h in enumerate(STANDARD_HEADERS, start=1):
        c = sd.cell(4, j, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    rows = [list(row) + [None] * (STANDARD_FIELDS - len(row)) for row in data["standards"]]
    for i in range(MAX_STANDARDS):
        for j in range(1, STANDARD_FIELDS + 1):
            c = sd.cell(5 + i, j)
            c.font = f_in
            c.fill = fill_in
            c.border = box
            if j in (3, 4, 5, 6, 8, 9, 10, 11):
                c.number_format = "0.00"
            if i < len(rows) and rows[i][j - 1] not in (None, ""):
                c.value = rows[i][j - 1]
    dv_role = DataValidation(type="list", formula1='"' + ",".join(ROLES) + '"', allow_blank=True)
    sd.add_data_validation(dv_role)
    dv_role.add(f"B5:B{4 + MAX_STANDARDS}")
    sd.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    blob = buf.getvalue()
    if path is not None:
        Path(path).write_bytes(blob)
    return blob


def _number(text, whole=False):
    value = float(text) if not isinstance(text, (int, float)) else text
    if whole:
        if float(value) != int(value):
            raise ValueError
        return int(value)
    return float(value)


def _read_value(key, kind, raw, problems, where):
    empty = raw is None or (isinstance(raw, str) and not raw.strip())
    text = raw.strip() if isinstance(raw, str) else raw
    if kind in ("switch", "switch_or_empty", "model", "onoff"):
        allowed = MODELS if kind == "model" else ONOFF if kind == "onoff" else SWITCH
        if empty:
            if kind == "switch_or_empty":
                return None
            problems.append(f"{where} ({key}) is empty; it needs one of {', '.join(allowed)}")
            return None
        if isinstance(text, bool):
            text = "On" if text else "Off"
        hit = next((a for a in allowed if str(text).lower() == a.lower()), None)
        if hit is None:
            problems.append(f"{where} ({key}) is {_text(raw)!r}; it needs one of {', '.join(allowed)}")
        return hit
    if kind == "positions":
        if empty:
            return []
        if isinstance(text, (int, float)) and not isinstance(text, bool):
            try:
                return [_number(text, whole=True)]
            except ValueError:
                problems.append(f"{where} ({key}) is {_text(raw)!r}; it needs whole positions such as 59, 61")
                return []
        tokens = [t.strip() for t in str(text).replace(";", ",").split(",") if t.strip()]
        if not all(t.isdigit() for t in tokens):
            problems.append(f"{where} ({key}) is {_text(raw)!r}; it needs whole positions such as 59, 61")
            return []
        return [int(t) for t in tokens]
    if empty:
        if kind in ("number_or_empty", "whole_or_empty"):
            return None
        problems.append(f"{where} ({key}) is empty; it needs a number")
        return None
    try:
        return _number(text, whole=(kind == "whole_or_empty"))
    except (TypeError, ValueError):
        what = "a whole number" if kind == "whole_or_empty" else "a number"
        problems.append(f"{where} ({key}) is {_text(raw)!r}; it needs {what}")
        return None


def read_lab_config(source) -> tuple[Config, list[str]]:
    """Read an Excel lab configuration. Returns the configuration and notes on what was
    filled from the defaults or left unread; raises LabConfigError on any cell it cannot use.
    ``source`` is a path or the file's bytes."""
    try:
        wb = load_workbook(io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source,
                           data_only=True)
    except Exception as exc:   # not an .xlsx file at all
        raise LabConfigError([f"the file could not be opened as an Excel workbook ({type(exc).__name__})"]) \
            from exc
    problems, notes = [], []
    if "Settings" not in wb.sheetnames or "Standards" not in wb.sheetnames:
        raise LabConfigError(["a lab configuration has a Settings sheet and a Standards sheet; this file has "
                              + ", ".join(wb.sheetnames)])
    st, sd = wb["Settings"], wb["Standards"]
    kinds = {key: kind for key, _, kind, _ in SETTINGS if key}
    header = None
    for r in range(1, min(st.max_row, 20) + 1):
        if [st.cell(r, j).value for j in (1, 2, 4)] == ["Setting", "Value", "Key"]:
            header = r
            break
    if header is None:
        raise LabConfigError(["the Settings sheet has no header row reading Setting, Value, What it means, Key"])
    data, seen = {}, set()
    for r in range(header + 1, st.max_row + 1):
        key = st.cell(r, 4).value
        if key is None or (isinstance(key, str) and not key.strip()):
            continue
        key = str(key).strip()
        if key not in kinds:
            notes.append(f"Settings row {r}: the key {key!r} is not one this version reads, so it was left out")
            continue
        if key in seen:
            problems.append(f"Settings row {r}: {key} appears twice")
            continue
        seen.add(key)
        data[key] = _read_value(key, kinds[key], st.cell(r, 2).value, problems, f"Settings B{r}")
    missing = [k for k in KEYS if k not in seen]
    if missing:
        notes.append("not in the file, so the package defaults were used: " + ", ".join(missing))

    hdr_row = None
    for r in range(1, min(sd.max_row, 20) + 1):
        if str(sd.cell(r, 1).value or "").strip() == STANDARD_HEADERS[0]:
            hdr_row = r
            break
    if hdr_row is None:
        raise LabConfigError([f"the Standards sheet has no header row starting with {STANDARD_HEADERS[0]}"])
    rows = []
    for r in range(hdr_row + 1, sd.max_row + 1):
        cells = [sd.cell(r, j).value for j in range(1, STANDARD_FIELDS + 1)]
        if all(c is None or (isinstance(c, str) and not c.strip()) for c in cells):
            continue
        match = cells[0]
        if match is None or not str(match).strip():
            problems.append(f"Standards row {r} has values but no match text in column A")
            continue
        cells[0] = str(match).strip()
        role = cells[1]
        hit = next((x for x in ROLES if str(role or "").strip().lower() == x.lower()), None)
        if hit is None:
            problems.append(f"Standards B{r} ({cells[0]}) is {_text(role)!r}; the role is one of {', '.join(ROLES)}")
        cells[1] = hit
        for j in (2, 3, 4, 5, 7, 8, 9, 10):
            v = cells[j]
            if v is None or (isinstance(v, str) and not v.strip()):
                cells[j] = None
                continue
            try:
                cells[j] = _number(v)
            except (TypeError, ValueError):
                problems.append(f"Standards {chr(65 + j)}{r} ({cells[0]}) is {_text(v)!r}; it needs a number")
                cells[j] = None
        cells[6] = None if cells[6] is None or not str(cells[6]).strip() else str(cells[6]).strip()
        rows.append(cells)
    if len(rows) > MAX_STANDARDS:
        problems.append(f"Standards holds {len(rows)} rows; the workbook takes at most {MAX_STANDARDS}")
    if problems:
        raise LabConfigError(problems)
    data["standards"] = standards_from_rows(rows)
    return config_from_dict(data), notes


def config_from_xlsx(source) -> Config:
    """The configuration in an Excel lab configuration (see read_lab_config)."""
    return read_lab_config(source)[0]
