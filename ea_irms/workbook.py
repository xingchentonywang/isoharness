"""Write the correction workbook for one sequence.

The workbook carries live formulas rather than computed values, so a student can
follow every step and change a standard's accepted value without the sheet going
stale. `ea_irms.pipeline` computes the same numbers in Python; the test suite
holds the two to each other.
"""
from __future__ import annotations

from pathlib import Path

"""Build the EA-IRMS student correction template (.xlsx), formula-driven."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule


from .metadata import (NOT_MATCHED, RUN_FIELDS, clean_run_info, critical_range_factor,
                       excel_trim, match_positions, match_summary, output_columns, output_header, summary_text)
from .cells import put
from .parse import _read_rows, parse_export
from .pipeline import (ANCHOR_RATIO_LIMIT, ANCHOR_SPREAD_PERMIL, MAX_STANDARDS,
                       MIN_OUTLIER_MONITORS, Config, assign_roles)

VERSION = "1.4.14"
NSTD = MAX_STANDARDS      # rows in the Setup standards table (16 from 1.4.2, 13 before)
# 1.4.10: the largest export the workbook is built for. Before 1.4.10 it held positions 1 to 120 and
# 1,000 peak records and left out the rest without a word; it now grows to fit, up to these.
MAX_POSITION = 400
MAX_RECORDS = 4000


# 1.4.10: how long LibreOffice may take over one workbook, in seconds (600 before). The workbooks of
# the nine reference runs took 4 to 9 s, and one of 400 positions and 3,474 peak records 14 s.
BAKE_TIMEOUT_S = 120.0


def bake_values(path, timeout: float | None = None) -> bool:
    """Have LibreOffice compute every formula and store the results in the file.

    openpyxl writes formulas without cached values. Excel computes them when it
    opens the file, but a preview pane, a file card, Protected View, or a viewer
    that does not calculate shows every formula cell empty. When LibreOffice is
    on the PATH this recalculates the workbook in place, keeping every formula,
    chart, validation and format, so the numbers show anywhere. Returns whether
    it ran. From 1.4.10 LibreOffice runs in a process group of its own, and the
    whole group is stopped after ``timeout`` seconds (BAKE_TIMEOUT_S), so no
    soffice.bin is left holding the profile.
    """
    import os
    import shutil
    import signal
    import subprocess
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    path = Path(path)
    with tempfile.TemporaryDirectory() as tmp:
        try:
            proc = subprocess.Popen([soffice, "--headless", "--calc", "--convert-to", "xlsx",
                                     "--outdir", tmp, str(path)],
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
        except OSError:
            return False
        try:
            proc.communicate(timeout=BAKE_TIMEOUT_S if timeout is None else timeout)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(proc.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                pass
            proc.communicate()
            return False
        if proc.returncode != 0:
            return False
        produced = Path(tmp) / path.name
        if not produced.exists() or produced.stat().st_size == 0:
            return False
        shutil.copyfile(produced, path)
    return True


def build_workbook(export_path, out_path, config: Config | None = None,
                   run_name: str | None = None, run_date: str | None = None,
                   recalc: bool = False, samples=None, run_info: dict | None = None,
                   file_name: str | None = None, config_notes: list | None = None):
    """Build the workbook from an Isodat export. Returns ``out_path``.

    With ``recalc`` the finished file is passed through :func:`bake_values` when
    LibreOffice is available, so it carries computed values as well as formulas.
    From 1.4.8 ``samples`` (an ea_irms.metadata.SampleSheet) is matched to the run's samples, its
    rows written to a Metadata sheet and after the Results columns, and the replicates it names
    compared on a Replicates sheet; ``run_info`` holds Operator, Instrument, Project, and Notes.
    From 1.4.10 ``config_notes`` (what ea_irms.pipeline.scope_to_run said about the configuration) is
    written under the Setup sheet's title.
    """
    config = config or Config()
    if len(config.standards) > NSTD:
        raise ValueError(f"the standards table holds {NSTD} rows and the configuration has "
                         f"{len(config.standards)}; the workbook would drop the rest silently, "
                         "so merge or remove rows first")
    SRC = str(export_path)
    OUT = str(out_path)
    _CFG = config
    _RUN = run_name or Path(export_path).stem
    _DATE = run_date or ""
    _FILE = file_name or Path(export_path).name
    # 1.4.8: the sample sheet's matches and the replicate groups, found as the engine finds them
    _INFO = clean_run_info(run_info) if run_info else {}
    _META = sorted(parse_export(SRC), key=lambda q: q.row)
    assign_roles(_META, config)
    _GROUPS = match_positions(_META, samples)
    _MATCH = match_summary(_META, samples)
    _MCOLS = output_columns(samples, [_META])
    header, records = _read_rows(SRC)

    # 1.4.10: the Raw and Work sheets are sized from the export, at least as large as before:
    # Work has one row per position number from 1, so a run numbered 101 to 156 needs 156.
    # Up to the limits below; a larger export is refused rather than cut short.
    if any(q.row < 1 for q in _META):
        raise ValueError("the export has a position numbered below 1; the workbook numbers its rows "
                         "from position 1")
    top_pos = max((q.row for q in _META), default=0)
    if top_pos > MAX_POSITION:
        raise ValueError(f"the export's positions run to {top_pos}, and the workbook holds positions 1 "
                         f"to {MAX_POSITION}")
    if len(records) > MAX_RECORDS:
        raise ValueError(f"the export has {len(records)} peak records, and the workbook holds "
                         f"{MAX_RECORDS}")
    MAXRAW = max(1000, len(records))     # raw data rows 2..MAXRAW+1 (1001 before 1.4.10)
    R0, NSAMP = 6, max(120, top_pos)    # Work data rows 6..NSAMP+5 (125 before 1.4.10)
    UT_N, UT_C = 261, 281  # first data row of the two uncertainty tables on Fits (1.4.2: moved to the end)

    # ---------------------------------------------------------------- styling
    ARIAL = 'Arial'
    F_TITLE = Font(name=ARIAL, size=14, bold=True)
    F_H1    = Font(name=ARIAL, size=11, bold=True)
    F_HDR   = Font(name=ARIAL, size=9, bold=True, color='FFFFFF')
    F_GRP   = Font(name=ARIAL, size=9, bold=True)
    F_BODY  = Font(name=ARIAL, size=10)
    F_SM    = Font(name=ARIAL, size=9)
    F_IN    = Font(name=ARIAL, size=9, color='0000FF')       # student input
    F_NOTE  = Font(name=ARIAL, size=9, italic=True, color='595959')
    FILL_HDR  = PatternFill('solid', fgColor='1F4E79')
    FILL_GRP  = PatternFill('solid', fgColor='DDEBF7')
    FILL_IN   = PatternFill('solid', fgColor='FFF2CC')       # yellow = type here
    FILL_HELP = PatternFill('solid', fgColor='F2F2F2')       # grey = machinery
    FILL_OUT  = PatternFill('solid', fgColor='E2EFDA')       # green = reported value
    FILL_RERUN = PatternFill('solid', fgColor='F8CBAD')      # orange = rerun (1.4.4)
    FILL_CHECK = PatternFill('solid', fgColor='FCE4D6')      # pale orange = check (1.4.4)
    THIN = Side(style='thin', color='BFBFBF')
    BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    wb = Workbook()

    # ================================================================ README
    ws = wb.active
    ws.title = 'README'
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 104

    def rd(row, text, style=F_BODY):
        c = ws.cell(row, 2, text); c.font = style; c.alignment = Alignment(wrap_text=True, vertical='top')
        return row + 1

    r = 2
    ws.cell(r, 2, 'EA-IRMS correction template').font = F_TITLE
    r += 2
    r = rd(r, 'For a Thermo Flash EA coupled to a Delta or 253 IRMS. The workbook takes the Isodat export and returns '
              f'%C, %N, C:N, delta-13C, and delta-15N, with every correction visible as a live formula. ea_irms {VERSION}.')
    r += 1
    r = rd(r, 'How to use it', F_H1)
    r = rd(r, f'1.  Open the Raw sheet. Select rows 2 to {MAXRAW + 1} and press Delete, then paste your Isodat export '
              'starting in cell A2. Do not paste over row 1, whose headers every other sheet looks for. The sheet '
              f'comes loaded with the export this workbook was built from ({_FILE}), so delete that first. This '
              f'workbook has room for {MAXRAW} peak records and positions 1 to {NSAMP}; build a new one for an '
              'export that needs more.')
    r = rd(r, '2.  Open Setup and check the standards table against the certificates and against the names the '
              'sequence used. A material is recognized when the text in column A appears anywhere in Identifier 1, '
              'ignoring case, so "Ace Drift" catches Ace Drift1 through Ace Drift 13. When two rows match, the lower '
              'row wins, so put the more specific text lower down. A capsule of another material whose name holds '
              'an anchor\'s text, such as USGS 64 Drift, is read as that anchor, and from 1.4.13 QC row 177 says '
              'when the anchors of one row do not read as one material (Fits rows 369 to 391).')
    r = rd(r, '3.  Open Work and type sample masses in the yellow column F, in mg. Leave a cell empty to use the '
              'mass the export already carries in its Amount column. Without a mass there is no %C or %N for that '
              'row; C:N, which comes from the two peaks through the k-factor lines, and the isotope values still '
              'come out.')
    r = rd(r, '4.  Read QC before Results. It flags check standards that missed, samples outside the calibrated '
              'size range, blanks that cannot support a blank correction, a reference gas that moved during the '
              'run, and positions whose background rose right after an oversized sample.')
    r = rd(r, '5.  Take the numbers from Results, each with its 1σ. Work is the audit trail behind them, and Fits '
              'holds every fitted coefficient with its r-squared, its residual scatter, and, for the corrections, '
              'the test that decided whether it was applied.')
    r = rd(r, '6.  Open Rerun. It lists the samples to measure again, with the reason and a mass to weigh, and '
              'the samples to check first. Its top line says whether the run as a whole has a problem.')
    r = rd(r, '7.  Metadata holds the run information and, when a sample sheet was given, each sample\'s row of it, '
              'which Results also carries after its own columns. Replicates compares the analyses of one sample '
              'in this run (1.4.8).')
    r += 1
    r = rd(r, 'Color', F_H1)
    r = rd(r, 'Yellow cells with blue text are yours to fill in. Gray columns are machinery and can be ignored. '
              'Green columns are the reported values. Everything else is a formula, and overtyping one breaks that '
              'row silently, so work on a copy. Orange rows on Results and Rerun are samples to measure again; '
              'pale orange rows are samples to check first.')
    r += 1
    r = rd(r, 'What the sheets do', F_H1)
    for a, b in [
        ('Raw',     'Paste target, unaltered. Gray helper columns label each peak record by channel and give it a key, '
                    'and, from 1.4.10, pick out each reference pulse Isodat did not designate (columns U to AD). '
                    'The sample CO2 peak is the one right before the first designated CO2 pulse; where two or more '
                    'come before it, Flags says so (1.4.12).'),
        ('Setup',   'Everything you are meant to change: processing switches, the standards table, and flag thresholds.'),
        ('Work',    'One row per autosampler position. Reads left to right: parsed peaks, then blank, linearity, '
                    'drift, and calibration corrections, then elemental content.'),
        ('Fits',    'Every coefficient fitted from this run, each with n, r-squared, and residual scatter.'),
        ('Results', 'The clean table to copy out or plot from.'),
        ('Rerun',   'The samples to measure again or to check, with the reason and a suggested mass.'),
        ('Replicates', 'Analyses of one sample in this run, their range, and the critical range of ISO 5725-6 (1.4.8).'),
        ('QC',      'Checks and flags, to read before Results.'),
        ('Metadata', 'The run information and the sample sheet\'s rows for this run\'s samples (1.4.8).'),
    ]:
        c = ws.cell(r, 2); c.value = f'{a}.  {b}'; c.font = F_SM
        c.alignment = Alignment(wrap_text=True, vertical='top'); r += 1
    r += 1
    r = rd(r, 'How the corrections work', F_H1)
    r = rd(r, 'The corrections run in the order blank, linearity, drift, and calibration, each fitted from this '
              'run\'s own standards. Because linearity is fitted after the blank correction, the two cannot count '
              'twice a size dependence that the blank already explains.')
    r = rd(r, 'Calibration regresses accepted values on measured values and applies that line as it is. A line '
              'fitted the other way, measured on accepted, has to be inverted, and subtracting its intercept over '
              'its slope from each value shifts every value by a constant without applying the slope. The line '
              'needs anchors of two materials with different accepted values; when every anchor carries the same '
              'value, nothing is calibrated, and the top line of Rerun says why (1.4.10). From 1.4.12 the line '
              'itself is judged too, its slope against 1 and the anchors\' residual SD about it, at the Note and '
              'Caution limits of Setup B73 to B76 (Fits rows 360 to 367, QC rows 23 and 25); a line beyond either '
              'limit is a problem of the whole run on Rerun.')
    r = rd(r, 'Drift is fitted twice, against autosampler position, the default, and against the mass-28 background '
              'on the reference peaks, and QC reports the residual scatter of both. Residual scatter cannot decide '
              'between the two, since both were fitted to the same drift standards, so run labeled check standards '
              'and let their offsets decide.')
    r += 1
    r = rd(r, 'Before trusting a drift correction, read the settling section of QC. While the instrument comes to rest '
              'at the start of a run, the mass-28 background decays toward a floor and the standards follow it down '
              'until both stop, and a drift fit reads that transient as drift. QC finds the position where the '
              'background settles, as the first three readings in a row near the floor, reports the drift slope '
              'separately before and after it, and says which pattern the run shows. A spike later in the run is '
              'carryover from an oversized sample and does not move that position. A straight line in position '
              'fits a transient badly, under-correcting the early positions and over-correcting the late ones, so '
              'Setup B14 can leave the unsettled positions out of the drift fit; waiting for the background to reach '
              'its floor before the first sample avoids the transient.')
    r += 1
    r = rd(r, 'QC also lists the scatter of the drift standards after each correction in turn. A blank or drift '
              'correction that does not lower it has not improved this run; linearity is judged by the check '
              'standards instead, since the drift standards are all one size.')
    r += 1
    r = rd(r, 'Only corrections the run supports', F_H1)
    r = rd(r, 'On Auto, the default, a correction is applied only when this run supports it, the blank when at least '
              'three blank capsules agree to within 2 permil, and linearity and drift when the fitted slope differs '
              'from zero at the significance level in Setup B15 (a two-sided t-test of the slope). That follows the '
              'FIRMS Good Practice Guide for IRMS (Dunn and Carter, 2018, section 6.4), which asks that a correction '
              'be shown necessary before it is applied. With five or six standards the test has little power, so it '
              'can miss a slope that is there; set the switch to On to apply it regardless. Setup B11 sets linearity for both '
              'elements unless B18 sets carbon on its own. Each drift monitor is also compared with the line through '
              'the others, and one that sits far off is flagged and kept in the fit; if the capsule was the cause, '
              'list its position in Setup B16 or B17 and that element\'s fit leaves it out. Those lists, B14, and '
              'B37 belong to this run: a lab configuration made for another run leaves them empty here, and the '
              'line under the title of Setup says so (1.4.10).')
    r += 1
    r = rd(r, 'Uncertainty', F_H1)
    r = rd(r, 'Results gives every calibrated value with its uncertainty beside it as one standard deviation, 1σ '
              '(from 1.4.5; earlier versions put U = k x 1σ there). The checks on QC and on the Rerun sheet compare '
              'a standard\'s deviation from its accepted value with k x 1σ, k from Setup B32 (2 by default). '
              'The budget follows section 6.5 of the guide and has three terms, the 1σ certificate uncertainty of '
              'each anchor material (Setup columns H and I), the scatter of its replicates in this run, and the '
              'repeatability of one sample analysis (Setup B30 and B31, or the pooled anchor scatter when those are '
              'empty), each propagated through the calibration line to first order, by its linear effect on the '
              'result, which a Kragten spreadsheet estimates by changing each input by its 1σ. A material '
              'analyzed once has no scatter of its own, so the error of its mean takes that repeatability (from '
              '1.4.12; before, it counted as 0 and the 1σ was too small). It leaves out '
              'the blank, linearity, and drift corrections and day-to-day reproducibility, which a control chart of '
              'an in-house standard measures.')
    r = rd(r, 'Elemental content comes from a regression of peak area on micrograms of C or N across every standard '
              'with both a known mass and a known composition. The intercept of that line is the area blank and is '
              'subtracted, so the k-factor is not forced through the origin. A peak at or below the intercept holds '
              'no N or C that the line can detect, so that position has no %N (or %C) and no C:N, and Flags says '
              'below detection (1.4.10).')
    r = rd(r, 'From 1.4.3 Results also carries the 1σ of %N, %C, and C:N, each from five terms combined in '
              'quadrature (Fits row 300 onwards): the instrument\'s repeatability, from the drift monitors\' measured '
              'over certified ratios by successive differences in position, less the part their own weighing '
              'explains; the position\'s weighing, the balance\'s 1σ in Setup B33 over its mass; the calibration, the '
              'interpolated elemental drift factor when that correction is in force and the k-factor line at the '
              'position\'s micrograms when it is not; the certificate uncertainty of the material that sets the scale, '
              'from Setup columns J and K; and trueness, the anchors\' RMS deviation from their certified values with '
              'their certificate uncertainty, as in the Nordtest TR 537 handbook. C:N takes the same terms without '
              'the weighing, which cancels in a ratio. Empty inputs count as 0. Entering the certificates raises every '
              '1σ; entering the balance raises 1σ for positions lighter than the drift monitors and lowers it '
              'slightly for heavier ones, because the monitors\' own weighing then comes out of the repeatability. '
              'QC lists how many anchors and linearity standards fall within k x 1σ of their certified values.')
    r += 1
    r = rd(r, 'Elemental drift (1.4.2)', F_H1)
    r = rd(r, 'That k-factor is one line for the whole run. When the sensitivity of the detector moves during the '
              'run, the drift monitors read high or low in %N and %C by position, and so do the samples around '
              'them. Each monitor\'s measured over certified %N (and %C) gives a factor at its position; between '
              'monitors the factor is interpolated in position, beyond the first and last it is held, and every '
              'position\'s micrograms are divided by the factor at its position. On Auto the correction is applied '
              'when a monitor is further than Setup B36 (percent) from its certified value; On applies it always, '
              'Off never. A monitor whose capsule was weighed or entered wrongly would be carried into the samples, '
              'so list such positions in Setup B37 and they are left out. Results reports the corrected %N, %C, and '
              'C:N, with the factors and the values as measured in its last columns; the isotope values are not '
              'touched.')
    r = rd(r, 'The correction does not use the anchors\' certified %N and %C, so from 1.4.13 the anchors test it. '
              'Fits rows 393 to 404 give their RMS deviation from certified as measured and as corrected, and a '
              'correction that adds more than half of Setup B36 to it is a Note, and more than all of it a Caution '
              '(QC rows 178 to 181, and a problem of the whole run on Rerun). In one run the correction took the '
              'anchors\' %N from 0.65 to 6.5 percent RMS, because a monitor beside them was left out and the factor '
              'of the next one was held flat over them.')
    r += 1
    r = rd(r, 'Exports without reference peaks (1.4.2)', F_H1)
    r = rd(r, 'An export cut down to the sample peaks (Peak Nr 3 and 4) carries no reference peaks to count from. '
              'Where a position made no sample CO2, the first exported CO2 peak is then a reference pulse that '
              'Isodat did not flag, and it would be read as the sample. Such a pulse matches the designated pulses '
              'the export does carry (at positions with no sample gas) in amplitude, within Setup B39, and reads '
              'the working-gas value in d13C, within Setup B40; a sample peak is rejected only if it matches both. '
              'The template rejects such a pulse, reports the position without a carbon value, and says so in '
              'Flags. Leave B39 empty to switch the check off. The designated pulses\' amplitude is interpolated '
              'between the positions that carry one and held flat before the first and after the last, where it '
              'cannot follow the pulses\' drift, so from 1.4.12 those positions use the tolerance of Setup B41, or '
              'B39 where that is wider.')
    r += 1
    r = rd(r, 'Rerun marks (1.4.4, 1.4.7, 1.4.8, 1.4.14)', F_H1)
    r = rd(r, 'A sample is marked for what measuring it again can fix, namely the amount, the peak, or the state of '
              'the instrument at its position. It is marked for a rerun of N, C, or both when its peak lies below '
              'the range of the linearity standards (Fits rows 81 to 104) or under the blank multiple in Setup B25; '
              'when its peak reaches the instrument\'s largest amplitude, Setup B69 (1.4.7; a peak above the '
              'linearity range but under it is not marked, and with B69 empty every peak above the range is, as '
              'before 1.4.7); when a '
              'reference pulse at its position is off, further than Setup B67 robust SDs from the median of the '
              'run\'s pulses (from 1.4.10 every pulse Isodat did not designate counts, two of CO2 at each position '
              'in a full export); when the 1σ of its final d15N or d13C is above Setup B70 (1.4.7), which a run '
              'with better anchors can fix; or when it has replicates in the run whose range is above the critical '
              'range of ISO 5725-6 (1.4.8, Replicates sheet, switched by Setup B71). A sample with neither peak is '
              'marked for a rerun of N and C. It is marked for a check when it '
              'ran before the background settled, when it follows an oversized sample and its background rose by '
              'more than Setup B28 (possible carryover, a check from 1.4.14 and a rerun of N and C before), or when '
              'it has the peak of only one element. A value outside the anchors\' '
              'range is not marked, because a rerun would extrapolate the same way; the fix is an anchor that '
              'brackets it.')
    r = rd(r, 'The suggested mass puts each peak inside the linearity range and above the blank multiple, read from '
              'this analysis\'s own peak area per mg above the blank (from 1.4.10; a peak near the blank\'s size is '
              'mostly blank, and counting all of it as sample suggested far too little), and it says when it is more '
              'than the heaviest capsule in this run. When no one mass does that for both N and C, the two '
              'elements need two capsules and both ranges are given. Problems of the whole run, a check standard '
              'off, anchors outside k x 1σ, a reference-gas SD at or above Setup B68 (the FIRMS guide\'s '
              'zero-enrichment limit, section 4.3.2), an element with no calibration (1.4.10), a calibration '
              'line beyond the limits of Setup B73 to B76 (1.4.12), anchors of one row of the standards table that '
              'do not read as one material, or an elemental drift correction that moves the anchors further from '
              'their certified %N or %C (1.4.13), are listed once at the top of Rerun and not on every sample.')
    r = rd(r, 'Replicates are the analyses of one sample in this run as the sample sheet names them: samples matched '
              'to the sheet with the same name, and samples whose own row gives the same Replicate of. A sample '
              'matched by an ID its name begins with, such as CI 2-2-H+S by the row CI 2-2-H, takes that row\'s '
              'columns but not its Replicate of, so it is a replicate only of samples with its own name (1.4.10). '
              'Their range is compared with '
              'f(n) x s, the critical range of ISO 5725-6 (section 5.2), where s is the repeatability of one analysis '
              'in use times the calibration slope and f(n) is 2.77 for two analyses (the 2.8 of r = 2.8 s, '
              'unrounded), 3.31 for three, and 3.63 for four. When the range is above it, every one of them is '
              'marked for a rerun of that element, since the data cannot say which analysis is off; the standard\'s '
              'next step is more analyses and their median. Without a sample sheet nothing is compared.')
    r += 1
    r = rd(r, 'Known limits', F_H1)
    r = rd(r, 'The outlier test on the drift monitors needs five of them in the fit; with four it has one degree of '
              'freedom and flags residuals of a hundredth of a permil, so from 1.4.2 it is not run below five.')
    r = rd(r, 'Linearity is fitted as a straight line. On the 2026-09-15 run the fitted slope moved 36 percent when '
              'the smallest standard was included, which suggests a response curved across the full size range. A '
              'sample far from the standards will carry that curvature as an error, and the size flags in QC mark '
              'those rows.')
    r = rd(r, 'Two blank capsules whose delta-15N differ by 22 permil cannot constrain a blank correction, and QC '
              'says so. On Auto the correction is applied only when at least three blanks agree to within 2 permil. '
              'A sample barely larger than the blank is flagged separately whatever the blanks look like, because '
              'the mass balance carries an error in the blank\'s d value into the sample with the weight '
              'A_blank / (A - A_blank), which is 1 for a peak twice the blank.')
    r = rd(r, 'Precision quoted from the drift standards understates the scatter of an unknown, because the drift '
              'correction was fitted to those standards. QC reports both that number and the offsets of any check '
              'standards, which test the correction independently, and the Replicates sheet compares repeated '
              'samples.')
    r += 1
    c = ws.cell(r, 2, f'ea_irms {VERSION}. Built at the Wang Stable Isotope Biogeochemistry Lab, '
                      f'Boston College, and free to use and adapt. Cite the version when you publish '
                      f'data corrected with it.')
    c.font = F_NOTE

    # ================================================================ Raw
    raw = wb.create_sheet('Raw')
    RAW_HDR = ['Row', 'Identifier 1', 'Identifier 2', 'Peak Nr', 'Amount', 'Ampl  28', 'BGD 28', 'Area 28',
               'Is Ref _', 'Ampl  44', 'Area 44', 'Peak Nr  Flash TCD', 'rArea Flash TCD',
               'd 15N/14N', 'd 13C/12C', 'BGD Flash TCD']
    for j, h in enumerate(RAW_HDR, start=1):
        c = raw.cell(1, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    for j, h in enumerate(['Channel', 'Key'], start=18):
        c = raw.cell(1, j, h); c.font = F_GRP; c.fill = FILL_HELP
        c.alignment = Alignment(horizontal='center')
    raw.cell(1, 20, 'gray helper columns, computed').font = F_NOTE

    order = [header.index(h) if h in header else None for h in RAW_HDR]
    for i, record in enumerate(records, start=1):
        for j, src_j in enumerate(order):
            v = record[src_j] if src_j is not None and src_j < len(record) else ''
            if v is None or v == '':
                continue
            if j in (0, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14) and isinstance(v, str):
                try: v = float(v)                 # csv exports and BGD 28 come as text
                except ValueError: pass
            c = put(raw, i + 1, j + 1, v); c.font = F_SM     # text stays text (1.4.10)
    # 1.4.8: the identifier each Work row shows (the first Raw record of its position)
    RAW_ID = {}
    if 'Row' in header and 'Identifier 1' in header:
        i_row, i_id = header.index('Row'), header.index('Identifier 1')
        for record in records:
            try:
                k_ = int(float(record[i_row]))
            except (TypeError, ValueError, IndexError):
                continue
            RAW_ID.setdefault(k_, record[i_id] if i_id < len(record) else '')

    def lit(pos):
        """The identifier at a position as Excel's TRIM gives it, quoted for a formula."""
        return excel_trim(RAW_ID.get(pos, '')).replace('"', '""')
    REP_TOP = 9                                  # first member row on the Replicates sheet
    REP_FIRST, REP_OF, r_ = {}, {}, REP_TOP
    for g_ in _GROUPS:
        REP_FIRST[g_['group']] = r_
        r_ += len(g_['rows'])
        for row_ in g_['rows']:
            REP_OF[row_] = g_['group']

    # 1.4.10: every reference pulse Isodat did not designate, one record at a time (columns U to AD):
    # an unflagged mass-28 peak before the position's last one, and an unflagged mass-44 peak after its
    # first designated one, with the position for the trend, the distance from the run's median (Fits
    # B340 and B344) and a chart feed. Before 1.4.10 the checks read only the first such pulse.
    WKA = f'Work!$A${R0}:$A${R0 + NSAMP - 1}'
    at = lambda col, i: f'INDEX(Work!${col}${R0}:${col}${R0 + NSAMP - 1},MATCH($A{i},{WKA},0))'   # noqa: E731
    PULSE_HDR = ['pulse d15N', 'pulse x N', 'pulse d13C', 'pulse x C', 'pulse dev N', 'pulse adev N',
                 'pulse dev C', 'pulse adev C', 'chart pulse N', 'chart pulse C']
    for j, h in enumerate(PULSE_HDR, start=21):
        c = raw.cell(1, j, h); c.font = F_GRP; c.fill = FILL_HELP
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    for i in range(2, MAXRAW + 2):
        c = raw.cell(i, 18, f'=IF($D{i}="","",IF($F{i}<>"","N",IF($J{i}<>"","C","")))')
        c.font = F_SM; c.fill = FILL_HELP
        c = raw.cell(i, 19, f'=IF($R{i}="","",$A{i}&"|"&$R{i}&"|"&$D{i})')
        c.font = F_SM; c.fill = FILL_HELP
        hN, fC = at('H', i), at('BU', i)
        cells = [
            f'=IF(AND($R{i}="N",$I{i}=0,ISNUMBER($N{i})),IF($D{i}<IFERROR({hN},0),$N{i},""),"")',
            f'=IF($U{i}="","",$A{i})',
            f'=IF(AND($R{i}="C",$I{i}=0,ISNUMBER($O{i})),IF(AND(ISNUMBER(IFERROR({fC},"")),'
            f'IFERROR({fC},9999)<>9999,$D{i}>IFERROR({fC},9999)),$O{i},""),"")',
            f'=IF($W{i}="","",$A{i})',
            f'=IF(OR($U{i}="",Fits!$B$340=""),0,$U{i}-Fits!$B$340)',
            f'=IF(OR($U{i}="",Fits!$B$340=""),"",ABS($U{i}-Fits!$B$340))',
            f'=IF(OR($W{i}="",Fits!$B$344=""),0,$W{i}-Fits!$B$344)',
            f'=IF(OR($W{i}="",Fits!$B$344=""),"",ABS($W{i}-Fits!$B$344))',
            f'=IF($U{i}="",NA(),$U{i})',
            f'=IF($W{i}="",NA(),$W{i})',
        ]
        for j, v in enumerate(cells, start=21):
            c = raw.cell(i, j, v); c.font = F_SM; c.fill = FILL_HELP
            if j in (21, 23, 25, 26, 27, 28): c.number_format = '0.000'
    raw.freeze_panes = 'A2'
    for col, w in zip('ABCDEFGHIJKLMNOPQRS',
                      [6,16,11,8,9,9,8,9,8,9,9,11,12,11,11,11,3,9,13]):
        raw.column_dimensions[col].width = w
    for k in range(21, 31):
        raw.column_dimensions[gcl(k)].width = 9

    # ================================================================ Setup
    st = wb.create_sheet('Setup')
    st.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGH', [34, 14, 13, 13, 10, 10, 34, 4]):
        st.column_dimensions[col].width = w

    def lab(ws_, r_, txt, style=F_BODY, col=1):
        c = ws_.cell(r_, col, txt); c.font = style; return c

    def inp(ws_, r_, col, val, fmt=None):
        c = put(ws_, r_, col, val); c.font = F_IN; c.fill = FILL_IN; c.border = BOX
        if fmt: c.number_format = fmt
        return c

    st.cell(1, 1, 'Setup').font = F_TITLE
    st.cell(2, 1, 'Yellow cells are yours. Everything else on this sheet is a label.').font = F_NOTE
    if config_notes:
        # 1.4.10: what the configuration set for another run, or for no run named, and how it was used here
        c = put(st, 3, 1, 'From the configuration: ' + '; '.join(str(t) for t in config_notes) + '.')
        c.font = Font(name='Arial', size=9, bold=True, color='A33A00')
        c.alignment = Alignment(wrap_text=True, vertical='top')
        st.merge_cells('A3:G3')
        st.row_dimensions[3].height = 15 * max(2, min(8, 1 + len(c.value) // 150))

    lab(st, 4, 'Run information', F_H1)
    for i, (k, v) in enumerate([('Run name', _RUN), ('Run date', _DATE), ('Operator', _INFO.get('Operator', ''))]):
        lab(st, 5 + i, k, F_SM); inp(st, 5 + i, 2, v)

    lab(st, 9, 'Processing options', F_H1)
    OPTS = [
        ('Blank correction',     _CFG.blank_correction,
                                        'On / Off / Auto. Auto applies it only when at least three blanks agree '
                                        'to within 2 permil.'),
        ('Linearity correction', _CFG.linearity_correction,
                                        'On / Off / Auto. Auto applies the fitted slope only when it differs from '
                                        'zero at the significance level below. Sets nitrogen, and carbon too '
                                        'unless B18 is filled.'),
        ('Drift correction',     _CFG.drift_correction,
                                        'On / Off / Auto. Auto applies the fitted slope only when it differs from '
                                        'zero at the significance level below.'),
        ('Drift model',          _CFG.drift_model,
                                        'Row / Background / Auto. Auto takes the lower residual scatter, '
                                            'which is not an independent test. See the note under the table.'),
    ]
    for i, (k, v, note) in enumerate(OPTS):
        lab(st, 10 + i, k, F_SM); inp(st, 10 + i, 2, v); lab(st, 10 + i, note, F_NOTE, 7)
    dv_bk = DataValidation(type='list', formula1='"On,Off,Auto"', allow_blank=False)
    st.add_data_validation(dv_bk)
    dv_dm = DataValidation(type='list', formula1='"Auto,Row,Background"', allow_blank=False)
    st.add_data_validation(dv_dm)
    dv_bk.add('B10'); dv_bk.add('B11'); dv_bk.add('B12'); dv_dm.add('B13')

    lab(st, 14, 'First position to use in the drift fit', F_SM)
    inp(st, 14, 2, _CFG.first_drift_position or '')
    lab(st, 14, 'Leave empty to use every position. QC suggests a value. It applies to the drift monitors '
                'only; the calibration anchors are never left out.', F_NOTE, 7)
    lab(st, 15, 'Significance level for Auto', F_SM)
    inp(st, 15, 2, _CFG.significance_level, '0.00')
    lab(st, 15, 'A slope is applied on Auto when its two-sided p-value is below this. The FIRMS guide asks '
                'that a correction be shown necessary before it is applied.', F_NOTE, 7)
    lab(st, 16, 'Positions left out of the nitrogen drift fit', F_SM)
    inp(st, 16, 2, ', '.join(str(v) for v in (_CFG.exclude_drift_n or [])))
    lab(st, 16, 'Comma-separated autosampler positions, e.g. 59. Use it for a drift monitor QC calls an '
                'outlier when the cause is the capsule, not the instrument.', F_NOTE, 7)
    lab(st, 17, 'Positions left out of the carbon drift fit', F_SM)
    inp(st, 17, 2, ', '.join(str(v) for v in (_CFG.exclude_drift_c or [])))
    lab(st, 17, 'Separate lists, because a capsule can go wrong in one element only.', F_NOTE, 7)
    for rr in (16, 17):
        st.cell(rr, 2).number_format = '@'
    lab(st, 18, 'Linearity correction, carbon only', F_SM)
    inp(st, 18, 2, _CFG.linearity_correction_c or '')
    lab(st, 18, 'Leave empty to follow B11. On / Off / Auto, for a run where nitrogen and carbon need '
                'different answers (1.4.1).', F_NOTE, 7)
    dv_lc = DataValidation(type='list', formula1='"On,Off,Auto"', allow_blank=True)
    st.add_data_validation(dv_lc)
    dv_lc.add('B18')

    lab(st, 19, 'Flag thresholds', F_H1)
    # 1.4.13: an area limit of 0 is written as 0, not left empty; empty means the linearity standards' range,
    # while the engine reads 0 as a limit of 0 (found while fixing R-24)
    THR = [
        ('Check standard offset (permil)', _CFG.check_offset_permil,
                                              'Flag a check standard further than this from its accepted value.'),
        ('Minimum Area 28', '' if _CFG.min_area_n is None else _CFG.min_area_n,
                                              'Leave empty to use the smallest linearity standard in the run.'),
        ('Minimum Area 44', '' if _CFG.min_area_c is None else _CFG.min_area_c,
                                              'Leave empty to use the smallest linearity standard in the run.'),
        ('Maximum Area 28', '' if _CFG.max_area_n is None else _CFG.max_area_n,
                                              'Leave empty to use the largest linearity standard in the run.'),
        ('Maximum Area 44', '' if _CFG.max_area_c is None else _CFG.max_area_c,
                                              'Leave empty to use the largest linearity standard in the run.'),
        ('Minimum peak area, x the blank', _CFG.min_blank_multiple or '',
                                              'A peak under this multiple of the blank is flagged, since there the '
                                              'blank correction carries the blank\'s scatter into the value. Leave '
                                              'empty to skip the check.'),
        ('Settled when background excess is under', _CFG.settled_threshold or '',
                                               'Fraction of the run\'s starting excess left in the reference-peak '
                                               'background. Used only to suggest a first position and to flag.'),
        ('Drift monitor outlier, studentized residual', _CFG.monitor_outlier_t,
                                               'A drift monitor further than this from the line through the '
                                               'others, in units of their scatter, is flagged.'),
        ('Carryover: background rise after an oversized sample (mV)',
                                               '' if _CFG.carryover_mv is None else _CFG.carryover_mv,
                                               'Flags a position when one of the two before it is above the linearity '
                                               'range and its sample-peak background rose by more than this. Leave '
                                               'empty to skip the check.'),
    ]
    for i, (k, v, note) in enumerate(THR):
        lab(st, 20 + i, k, F_SM); inp(st, 20 + i, 2, v); lab(st, 20 + i, note, F_NOTE, 7)
    st.cell(26, 2).number_format = '0.00'
    st.cell(28, 2).number_format = '0.0'

    lab(st, 29, 'Uncertainty', F_H1)
    UNC = [
        ('Sample repeatability, d15N (1σ, permil)', '' if _CFG.u_sample_n is None else _CFG.u_sample_n,
                                               'One analysis of one sample. Leave empty to use the pooled scatter '
                                               'of the anchor replicates in this run; the long-term SD of a '
                                               'matrix-matched in-house standard is a better value.'),
        ('Sample repeatability, d13C (1σ, permil)', '' if _CFG.u_sample_c is None else _CFG.u_sample_c, ''),
        ('Multiple of 1σ for the checks (k)', _CFG.coverage_k,
                                               'Results gives 1σ. The checks compare a standard\'s deviation from '
                                               'its accepted value with k x 1σ, the anchors and linearity standards '
                                               'on QC and the anchors on Rerun. k = 2 is about 95 percent only when '
                                               'the replicates behind 1σ are many; with few, take k from '
                                               'Student\'s t.'),
    ]
    for i, (k, v, note) in enumerate(UNC):
        lab(st, 30 + i, k, F_SM); inp(st, 30 + i, 2, v); lab(st, 30 + i, note, F_NOTE, 7)
    st.cell(30, 2).number_format = '0.000'; st.cell(31, 2).number_format = '0.000'
    st.cell(32, 2).number_format = '0.0'
    # 1.4.3: the balance, for the weighing term of %N and %C
    lab(st, 33, 'Balance, 1σ of one weighing (mg)', F_SM)
    inp(st, 33, 2, '' if _CFG.balance_u_mg is None else _CFG.balance_u_mg, '0.0000')
    lab(st, 33, 'From the balance\'s certificate (its repeatability, 1σ). Leave empty to leave the weighing '
                'term out of the uncertainty of %N and %C; Fits says so.', F_NOTE, 7)

    # 1.4.2: elemental drift from the drift monitors, and the reference-pulse
    # check for exports without reference peaks.
    lab(st, 34, 'Elemental drift correction (1.4.2)', F_H1)
    lab(st, 35, 'Elemental drift correction', F_SM); inp(st, 35, 2, _CFG.elemental_correction)
    lab(st, 35, 'On / Off / Auto. Divides every position\'s %N and %C by the factor the drift monitors '
                'show at that position (measured over certified, interpolated between monitors). Auto '
                'applies it when a monitor is further off than the percentage below.', F_NOTE, 7)
    lab(st, 36, 'Applied on Auto when a monitor is off by more than (percent)', F_SM)
    inp(st, 36, 2, _CFG.elemental_threshold_pct, '0.0')
    lab(st, 36, 'The largest relative bias of the monitors used, in percent of the certified value.', F_NOTE, 7)
    lab(st, 37, 'Positions left out of the elemental correction', F_SM)
    inp(st, 37, 2, ', '.join(str(v) for v in (_CFG.exclude_elemental or [])))
    st.cell(37, 2).number_format = '@'
    lab(st, 37, 'Comma-separated. For a monitor whose capsule was weighed or entered wrongly; when the anchors '
                'beside it and the reference pulses do not share its change, the instrument did not cause it.', F_NOTE, 7)
    dv_el = DataValidation(type='list', formula1='"On,Off,Auto"', allow_blank=False)
    st.add_data_validation(dv_el); dv_el.add('B35')
    lab(st, 38, 'Reference-pulse check for exports without reference peaks (1.4.2)', F_H1)
    lab(st, 39, 'Amplitude tolerance, fraction of the designated pulses', F_SM)
    inp(st, 39, 2, '' if _CFG.ref_pulse_ampl_tol is None else _CFG.ref_pulse_ampl_tol, '0.000')
    lab(st, 39, 'An unflagged CO2 peak within this fraction of the designated pulses\' mean amplitude, and '
                'within the d13C window below, is a reference pulse and is not read as the sample. Only in '
                'an export with no N2 reference peaks. Leave empty to switch the check off.', F_NOTE, 7)
    lab(st, 40, 'd13C window around the designated pulses (permil)', F_SM)
    inp(st, 40, 2, _CFG.ref_pulse_d13c_window, '0.00')
    # 1.4.12 (Tony, 2026-10-14; finding W-03): before the first and after the last designated pulse
    lab(st, 41, 'Amplitude tolerance before the first and after the last designated pulse', F_SM)
    inp(st, 41, 2, '' if _CFG.ref_pulse_ampl_tol_outside is None else _CFG.ref_pulse_ampl_tol_outside, '0.000')
    lab(st, 41, 'There the yardstick is held flat and cannot follow the pulses\' drift, so the tolerance is wider '
                '(1.4.12). A value below B39 counts as B39; leave empty to use B39 everywhere.', F_NOTE, 7)

    lab(st, 42, 'Standards', F_H1)
    st.cell(43, 1, 'Accepted values below are the published USGS values as of 2026-09-16, and 1σ is the '
                   'uncertainty the producer lists, one standard deviation. Check both against your certificates. '
                   '1σ(%C) and 1σ(%N) (1.4.3) are left empty for your certificates; empty counts as 0.').font = F_NOTE
    SHDR = ['Match text in Identifier 1', 'Role', 'Accepted d13C', 'Accepted d15N', '%C', '%N', 'Material',
            '1σ(d13C)', '1σ(d15N)', '1σ(%C)', '1σ(%N)']
    for j, h in enumerate(SHDR, start=1):
        c = st.cell(44, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    STD_ROWS = [tuple(r) + (None,) * (11 - len(r)) for r in _CFG.standards]
    SFIRST, SLAST = 45, 45 + NSTD - 1        # 16 slots from 1.4.2, the last three left empty for the student
    for i in range(NSTD):
        for j in range(1, 12):
            c = st.cell(SFIRST + i, j); c.font = F_IN; c.fill = FILL_IN; c.border = BOX
            if j in (3, 4, 5, 6, 8, 9, 10, 11): c.number_format = '0.00'
        if i < len(STD_ROWS):
            for j, v in enumerate(STD_ROWS[i][:11], start=1):
                put(st, SFIRST + i, j, v if v != '' else None)
    dv_role = DataValidation(type='list', formula1='"Anchor,Check,Drift,Linearity,Blank,Exclude,Sample"',
                             allow_blank=True)
    st.add_data_validation(dv_role); dv_role.add(f'B{SFIRST}:B{SLAST}')

    lab(st, SLAST + 2, 'Constants', F_H1)
    for i, (k, v, note) in enumerate([('Atomic mass C (g/mol)', 12.011, 'IUPAC'),
                                      ('Atomic mass N (g/mol)', 14.007, 'IUPAC')]):
        lab(st, SLAST + 3 + i, k, F_SM)
        c = st.cell(SLAST + 3 + i, 2, v); c.font = F_SM; c.number_format = '0.000'
        lab(st, SLAST + 3 + i, note, F_NOTE, 7)
    AMC, AMN = f'Setup!$B${SLAST+3}', f'Setup!$B${SLAST+4}'

    # 1.4.4: the rerun marks' two settings, below the constants so nothing above moves
    lab(st, SLAST + 6, 'Rerun marks (1.4.4)', F_H1)
    lab(st, SLAST + 7, 'Reference pulse off at a position, robust SDs from the run\'s median', F_SM)
    inp(st, SLAST + 7, 2, '' if _CFG.ref_pulse_off_sd is None else _CFG.ref_pulse_off_sd, '0.0')
    lab(st, SLAST + 7, 'A sample whose undesignated N2 or CO2 reference pulse lies further than this from the '
                       'median of the run\'s pulses is marked for a rerun of that element. Leave empty to switch '
                       'the check off.', F_NOTE, 7)
    lab(st, SLAST + 8, 'Reference-gas SD that is a problem for the whole run (permil)', F_SM)
    inp(st, SLAST + 8, 2, '' if _CFG.ref_gas_sd_limit is None else _CFG.ref_gas_sd_limit, '0.00')
    lab(st, SLAST + 8, 'The FIRMS guide gives 0.1 permil as the usual limit for the SD of N2 and CO2 working-gas '
                       'pulses in its zero-enrichment test (section 4.3.2). At or above this the Rerun sheet says '
                       'so once. Leave empty to switch it off.', F_NOTE, 7)
    S_POFF, S_RGSD = f'Setup!$B${SLAST+7}', f'Setup!$B${SLAST+8}'
    # 1.4.7 (Tony, 2026-10-13): the instrument's largest amplitude, and the largest 1σ before a rerun
    lab(st, SLAST + 9, 'Largest peak amplitude the instrument measures (mV)', F_SM)
    inp(st, SLAST + 9, 2, '' if _CFG.max_amplitude_mv is None else _CFG.max_amplitude_mv, '0')
    lab(st, SLAST + 9, 'A peak above the linearity range is marked for a rerun only at or above this amplitude, and '
                       'a peak at or above it is marked whatever the range (1.4.7). Leave empty to mark every peak '
                       'above the linearity range, as before 1.4.7.', F_NOTE, 7)
    lab(st, SLAST + 10, 'Largest 1σ of the final d15N or d13C before a rerun (permil)', F_SM)
    inp(st, SLAST + 10, 2, '' if _CFG.rerun_max_1sigma is None else _CFG.rerun_max_1sigma, '0.00')
    lab(st, SLAST + 10, 'A sample whose final d15N or d13C has a 1σ above this is marked for a rerun of that '
                        'element (1.4.7). The 1σ leaves out the blank, linearity, and drift corrections. Leave '
                        'empty to switch it off.', F_NOTE, 7)
    S_AMAX, S_USIG = f'Setup!$B${SLAST+9}', f'Setup!$B${SLAST+10}'
    # 1.4.8: replicates of a sample, named by the sample sheet, against the critical range of ISO 5725-6
    lab(st, SLAST + 11, 'Compare replicates of a sample (On or Off)', F_SM)
    inp(st, SLAST + 11, 2, _CFG.rerun_replicates if _CFG.rerun_replicates in ('On', 'Off') else 'On')
    lab(st, SLAST + 11, 'Replicates are analyses of one sample in this run, as the sample sheet names them (Metadata '
                        'sheet). On marks them for a rerun of an element when their range is above the critical '
                        'range of ISO 5725-6 (Replicates sheet, 1.4.8).', F_NOTE, 7)
    dv_rep = DataValidation(type='list', formula1='"On,Off"', allow_blank=False)
    st.add_data_validation(dv_rep); dv_rep.add(f'B{SLAST + 11}')
    S_REP = f'Setup!$B${SLAST+11}'
    # 1.4.12 (Tony, 2026-10-14; finding W-02): the calibration line itself, at two levels
    lab(st, SLAST + 12, 'Calibration line (1.4.12)', F_H1)
    CAL_LIMITS = [
        ('Calibration slope further from 1 than this is a Note', _CFG.cal_slope_note, '0.000',
         'Either level is a problem of the whole run on the Rerun sheet, and QC rows 23 and 25 say which. Leave a '
         'cell empty to switch that test off.'),
        ('Calibration slope further from 1 than this is a Caution', _CFG.cal_slope_caution, '0.000', ''),
        ('Anchor residual SD about the line above this is a Note (permil)', _CFG.cal_sd_note, '0.00',
         'The scatter of the anchors about the calibration line (Fits rows 59 and 66).'),
        ('Anchor residual SD about the line above this is a Caution (permil)', _CFG.cal_sd_caution, '0.00', ''),
    ]
    for k_, (label_, v_, fmt_, note_) in enumerate(CAL_LIMITS):
        row_ = SLAST + 13 + k_
        lab(st, row_, label_, F_SM)
        inp(st, row_, 2, '' if v_ is None else v_, fmt_)
        if note_:
            lab(st, row_, note_, F_NOTE, 7)
    S_CSN, S_CSC = f'Setup!$B${SLAST+13}', f'Setup!$B${SLAST+14}'
    S_CDN, S_CDC = f'Setup!$B${SLAST+15}', f'Setup!$B${SLAST+16}'

    S_LIN, S_DRIFT, S_DM = 'Setup!$B$11', 'Setup!$B$12', 'Setup!$B$13'
    S_LINC = 'IF(Setup!$B$18<>"",Setup!$B$18,Setup!$B$11)'   # carbon switch, empty = follow B11 (1.4.1)
    S_FIRST, S_ALPHA, S_EXN, S_EXC = 'Setup!$B$14', 'Setup!$B$15', 'Setup!$B$16', 'Setup!$B$17'
    S_THR, S_MINN, S_MINC, S_MAXN, S_MAXC = ('Setup!$B$20', 'Setup!$B$21', 'Setup!$B$22',
                                             'Setup!$B$23', 'Setup!$B$24')
    S_MULT, S_SETTLE, S_TOUT = 'Setup!$B$25', 'Setup!$B$26', 'Setup!$B$27'
    S_CARRY = 'Setup!$B$28'
    S_USN, S_USC, S_K = 'Setup!$B$30', 'Setup!$B$31', 'Setup!$B$32'
    S_ELEM, S_ETHR, S_EXE = 'Setup!$B$35', 'Setup!$B$36', 'Setup!$B$37'      # 1.4.2
    S_RPA, S_RPD = 'Setup!$B$39', 'Setup!$B$40'                               # 1.4.2
    S_RPO = 'Setup!$B$41'                                                       # 1.4.12
    S_BAL = 'Setup!$B$33'                                                       # 1.4.3
    SROLE  = f'Setup!$B${SFIRST}:$B${SLAST}'
    SD13   = f'Setup!$C${SFIRST}:$C${SLAST}'
    SD15   = f'Setup!$D${SFIRST}:$D${SLAST}'
    SPC    = f'Setup!$E${SFIRST}:$E${SLAST}'
    SUPC   = f'Setup!$J${SFIRST}:$J${SLAST}'      # 1.4.3
    SUPN   = f'Setup!$K${SFIRST}:$K${SLAST}'
    SPN    = f'Setup!$F${SFIRST}:$F${SLAST}'
    for col, w in zip('ABCDEFGHI', [40, 14, 13, 13, 10, 10, 34, 9, 9]):
        st.column_dimensions[col].width = w

    # ================================================================ Work
    wk = wb.create_sheet('Work')
    from .pipeline import _NO_BGD                # finding R-23: the engine's two flag texts
    NO_BGD_N, NO_BGD_C = _NO_BGD['n'], _NO_BGD['c']
    COLS = [
        ('Row','A',6,'k'),('Identifier','B',16,'k'),('Std row','C',7,'h'),('Type','D',10,'k'),
        ('Mass from export (mg)','E',11,'k'),('Mass entered (mg)','F',11,'i'),('Mass used (mg)','G',11,'k'),
        ('peak N','H',7,'h'),('key N','I',9,'h'),('peak C','J',7,'h'),('key C','K',9,'h'),
        ('Area 28','L',9,'k'),('Ampl 28','M',9,'k'),('BGD 28','N',9,'k'),('d15N raw','O',10,'k'),
        ('Area 44','P',9,'k'),('Ampl 44','Q',9,'k'),('d13C raw','R',10,'k'),
        ('Accepted d15N','S',11,'k'),('Accepted d13C','T',11,'k'),('Certified %N','U',10,'k'),('Certified %C','V',10,'k'),
        ('d15N blank-corr','W',12,'k'),('d15N linearity-corr','X',13,'k'),('drift x (N)','Y',10,'h'),
        ('d15N drift-corr','Z',12,'k'),('d15N calibrated','AA',12,'o'),
        ('d13C blank-corr','AB',12,'k'),('d13C linearity-corr','AC',13,'k'),('drift x (C)','AD',10,'h'),
        ('d13C drift-corr','AE',12,'k'),('d13C calibrated','AF',12,'o'),
        ('ug N','AG',9,'k'),('ug C','AH',9,'k'),('%N','AI',8,'o'),('%C','AJ',8,'o'),
        ('C:N molar','AK',10,'o'),('C:N mass','AL',10,'o'),
        ('lin x N','AM',8,'h'),('lin y N','AN',8,'h'),('lin x C','AO',8,'h'),('lin y C','AP',8,'h'),
        ('drift x row','AQ',9,'h'),('drift x bgd','AR',9,'h'),('drift y N','AS',9,'h'),('drift y C','AT',9,'h'),
        ('cal x N','AU',8,'h'),('cal y N','AV',8,'h'),('cal x C','AW',8,'h'),('cal y C','AX',8,'h'),
        ('kfit x N','AY',9,'h'),('kfit y N','AZ',9,'h'),('kfit x C','BA',9,'h'),('kfit y C','BB',9,'h'),
        ('blank d15N','BC',9,'h'),('blank A28','BD',9,'h'),('blank d13C','BE',9,'h'),('blank A44','BF',9,'h'),
        ('drift cal d15N','BG',10,'h'),('drift cal d13C','BH',10,'h'),
        ('ref BGD 28','BI',10,'k'),('settling frac','BJ',10,'k'),
        ('drift x early','BK',9,'h'),('drift y early','BL',9,'h'),
        ('drift x late','BM',9,'h'),('drift y late','BN',9,'h'),
        ('stage raw','BO',9,'h'),('stage blank','BP',9,'h'),('stage lin','BQ',9,'h'),('stage drift','BR',9,'h'),
        ('ref peak N','BS',8,'h'),('ref pulse d15N','BT',10,'k'),('ref flagged C','BU',8,'h'),
        ('ref peak C','BV',8,'h'),('ref pulse d13C','BW',10,'k'),
        ('excl N','BX',6,'h'),('excl C','BY',6,'h'),('drift x N','BZ',9,'h'),('drift x C','CA',9,'h'),
        ('t drift N','CB',9,'k'),('t drift C','CC',9,'k'),
        ('cal one N','CD',7,'h'),('cal dev N','CE',9,'h'),('cal res N','CF',9,'h'),('cal sq N','CG',9,'h'),
        ('cal one C','CH',7,'h'),('cal dev C','CI',9,'h'),('cal res C','CJ',9,'h'),('cal sq C','CK',9,'h'),
        ('u d15N','CL',9,'o'),('U d15N','CM',9,'o'),('u d13C','CN',9,'o'),('U d13C','CO',9,'o'),
        ('settle streak','CP',8,'h'),('ref x N','CQ',7,'h'),('ref x C','CR',7,'h'),
    ]
    # One scalar match test per standards-table slot. A single SUMPRODUCT over
    # SEARCH(range, cell) computed correctly in LibreOffice and returned 0 in
    # Excel for Mac, which silently made every position a Sample; thirteen
    # plain cells cannot do that.
    MATCH_COLS = [gcl(wk_col) for wk_col in range(column_index_from_string('CS'),
                                                  column_index_from_string('CS') + NSTD)]
    COLS += [(f'match {k + 1}', col, 7, 'h') for k, col in enumerate(MATCH_COLS)]
    # Chart feeds. Excel plots a formula's empty string as zero, so every point a
    # chart should skip is NA() here; Excel and LibreOffice both leave those out.
    CHART_NAMES = ['ch pos', 'ch bgd', 'ch bgd sample', 'ch bgd monitor', 'mon bgd', 'mon bgd x',
                   'ch lin x N', 'ch lin y N', 'ch samp x N', 'ch samp y N', 'ch drift used N',
                   'ch drift out N', 'ch drift after N', 'ch cal y N', 'ch acc min N', 'ch acc max N', 'ch pulse N',
                   'ch lin x C', 'ch lin y C', 'ch samp x C', 'ch samp y C', 'ch drift used C',
                   'ch drift out C', 'ch drift after C', 'ch cal y C', 'ch acc min C', 'ch acc max C', 'ch pulse C',
                   'ch carry']
    first_chart_col = column_index_from_string(MATCH_COLS[-1]) + 1
    CH = {name: gcl(first_chart_col + k) for k, name in enumerate(CHART_NAMES)}
    COLS += [(name, CH[name], 7, 'h') for name in CHART_NAMES]
    # Carryover and settling helpers (1.4.0), after the chart feeds so no earlier
    # column moves. Plain IF, INDEX, MATCH and running counts only.
    HELP_NAMES = ['pos used', 'oversized', 'exposed', 'carry ref', 'carry rise', 'carry flag',
                  'carry rank', 'after oversized', 'ref first5', 'settle start']
    first_help_col = first_chart_col + len(CHART_NAMES)
    HX = {name: gcl(first_help_col + k) for k, name in enumerate(HELP_NAMES)}
    COLS += [(name, HX[name], 7, 'h') for name in HELP_NAMES]
    # Elemental drift and the reference-pulse check (1.4.2), after everything
    # else so no earlier column moves.
    ELEM_NAMES = ['desig x', 'desig ampl', 'pulse ampl', 'rej C', 'excl E',
                  'elem bias N', 'elem bias C', 'elem x N', 'elem f N',
                  'elem x C', 'elem f C', 'elem factor N', 'elem factor C',
                  '%N corrected', '%C corrected', 'C:N molar corrected', 'C:N mass corrected']
    first_elem_col = first_help_col + len(HELP_NAMES)
    EX = {name: gcl(first_elem_col + k) for k, name in enumerate(ELEM_NAMES)}
    COLS += [(name, EX[name], 10 if name.endswith('corrected') else 7,
              'o' if name.endswith('corrected') else 'h') for name in ELEM_NAMES]
    # Uncertainty of %N, %C, and C:N (1.4.3), after everything else so no earlier
    # column moves.
    UNC_NAMES = ['u%N cert', 'u%C cert', 'cert rel N', 'cert rel C',
                 'd2 N', 'w2 N', 'd2 C', 'w2 C', 'cn x', 'cn f', 'cn d2',
                 'mult N', 'mult C', 'mult CN', 'kfit N', 'kfit C', 'cal2 N', 'cal2 C', 'cal2 CN',
                 'scale2 N', 'scale2 C', 'adev N', 'adev C', 'adev CN', 'acert2 N', 'acert2 C', 'acert2 CN',
                 'urel N', 'urel C', 'urel CN',
                 'u %N', 'U %N', 'u %C', 'U %C', 'u CN molar', 'U CN molar', 'u CN mass', 'U CN mass',
                 'inU N', 'inU C', 'inU CN']
    UOUT = ('u %N', 'U %N', 'u %C', 'U %C', 'u CN molar', 'U CN molar', 'u CN mass', 'U CN mass')
    first_unc_col = first_elem_col + len(ELEM_NAMES)
    UX = {name: gcl(first_unc_col + k) for k, name in enumerate(UNC_NAMES)}
    COLS += [(name, UX[name], 9 if name in UOUT else 7, 'o' if name in UOUT else 'h') for name in UNC_NAMES]
    # Rerun marks (1.4.4), after everything else so no earlier column moves. The
    # flag pieces are the Flags text of Results, one piece per cell, so that the
    # marks read exactly what Flags says.
    RERUN_NAMES = ['pdev N', 'apdev N', 'poff N', 'pdev C', 'apdev C', 'poff C',
                   'fl noN', 'fl noC', 'fl rejC', 'fl sizeN', 'fl sizeC', 'fl checkN', 'fl checkC',
                   'fl blankN', 'fl blankC', 'fl settle', 'fl carry', 'fl pulseN', 'fl pulseC',
                   'chk off', 'chk list',
                   'per mg N', 'lo N', 'hi N', 'per mg C', 'lo C', 'hi C', 'mass lo', 'mass hi',
                   'rr N', 'rr C', 'rr both', 'rr check', 'rerun', 'why raw', 'why', 'suggested mass', 'rerun rank']
    ROUT = ('rerun', 'why', 'suggested mass')
    first_rerun_col = first_unc_col + len(UNC_NAMES)
    RX = {name: gcl(first_rerun_col + k) for k, name in enumerate(RERUN_NAMES)}
    COLS += [(name, RX[name], 12 if name in ROUT else 7, 'o' if name in ROUT else 'h') for name in RERUN_NAMES]
    # 1.4.7: Tony's two rules, after everything else so no earlier column moves: a peak at or above the
    # instrument's largest amplitude (Setup B69), a 1σ above Setup B70, and the size flags that count.
    RERUN2_NAMES = ['fl amplN', 'fl amplC', 'fl uN', 'fl uC', 'why sizeN', 'why sizeC']
    first_rerun2_col = first_rerun_col + len(RERUN_NAMES)
    RX.update({name: gcl(first_rerun2_col + k) for k, name in enumerate(RERUN2_NAMES)})
    COLS += [(name, RX[name], 7, 'h') for name in RERUN2_NAMES]
    # 1.4.8: the replicate group of the position (from the sample sheet, kept only while the identifier is
    # the one the sheet matched), and the rerun reasons the Replicates sheet gives it
    REP_NAMES = ['rep group', 'fl repN', 'fl repC']
    first_rep_col = first_rerun2_col + len(RERUN2_NAMES)
    RX.update({name: gcl(first_rep_col + k) for k, name in enumerate(REP_NAMES)})
    COLS += [(name, RX[name], 7, 'h') for name in REP_NAMES]
    # 1.4.10: each position's reference pulses, all of them (Raw columns U to AB)
    PULSE_NAMES = ['np N', 'pmax N', 'pmin N', 'np C', 'pmax C', 'pmin C']
    first_pulse_col = first_rep_col + len(REP_NAMES)
    RX.update({name: gcl(first_pulse_col + k) for k, name in enumerate(PULSE_NAMES)})
    COLS += [(name, RX[name], 7, 'h') for name in PULSE_NAMES]
    # 1.4.12 (finding W-09): how many CO2 peaks came before the designated pulse, and the flag that says the
    # last of them was taken
    CPRE_NAMES = ['c before', 'fl cpre']
    first_cpre_col = first_pulse_col + len(PULSE_NAMES)
    RX.update({name: gcl(first_cpre_col + k) for k, name in enumerate(CPRE_NAMES)})
    COLS += [(name, RX[name], 7, 'h') for name in CPRE_NAMES]
    # 1.4.13: the anchor checks, after everything else so no earlier column moves. For finding W-13 each
    # anchor's C/N peak-area ratio and raw d15N and d13C, as numbers with 0 where they do not apply and a 0/1
    # column saying where they do, so Fits can take a smallest and a largest value per row of the standards
    # table without array formulas; and each anchor's %N and %C as measured against certified, at the
    # analyses whose corrected deviation adev N and adev C hold, for the check of the elemental correction.
    ANC_NAMES = ['am ok ratio', 'am ratio', 'am ok d15N', 'am d15N', 'am ok d13C', 'am d13C', 'mdev N', 'mdev C']
    first_anc_col = first_cpre_col + len(CPRE_NAMES)
    RX.update({name: gcl(first_anc_col + k) for k, name in enumerate(ANC_NAMES)})
    COLS += [(name, RX[name], 7, 'h') for name in ANC_NAMES]
    GROUPS = [('A','D','identification'),('E','G','mass'),('H','K','peak lookup'),('L','R','as measured'),
              ('S','V','from the standards table'),('W','AA','nitrogen corrections'),
              ('AB','AF','carbon corrections'),('AG','AL','elemental content'),
              ('AM','BH','fit inputs and helpers, ignore'),('BI','BJ','instrument settling'),
              ('BK','BR','fit inputs and helpers, ignore'),('BS','BW','reference pulses'),
              ('BX','CK','fit inputs and helpers, ignore'),('CL','CO','uncertainty: 1σ, and k x 1σ for the checks'),
              ('CP','DE','fit inputs and helpers, ignore'),(CH['ch pos'],CH['ch carry'],'chart feeds, ignore'),
              (HX['pos used'],HX['settle start'],'carryover and settling helpers, ignore'),
              (EX['desig x'],EX['elem factor C'],'reference-pulse and elemental drift helpers, ignore'),
              (EX['%N corrected'],EX['C:N mass corrected'],'elemental content, drift-corrected'),
              (UX['u%N cert'],UX['urel CN'],'uncertainty of %N, %C, and C:N: helpers, ignore'),
              (UX['u %N'],UX['U CN mass'],'uncertainty of %N, %C, and C:N: 1σ, and k x 1σ for the checks'),
              (UX['inU N'],UX['inU CN'],'within k x 1σ, for QC'),
              (RX['pdev N'],RX['rr check'],'rerun marks: helpers, ignore'),
              (RX['rerun'],RX['suggested mass'],'rerun marks'),
              (RX['rerun rank'],RX['rerun rank'],'helper'),
              (RX['fl amplN'],RX['why sizeC'],'rerun marks (1.4.7): helpers, ignore'),
              (RX['rep group'],RX['fl repC'],'replicates (1.4.8): helpers, ignore'),
              (RX['np N'],RX['pmin C'],'reference pulses (1.4.10): helpers, ignore'),
              (RX['c before'],RX['fl cpre'],'CO2 peaks before the reference pulse (1.4.12): helpers, ignore'),
              (RX['am ok ratio'],RX['mdev C'],'anchor checks (1.4.13): helpers, ignore')]
    for a, b, t in GROUPS:
        wk.cell(4, wk[a + '1'].column, t).font = F_GRP
        wk.cell(4, wk[a + '1'].column).fill = FILL_GRP
    # 1.4.5: uncertainty columns are headed 1σ (u in the code) and kσ (U in the code, k x 1σ)
    SHOWN_AS = {'u d15N': '1σ d15N', 'U d15N': 'kσ d15N', 'u d13C': '1σ d13C', 'U d13C': 'kσ d13C',
                'u%N cert': '1σ %N cert', 'u%C cert': '1σ %C cert',
                'urel N': '1σ rel N', 'urel C': '1σ rel C', 'urel CN': '1σ rel CN',
                'u %N': '1σ %N', 'U %N': 'kσ %N', 'u %C': '1σ %C', 'U %C': 'kσ %C',
                'u CN molar': '1σ CN molar', 'U CN molar': 'kσ CN molar',
                'u CN mass': '1σ CN mass', 'U CN mass': 'kσ CN mass',
                'inU N': 'in kσ N', 'inU C': 'in kσ C', 'inU CN': 'in kσ CN'}
    for name, col, width, kind in COLS:
        j = wk[col + '1'].column
        c = wk.cell(5, j, SHOWN_AS.get(name, name)); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True, vertical='bottom')
        wk.column_dimensions[col].width = width
    wk.cell(1, 1, 'Work').font = F_TITLE
    wk.cell(2, 1, 'One row per autosampler position. Type sample masses in the yellow column F, in mg. '
                  'Everything else is a formula. Gray columns are machinery.').font = F_NOTE

    FILLS = {'i': FILL_IN, 'o': FILL_OUT, 'h': FILL_HELP, 'k': None}
    FONTS = {'i': F_IN, 'o': F_SM, 'h': F_SM, 'k': F_SM}
    NUMF = {'L':'0.000','M':'0','N':'0.0','BI':'0.0','BJ':'0.00','O':'0.000','P':'0.000','Q':'0','R':'0.000',
            'S':'0.00','T':'0.00','U':'0.00','V':'0.00','W':'0.000','X':'0.000','Y':'0.0',
            'Z':'0.000','AA':'0.00','AB':'0.000','AC':'0.000','AD':'0.0','AE':'0.000','AF':'0.00',
            'AG':'0.00','AH':'0.00','AI':'0.00','AJ':'0.00','AK':'0.00','AL':'0.00',
            'E':'0.000','F':'0.000','G':'0.000','BT':'0.000','BW':'0.000','BZ':'0.0','CA':'0.0',
            'CB':'0.0','CC':'0.0','CE':'0.000','CF':'0.000','CG':'0.00','CI':'0.000','CJ':'0.000','CK':'0.00',
            'CL':'0.000','CM':'0.000','CN':'0.000','CO':'0.000'}
    NUMF.update({EX['desig ampl']: '0', EX['pulse ampl']: '0.0', EX['elem bias N']: '0.0000', EX['elem bias C']: '0.0000', EX['elem f N']: '0.0000',
                 EX['elem f C']: '0.0000', EX['elem factor N']: '0.0000', EX['elem factor C']: '0.0000',
                 EX['%N corrected']: '0.00', EX['%C corrected']: '0.00',
                 EX['C:N molar corrected']: '0.00', EX['C:N mass corrected']: '0.00'})
    NUMF.update({UX[n]: '0.00' for n in ('u%N cert', 'u%C cert')})
    NUMF.update({UX[n]: '0.0000' for n in ('cert rel N', 'cert rel C', 'cn f', 'mult N', 'mult C', 'mult CN',
                                           'kfit N', 'kfit C', 'adev N', 'adev C', 'adev CN',
                                           'urel N', 'urel C', 'urel CN')})
    NUMF.update({UX[n]: '0.000' for n in UOUT})
    NUMF.update({RX[n]: '0.000' for n in ('pdev N', 'apdev N', 'pdev C', 'apdev C', 'lo N', 'hi N', 'lo C',
                                          'hi C', 'mass lo', 'mass hi')})
    NUMF.update({RX[n]: '0.0' for n in ('per mg N', 'per mg C')})
    NUMF.update({RX['am ratio']: '0.000', RX['am d15N']: '0.000', RX['am d13C']: '0.000',
                 RX['mdev N']: '0.0000', RX['mdev C']: '0.0000'})
    NUMF.update({UX[n]: '0.00000000' for n in ('d2 N', 'w2 N', 'd2 C', 'w2 C', 'cn d2', 'cal2 N', 'cal2 C',
                                               'cal2 CN', 'scale2 N', 'scale2 C', 'acert2 N', 'acert2 C',
                                               'acert2 CN')})

    W = lambda col: f"Work!${col}${R0}:${col}${R0+NSAMP-1}"      # noqa: E731
    RA, RB = 2, MAXRAW + 1
    rawA = f'Raw!$A${RA}:$A${RB}'; rawB = f'Raw!$B${RA}:$B${RB}'; rawD = f'Raw!$D${RA}:$D${RB}'
    rawE = f'Raw!$E${RA}:$E${RB}'; rawF = f'Raw!$F${RA}:$F${RB}'; rawG = f'Raw!$G${RA}:$G${RB}'
    rawH = f'Raw!$H${RA}:$H${RB}'; rawI = f'Raw!$I${RA}:$I${RB}'; rawJ = f'Raw!$J${RA}:$J${RB}'
    rawK = f'Raw!$K${RA}:$K${RB}'; rawN = f'Raw!$N${RA}:$N${RB}'; rawO = f'Raw!$O${RA}:$O${RB}'
    rawR = f'Raw!$R${RA}:$R${RB}'; rawS = f'Raw!$S${RA}:$S${RB}'
    OFF = SFIRST - 1

    def pullN(i, src):
        return (f'=IF($I{i}="","",IF(INDEX({rawI},MATCH($I{i},{rawS},0))<>0,"",'
                f'INDEX({src},MATCH($I{i},{rawS},0))))')
    def pullC(i, src):
        return (f'=IF($K{i}="","",IF(INDEX({rawI},MATCH($K{i},{rawS},0))<>0,"",'
                f'INDEX({src},MATCH($K{i},{rawS},0))))')
    def stdpull(i, rng):
        return f'=IF($B{i}="","",IF($C{i}=0,"",IF(INDEX({rng},$C{i}-{OFF})="","",INDEX({rng},$C{i}-{OFF}))))'

    for n in range(NSAMP):
        i = R0 + n
        f = {}
        f['A'] = n + 1
        f['B'] = f'=IFERROR(INDEX({rawB},MATCH($A{i},{rawA},0)),"")'
        f['C'] = f'=IF($B{i}="","",MAX(${MATCH_COLS[0]}{i}:${MATCH_COLS[-1]}{i}))'
        for k, col in enumerate(MATCH_COLS):
            sr = SFIRST + k
            f[col] = (f'=IF($B{i}="","",IF(Setup!$A${sr}="",0,'
                      f'IF(ISNUMBER(SEARCH(Setup!$A${sr},$B{i})),{sr},0)))')
        f['D'] = f'=IF($B{i}="","",IF($C{i}=0,"Sample",INDEX({SROLE},$C{i}-{OFF})))'
        f['E'] = (f'=IF($B{i}="","",IFERROR(IF(INDEX({rawE},MATCH($A{i},{rawA},0))=0,"",'
                  f'INDEX({rawE},MATCH($A{i},{rawA},0))),""))')
        f['G'] = f'=IF($B{i}="","",IF($F{i}<>"",$F{i},IF($E{i}="","",$E{i})))'
        f['H'] = f'=IF($B{i}="","",SUMPRODUCT(MAX(({rawA}=$A{i})*({rawR}="N")*{rawD})))'
        f['I'] = f'=IF(OR($H{i}="",$H{i}=0),"",$A{i}&"|N|"&$H{i})'
        # 1.4.12 (finding W-09): with a designated CO2 pulse, the CO2 peak right before it, as the engine
        # takes it; before 1.4.12 the first CO2 peak, which differs where two or more come before the pulse
        pre = f'({rawA}=$A{i})*({rawR}="C")*({rawD}<$BU{i})'
        f['J'] = (f'=IF($B{i}="","",IF(AND(ISNUMBER($BU{i}),$BU{i}<>9999,SUMPRODUCT({pre})>0),'
                  f'SUMPRODUCT(MAX({pre}*{rawD})),'
                  f'SUMPRODUCT(MIN(({rawA}=$A{i})*({rawR}="C")*{rawD}'
                  f'+(1-({rawA}=$A{i})*({rawR}="C"))*9999))))')
        f['K'] = f'=IF(OR($J{i}="",$J{i}=0,$J{i}=9999),"",IF(${EX["rej C"]}{i}=1,"",$A{i}&"|C|"&$J{i}))'
        # 1.4.2: an unflagged CO2 peak that matches the designated pulses in
        # amplitude (interpolated in position between them) and d13C, in an
        # export without N2 reference peaks.
        ckey = f'$A{i}&"|C|"&$J{i}'
        dkey = f'$A{i}&"|C|"&$BU{i}'
        DX, DA, PA = EX['desig x'], EX['desig ampl'], EX['pulse ampl']
        f[DX] = (f'=IF(OR($B{i}="",$BU{i}="",$BU{i}=9999),"",'
                 f'IF(ISNUMBER(INDEX({rawJ},MATCH({dkey},{rawS},0))),$A{i},""))')
        f[DA] = f'=IF(${DX}{i}="","",INDEX({rawJ},MATCH({dkey},{rawS},0)))'
        XR, FR = W(DX), W(DA)
        k = f'COUNTIF({XR},"<="&$A{i})'
        fk = f'INDEX({FR},MATCH(SMALL({XR},{k}),{XR},0))'
        fk1 = f'INDEX({FR},MATCH(SMALL({XR},{k}+1),{XR},0))'
        xk, xk1 = f'SMALL({XR},{k})', f'SMALL({XR},{k}+1)'
        f[PA] = (f'=IF($B{i}="","",IF(COUNT({XR})=0,"",'
                 f'IF({k}=0,INDEX({FR},MATCH(SMALL({XR},1),{XR},0)),'
                 f'IF({k}>=COUNT({XR}),INDEX({FR},MATCH(SMALL({XR},COUNT({XR})),{XR},0)),'
                 f'{fk}+($A{i}-{xk})/({xk1}-{xk})*({fk1}-{fk})))))')
        f[EX['rej C']] = (f'=IF(OR($J{i}="",$J{i}=0,$J{i}=9999),0,IF(Fits!$B$240<>"Yes",0,'
                          f'IF(INDEX({rawI},MATCH({ckey},{rawS},0))<>0,0,IF(${PA}{i}="",0,'
                          f'IF(AND(ISNUMBER(INDEX({rawJ},MATCH({ckey},{rawS},0))),'
                          f'ISNUMBER(INDEX({rawO},MATCH({ckey},{rawS},0))),'
                          f'ABS(INDEX({rawJ},MATCH({ckey},{rawS},0))-${PA}{i})<='
                          f'IF(OR($A{i}<MIN({XR}),$A{i}>MAX({XR})),IF({S_RPO}="",{S_RPA},MAX({S_RPA},{S_RPO})),{S_RPA})'
                          f'*${PA}{i},'
                          f'ABS(INDEX({rawO},MATCH({ckey},{rawS},0))-Fits!$B$239)<={S_RPD}),1,0)))))')
        f['L'] = pullN(i, rawH); f['M'] = pullN(i, rawF); f['N'] = pullN(i, rawG); f['O'] = pullN(i, rawN)
        f['P'] = pullC(i, rawK); f['Q'] = pullC(i, rawJ); f['R'] = pullC(i, rawO)
        f['S'] = stdpull(i, SD15); f['T'] = stdpull(i, SD13)
        f['U'] = stdpull(i, SPN);  f['V'] = stdpull(i, SPC)
        f['W'] = (f'=IF($O{i}="","",IF(OR(Fits!$C$5<>"Yes",$D{i}="Blank",$L{i}<=Fits!$B$6),'
                  f'$O{i},($L{i}*$O{i}-Fits!$B$6*Fits!$B$7)/($L{i}-Fits!$B$6)))')
        f['X'] = f'=IF($W{i}="","",$W{i}-Fits!$B$20*($L{i}-Fits!$B$16))'
        # Finding R-23 of the 2026-10-14 review: under the Background model a position with no reference
        # background has no drift x (Y and AD empty) and keeps its linearity-corrected value; before, its
        # position number stood in for the background
        f['Y'] = f'=IF($W{i}="","",IF(Fits!$B$38="Row",$A{i},IF($BI{i}="","",$BI{i})))'
        f['Z'] = f'=IF($X{i}="","",IF($Y{i}="",$X{i},$X{i}-Fits!$B$39*($Y{i}-Fits!$B$40)))'
        f['AA'] = (f'=IF(OR($Z{i}="",Fits!$C$55<>"Yes",NOT(ISNUMBER(Fits!$B$56))),"",'
                   f'Fits!$B$56*$Z{i}+Fits!$B$57)')
        f['AB'] = (f'=IF($R{i}="","",IF(OR(Fits!$C$9<>"Yes",$D{i}="Blank",$P{i}<=Fits!$B$10),'
                   f'$R{i},($P{i}*$R{i}-Fits!$B$10*Fits!$B$11)/($P{i}-Fits!$B$10)))')
        f['AC'] = f'=IF($AB{i}="","",$AB{i}-Fits!$B$28*($P{i}-Fits!$B$24))'
        f['AD'] = f'=IF($AB{i}="","",IF(Fits!$B$50="Row",$A{i},IF($BI{i}="","",$BI{i})))'
        f['AE'] = f'=IF($AC{i}="","",IF($AD{i}="",$AC{i},$AC{i}-Fits!$B$51*($AD{i}-Fits!$B$52)))'
        f['AF'] = (f'=IF(OR($AE{i}="",Fits!$C$62<>"Yes",NOT(ISNUMBER(Fits!$B$63))),"",'
                   f'Fits!$B$63*$AE{i}+Fits!$B$64)')
        f['AG'] = f'=IF(OR($L{i}="",Fits!$B$69<3),"",($L{i}-Fits!$B$71)/Fits!$B$70)'
        f['AH'] = f'=IF(OR($P{i}="",Fits!$B$75<3),"",($P{i}-Fits!$B$77)/Fits!$B$76)'
        # 1.4.10: no %N, %C or C:N where the peak is at or below the k-factor line's intercept
        f['AI'] = f'=IF(OR($AG{i}="",$G{i}="",$G{i}<=0,$AG{i}<=0),"",$AG{i}/($G{i}*1000)*100)'
        f['AJ'] = f'=IF(OR($AH{i}="",$G{i}="",$G{i}<=0,$AH{i}<=0),"",$AH{i}/($G{i}*1000)*100)'
        f['AK'] = f'=IF(OR($AG{i}="",$AH{i}="",$AG{i}<=0,$AH{i}<=0),"",($AH{i}/{AMC})/($AG{i}/{AMN}))'
        f['AL'] = f'=IF(OR($AG{i}="",$AH{i}="",$AG{i}<=0,$AH{i}<=0),"",$AH{i}/$AG{i})'
        f['AM'] = f'=IF($D{i}="Linearity",IF($L{i}="","",$L{i}),"")'
        f['AN'] = f'=IF($D{i}="Linearity",IF($W{i}="","",$W{i}),"")'
        f['AO'] = f'=IF($D{i}="Linearity",IF($P{i}="","",$P{i}),"")'
        f['AP'] = f'=IF($D{i}="Linearity",IF($AB{i}="","",$AB{i}),"")'
        G = f'IF($A{i}<IF({S_FIRST}="",0,{S_FIRST}),"",'
        f['AQ'] = f'=IF($D{i}="Drift",{G}$A{i}),"")'
        f['AR'] = f'=IF($D{i}="Drift",{G}IF($BI{i}="","",$BI{i})),"")'
        f['AS'] = f'=IF($D{i}="Drift",{G}IF(OR($X{i}="",$BX{i}=1),"",$X{i})),"")'
        f['AT'] = f'=IF($D{i}="Drift",{G}IF(OR($AC{i}="",$BY{i}=1),"",$AC{i})),"")'
        f['AU'] = f'=IF($D{i}="Anchor",IF(OR($Z{i}="",$S{i}=""),"",$Z{i}),"")'
        f['AV'] = f'=IF($D{i}="Anchor",IF(OR($Z{i}="",$S{i}=""),"",$S{i}),"")'
        f['AW'] = f'=IF($D{i}="Anchor",IF(OR($AE{i}="",$T{i}=""),"",$AE{i}),"")'
        f['AX'] = f'=IF($D{i}="Anchor",IF(OR($AE{i}="",$T{i}=""),"",$T{i}),"")'
        f['AY'] = (f'=IF(OR($D{i}="Blank",$D{i}="Exclude",$D{i}="Sample",$D{i}="",$G{i}="",$U{i}="",$L{i}=""),'
                   f'"",$G{i}*1000*$U{i}/100)')
        f['AZ'] = f'=IF($AY{i}="","",$L{i})'
        f['BA'] = (f'=IF(OR($D{i}="Blank",$D{i}="Exclude",$D{i}="Sample",$D{i}="",$G{i}="",$V{i}="",$P{i}=""),'
                   f'"",$G{i}*1000*$V{i}/100)')
        f['BB'] = f'=IF($BA{i}="","",$P{i})'
        f['BC'] = f'=IF(AND($D{i}="Blank",$L{i}<>""),$O{i},"")'
        f['BD'] = f'=IF($BC{i}="","",$L{i})'
        f['BE'] = f'=IF(AND($D{i}="Blank",$P{i}<>""),$R{i},"")'
        f['BF'] = f'=IF($BE{i}="","",$P{i})'
        f['BG'] = f'=IF(AND($D{i}="Drift",$BX{i}<>1),IF($AA{i}="","",$AA{i}),"")'
        f['BH'] = f'=IF(AND($D{i}="Drift",$BY{i}<>1),IF($AF{i}="","",$AF{i}),"")'
        # Finding R-23: a designated N2 peak whose BGD 28 cell is empty gives no background, as in the engine.
        # INDEX of the empty cell alone gave a 0 that LibreOffice's SMALL counted in the floor (Fits B83).
        ref_bgd = (f'INDEX({rawG},MATCH($A{i}&"|N|"'
                   f'&SUMPRODUCT(MAX(({rawA}=$A{i})*({rawR}="N")*({rawI}=1)*{rawD})),{rawS},0))')
        f['BI'] = f'=IF($B{i}="","",IFERROR(IF(ISNUMBER({ref_bgd}),{ref_bgd},""),""))'
        f['BJ'] = f'=IF(OR($BI{i}="",Fits!$B$85<=0),"",($BI{i}-Fits!$B$83)/Fits!$B$85)'
        f['BK'] = f'=IF($AS{i}="","",IF($A{i}>=Fits!$B$86,"",$A{i}))'
        f['BL'] = f'=IF($BK{i}="","",$X{i})'
        f['BM'] = f'=IF($AS{i}="","",IF($A{i}<Fits!$B$86,"",$A{i}))'
        f['BN'] = f'=IF($BM{i}="","",$X{i})'
        f['BO'] = f'=IF(AND($D{i}="Drift",$BX{i}<>1),IF($O{i}="","",$O{i}),"")'
        f['BP'] = f'=IF(AND($D{i}="Drift",$BX{i}<>1),IF($W{i}="","",$W{i}),"")'
        f['BQ'] = f'=IF(AND($D{i}="Drift",$BX{i}<>1),IF($X{i}="","",$X{i}),"")'
        f['BR'] = f'=IF(AND($D{i}="Drift",$BX{i}<>1),IF($Z{i}="","",$Z{i}),"")'
        # Reference pulses Isodat did not designate: for nitrogen the unflagged mass-28
        # peaks before the sample peak, for carbon the unflagged mass-44 peaks after
        # the designated reference. Their d values are measured, not assigned.
        selN = f'({rawA}=$A{i})*({rawR}="N")*({rawI}=0)*({rawD}<$H{i})'
        f['BS'] = (f'=IF(OR($B{i}="",$H{i}=""),"",SUMPRODUCT(MIN({selN}*{rawD}+(1-{selN})*9999)))')
        f['BT'] = (f'=IF(OR($BS{i}="",$BS{i}=9999,$BS{i}=0),"",'
                   f'IFERROR(INDEX({rawN},MATCH($A{i}&"|N|"&$BS{i},{rawS},0)),""))')
        selF = f'({rawA}=$A{i})*({rawR}="C")*({rawI}=1)'
        f['BU'] = f'=IF($B{i}="","",SUMPRODUCT(MIN({selF}*{rawD}+(1-{selF})*9999)))'
        selC = f'({rawA}=$A{i})*({rawR}="C")*({rawI}=0)*({rawD}>$BU{i})'
        f['BV'] = (f'=IF(OR($BU{i}="",$BU{i}=9999),"",SUMPRODUCT(MIN({selC}*{rawD}+(1-{selC})*9999)))')
        f['BW'] = (f'=IF(OR($BV{i}="",$BV{i}=9999,$BV{i}=0),"",'
                   f'IFERROR(INDEX({rawO},MATCH($A{i}&"|C|"&$BV{i},{rawS},0)),""))')
        # Positions left out of a drift fit, one element at a time.
        f['BX'] = (f'=IF($B{i}="","",IF(ISNUMBER(SEARCH(","&$A{i}&",",'
                   f'","&SUBSTITUTE({S_EXN}," ","")&",")),1,0))')
        f['BY'] = (f'=IF($B{i}="","",IF(ISNUMBER(SEARCH(","&$A{i}&",",'
                   f'","&SUBSTITUTE({S_EXC}," ","")&",")),1,0))')
        f['BZ'] = f'=IF($AS{i}="","",IF(Fits!$B$38="Row",$AQ{i},IF($AR{i}="","",$AR{i})))'
        f['CA'] = f'=IF($AT{i}="","",IF(Fits!$B$50="Row",$AQ{i},IF($AR{i}="","",$AR{i})))'
        # Externally studentized residual of each drift monitor: leave-one-out for
        # a monitor in the fit, prediction form for one left out.
        def stud(ycol, xrow, excl, model, srow, sbgd, nrow, mrow, sxxrow, irow, sdrow, sserow):
            x = f'IF(Fits!$B${model}="Row",$AQ{i},$AR{i})'
            slope = f'IF(Fits!$B${model}="Row",Fits!$B${srow},Fits!$B${sbgd})'
            e = f'(${ycol}{i}-({slope}*{x}+Fits!$B${irow}))'
            h = f'(1/Fits!$B${nrow}+({x}-Fits!$B${mrow})^2/Fits!$B${sxxrow})'
            inc = (f'IF(OR(Fits!$B${nrow}<{MIN_OUTLIER_MONITORS},{h}>=1),"",IF((Fits!$B${sserow}-{e}^2/(1-{h}))<=0,"",'
                   f'{e}/(SQRT((Fits!$B${sserow}-{e}^2/(1-{h}))/(Fits!$B${nrow}-3))*SQRT(1-{h}))))')
            exc = f'IF(OR(Fits!$B${sdrow}="",Fits!$B${sdrow}=0),"",{e}/(Fits!$B${sdrow}*SQRT(1+{h})))'
            G_ = f'$A{i}<IF({S_FIRST}="",0,{S_FIRST})'
            return (f'=IF(OR($D{i}<>"Drift",${ycol}{i}="",{G_},Fits!$B${nrow}<3,Fits!$B${irow}="",'
                    f'Fits!$B${sxxrow}=0,{x}=""),"",IF(${excl}{i}=1,{exc},{inc}))')
        f['CB'] = stud('X', 'AQ', 'BX', 38, 32, 35, 107, 108, 109, 110, 111, 116)
        f['CC'] = stud('AC', 'AQ', 'BY', 50, 44, 47, 120, 121, 122, 123, 124, 129)
        # Calibration helpers for the uncertainty budget.
        f['CD'] = f'=IF($AU{i}="","",1)'
        f['CE'] = f'=IF($AU{i}="","",$AU{i}-Fits!$B$133)'
        f['CF'] = (f'=IF(OR($AU{i}="",Fits!$C$55<>"Yes"),"",($AV{i}-(Fits!$B$56*$AU{i}+Fits!$B$57))'
                   f'-Fits!$B$56*($AU{i}-Fits!$B$133))')
        f['CG'] = f'=IF($AU{i}="","",$AU{i}^2)'
        f['CH'] = f'=IF($AW{i}="","",1)'
        f['CI'] = f'=IF($AW{i}="","",$AW{i}-Fits!$B$153)'
        f['CJ'] = (f'=IF(OR($AW{i}="",Fits!$C$62<>"Yes"),"",($AX{i}-(Fits!$B$63*$AW{i}+Fits!$B$64))'
                   f'-Fits!$B$63*($AW{i}-Fits!$B$153))')
        f['CK'] = f'=IF($AW{i}="","",$AW{i}^2)'
        # Standard uncertainty of the calibrated value: the sample's own
        # repeatability through the slope, plus each anchor material's accepted
        # value and measured mean through their exact effect on the fitted line.
        def unc(xcol, nrow, srow, irow, mrow, sxxrow, urow, t0):
            T = lambda col: f'Fits!${col}${t0}:${col}${t0+NSTD-1}'      # noqa: E731
            dx = f'(${xcol}{i}-Fits!$B${mrow})'
            return (f'=IF(OR(${xcol}{i}="",Fits!$C${nrow}<>"Yes",Fits!$B${nrow}<3,Fits!$B${mrow}="",Fits!$B${sxxrow}="",'
                    f'Fits!$B${sxxrow}=0,Fits!$B${urow}=""),"",'
                    f'SQRT((Fits!$B${urow}*Fits!$B${srow})^2'
                    f'+SUMPRODUCT(({T("J")}*({T("B")}/Fits!$B${nrow}+{T("H")}*{dx}/Fits!$B${sxxrow}))^2)'
                    f'+SUMPRODUCT(({T("G")}*({dx}*{T("I")}/Fits!$B${sxxrow}'
                    f'-{T("B")}*Fits!$B${srow}/Fits!$B${nrow}))^2)))')
        f['CL'] = unc('Z', 55, 56, 57, 133, 134, 136, UT_N)
        f['CM'] = f'=IF($CL{i}="","",{S_K}*$CL{i})'
        f['CN'] = unc('AE', 62, 63, 64, 153, 154, 156, UT_C)
        f['CO'] = f'=IF($CN{i}="","",{S_K}*$CN{i})'
        # Carryover (1.4.0). A position is exposed when one of the two before it is
        # above the linearity range; its rise is measured against the last
        # sample-peak background that was not exposed.
        OVc, EXc, CRc = HX['oversized'], HX['exposed'], HX['carry ref']
        CXc, CFc, SSc = HX['carry rise'], HX['carry flag'], HX['settle start']
        f[HX['pos used']] = f'=IF($B{i}="","",$A{i})'
        f[OVc] = (f'=IF(OR(AND(ISNUMBER($L{i}),Fits!$B$97>0,$L{i}>Fits!$B$97),'
                  f'AND(ISNUMBER($P{i}),Fits!$B$104>0,$P{i}>Fits!$B$104)),1,0)')
        if n == 0:
            f[EXc] = 0
            f[CRc] = f'=IF(AND(${EXc}{i}=0,ISNUMBER($N{i})),$N{i},"")'
            f[HX['after oversized']] = 0
        else:
            before = f'${OVc}{i-1}=1' if n == 1 else f'OR(${OVc}{i-1}=1,${OVc}{i-2}=1)'
            f[EXc] = f'=IF({before},1,0)'
            f[CRc] = f'=IF(AND(${EXc}{i}=0,ISNUMBER($N{i})),$N{i},${CRc}{i-1})'
            f[HX['after oversized']] = f'=IF(AND(${OVc}{i-1}=1,OR($D{i}="Sample",$D{i}="Check")),1,0)'
        f[CXc] = f'=IF(AND(${EXc}{i}=1,ISNUMBER($N{i}),ISNUMBER(${CRc}{i})),$N{i}-${CRc}{i},"")'
        f[CFc] = f'=IF({S_CARRY}="",0,IF(ISNUMBER(${CXc}{i}),IF(${CXc}{i}>{S_CARRY},1,0),0))'
        f[HX['carry rank']] = f'=IF(${CFc}{i}=1,COUNTIF(${CFc}${R0}:${CFc}{i},1),"")'
        f[CH['ch carry']] = f'=IF(AND(${CFc}{i}=1,ISNUMBER($N{i})),$N{i},NA())'
        # Settling (1.4.0): the scale comes from the first five reference-peak
        # backgrounds (Fits B85), and the run settles at the first of three
        # readings in a row at or under the threshold, skipping exposed rows.
        f[HX['ref first5']] = f'=IF(ISNUMBER($BI{i}),IF(COUNT($BI${R0}:$BI{i})<=5,$BI{i},""),"")'
        skip = f'OR(NOT(ISNUMBER($BJ{i})),${EXc}{i}=1)'
        below = f'$BJ{i}<={S_SETTLE}'
        if n == 0:
            f['CP'] = f'=IF({skip},0,IF({below},1,0))'
            f[SSc] = f'=IF({skip},"",IF({below},$A{i},""))'
        else:
            f['CP'] = f'=IF({skip},$CP{i-1},IF({below},$CP{i-1}+1,0))'
            f[SSc] = f'=IF({skip},${SSc}{i-1},IF({below},IF($CP{i-1}=0,$A{i},${SSc}{i-1}),""))'
        f['CQ'] = f'=IF($BT{i}="","",$A{i})'
        f['CR'] = f'=IF($BW{i}="","",$A{i})'
        # Elemental drift (1.4.2): each monitor's measured over certified %N and
        # %C, the factor interpolated in position between the monitors used,
        # and the corrected elemental content.
        f[EX['excl E']] = (f'=IF($B{i}="","",IF(ISNUMBER(SEARCH(","&$A{i}&",",'
                           f'","&SUBSTITUTE({S_EXE}," ","")&",")),1,0))')
        for el, meas, cert, inforce_row in (('N', 'AI', 'U', 246), ('C', 'AJ', 'V', 249)):
            B, X_, F_, FA = EX[f'elem bias {el}'], EX[f'elem x {el}'], EX[f'elem f {el}'], EX[f'elem factor {el}']
            f[B] = (f'=IF(AND($D{i}="Drift",ISNUMBER(${meas}{i}),ISNUMBER(${cert}{i}),${cert}{i}>0),'
                    f'${meas}{i}/${cert}{i}-1,"")')
            f[X_] = f'=IF(AND(${B}{i}<>"",${EX["excl E"]}{i}<>1),$A{i},"")'
            f[F_] = f'=IF(${X_}{i}="","",1+${B}{i})'
            XR, FR = W(X_), W(F_)
            k = f'COUNTIF({XR},"<="&$A{i})'
            fk = f'INDEX({FR},MATCH(SMALL({XR},{k}),{XR},0))'
            fk1 = f'INDEX({FR},MATCH(SMALL({XR},{k}+1),{XR},0))'
            xk, xk1 = f'SMALL({XR},{k})', f'SMALL({XR},{k}+1)'
            f[FA] = (f'=IF($B{i}="","",IF(Fits!$B${inforce_row}<>"Yes",1,IF(COUNT({XR})=0,1,'
                     f'IF({k}=0,INDEX({FR},MATCH(SMALL({XR},1),{XR},0)),'
                     f'IF({k}>=COUNT({XR}),INDEX({FR},MATCH(SMALL({XR},COUNT({XR})),{XR},0)),'
                     f'{fk}+($A{i}-{xk})/({xk1}-{xk})*({fk1}-{fk}))))))')
        FAN, FAC = EX['elem factor N'], EX['elem factor C']
        f[EX['%N corrected']] = f'=IF(OR($AI{i}="",${FAN}{i}=""),"",$AI{i}/${FAN}{i})'
        f[EX['%C corrected']] = f'=IF(OR($AJ{i}="",${FAC}{i}=""),"",$AJ{i}/${FAC}{i})'
        f[EX['C:N molar corrected']] = f'=IF(OR($AK{i}="",${FAN}{i}="",${FAC}{i}=""),"",$AK{i}*${FAN}{i}/${FAC}{i})'
        f[EX['C:N mass corrected']] = f'=IF(OR($AL{i}="",${FAN}{i}="",${FAC}{i}=""),"",$AL{i}*${FAN}{i}/${FAC}{i})'
        f[CH['ch pos']] = f'=IF($B{i}="",NA(),$A{i})'
        for el, linx, liny, area, used, linv, driftv, calv, accmin, accmax, pulse, rug in (
                ('N', 'AM', 'AN', 'L', 'AS', 'X', 'Z', 'AA', 'Fits!$B$138', 'Fits!$B$139', 'BT', 'Fits!$B$98'),
                ('C', 'AO', 'AP', 'P', 'AT', 'AC', 'AE', 'AF', 'Fits!$B$158', 'Fits!$B$159', 'BW', 'Fits!$B$105')):
            f[CH[f'ch lin x {el}']] = f'=IF(${liny}{i}="",NA(),${linx}{i})'
            f[CH[f'ch lin y {el}']] = f'=IF(${liny}{i}="",NA(),${liny}{i})'
            # The unknowns are drawn as a rug along the bottom of the linearity chart,
            # at a level just under the standards, so their sizes show without their
            # own isotope spread swamping the standards' trend.
            f[CH[f'ch samp x {el}']] = f'=IF(AND($D{i}="Sample",${area}{i}<>""),${area}{i},NA())'
            f[CH[f'ch samp y {el}']] = f'=IF(AND($D{i}="Sample",${area}{i}<>"",{rug}<>""),{rug},NA())'
            f[CH[f'ch drift used {el}']] = f'=IF(${used}{i}="",NA(),${used}{i})'
            f[CH[f'ch drift out {el}']] = (f'=IF(AND($D{i}="Drift",${used}{i}="",${linv}{i}<>""),'
                                           f'${linv}{i},NA())')
            f[CH[f'ch drift after {el}']] = f'=IF(AND($D{i}="Drift",${driftv}{i}<>""),${driftv}{i},NA())'
            f[CH[f'ch cal y {el}']] = (f'=IF(AND(OR($D{i}="Sample",$D{i}="Check"),${calv}{i}<>""),'
                                       f'${calv}{i},NA())')
            f[CH[f'ch acc min {el}']] = f'=IF(OR($B{i}="",{accmin}=""),NA(),{accmin})'
            f[CH[f'ch acc max {el}']] = f'=IF(OR($B{i}="",{accmax}=""),NA(),{accmax})'
            f[CH[f'ch pulse {el}']] = f'=IF(${pulse}{i}="",NA(),${pulse}{i})'
        f[CH['ch bgd']] = f'=IF($BI{i}="",NA(),$BI{i})'
        # The background under the sample peak is in every export, filtered or not.
        # It also carries what a large sample before it left in the system (on
        # 2026-09-22 the three highest values followed the three largest nitrogen
        # peaks, and the largest itself read 23.1 against a median of 23.2), so the
        # drift monitors, all one size, trace the instrument unless one follows a
        # large sample. See README, "Carryover after a large sample".
        f[CH['ch bgd sample']] = f'=IF($N{i}="",NA(),$N{i})'
        f[CH['ch bgd monitor']] = f'=IF(AND($D{i}="Drift",$N{i}<>""),$N{i},NA())'
        f[CH['mon bgd']] = f'=IF(AND($D{i}="Drift",$N{i}<>""),$N{i},"")'
        f[CH['mon bgd x']] = f'=IF(${CH["mon bgd"]}{i}="","",$A{i})'
        # Uncertainty of %N, %C, and C:N (1.4.3). Fits rows 301 to 335 hold the run's terms.
        UXc = UX
        f[UXc['u%N cert']] = stdpull(i, SUPN)
        f[UXc['u%C cert']] = stdpull(i, SUPC)
        for el, certc, ucol in (('N', 'U', 'u%N cert'), ('C', 'V', 'u%C cert')):
            f[UXc[f'cert rel {el}']] = (f'=IF(OR($B{i}="",${certc}{i}="",${certc}{i}=0),"",'
                                        f'IF(${UXc[ucol]}{i}="",0,${UXc[ucol]}{i}/${certc}{i}))')
        for el in ('N', 'C'):
            X_, F_ = EX[f'elem x {el}'], EX[f'elem f {el}']
            XR, FR = W(X_), W(F_)
            prev = f'COUNTIF({XR},"<"&$A{i})'
            f[UXc[f'd2 {el}']] = (f'=IF(${X_}{i}="","",IF({prev}=0,"",'
                                  f'(${F_}{i}-INDEX({FR},MATCH(SMALL({XR},{prev}),{XR},0)))^2))')
            f[UXc[f'w2 {el}']] = f'=IF(${X_}{i}="","",IF(OR({S_BAL}="",$G{i}="",$G{i}=0),0,({S_BAL}/$G{i})^2))'
        f[UXc['cn x']] = (f'=IF(AND($D{i}="Drift",ISNUMBER($AK{i}),ISNUMBER($U{i}),ISNUMBER($V{i}),'
                          f'${EX["excl E"]}{i}<>1),IF(AND($U{i}>0,$V{i}>0),$A{i},""),"")')
        f[UXc['cn f']] = f'=IF(${UXc["cn x"]}{i}="","",$AK{i}/(($V{i}/{AMC})/($U{i}/{AMN})))'
        CNXR, CNFR = W(UXc['cn x']), W(UXc['cn f'])
        prev = f'COUNTIF({CNXR},"<"&$A{i})'
        f[UXc['cn d2']] = (f'=IF(${UXc["cn x"]}{i}="","",IF({prev}=0,"",'
                           f'(${UXc["cn f"]}{i}-INDEX({CNFR},MATCH(SMALL({CNXR},{prev}),{CNXR},0)))^2))')
        for name, XR in (('mult N', W(EX['elem x N'])), ('mult C', W(EX['elem x C'])), ('mult CN', CNXR)):
            k = f'COUNTIF({XR},"<="&$A{i})'
            w = f'(($A{i}-SMALL({XR},{k}))/(SMALL({XR},{k}+1)-SMALL({XR},{k})))'
            f[UXc[name]] = (f'=IF($B{i}="","",IF(COUNT({XR})=0,1,IF(OR($A{i}<=MIN({XR}),$A{i}>=MAX({XR})),1,'
                            f'(1-{w})^2+{w}^2)))')
        for el, ug, nrow, brow, sdc, mxc, sxxc in (('N', 'AG', 69, 70, 312, 313, 314),
                                                   ('C', 'AH', 75, 76, 325, 326, 327)):
            f[UXc[f'kfit {el}']] = (f'=IF(OR(${ug}{i}="",Fits!$B${nrow}<3,Fits!$B${sdc}="",Fits!$B${sxxc}="",'
                                    f'Fits!$B${sxxc}=0,Fits!$B${brow}=0),"",IF(${ug}{i}<=0,"",'
                                    f'(Fits!$B${sdc}/ABS(Fits!$B${brow}))*SQRT(1/Fits!$B${nrow}'
                                    f'+(${ug}{i}-Fits!$B${mxc})^2/Fits!$B${sxxc})/${ug}{i}))')
        for el, inforce, rep in (('N', 246, 304), ('C', 249, 317)):
            f[UXc[f'cal2 {el}']] = (f'=IF($B{i}="","",IF(Fits!$B${inforce}="Yes",'
                                    f'IF(Fits!$B${rep}="","",${UXc[f"mult {el}"]}{i}*Fits!$B${rep}^2),'
                                    f'IF(${UXc[f"kfit {el}"]}{i}="","",${UXc[f"kfit {el}"]}{i}^2)))')
        f[UXc['cal2 CN']] = (f'=IF($B{i}="","",IF(AND(Fits!$B$246="Yes",Fits!$B$249="Yes"),'
                             f'IF(Fits!$B$330="","",${UXc["mult CN"]}{i}*Fits!$B$330^2),'
                             f'IF(OR(${UXc["cal2 N"]}{i}="",${UXc["cal2 C"]}{i}=""),"",'
                             f'${UXc["cal2 N"]}{i}+${UXc["cal2 C"]}{i})))')
        for el, inforce, X_, kcol in (('N', 246, EX['elem x N'], 'AY'), ('C', 249, EX['elem x C'], 'BA')):
            cr = f'N(${UXc[f"cert rel {el}"]}{i})^2'
            f[UXc[f'scale2 {el}']] = (f'=IF(Fits!$B${inforce}="Yes",IF(${X_}{i}="","",{cr}),'
                                      f'IF(${kcol}{i}="","",{cr}))')
        for el, val, certc in (('N', EX['%N corrected'], 'U'), ('C', EX['%C corrected'], 'V')):
            f[UXc[f'adev {el}']] = (f'=IF(AND($D{i}="Anchor",ISNUMBER(${val}{i}),ISNUMBER(${certc}{i})),'
                                    f'IF(${certc}{i}>0,${val}{i}/${certc}{i}-1,""),"")')
            f[UXc[f'acert2 {el}']] = f'=IF(${UXc[f"adev {el}"]}{i}="","",N(${UXc[f"cert rel {el}"]}{i})^2)'
        CNC = EX['C:N molar corrected']
        CERTCN = f'(($V{i}/{AMC})/($U{i}/{AMN}))'
        f[UXc['adev CN']] = (f'=IF(AND($D{i}="Anchor",ISNUMBER(${CNC}{i}),ISNUMBER($U{i}),ISNUMBER($V{i})),'
                             f'IF(AND($U{i}>0,$V{i}>0),${CNC}{i}/{CERTCN}-1,""),"")')
        f[UXc['acert2 CN']] = (f'=IF(${UXc["adev CN"]}{i}="","",N(${UXc["cert rel N"]}{i})^2'
                               f'+N(${UXc["cert rel C"]}{i})^2)')
        for el, val, inst, scale, true in (('N', EX['%N corrected'], 306, 307, 311),
                                           ('C', EX['%C corrected'], 319, 320, 324)):
            f[UXc[f'urel {el}']] = (f'=IF(OR(${val}{i}="",Fits!$B${inst}="",$G{i}="",$G{i}=0,'
                                    f'${UXc[f"cal2 {el}"]}{i}=""),"",SQRT(Fits!$B${inst}^2'
                                    f'+IF({S_BAL}="",0,({S_BAL}/$G{i})^2)+${UXc[f"cal2 {el}"]}{i}'
                                    f'+Fits!$B${scale}^2+IF(Fits!$B${true}="",0,Fits!$B${true}^2)))')
        f[UXc['urel CN']] = (f'=IF(OR(${CNC}{i}="",Fits!$B$330="",${UXc["cal2 CN"]}{i}=""),"",'
                             f'SQRT(Fits!$B$330^2+${UXc["cal2 CN"]}{i}+Fits!$B$331^2'
                             f'+IF(Fits!$B$335="",0,Fits!$B$335^2)))')
        for out, val, rel in (('%N', EX['%N corrected'], 'urel N'), ('%C', EX['%C corrected'], 'urel C'),
                              ('CN molar', EX['C:N molar corrected'], 'urel CN'),
                              ('CN mass', EX['C:N mass corrected'], 'urel CN')):
            f[UXc[f'u {out}']] = f'=IF(${UXc[rel]}{i}="","",${val}{i}*${UXc[rel]}{i})'
            f[UXc[f'U {out}']] = f'=IF(${UXc[f"u {out}"]}{i}="","",{S_K}*${UXc[f"u {out}"]}{i})'
        for el, val, certc, Ucol in (('N', EX['%N corrected'], 'U', 'U %N'), ('C', EX['%C corrected'], 'V', 'U %C')):
            f[UXc[f'inU {el}']] = (f'=IF(AND(OR($D{i}="Anchor",$D{i}="Linearity"),ISNUMBER(${UXc[Ucol]}{i}),'
                                   f'ISNUMBER(${certc}{i})),IF(ABS(${val}{i}-${certc}{i})<=${UXc[Ucol]}{i},1,0),"")')
        f[UXc['inU CN']] = (f'=IF(AND(OR($D{i}="Anchor",$D{i}="Linearity"),ISNUMBER(${UXc["U CN molar"]}{i}),'
                            f'ISNUMBER($U{i}),ISNUMBER($V{i})),IF(AND($U{i}>0,$V{i}>0),'
                            f'IF(ABS(${CNC}{i}-{CERTCN})<=${UXc["U CN molar"]}{i},1,0),""),"")')
        # Rerun marks (1.4.4). Fits rows 339 to 356 hold the run's numbers.
        R_ = lambda name: f'${RX[name]}{i}'      # noqa: E731
        # 1.4.12 (finding W-09): CO2 peaks before the designated pulse; with two or more, Flags says so
        f[RX['c before']] = (f'=IF(OR($B{i}="",NOT(ISNUMBER($BU{i}))),0,IF($BU{i}=9999,0,'
                             f'SUMPRODUCT(({rawA}=$A{i})*({rawR}="C")*({rawD}<$BU{i}))))')
        f[RX['fl cpre']] = (f'=IF($B{i}="","",IF(AND($P{i}<>"",N({R_("c before")})>=2),'
                            f'{R_("c before")}&" CO2 peaks before the reference pulse, the last one used; ",""))')
        # 1.4.10: a position's pulse furthest from the run's median, of all its pulses (Raw columns U to AB);
        # of two equally far, the one above it
        for el, val, dev, med in (('N', 'U', 'Y', '$B$340'), ('C', 'W', 'AA', '$B$344')):
            sel = f'(Raw!$A${RA}:$A${RB}=$A{i})'
            f[RX[f'np {el}']] = f'=IF($B{i}="","",COUNTIFS(Raw!$A${RA}:$A${RB},$A{i},Raw!${val}${RA}:${val}${RB},">-1E+30"))'
            f[RX[f'pmax {el}']] = f'=IF(OR(N({R_(f"np {el}")})=0,Fits!{med}=""),"",SUMPRODUCT(MAX({sel}*Raw!${dev}${RA}:${dev}${RB})))'
            f[RX[f'pmin {el}']] = f'=IF({R_(f"pmax {el}")}="","",SUMPRODUCT(MIN({sel}*Raw!${dev}${RA}:${dev}${RB})))'
            f[RX[f'pdev {el}']] = (f'=IF({R_(f"pmax {el}")}="","",IF({R_(f"pmax {el}")}>=-{R_(f"pmin {el}")},'
                                   f'{R_(f"pmax {el}")},{R_(f"pmin {el}")}))')
        for el, med, lim in (('N', '$B$340', '$B$342'), ('C', '$B$344', '$B$346')):
            f[RX[f'apdev {el}']] = f'=IF({R_(f"pdev {el}")}="","",ABS({R_(f"pdev {el}")}))'
            f[RX[f'poff {el}']] = (f'=IF(OR({R_(f"pdev {el}")}="",Fits!{lim}=""),0,'
                                   f'IF(ABS({R_(f"pdev {el}")})>Fits!{lim},1,0))')
        f[RX['fl noN']] = f'=IF($B{i}="","",IF($L{i}="","no N peak; ",""))'
        f[RX['fl noC']] = f'=IF($B{i}="","",IF($P{i}="","no C peak; ",""))'
        f[RX['fl rejC']] = f'=IF($B{i}="","",IF(${EX["rej C"]}{i}=1,"CO2 peak rejected as a reference pulse; ",""))'
        for el, a, lo, hi in (('N', 'L', '81', '97'), ('C', 'P', '82', '104')):
            f[RX[f'fl size{el}']] = (f'=IF($B{i}="","",IF(${a}{i}="","",IF(Fits!$B${lo}<=0,"",'
                                     f'IF(${a}{i}<Fits!$B${lo},"{el} below linearity range; ",'
                                     f'IF(Fits!$B${hi}<=0,"",IF(${a}{i}>Fits!$B${hi},"{el} above linearity range; ",""))))))')
        for el, name, got, acc in (('N', 'd15N', 'AA', 'S'), ('C', 'd13C', 'AF', 'T')):
            f[RX[f'fl check{el}']] = (f'=IF($B{i}="","",IF($D{i}<>"Check","",IF(${got}{i}="","",IF(${acc}{i}="","",'
                                      f'IF(ABS(${got}{i}-${acc}{i})<={S_THR},"",'
                                      f'"check {name} off by "&TEXT(${got}{i}-${acc}{i},"+0.00;-0.00")&"; ")))))')
        for el, a, blank in (('N', 'L', '6'), ('C', 'P', '10')):
            f[RX[f'fl blank{el}']] = (f'=IF($B{i}="","",IF($D{i}="Blank","",IF(${a}{i}="","",IF({S_MULT}="","",'
                                      f'IF(Fits!$B${blank}<=0,"",IF(${a}{i}>=Fits!$B${blank}*{S_MULT},"",'
                                      f'"{el} peak only "&TEXT(${a}{i}/Fits!$B${blank},"0.0")&"x the blank; "))))))')
        f[RX['fl settle']] = f'=IF($B{i}="","",IF($A{i}>=Fits!$B$86,"","ran before the background settled; "))'
        f[RX['fl carry']] = (f'=IF($B{i}="","",IF(${HX["carry flag"]}{i}=1,'
                             f'"possible carryover from an oversized sample; ",""))')
        for el, name in (('N', 'N2'), ('C', 'CO2')):
            f[RX[f'fl pulse{el}']] = (f'=IF($B{i}="","",IF({R_(f"poff {el}")}=1,"{name} reference pulse off by "'
                                      f'&TEXT({R_(f"pdev {el}")},"+0.00;-0.00")&"; ",""))')
        for el, gas, area, ampl in (('N', 'N2', 'L', 'M'), ('C', 'CO2', 'P', 'Q')):
            f[RX[f'fl ampl{el}']] = (f'=IF($B{i}="","",IF(NOT(ISNUMBER({S_AMAX})),"",IF(${area}{i}="","",'
                                     f'IF(NOT(ISNUMBER(${ampl}{i})),"",IF(${ampl}{i}>={S_AMAX},'
                                     f'"{gas} peak at or above "&TEXT({S_AMAX},"0")&" mV; ","")))))')
        for el, name, ucol in (('N', 'd15N', 'CL'), ('C', 'd13C', 'CN')):
            f[RX[f'fl u{el}']] = (f'=IF($B{i}="","",IF(NOT(ISNUMBER({S_USIG})),"",IF(NOT(ISNUMBER(${ucol}{i})),"",'
                                  f'IF(${ucol}{i}>{S_USIG},"1σ({name}) "&TEXT(${ucol}{i},"0.00")&" above "'
                                  f'&TEXT({S_USIG},"0.00")&"; ",""))))')
            f[RX[f'why size{el}']] = (f'=IF(OR(LEFT({R_(f"fl size{el}")},7)="{el} below",NOT(ISNUMBER({S_AMAX}))),'
                                      f'{R_(f"fl size{el}")},"")')
        f[RX['chk off']] = f'=IF(AND($B{i}<>"",OR({R_("fl checkN")}<>"",{R_("fl checkC")}<>"")),1,0)'
        if n == 0:
            f[RX['chk list']] = f'=IF({R_("chk off")}=1,$A{i}&"","")'
        else:
            prev = f'${RX["chk list"]}{i-1}'
            f[RX['chk list']] = f'={prev}&IF({R_("chk off")}=1,IF({prev}="","",", ")&$A{i},"")'
        # 1.4.10: the yield per mg above the blank (Fits B6 and B10), since a capsule of m mg gives the
        # blank's area plus m times that yield
        for el, a, loa, hia, bl in (('N', 'L', '$B$347', '$B$348', '$B$6'), ('C', 'P', '$B$349', '$B$350', '$B$10')):
            per = R_(f'per mg {el}')
            ab = f'MAX(Fits!{bl},0)'
            f[RX[f'per mg {el}']] = (f'=IF(AND(ISNUMBER(${a}{i}),ISNUMBER($G{i})),'
                                     f'IF(AND(${a}{i}-{ab}>0,$G{i}>0),(${a}{i}-{ab})/$G{i},""),"")')
            f[RX[f'lo {el}']] = (f'=IF(OR({per}="",Fits!{loa}-{ab}<=0,Fits!{hia}-{ab}<=0,Fits!{loa}>Fits!{hia}),"",'
                                 f'(Fits!{loa}-{ab})/{per})')
            f[RX[f'hi {el}']] = f'=IF({R_(f"lo {el}")}="","",(Fits!{hia}-{ab})/{per})'
        f[RX['mass lo']] = f'=IF(AND({R_("lo N")}="",{R_("lo C")}=""),"",MAX({R_("lo N")},{R_("lo C")}))'
        f[RX['mass hi']] = f'=IF(AND({R_("hi N")}="",{R_("hi C")}=""),"",MIN({R_("hi N")},{R_("hi C")}))'
        pos_ = n + 1
        if pos_ in REP_OF:
            g_, top_ = REP_OF[pos_], REP_FIRST[REP_OF[pos_]]
            f[RX['rep group']] = f'=IF(TRIM($B{i})="{lit(pos_)}",{g_},"")'
            for el, name, cnt, rng, lim, over in (('N', 'd15N', 'G', 'H', 'I', 'J'), ('C', 'd13C', 'K', 'L', 'M', 'N')):
                f[RX[f'fl rep{el}']] = (f'=IF({R_("rep group")}="","",IF(Replicates!${over}${top_}="yes",'
                                        f'"replicates at positions "&Replicates!$O${top_}&" differ by "'
                                        f'&TEXT(Replicates!${rng}${top_},"0.00")&" in {name}, above "'
                                        f'&TEXT(Replicates!${lim}${top_},"0.00")&"; ",""))')
        f[RX['rr N']] = (f'=IF($D{i}<>"Sample",0,IF(OR({R_("why sizeN")}<>"",{R_("fl blankN")}<>"",'
                         f'{R_("fl pulseN")}<>"",{R_("fl amplN")}<>"",{R_("fl uN")}<>"",{R_("fl repN")}<>""),1,0))')
        f[RX['rr C']] = (f'=IF($D{i}<>"Sample",0,IF(OR({R_("why sizeC")}<>"",{R_("fl blankC")}<>"",'
                         f'{R_("fl pulseC")}<>"",{R_("fl amplC")}<>"",{R_("fl uC")}<>"",{R_("fl repC")}<>""),1,0))')
        f[RX['rr both']] = (f'=IF($D{i}<>"Sample",0,IF(AND({R_("fl noN")}<>"",{R_("fl noC")}<>""),1,0))')
        # Finding R-23: under the Background model with the drift correction in force, a position with no
        # reference background (Y or AD empty) is marked for a check; from 1.4.14 (Tony Wang, 2026-10-16)
        # possible carryover is a check too, no longer a rerun of N and C
        nobg = (f'AND($X{i}<>"",$Y{i}="",Fits!$B$115="Yes")', f'AND($AC{i}<>"",$AD{i}="",Fits!$B$128="Yes")')
        f[RX['rr check']] = (f'=IF($D{i}<>"Sample",0,IF(OR({R_("fl settle")}<>"",{R_("fl carry")}<>"",'
                             f'AND({R_("fl noN")}<>"",{R_("fl noC")}=""),AND({R_("fl noN")}="",{R_("fl noC")}<>""),'
                             f'{nobg[0]},{nobg[1]}),1,0))')
        f[RX['rerun']] = (f'=IF($D{i}<>"Sample","",IF(OR({R_("rr both")}=1,AND({R_("rr N")}=1,{R_("rr C")}=1)),'
                          f'"N and C",IF({R_("rr N")}=1,"N",IF({R_("rr C")}=1,"C",IF({R_("rr check")}=1,"Check","")))))')
        why = [R_(k) for k in ('fl noN', 'why sizeN', 'fl noC', 'fl rejC', 'why sizeC')]
        why += [f'IF({c},"{t}; ","")' for c, t in zip(nobg, (NO_BGD_N, NO_BGD_C))]
        why += [R_(k) for k in ('fl blankN', 'fl blankC', 'fl settle', 'fl carry', 'fl pulseN', 'fl pulseC',
                                'fl amplN', 'fl amplC', 'fl uN', 'fl uC', 'fl repN', 'fl repC')]
        f[RX['why raw']] = f'=IF({R_("rerun")}="","",' + '&'.join(why) + ')'
        f[RX['why']] = f'=IF({R_("why raw")}="","",LEFT({R_("why raw")},LEN({R_("why raw")})-2))'
        # 1.4.10: a mass above every capsule weighed in this run (Fits B358) says so
        ml = R_("mass lo")
        beyond = (f'IF(AND(ISNUMBER(Fits!$B$358),{ml}-Fits!$B$358>1E-9*MAX(1,ABS(Fits!$B$358))),'
                  f'" (more than the heaviest capsule in this run)","")')
        f[RX['suggested mass']] = (
            f'=IF(OR({R_("rerun")}="",{R_("rerun")}="Check",{ml}=""),"",'
            f'IF({ml}<={R_("mass hi")},TEXT({ml},"0.000")&" to "&TEXT({R_("mass hi")},"0.000")&{beyond},'
            f'IF(AND({R_("lo N")}<>"",{R_("lo C")}<>""),"N "&TEXT({R_("lo N")},"0.000")&" to "'
            f'&TEXT({R_("hi N")},"0.000")&"; C "&TEXT({R_("lo C")},"0.000")&" to "&TEXT({R_("hi C")},"0.000")'
            f'&" (two capsules)"&{beyond},"none")))')
        f[RX['rerun rank']] = (f'=IF({R_("rerun")}="","",COUNTIF(${RX["rerun"]}${R0}:${RX["rerun"]}{i},"?*"))')
        # 1.4.13 (finding W-13): an anchor's C/N peak-area ratio, from both peaks above 0 (a CO2 peak set
        # aside as a reference pulse leaves P empty), and its raw d15N and d13C; Fits rows 373 to 388 read them
        okR, okN, okC = RX['am ok ratio'], RX['am ok d15N'], RX['am ok d13C']
        f[okR] = f'=IF(AND($D{i}="Anchor",ISNUMBER($L{i}),ISNUMBER($P{i})),IF(AND($L{i}>0,$P{i}>0),1,0),0)'
        f[RX['am ratio']] = f'=IF(${okR}{i}=1,$P{i}/$L{i},0)'
        f[okN] = f'=IF(AND($D{i}="Anchor",ISNUMBER($O{i})),1,0)'
        f[RX['am d15N']] = f'=IF(${okN}{i}=1,$O{i},0)'
        f[okC] = f'=IF(AND($D{i}="Anchor",ISNUMBER($R{i})),1,0)'
        f[RX['am d13C']] = f'=IF(${okC}{i}=1,$R{i},0)'
        # 1.4.13: the anchor's %N and %C as measured, over certified, less one, where adev holds the corrected
        # value's; Fits rows 393 to 404 compare the two
        f[RX['mdev N']] = f'=IF(ISNUMBER(${UX["adev N"]}{i}),$AI{i}/$U{i}-1,"")'
        f[RX['mdev C']] = f'=IF(ISNUMBER(${UX["adev C"]}{i}),$AJ{i}/$V{i}-1,"")'
        for name, col, width, kind in COLS:
            c = wk.cell(i, wk[col + '1'].column, f.get(col))
            c.font = FONTS[kind]
            if FILLS[kind]: c.fill = FILLS[kind]
            if col in NUMF: c.number_format = NUMF[col]
    wk.freeze_panes = 'C6'

    # ================================================================ Fits
    ft = wb.create_sheet('Fits')
    ft.sheet_view.showGridLines = False
    for col, w in zip('ABCD', [40, 14, 8, 62]):
        ft.column_dimensions[col].width = w
    ft.cell(1, 1, 'Fits').font = F_TITLE
    ft.cell(2, 1, 'Every coefficient below is fitted from this run\'s own standards.').font = F_NOTE

    def sec(r_, title):
        c = ft.cell(r_, 1, title); c.font = F_H1; c.fill = FILL_GRP
        ft.cell(r_, 2).fill = FILL_GRP; ft.cell(r_, 3).fill = FILL_GRP; ft.cell(r_, 4).fill = FILL_GRP

    def fit(r_, label, formula, fmt='0.0000', note=''):
        ft.cell(r_, 1, label).font = F_SM
        c = ft.cell(r_, 2, formula); c.font = F_SM; c.number_format = fmt; c.border = BOX
        if note: ft.cell(r_, 4, note).font = F_NOTE

    sec(4, 'Blank, from the blank capsules')
    fit(5,  'n with a nitrogen peak',       f'=COUNT({W("BC")})', '0')
    fit(6,  'mean Area 28',                 f'=IFERROR(AVERAGE({W("BD")}),0)', '0.000')
    fit(7,  'mean d15N (permil)',           f'=IFERROR(AVERAGE({W("BC")}),0)', '0.000')
    fit(8,  'SD of blank d15N (permil)',    f'=IFERROR(STDEV({W("BC")}),"")', '0.000',
            'A large SD here means the blank correction is not supported.')
    ft.cell(5, 3, '=IF(Setup!$B$10="Off","No",IF($B$5<1,"No",IF(Setup!$B$10="On","Yes",'
                    'IF(AND($B$5>=3,ISNUMBER($B$8),$B$8<=2),"Yes","No"))))').font = F_SM
    ft.cell(5, 3).border = BOX
    ft.cell(5, 4, 'Nitrogen blank correction in force?  On applies it whenever a blank exists; '
                  'Auto also requires three blanks and an SD of 2 permil or better.').font = F_NOTE
    fit(9,  'n with a carbon peak',          f'=COUNT({W("BE")})', '0')
    ft.cell(9, 3, '=IF(Setup!$B$10="Off","No",IF($B$9<1,"No",IF(Setup!$B$10="On","Yes",'
                    'IF(AND($B$9>=3,ISNUMBER($B$12),$B$12<=2),"Yes","No"))))').font = F_SM
    ft.cell(9, 3).border = BOX
    ft.cell(9, 4, 'Carbon blank correction in force?').font = F_NOTE
    fit(10, 'mean Area 44',                  f'=IFERROR(AVERAGE({W("BF")}),0)', '0.000')
    fit(11, 'mean d13C (permil)',            f'=IFERROR(AVERAGE({W("BE")}),0)', '0.000')
    fit(12, 'SD of blank d13C (permil)',     f'=IFERROR(STDEV({W("BE")}),"")', '0.000')

    sec(14, 'Linearity, nitrogen: d15N against Area 28')
    fit(15, 'n linearity standards',  f'=COUNT({W("AM")})', '0')
    fit(16, 'pivot, mean Area 28',    f'=IFERROR(AVERAGE({W("AM")}),0)', '0.000',
            'The pivot cancels in the calibration and does not affect the result.')
    fit(17, 'fitted slope (permil per area unit)', f'=IFERROR(SLOPE({W("AN")},{W("AM")}),0)')
    fit(18, 'r-squared',              f'=IFERROR(RSQ({W("AN")},{W("AM")}),"")', '0.000')
    fit(19, 'residual SD (permil)',   f'=IFERROR(STEYX({W("AN")},{W("AM")}),"")', '0.000')
    fit(20, 'slope applied',          f'=IF($B$96="Yes",$B$17,0)', '0.0000',
            'Zero when the correction is off, not supported by the data on Auto, or fewer than three '
            'standards were found. Row 92 onwards has the test.')

    sec(22, 'Linearity, carbon: d13C against Area 44')
    fit(23, 'n linearity standards',  f'=COUNT({W("AO")})', '0')
    fit(24, 'pivot, mean Area 44',    f'=IFERROR(AVERAGE({W("AO")}),0)', '0.000')
    fit(25, 'fitted slope',           f'=IFERROR(SLOPE({W("AP")},{W("AO")}),0)')
    fit(26, 'r-squared',              f'=IFERROR(RSQ({W("AP")},{W("AO")}),"")', '0.000')
    fit(27, 'residual SD (permil)',   f'=IFERROR(STEYX({W("AP")},{W("AO")}),"")', '0.000')
    fit(28, 'slope applied',          f'=IF($B$103="Yes",$B$25,0)')

    sec(30, 'Drift, nitrogen: linearity-corrected d15N of the drift standards')
    fit(31, 'n drift standards',              f'=COUNT({W("AS")})', '0')
    fit(32, 'slope against row',              f'=IFERROR(SLOPE({W("AS")},{W("AQ")}),0)')
    fit(33, 'r-squared against row',          f'=IFERROR(RSQ({W("AS")},{W("AQ")}),"")', '0.000')
    fit(34, 'residual SD against row',        f'=IFERROR(STEYX({W("AS")},{W("AQ")}),"")', '0.000')
    fit(35, 'slope against background',       f'=IFERROR(SLOPE({W("AS")},{W("AR")}),0)')
    fit(36, 'r-squared against background',   f'=IFERROR(RSQ({W("AS")},{W("AR")}),"")', '0.000')
    fit(37, 'residual SD against background', f'=IFERROR(STEYX({W("AS")},{W("AR")}),"")', '0.000')
    fit(38, 'model used', f'=IF({S_DM}="Row","Row",IF({S_DM}="Background","Background",'
                          f'IF(AND(ISNUMBER($B$37),ISNUMBER($B$34)),IF($B$37<$B$34,"Background","Row"),"Row")))',
            '@', 'On Auto, whichever model left the smaller residual SD. That is not an independent test, '
                 'because both models were fitted to these same standards. Judge by the check standards on QC.')
    fit(39, 'slope applied', f'=IF($B$115="Yes",IF($B$38="Row",$B$32,$B$35),0)')
    fit(40, 'pivot applied', '=$B$108', '0.000', 'Mean of the x values used in the fit.')

    sec(42, 'Drift, carbon: linearity-corrected d13C of the drift standards')
    fit(43, 'n drift standards',              f'=COUNT({W("AT")})', '0')
    fit(44, 'slope against row',              f'=IFERROR(SLOPE({W("AT")},{W("AQ")}),0)')
    fit(45, 'r-squared against row',          f'=IFERROR(RSQ({W("AT")},{W("AQ")}),"")', '0.000')
    fit(46, 'residual SD against row',        f'=IFERROR(STEYX({W("AT")},{W("AQ")}),"")', '0.000')
    fit(47, 'slope against background',       f'=IFERROR(SLOPE({W("AT")},{W("AR")}),0)',
            '0.0000', 'The mass-28 background stands in for a mass-44 background the export does not carry.')
    fit(48, 'r-squared against background',   f'=IFERROR(RSQ({W("AT")},{W("AR")}),"")', '0.000')
    fit(49, 'residual SD against background', f'=IFERROR(STEYX({W("AT")},{W("AR")}),"")', '0.000')
    fit(50, 'model used', f'=IF({S_DM}="Row","Row",IF({S_DM}="Background","Background",'
                          f'IF(AND(ISNUMBER($B$49),ISNUMBER($B$46)),IF($B$49<$B$46,"Background","Row"),"Row")))', '@')
    fit(51, 'slope applied', f'=IF($B$128="Yes",IF($B$50="Row",$B$44,$B$47),0)')
    fit(52, 'pivot applied', '=$B$121', '0.000')

    # 1.4.10: the scale needs anchors of two materials with different accepted values (column C);
    # with one, the line is flat and would give every sample that material's value.
    for top, el, x_, y_, lo_, hi_ in ((54, 'nitrogen', 'AU', 'AV', 138, 139), (61, 'carbon', 'AW', 'AX', 158, 159)):
        sec(top, f'Calibration, {el}: accepted regressed on measured')
        n_ = top + 1
        ok = f'$C${n_}="Yes"'
        fit(n_, 'n anchor measurements', f'=COUNT({W(x_)})', '0')
        ft.cell(n_, 3, f'=IF(AND($B${n_}>=2,ISNUMBER($B${lo_}),ISNUMBER($B${hi_})),IF($B${hi_}>$B${lo_},"Yes","No"),"No")'
                ).font = F_SM
        ft.cell(n_, 3).border = BOX
        ft.cell(n_, 4, 'Calibrated? Needs anchors of two materials with different accepted values (rows '
                       f'{lo_} and {hi_}); otherwise no value on this scale is reported (1.4.10).').font = F_NOTE
        fit(n_ + 1, 'slope',      f'=IF({ok},IFERROR(SLOPE({W(y_)},{W(x_)}),""),"")')
        fit(n_ + 2, 'intercept',  f'=IF({ok},IFERROR(INTERCEPT({W(y_)},{W(x_)}),""),"")')
        fit(n_ + 3, 'r-squared',  f'=IF({ok},IFERROR(RSQ({W(y_)},{W(x_)}),""),"")', '0.00000')
        fit(n_ + 4, 'residual SD (permil)', f'=IF({ok},IFERROR(STEYX({W(y_)},{W(x_)}),""),"")', '0.000',
            'Applied as accepted = slope x measured + intercept, never inverted.' if el == 'nitrogen' else '')

    sec(68, 'Elemental k-factor, nitrogen: Area 28 against micrograms N')
    fit(69, 'n standards used', f'=COUNT({W("AY")})', '0')
    fit(70, 'slope, area per ug N', f'=IFERROR(SLOPE({W("AZ")},{W("AY")}),"")', '0.00000')
    fit(71, 'intercept, area blank', f'=IFERROR(INTERCEPT({W("AZ")},{W("AY")}),"")', '0.000',
            'Subtracted from every peak, so the k-factor is not forced through the origin.')
    fit(72, 'r-squared', f'=IFERROR(RSQ({W("AZ")},{W("AY")}),"")', '0.0000')

    sec(74, 'Elemental k-factor, carbon: Area 44 against micrograms C')
    fit(75, 'n standards used', f'=COUNT({W("BA")})', '0')
    fit(76, 'slope, area per ug C', f'=IFERROR(SLOPE({W("BB")},{W("BA")}),"")', '0.00000')
    fit(77, 'intercept, area blank', f'=IFERROR(INTERCEPT({W("BB")},{W("BA")}),"")', '0.000')
    fit(78, 'r-squared', f'=IFERROR(RSQ({W("BB")},{W("BA")}),"")', '0.0000')

    sec(80, 'Size limits used by the flags')
    fit(81, 'minimum trusted Area 28', f'=IF({S_MINN}<>"",{S_MINN},IFERROR(MIN({W("AM")}),0))', '0.000',
            'Defaults to the smallest linearity standard in the run.')
    fit(82, 'minimum trusted Area 44', f'=IF({S_MINC}<>"",{S_MINC},IFERROR(MIN({W("AO")}),0))', '0.000')

    sec(84, 'Instrument settling, from the mass-28 background on the reference peaks')
    BIR = W('BI')
    fit(83, 'background floor',
        f'=IFERROR(AVERAGE(SMALL({BIR},1),SMALL({BIR},2),SMALL({BIR},3),SMALL({BIR},4),SMALL({BIR},5)),0)', '0.00')
    ft.cell(83, 4, 'Mean of the five lowest reference-peak backgrounds in the run.').font = F_NOTE
    fit(85, 'starting excess above the floor', f'=IFERROR(MAX({W(HX["ref first5"])}),0)-$B$83', '0.00',
            'Highest of the first five reference-peak backgrounds, less the floor. Under 1 mV the run '
            'counts as settled from the start.')
    fit(86, 'settled from position',
        f'=IF(OR({S_SETTLE}="",{S_SETTLE}=0,$B$85<1,$B$235<1),1,IFERROR(INDEX({W(HX["settle start"])},'
        f'MATCH(3,{W("CP")},0)),MAX({W(HX["pos used"])})+1))', '0',
        'First of three readings in a row at or under Setup B26, skipping readings within two positions '
        'of an oversized sample. A later spike does not move it. One past the last position means the '
        'background never settled. From 1.4.2 a start within 1 mV of the run\'s median background '
        '(row 235) counts as settled from the start, whatever the floor.')
    fit(87, 'drift slope before that position', f'=IFERROR(SLOPE({W("BL")},{W("BK")}),"")')
    fit(88, 'n before',                        f'=COUNT({W("BK")})', '0')
    fit(89, 'drift slope from that position',  f'=IFERROR(SLOPE({W("BN")},{W("BM")}),"")')
    fit(90, 'n from',                          f'=COUNT({W("BM")})', '0')
    ft.cell(90, 4, 'A run that settled shows a steep slope before and a flat one after.').font = F_NOTE

    # ---- is each correction supported by this run? ------------------------
    def inforce(r_, switch, nrow, prow, minimum=3):
        return (f'=IF({switch}="Off","No",IF($B${nrow}<{minimum},"No",IF({switch}="On","Yes",'
                f'IF(AND(ISNUMBER($B${prow}),$B${prow}<{S_ALPHA}),"Yes","No"))))')
    sec(92, 'Linearity, nitrogen: is the slope real?')
    fit(93, 'SE of slope',        f'=IFERROR(STEYX({W("AN")},{W("AM")})/SQRT(DEVSQ({W("AM")})),"")', '0.00000')
    fit(94, 't',                  '=IFERROR($B$17/$B$93,"")', '0.00')
    fit(95, 'p, two-sided',       '=IFERROR(TDIST(ABS($B$94),$B$15-2,2),"")', '0.000',
            'Probability of a slope this large from a flat response. Auto applies the slope only below the '
            'significance level in Setup.')
    fit(96, 'in force',           inforce(96, S_LIN, 15, 95), '@')
    fit(97, 'maximum trusted Area 28', f'=IF({S_MAXN}<>"",{S_MAXN},IFERROR(MAX({W("AM")}),0))', '0.000',
            'Defaults to the largest linearity standard in the run.')
    fit(98, 'chart rug level', f'=IFERROR(MIN({W("AN")})-MAX(0.05,0.15*(MAX({W("AN")})-MIN({W("AN")}))),"")',
        '0.000', 'Where the QC linearity chart draws the unknowns\' sizes, just under the standards.')

    sec(99, 'Linearity, carbon: is the slope real?')
    fit(100, 'SE of slope',       f'=IFERROR(STEYX({W("AP")},{W("AO")})/SQRT(DEVSQ({W("AO")})),"")', '0.00000')
    fit(101, 't',                 '=IFERROR($B$25/$B$100,"")', '0.00')
    fit(102, 'p, two-sided',      '=IFERROR(TDIST(ABS($B$101),$B$23-2,2),"")', '0.000')
    fit(103, 'in force',          inforce(103, S_LINC, 23, 102), '@')
    fit(104, 'maximum trusted Area 44', f'=IF({S_MAXC}<>"",{S_MAXC},IFERROR(MAX({W("AO")}),0))', '0.000')
    fit(105, 'chart rug level', f'=IFERROR(MIN({W("AP")})-MAX(0.05,0.15*(MAX({W("AP")})-MIN({W("AP")}))),"")', '0.000')

    sec(106, 'Drift, nitrogen: is the slope real?')
    fit(107, 'n monitors used',            f'=COUNT({W("BZ")})', '0',
            'Drift standards from the first position in Setup B14, minus any listed in Setup B16.')
    fit(108, 'mean x used',                f'=IFERROR(AVERAGE({W("BZ")}),0)', '0.000')
    fit(109, 'Sxx',                        f'=IFERROR(DEVSQ({W("BZ")}),0)', '0.000')
    fit(110, 'intercept, model in use',    f'=IFERROR(INTERCEPT({W("AS")},{W("BZ")}),"")')
    fit(111, 'residual SD, model in use',  f'=IFERROR(STEYX({W("AS")},{W("BZ")}),"")', '0.000')
    fit(112, 'SE of slope',                '=IFERROR($B$111/SQRT($B$109),"")', '0.00000')
    fit(113, 't',                          '=IFERROR(IF($B$38="Row",$B$32,$B$35)/$B$112,"")', '0.00')
    fit(114, 'p, two-sided',               '=IFERROR(TDIST(ABS($B$113),$B$107-2,2),"")', '0.000')
    fit(115, 'in force',                   inforce(115, S_DRIFT, 107, 114), '@')
    fit(116, 'residual sum of squares',    '=IFERROR($B$111^2*($B$107-2),"")', '0.0000')
    fit(117, 'monitors flagged as outliers',
        f'=COUNTIFS({W("CB")},">"&{S_TOUT},{W("BX")},0)+COUNTIFS({W("CB")},"<"&-{S_TOUT},{W("BX")},0)', '0',
        'Each monitor against the line through the others, in units of their scatter (Work column CB).')

    sec(119, 'Drift, carbon: is the slope real?')
    fit(120, 'n monitors used',            f'=COUNT({W("CA")})', '0')
    fit(121, 'mean x used',                f'=IFERROR(AVERAGE({W("CA")}),0)', '0.000')
    fit(122, 'Sxx',                        f'=IFERROR(DEVSQ({W("CA")}),0)', '0.000')
    fit(123, 'intercept, model in use',    f'=IFERROR(INTERCEPT({W("AT")},{W("CA")}),"")')
    fit(124, 'residual SD, model in use',  f'=IFERROR(STEYX({W("AT")},{W("CA")}),"")', '0.000')
    fit(125, 'SE of slope',                '=IFERROR($B$124/SQRT($B$122),"")', '0.00000')
    fit(126, 't',                          '=IFERROR(IF($B$50="Row",$B$44,$B$47)/$B$125,"")', '0.00')
    fit(127, 'p, two-sided',               '=IFERROR(TDIST(ABS($B$126),$B$120-2,2),"")', '0.000')
    fit(128, 'in force',                   inforce(128, S_DRIFT, 120, 127), '@')
    fit(129, 'residual sum of squares',    '=IFERROR($B$124^2*($B$120-2),"")', '0.0000')
    fit(130, 'monitors flagged as outliers',
        f'=COUNTIFS({W("CC")},">"&{S_TOUT},{W("BY")},0)+COUNTIFS({W("CC")},"<"&-{S_TOUT},{W("BY")},0)', '0')

    # ---- calibration statistics and the uncertainty budget ---------------
    def caltable(r0, hdr, xcol, onecol, devcol, rescol, sqcol, ucol_setup, urow):
        sec(r0 - 1, hdr)
        heads = ['Material', 'n', 'sum x', 'sum x^2', 'mean measured', 'SD', '1σ(mean)',
                 'S1', 'S2', '1σ(accepted)']
        for j, h in enumerate(heads, start=1):
            c = ft.cell(r0, j, h); c.font = F_HDR; c.fill = FILL_HDR
            c.alignment = Alignment(horizontal='center', wrap_text=True)
        for k in range(NSTD):
            i, sr = r0 + 1 + k, SFIRST + k
            vals = [
                f'=IF(Setup!$A${sr}="","",Setup!$A${sr})',
                f'=SUMIF({W("C")},{sr},{W(onecol)})',
                f'=SUMIF({W("C")},{sr},{W(xcol)})',
                f'=SUMIF({W("C")},{sr},{W(sqcol)})',
                f'=IF($B{i}>0,$C{i}/$B{i},"")',
                f'=IF($B{i}>1,SQRT(MAX(0,($D{i}-$C{i}^2/$B{i})/($B{i}-1))),0)',
                # 1.4.12 (finding W-10): a material analyzed once takes the repeatability in use
                f'=IF($B{i}>1,$F{i}/SQRT($B{i}),IF($B{i}=1,IF(ISNUMBER($B${urow}),$B${urow},0),0))',
                f'=SUMIF({W("C")},{sr},{W(devcol)})',
                f'=SUMIF({W("C")},{sr},{W(rescol)})',
                f'=IF(ISNUMBER({ucol_setup}${sr}),{ucol_setup}${sr},0)',
            ]
            for j, v in enumerate(vals, start=1):
                c = ft.cell(i, j, v); c.font = F_SM
                c.number_format = '0' if j == 2 else '0.0000'
        ft.cell(r0 + NSTD + 1, 1, 'One row per line of the standards table. n, mean and SD are of the anchor '
                            'replicates after drift correction; S1 and S2 are the sums the uncertainty '
                            'weights need; 1σ(accepted) is the certificate value from Setup. 1σ(mean) is SD over '
                            f'the square root of n, and for a material analyzed once the repeatability in use, row '
                            f'{urow} (1.4.12).').font = F_NOTE

    sec(132, 'Calibration, nitrogen: statistics behind the uncertainty')
    fit(133, 'mean measured, all anchors', f'=IFERROR(AVERAGE({W("AU")}),"")', '0.0000')
    fit(134, 'Sxx',                        f'=IFERROR(DEVSQ({W("AU")}),"")', '0.0000')
    fit(135, 'pooled SD of anchor replicates',
        f'=IF($C$55<>"Yes","",IFERROR(SQRT(SUMPRODUCT(($B${UT_N}:$B${UT_N+NSTD-1}-1)*($B${UT_N}:$B${UT_N+NSTD-1}>1)*$F${UT_N}:$F${UT_N+NSTD-1}^2)'
        f'/SUMPRODUCT(($B${UT_N}:$B${UT_N+NSTD-1}-1)*($B${UT_N}:$B${UT_N+NSTD-1}>1))),""))', '0.000',
        'Within-material scatter of the anchors, pooled. Stands in for the repeatability of a single '
        'sample analysis when Setup gives none.')
    fit(136, 'sample repeatability in use', f'=IF({S_USN}<>"",{S_USN},$B$135)', '0.000')
    fit(137, 'its source', f'=IF({S_USN}<>"","Setup","pooled anchor replicates")', '@')
    fit(138, 'accepted d15N, lowest anchor',  f'=IFERROR(MIN({W("AV")}),"")', '0.00')
    fit(139, 'accepted d15N, highest anchor', f'=IFERROR(MAX({W("AV")}),"")', '0.00',
            'A sample outside this range is extrapolated beyond the anchors and is flagged.')
    caltable(UT_N - 1, f'Uncertainty table, nitrogen (rows {UT_N} to {UT_N + NSTD - 1})', 'AU', 'CD', 'CE', 'CF', 'CG', 'Setup!$I', 136)

    sec(152, 'Calibration, carbon: statistics behind the uncertainty')
    fit(153, 'mean measured, all anchors', f'=IFERROR(AVERAGE({W("AW")}),"")', '0.0000')
    fit(154, 'Sxx',                        f'=IFERROR(DEVSQ({W("AW")}),"")', '0.0000')
    fit(155, 'pooled SD of anchor replicates',
        f'=IF($C$62<>"Yes","",IFERROR(SQRT(SUMPRODUCT(($B${UT_C}:$B${UT_C+NSTD-1}-1)*($B${UT_C}:$B${UT_C+NSTD-1}>1)*$F${UT_C}:$F${UT_C+NSTD-1}^2)'
        f'/SUMPRODUCT(($B${UT_C}:$B${UT_C+NSTD-1}-1)*($B${UT_C}:$B${UT_C+NSTD-1}>1))),""))', '0.000')
    fit(156, 'sample repeatability in use', f'=IF({S_USC}<>"",{S_USC},$B$155)', '0.000')
    fit(157, 'its source', f'=IF({S_USC}<>"","Setup","pooled anchor replicates")', '@')
    fit(158, 'accepted d13C, lowest anchor',  f'=IFERROR(MIN({W("AX")}),"")', '0.00')
    fit(159, 'accepted d13C, highest anchor', f'=IFERROR(MAX({W("AX")}),"")', '0.00')
    caltable(UT_C - 1, f'Uncertainty table, carbon (rows {UT_C} to {UT_C + NSTD - 1})', 'AW', 'CH', 'CI', 'CJ', 'CK', 'Setup!$H', 156)

    # ---- working gas, from the pulses Isodat did not designate --------------
    sec(199, 'Reference pulses: did the working-gas measurement hold across the run?')
    ft.cell(200, 4, 'The designated reference pulse carries the assigned value by definition, so only the '
                    'other pulses can show movement. Their d values are relative to the designated pulse '
                    'of the same acquisition.').font = F_NOTE
    # 1.4.10: every pulse Isodat did not designate (Raw columns U to X), not only the first at each position
    for base, col, name, xcol, npc in ((201, 'U', 'N2', 'V', 'np N'), (208, 'W', 'CO2', 'X', 'np C')):
        R = f'Raw!${col}${RA}:${col}${RB}'; A_ = rawA
        fit(base,     f'{name} pulses read',            f'=COUNT({R})', '0')
        fit(base + 6, f'positions with such a {name} pulse', f'=COUNTIF({W(RX[npc])},">0")', '0')
        fit(base + 1, 'mean d, relative to the designated pulse', f'=IFERROR(AVERAGE({R}),"")', '0.000')
        fit(base + 2, 'SD (permil)',                     f'=IFERROR(STDEV({R}),"")', '0.000')
        fit(base + 3, 'range (permil)',                  f'=IF($B${base}<2,"",MAX({R})-MIN({R}))', '0.000')
        fit(base + 4, 'trend, permil per position',      f'=IFERROR(SLOPE({R},{A_}),"")', '0.00000')
        fit(base + 5, 'p of the trend, two-sided',
            f'=IFERROR(TDIST(ABS($B${base+4}/(STEYX({R},{A_})/SQRT(DEVSQ(Raw!${xcol}${RA}:${xcol}${RB})))),$B${base}-2,2),"")',
            '0.000')

    # ---- what the run offers as an independent check -------------------------
    sec(216, 'Check standards')
    sec(219, 'Sample-peak background on the drift monitors')
    ft.cell(220, 4, 'The mass-28 background under each monitor\'s own peak, from column BGD 28. Present in '
                    'every export; the reference-peak background above needs the reference peaks. A monitor '
                    'run right after a large sample can read high from what that sample left behind.').font = F_NOTE
    MB, MBX = W(CH['mon bgd']), W(CH['mon bgd x'])
    fit(221, 'monitors with a background', f'=COUNT({MB})', '0')
    fit(222, 'lowest',                     f'=IFERROR(MIN({MB}),"")', '0.0')
    fit(223, 'highest',                    f'=IFERROR(MAX({MB}),"")', '0.0')
    fit(224, 'trend, per position',        f'=IFERROR(SLOPE({MB},{MBX}),"")', '0.000')
    fit(225, 'p of the trend, two-sided',
        f'=IFERROR(TDIST(ABS($B$224/(STEYX({MB},{MBX})/SQRT(DEVSQ({MBX})))),$B$221-2,2),"")', '0.000')

    sec(227, 'Carryover after oversized samples')
    fit(228, 'rise that flags a position (mV)', f'=IF({S_CARRY}="","",{S_CARRY})', '0.0',
        'Setup B28. A position is exposed when one of the two before it is above the linearity range '
        '(B97, B104); its rise is its sample-peak background less the last one that was not exposed.')
    fit(229, 'positions above the linearity range', f'=SUM({W(HX["oversized"])})', '0')
    fit(230, 'of those, followed directly by a sample or check', f'=SUM({W(HX["after oversized"])})', '0')
    fit(231, 'positions flagged for possible carryover', f'=COUNTIF({W(HX["carry flag"])},1)', '0')
    fit(232, 'largest rise (mV)',
        f'=IF(COUNT({W(HX["carry rise"])})=0,"",MAX({W(HX["carry rise"])}))', '0.0')

    # 1.4.2: settling against the run's level, the reference-pulse check, and
    # the elemental drift correction.
    fit(234, 'run level, median reference-peak background', f'=IFERROR(MEDIAN({BIR}),0)', '0.00',
        'The middle reading of the run. A start within 1 mV of it was settled from the start.')
    fit(235, 'starting excess above the run level', f'=IFERROR(MAX({W(HX["ref first5"])}),0)-$B$234', '0.00')
    sec(236, 'Reference-pulse check for exports without reference peaks')
    NDES = f'COUNTIFS({rawR},"C",{rawI},1,{rawJ},">-1E+30",{rawO},">-1E+30")'
    fit(237, 'designated CO2 pulses in the export', f'={NDES}', '0',
        'Flagged mass-44 peaks anywhere in the export; an export cut to the sample peaks carries them '
        'only at positions with no sample gas.')
    fit(238, 'their mean amplitude', f'=IF($B$237=0,"",SUMIFS({rawJ},{rawR},"C",{rawI},1,{rawO},">-1E+30")/$B$237)', '0.0',
        'For the comparison each peak is held against the pulses interpolated in position (Work, pulse ampl).')
    fit(239, 'their mean d13C (permil)', f'=IF($B$237=0,"",SUMIFS({rawO},{rawR},"C",{rawI},1,{rawJ},">-1E+30")/$B$237)', '0.000')
    fit(240, 'check active', f'=IF(OR({S_RPA}="",$B$237=0,COUNT({BIR})>0,COUNT({W(EX["desig x"])})=0),"No","Yes")', '@',
        'Only in an export with no N2 reference peaks, with Setup B39 filled and at least one designated pulse.')
    fit(241, 'CO2 peaks rejected as reference pulses', f'=COUNTIF({W(EX["rej C"])},1)', '0')
    sec(243, 'Elemental drift, from the drift monitors')
    for base, el, name in ((244, 'N', 'nitrogen'), (247, 'C', 'carbon')):
        XR, FR = W(EX[f'elem x {el}']), W(EX[f'elem f {el}'])
        fit(base, f'n monitors used, {name}', f'=COUNT({XR})', '0',
            'Drift monitors with a mass, minus the positions in Setup B37.' if el == 'N' else '')
        fit(base + 1, f'largest relative bias, {name} (percent)',
            f'=IF($B${base}=0,"",MAX(MAX({FR})-1,1-MIN({FR}))*100)', '0.00',
            'Measured over certified, less one, of the monitor furthest off.' if el == 'N' else '')
        fit(base + 2, f'in force, {name}',
            f'=IF({S_ELEM}="Off","No",IF($B${base}<2,"No",IF({S_ELEM}="On","Yes",'
            f'IF(AND(ISNUMBER($B${base+1}),$B${base+1}>{S_ETHR}),"Yes","No"))))', '@',
            'Off, fewer than two monitors, On, or Auto against Setup B36.' if el == 'N' else '')
    fit(250, 'positions left out', f'=IF({S_EXE}="","none",{S_EXE})', '@')

    # 1.4.3: u of %N, %C, and C:N. Fractions of the value, 1 SD.
    sec(300, 'Uncertainty of %N, %C, and C:N (1.4.3), as fractions of the value, 1σ')
    fit(301, 'balance, 1σ of one weighing (mg)', f'=IF({S_BAL}="","",{S_BAL})', '0.0000',
        'Setup B33. Empty: the weighing term is left out, here and in the monitors\' repeatability.')
    for base, el, name, X_, ug, area in ((303, 'N', 'nitrogen', EX['elem x N'], 'AY', 'AZ'),
                                         (316, 'C', 'carbon', EX['elem x C'], 'BA', 'BB')):
        b = base
        fit(b, f'drift monitors, {name}', f'=COUNT({W(X_)})', '0',
            'The monitors of the elemental drift correction (Fits row 244 or 247).' if el == 'N' else '')
        fit(b + 1, f'repeatability, {name}',
            f'=IF($B${b}<3,"",SQRT(SUM({W(UX[f"d2 {el}"])})/(2*($B${b}-1))))', '0.00000',
            'Measured over certified, by successive differences in position: sqrt(sum d^2 / (2 (n - 1))), '
            'so a drift slow next to the monitors\' spacing is not counted as scatter.' if el == 'N' else '')
        fit(b + 2, f'monitors\' own weighing, {name}', f'=IF($B${b}=0,0,SQRT(SUM({W(UX[f"w2 {el}"])})/$B${b}))',
            '0.00000', 'RMS of the balance\'s 1σ over each monitor\'s mass.' if el == 'N' else '')
        fit(b + 3, f'instrument repeatability, {name}',
            f'=IF($B${b+1}="","",SQRT(MAX(0,$B${b+1}^2-$B${b+2}^2)))', '0.00000',
            'The repeatability less the monitors\' own weighing, so a sample\'s weighing is not counted twice.'
            if el == 'N' else '')
        fit(b + 4, f'certificate, material setting the scale, {name}',
            f'=IF(COUNT({W(UX[f"scale2 {el}"])})=0,0,SQRT(AVERAGE({W(UX[f"scale2 {el}"])})))', '0.00000',
            'RMS of 1σ(%N)/%N from Setup over the monitors when the correction is in force, otherwise over the '
            'weighed standards of the k-factor.' if el == 'N' else '')
        fit(b + 5, f'anchor analyses, {name}', f'=COUNT({W(UX[f"adev {el}"])})', '0')
        fit(b + 6, f'anchors\' RMS deviation from certified, {name}',
            f'=IF($B${b+5}<2,"",SQRT(SUMSQ({W(UX[f"adev {el}"])})/$B${b+5}))', '0.00000',
            'Reported over certified, less one; needs two anchor analyses.' if el == 'N' else '')
        fit(b + 7, f'anchors\' certificate, {name}',
            f'=IF($B${b+5}=0,"",SQRT(AVERAGE({W(UX[f"acert2 {el}"])})))', '0.00000')
        fit(b + 8, f'trueness, {name}', f'=IF($B${b+6}="","",SQRT($B${b+6}^2+N($B${b+7})^2))', '0.00000',
            'The Nordtest TR 537 bias term: RMS deviation and certificate combined.' if el == 'N' else '')
        fit(b + 9, f'k-factor residual SD, {name} (area)', f'=IFERROR(STEYX({W(area)},{W(ug)}),"")', '0.0000')
        fit(b + 10, f'k-factor mean micrograms, {name}', f'=IFERROR(AVERAGE({W(ug)}),"")', '0.000')
        fit(b + 11, f'k-factor Sxx, {name} (micrograms squared)', f'=IFERROR(DEVSQ({W(ug)}),"")', '0.000')
    fit(329, 'drift monitors, C:N', f'=COUNT({W(UX["cn x"])})', '0')
    fit(330, 'repeatability, C:N', f'=IF($B$329<3,"",SQRT(SUM({W(UX["cn d2"])})/(2*($B$329-1))))', '0.00000',
        'The monitors\' measured over certified C:N; weighing cancels in a ratio.')
    fit(331, 'certificate, materials setting the scale, C:N', '=SQRT($B$307^2+$B$320^2)', '0.00000')
    fit(332, 'anchor analyses, C:N', f'=COUNT({W(UX["adev CN"])})', '0')
    fit(333, 'anchors\' RMS deviation from certified, C:N',
        f'=IF($B$332<2,"",SQRT(SUMSQ({W(UX["adev CN"])})/$B$332))', '0.00000')
    fit(334, 'anchors\' certificate, C:N', f'=IF($B$332=0,"",SQRT(AVERAGE({W(UX["acert2 CN"])})))', '0.00000')
    fit(335, 'trueness, C:N', '=IF($B$333="","",SQRT($B$333^2+N($B$334)^2))', '0.00000')

    # 1.4.4: rerun marks
    sec(338, 'Rerun marks (1.4.4)')
    for base, col, adev, name in ((339, 'U', 'Z', 'N2'), (343, 'W', 'AB', 'CO2')):
        R = f'Raw!${col}${RA}:${col}${RB}'
        fit(base, f'{name} reference pulses read', f'=COUNT({R})', '0',
            'Every pulse Isodat did not designate (Raw columns U and W, 1.4.10).' if name == 'N2' else '')
        fit(base + 1, 'their median (permil)', f'=IF($B${base}<5,"",MEDIAN({R}))', '0.000',
            'Needs five pulses.')
        fit(base + 2, 'robust SD (permil)', f'=IF($B${base+1}="","",1.482602218505602*MEDIAN(Raw!${adev}${RA}:${adev}${RB}))',
            '0.0000', '1.4826 x the median absolute deviation from the median.')
        fit(base + 3, 'a pulse is off beyond (permil)',
            f'=IF(OR({S_POFF}="",N({S_POFF})<=0,$B${base+2}="",N($B${base+2})<=0),"",{S_POFF}*$B${base+2})',
            '0.000', 'Setup B67 x the robust SD.')
    fit(358, 'heaviest capsule in this run (mg)', f'=IF(COUNT({W("G")})=0,"",IF(MAX({W("G")})>0,MAX({W("G")}),""))',
        '0.000', 'A suggested mass above it says so (1.4.10).')
    fit(347, 'smallest Area 28 a rerun aims for',
        f'=MAX(MAX($B$81,0),IF(OR({S_MULT}="",$B$6<=0),0,{S_MULT}*$B$6))', '0.000',
        'The linearity range, and the blank multiple of Setup B25.')
    fit(348, 'largest Area 28 a rerun aims for', '=MAX($B$97,0)', '0.000')
    fit(349, 'smallest Area 44 a rerun aims for',
        f'=MAX(MAX($B$82,0),IF(OR({S_MULT}="",$B$10<=0),0,{S_MULT}*$B$10))', '0.000')
    fit(350, 'largest Area 44 a rerun aims for', '=MAX($B$104,0)', '0.000')
    MARK = W(RX['rerun'])
    fit(351, 'samples marked for a rerun',
        f'=COUNTIF({MARK},"N")+COUNTIF({MARK},"C")+COUNTIF({MARK},"N and C")', '0')
    fit(352, 'samples marked for a check', f'=COUNTIF({MARK},"Check")', '0')
    NCHK = f'COUNTIF({W(RX["chk off"])},1)'
    fit(353, 'whole run: check standards',
        f'=IF({NCHK}=0,"","check standard off by more than "&TEXT({S_THR},"0.00")&" permil at "'
        f'&IF({NCHK}=1,"position ","positions ")&Work!${RX["chk list"]}${R0+NSAMP-1}&"; ")', '@')
    parts = []
    for name, inu in (('%N', 'inU N'), ('%C', 'inU C'), ('C:N', 'inU CN')):
        within = f'COUNTIFS({W("D")},"Anchor",{W(UX[inu])},1)'
        total = f'COUNTIFS({W("D")},"Anchor",{W(UX[inu])},">=0")'
        parts.append(f'IF(AND({total}>0,{within}<{total}),"anchors outside "&{S_K}&"σ, {name}: "'
                     f'&TEXT({total}-{within},"0")&" of "&TEXT({total},"0")&"; ","")')
    fit(354, 'whole run: anchors outside kσ', '=' + '&'.join(parts), '@')
    fit(355, 'whole run: reference gas',
        '=' + '&'.join(f'IF(AND({S_RGSD}<>"",$B${b}>=3,ISNUMBER($B${b+2})),IF($B${b+2}>={S_RGSD},'
                       f'"{nm} reference pulses SD "&TEXT($B${b+2},"0.00")&" permil, at or above "'
                       f'&TEXT({S_RGSD},"0.00")&"; ",""),"")' for b, nm in ((201, 'N2'), (208, 'CO2'))), '@')
    # 1.4.10: an element the anchors cannot calibrate comes first
    cal_parts = []
    for n_, name, z_ in ((55, 'd15N', 'Z'), (62, 'd13C', 'AE')):
        cal_parts.append(f'IF($C${n_}="Yes","",IF($B${n_}>=1,"no {name} calibrated: every anchor carries the same '
                         f'accepted {name}, and the scale needs two materials with different values; ",'
                         f'IF(COUNTIFS({W("D")},"Sample",{W(z_)},">-1E+30")>0,"no {name} calibrated: no anchor has '
                         f'a measured {name}; ","")))')
    fit(357, 'whole run: calibration', '=' + '&'.join(cal_parts), '@')
    # 1.4.12 (Tony, 2026-10-14; finding W-02): the calibration line itself, its slope against 1 and the
    # anchors' residual SD about it, at the Note and Caution limits of Setup B73 to B76
    sec(360, 'Calibration line: slope against 1 and the anchors\' residual SD (1.4.12)')

    def cal_level(off_r, sd_r):
        over = lambda v, lim: f'AND({lim}<>"",ISNUMBER({v}),{v}>{lim})'          # noqa: E731
        off, sd = f'$B${off_r}', f'$B${sd_r}'
        return (f'=IF(NOT(ISNUMBER({off})),"",IF(OR({over(off, S_CSC)},{over(sd, S_CDC)}),"Caution",'
                f'IF(OR({over(off, S_CSN)},{over(sd, S_CDN)}),"Note","OK")))')
    calline = []
    for base, name, slope_r, sd_r, ok in ((361, 'd15N', 56, 59, '$C$55'), (364, 'd13C', 63, 66, '$C$62')):
        fit(base, f'{name}: slope, distance from 1',
            f'=IF(AND({ok}="Yes",ISNUMBER($B${slope_r})),ABS($B${slope_r}-1),"")', '0.0000')
        fit(base + 1, f'{name}: anchor residual SD (permil)',
            f'=IF(AND({ok}="Yes",ISNUMBER($B${sd_r})),$B${sd_r},"")', '0.000')
        fit(base + 2, f'{name}: calibration line', cal_level(base, base + 1), '@',
            'OK, Note or Caution, at the limits of Setup B73 to B76.' if name == 'd15N' else '')
        off, sd = f'$B${base}', f'$B${base + 1}'
        for v, note, caution, text in (
                (off, S_CSN, S_CSC, f'"{name} calibration slope "&TEXT($B${slope_r},"0.0000")&", more than "&LIM&" from 1; "'),
                (sd, S_CDN, S_CDC, f'"{name} anchor residual SD "&TEXT({sd},"0.000")&" permil, above "&LIM&"; "')):
            lim = f'IF(AND({caution}<>"",{v}>{caution}),{caution},{note})'
            calline.append(f'IF(NOT(ISNUMBER({v})),"",IF(OR(AND({caution}<>"",{v}>{caution}),AND({note}<>"",{v}>{note})),'
                           + text.replace('LIM', lim) + ',""))')
    fit(367, 'whole run: calibration line', '=' + '&'.join(calline), '@')

    def above(v, lim):
        """v above lim with the engine's slack (pipeline._above); FALSE when v is not a number (1.4.13)."""
        return f'IF(ISNUMBER({v}),{v}-{lim}>1E-9*MAX(1,ABS({lim})),FALSE)'

    # 1.4.13 (finding W-13): do the anchors of each row of the standards table read as one material? Analyses
    # of one material agree in their C/N peak-area ratio and their raw d15N and d13C; capsules of another
    # material whose names contain an anchor's match text do not.
    AM0 = 369
    sec(AM0, 'Anchors read as one material each (1.4.13)')
    AM_R, AM_S = f'$B${AM0 + 1}', f'$B${AM0 + 2}'
    fit(AM0 + 1, 'limit: largest over smallest C/N peak-area ratio', ANCHOR_RATIO_LIMIT, '0.00',
        'A fixed limit in this version (finding W-13): over 19 runs the quotient within one anchor material was '
        '1.003 to 1.145, and 4.10 where acetanilide capsules were labeled USGS 64 Lin, Drift, or Bypass.')
    fit(AM0 + 2, 'limit: spread of the raw d15N or d13C (permil)', ANCHOR_SPREAD_PERMIL, '0.0',
        'A fixed limit in this version: 0.03 to 0.6 permil in the same runs where the calibration held.')
    AMH, AMT = AM0 + 3, AM0 + 4
    for j, h in enumerate(['Material', 'n anchors', 'n with both peaks', 'ratio, smallest', 'ratio, largest',
                           'largest over smallest', 'd15N, lowest', 'd15N, highest', 'd15N spread', 'd13C, lowest',
                           'd13C, highest', 'd13C spread', 'level', 'problem of the whole run'], start=1):
        c = ft.cell(AMH, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    WC_ = W('C')

    def extreme(fn, ok, x, sr):
        """The smallest (MIN) or largest (MAX) x of the anchors of standards row sr where ok is 1. Every other
        position is set 1E30 beyond the values instead of being offset, so each value is taken exactly."""
        cond = f'({WC_}={sr})*{W(RX[ok])}'
        return f'SUMPRODUCT({fn}({cond}*{W(RX[x])}{"+" if fn == "MIN" else "-"}(1-{cond})*1E+30))'
    for k in range(NSTD):
        t, sr = AMT + k, SFIRST + k
        cnt = {q: f'SUMPRODUCT(({WC_}={sr})*{W(RX["am ok " + q])})' for q in ('ratio', 'd15N', 'd13C')}
        rx_, nsp, csp = f'$F{t}', f'$I{t}', f'$L{t}'
        vals = [
            f'=IF(Setup!$A${sr}="","",Setup!$A${sr})',
            f'=COUNTIFS({WC_},{sr},{W("D")},"Anchor")',
            f'={cnt["ratio"]}',
            f'=IF($C{t}=0,"",{extreme("MIN", "am ok ratio", "am ratio", sr)})',
            f'=IF($C{t}=0,"",{extreme("MAX", "am ok ratio", "am ratio", sr)})',
            f'=IF($C{t}<2,"",$E{t}/$D{t})',
            f'=IF({cnt["d15N"]}=0,"",{extreme("MIN", "am ok d15N", "am d15N", sr)})',
            f'=IF({cnt["d15N"]}=0,"",{extreme("MAX", "am ok d15N", "am d15N", sr)})',
            f'=IF({cnt["d15N"]}<2,"",$H{t}-$G{t})',
            f'=IF({cnt["d13C"]}=0,"",{extreme("MIN", "am ok d13C", "am d13C", sr)})',
            f'=IF({cnt["d13C"]}=0,"",{extreme("MAX", "am ok d13C", "am d13C", sr)})',
            f'=IF({cnt["d13C"]}<2,"",$K{t}-$J{t})',
            (f'=IF($B{t}=0,"",IF(OR({above(rx_, AM_R)},{above(nsp, AM_S)},{above(csp, AM_S)}),"Caution",'
             f'IF(COUNT({rx_},{nsp},{csp})>0,"OK","")))'),
            (f'=IF($M{t}<>"Caution","","anchors labeled "&$A{t}&" do not read as one material: "'
             f'&IF({above(rx_, AM_R)},"C/N peak-area ratio "&TEXT($D{t},"0.00")&" to "&TEXT($E{t},"0.00")'
             f'&", "&TEXT($F{t},"0.00")&" times; ","")'
             f'&IF({above(nsp, AM_S)},"raw d15N "&TEXT($G{t},"0.00")&" to "&TEXT($H{t},"0.00")&" permil; ","")'
             f'&IF({above(csp, AM_S)},"raw d13C "&TEXT($J{t},"0.00")&" to "&TEXT($K{t},"0.00")&" permil; ",""))'),
        ]
        for j, v in enumerate(vals, start=1):
            c = ft.cell(t, j, v); c.font = F_SM
            c.number_format = '0' if j in (2, 3) else '0.0000' if j in (4, 5, 6) else '@' if j in (1, 13, 14) else '0.000'
    AM_LEVEL, AM_TEXT = AMT + NSTD, AMT + NSTD + 1
    fit(AM_LEVEL, 'anchors read as one material, every row',
        f'=IF(COUNTIF($M${AMT}:$M${AMT + NSTD - 1},"Caution")>0,"Caution",'
        f'IF(COUNTIF($M${AMT}:$M${AMT + NSTD - 1},"OK")>0,"OK",""))', '@',
        'Caution when a row is, OK when any row could be judged, and empty when none could.')
    fit(AM_TEXT, 'whole run: anchors read as one material', '=' + '&'.join(f'$N${AMT + k}' for k in range(NSTD)), '@')
    ft.cell(AM_TEXT + 1, 1, 'One row per line of the standards table, from its anchor analyses only (Work columns C and '
                            'D). The ratio is Area 44 over Area 28 of an analysis with both peaks, and a quotient or a '
                            f'spread needs two values; a row is a Caution when its quotient is above B{AM0 + 1} or a '
                            f'spread above B{AM0 + 2}.').font = F_NOTE

    # 1.4.13: the elemental drift correction against the anchors, whose certified %N and %C it does not use
    EC0 = AM_TEXT + 3
    sec(EC0, 'Elemental drift correction against the anchors (1.4.13)')
    ec_text = []
    for base, el, name, pct, inf in ((EC0 + 1, 'N', 'nitrogen', '%N', 246), (EC0 + 6, 'C', 'carbon', '%C', 249)):
        on = f'$B${inf}="Yes"'
        fit(base, f'anchor analyses, {name}', f'=IF({on},COUNT({W(UX[f"adev {el}"])}),"")', '0',
            'Anchors with a certified and a reported value, as in row 308; empty when the correction is not in '
            'force (rows 246 and 249).' if el == 'N' else '')
        fit(base + 1, f'anchors\' RMS deviation, {pct} as measured',
            f'=IF(OR(NOT({on}),N($B${base})<2),"",SQRT(SUMSQ({W(RX[f"mdev {el}"])})/$B${base}))', '0.00000',
            'From certified: measured over certified, less one, before the elemental drift correction; needs two '
            'anchors.' if el == 'N' else '')
        fit(base + 2, f'anchors\' RMS deviation, {pct} corrected',
            f'=IF(OR(NOT({on}),N($B${base})<2),"",SQRT(SUMSQ({W(UX[f"adev {el}"])})/$B${base}))', '0.00000',
            'The same after it, as in row 309.' if el == 'N' else 'The same after it, as in row 322.')
        fit(base + 3, f'added by the correction, {name}', f'=IF($B${base + 1}="","",$B${base + 2}-$B${base + 1})',
            '0.00000')
        add_ = f'$B${base + 3}'
        fit(base + 4, f'{name}: the correction at the anchors',
            f'=IF({add_}="","",IF({above(add_, "(Setup!$B$36/100)")},"Caution",'
            f'IF({above(add_, "(Setup!$B$36/100/2)")},"Note","OK")))', '@',
            'A Caution when the correction adds more than Setup B36, the gate Auto applies it at, and a Note when '
            'more than half of that.' if el == 'N' else '')
        lv_ = f'$B${base + 4}'
        ec_text.append(f'IF(OR({lv_}="Note",{lv_}="Caution"),"elemental drift correction moves the anchors\' {pct} "'
                       f'&TEXT({add_}*100,"0.00")&" points further from certified (RMS "&TEXT($B${base + 1}*100,"0.00")'
                       f'&"% as measured, "&TEXT($B${base + 2}*100,"0.00")&"% corrected); ","")')
    EC_TEXT = EC0 + 11
    fit(EC_TEXT, 'whole run: elemental correction at the anchors', '=' + '&'.join(ec_text), '@')
    # 1.4.13: in the engine's order, with the two new checks after the calibration line and after the anchors
    # outside k sigma
    ALL = f'$B$357&$B$367&$B${AM_TEXT}&$B$353&$B$354&$B${EC_TEXT}&$B$355'
    fit(356, 'problems of the whole run', f'=IF({ALL}="","none found",LEFT({ALL},LEN({ALL})-2))', '@')

    fit(217, 'check standards with a calibrated value',
        f'=COUNTIFS({W("D")},"Check",{W("AA")},">-1E+30")+COUNTIFS({W("D")},"Check",{W("AF")},">-1E+30")'
        f'-COUNTIFS({W("D")},"Check",{W("AA")},">-1E+30",{W("AF")},">-1E+30")', '0',
        'Materials with the Check role that were measured in this run. Zero means nothing independent '
        'of the anchors tests the calibration.')

    # ================================================================ Results
    rs = wb.create_sheet('Results')
    RHDR = [('Row', 6), ('Identifier', 18), ('Type', 10), ('Mass (mg)', 10),
            ('%C', 8), ('1σ(%C)', 8), ('%N', 8), ('1σ(%N)', 8),
            ('C:N molar', 10), ('1σ(C:N molar)', 10), ('C:N mass', 10), ('1σ(C:N mass)', 10),
            ('d13C (permil)', 12), ('1σ(d13C)', 9), ('d15N (permil)', 12), ('1σ(d15N)', 9), ('Flags', 54),
            ('Rerun', 9), ('Why', 46), ('Suggested mass (mg)', 24),
            ('Elemental drift factor N', 10), ('Elemental drift factor C', 10),
            ('%C as measured', 9), ('%N as measured', 9)]
    NRES = len(RHDR)                                      # the 24 columns every workbook has
    MHDR = output_header(samples, _MCOLS)                 # 1.4.8: the sample sheet's, when one was given
    RHDR += [(h, 16 if j == 0 else 11 if j == 1 else 14) for j, h in enumerate(MHDR)]
    rs.cell(1, 1, 'Results').font = F_TITLE
    rs.cell(2, 1, 'Read QC before using these. Isotope values are on the scale set by the anchor standards '
                  'in Setup; %C and %N need a mass in Work column F or G (C:N comes from the two peaks and needs '
                  'none), and from 1.4.2 the three carry the '
                  'elemental drift correction when it is in force (Fits row 243 onwards), with the factors and '
                  'the values as measured in the last four columns. Each value is followed by its 1σ, one standard '
                  'deviation (1.4.5); the checks on QC and Rerun use k x 1σ, k from Setup B32. Flags is followed by '
                  'the rerun mark (1.4.4), its reason, and a suggested mass, which the Rerun sheet lists. For the '
                  'isotope values 1σ comes from the scale '
                  'calibration and the sample repeatability (Fits rows 132 onwards) and leaves out the blank, '
                  'linearity, and drift corrections and day-to-day reproducibility; for %C, %N, and C:N it has five '
                  'terms (Fits rows 300 onwards), and stays incomplete until Setup B33 and Setup columns J and K '
                  'are filled in.' + (' From 1.4.8 the columns after %N as measured come from the sample sheet '
                                      '(Metadata sheet): the ID a sample matched, how it matched, and the sheet\'s '
                                      'columns that hold a value for a sample of this run.' if MHDR else '')
                  ).font = F_NOTE
    for j, (h, w_) in enumerate(RHDR, start=1):
        c = put(rs, 4, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
        rs.column_dimensions[gcl(j)].width = w_
    for n in range(NSAMP):
        i, k = 5 + n, R0 + n
        PCc, PNc = EX['%C corrected'], EX['%N corrected']
        CNm, CNw = EX['C:N molar corrected'], EX['C:N mass corrected']
        pull = lambda col: f'=IF(Work!${col}{k}="","",Work!${col}{k})'      # noqa: E731
        vals = [f'=IF(Work!$B{k}="","",Work!$A{k})', f'=IF(Work!$B{k}="","",Work!$B{k})',
                f'=IF(Work!$B{k}="","",Work!$D{k})', f'=IF(Work!$G{k}="","",Work!$G{k})',
                pull(PCc), pull(UX['u %C']), pull(PNc), pull(UX['u %N']),
                pull(CNm), pull(UX['u CN molar']), pull(CNw), pull(UX['u CN mass']),
                f'=IF(Work!$AF{k}="","",Work!$AF{k})', pull('CN'), f'=IF(Work!$AA{k}="","",Work!$AA{k})', pull('CL'),
                (f'=IF(Work!$B{k}="","",'
                 # 1.4.12: in the engine's order, the N peak's flags and then the C peak's
                 f'Work!${RX["fl noN"]}{k}&Work!${RX["fl sizeN"]}{k}&Work!${RX["fl noC"]}{k}'
                 f'&Work!${RX["fl rejC"]}{k}&Work!${RX["fl sizeC"]}{k}&Work!${RX["fl cpre"]}{k}'
                 f'&IF(OR(Work!$D{k}="Sample",Work!$D{k}="Check"),IF(Work!$AA{k}="","",IF(Fits!$B$138="","",'
                 f'IF(AND(Work!$AA{k}>=Fits!$B$138,Work!$AA{k}<=Fits!$B$139),"","d15N outside anchor range; "))),"")'
                 f'&IF(OR(Work!$D{k}="Sample",Work!$D{k}="Check"),IF(Work!$AF{k}="","",IF(Fits!$B$158="","",'
                 f'IF(AND(Work!$AF{k}>=Fits!$B$158,Work!$AF{k}<=Fits!$B$159),"","d13C outside anchor range; "))),"")'
                 f'&Work!${RX["fl checkN"]}{k}&Work!${RX["fl checkC"]}{k}'
                 f'&IF(Work!$D{k}<>"Drift","",IF(Work!$BX{k}=1,"excluded from N drift fit; ",'
                 f'IF(Work!$CB{k}="","",IF(ABS(Work!$CB{k})>{S_TOUT},"N drift monitor outlier; ",""))))'
                 f'&IF(Work!$D{k}<>"Drift","",IF(Work!$BY{k}=1,"excluded from C drift fit; ",'
                 f'IF(Work!$CC{k}="","",IF(ABS(Work!$CC{k})>{S_TOUT},"C drift monitor outlier; ",""))))'
                 f'&IF(AND(Work!$D{k}="Drift",Work!${EX["excl E"]}{k}=1),"excluded from elemental correction; ","")'
                 # finding R-23: no reference background under the Background model, so no drift correction
                 f'&IF(AND(Work!$X{k}<>"",Work!$Y{k}="",Fits!$B$115="Yes"),"{NO_BGD_N}; ","")'
                 f'&IF(AND(Work!$AC{k}<>"",Work!$AD{k}="",Fits!$B$128="Yes"),"{NO_BGD_C}; ","")'
                 f'&Work!${RX["fl blankN"]}{k}&Work!${RX["fl blankC"]}{k}'
                 f'&Work!${RX["fl settle"]}{k}&Work!${RX["fl carry"]}{k}'
                 f'&Work!${RX["fl pulseN"]}{k}&Work!${RX["fl pulseC"]}{k}'
                 f'&IF(AND(ISNUMBER(Work!$AG{k}),ISNUMBER(Work!$G{k})),IF(AND(Work!$AG{k}<=0,Work!$G{k}>0),'
                 f'"%N below detection; ",""),"")'
                 f'&IF(AND(ISNUMBER(Work!$AH{k}),ISNUMBER(Work!$G{k})),IF(AND(Work!$AH{k}<=0,Work!$G{k}>0),'
                 f'"%C below detection; ",""),""))'),
                f'=IF(Work!$B{k}="","",Work!${RX["rerun"]}{k})', f'=IF(Work!$B{k}="","",Work!${RX["why"]}{k})',
                f'=IF(Work!$B{k}="","",Work!${RX["suggested mass"]}{k})',
                f'=IF(Work!$B{k}="","",Work!${EX["elem factor N"]}{k})',
                f'=IF(Work!$B{k}="","",Work!${EX["elem factor C"]}{k})',
                f'=IF(Work!$AJ{k}="","",Work!$AJ{k})', f'=IF(Work!$AI{k}="","",Work!$AI{k})']
        for j, v in enumerate(vals, start=1):
            c = rs.cell(i, j, v); c.font = F_SM
            if j in (5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 23, 24): c.number_format = '0.00'
            if j == 4: c.number_format = '0.000'
            if j in (14, 16): c.number_format = '0.000'
            if j in (21, 22): c.number_format = '0.0000'
            if 5 <= j <= 16: c.fill = FILL_OUT
    # 1.4.8: the metadata cells read the Metadata sheet's table, while the identifier is the one matched
    MD_TOP = 24                                           # first row of the samples table on Metadata
    MD_ROW = {}
    if samples is not None:
        for t_, q in enumerate([q for q in _META if q.role == 'Sample' and 1 <= q.row <= NSAMP]):
            MD_ROW[q.row] = MD_TOP + t_
        for pos_, t_ in MD_ROW.items():
            k = R0 + pos_ - 1
            for j in range(len(MHDR)):
                L = gcl(3 + (j if j < 2 else j + 1))      # C, D, then F onwards (E is the replicate group)
                c = rs.cell(4 + pos_, NRES + 1 + j,
                            f'=IF(TRIM(Work!$B{k})="{lit(pos_)}",IF(Metadata!${L}${t_}="","",Metadata!${L}${t_}),"")')
                c.font = F_SM
    rs.freeze_panes = 'A5'
    RLAST = gcl(len(RHDR))
    rs.conditional_formatting.add(f'A5:{RLAST}{5+NSAMP-1}', FormulaRule(formula=['$R5="Check"'], fill=FILL_CHECK))
    rs.conditional_formatting.add(f'A5:{RLAST}{5+NSAMP-1}',
                                  FormulaRule(formula=['AND($R5<>"",$R5<>"Check")'], fill=FILL_RERUN))

    # ================================================================ Rerun (1.4.4)
    rr = wb.create_sheet('Rerun')
    rr.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFG', [6, 18, 10, 10, 10, 60, 30]):
        rr.column_dimensions[col].width = w
    rr.cell(1, 1, 'Rerun').font = F_TITLE
    rr.cell(2, 1, 'Samples to measure again (Rerun: N, C, or N and C) and to check first, in the order they ran. '
                  'A mark reads what a rerun can fix; README, Rerun marks, says which flags count (from 1.4.7 a '
                  'peak above the linearity range only at Setup B69, and a 1σ above Setup B70; from 1.4.8 replicates '
                  'above the critical range, on the Replicates sheet; from 1.4.14 possible carryover marks a check, '
                  'not a rerun). The suggested mass puts the peaks inside the '
                  'linearity range, from this analysis\'s area per mg above the blank (1.4.10).'
                  ).font = F_NOTE
    rr.merge_cells("A2:G2"); rr.row_dimensions[2].height = 56
    rr['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    rr.cell(4, 1, 'Problems of the whole run').font = F_H1
    rr.cell(5, 1, '=Fits!$B$356').font = F_SM
    rr.merge_cells('A5:G5'); rr.row_dimensions[5].height = 30
    rr['A5'].alignment = Alignment(wrap_text=True, vertical='top')
    for j, h in enumerate(['Row', 'Identifier', 'Type', 'Mass (mg)', 'Rerun', 'Why', 'Suggested mass (mg)'],
                          start=1):
        c = rr.cell(7, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    RANK = W(RX['rerun rank'])
    for n in range(NSAMP):
        i = 8 + n
        hit = f'MATCH({n + 1},{RANK},0)'
        vals = [f'=IFERROR(INDEX({W("A")},{hit}),"")']
        vals += [f'=IF($A{i}="","",INDEX({W(col)},{hit}))'
                 for col in ('B', 'D', 'G', RX['rerun'], RX['why'], RX['suggested mass'])]
        for j, v in enumerate(vals, start=1):
            c = rr.cell(i, j, v); c.font = F_SM
            if j == 4: c.number_format = '0.000'
    rr.freeze_panes = 'A8'
    rr.conditional_formatting.add(f'A8:G{8+NSAMP-1}', FormulaRule(formula=['$E8="Check"'], fill=FILL_CHECK))
    rr.conditional_formatting.add(f'A8:G{8+NSAMP-1}',
                                  FormulaRule(formula=['AND($E8<>"",$E8<>"Check")'], fill=FILL_RERUN))

    # ================================================================ Replicates (1.4.8)
    rp = wb.create_sheet('Replicates')
    rp.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGHIJKLMNOPQR', [7, 18, 8, 18, 10, 10, 7, 9, 9, 7, 7, 9, 9, 7, 22, 3, 5, 9]):
        rp.column_dimensions[col].width = w
    rp.cell(1, 1, 'Replicates').font = F_TITLE
    rp.cell(2, 1, 'Analyses of one sample in this run as the sample sheet names them, that is, the samples matched to '
                  'it with the same name or the same Replicate of (Metadata sheet). When their range in d15N or d13C is '
                  'above the critical range of ISO 5725-6, f(n) x s, each of them is marked for a rerun of that '
                  'element, unless Setup B71 is Off. s is the repeatability of one analysis in use (Fits B136 and B156, '
                  'from Setup B30 and B31 or the pooled anchor replicates) times the calibration slope, and f(n), for '
                  'the n analyses with a value, is the 0.95 quantile of the range of n values from one normal '
                  'distribution in units of its SD (columns Q and R), 2.77 for two, which ISO 5725-6 rounds to 2.8 in '
                  'r = 2.8 s.').font = F_NOTE
    rp.merge_cells('A2:O2'); rp.row_dimensions[2].height = 66
    rp['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    NF = max([20] + [len(g_['rows']) for g_ in _GROUPS])
    FT = f'$R$4:$R${4 + NF - 2}'
    for r_, (k_, v_, note_) in enumerate((
            ('s, d15N (permil)', '=IF(OR(Fits!$C$55<>"Yes",Fits!$B$55<3,Fits!$B$134="",Fits!$B$134=0,Fits!$B$136=""),"",'
                                 'Fits!$B$136*ABS(Fits!$B$56))', 'Fits B136 x |Fits B56|'),
            ('s, d13C (permil)', '=IF(OR(Fits!$C$62<>"Yes",Fits!$B$62<3,Fits!$B$154="",Fits!$B$154=0,Fits!$B$156=""),"",'
                                 'Fits!$B$156*ABS(Fits!$B$63))', 'Fits B156 x |Fits B63|'),
            ('Comparison', f'={S_REP}', 'Setup B71')), start=4):
        rp.cell(r_, 1, k_).font = F_SM
        c = rp.cell(r_, 2, v_); c.font = F_SM; c.number_format = '0.000'
        rp.cell(r_, 3, note_).font = F_NOTE
    for j, h in enumerate(('n', 'f(n)'), start=17):
        c = rp.cell(3, j, h); c.font = F_HDR; c.fill = FILL_HDR; c.alignment = Alignment(horizontal='center')
    for k_, nn in enumerate(range(2, NF + 1)):
        rp.cell(4 + k_, 17, nn).font = F_SM
        c = rp.cell(4 + k_, 18, critical_range_factor(nn)); c.font = F_SM; c.number_format = '0.000000'
    RPH = ['Group', 'Sample', 'Position', 'Identifier', 'd15N (permil)', 'd13C (permil)',
           'n d15N', 'Range d15N', 'Limit d15N', 'Over d15N', 'n d13C', 'Range d13C', 'Limit d13C', 'Over d13C',
           'Positions']
    for j, h in enumerate(RPH, start=1):
        c = rp.cell(8, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    if not _GROUPS:
        # 1.4.13 (finding R-34): say that no group formed, and why a name ending in Dup was not compared
        dup = ('Names ending in Dup are compared only when the sample sheet\'s Replicate of column names them: CI 3-2 '
               'Dup is a replicate of CI 3-2 only when its own row of the sheet has CI 3-2 in that column.')
        c = rp.cell(REP_TOP, 1, ('No sample sheet was given, so no replicate group formed and no replicates are '
                                 'compared. ' if samples is None else
                                 'No replicate group formed in this run: no two samples matched to the sample sheet '
                                 'share a name or a Replicate of, so no replicates are compared. ') + dup)
        c.font = F_NOTE
        c.alignment = Alignment(wrap_text=True, vertical='top')
        rp.merge_cells(start_row=REP_TOP, start_column=1, end_row=REP_TOP, end_column=15)
        rp.row_dimensions[REP_TOP].height = 30
    for g_ in _GROUPS:
        top_, bot_ = REP_FIRST[g_['group']], REP_FIRST[g_['group']] + len(g_['rows']) - 1
        for k_, pos_ in enumerate(g_['rows']):
            i_, w_ = top_ + k_, R0 + pos_ - 1
            vals = [g_['group'], g_['name'], pos_, f'=IF(Work!$B{w_}="","",Work!$B{w_})']
            for col_ in ('AA', 'AF'):
                vals.append(f'=IF(AND(Work!$D{w_}="Sample",Work!${RX["rep group"]}{w_}={g_["group"]}),'
                            f'IF(ISNUMBER(Work!${col_}{w_}),Work!${col_}{w_},""),"")')
            for j, v in enumerate(vals, start=1):
                c = put(rp, i_, j, v) if j == 2 else rp.cell(i_, j, v); c.font = F_SM
                if j in (5, 6): c.number_format = '0.00'
        for (cnt, rng, lim, over), val, srow in ((('G', 'H', 'I', 'J'), 'E', 4), (('K', 'L', 'M', 'N'), 'F', 5)):
            span = f'{val}{top_}:{val}{bot_}'
            for col_, v in ((cnt, f'=COUNT({span})'),
                            (rng, f'=IF({cnt}{top_}<2,"",MAX({span})-MIN({span}))'),
                            (lim, f'=IF(OR({cnt}{top_}<2,$B${srow}=""),"",INDEX({FT},{cnt}{top_}-1)*$B${srow})'),
                            (over, f'=IF(OR({rng}{top_}="",{lim}{top_}=""),"",'
                                   f'IF(AND($B$6="On",{rng}{top_}>{lim}{top_}),"yes",""))')):
                c = rp[f'{col_}{top_}']; c.value = v; c.font = F_SM
                if col_ in (rng, lim): c.number_format = '0.000'
        c = rp.cell(top_, 15, ', '.join(str(r) for r in g_['rows'])); c.font = F_SM
        rp.conditional_formatting.add(f'A{top_}:O{bot_}',
                                      FormulaRule(formula=[f'OR($J${top_}="yes",$N${top_}="yes")'], fill=FILL_RERUN))
    rp.freeze_panes = 'A9'

    # ================================================================ QC
    qc = wb.create_sheet('QC')
    qc.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGHI', [40, 13, 13, 11, 12, 11, 12, 12, 11]):
        qc.column_dimensions[col].width = w
    qc.cell(1, 1, 'QC').font = F_TITLE
    qc.cell(2, 1, 'Read this before taking any number off the Results sheet.').font = F_NOTE

    def qsec(r_, title):
        c = qc.cell(r_, 1, title); c.font = F_H1; c.fill = FILL_GRP
        for j in range(2, 10): qc.cell(r_, j).fill = FILL_GRP

    def qline(r_, label, formula, fmt='0.000', verdict=None):
        qc.cell(r_, 1, label).font = F_SM
        c = qc.cell(r_, 2, formula); c.font = F_SM; c.number_format = fmt; c.border = BOX
        if verdict:
            qc.cell(r_, 4, verdict).font = F_SM

    WB_ = W('B'); WD_ = W('D'); WA_ = W('A'); WAA = W('AA'); WAF = W('AF'); WC_ = W('C')

    qsec(4, 'Run')
    qline(5, 'Autosampler positions detected', f'=COUNTIF({WB_},"?*")', '0')
    qline(6, 'IRMS peak records read from Raw', f'=COUNT({rawD})', '0')
    qline(7, 'Unknowns (role Sample)',         f'=COUNTIF({WD_},"Sample")', '0')
    qline(8, 'Positions with no nitrogen peak',
          f'=SUMPRODUCT(({WB_}<>"")*1)-COUNT(Work!$L${R0}:$L${R0+NSAMP-1})', '0')

    qsec(10, 'Working gas, from the reference pulses Isodat did not designate')
    qc.cell(11, 1, 'The designated pulse carries the assigned value by definition, so it cannot show movement. '
                   'An export cut down to the sample peaks carries no other pulses, and then this section says '
                   'so.').font = F_NOTE
    for r_, base, name in ((12, 201, 'N2'), (16, 208, 'CO2')):
        qline(r_,     f'{name} pulses read',              f'=Fits!$B${base}', '0',
              f'=IF(Fits!$B${base}<3,"no reference pulses in this export; nothing to check","ok")')
        qline(r_ + 1, f'{name} pulse SD (permil)',        f'=Fits!$B${base+2}')
        qline(r_ + 2, f'{name} pulse trend, permil per position', f'=Fits!$B${base+4}', '0.00000',
              f'=IF(Fits!$B${base}<3,"",IF(AND(ISNUMBER(Fits!$B${base+5}),Fits!$B${base+5}<{S_ALPHA}),'
              f'"CHECK: the pulses trend across the run (p="&TEXT(Fits!$B${base+5},"0.000")&"), so the '
              f'reference measurement itself moved","ok, no trend (p="&TEXT(Fits!$B${base+5},"0.00")&")"))')
        qline(r_ + 3, f'{name} pulse range (permil)',     f'=Fits!$B${base+3}')

    qsec(21, 'Calibration')
    qline(22, 'Nitrogen anchors used', '=Fits!$B$55', '0',
          '=IF(Fits!$B$55<2,"CHECK: need at least two anchor measurements",IF(Fits!$C$55<>"Yes",'
          '"CHECK: every anchor carries the same accepted d15N; two materials with different values set the '
          'scale, so no d15N is reported","ok"))')
    # 1.4.12 (finding W-02): the slope and the anchors' residual SD against Setup B73 to B76
    qline(23, 'Nitrogen calibration slope', '=Fits!$B$56', '0.0000',
          f'=IF(Fits!$B$363="","",IF(Fits!$B$363="OK","ok","CHECK, "&Fits!$B$363&": the slope or the anchors\' '
          f'residual SD (Fits rows 361 and 362) is beyond the limits in Setup B{SLAST+13} to B{SLAST+16}"))')
    qline(24, 'Carbon anchors used', '=Fits!$B$62', '0',
          '=IF(Fits!$B$62<2,"CHECK: need at least two anchor measurements",IF(Fits!$C$62<>"Yes",'
          '"CHECK: every anchor carries the same accepted d13C; two materials with different values set the '
          'scale, so no d13C is reported","ok"))')
    qline(25, 'Carbon calibration slope', '=Fits!$B$63', '0.0000',
          f'=IF(Fits!$B$366="","",IF(Fits!$B$366="OK","ok","CHECK, "&Fits!$B$366&": the slope or the anchors\' '
          f'residual SD (Fits rows 364 and 365) is beyond the limits in Setup B{SLAST+13} to B{SLAST+16}"))')
    qline(26, 'Check standards measured in this run', '=Fits!$B$217', '0',
          '=IF(Fits!$B$217<1,"CHECK: nothing independent of the anchors tests this calibration; run a '
          'check standard (USGS66, or an in-house material with an assigned value) in every sequence","ok")')
    qline(27, 'Samples outside the anchor d15N range',
          f'=COUNTIF(Results!$Q$5:$Q${5+NSAMP-1},"*d15N outside anchor range*")', '0',
          '=IF($B$27=0,"ok","CHECK: these values are extrapolated beyond the anchors; add an anchor that '
          'brackets them")')
    qline(28, 'Samples outside the anchor d13C range',
          f'=COUNTIF(Results!$Q$5:$Q${5+NSAMP-1},"*d13C outside anchor range*")', '0',
          '=IF($B$28=0,"ok","CHECK: these values are extrapolated beyond the anchors; add an anchor that '
          'brackets them")')

    qsec(30, 'Is each correction supported by this run?')
    qc.cell(31, 1, 'A slope is applied on Auto only when p is below the significance level in Setup B15. '
                   'The FIRMS guide (section 6.4) asks that a correction be shown necessary before it is '
                   'applied.').font = F_NOTE
    def decision(r_, label, slope, se, p, force):
        qline(r_, label, f'={slope}', '0.00000',
              f'=IF({slope}="","no fit",IF({p}="","too few standards to test",'
              f'"slope "&TEXT({slope},"0.00000")&" +/- "&TEXT({se},"0.00000")&", p = "&TEXT({p},"0.000")'
              f'&IF({force}="Yes",", applied",", not applied")))')
    decision(32, 'Nitrogen linearity slope, permil per area unit', 'Fits!$B$17', 'Fits!$B$93', 'Fits!$B$95', 'Fits!$B$96')
    decision(33, 'Carbon linearity slope',                        'Fits!$B$25', 'Fits!$B$100', 'Fits!$B$102', 'Fits!$B$103')
    decision(34, 'Nitrogen drift slope, per position or per background unit',
             'IF(Fits!$B$38="Row",Fits!$B$32,Fits!$B$35)', 'Fits!$B$112', 'Fits!$B$114', 'Fits!$B$115')
    decision(35, 'Carbon drift slope',
             'IF(Fits!$B$50="Row",Fits!$B$44,Fits!$B$47)', 'Fits!$B$125', 'Fits!$B$127', 'Fits!$B$128')
    qc.cell(36, 1, 'With five or six standards the test has little power, so it can miss a slope that is there; '
                   'set the switch to On to apply it regardless. The outlier test below needs five monitors in the fit and '
                   'is not run with fewer (1.4.2).').font = F_NOTE
    qline(37, 'Nitrogen drift monitors flagged as outliers', '=Fits!$B$117', '0',
          f'=IF(Fits!$B$117=0,"ok","CHECK: a monitor sits far from the line through the others (Work column CB). '
          f'If the capsule is the cause, list its position in Setup B16 and read QC again")')
    qline(38, 'Carbon drift monitors flagged as outliers', '=Fits!$B$130', '0',
          f'=IF(Fits!$B$130=0,"ok","CHECK: a monitor sits far from the line through the others (Work column CC). '
          f'If the capsule is the cause, list its position in Setup B17 and read QC again")')
    qline(39, 'Nitrogen model in use', '=Fits!$B$38', '@')
    qline(40, 'Residual SD, row model (permil)', '=Fits!$B$34')
    qline(41, 'Residual SD, background model (permil)', '=Fits!$B$37')
    qline(42, 'Carbon model in use', '=Fits!$B$50', '@')
    qline(43, 'Residual SD, row model (permil)', '=Fits!$B$46')
    qline(44, 'Residual SD, background model (permil)', '=Fits!$B$49')
    qc.cell(45, 1, 'Residual SD is the scatter left in the drift standards themselves, so it cannot decide '
                   'between the two models; the check standard offsets below can.').font = F_NOTE

    qsec(47, 'Blank')
    qline(48, 'Blank correction in force, nitrogen', '=Fits!$C$5', '@')
    NEXB = f'COUNTIFS({WD_},"Exclude",{WB_},"*blank*",{WB_},"<>*hole*")'
    qc.cell(48, 4, f'=IF(Setup!$B$10="Off","switched off in Setup",'
                   f'IF(Fits!$B$5=0,IF({NEXB}>0,"no blanks in the fit: "&TEXT({NEXB},"0")&" blank capsules are '
                   f'set to Exclude in Setup","no blank capsules found"),'
                   'IF(Fits!$C$5="Yes",IF(Fits!$B$5<3,"CHECK: running on fewer than three blanks",'
                   'IF(Fits!$B$8>2,"CHECK: blanks disagree by more than 2 permil","ok")),'
                   'IF(Fits!$B$5<3,"not applied: fewer than three blanks",'
                   '"not applied: blanks disagree by more than 2 permil"))))').font = F_SM
    qline(49, 'Blank capsules with a nitrogen peak', '=Fits!$B$5', '0')
    qline(50, 'SD of blank d15N (permil)', '=Fits!$B$8')
    qline(51, 'Blank correction in force, carbon', '=Fits!$C$9', '@')
    qline(52, 'Blank capsules with a carbon peak', '=Fits!$B$9', '0')
    qline(53, 'SD of blank d13C (permil)', '=Fits!$B$12')
    qc.cell(54, 1, 'A blank that gives no peak has not been measured, only bounded by the detection '
                   'limit. The guide asks that its size and d value be determined, which the two-standard '
                   'decreasing-size series does.').font = F_NOTE

    qsec(56, 'Check standards and drift monitors')
    CHDR = ['Material', 'Role', 'n', 'mean d15N', 'accepted', 'offset', 'mean d13C', 'accepted', 'offset']
    for j, h in enumerate(CHDR, start=1):
        c = qc.cell(57, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    for k in range(NSTD):
        i, sr = 58 + k, SFIRST + k
        cond = f'OR(Setup!$B${sr}="Check",Setup!$B${sr}="Drift",Setup!$B${sr}="Anchor")'
        vals = [
            f'=IF(AND(Setup!$A${sr}<>"",{cond}),Setup!$A${sr},"")',
            f'=IF(AND(Setup!$A${sr}<>"",{cond}),Setup!$B${sr},"")',
            f'=IF($A{i}="","",COUNTIFS({WC_},{sr},{WAA},">-1E+30"))',
            f'=IF($A{i}="","",IFERROR(AVERAGEIFS({WAA},{WC_},{sr}),""))',
            f'=IF($A{i}="","",IF(Setup!$D${sr}="","",Setup!$D${sr}))',
            f'=IF(OR($D{i}="",$E{i}=""),"",$D{i}-$E{i})',
            f'=IF($A{i}="","",IFERROR(AVERAGEIFS({WAF},{WC_},{sr}),""))',
            f'=IF($A{i}="","",IF(Setup!$C${sr}="","",Setup!$C${sr}))',
            f'=IF(OR($G{i}="",$H{i}=""),"",$G{i}-$H{i})',
        ]
        for j, v in enumerate(vals, start=1):
            c = qc.cell(i, j, v); c.font = F_SM
            if j in (4, 5, 6, 7, 8, 9): c.number_format = '0.00'
            if j == 3: c.number_format = '0'
            if j in (6, 9): c.fill = FILL_OUT
    qc.cell(58 + NSTD, 1, 'An anchor is forced onto its accepted value by construction, so only the check rows '
                          'are an independent test. Offsets are in permil.').font = F_NOTE

    # The rows below moved down by three in 1.4.2, with the standards table.
    Q = 58 + NSTD + 2
    qsec(Q, 'Uncertainty')
    qline(Q + 1, 'Sample repeatability in use, d15N (permil)', '=Fits!$B$136', '0.000', '=Fits!$B$137')
    qline(Q + 2, 'Sample repeatability in use, d13C (permil)', '=Fits!$B$156', '0.000', '=Fits!$B$157')
    qline(Q + 3, 'Multiple of 1σ for the checks (k)', f'={S_K}', '0.0')
    qline(Q + 4, 'Typical 1σ(d15N) of the unknowns (permil)',
          f'=IFERROR(AVERAGEIFS({W("CL")},{WD_},"Sample"),"")')
    qline(Q + 5, 'Typical 1σ(d13C) of the unknowns (permil)',
          f'=IFERROR(AVERAGEIFS({W("CN")},{WD_},"Sample"),"")')
    qc.cell(Q + 6, 1, '1σ combines the certificate uncertainty of each anchor, the scatter of its replicates, and '
                      'the sample repeatability, propagated through the fitted line to first order, the effect a '
                      'Kragten spreadsheet estimates by changing each input by its 1σ. It '
                      'leaves out the blank, linearity, and drift corrections and the day-to-day reproducibility '
                      'a control chart would give.').font = F_NOTE
    qline(Q + 7, 'SD of the drift standards, d15N (permil)', f'=IFERROR(STDEV(Work!$BG${R0}:$BG${R0+NSAMP-1}),"")')
    qline(Q + 8, 'SD of the drift standards, d13C (permil)', f'=IFERROR(STDEV(Work!$BH${R0}:$BH${R0+NSAMP-1}),"")')
    qc.cell(Q + 9, 1, 'These understate the scatter of an unknown, because the drift correction was fitted to these '
                      'same standards; the check standard offsets above test it independently.').font = F_NOTE

    qsec(Q + 11, 'Flags')
    qline(Q + 12, 'Rows carrying a flag on Results', f'=COUNTIF(Results!$Q$5:$Q${5+NSAMP-1},"?*")', '0')
    qline(Q + 13, 'Minimum trusted Area 28', '=Fits!$B$81')
    qline(Q + 14, 'Maximum trusted Area 28', '=Fits!$B$97')
    qline(Q + 15, 'Minimum trusted Area 44', '=Fits!$B$82')
    qline(Q + 16, 'Maximum trusted Area 44', '=Fits!$B$104')
    qline(Q + 17, 'Samples outside the linearity range',
          f'=COUNTIF(Results!$Q$5:$Q${5+NSAMP-1},"*linearity range*")', '0',
          f'=IF($B${Q + 17}=0,"ok","CHECK: weigh samples to land inside the standards, or extend the size series")')

    qsec(96, 'Instrument settling')
    qc.cell(97, 1, 'A run that began before the mass-28 background reached its floor carries a settling '
                   'transient, which a straight line in position does not follow.').font = F_NOTE
    NBI = f'COUNT({W("BI")})'      # reference-peak backgrounds read; zero for an export cut to peaks 3 and 4
    qline(98, 'Reference-peak background floor', f'=IF({NBI}=0,"",Fits!$B$83)', '0.00',
          f'=IF({NBI}=0,"no reference peaks in this export, so rows 98 to 105 have nothing to read; '
          f'rows 106 and 107 use the background under the drift monitors instead","")')
    qline(99, 'Starting excess above the floor', f'=IF({NBI}=0,"",Fits!$B$85)', '0.00')
    qline(100, 'Background settled from position', f'=IF({NBI}=0,"",Fits!$B$86)', '0',
          f'=IF({NBI}=0,"",IF(Fits!$B$86<=2,"ok, settled from the start",'
          f'"CHECK: positions 1 to "&TEXT(Fits!$B$86-1,"0")&" ran before the background settled"))')
    qline(101, 'First position used in the drift fit', f'=IF({S_FIRST}="","all",{S_FIRST})', '@',
          f'=IF(OR({NBI}=0,{S_FIRST}<>""),"",IF(Fits!$B$86<=2,"","consider setting Setup B14 to "'
          f'&TEXT(Fits!$B$86,"0")))')
    qline(102, 'Linearity standards before that position',
          f'=IF({NBI}=0,"",COUNTIFS({WD_},"Linearity",{WA_},"<"&Fits!$B$86))', '0',
          f'=IF({NBI}=0,"",IF($B$102=0,"ok",'
          f'"CHECK: size and time are confounded; the linearity fit sits inside the unsettled window"))')
    qline(103, 'Calibration anchors before that position',
          f'=IF({NBI}=0,"",COUNTIFS({WD_},"Anchor",{WA_},"<"&Fits!$B$86))', '0',
          f'=IF({NBI}=0,"",IF($B$103=0,"ok","CHECK: the scale was set while the instrument was still settling"))')
    qline(104, 'Drift slope before the settled position', f'=IF({NBI}=0,"",Fits!$B$87)')
    qline(105, 'Drift slope from the settled position',   f'=IF({NBI}=0,"",Fits!$B$89)', '0.000',
          f'=IF({NBI}=0,"",IF(OR(Fits!$B$88<3,Fits!$B$90<3),"too few standards in one segment to compare",'
          f'IF(ABS(Fits!$B$89)<0.4*ABS(Fits!$B$87),"settling transient, not steady drift","slopes are comparable")))')
    qline(106, 'Sample-peak background on the drift monitors, lowest to highest',
          '=IF(Fits!$B$221<2,"",TEXT(Fits!$B$222,"0.0")&" to "&TEXT(Fits!$B$223,"0.0"))', '@')
    qline(107, 'Its trend, per position', '=Fits!$B$224', '0.000',
          '=IF(Fits!$B$221<3,"too few monitors",IF(AND(ISNUMBER(Fits!$B$225),Fits!$B$225<'+S_ALPHA+'),'
          '"CHECK: the background under the monitors moved across the run (p="&TEXT(Fits!$B$225,"0.000")&")",'
          '"ok, no trend (p="&TEXT(Fits!$B$225,"0.00")&")"))')
    # Two short lines, since the background chart starts at column K and hides
    # whatever a long note spills into it.
    qc.cell(108, 1, 'The drift monitors are all one size, so their background follows the instrument.').font = F_NOTE
    qc.cell(109, 1, 'A high background right after a large sample points to carryover from that sample.').font = F_NOTE

    qsec(110, 'Does each correction reduce the scatter of the drift standards?')
    STAGES = [('as measured', 'BO'), ('after blank', 'BP'), ('after linearity', 'BQ'),
              ('after drift', 'BR'), ('after calibration', 'BG')]
    for i, (nm, col) in enumerate(STAGES):
        qline(111 + i, nm, f'=IFERROR(STDEV({W(col)}),"")')
    qc.cell(116, 4, 'SD in permil.').font = F_NOTE
    qc.cell(117, 1, 'This table can judge the blank and drift corrections but not the linearity one, because the '
                    'drift standards are all about the same size and a size-dependent correction moves them together. '
                    'Judge linearity by the check standards and by the spread of the linearity series itself.').font = F_NOTE

    qsec(119, 'Carryover after oversized samples')
    qc.cell(120, 1, 'A sample above the linearity range can leave something in the EA that raises the background '
                    'of the next one or two analyses.').font = F_NOTE
    qline(121, 'Positions above the linearity range', '=Fits!$B$229', '0')
    qline(122, 'Of those, followed directly by a sample or check', '=Fits!$B$230', '0',
          '=IF($B$122=0,"ok","CHECK: put a blank capsule or a drift monitor after each oversized sample")')
    qline(123, 'Positions flagged for possible carryover', '=Fits!$B$231', '0',
          '=IF(Fits!$B$228="","check switched off in Setup B28",IF($B$123=0,"ok",'
          '"CHECK: listed below, and marked in red on the background chart at K95"))')
    for j, h in enumerate(['Flagged position', 'Position', 'Rise (mV)', 'Share of peak height'], start=1):
        c = qc.cell(124, j, h); c.font = F_HDR; c.fill = FILL_HDR
    RANK = W(HX['carry rank'])
    for j in range(10):
        r_ = 125 + j
        at = f'MATCH({j + 1},{RANK},0)'
        cells = [f'=IFERROR(INDEX({W("B")},{at}),"")', f'=IFERROR(INDEX({W("A")},{at}),"")',
                 f'=IFERROR(INDEX({W(HX["carry rise"])},{at}),"")',
                 f'=IFERROR(INDEX({W(HX["carry rise"])},{at})/INDEX({W("M")},{at}),"")']
        for col, (v, fmt) in enumerate(zip(cells, ['@', '0', '0.0', '0.00%']), start=1):
            c = qc.cell(r_, col, v); c.font = F_SM; c.number_format = fmt
    qc.cell(135, 1, '=IF(Fits!$B$231>10,"and "&TEXT(Fits!$B$231-10,"0")&" more, all flagged on Results","")').font = F_SM
    qc.cell(136, 1, 'The rise is measured against the last background before the oversized samples, so a slow '
                    'change across the run does not count. The same rise weighs more under a small peak.').font = F_NOTE

    # 1.4.2
    qsec(138, 'Elemental drift, from the drift monitors')
    qc.cell(139, 1, 'The k-factor is one line for the whole run; when the sensitivity moves, the drift monitors read '
                    'high or low in %N and %C by position, and so do the samples around them. Results carries the '
                    'corrected values when the correction is in force.').font = F_NOTE
    for r_, base, el, name in ((140, 244, 'N', '%N'), (142, 247, 'C', '%C')):
        qline(r_, f'Drift monitors used, {name}', f'=Fits!$B${base}', '0',
              f'=IF(Fits!$B${base}<2,"too few monitors with a mass and a certified value",'
              f'IF(Fits!$B${base+2}="Yes","corrected: {name} divided by the factor at each position",'
              f'"not corrected"&IF({S_ELEM}="Off"," (switched off in Setup B35)",'
              f'IF(Fits!$B${base+1}>{S_ETHR},""," (under the "&TEXT({S_ETHR},"0.0")&" percent threshold in Setup B36)"))))')
        qline(r_ + 1, f'Largest monitor bias, {name} (percent)', f'=Fits!$B${base+1}', '0.00',
              f'=IF(Fits!$B${base}<2,"",IF(Fits!$B${base+1}>{S_ETHR},"CHECK: a monitor is off by more than "'
              f'&TEXT({S_ETHR},"0.0")&" percent; if the anchors beside it and the reference pulses do not share the '
              f'change, the capsule is the cause, so list its position in Setup B37","ok"))')
    qline(144, 'Positions left out (Setup B37)', '=Fits!$B$250', '@')
    qsec(146, 'Exports without reference peaks: CO2 peaks that are reference pulses')
    qline(147, 'Designated CO2 pulses in the export', '=Fits!$B$237', '0',
          f'=IF(Fits!$B$240="Yes","check active: an export with no N2 reference peaks",'
          f'"check not active: "&IF({S_RPA}="","switched off in Setup B39",IF(COUNT({W("BI")})>0,'
          f'"the export carries reference peaks, so the sample CO2 peak is identified from them",'
          f'"no designated pulses in the export to compare with")))')
    qline(148, 'CO2 peaks rejected as reference pulses', '=Fits!$B$241', '0',
          '=IF(Fits!$B$240<>"Yes","",IF(Fits!$B$241=0,"ok","those positions made no sample CO2; Results shows '
          'no C peak and says why in Flags"))')
    # 1.4.3
    qsec(150, 'Uncertainty of %N, %C, and C:N')
    qc.cell(151, 1, '1σ on Results has five terms (Fits row 300 onwards). Empty inputs count as 0, so 1σ is '
                    'incomplete until the balance (Setup B33) and the certificates (Setup columns J and K) are '
                    'filled in. The anchors and the linearity standards have known values, so how many fall within '
                    'k x 1σ of them (k in Setup B32) is a check on the budget.').font = F_NOTE
    qline(152, 'Balance, 1σ of one weighing (mg)', '=Fits!$B$301', '0.0000',
          '=IF(Fits!$B$301="","not given in Setup B33: the weighing term is left out, so 1σ is too small for '
          'light samples","ok")')
    for r_, el, name, nrow, rep, true, inu in ((153, 'N', '%N', 303, 304, 311, 'inU N'),
                                                (157, 'C', '%C', 316, 317, 324, 'inU C')):
        qline(r_, f'Repeatability of {name} (percent, 1σ)', f'=IF(Fits!$B${rep}="","",Fits!$B${rep}*100)', '0.00',
              f'=IF(Fits!$B${nrow}<3,"fewer than three drift monitors: no 1σ for {name}","ok")')
        qline(r_ + 1, f'Trueness of {name} from the anchors (percent)', f'=IF(Fits!$B${true}="","",Fits!$B${true}*100)',
              '0.00', f'=IF(Fits!$B${true}="","fewer than two anchor analyses: no trueness term","ok")')
        for k_, rl, word in ((2, 'Anchor', 'Anchors'), (3, 'Linearity', 'Linearity standards')):
            within = f'COUNTIFS({WD_},"{rl}",{W(UX[inu])},1)'
            total = f'COUNTIFS({WD_},"{rl}",{W(UX[inu])},">=0")'
            qline(r_ + k_, f'="{word} within "&{S_K}&"σ, {name}"', f'={within}&" of "&{total}', '@',
                  f'=IF({total}=0,"",IF({within}<{total},"CHECK: "&TEXT({total}-{within},"0")'
                  f'&" outside "&{S_K}&"σ","ok"))')
    qline(161, 'Repeatability of C:N (percent, 1σ)', '=IF(Fits!$B$330="","",Fits!$B$330*100)', '0.00',
          '=IF(Fits!$B$329<3,"fewer than three drift monitors: no 1σ for C:N","ok")')
    qline(162, 'Trueness of C:N from the anchors (percent)', '=IF(Fits!$B$335="","",Fits!$B$335*100)', '0.00',
          '=IF(Fits!$B$335="","fewer than two anchor analyses: no trueness term","ok")')
    for r_, rl, word in ((163, 'Anchor', 'Anchors'), (164, 'Linearity', 'Linearity standards')):
        within = f'COUNTIFS({WD_},"{rl}",{W(UX["inU CN"])},1)'
        total = f'COUNTIFS({WD_},"{rl}",{W(UX["inU CN"])},">=0")'
        qline(r_, f'="{word} within "&{S_K}&"σ, C:N"', f'={within}&" of "&{total}', '@',
              f'=IF({total}=0,"",IF({within}<{total},"CHECK: "&TEXT({total}-{within},"0")'
              f'&" outside "&{S_K}&"σ","ok"))')

    # 1.4.4
    qsec(166, 'Rerun marks')
    qc.cell(167, 1, 'A sample is marked for a rerun when its peak lies below the linearity range or under the blank '
                    'multiple, reaches the instrument\'s largest amplitude (Setup B69), or meets a reference pulse '
                    'that is off, when its 1σ is above Setup B70, '
                    'or when its replicates differ by more than the critical range (Replicates sheet), and for a rerun '
                    'of N and C when it has neither peak. It is marked for a check when it ran before the background '
                    'settled, follows an oversized sample with a background rise (possible carryover, a check from '
                    '1.4.14), lacks one element\'s peak, or, under the Background drift model, has no reference '
                    'background and so no drift correction. The Rerun sheet lists every mark.').font = F_NOTE
    qline(168, 'Samples marked for a rerun', '=Fits!$B$351', '0',
          '=IF(Fits!$B$351=0,"ok","see the Rerun sheet")')
    qline(169, 'Samples marked for a check', '=Fits!$B$352', '0',
          '=IF(Fits!$B$352=0,"ok","see the Rerun sheet")')
    qline(170, 'N2 reference pulses off at their position', f'=COUNTIF({W(RX["poff N"])},1)', '0',
          '=IF(Fits!$B$342="","not judged: fewer than five pulses, or switched off in Setup B67","ok")')
    qline(171, 'CO2 reference pulses off at their position', f'=COUNTIF({W(RX["poff C"])},1)', '0',
          '=IF(Fits!$B$346="","not judged: fewer than five pulses, or switched off in Setup B67","ok")')
    qline(172, 'Problems of the whole run', '=IF(Fits!$B$356="none found","none","listed below")', '@',
          '=IF(Fits!$B$356="none found","ok","CHECK")')
    qc.cell(173, 1, '=IF(Fits!$B$356="none found","",Fits!$B$356)').font = F_SM

    # 1.4.13: the anchors as one material each (finding W-13), and the elemental drift correction against them
    qsec(175, 'Anchors: one material each, and a test of the elemental drift correction (1.4.13)')
    qc.cell(176, 1, 'Analyses of one material agree in their C/N peak-area ratio and their raw delta values, so anchors '
                    'of one row of the standards table that do not are more than one material, such as drift monitors '
                    'named USGS 64 Drift. The elemental drift correction does not use the anchors\' certified %N and %C, '
                    'so they test it.').font = F_NOTE
    qline(177, 'Anchors of each row read as one material', f'=IF(Fits!$B${AM_LEVEL}="","",Fits!$B${AM_LEVEL})', '@',
          f'=IF(Fits!$B${AM_LEVEL}="","not judged: no anchor material has two analyses",IF(Fits!$B${AM_LEVEL}="OK",'
          f'"ok","CHECK: the anchors of a row of the standards table do not read as one material (Fits rows {AMT} to '
          f'{AMT + NSTD - 1}); a capsule whose name holds that row\'s match text is read as that anchor, so rename it '
          f'or lengthen the match text, unless the anchors\' own measurement failed (QC rows 23 and 25)"))')
    for r_, base, inf, pct in ((178, EC0 + 1, 246, '%N'), (180, EC0 + 6, 249, '%C')):
        qline(r_, f'Anchors\' RMS deviation, {pct} as measured (percent)',
              f'=IF(Fits!$B${base + 1}="","",Fits!$B${base + 1}*100)', '0.00')
        qline(r_ + 1, f'Anchors\' RMS deviation, {pct} corrected (percent)',
              f'=IF(Fits!$B${base + 2}="","",Fits!$B${base + 2}*100)', '0.00',
              f'=IF(Fits!$B${inf}<>"Yes","not judged: the elemental drift correction is not in force for {pct}",'
              f'IF(Fits!$B${base + 4}="","not judged: fewer than two anchors with a {pct}",IF(Fits!$B${base + 4}="OK",'
              f'"ok","CHECK, "&Fits!$B${base + 4}&": the correction moves the anchors "&TEXT(Fits!$B${base + 3}*100,'
              f'"0.00")&" points further from certified {pct}; check the drift monitors beside them and Setup B37, or '
              f'set Setup B35 to Off")))')

    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.trendline import Trendline

    def feed(col, sheet=None):
        if sheet is raw:          # 1.4.10: the pulses chart reads every pulse from Raw
            return Reference(raw, min_col=raw[col + '1'].column, min_row=RA, max_row=RB)
        return Reference(wk, min_col=wk[col + '1'].column, min_row=R0, max_row=R0 + NSAMP - 1)

    def add_series(ch, xcol, ycol, name, colour, marker=True, line=False, trend=False, size=5, hollow=False,
                   sheet=None):
        sr = Series(feed(ycol, sheet), feed(xcol, sheet))
        sr.tx = SeriesLabel(v=name)
        if marker:
            sr.marker = Marker(symbol='circle', size=size)
            sr.marker.graphicalProperties.solidFill = 'FFFFFF' if hollow else colour
            sr.marker.graphicalProperties.line.solidFill = colour
        else:
            sr.marker = Marker(symbol='none')
        if line:
            sr.graphicalProperties.line.solidFill = colour
            sr.graphicalProperties.line.width = 15000
        else:
            sr.graphicalProperties.line.noFill = True
        if trend:
            sr.trendline = Trendline(trendlineType='linear', dispEq=False, dispRSqr=False)
        ch.series.append(sr)
        return sr

    def chart(anchor, title, xtitle, ytitle, log_x=False):
        ch = ScatterChart(); ch.title = title; ch.style = 2; ch.height = 7.4; ch.width = 13.5
        ch.x_axis.title = xtitle; ch.y_axis.title = ytitle
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.legend.position = 'b'
        # 1.4.13: Excel draws a title, an axis title, or the legend over the plot area when the file leaves its
        # overlay unset (Excel for Mac, 2026-10-15), so each says it does not overlay. The x axis's tick labels
        # sat on the y = 0 line, at the top of a chart whose values are all negative and across the data of
        # one that spans zero, so both axes put their labels at the low end, and the x axis is at the bottom
        # (openpyxl writes a scatter chart's x axis at the left).
        ch.title.overlay = False
        ch.x_axis.title.overlay = False; ch.y_axis.title.overlay = False
        ch.legend.overlay = False
        ch.x_axis.tickLblPos = 'low'; ch.x_axis.axPos = 'b'
        ch.y_axis.tickLblPos = 'low'
        if log_x:
            ch.x_axis.scaling.logBase = 10
        qc.add_chart(ch, anchor)
        return ch

    BLUE, ORANGE, GREEN, YELLOW, GREY, RED = '1F4E79', 'C55A11', '1B7F5A', 'B8860B', '7F7F7F', 'B2182B'
    # Working gas, beside rows 10 to 19
    for anchor, el, gas, ycol in (('K10', 'N', 'N2', 'AC'), ('T10', 'C', 'CO2', 'AD')):
        ch = chart(anchor, f'{gas} reference pulses Isodat did not designate', 'Autosampler position',
                   ('d15N' if el == 'N' else 'd13C') + ' of the pulse (permil)')
        add_series(ch, 'A', ycol, 'pulses', GREY, trend=True, sheet=raw)
    # Calibration, beside rows 21 to 28: are the unknowns bracketed by the anchors?
    for anchor, el, name in (('K25', 'N', 'd15N'), ('T25', 'C', 'd13C')):
        ch = chart(anchor, f'Calibrated {name} of the unknowns against the anchor span', 'Autosampler position',
                   f'{name} (permil)')
        add_series(ch, CH['ch pos'], CH[f'ch cal y {el}'], 'unknowns and checks', BLUE)
        add_series(ch, CH['ch pos'], CH[f'ch acc min {el}'], 'lowest anchor', ORANGE, marker=False, line=True)
        add_series(ch, CH['ch pos'], CH[f'ch acc max {el}'], 'highest anchor', ORANGE, marker=False, line=True)
    # Corrections, beside rows 30 to 45: linearity, then drift
    for anchor, el, area, name in (('K40', 'N', 'AM', 'Area 28 (Vs)'), ('T40', 'C', 'AO', 'Area 44 (Vs)')):
        ch = chart(anchor, ('Nitrogen' if el == 'N' else 'Carbon') + ' linearity: size series and unknowns',
                   name + ', log scale', ('d15N' if el == 'N' else 'd13C') + ' after blank (permil)', log_x=True)
        add_series(ch, CH[f'ch samp x {el}'], CH[f'ch samp y {el}'], 'unknowns, by size', GREY, size=4)
        add_series(ch, CH[f'ch lin x {el}'], CH[f'ch lin y {el}'], 'linearity standards', YELLOW, size=6, trend=True)
    for anchor, el in (('K55', 'N'), ('T55', 'C')):
        ch = chart(anchor, ('Nitrogen' if el == 'N' else 'Carbon') + ' drift monitors, before and after the correction',
                   'Autosampler position', ('d15N' if el == 'N' else 'd13C') + ', raw scale (permil)')
        add_series(ch, CH['ch pos'], CH[f'ch drift used {el}'], 'after linearity, in the fit', GREEN, size=6, trend=True)
        add_series(ch, CH['ch pos'], CH[f'ch drift out {el}'], 'after linearity, left out', GREEN, size=6, hollow=True)
        add_series(ch, CH['ch pos'], CH[f'ch drift after {el}'], 'after the drift correction', RED, size=4)
    # Settling, beside rows 93 to 102
    ch = chart('K96', 'Mass-28 background across the run', 'Autosampler position', 'Background (mV)')
    add_series(ch, CH['ch pos'], CH['ch bgd sample'], 'under each sample peak', GREY, size=4)
    add_series(ch, CH['ch pos'], CH['ch bgd monitor'], 'under the drift monitors', GREEN, size=6, trend=True)
    add_series(ch, CH['ch pos'], CH['ch bgd'], 'on the reference peaks', BLUE, line=True)
    add_series(ch, CH['ch pos'], CH['ch carry'], 'flagged for possible carryover', RED, size=7)
    qc.cell(3, 11, 'Charts sit beside the section they belong to. A trendline is Excel\'s own linear fit '
                   'through the plotted points; for the drift monitors in the fit it is the row-model slope '
                   'Fits reports.').font = F_NOTE

    # ================================================================ Metadata (1.4.8)
    md = wb.create_sheet('Metadata')
    md.sheet_view.showGridLines = False
    for col, w in zip('ABCDE', [22, 30, 18, 12, 10]):
        md.column_dimensions[col].width = w
    md.cell(1, 1, 'Metadata').font = F_TITLE
    md.cell(2, 1, 'The run information and the sample sheet\'s rows for this run\'s samples, as isoharness.org '
                  'received them. Unlike the rest of the workbook, the sample table holds plain values, because it '
                  'belongs to the run this workbook was built from; Results reads it only at positions whose '
                  'identifier is still the one matched.').font = F_NOTE
    md.merge_cells('A2:H2'); md.row_dimensions[2].height = 40
    md['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    md.cell(4, 1, 'Run information').font = F_H1
    info_rows = [('Run', '=IF(Setup!$B$5="","",Setup!$B$5)', None), ('File', _FILE, 'text'),
                 ('Run date', '=IF(Setup!$B$6="","",Setup!$B$6)', None),
                 ('Operator', '=IF(Setup!$B$7="","",Setup!$B$7)', None)]
    info_rows += [(k_, _INFO.get(k_, ''), 'in') for k_ in RUN_FIELDS if k_ != 'Operator']
    for r_, (k_, v_, kind_) in enumerate(info_rows, start=5):
        md.cell(r_, 1, k_).font = F_SM
        if kind_ == 'in':
            inp(md, r_, 2, v_)
            md.merge_cells(start_row=r_, start_column=2, end_row=r_, end_column=8)
            md.cell(r_, 2).alignment = Alignment(wrap_text=True, vertical='top')
        else:
            c = put(md, r_, 2, v_) if kind_ == 'text' else md.cell(r_, 2, v_); c.font = F_SM
    md.cell(5 + len(info_rows), 1, 'Operator is Setup B7; type the others here.').font = F_NOTE
    md.cell(13, 1, 'Sample sheet').font = F_H1
    if samples is None:
        md.cell(14, 1, 'No sample sheet was given.').font = F_NOTE
    else:
        sheet_rows = [('File', samples.name), ('Sheet', samples.sheet_name or '(a text file)'),
                      ('ID column', samples.id_column), ('Header row', samples.header_row),
                      ('Rows', len(samples.rows)), ('Samples matched', summary_text(_MATCH)),
                      ('Replicate groups', len(_GROUPS)), ('Notes', '; '.join(samples.notes) or 'none')]
        for r_, (k_, v_) in enumerate(sheet_rows, start=14):
            md.cell(r_, 1, k_).font = F_SM
            c = put(md, r_, 2, v_); c.font = F_SM; c.alignment = Alignment(horizontal='left')
        md.cell(MD_TOP - 2, 1, 'This run\'s samples').font = F_H1
        mhead = ['Row', 'Identifier', 'Sample ID (sheet)', 'Matched by', 'Replicate group'] + [h for _, h in _MCOLS]
        for j, h in enumerate(mhead, start=1):
            c = put(md, MD_TOP - 1, j, h); c.font = F_HDR; c.fill = FILL_HDR
            c.alignment = Alignment(horizontal='center', wrap_text=True)
            if j > 5:
                md.column_dimensions[gcl(j)].width = 14
        for q in _META:
            if q.row not in MD_ROW:
                continue
            t_ = MD_ROW[q.row]
            vals = [q.row, q.identifier, q.sample_id, q.matched_by or NOT_MATCHED, q.replicate_group]
            vals += [samples.rows[q.sheet_index][j_] if q.sheet_index is not None else None for j_, _ in _MCOLS]
            for j, v in enumerate(vals, start=1):
                c = put(md, t_, j, v); c.font = F_SM
        md.freeze_panes = f'A{MD_TOP}'

    for s in wb.worksheets:
        s.sheet_properties.tabColor = '1F4E79' if s.title in ('Results', 'QC') else None
    wb.save(OUT)
    if recalc:
        bake_values(OUT)

    return Path(OUT)
