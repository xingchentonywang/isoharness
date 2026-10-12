"""Write the correction workbook for one sequence.

The workbook carries live formulas rather than computed values, so a student can
follow every step and change a standard's accepted value without the sheet going
stale. `ea_irms.pipeline` computes the same numbers in Python; the test suite
holds the two to each other.
"""
from __future__ import annotations

from pathlib import Path

"""Build the EA-IRMS student correction template (.xlsx), formula-driven."""
import collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation


from .parse import _read_rows
from .pipeline import DEFAULT_STANDARDS, MAX_STANDARDS, MIN_OUTLIER_MONITORS, Config

VERSION = "1.4.2"
NSTD = MAX_STANDARDS      # rows in the Setup standards table (16 from 1.4.2, 13 before)


def bake_values(path) -> bool:
    """Have LibreOffice compute every formula and store the results in the file.

    openpyxl writes formulas without cached values. Excel computes them when it
    opens the file, but a preview pane, a file card, Protected View, or a viewer
    that does not calculate shows every formula cell empty. When LibreOffice is
    on the PATH this recalculates the workbook in place, keeping every formula,
    chart, validation and format, so the numbers show anywhere. Returns whether
    it ran.
    """
    import shutil
    import subprocess
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    path = Path(path)
    with tempfile.TemporaryDirectory() as tmp:
        try:
            subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx",
                            "--outdir", tmp, str(path)],
                           check=True, capture_output=True, timeout=600)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError):
            return False
        produced = Path(tmp) / path.name
        if not produced.exists() or produced.stat().st_size == 0:
            return False
        shutil.copyfile(produced, path)
    return True


def build_workbook(export_path, out_path, config: Config | None = None,
                   run_name: str | None = None, run_date: str | None = None,
                   recalc: bool = False):
    """Build the workbook from an Isodat export. Returns ``out_path``.

    With ``recalc`` the finished file is passed through :func:`bake_values` when
    LibreOffice is available, so it carries computed values as well as formulas.
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

    MAXRAW = 1000          # raw data rows 2..1001
    R0, NSAMP = 6, 120     # Work data rows 6..125
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
    r = rd(r, 'Thermo Flash EA coupled to a Delta/253 IRMS. Takes the Isodat export and returns %C, %N, C:N, '
              f'delta-13C and delta-15N, with every correction visible as a live formula. ea_irms {VERSION}.')
    r += 1
    r = rd(r, 'How to use it', F_H1)
    r = rd(r, '1.  Open the Raw sheet. Select rows 2 to 1001 and press Delete, then paste your Isodat export '
              'starting in cell A2. Do not paste over row 1: the headers there are what every other sheet looks for. '
              'The sheet ships loaded with the 2026-09-15 test run as a worked example, so delete that first.')
    r = rd(r, '2.  Open Setup. Check the standards table against the certificates and against what the sequence '
              'actually called things. A material is recognised by the text in column A appearing anywhere in '
              'Identifier 1, so "Ace Drift" catches Ace Drift1 through Ace Drift 13. Matching ignores case. '
              'When two rows match, the lower row wins, so put the more specific text lower down.')
    r = rd(r, '3.  Open Work and type sample masses in the yellow column F, in mg. Leave a cell empty to use the '
              'mass the export already carries in its Amount column. Without a mass there is no %C, %N or C:N for '
              'that row, though the isotope values still come out.')
    r = rd(r, '4.  Read QC before you read anything else. It flags check standards that missed, samples outside '
              'the calibrated size range, blanks that cannot support a blank correction, a reference gas that moved '
              'during the run, and positions whose background rose right after an oversized sample.')
    r = rd(r, '5.  Take the numbers from Results, with their uncertainties. Work is the audit trail behind them '
              'and Fits holds every fitted coefficient with its r-squared, residual scatter and, for the '
              'corrections, the test that decided whether it was applied.')
    r += 1
    r = rd(r, 'Colour', F_H1)
    r = rd(r, 'Yellow cells with blue text are yours to fill in. Grey columns are machinery and can be ignored. '
              'Green columns are the reported values. Everything else is a formula, and overtyping one breaks that '
              'row silently, so work on a copy.')
    r += 1
    r = rd(r, 'What the sheets do', F_H1)
    for a, b in [
        ('Raw',     'Paste target, unaltered. Two grey helper columns label each peak record by channel and give it a key.'),
        ('Setup',   'Everything you are meant to change: processing switches, the standards table, flag thresholds.'),
        ('Work',    'One row per autosampler position. Reads left to right: parsed peaks, then blank, linearity, '
                    'drift and calibration corrections, then elemental content.'),
        ('Fits',    'Every coefficient fitted from this run, each with n, r-squared and residual scatter.'),
        ('Results', 'The clean table to copy out or plot from.'),
        ('QC',      'Checks and flags. Read this first.'),
    ]:
        c = ws.cell(r, 2); c.value = f'{a}.  {b}'; c.font = F_SM
        c.alignment = Alignment(wrap_text=True, vertical='top'); r += 1
    r += 1
    r = rd(r, 'How the corrections work', F_H1)
    r = rd(r, 'Blank, then linearity, then drift, then calibration, in that order, each fitted from this run\'s own '
              'standards rather than from a coefficient typed in by hand. Because linearity is fitted after the blank '
              'correction, the two cannot double-count a size dependence that the blank already explains.')
    r = rd(r, 'Calibration regresses accepted values on measured values and applies that line directly. It is never '
              'fitted in the other direction and then inverted, which is where the arithmetic usually goes wrong: '
              'dividing the intercept by the slope and subtracting it shifts every value by a constant and never '
              'applies the slope at all.')
    r = rd(r, 'Drift is fitted twice, against autosampler position and against the mass-28 background on the '
              'reference peaks, and QC reports the residual scatter of both. The default is autosampler position. '
              'Residual scatter cannot decide between the two, since both were fitted to the same drift standards, so '
              'run labelled check standards and let their offsets decide.')
    r += 1
    r = rd(r, 'Before trusting any drift correction, read the settling section of QC. What looks like drift is often '
              'the instrument still coming to rest: the mass-28 background decays towards a floor over the first '
              'several dozen positions and the standards follow it down, then both stop. QC finds the position where '
              'the background settles, as the first three readings in a row near the floor, reports the drift slope '
              'separately before and after it, and says which pattern the run shows. A spike later in the run is '
              'carryover from an oversized sample, not settling, and does not move that position. A straight line in '
              'position is the wrong shape for a transient, under-correcting the early positions and over-correcting '
              'the late ones. Setup B14 excludes the unsettled positions from the drift fit. The real fix is upstream: '
              'wait for the background to reach its floor before the first real sample, and there is no drift left '
              'to correct.')
    r += 1
    r = rd(r, 'QC also lists the scatter of the drift standards after each correction in turn. A correction that does '
              'not lower that number is not earning its place in the run you are looking at.')
    r += 1
    r = rd(r, 'Only corrections the run supports', F_H1)
    r = rd(r, 'On Auto, which is the default, a correction is applied only when this run supports it: the blank when '
              'at least three blank capsules agree to within 2 per mil, linearity and drift when the fitted slope '
              'differs from zero at the significance level in Setup B15 (a two-sided t-test of the slope). That '
              'follows the FIRMS Good Practice Guide for IRMS (Dunn and Carter, 2018, section 6.4), which asks that '
              'a correction be shown necessary before it is applied. With five or six standards the test has little '
              'power, so a slope that fails it is unproven rather than disproven; set the switch to On to apply it '
              'regardless. Setup B11 sets linearity for both elements unless B18 sets carbon on its own. '
              'Each drift monitor is also compared against the line through the others, and one that '
              'sits far off is flagged rather than removed: if the capsule was the cause, list its position in Setup '
              'B16 or B17 and that element\'s fit leaves it out.')
    r += 1
    r = rd(r, 'Uncertainty', F_H1)
    r = rd(r, 'Results carries a standard uncertainty u for every calibrated value and U = k x u with k from Setup '
              'B32. The budget follows section 6.5 of the guide: the 1-sigma certificate uncertainty of each anchor '
              'material (Setup columns H and I), the scatter of its replicates in this run, and the repeatability of '
              'one sample analysis (Setup B30 and B31, or the pooled anchor scatter when those are empty), each '
              'propagated through the calibration line as a Kragten spreadsheet would. It leaves out the blank, '
              'linearity and drift corrections and day-to-day reproducibility, so it is a floor. A control chart of '
              'an in-house standard is what supplies the rest.')
    r = rd(r, 'Elemental content comes from a regression of peak area on micrograms of C or N across every standard '
              'with both a known mass and a known composition. The intercept of that line is the area blank and is '
              'subtracted, so the k-factor is not forced through the origin.')
    r += 1
    r = rd(r, 'Elemental drift (1.4.2)', F_H1)
    r = rd(r, 'That k-factor is one line for the whole run. When the sensitivity of the detector moves during the '
              'run, the drift monitors read high or low in %N and %C by position, and so do the samples around '
              'them. Each monitor\'s measured over certified %N (and %C) gives a factor at its position; between '
              'monitors the factor is interpolated in position, beyond the first and last it is held, and every '
              'position\'s micrograms are divided by the factor at its position. On Auto the correction is applied '
              'when a monitor is further than Setup B36 (percent) from its certified value; On applies it always, '
              'Off never. A monitor whose capsule was weighed or entered wrongly would be carried into the samples, '
              'so list such positions in Setup B37 and they are left out. Results reports the corrected %N, %C and '
              'C:N, with the factors and the values as measured in its last columns; the isotope values are not '
              'touched.')
    r += 1
    r = rd(r, 'Exports without reference peaks (1.4.2)', F_H1)
    r = rd(r, 'An export cut down to the sample peaks (Peak Nr 3 and 4) carries no reference peaks to count from. '
              'Where a position made no sample CO2, the first exported CO2 peak is then a reference pulse that '
              'Isodat did not flag, and it would be read as the sample. Such a pulse matches the designated pulses '
              'the export does carry (at positions with no sample gas) in amplitude, within Setup B39, and reads '
              'the working-gas value in d13C, within Setup B40, which no sample does. The template rejects it, '
              'reports the position without a carbon value, and says so in Flags. Leave B39 empty to switch the '
              'check off.')
    r += 1
    r = rd(r, 'Known limits', F_H1)
    r = rd(r, 'The outlier test on the drift monitors needs five of them in the fit; with four it has one degree of '
              'freedom and flags residuals of a hundredth of a per mil, so from 1.4.2 it is not run below five.')
    r = rd(r, 'Linearity is fitted as a straight line. On the 2026-09-15 run the fitted slope moved 36 percent when '
              'the smallest standard was included, so the response is curved across the full size range and a sample '
              'far from the standards will carry that curvature as an error. The size flags in QC mark those rows.')
    r = rd(r, 'A blank correction is only as good as the blanks. Two blank capsules whose delta-15N differ by 22 per '
              'mil cannot constrain one, and QC says so rather than letting the correction run quietly. Setting the '
              'switch to Auto applies the correction only when at least three blanks agree to within 2 per mil. A '
              'sample barely larger than the blank is flagged separately, because there the mass balance amplifies '
              'the blank\'s uncertainty instead of removing it, whatever the blanks look like.')
    r = rd(r, 'Precision quoted from the drift standards understates the real figure, because those are the material '
              'the drift correction was fitted to. QC reports both that number and the scatter of any check standards '
              'and repeated unknowns, which is the honest one.')
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
    raw.cell(1, 20, 'grey helper columns, computed').font = F_NOTE

    header, records = _read_rows(SRC)
    order = [header.index(h) if h in header else None for h in RAW_HDR]
    for i, record in enumerate(records, start=1):
        for j, src_j in enumerate(order):
            v = record[src_j] if src_j is not None and src_j < len(record) else ''
            if v is None or v == '':
                continue
            if j in (0, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14) and isinstance(v, str):
                try: v = float(v)                 # csv exports and BGD 28 come as text
                except ValueError: pass
            c = raw.cell(i + 1, j + 1, v); c.font = F_SM
    NRAW_USED = len(records)

    for i in range(2, MAXRAW + 2):
        c = raw.cell(i, 18, f'=IF($D{i}="","",IF($F{i}<>"","N",IF($J{i}<>"","C","")))')
        c.font = F_SM; c.fill = FILL_HELP
        c = raw.cell(i, 19, f'=IF($R{i}="","",$A{i}&"|"&$R{i}&"|"&$D{i})')
        c.font = F_SM; c.fill = FILL_HELP
    raw.freeze_panes = 'A2'
    for col, w in zip('ABCDEFGHIJKLMNOPQRS',
                      [6,16,11,8,9,9,8,9,8,9,9,11,12,11,11,11,3,9,13]):
        raw.column_dimensions[col].width = w

    # ================================================================ Setup
    st = wb.create_sheet('Setup')
    st.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGH', [34, 14, 13, 13, 10, 10, 34, 4]):
        st.column_dimensions[col].width = w

    def lab(ws_, r_, txt, style=F_BODY, col=1):
        c = ws_.cell(r_, col, txt); c.font = style; return c

    def inp(ws_, r_, col, val, fmt=None):
        c = ws_.cell(r_, col, val); c.font = F_IN; c.fill = FILL_IN; c.border = BOX
        if fmt: c.number_format = fmt
        return c

    st.cell(1, 1, 'Setup').font = F_TITLE
    st.cell(2, 1, 'Yellow cells are yours. Everything else on this sheet is a label.').font = F_NOTE

    lab(st, 4, 'Run information', F_H1)
    for i, (k, v) in enumerate([('Run name', _RUN), ('Run date', _DATE), ('Analyst', '')]):
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
    lab(st, 14, 'Leave empty to use every position. QC suggests a value. The calibration anchors are '
                'never excluded, since they are always run early.', F_NOTE, 7)
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
    THR = [
        ('Check standard offset (permil)', _CFG.check_offset_permil,
                                              'Flag a check standard further than this from its accepted value.'),
        ('Minimum Area 28', _CFG.min_area_n or '',
                                              'Leave empty to use the smallest linearity standard in the run.'),
        ('Minimum Area 44', _CFG.min_area_c or '',
                                              'Leave empty to use the smallest linearity standard in the run.'),
        ('Maximum Area 28', _CFG.max_area_n or '',
                                              'Leave empty to use the largest linearity standard in the run.'),
        ('Maximum Area 44', _CFG.max_area_c or '',
                                              'Leave empty to use the largest linearity standard in the run.'),
        ('Minimum peak area, x the blank', _CFG.min_blank_multiple or '',
                                              'Below this multiple of the blank the blank correction amplifies '
                                              'rather than corrects. Leave empty to skip the check.'),
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
        ('Sample repeatability, d15N (1 SD, permil)', _CFG.u_sample_n or '',
                                               'One analysis of one sample. Leave empty to use the pooled scatter '
                                               'of the anchor replicates in this run; better is the long-term SD '
                                               'of a matrix-matched in-house standard.'),
        ('Sample repeatability, d13C (1 SD, permil)', _CFG.u_sample_c or '', ''),
        ('Coverage factor k', _CFG.coverage_k,
                                               'Expanded uncertainty U = k x u. k = 2 is about 95 percent only when '
                                               'the replicates behind u are many; with a handful, take k from '
                                               'Student\'s t.'),
    ]
    for i, (k, v, note) in enumerate(UNC):
        lab(st, 30 + i, k, F_SM); inp(st, 30 + i, 2, v); lab(st, 30 + i, note, F_NOTE, 7)
    st.cell(30, 2).number_format = '0.000'; st.cell(31, 2).number_format = '0.000'
    st.cell(32, 2).number_format = '0.0'

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
    lab(st, 37, 'Comma-separated. For a monitor whose capsule was weighed or entered wrongly, which the '
                'instrument did not cause: the anchors beside it and the reference pulses say which.', F_NOTE, 7)
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

    lab(st, 42, 'Standards', F_H1)
    st.cell(43, 1, 'Accepted values below are the published USGS values as of 2026-09-16, and u is the '
                   '1-sigma uncertainty the producer lists. Check both against your certificates.').font = F_NOTE
    SHDR = ['Match text in Identifier 1', 'Role', 'Accepted d13C', 'Accepted d15N', '%C', '%N', 'Material',
            'u(d13C)', 'u(d15N)']
    for j, h in enumerate(SHDR, start=1):
        c = st.cell(44, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    STD_ROWS = [tuple(r) + (None,) * (9 - len(r)) for r in _CFG.standards]
    SFIRST, SLAST = 45, 45 + NSTD - 1        # 16 slots from 1.4.2, the last three left empty for the student
    for i in range(NSTD):
        for j in range(1, 10):
            c = st.cell(SFIRST + i, j); c.font = F_IN; c.fill = FILL_IN; c.border = BOX
            if j in (3, 4, 5, 6, 8, 9): c.number_format = '0.00'
        if i < len(STD_ROWS):
            for j, v in enumerate(STD_ROWS[i][:9], start=1):
                st.cell(SFIRST + i, j, v if v != '' else None)
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

    S_BLANK, S_LIN, S_DRIFT, S_DM = 'Setup!$B$10', 'Setup!$B$11', 'Setup!$B$12', 'Setup!$B$13'
    S_LINC = 'IF(Setup!$B$18<>"",Setup!$B$18,Setup!$B$11)'   # carbon switch, empty = follow B11 (1.4.1)
    S_FIRST, S_ALPHA, S_EXN, S_EXC = 'Setup!$B$14', 'Setup!$B$15', 'Setup!$B$16', 'Setup!$B$17'
    S_THR, S_MINN, S_MINC, S_MAXN, S_MAXC = ('Setup!$B$20', 'Setup!$B$21', 'Setup!$B$22',
                                             'Setup!$B$23', 'Setup!$B$24')
    S_MULT, S_SETTLE, S_TOUT = 'Setup!$B$25', 'Setup!$B$26', 'Setup!$B$27'
    S_CARRY = 'Setup!$B$28'
    S_USN, S_USC, S_K = 'Setup!$B$30', 'Setup!$B$31', 'Setup!$B$32'
    S_ELEM, S_ETHR, S_EXE = 'Setup!$B$35', 'Setup!$B$36', 'Setup!$B$37'      # 1.4.2
    S_RPA, S_RPD = 'Setup!$B$39', 'Setup!$B$40'                               # 1.4.2
    SMATCH = f'Setup!$A${SFIRST}:$A${SLAST}'
    SROLE  = f'Setup!$B${SFIRST}:$B${SLAST}'
    SD13   = f'Setup!$C${SFIRST}:$C${SLAST}'
    SD15   = f'Setup!$D${SFIRST}:$D${SLAST}'
    SPC    = f'Setup!$E${SFIRST}:$E${SLAST}'
    SPN    = f'Setup!$F${SFIRST}:$F${SLAST}'
    SU13   = f'Setup!$H${SFIRST}:$H${SLAST}'
    SU15   = f'Setup!$I${SFIRST}:$I${SLAST}'
    for col, w in zip('ABCDEFGHI', [40, 14, 13, 13, 10, 10, 34, 9, 9]):
        st.column_dimensions[col].width = w

    # ================================================================ Work
    wk = wb.create_sheet('Work')
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
    GROUPS = [('A','D','identification'),('E','G','mass'),('H','K','peak lookup'),('L','R','as measured'),
              ('S','V','from the standards table'),('W','AA','nitrogen corrections'),
              ('AB','AF','carbon corrections'),('AG','AL','elemental content'),
              ('AM','BH','fit inputs and helpers, ignore'),('BI','BJ','instrument settling'),
              ('BK','BR','fit inputs and helpers, ignore'),('BS','BW','reference pulses'),
              ('BX','CK','fit inputs and helpers, ignore'),('CL','CO','uncertainty, 1 SD and k x SD'),
              ('CP','DE','fit inputs and helpers, ignore'),(CH['ch pos'],CH['ch carry'],'chart feeds, ignore'),
              (HX['pos used'],HX['settle start'],'carryover and settling helpers, ignore'),
              (EX['desig x'],EX['elem factor C'],'reference-pulse and elemental drift helpers, ignore'),
              (EX['%N corrected'],EX['C:N mass corrected'],'elemental content, drift-corrected')]
    for a, b, t in GROUPS:
        wk.cell(4, wk[a + '1'].column, t).font = F_GRP
        wk.cell(4, wk[a + '1'].column).fill = FILL_GRP
    for name, col, width, kind in COLS:
        j = wk[col + '1'].column
        c = wk.cell(5, j, name); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True, vertical='bottom')
        wk.column_dimensions[col].width = width
    wk.cell(1, 1, 'Work').font = F_TITLE
    wk.cell(2, 1, 'One row per autosampler position. Type sample masses in the yellow column F, in mg. '
                  'Everything else is a formula. Grey columns are machinery.').font = F_NOTE

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
        f['J'] = (f'=IF($B{i}="","",SUMPRODUCT(MIN(({rawA}=$A{i})*({rawR}="C")*{rawD}'
                  f'+(1-({rawA}=$A{i})*({rawR}="C"))*9999)))')
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
                          f'ABS(INDEX({rawJ},MATCH({ckey},{rawS},0))-${PA}{i})<={S_RPA}*${PA}{i},'
                          f'ABS(INDEX({rawO},MATCH({ckey},{rawS},0))-Fits!$B$239)<={S_RPD}),1,0)))))')
        f['L'] = pullN(i, rawH); f['M'] = pullN(i, rawF); f['N'] = pullN(i, rawG); f['O'] = pullN(i, rawN)
        f['P'] = pullC(i, rawK); f['Q'] = pullC(i, rawJ); f['R'] = pullC(i, rawO)
        f['S'] = stdpull(i, SD15); f['T'] = stdpull(i, SD13)
        f['U'] = stdpull(i, SPN);  f['V'] = stdpull(i, SPC)
        f['W'] = (f'=IF($O{i}="","",IF(OR(Fits!$C$5<>"Yes",$D{i}="Blank",$L{i}<=Fits!$B$6),'
                  f'$O{i},($L{i}*$O{i}-Fits!$B$6*Fits!$B$7)/($L{i}-Fits!$B$6)))')
        f['X'] = f'=IF($W{i}="","",$W{i}-Fits!$B$20*($L{i}-Fits!$B$16))'
        f['Y'] = f'=IF($W{i}="","",IF(OR(Fits!$B$38="Row",$BI{i}=""),$A{i},$BI{i}))'
        f['Z'] = f'=IF($X{i}="","",$X{i}-Fits!$B$39*($Y{i}-Fits!$B$40))'
        f['AA'] = f'=IF(OR($Z{i}="",Fits!$B$55<2),"",Fits!$B$56*$Z{i}+Fits!$B$57)'
        f['AB'] = (f'=IF($R{i}="","",IF(OR(Fits!$C$9<>"Yes",$D{i}="Blank",$P{i}<=Fits!$B$10),'
                   f'$R{i},($P{i}*$R{i}-Fits!$B$10*Fits!$B$11)/($P{i}-Fits!$B$10)))')
        f['AC'] = f'=IF($AB{i}="","",$AB{i}-Fits!$B$28*($P{i}-Fits!$B$24))'
        f['AD'] = f'=IF($AB{i}="","",IF(OR(Fits!$B$50="Row",$BI{i}=""),$A{i},$BI{i}))'
        f['AE'] = f'=IF($AC{i}="","",$AC{i}-Fits!$B$51*($AD{i}-Fits!$B$52))'
        f['AF'] = f'=IF(OR($AE{i}="",Fits!$B$62<2),"",Fits!$B$63*$AE{i}+Fits!$B$64)'
        f['AG'] = f'=IF(OR($L{i}="",Fits!$B$69<3),"",($L{i}-Fits!$B$71)/Fits!$B$70)'
        f['AH'] = f'=IF(OR($P{i}="",Fits!$B$75<3),"",($P{i}-Fits!$B$77)/Fits!$B$76)'
        f['AI'] = f'=IF(OR($AG{i}="",$G{i}="",$G{i}=0),"",$AG{i}/($G{i}*1000)*100)'
        f['AJ'] = f'=IF(OR($AH{i}="",$G{i}="",$G{i}=0),"",$AH{i}/($G{i}*1000)*100)'
        f['AK'] = f'=IF(OR($AG{i}="",$AH{i}="",$AG{i}<=0),"",($AH{i}/{AMC})/($AG{i}/{AMN}))'
        f['AL'] = f'=IF(OR($AG{i}="",$AH{i}="",$AG{i}<=0),"",$AH{i}/$AG{i})'
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
        f['BI'] = (f'=IF($B{i}="","",IFERROR(INDEX({rawG},MATCH($A{i}&"|N|"'
                   f'&SUMPRODUCT(MAX(({rawA}=$A{i})*({rawR}="N")*({rawI}=1)*{rawD})),{rawS},0)),""))')
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
        f['CF'] = (f'=IF($AU{i}="","",($AV{i}-(Fits!$B$56*$AU{i}+Fits!$B$57))'
                   f'-Fits!$B$56*($AU{i}-Fits!$B$133))')
        f['CG'] = f'=IF($AU{i}="","",$AU{i}^2)'
        f['CH'] = f'=IF($AW{i}="","",1)'
        f['CI'] = f'=IF($AW{i}="","",$AW{i}-Fits!$B$153)'
        f['CJ'] = (f'=IF($AW{i}="","",($AX{i}-(Fits!$B$63*$AW{i}+Fits!$B$64))'
                   f'-Fits!$B$63*($AW{i}-Fits!$B$153))')
        f['CK'] = f'=IF($AW{i}="","",$AW{i}^2)'
        # Standard uncertainty of the calibrated value: the sample's own
        # repeatability through the slope, plus each anchor material's accepted
        # value and measured mean through their exact effect on the fitted line.
        def unc(xcol, nrow, srow, irow, mrow, sxxrow, urow, t0):
            T = lambda col: f'Fits!${col}${t0}:${col}${t0+NSTD-1}'      # noqa: E731
            dx = f'(${xcol}{i}-Fits!$B${mrow})'
            return (f'=IF(OR(${xcol}{i}="",Fits!$B${nrow}<3,Fits!$B${mrow}="",Fits!$B${sxxrow}="",'
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
    ft.cell(2, 1, 'Every coefficient below is fitted from this run. Nothing here is typed in by hand.').font = F_NOTE

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

    sec(54, 'Calibration, nitrogen: accepted regressed on measured')
    fit(55, 'n anchor measurements', f'=COUNT({W("AU")})', '0')
    fit(56, 'slope',      f'=IFERROR(SLOPE({W("AV")},{W("AU")}),"")')
    fit(57, 'intercept',  f'=IFERROR(INTERCEPT({W("AV")},{W("AU")}),"")')
    fit(58, 'r-squared',  f'=IFERROR(RSQ({W("AV")},{W("AU")}),"")', '0.00000')
    fit(59, 'residual SD (permil)', f'=IFERROR(STEYX({W("AV")},{W("AU")}),"")', '0.000',
            'Applied as accepted = slope x measured + intercept. Never inverted.')

    sec(61, 'Calibration, carbon: accepted regressed on measured')
    fit(62, 'n anchor measurements', f'=COUNT({W("AW")})', '0')
    fit(63, 'slope',      f'=IFERROR(SLOPE({W("AX")},{W("AW")}),"")')
    fit(64, 'intercept',  f'=IFERROR(INTERCEPT({W("AX")},{W("AW")}),"")')
    fit(65, 'r-squared',  f'=IFERROR(RSQ({W("AX")},{W("AW")}),"")', '0.00000')
    fit(66, 'residual SD (permil)', f'=IFERROR(STEYX({W("AX")},{W("AW")}),"")', '0.000')

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
    BIR = W('BI'); BJR = W('BJ'); AR_ = W('A')
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
    def caltable(r0, hdr, xcol, onecol, devcol, rescol, sqcol, ucol_setup):
        sec(r0 - 1, hdr)
        heads = ['Material', 'n', 'sum x', 'sum x^2', 'mean measured', 'SD', 'u(mean)',
                 'S1', 'S2', 'u(accepted)']
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
                f'=IF($B{i}>1,$F{i}/SQRT($B{i}),0)',
                f'=SUMIF({W("C")},{sr},{W(devcol)})',
                f'=SUMIF({W("C")},{sr},{W(rescol)})',
                f'=IF(ISNUMBER({ucol_setup}${sr}),{ucol_setup}${sr},0)',
            ]
            for j, v in enumerate(vals, start=1):
                c = ft.cell(i, j, v); c.font = F_SM
                c.number_format = '0' if j == 2 else '0.0000'
        ft.cell(r0 + NSTD + 1, 1, 'One row per line of the standards table. n, mean and SD are of the anchor '
                            'replicates after drift correction; S1 and S2 are the sums the uncertainty '
                            'weights need; u(accepted) is the 1-sigma certificate value from Setup.').font = F_NOTE

    sec(132, 'Calibration, nitrogen: statistics behind the uncertainty')
    fit(133, 'mean measured, all anchors', f'=IFERROR(AVERAGE({W("AU")}),"")', '0.0000')
    fit(134, 'Sxx',                        f'=IFERROR(DEVSQ({W("AU")}),"")', '0.0000')
    fit(135, 'pooled SD of anchor replicates',
        f'=IFERROR(SQRT(SUMPRODUCT(($B${UT_N}:$B${UT_N+NSTD-1}-1)*($B${UT_N}:$B${UT_N+NSTD-1}>1)*$F${UT_N}:$F${UT_N+NSTD-1}^2)'
        f'/SUMPRODUCT(($B${UT_N}:$B${UT_N+NSTD-1}-1)*($B${UT_N}:$B${UT_N+NSTD-1}>1))),"")', '0.000',
        'Within-material scatter of the anchors, pooled. Stands in for the repeatability of a single '
        'sample analysis when Setup gives none.')
    fit(136, 'sample repeatability in use', f'=IF({S_USN}<>"",{S_USN},$B$135)', '0.000')
    fit(137, 'its source', f'=IF({S_USN}<>"","Setup","pooled anchor replicates")', '@')
    fit(138, 'accepted d15N, lowest anchor',  f'=IFERROR(MIN({W("AV")}),"")', '0.00')
    fit(139, 'accepted d15N, highest anchor', f'=IFERROR(MAX({W("AV")}),"")', '0.00',
            'A sample outside this range is extrapolated beyond the anchors and is flagged.')
    caltable(UT_N - 1, f'Uncertainty table, nitrogen (rows {UT_N} to {UT_N + NSTD - 1})', 'AU', 'CD', 'CE', 'CF', 'CG', 'Setup!$I')

    sec(152, 'Calibration, carbon: statistics behind the uncertainty')
    fit(153, 'mean measured, all anchors', f'=IFERROR(AVERAGE({W("AW")}),"")', '0.0000')
    fit(154, 'Sxx',                        f'=IFERROR(DEVSQ({W("AW")}),"")', '0.0000')
    fit(155, 'pooled SD of anchor replicates',
        f'=IFERROR(SQRT(SUMPRODUCT(($B${UT_C}:$B${UT_C+NSTD-1}-1)*($B${UT_C}:$B${UT_C+NSTD-1}>1)*$F${UT_C}:$F${UT_C+NSTD-1}^2)'
        f'/SUMPRODUCT(($B${UT_C}:$B${UT_C+NSTD-1}-1)*($B${UT_C}:$B${UT_C+NSTD-1}>1))),"")', '0.000')
    fit(156, 'sample repeatability in use', f'=IF({S_USC}<>"",{S_USC},$B$155)', '0.000')
    fit(157, 'its source', f'=IF({S_USC}<>"","Setup","pooled anchor replicates")', '@')
    fit(158, 'accepted d13C, lowest anchor',  f'=IFERROR(MIN({W("AX")}),"")', '0.00')
    fit(159, 'accepted d13C, highest anchor', f'=IFERROR(MAX({W("AX")}),"")', '0.00')
    caltable(UT_C - 1, f'Uncertainty table, carbon (rows {UT_C} to {UT_C + NSTD - 1})', 'AW', 'CH', 'CI', 'CJ', 'CK', 'Setup!$H')

    # ---- working gas, from the pulses Isodat did not designate --------------
    sec(199, 'Reference pulses: did the working-gas measurement hold across the run?')
    ft.cell(200, 4, 'The designated reference pulse carries the assigned value by definition, so only the '
                    'other pulses can show movement. Their d values are relative to the designated pulse '
                    'of the same acquisition.').font = F_NOTE
    for base, col, name, xcol in ((201, 'BT', 'N2', 'CQ'), (208, 'BW', 'CO2', 'CR')):
        R = W(col); A_ = W('A')
        fit(base,     f'{name} pulses read',            f'=COUNT({R})', '0')
        fit(base + 1, 'mean d, relative to the designated pulse', f'=IFERROR(AVERAGE({R}),"")', '0.000')
        fit(base + 2, 'SD (permil)',                     f'=IFERROR(STDEV({R}),"")', '0.000')
        fit(base + 3, 'range (permil)',                  f'=IF($B${base}<2,"",MAX({R})-MIN({R}))', '0.000')
        fit(base + 4, 'trend, permil per position',      f'=IFERROR(SLOPE({R},{A_}),"")', '0.00000')
        fit(base + 5, 'p of the trend, two-sided',
            f'=IFERROR(TDIST(ABS($B${base+4}/(STEYX({R},{A_})/SQRT(DEVSQ({W(xcol)})))),$B${base}-2,2),"")',
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

    fit(217, 'check standards with a calibrated value',
        f'=COUNTIFS({W("D")},"Check",{W("AA")},">-1E+30")+COUNTIFS({W("D")},"Check",{W("AF")},">-1E+30")'
        f'-COUNTIFS({W("D")},"Check",{W("AA")},">-1E+30",{W("AF")},">-1E+30")', '0',
        'Materials with the Check role that were measured in this run. Zero means nothing independent '
        'of the anchors tests the calibration.')

    # ================================================================ Results
    rs = wb.create_sheet('Results')
    RHDR = [('Row', 6), ('Identifier', 18), ('Type', 10), ('Mass (mg)', 10), ('%C', 8), ('%N', 8),
            ('C:N molar', 10), ('C:N mass', 10), ('d13C (permil)', 12), ('d15N (permil)', 12), ('Flags', 54),
            ('u(d13C)', 9), ('U(d13C)', 9), ('u(d15N)', 9), ('U(d15N)', 9),
            ('Elemental drift factor N', 10), ('Elemental drift factor C', 10),
            ('%C as measured', 9), ('%N as measured', 9)]
    rs.cell(1, 1, 'Results').font = F_TITLE
    rs.cell(2, 1, 'Read QC before using these. Isotope values are on the scale set by the anchor standards '
                  'in Setup; %C, %N and C:N need a mass in Work column F or G, and from 1.4.2 they carry the '
                  'elemental drift correction when it is in force (Fits row 243 onwards), with the factors and '
                  'the values as measured in the last four columns. u is the standard uncertainty '
                  'from the scale calibration and the sample repeatability (Fits rows 132 onwards), and U is '
                  'k times u with k from Setup B32; neither includes the blank, linearity or drift '
                  'corrections or day-to-day reproducibility.').font = F_NOTE
    for j, (h, w_) in enumerate(RHDR, start=1):
        c = rs.cell(4, j, h); c.font = F_HDR; c.fill = FILL_HDR
        c.alignment = Alignment(horizontal='center', wrap_text=True)
        rs.column_dimensions[gcl(j)].width = w_
    rs.cell(4, 13, f'="U(d13C), k="&TEXT({S_K},"0.0")').font = F_HDR
    rs.cell(4, 15, f'="U(d15N), k="&TEXT({S_K},"0.0")').font = F_HDR
    for n in range(NSAMP):
        i, k = 5 + n, R0 + n
        PCc, PNc = EX['%C corrected'], EX['%N corrected']
        CNm, CNw = EX['C:N molar corrected'], EX['C:N mass corrected']
        vals = [f'=IF(Work!$B{k}="","",Work!$A{k})', f'=IF(Work!$B{k}="","",Work!$B{k})',
                f'=IF(Work!$B{k}="","",Work!$D{k})', f'=IF(Work!$G{k}="","",Work!$G{k})',
                f'=IF(Work!${PCc}{k}="","",Work!${PCc}{k})', f'=IF(Work!${PNc}{k}="","",Work!${PNc}{k})',
                f'=IF(Work!${CNm}{k}="","",Work!${CNm}{k})', f'=IF(Work!${CNw}{k}="","",Work!${CNw}{k})',
                f'=IF(Work!$AF{k}="","",Work!$AF{k})', f'=IF(Work!$AA{k}="","",Work!$AA{k})',
                (f'=IF(Work!$B{k}="","",'
                 f'IF(Work!$L{k}="","no N peak; ","")'
                 f'&IF(Work!$P{k}="","no C peak; ","")'
                 f'&IF(Work!${EX["rej C"]}{k}=1,"CO2 peak rejected as a reference pulse; ","")'
                 f'&IF(Work!$L{k}="","",IF(Fits!$B$81<=0,"",'
                 f'IF(Work!$L{k}<Fits!$B$81,"N below linearity range; ",'
                 f'IF(Fits!$B$97<=0,"",IF(Work!$L{k}>Fits!$B$97,"N above linearity range; ","")))))'
                 f'&IF(Work!$P{k}="","",IF(Fits!$B$82<=0,"",'
                 f'IF(Work!$P{k}<Fits!$B$82,"C below linearity range; ",'
                 f'IF(Fits!$B$104<=0,"",IF(Work!$P{k}>Fits!$B$104,"C above linearity range; ","")))))'
                 f'&IF(OR(Work!$D{k}="Sample",Work!$D{k}="Check"),IF(Work!$AA{k}="","",IF(Fits!$B$138="","",'
                 f'IF(AND(Work!$AA{k}>=Fits!$B$138,Work!$AA{k}<=Fits!$B$139),"","d15N outside anchor range; "))),"")'
                 f'&IF(OR(Work!$D{k}="Sample",Work!$D{k}="Check"),IF(Work!$AF{k}="","",IF(Fits!$B$158="","",'
                 f'IF(AND(Work!$AF{k}>=Fits!$B$158,Work!$AF{k}<=Fits!$B$159),"","d13C outside anchor range; "))),"")'
                 f'&IF(Work!$D{k}<>"Check","",IF(Work!$AA{k}="","",IF(Work!$S{k}="","",'
                 f'IF(ABS(Work!$AA{k}-Work!$S{k})<={S_THR},"",'
                 f'"check d15N off by "&TEXT(Work!$AA{k}-Work!$S{k},"+0.00;-0.00")&"; "))))'
                 f'&IF(Work!$D{k}<>"Check","",IF(Work!$AF{k}="","",IF(Work!$T{k}="","",'
                 f'IF(ABS(Work!$AF{k}-Work!$T{k})<={S_THR},"",'
                 f'"check d13C off by "&TEXT(Work!$AF{k}-Work!$T{k},"+0.00;-0.00")&"; "))))'
                 f'&IF(Work!$D{k}<>"Drift","",IF(Work!$BX{k}=1,"excluded from N drift fit; ",'
                 f'IF(Work!$CB{k}="","",IF(ABS(Work!$CB{k})>{S_TOUT},"N drift monitor outlier; ",""))))'
                 f'&IF(Work!$D{k}<>"Drift","",IF(Work!$BY{k}=1,"excluded from C drift fit; ",'
                 f'IF(Work!$CC{k}="","",IF(ABS(Work!$CC{k})>{S_TOUT},"C drift monitor outlier; ",""))))'
                 f'&IF(Work!$D{k}="Blank","",IF(Work!$L{k}="","",IF({S_MULT}="","",IF(Fits!$B$6<=0,"",'
                 f'IF(Work!$L{k}>=Fits!$B$6*{S_MULT},"",'
                 f'"N peak only "&TEXT(Work!$L{k}/Fits!$B$6,"0.0")&"x the blank; ")))))'
                 f'&IF(Work!$D{k}="Blank","",IF(Work!$P{k}="","",IF({S_MULT}="","",IF(Fits!$B$10<=0,"",'
                 f'IF(Work!$P{k}>=Fits!$B$10*{S_MULT},"",'
                 f'"C peak only "&TEXT(Work!$P{k}/Fits!$B$10,"0.0")&"x the blank; ")))))'
                 f'&IF(AND(Work!$D{k}="Drift",Work!${EX["excl E"]}{k}=1),"excluded from elemental correction; ","")'
                 f'&IF(Work!$A{k}>=Fits!$B$86,"","ran before the background settled; ")'
                 f'&IF(Work!${HX["carry flag"]}{k}=1,"possible carryover from an oversized sample; ",""))'),
                f'=IF(Work!$CN{k}="","",Work!$CN{k})', f'=IF(Work!$CO{k}="","",Work!$CO{k})',
                f'=IF(Work!$CL{k}="","",Work!$CL{k})', f'=IF(Work!$CM{k}="","",Work!$CM{k})',
                f'=IF(Work!$B{k}="","",Work!${EX["elem factor N"]}{k})',
                f'=IF(Work!$B{k}="","",Work!${EX["elem factor C"]}{k})',
                f'=IF(Work!$AJ{k}="","",Work!$AJ{k})', f'=IF(Work!$AI{k}="","",Work!$AI{k})']
        for j, v in enumerate(vals, start=1):
            c = rs.cell(i, j, v); c.font = F_SM
            if j in (5, 6, 7, 8, 18, 19): c.number_format = '0.00'
            if j in (9, 10): c.number_format = '0.00'
            if j == 4: c.number_format = '0.000'
            if j in (12, 13, 14, 15): c.number_format = '0.000'
            if j in (16, 17): c.number_format = '0.0000'
            if j in (5, 6, 7, 8, 9, 10, 12, 13, 14, 15): c.fill = FILL_OUT
    rs.freeze_panes = 'A5'

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
    qc.cell(11, 1, 'The designated pulse carries the assigned value by definition, so it can never move. '
                   'An export cut down to the sample peaks carries no pulses at all, and this section then '
                   'says so rather than reporting a number.').font = F_NOTE
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
          '=IF(Fits!$B$55<2,"CHECK: need at least two anchor measurements","ok")')
    qline(23, 'Nitrogen calibration slope', '=Fits!$B$56', '0.0000',
          '=IF(Fits!$B$56="","",IF(ABS(Fits!$B$56-1)>0.05,"CHECK: stretch factor far from 1","ok"))')
    qline(24, 'Carbon anchors used', '=Fits!$B$62', '0',
          '=IF(Fits!$B$62<2,"CHECK: need at least two anchor measurements","ok")')
    qline(25, 'Carbon calibration slope', '=Fits!$B$63', '0.0000',
          '=IF(Fits!$B$63="","",IF(ABS(Fits!$B$63-1)>0.05,"CHECK: stretch factor far from 1","ok"))')
    qline(26, 'Check standards measured in this run', '=Fits!$B$217', '0',
          '=IF(Fits!$B$217<1,"CHECK: nothing independent of the anchors tests this calibration; run a '
          'check standard (USGS66, or an in-house material with an assigned value) in every sequence","ok")')
    qline(27, 'Samples outside the anchor d15N range',
          f'=COUNTIF(Results!$K$5:$K${5+NSAMP-1},"*d15N outside anchor range*")', '0',
          '=IF($B$27=0,"ok","CHECK: these values are extrapolated beyond the anchors; add an anchor that '
          'brackets them")')
    qline(28, 'Samples outside the anchor d13C range',
          f'=COUNTIF(Results!$K$5:$K${5+NSAMP-1},"*d13C outside anchor range*")', '0',
          '=IF($B$28=0,"ok","CHECK: these values are extrapolated beyond the anchors; add an anchor that '
          'brackets them")')

    qsec(30, 'Is each correction supported by this run?')
    qc.cell(31, 1, 'A slope is applied on Auto only when p is below the significance level in Setup B15. '
                   'The FIRMS guide (section 6.4) asks that a correction be shown necessary before it is '
                   'applied, because over-correction adds as much error as it removes.').font = F_NOTE
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
    qc.cell(36, 1, 'With five or six standards the test has little power: a slope that fails it is unproven, '
                   'not disproven. Set the switch to On to apply it regardless. The outlier test below needs five '
                   'monitors in the fit and is not run with fewer (1.4.2).').font = F_NOTE
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
                   'between the two models. The check standard offsets below can.').font = F_NOTE

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
    qc.cell(54, 1, 'A blank that gives no peak at all has not been measured, only bounded by the detection '
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
                          'are an independent test. Offsets are in per mil.').font = F_NOTE

    # The rows below moved down by three in 1.4.2, with the standards table.
    Q = 58 + NSTD + 2
    qsec(Q, 'Uncertainty')
    qline(Q + 1, 'Sample repeatability in use, d15N (permil)', '=Fits!$B$136', '0.000', '=Fits!$B$137')
    qline(Q + 2, 'Sample repeatability in use, d13C (permil)', '=Fits!$B$156', '0.000', '=Fits!$B$157')
    qline(Q + 3, 'Coverage factor k', f'={S_K}', '0.0')
    qline(Q + 4, 'Typical u(d15N) of the unknowns (permil)',
          f'=IFERROR(AVERAGEIFS({W("CL")},{WD_},"Sample"),"")')
    qline(Q + 5, 'Typical u(d13C) of the unknowns (permil)',
          f'=IFERROR(AVERAGEIFS({W("CN")},{WD_},"Sample"),"")')
    qc.cell(Q + 6, 1, 'u combines the certificate uncertainty of each anchor, the scatter of its replicates, and '
                      'the sample repeatability, through the fitted line, as a Kragten spreadsheet would. It '
                      'leaves out the blank, linearity and drift corrections and the day-to-day reproducibility '
                      'a control chart would give, so it is a floor, not the whole budget.').font = F_NOTE
    qline(Q + 7, 'SD of the drift standards, d15N (permil)', f'=IFERROR(STDEV(Work!$BG${R0}:$BG${R0+NSAMP-1}),"")')
    qline(Q + 8, 'SD of the drift standards, d13C (permil)', f'=IFERROR(STDEV(Work!$BH${R0}:$BH${R0+NSAMP-1}),"")')
    qc.cell(Q + 9, 1, 'These understate the real precision, because the drift standards are the material the drift '
                      'correction was fitted to. The check standard offsets above are the honest test.').font = F_NOTE

    qsec(Q + 11, 'Flags')
    qline(Q + 12, 'Rows carrying a flag on Results', f'=COUNTIF(Results!$K$5:$K${5+NSAMP-1},"?*")', '0')
    qline(Q + 13, 'Minimum trusted Area 28', '=Fits!$B$81')
    qline(Q + 14, 'Maximum trusted Area 28', '=Fits!$B$97')
    qline(Q + 15, 'Minimum trusted Area 44', '=Fits!$B$82')
    qline(Q + 16, 'Maximum trusted Area 44', '=Fits!$B$104')
    qline(Q + 17, 'Samples outside the linearity range',
          f'=COUNTIF(Results!$K$5:$K${5+NSAMP-1},"*linearity range*")', '0',
          f'=IF($B${Q + 17}=0,"ok","CHECK: weigh samples to land inside the standards, or extend the size series")')

    qsec(96, 'Instrument settling')
    qc.cell(97, 1, 'A run that began before the mass-28 background reached its floor carries a settling '
                   'transient, not a steady drift, and a straight line in position is the wrong shape for it.').font = F_NOTE
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
          f'"CHECK: size and time are confounded, the linearity fit sits inside the unsettled window"))')
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
    qc.cell(109, 1, 'A high background right after a large sample points to carryover from that sample, '
                    'not to instrument drift.').font = F_NOTE

    qsec(110, 'Does each correction reduce the scatter of the drift standards?')
    STAGES = [('as measured', 'BO'), ('after blank', 'BP'), ('after linearity', 'BQ'),
              ('after drift', 'BR'), ('after calibration', 'BG')]
    for i, (nm, col) in enumerate(STAGES):
        qline(111 + i, nm, f'=IFERROR(STDEV({W(col)}),"")')
    qc.cell(116, 4, 'SD in per mil.').font = F_NOTE
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

    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.trendline import Trendline

    def feed(col):
        return Reference(wk, min_col=wk[col + '1'].column, min_row=R0, max_row=R0 + NSAMP - 1)

    def add_series(ch, xcol, ycol, name, colour, marker=True, line=False, trend=False, size=5, hollow=False):
        sr = Series(feed(ycol), feed(xcol))
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
        if log_x:
            ch.x_axis.scaling.logBase = 10
        qc.add_chart(ch, anchor)
        return ch

    BLUE, ORANGE, GREEN, YELLOW, GREY, RED = '1F4E79', 'C55A11', '1B7F5A', 'B8860B', '7F7F7F', 'B2182B'
    # Working gas, beside rows 10 to 19
    for anchor, el, gas in (('K10', 'N', 'N2'), ('T10', 'C', 'CO2')):
        ch = chart(anchor, f'{gas} reference pulses Isodat did not designate', 'Autosampler position',
                   ('d15N' if el == 'N' else 'd13C') + ' of the pulse (permil)')
        add_series(ch, CH['ch pos'], CH[f'ch pulse {el}'], 'pulses', GREY, trend=True)
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

    for s in wb.worksheets:
        s.sheet_properties.tabColor = '1F4E79' if s.title in ('Results', 'QC') else None
    wb.save(OUT)
    if recalc:
        bake_values(OUT)

    return Path(OUT)
