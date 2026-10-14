"""Read a Thermo Isodat EA-IRMS export and identify the sample peaks.

Peak position in the export is not a reliable identifier. A position with no
sample peak leaves the usual index pointing at a reference-gas peak, and the
2026-09-15 run has five of those. Everything here keys on the ``Is Ref _`` flag
and on which amplitude column is populated.
"""
from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path

from .xlsxguard import MAX_READ_COLUMNS, MAX_READ_ROWS, UploadTooLarge, check_xlsx

REQUIRED = [
    "Row", "Identifier 1", "Peak Nr", "Amount", "Ampl  28", "BGD 28",
    "Area 28", "Is Ref _", "Ampl  44", "Area 44", "d 15N/14N", "d 13C/12C",
]


class ExportFormatError(ValueError):
    """The file is not a Thermo EA-IRMS export this parser recognises."""


@dataclass
class Position:
    """One autosampler position."""
    row: int
    identifier: str
    amount: float | None = None          # mg, as the export carries it
    area_n: float | None = None
    ampl_n: float | None = None
    bgd_n: float | None = None           # background under the sample N2 peak
    ref_bgd_n: float | None = None       # background on the N2 reference peak
    ref_area_n: float | None = None
    d15n: float | None = None
    area_c: float | None = None
    ampl_c: float | None = None
    d13c: float | None = None
    ref_area_c: float | None = None
    ref_pulse_d15n: float | None = None  # first N2 reference pulse Isodat did not designate
    ref_pulse_d13c: float | None = None  # first CO2 reference pulse after the designated one
    # 1.4.10: every reference pulse Isodat did not designate, in peak order (the full exports carry
    # one N2 pulse and two CO2 pulses at each position); the reference-gas checks read them all
    ref_pulses_d15n: list = field(default_factory=list)
    ref_pulses_d13c: list = field(default_factory=list)
    n_peaks: int = 0
    notes: list[str] = field(default_factory=list)
    # 1.4.2: amplitude and d13C of every flagged (designated) CO2 pulse in the
    # row, the yardstick for an unflagged pulse mistaken for the sample peak
    desig_c: list = field(default_factory=list)
    c_rejected: bool = False             # the CO2 peak was a reference pulse (set by process)

    @property
    def has_n(self) -> bool:
        return self.d15n is not None

    @property
    def has_c(self) -> bool:
        return self.d13c is not None


def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


EMPTY_RUN_STOP = 2000      # consecutive empty rows of an .xlsx sheet that end the reading (1.4.10)


def _too_many(path: Path, n: int):
    return ExportFormatError(f"{path.name} has more than {MAX_READ_ROWS} rows, more than an Isodat export of "
                             f"one run holds and the most this reads (it holds {n} or more)")


def _read_rows(path: str | Path):
    """Return (header, rows) from .xls, .xlsx or .csv without guessing.

    From 1.4.10 an .xlsx file is checked before it is opened (ea_irms.xlsxguard) and read as a stream,
    and every format stops at MAX_READ_ROWS rows of at most MAX_READ_COLUMNS columns."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".xls":
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        sheet = book.sheet_by_index(0)
        if sheet.nrows > MAX_READ_ROWS + 1:
            raise _too_many(path, sheet.nrows - 1)
        ncols = min(sheet.ncols, MAX_READ_COLUMNS)
        header = [sheet.cell_value(0, c) for c in range(ncols)] if sheet.nrows else []
        rows = [[sheet.cell_value(r, c) for c in range(min(ncols, sheet.row_len(r)))]
                for r in range(1, sheet.nrows)]
        book.release_resources()
        return header, rows
    if suffix in (".xlsx", ".xlsm"):
        import openpyxl
        try:
            check_xlsx(path, what=path.name)
        except UploadTooLarge as exc:
            raise ExportFormatError(str(exc)) from exc
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            data, empty = [], 0
            for row in ws.iter_rows(values_only=True, max_col=MAX_READ_COLUMNS):
                if data and all(v is None for v in row):
                    empty += 1                   # rows with nothing in them, which a sheet's stated
                    if empty >= EMPTY_RUN_STOP:  # size can add by the million, are passed over
                        break
                    continue
                empty = 0
                data.append(list(row))
                if len(data) > MAX_READ_ROWS + 1:
                    raise _too_many(path, len(data) - 1)
        finally:
            wb.close()
        if not data:
            raise ExportFormatError(f"{path.name} is empty")
        return list(data[0]), data[1:]
    if suffix in (".csv", ".txt", ".tsv"):
        import csv
        delim = "\t" if suffix == ".tsv" else ","
        data = []
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.reader(fh, delimiter=delim):
                data.append(row[:MAX_READ_COLUMNS])
                if len(data) > MAX_READ_ROWS + 1:
                    raise _too_many(path, len(data) - 1)
        if not data:
            raise ExportFormatError(f"{path.name} is empty")
        return data[0], data[1:]
    raise ExportFormatError(f"cannot read {path.suffix} exports")


def parse_export(path: str | Path) -> list[Position]:
    """Parse an export into one :class:`Position` per autosampler row."""
    header, rows = _read_rows(path)
    header = [str(h).strip() for h in header]
    index = {h: i for i, h in enumerate(header)}
    missing = [c for c in REQUIRED if c not in index]
    if missing:
        raise ExportFormatError(
            "export is missing the columns " + ", ".join(missing)
            + ". Check that the Isodat export template includes them."
        )

    def cell(record, name):
        i = index[name]
        return record[i] if i < len(record) else None

    grouped: "OrderedDict[int, list]" = OrderedDict()
    for record in rows:
        row = _num(cell(record, "Row"))
        if row is None:
            continue
        grouped.setdefault(int(row), []).append(record)

    out: list[Position] = []
    for row, records in grouped.items():
        peaks = [r for r in records if _num(cell(r, "Peak Nr")) is not None]
        mass28 = [r for r in peaks if _num(cell(r, "Ampl  28")) is not None]
        mass44 = [r for r in peaks if _num(cell(r, "Ampl  44")) is not None]
        is_ref = lambda r: _num(cell(r, "Is Ref _")) == 1.0   # noqa: E731

        pos = Position(
            row=row,
            identifier=str(cell(records[0], "Identifier 1") or "").strip(),
            amount=_num(cell(records[0], "Amount")) or None,
            n_peaks=len(peaks),
        )

        # Sample N2: the last mass-28 peak, valid only when it is not a reference.
        if mass28:
            last = mass28[-1]
            if not is_ref(last):
                pos.area_n = _num(cell(last, "Area 28"))
                pos.ampl_n = _num(cell(last, "Ampl  28"))
                pos.bgd_n = _num(cell(last, "BGD 28"))
                pos.d15n = _num(cell(last, "d 15N/14N"))
            else:
                pos.notes.append("no sample nitrogen peak")
            refs = [r for r in mass28 if is_ref(r)]
            if refs:
                pos.ref_bgd_n = _num(cell(refs[-1], "BGD 28"))
                pos.ref_area_n = _num(cell(refs[-1], "Area 28"))
            # Reference pulses other than the designated one: every unflagged
            # mass-28 peak before the last one (the last is the sample when it
            # is unflagged, and a flagged reference otherwise).
            pulses = [r for r in mass28[:-1] if not is_ref(r)]
            if pulses:
                pos.ref_pulse_d15n = _num(cell(pulses[0], "d 15N/14N"))
            pos.ref_pulses_d15n = [v for v in (_num(cell(r, "d 15N/14N")) for r in pulses) if v is not None]
        else:
            pos.notes.append("no mass-28 peaks")

        # Sample CO2: the mass-44 peak before the first reference one.
        if mass44:
            first_ref = next((i for i, r in enumerate(mass44) if is_ref(r)), None)
            sample = None
            if first_ref:                       # not None and not 0
                sample = mass44[first_ref - 1]
            elif first_ref is None:
                # A filtered export (Peak Nr 3 and 4 only) carries no reference
                # peaks to count from. The sample CO2 is then the first mass-44
                # peak; a position whose first peak is a flagged reference had no
                # sample gas, and whatever follows it is the reference train.
                sample = mass44[0]
                pos.notes.append("filtered export: no reference peaks to confirm the CO2 peak")
            if sample is not None:
                pos.area_c = _num(cell(sample, "Area 44"))
                pos.ampl_c = _num(cell(sample, "Ampl  44"))
                pos.d13c = _num(cell(sample, "d 13C/12C"))
            else:
                pos.notes.append("no sample carbon peak")
            refs = [r for r in mass44 if is_ref(r)]
            if refs:
                pos.ref_area_c = _num(cell(refs[0], "Area 44"))
                pos.desig_c = [(_num(cell(r, "Ampl  44")), _num(cell(r, "d 13C/12C"))) for r in refs]
            if first_ref is not None:
                pulses = [r for r in mass44[first_ref + 1:] if not is_ref(r)]
                if pulses:
                    pos.ref_pulse_d13c = _num(cell(pulses[0], "d 13C/12C"))
                pos.ref_pulses_d13c = [v for v in (_num(cell(r, "d 13C/12C")) for r in pulses) if v is not None]
        out.append(pos)
    return out
