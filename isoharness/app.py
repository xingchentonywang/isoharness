"""The isoharness web API: EA-IRMS exports in, corrected values and files out.

Nothing is stored. Each request writes the upload to a temporary folder, runs
ea_irms on it, and removes the folder before the response leaves. The engine is
the only source of the numbers; the browser's preview port is checked against
it in continuous integration.

The site itself (``web/``) is served from the same address: the page at ``/``,
the preview port at ``/engine.js`` (``port/engine.js``, the file CI checks).
Uploads are limited per client address in any hour: workbooks, which LibreOffice
computes, to ``ISOHARNESS_UPLOADS_PER_HOUR`` (default 60), and the requests the engine
answers alone (processing a run, the results tables, reading a configuration) to
``ISOHARNESS_READS_PER_HOUR`` (default 600, from isoharness 0.2.0, since a batch sends
one request per run); 0 turns either off. Several runs can be combined into one results
table (isoharness 0.2.0), at most ``ISOHARNESS_MAX_BATCH_RUNS`` (default 20) in a request.
A sample sheet and each run's information (isoharness 0.3.0) go with the request that needs
them, as the export does, and are read for that request only.

    uvicorn isoharness.app:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
from collections import deque
from contextlib import ExitStack, contextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles

from ea_irms.labconfig import LabConfigError, config_to_xlsx, read_lab_config
from ea_irms.metadata import (MAX_ROWS, SHEET_SUFFIXES, SampleSheetError, clean_run_info, match_positions,
                              match_summary, output_columns, output_header, read_sample_sheet, sample_names,
                              template_xlsx)
from ea_irms.parse import ExportFormatError, parse_export
from ea_irms.pipeline import Config, config_from_dict, config_to_dict, process
from ea_irms.workbook import VERSION, bake_values, build_workbook

from . import __version__
from .batch import BatchRun, combined_csv, combined_xlsx, label_runs
from .serialize import clean, port_input, position_record, results_csv, sample_records

log = logging.getLogger("isoharness")

SUFFIXES = (".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".tsv")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CHUNK = 1 << 16

# One LibreOffice at a time: concurrent conversions share a user profile and fail.
_libreoffice = threading.Lock()

ROOT = Path(__file__).resolve().parents[1]
WEB = ROOT / "web"
PORT_ENGINE = ROOT / "port" / "engine.js"
# The 2026-09-15 test run, offered on the page as an example: in the image it sits in
# warmup/, where the Dockerfile copies it; in the repository, in tests/data/.
EXAMPLE = next((p for p in (ROOT / "warmup" / "20260915_raw.csv",
                            ROOT / "tests" / "data" / "20260915_raw.csv") if p.is_file()), None)
LIMITED = ("/api/process", "/api/workbook", "/api/results.csv", "/api/config/read",
           "/api/batch/results.xlsx", "/api/batch/results.csv",
           "/api/samplesheet/read", "/api/samplesheet/template.xlsx")
BATCH = ("/api/batch/results.xlsx", "/api/batch/results.csv")
# The workbook is the costly request, since LibreOffice computes it: for the test run a workbook
# request to isoharness.org took 8.96 s on 2026-10-13. The engine answers the others alone; for the
# same run /api/process reported 0.011 s of processing, measured locally that day. So each kind
# has a limit of its own.
HEAVY = ("/api/workbook",)
WINDOW_S = 3600.0
# The page's own scripts and styles, Google Fonts, and calls back to this address.
CSP = ("default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' "
       "https://fonts.googleapis.com; font-src https://fonts.gstatic.com; "
       "img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; "
       "base-uri 'none'; frame-ancestors 'none'; form-action 'self'")


def max_upload_bytes() -> int:
    return int(float(os.environ.get("ISOHARNESS_MAX_UPLOAD_MB", "10")) * 1024 * 1024)


def uploads_per_hour() -> int:
    """Workbooks per address in an hour."""
    return int(os.environ.get("ISOHARNESS_UPLOADS_PER_HOUR", "60"))


def reads_per_hour() -> int:
    """Requests the engine answers alone, per address in an hour (isoharness 0.2.0)."""
    return int(os.environ.get("ISOHARNESS_READS_PER_HOUR", "600"))


def max_batch_runs() -> int:
    return int(os.environ.get("ISOHARNESS_MAX_BATCH_RUNS", "20"))


class UploadLimiter:
    """At most N limited requests per client address in any hour, counted in memory.

    The count lives in one process, so it holds per container; the cap on
    instances set at deployment bounds the rest. The address is the one uvicorn
    reports, which behind Cloud Run's front end comes from X-Forwarded-For.
    """

    def __init__(self):
        self._hits: dict[str, deque] = {}
        self._lock = threading.Lock()

    def reset(self):
        with self._lock:
            self._hits.clear()

    def wait_seconds(self, key: str, limit: int, now: float | None = None) -> int | None:
        """Count one request; None if it may proceed, else the seconds until it could."""
        now = time.monotonic() if now is None else now
        with self._lock:
            hits = self._hits.setdefault(key, deque())
            while hits and now - hits[0] >= WINDOW_S:
                hits.popleft()
            if len(hits) >= limit:
                return max(1, int(WINDOW_S - (now - hits[0])) + 1)
            hits.append(now)
            if len(self._hits) > 10000:          # forget addresses with nothing in the window
                for k in [k for k, v in self._hits.items() if not v or now - v[-1] >= WINDOW_S]:
                    del self._hits[k]
            return None


limiter = UploadLimiter()


app = FastAPI(
    title="isoharness API",
    version=__version__,
    description="Stable isotope data reduction. These endpoints serve the EA-IRMS page, with ea_irms. "
                "Uploaded files are deleted as soon as they are processed, and nothing is stored.",
    docs_url="/api/docs",
    redoc_url=None,
    openapi_url="/api/openapi.json",
)


@app.middleware("http")
async def limit_uploads(request: Request, call_next):
    """Refuse a client address its next upload once it has made the hour's allowance: workbooks
    and the other uploads are counted apart, each against its own limit."""
    if request.method == "POST" and request.url.path in LIMITED:
        heavy = request.url.path in HEAVY
        limit = uploads_per_hour() if heavy else reads_per_hour()
        if limit > 0:
            key = (request.client.host if request.client else "unknown") + (" workbook" if heavy else "")
            wait = limiter.wait_seconds(key, limit)
            if wait is not None:
                minutes = max(1, round(wait / 60))
                what = "workbook requests" if heavy else "uploads"
                return JSONResponse(
                    {"detail": f"this address has sent {limit} {what} in the last hour, the most this "
                               f"service takes; try again in about {minutes} minute"
                               f"{'' if minutes == 1 else 's'}"},
                    status_code=429, headers={"Retry-After": str(wait)})
    return await call_next(request)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    if not request.url.path.startswith("/api/"):
        # The API's own docs page loads its scripts from a CDN, so it keeps no policy.
        response.headers.setdefault("Content-Security-Policy", CSP)
    return response


@app.middleware("http")
async def refuse_large_bodies(request: Request, call_next):
    """Turn away a declared body far over the limit before reading it. The form's
    own fields and boundaries get 1 MB of room; the file itself is checked again
    as it is written."""
    declared = request.headers.get("content-length", "")
    if declared.isdigit() and int(declared) > max_upload_bytes() + 1024 * 1024:
        limit = max_upload_bytes() // (1024 * 1024)
        what = ("the exports together are" if request.url.path in BATCH else "the upload is")
        return JSONResponse({"detail": f"{what} larger than the {limit} MB this service accepts"
                                       f"{' in one request' if request.url.path in BATCH else ''}"},
                            status_code=413)
    return await call_next(request)


def _stem(name: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._")
    return stem[:100] or "export"


@contextmanager
def _upload(file: UploadFile):
    """Write the upload to a private temporary folder; remove it afterwards."""
    name = Path(file.filename or "").name
    suffix = Path(name).suffix.lower()
    if suffix not in SUFFIXES:
        raise HTTPException(415, f"{name or 'the file'} is not an export ea_irms reads; "
                                 f"use one of {', '.join(SUFFIXES)}")
    limit = max_upload_bytes()
    digest = hashlib.sha256()
    size = 0
    with tempfile.TemporaryDirectory(prefix="isoharness-") as tmp:
        path = Path(tmp) / f"{_stem(name)}{suffix}"
        with open(path, "wb") as out:
            while True:
                block = file.file.read(CHUNK)
                if not block:
                    break
                size += len(block)
                if size > limit:
                    raise HTTPException(413, f"{name} is larger than the {limit // (1024 * 1024)} MB "
                                             "this service accepts")
                digest.update(block)
                out.write(block)
        if size == 0:
            raise HTTPException(422, f"{name} is empty")
        yield path, {"name": name, "bytes": size, "sha256": digest.hexdigest()}


def _config(text: str | None) -> Config:
    if text is None or not text.strip():
        return Config()
    try:
        data = json.loads(text)
        if not isinstance(data, dict):
            raise ValueError("the configuration must be a JSON object")
        return config_from_dict(data)
    except (ValueError, TypeError, KeyError) as exc:
        raise HTTPException(422, f"the configuration could not be read: {exc}") from exc


def _parse(path: Path, name: str):
    try:
        positions = parse_export(path)
        if not positions:
            raise HTTPException(422, f"{name} holds the export's columns but no positions")
        return positions
    except HTTPException:
        raise
    except ExportFormatError as exc:
        raise HTTPException(422, f"cannot read {name}: {exc}") from exc
    except Exception as exc:  # a corrupt or mislabeled file is the caller's input, not our fault
        log.warning("parse failed for %s: %r", name, exc)
        raise HTTPException(422, f"cannot read {name}: the file could not be opened as "
                                 f"{path.suffix} ({type(exc).__name__})") from exc


def _process(positions, config: Config, name: str, samples=None):
    try:
        return process(positions, config, samples=samples)
    except (ValueError, ZeroDivisionError, IndexError, TypeError) as exc:
        # TypeError: a configuration field of the wrong kind, such as text where a number belongs
        log.warning("process failed for %s: %r", name, exc)
        raise HTTPException(422, f"the run in {name} could not be corrected: {exc}") from exc


def _sheet(upload: UploadFile | None, id_column: str | None = None):
    """The sample sheet sent with a request (isoharness 0.3.0), read for this request only; None without one."""
    if upload is None or not (upload.filename or "").strip():
        return None
    name = Path(upload.filename or "").name
    if Path(name).suffix.lower() not in SHEET_SUFFIXES:
        raise HTTPException(415, f"{name or 'the file'} is not a sample sheet ea_irms reads; use one of "
                                 f"{', '.join(SHEET_SUFFIXES)}")
    limit = max_upload_bytes()
    data = upload.file.read(limit + 1)
    if len(data) > limit:
        raise HTTPException(413, f"{name} is larger than the {limit // (1024 * 1024)} MB this service accepts")
    try:
        return read_sample_sheet(data, name, id_column)
    except SampleSheetError as exc:
        raise HTTPException(422, f"the sample sheet could not be read: {exc}") from exc
    except Exception as exc:  # a corrupt file is the caller's input
        log.warning("sample sheet failed for %s: %r", name, exc)
        raise HTTPException(422, f"the sample sheet {name} could not be read ({type(exc).__name__})") from exc


def _info(value, where: str = "run_info") -> dict:
    """Run information, a JSON object or its text, as ea_irms.metadata keeps it."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return {}
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except ValueError as exc:
            raise HTTPException(422, f"{where} could not be read as JSON: {exc}") from exc
    try:
        return clean_run_info(value)
    except ValueError as exc:
        raise HTTPException(422, f"{where}: {exc}") from exc


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/api/version")
def version():
    return {"isoharness": __version__, "ea_irms": VERSION,
            "libreoffice": bool(shutil.which("soffice") or shutil.which("libreoffice"))}


@app.get("/api/defaults")
def defaults():
    """The package default configuration, in the form /api/process accepts."""
    return {"ea_irms": VERSION, "config": config_to_dict(Config())}


@app.post("/api/process")
def process_run(file: UploadFile = File(...), config: str | None = Form(None),
                run_date: str | None = Form(None), sample_sheet: UploadFile | None = File(None),
                id_column: str | None = Form(None), run_info: str | None = Form(None)):
    """Correct one run. Returns every position's values and flags, the fitted steps,
    and the block the browser's preview port recomputes from. With a sample sheet (0.3.0)
    each sample is matched to it and its replicates are compared for the rerun marks."""
    t0 = time.perf_counter()
    with _upload(file) as (path, info):
        cfg = _config(config)
        sheet = _sheet(sample_sheet, id_column)
        rinfo = _info(run_info)
        positions = _parse(path, info["name"])
        positions, fits = _process(positions, cfg, info["name"], sheet)
        body = {
            "isoharness": __version__,
            "ea_irms": VERSION,
            "file": info,
            "config": config_to_dict(cfg),
            "positions": [position_record(p) for p in positions],
            "fits": clean(fits),
            "port": port_input(positions, fits, cfg, info["name"], info["bytes"], run_date),
            "metadata": clean({"run_info": rinfo, "sheet": sheet.describe() if sheet else None,
                               "match": fits["metadata"]["match"], "samples": sample_records(positions),
                               "columns": output_header(sheet, output_columns(sheet, [positions]))}),
        }
    body["seconds"] = round(time.perf_counter() - t0, 4)
    return body


@app.post("/api/workbook")
def workbook(file: UploadFile = File(...), config: str | None = Form(None),
             run_name: str | None = Form(None), run_date: str | None = Form(None),
             values: bool = Form(True), sample_sheet: UploadFile | None = File(None),
             id_column: str | None = Form(None), run_info: str | None = Form(None)):
    """The correction workbook for one run. With ``values`` (the default) LibreOffice
    computes every formula, so the numbers show in any viewer. A sample sheet and the run
    information (0.3.0) fill its Metadata sheet and the columns after Results."""
    with _upload(file) as (path, info):
        cfg = _config(config)
        sheet = _sheet(sample_sheet, id_column)
        rinfo = _info(run_info)
        _parse(path, info["name"])
        out = path.with_name(f"{path.stem}_ea_irms_v{VERSION}.xlsx")
        try:
            build_workbook(path, out, cfg, run_name=run_name or Path(info["name"]).stem,
                           run_date=run_date or None, samples=sheet, run_info=rinfo, file_name=info["name"])
        except ValueError as exc:
            raise HTTPException(422, f"the workbook could not be built: {exc}") from exc
        computed = False
        if values:
            with _libreoffice:
                computed = bake_values(out)
        data = out.read_bytes()
    return Response(data, media_type=XLSX, headers={
        "Content-Disposition": f'attachment; filename="{out.name}"',
        "X-Isoharness-Values": "computed" if computed else "formulas-only",
    })


@app.post("/api/results.csv")
def results(file: UploadFile = File(...), config: str | None = Form(None),
            sample_sheet: UploadFile | None = File(None), id_column: str | None = Form(None)):
    """The workbook's Results sheet as CSV, from the engine at full precision; with a sample
    sheet (0.3.0), its columns after them."""
    with _upload(file) as (path, info):
        cfg = _config(config)
        sheet = _sheet(sample_sheet, id_column)
        positions, _ = _process(_parse(path, info["name"]), cfg, info["name"], sheet)
        text = results_csv(positions, cfg, sheet)
        name = f"{path.stem}_results_ea_irms_v{VERSION}.csv"
    return Response(text, media_type="text/csv", headers={
        "Content-Disposition": f'attachment; filename="{name}"'})


def _json_list(text: str | None, field: str, n: int) -> list:
    """A form field holding a JSON list with one entry per file; empty means none for every file."""
    if text is None or not text.strip():
        return [None] * n
    try:
        data = json.loads(text)
    except ValueError as exc:
        raise HTTPException(422, f"{field} could not be read as JSON: {exc}") from exc
    if not isinstance(data, list) or len(data) != n:
        raise HTTPException(422, f"{field} must be a JSON list with one entry for each of the {n} files")
    return data


def _batch(files: list[UploadFile], configs: str | None, run_dates: str | None, sheet=None,
           run_infos: str | None = None) -> list[BatchRun]:
    """Correct each export on its own, with its own configuration, as /api/process does."""
    if not files:
        raise HTTPException(422, "send the exports in the files field")
    if len(files) > max_batch_runs():
        raise HTTPException(422, f"{len(files)} exports were sent; at most {max_batch_runs()} runs are combined "
                                 "in one request")
    cfgs = None
    if configs is not None and configs.strip():
        try:
            one = json.loads(configs)
        except ValueError as exc:
            raise HTTPException(422, f"configs could not be read as JSON: {exc}") from exc
        if isinstance(one, dict):
            cfgs = [one] * len(files)      # one configuration for every run
    if cfgs is None:
        cfgs = _json_list(configs, "configs", len(files))
    dates = _json_list(run_dates, "run_dates", len(files))
    infos = None
    if run_infos is not None and run_infos.strip():
        try:
            one = json.loads(run_infos)
        except ValueError as exc:
            raise HTTPException(422, f"run_infos could not be read as JSON: {exc}") from exc
        if isinstance(one, dict):
            infos = [one] * len(files)      # one run information for every run
    if infos is None:
        infos = _json_list(run_infos, "run_infos", len(files))
    runs, total, limit = [], 0, max_upload_bytes()
    with ExitStack() as stack:
        for f, data, date, rinfo in zip(files, cfgs, dates, infos):
            path, info = stack.enter_context(_upload(f))
            total += info["bytes"]
            if total > limit:
                raise HTTPException(413, f"the exports together are larger than the {limit // (1024 * 1024)} MB "
                                         "this service accepts in one request")
            if data is not None and not isinstance(data, dict):
                raise HTTPException(422, f"the configuration for {info['name']} is not a JSON object")
            try:
                cfg = _config(json.dumps(data) if data is not None else None)
            except HTTPException as exc:
                raise HTTPException(exc.status_code, f"{info['name']}: {exc.detail}") from exc
            run_info = _info(rinfo, f"the run information for {info['name']}")
            positions, fits = _process(_parse(path, info["name"]), cfg, info["name"], sheet)
            when = str(date).strip()[:40] if isinstance(date, (str, int, float)) and str(date).strip() else None
            runs.append(BatchRun(file=info["name"], positions=positions, fits=fits, config=cfg, run_date=when,
                                 run_info=run_info))
    return label_runs(runs)


@app.post("/api/batch/results.xlsx")
def batch_results_xlsx(files: list[UploadFile] = File(...), configs: str | None = Form(None),
                       run_dates: str | None = Form(None), config_name: str | None = Form(None),
                       sample_sheet: UploadFile | None = File(None), id_column: str | None = Form(None),
                       run_infos: str | None = Form(None)):
    """Several runs' results in one Excel file (isoharness 0.2.0): Results, every position of every run
    with a Run column first; Runs, one row per run; Repeats, the samples measured in more than one run.
    ``configs`` is a JSON list with one configuration per file, or one object for all of them;
    ``run_dates`` a JSON list of dates or nulls. Each run is corrected on its own. One sample sheet
    goes with all the runs, and ``run_infos`` is a JSON list with each run's information, or one
    object for all of them (0.3.0)."""
    sheet = _sheet(sample_sheet, id_column)
    runs = _batch(files, configs, run_dates, sheet, run_infos)
    name = f"combined_{len(runs)}_run{'s' if len(runs) != 1 else ''}_results_ea_irms_v{VERSION}.xlsx"
    data = combined_xlsx(runs, (config_name or "").strip()[:200] or None, sheet)
    return Response(data, media_type=XLSX, headers={"Content-Disposition": f'attachment; filename="{name}"'})


@app.post("/api/batch/results.csv")
def batch_results_csv(files: list[UploadFile] = File(...), configs: str | None = Form(None),
                      run_dates: str | None = Form(None), sample_sheet: UploadFile | None = File(None),
                      id_column: str | None = Form(None), run_infos: str | None = Form(None)):
    """The combined Results sheet as CSV: Run, then the columns of /api/results.csv, at full precision."""
    sheet = _sheet(sample_sheet, id_column)
    runs = _batch(files, configs, run_dates, sheet, run_infos)
    name = f"combined_{len(runs)}_run{'s' if len(runs) != 1 else ''}_results_ea_irms_v{VERSION}.csv"
    return Response(combined_csv(runs, sheet), media_type="text/csv",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


class _Sample:
    """A sample the page lists, for matching without its export."""
    role = "Sample"

    def __init__(self, row, identifier):
        self.row, self.identifier = row, identifier


def _page_runs(text: str | None) -> list[dict]:
    """The runs the page sends with a sample sheet: [{"name": ..., "samples": [[row, identifier], ...]}]."""
    if text is None or not text.strip():
        return []
    try:
        data = json.loads(text)
    except ValueError as exc:
        raise HTTPException(422, f"runs could not be read as JSON: {exc}") from exc
    if not isinstance(data, list) or len(data) > max_batch_runs():
        raise HTTPException(422, f"runs must be a JSON list of at most {max_batch_runs()} runs")
    out = []
    for k, r in enumerate(data, start=1):
        if not isinstance(r, dict) or not isinstance(r.get("samples", []), list):
            raise HTTPException(422, f"run {k} of runs must be an object with a samples list")
        samples = []
        for item in r.get("samples", []):
            if (not isinstance(item, (list, tuple)) or len(item) != 2 or not isinstance(item[0], int)
                    or isinstance(item[0], bool) or not isinstance(item[1], (str, int, float))):
                raise HTTPException(422, f"each sample of run {k} must be [position, identifier]")
            samples.append(_Sample(item[0], str(item[1])[:200]))
        if len(samples) > 1000:
            raise HTTPException(422, f"run {k} lists more than 1000 samples")
        out.append({"name": str(r.get("name") or f"run {k}")[:200], "samples": samples})
    return out


@app.post("/api/samplesheet/read")
def samplesheet_read(file: UploadFile = File(...), id_column: str | None = Form(None),
                     runs: str | None = Form(None)):
    """Read a sample sheet (isoharness 0.3.0) and, for the runs the page lists, match their samples
    to it: how each matched, which did not, and the replicate groups. Nothing is kept."""
    sheet = _sheet(file, id_column)
    if sheet is None:
        raise HTTPException(422, "send the sample sheet in the file field")
    listed = _page_runs(runs)
    out = []
    for r in listed:
        groups = match_positions(r["samples"], sheet)
        out.append({"name": r["name"], "samples": sample_records(r["samples"]),
                    "match": match_summary(r["samples"], sheet),
                    "groups": [{k: g[k] for k in ("group", "name", "rows", "by")} for g in groups]})
    return clean({"ea_irms": VERSION, "sheet": sheet.describe(), "runs": out,
                  "columns": output_header(sheet, output_columns(sheet, [r["samples"] for r in listed]))})


@app.post("/api/samplesheet/template.xlsx")
async def samplesheet_template(request: Request):
    """A sample sheet to fill in (isoharness 0.3.0), with one row per name sent:
    ``{"names": ["CI 3-2-H", ...]}``. The page sends the samples of the runs it holds."""
    try:
        body = await request.json()
    except ValueError as exc:
        raise HTTPException(422, "send a JSON object with the sample names in its names field") from exc
    names = body.get("names", []) if isinstance(body, dict) else None
    if not isinstance(names, list) or not all(isinstance(n, (str, int, float)) and not isinstance(n, bool)
                                              for n in names):
        raise HTTPException(422, "send a JSON object with the sample names in its names field")
    if len(names) > MAX_ROWS:
        raise HTTPException(422, f"at most {MAX_ROWS} names fit in one sample sheet")
    data = template_xlsx(sample_names([[str(n)[:200] for n in names]]))
    return Response(data, media_type=XLSX, headers={
        "Content-Disposition": f'attachment; filename="sample_sheet_ea_irms_v{VERSION}.xlsx"'})


def _xlsx_name(name, suffix: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name or "")).strip("._")[:100] or "lab"
    return f"{stem}{suffix}"


@app.get("/api/config-template")
def config_template():
    """The lab configuration as an Excel template, filled with the package defaults (ea_irms 1.4.6)."""
    name = f"lab_configuration_ea_irms_v{VERSION}.xlsx"
    return Response(config_to_xlsx(Config()), media_type=XLSX, headers={
        "Content-Disposition": f'attachment; filename="{name}"', "Cache-Control": "no-cache"})


@app.post("/api/config.xlsx")
async def config_xlsx(request: Request):
    """A configuration sent as JSON, ``{"config": {...}, "name": "run"}``, written as an Excel
    lab configuration; missing fields take the package defaults, as /api/process does."""
    try:
        body = await request.json()
    except ValueError as exc:
        raise HTTPException(422, "send a JSON object with the configuration in its config field") from exc
    if not isinstance(body, dict) or not isinstance(body.get("config", {}), dict):
        raise HTTPException(422, "send a JSON object with the configuration in its config field")
    cfg = _config(json.dumps(body.get("config") or {}))
    name = _xlsx_name(body.get("name"), "_config.xlsx")
    return Response(config_to_xlsx(cfg), media_type=XLSX, headers={
        "Content-Disposition": f'attachment; filename="{name}"'})


@app.post("/api/config/read")
def config_read(file: UploadFile = File(...)):
    """Read a lab configuration, the Excel template (.xlsx) or JSON, into the form /api/process
    accepts. Returns the configuration and notes on anything filled from the defaults or left out."""
    name = Path(file.filename or "").name
    suffix = Path(name).suffix.lower()
    if suffix not in (".xlsx", ".json"):
        raise HTTPException(415, f"{name or 'the file'} is not a lab configuration; use the Excel template "
                                 "(.xlsx) or a .json file")
    limit = max_upload_bytes()
    data = file.file.read(limit + 1)
    if len(data) > limit:
        raise HTTPException(413, f"{name} is larger than the {limit // (1024 * 1024)} MB this service accepts")
    if not data:
        raise HTTPException(422, f"{name} is empty")
    if suffix == ".json":
        cfg, notes = _config(data.decode("utf-8", "replace")), []
    else:
        try:
            cfg, notes = read_lab_config(data)
        except LabConfigError as exc:
            raise HTTPException(422, f"{name} could not be read: {exc}") from exc
    return {"name": name, "ea_irms": VERSION, "config": config_to_dict(cfg), "notes": notes}


@app.get("/api/example")
def example():
    """The 2026-09-15 test run, the export the page's example uses; the repository holds it too."""
    if EXAMPLE is None:
        raise HTTPException(404, "the example run is not in this deployment")
    return FileResponse(EXAMPLE, media_type="text/csv", filename="20260915_raw.csv",
                        headers={"Cache-Control": "no-cache"})


@app.get("/engine.js", include_in_schema=False)
def preview_port():
    """The preview port the page runs, the same file continuous integration checks."""
    return FileResponse(PORT_ENGINE, media_type="text/javascript",
                        headers={"Cache-Control": "no-cache"})


class _Site(StaticFiles):
    """The page's files, revalidated on each load so a new release shows at once."""

    async def get_response(self, path, scope):
        response = await super().get_response(path, scope)
        response.headers.setdefault("Cache-Control", "no-cache")
        return response


if (WEB / "index.html").is_file():
    # The home page, web/index.html, lists the instruments. Each instrument's page has a path
    # without a trailing slash, so that its relative links (scripts, api/...) reach the root.
    @app.api_route("/ea-irms", methods=["GET", "HEAD"], include_in_schema=False)
    def ea_irms_page():
        return FileResponse(WEB / "ea-irms.html", media_type="text/html",
                            headers={"Cache-Control": "no-cache"})

    @app.get("/ea-irms/", include_in_schema=False)
    def ea_irms_slash():
        return RedirectResponse("/ea-irms", status_code=308)

    app.mount("/", _Site(directory=WEB, html=True), name="site")
else:
    @app.get("/")
    def root():
        return {"service": "isoharness", "version": __version__, "docs": "/api/docs"}
