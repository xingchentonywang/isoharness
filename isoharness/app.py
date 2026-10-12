"""The isoharness web API: one EA-IRMS export in, corrected values and files out.

Nothing is stored. Each request writes the upload to a temporary folder, runs
ea_irms on it, and removes the folder before the response leaves. The engine is
the only source of the numbers; the browser's preview port is checked against
it in continuous integration.

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
from contextlib import contextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse, Response

from ea_irms.parse import ExportFormatError, parse_export
from ea_irms.pipeline import Config, config_from_dict, config_to_dict, process
from ea_irms.workbook import VERSION, bake_values, build_workbook

from . import __version__
from .serialize import clean, port_input, position_record, results_csv

log = logging.getLogger("isoharness")

SUFFIXES = (".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".tsv")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CHUNK = 1 << 16

# One LibreOffice at a time: concurrent conversions share a user profile and fail.
_libreoffice = threading.Lock()


def max_upload_bytes() -> int:
    return int(float(os.environ.get("ISOHARNESS_MAX_UPLOAD_MB", "10")) * 1024 * 1024)


app = FastAPI(
    title="isoharness API",
    version=__version__,
    description="EA-IRMS corrections with ea_irms. Uploads are processed and discarded; "
                "nothing is stored.",
    docs_url="/api/docs",
    redoc_url=None,
    openapi_url="/api/openapi.json",
)


@app.middleware("http")
async def refuse_large_bodies(request: Request, call_next):
    """Turn away a declared body far over the limit before reading it. The form's
    own fields and boundaries get 1 MB of room; the file itself is checked again
    as it is written."""
    declared = request.headers.get("content-length", "")
    if declared.isdigit() and int(declared) > max_upload_bytes() + 1024 * 1024:
        limit = max_upload_bytes() // (1024 * 1024)
        return JSONResponse({"detail": f"the upload is larger than the {limit} MB this service accepts"},
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


def _process(positions, config: Config, name: str):
    try:
        return process(positions, config)
    except (ValueError, ZeroDivisionError, IndexError) as exc:
        log.warning("process failed for %s: %r", name, exc)
        raise HTTPException(422, f"the run in {name} could not be corrected: {exc}") from exc


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
                run_date: str | None = Form(None)):
    """Correct one run. Returns every position's values and flags, the fitted steps,
    and the block the browser's preview port recomputes from."""
    t0 = time.perf_counter()
    with _upload(file) as (path, info):
        cfg = _config(config)
        positions = _parse(path, info["name"])
        positions, fits = _process(positions, cfg, info["name"])
        body = {
            "isoharness": __version__,
            "ea_irms": VERSION,
            "file": info,
            "config": config_to_dict(cfg),
            "positions": [position_record(p) for p in positions],
            "fits": clean(fits),
            "port": port_input(positions, fits, cfg, info["name"], info["bytes"], run_date),
        }
    body["seconds"] = round(time.perf_counter() - t0, 4)
    return body


@app.post("/api/workbook")
def workbook(file: UploadFile = File(...), config: str | None = Form(None),
             run_name: str | None = Form(None), run_date: str | None = Form(None),
             values: bool = Form(True)):
    """The correction workbook for one run. With ``values`` (the default) LibreOffice
    computes every formula, so the numbers show in any viewer."""
    with _upload(file) as (path, info):
        cfg = _config(config)
        _parse(path, info["name"])
        out = path.with_name(f"{path.stem}_ea_irms_v{VERSION}.xlsx")
        try:
            build_workbook(path, out, cfg, run_name=run_name or Path(info["name"]).stem,
                           run_date=run_date or None)
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
def results(file: UploadFile = File(...), config: str | None = Form(None)):
    """The workbook's Results sheet as CSV, from the engine at full precision."""
    with _upload(file) as (path, info):
        cfg = _config(config)
        positions, _ = _process(_parse(path, info["name"]), cfg, info["name"])
        text = results_csv(positions, cfg)
        name = f"{path.stem}_results_ea_irms_v{VERSION}.csv"
    return Response(text, media_type="text/csv", headers={
        "Content-Disposition": f'attachment; filename="{name}"'})


@app.get("/")
def root():
    return {"service": "isoharness", "version": __version__, "docs": "/api/docs"}
