# isoharness

isoharness corrects Thermo EA-IRMS carbon and nitrogen runs on the web. Its numbers come from ea_irms, the Python engine behind the lab's EA-IRMS correction workbook, which this repository carries at version 1.4.2.

## Where it stands

This is the first of four phases, which puts the engine behind a web API, LibreOffice in the container image to compute the workbook's formulas, and the browser's preview port under a check against the engine in continuous integration. The web page comes in phase two, and sign-in, saving, and history in phase three. The service stores nothing; each upload is processed in a temporary folder that is removed before the response is sent.

## Layout

| Path | What it is |
| --- | --- |
| `ea_irms/` | The engine, 1.4.2, unchanged (parsing, the corrections, the workbook, and the Postgres store that phase three will use) |
| `isoharness/` | The web API (FastAPI) |
| `port/` | The preview port (the steps that depend on a user's choices, in JavaScript) and its test |
| `scripts/` | `port_cases.py` (engine outputs for the port test), `smoke.py` (checks a running service), `reference_check.py` (the engine check on reference runs kept outside this repository) |
| `db/` | The Postgres schema for accounts and saved runs (phase three) |
| `tests/` | The engine, store, isolation, and API tests, with the 2026-09-15 test run in `tests/data/` |
| `Dockerfile` | One image with the API, ea_irms, and LibreOffice |
| `.github/workflows/ci.yml` | The tests, the port check, and an image build with a smoke test, on every push |

## Running it

```
docker build -t isoharness .
docker run -p 8000:8000 isoharness
```

Then http://localhost:8000/api/docs lists the endpoints and lets you try each one. Without Docker, `pip install -r requirements.txt` and `uvicorn isoharness.app:app` run the same service, but its workbooks carry formulas only unless LibreOffice (`soffice`) is on the PATH.

## The API

| Endpoint | Returns |
| --- | --- |
| `GET /api/health` | `{"status": "ok"}` |
| `GET /api/version` | The isoharness and ea_irms versions, and whether LibreOffice is present |
| `GET /api/defaults` | The package default configuration, in the form the POST endpoints accept |
| `POST /api/process` | Every position's values and flags, the fitted steps, and the block the preview port recomputes from |
| `POST /api/workbook` | The correction workbook (.xlsx), with its formulas computed by LibreOffice unless `values=false` |
| `POST /api/results.csv` | The workbook's Results sheet as CSV, at full precision |

Each POST takes a multipart form with `file`, an export as .xls, .xlsx, .xlsm, .csv, .txt, or .tsv of at most 10 MB (`ISOHARNESS_MAX_UPLOAD_MB` changes the limit), and optionally `config`, a lab configuration as JSON in the form `/api/defaults` returns. `/api/workbook` also takes `run_name` and `run_date` for its Setup sheet. A file ea_irms cannot read is refused with a message that names the problem, such as the columns it lacks.

The service converts one workbook at a time, because LibreOffice conversions that run together share a profile; in a test on 2026-10-11, one of three conversions started together failed, so the image runs a single service process. Three requests sent together came back in 5.1, 6.8, and 8.5 s, one after another. The image starts LibreOffice once while it is built, which in a test on 2026-10-11 brought the first workbook from a fresh container down from 3.42 s to 2.47 s (one trial each).

## The preview port

`port/engine.js` recomputes blank, linearity, drift, calibration, elemental content, and U in the browser, so a page can redraw on every switch, while the API's values stay the ones saved and downloaded. It reads only the user's decisions (the switches, the drift model, exclusions, and the sample repeatability); the significance level, the coverage factor, the elemental gate, and the reference-pulse windows are fixed in it at the package defaults. A lab configuration that changes those needs them passed to the port, or the preview marked as approximate, which is phase two's to settle.

```
python scripts/port_cases.py tests/data/20260915_raw.csv --out build/port
node port/test_port.js build/port/port_input.json build/port/cases.json
```

## Checks

The tests run with `pip install -r requirements-dev.txt` and `python -m pytest`. The store and isolation tests need Postgres at `EA_TEST_SUPER_DSN` and `EA_TEST_APP_DSN` (see `.github/workflows/ci.yml`), with a UTF8 database, since a SQL_ASCII one returns text as bytes and fails seven of them; the workbook tests need LibreOffice.

`python scripts/smoke.py http://localhost:8000 tests/data/20260915_raw.csv` checks a running service, requiring the workbook it returns to have its formulas computed and to hold the values `/api/process` returned.

`python scripts/reference_check.py MANIFEST.json --url http://localhost:8000` is the engine check before a release, on reference runs that stay outside this repository (the manifest format is in the script). On each run the API's values must equal its workbook, they must equal ea_irms called directly, and the workbook must equal the one delivered from the same export, cell by cell.

Results on 2026-10-11, with the smoke test and the reference check run against the image built from this repository:

| Check | Result |
| --- | --- |
| `pytest` | 130 passed |
| Port, 2026-09-15 test run | 643 decision sets, 740,736 comparisons, 0 mismatches, worst relative difference 1.35e-13 |
| Port, one reference run | 643 decision sets, 830,756 comparisons, 0 mismatches, worst relative difference 1.31e-12 |
| Smoke test against the image | 422 values and 56 flag lists, 0 problems |
| Reference check against the image, nine runs | 25,527 values against the workbook, 33,995 against the engine, and 778,464 cell comparisons (formulas and computed values) against the workbooks delivered for eight of the runs; 0 mismatches |

## Data in this repository

The only real run here is the 2026-09-15 test run (standards, blanks, and one position named Random), in `tests/data/` as a CSV export and as the reference workbook the engine tests compare against. Every other run is lab data, checked with `scripts/reference_check.py` outside this repository.

## License

MIT, copyright 2026 Xingchen Wang; see `LICENSE`.
