# isoharness

isoharness is a website for stable isotope data reduction, one instrument at a time. EA-IRMS comes first: it corrects Thermo EA-IRMS carbon and nitrogen runs, with numbers from ea_irms, the Python engine behind the lab's EA-IRMS correction workbook, which this repository carries at version 1.4.3. GasBench N₂O (nitrate by the denitrifier method, and organic nitrogen after persulfate oxidation) and GasBench carbonate are planned.

## Where it stands

Phase one put the engine behind a web API, LibreOffice in the container image to compute the workbook's formulas, and the browser's preview port under a check against the engine in continuous integration. Phase two adds the web pages: a home page at `/` that lists the instruments, and the EA-IRMS page at `/ea-irms`, where you upload an export, check the instrument's reference gas and background, go through the eight correction steps, and download the workbook, the results table, and the configuration. Anyone can use the page without signing in. Sign-in, saving, and history come in phase three. The service stores nothing; each upload is processed in a temporary folder that is removed before the response is sent. It runs at [isoharness.org](https://isoharness.org), on Google Cloud Run behind Firebase Hosting, and [DEPLOY.md](DEPLOY.md) has the steps.

## Layout

| Path | What it is |
| --- | --- |
| `ea_irms/` | The engine, 1.4.3 (parsing, the corrections, the workbook, and the Postgres store that phase three will use); 1.4.3 adds the uncertainty of %N, %C and C:N and puts each value's U beside it on the Results sheet |
| `isoharness/` | The web API (FastAPI) |
| `web/` | The pages the API serves: `index.html`, the home page at `/`, which lists the instruments; `ea-irms.html`, the EA-IRMS page at `/ea-irms` (the approved prototype's markup, as a template); `dc.js` (draws the template), `logic.js` (the views), and Preact 10.29.8 in `vendor/preact/` with its MIT license |
| `port/` | The preview port (the steps that depend on a user's choices, in JavaScript) and its test |
| `scripts/` | `port_cases.py` (engine outputs for the port test), `smoke.py` (checks a running service), `e2e_web.py` (the page in Chromium against a service it starts), `reference_check.py` (the engine check on reference runs kept outside this repository) |
| `db/` | The Postgres schema for accounts and saved runs (phase three) |
| `tests/` | The engine, store, isolation, and API tests, with the 2026-09-15 test run in `tests/data/` |
| `Dockerfile` | One image with the API, the page, ea_irms, and LibreOffice |
| `DEPLOY.md` | The steps to run it on Google Cloud Run from this repository and to serve it at a domain through Firebase Hosting |
| `firebase.json`, `public/` | Firebase Hosting's configuration, which passes every request to the Cloud Run service; `public/` holds only a placeholder, so Hosting serves no files of its own |
| `.github/workflows/ci.yml` | The tests, the port check on three runs, the page in a browser, and an image build with a smoke test, on every push |

## Running it

```
docker build -t isoharness .
docker run -p 8000:8000 isoharness
```

Then http://localhost:8000 is the home page, http://localhost:8000/ea-irms the EA-IRMS page, and http://localhost:8000/api/docs lists the endpoints and lets you try each one. Without Docker, `pip install -r requirements.txt` and `uvicorn isoharness.app:app` run the same service, but its workbooks carry formulas only unless LibreOffice (`soffice`) is on the PATH.

## The API

These endpoints serve the EA-IRMS page. An instrument added later gets endpoints of its own, and these keep their paths.

| Endpoint | Returns |
| --- | --- |
| `GET /api/health` | `{"status": "ok"}` |
| `GET /api/version` | The isoharness and ea_irms versions, and whether LibreOffice is present |
| `GET /api/defaults` | The package default configuration, in the form the POST endpoints accept |
| `GET /api/example` | The 2026-09-15 test run, which the page offers as an example |
| `POST /api/process` | Every position's values and flags, the fitted steps, and the block the preview port recomputes from |
| `POST /api/workbook` | The correction workbook (.xlsx), with its formulas computed by LibreOffice unless `values=false` |
| `POST /api/results.csv` | The workbook's Results sheet as CSV, at full precision: each value followed by its U, then Flags and the u columns |

Each POST takes a multipart form with `file`, an export as .xls, .xlsx, .xlsm, .csv, .txt, or .tsv of at most 10 MB (`ISOHARNESS_MAX_UPLOAD_MB` changes the limit), and optionally `config`, a lab configuration as JSON in the form `/api/defaults` returns. `/api/workbook` also takes `run_name` and `run_date` for its Setup sheet. A file ea_irms cannot read is refused with a message that names the problem, such as the columns it lacks.

The service converts one workbook at a time, because LibreOffice conversions that run together share a profile; in a test on 2026-10-11, one of three conversions started together failed, so the image runs a single service process. Three requests sent together came back in 5.1, 6.8, and 8.5 s, one after another. The image starts LibreOffice once while it is built, which in a test on 2026-10-11 brought the first workbook from a fresh container down from 3.42 s to 2.47 s (one trial each).

## Uncertainty of %N, %C and C:N (ea_irms 1.4.3)

Each %N, %C and C:N carries u and U = k·u, with the k of the isotope values. For %N and %C, u combines five terms, each a fraction of the value: the instrument's repeatability, from the drift monitors' measured over certified ratios by successive differences in position, less the part their own weighing explains; the position's weighing, the balance's standard uncertainty (`balance_u_mg`, Setup B33) over its mass; the calibration, the interpolated elemental drift factor when that correction is in force and the k-factor line at the position's micrograms when it is not; the certificate uncertainty of the material that sets the scale (standards table columns u(%C) and u(%N)); and trueness, the anchors' RMS deviation from their certified values combined with their certificate uncertainty, as in the Nordtest TR 537 handbook. C:N takes the same terms without the weighing, which cancels in a ratio. Empty inputs count as 0. The workbook computes every term as a formula (Fits rows 300 to 335), and QC counts the anchors and linearity standards that fall within their U. The page does not show these yet.

## The preview port

`port/engine.js` recomputes blank, linearity, drift, calibration, elemental content, and U in the browser, so the page redraws on every switch, while the API's values stay the ones downloaded. It reads the user's decisions (the switches, the drift model, exclusions, and the sample repeatability) and, from the block's `settings`, the configuration's settings that act inside those steps (the significance level, the monitor outlier limit, the check-standard offset, the blank multiple, the coverage factor, the elemental gate, and the first drift position). Each position in the block also says whether the server set its CO2 peak aside as a reference pulse, which the port flags as the engine does. Settings that act before the block is built (the standards, the size limits, settling, carryover, and the reference-pulse windows) are applied by the server, so the page sends a newly loaded configuration back to `/api/process`.

On every upload the page also runs the port on the configuration's own decisions and compares each position's corrected values, U, %N, %C, C:N, elemental factors, and flags with the engine's from `/api/process`; the upload screen says whether they agree, and the steps warn when they do not.

The port check runs the engine on 643 decision sets with the configuration's settings (all 243 combinations of the switches and the drift model, then 400 random sets) and 360 more under twelve non-default settings, and requires the port to agree within 1e-9 relative on every value, flag list, and fit. Continuous integration runs it three ways on the 2026-09-15 test run: as exported, with a configuration that makes the acetanilide bypasses a check standard (`tests/data/config_check_standard.json`, accepted values set 0.3‰ from the run's own mean so that the check offset matters), and rewritten as a filtered export, so that the reference-pulse check sets CO2 peaks aside.

```
python scripts/port_cases.py tests/data/20260915_raw.csv --out build/port
node port/test_port.js build/port/port_input.json build/port/cases.json
python scripts/port_cases.py tests/data/20260915_raw.csv --config tests/data/config_check_standard.json --out build/port_check
node port/test_port.js build/port_check/port_input.json build/port_check/cases.json
python scripts/port_cases.py tests/data/20260915_raw.csv --as-filtered --out build/port_filtered
node port/test_port.js build/port_filtered/port_input.json build/port_filtered/cases.json
```

## Checks

The tests run with `pip install -r requirements-dev.txt` and `python -m pytest`. The store and isolation tests need Postgres at `EA_TEST_SUPER_DSN` and `EA_TEST_APP_DSN` (see `.github/workflows/ci.yml`), with a UTF8 database, since a SQL_ASCII one returns text as bytes and fails seven of them; the workbook tests need LibreOffice.

`python scripts/smoke.py http://localhost:8000 tests/data/20260915_raw.csv` checks a running service, requiring the workbook it returns to have its formulas computed and to hold the values `/api/process` returned.

`python scripts/e2e_web.py` checks the page end to end: it starts the service on a free port, opens the home page and follows its link to the EA-IRMS page, drives that page in Chromium through the upload, a lab configuration, all ten screens from the instrument step to the results, five decisions, and the landing page's example run, and requires every value on the results screen to equal, to the shown precision, the results table the service returns for the same decisions built without the page, and the downloaded configuration to reproduce it. It fails on any console error or failed request, and saves screenshots (`--shots`). With `--url` it checks a service already running, such as the image or a deployment. It needs Playwright for Python (in `requirements-dev.txt`) with its Chromium (`python -m playwright install chromium`), and LibreOffice.

`python scripts/reference_check.py MANIFEST.json --url http://localhost:8000` is the engine check before a release, on reference runs that stay outside this repository (the manifest format is in the script). On each run the API's values must equal its workbook, they must equal ea_irms called directly, and the workbook must equal the one delivered from the same export, cell by cell.

Results on 2026-10-11, with the smoke test and the reference check run against the image built from this repository:

| Check | Result |
| --- | --- |
| `pytest` | 134 passed, none skipped |
| Port, 2026-09-15 test run, as exported | 1,003 decision sets (360 under non-default settings), 1,155,456 comparisons, 0 mismatches, worst relative difference 1.68e-13 |
| Port, the same with a check standard | 1,003 decision sets, 1,155,456 comparisons, 0 mismatches, worst relative difference 1.68e-13 |
| Port, the same as a filtered export (5 CO2 peaks set aside) | 1,003 decision sets, 1,155,456 comparisons, 0 mismatches, worst relative difference 1.10e-13 |
| Port, the nine reference runs with their delivered configurations | 1,003 decision sets each, 10,780,244 comparisons, 0 mismatches, worst relative difference 1.06e-11 |
| Smoke test against the image | 422 values and 56 flag lists, 0 problems |
| Page end to end (`scripts/e2e_web.py`), with uvicorn and against the image | 10 screens and five decisions; 208 values and 26 flag lists on the results screen equal the service's results table to the shown precision (all 208 formatted identically); workbook (422 computed values), table, and configuration round trip exact; 0 console errors, 0 failed requests |
| Reference check against the image, nine runs | 25,527 values against the workbook, 33,995 against the engine, and 778,464 cell comparisons (formulas and computed values) against the workbooks delivered for eight of the runs; 0 mismatches |

## Data in this repository

The only real run here is the 2026-09-15 test run (standards, blanks, and one position named Random), in `tests/data/` as a CSV export and as the reference workbook the engine tests compare against. Every other run is lab data, checked with `scripts/reference_check.py` outside this repository.

## License

MIT, copyright 2026 Xingchen Wang; see `LICENSE`.
