# isoharness

isoharness is a website for stable isotope data reduction, one instrument at a time. EA-IRMS comes first: it corrects Thermo EA-IRMS carbon and nitrogen runs, with numbers from ea_irms, the Python engine behind the lab's EA-IRMS correction workbook, which this repository carries at version 1.4.6. GasBench N₂O (nitrate by the denitrifier method, and organic nitrogen after persulfate oxidation) and GasBench carbonate are planned.

## Where it stands

Phase one put the engine behind a web API, LibreOffice in the container image to compute the workbook's formulas, and the browser's preview port under a check against the engine in continuous integration. Phase two adds the web pages: a home page at `/` that lists the instruments, and the EA-IRMS page at `/ea-irms`, where you upload an export, or several at once, check the instrument's reference gas and background, go through the eight correction steps, and download the workbook, the results table, and the configuration, and for several runs the combined results. Anyone can use the page without signing in. Sign-in, saving, and history come in phase three. The service stores nothing; each upload is processed in a temporary folder that is removed before the response is sent. It runs at [isoharness.org](https://isoharness.org), on Google Cloud Run behind Firebase Hosting, and [DEPLOY.md](DEPLOY.md) has the steps.

## Layout

| Path | What it is |
| --- | --- |
| `ea_irms/` | The engine, 1.4.6 (parsing, the corrections, the workbook, and the Postgres store that phase three will use); 1.4.3 adds the uncertainty of %N, %C, and C:N and puts each value's uncertainty beside it on the Results sheet, 1.4.4 marks the samples to measure again, 1.4.5 reports every uncertainty as 1σ, and 1.4.6 reads and writes the lab configuration as an Excel file |
| `isoharness/` | The web API (FastAPI), 0.2.0; 0.2.0 adds several runs at once (`batch.py`, the combined results) |
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
| `POST /api/results.csv` | The workbook's Results sheet as CSV, at full precision: each value followed by its 1σ (headed 1sigma), then Flags, the rerun mark with its reason and a suggested mass, the elemental drift factors, and %C and %N as measured |
| `GET /api/config-template` | The lab configuration as an Excel template, filled with the package defaults |
| `POST /api/config.xlsx` | A configuration sent as JSON, `{"config": {...}, "name": "run"}`, written as an Excel lab configuration; missing fields take the package defaults |
| `POST /api/config/read` | A lab configuration file, the Excel template or JSON, read into the form the other POST endpoints accept, with notes on anything filled from the defaults or left out |
| `POST /api/batch/results.xlsx` | Several runs' results in one Excel file (isoharness 0.2.0): Results, every position of every run with a Run column first; Runs, one row per run; Repeats, the samples measured in more than one run |
| `POST /api/batch/results.csv` | The combined Results sheet as CSV: Run, then the columns of `/api/results.csv`, at full precision |

Each POST takes a multipart form with `file`, an export as .xls, .xlsx, .xlsm, .csv, .txt, or .tsv of at most 10 MB (`ISOHARNESS_MAX_UPLOAD_MB` changes the limit), and optionally `config`, a lab configuration as JSON in the form `/api/defaults` returns. `/api/workbook` also takes `run_name` and `run_date` for its Setup sheet. A file ea_irms cannot read is refused with a message that names the problem, such as the columns it lacks. The two batch endpoints take the exports in `files` instead, at most 20 in a request (`ISOHARNESS_MAX_BATCH_RUNS`) and 10 MB together; `configs`, a JSON list with one configuration per file, in order, or one JSON object for all of them; `run_dates`, a JSON list of dates or nulls; and, for the Excel file, `config_name`, the name of the lab configuration to print on its README sheet. Each counts as one upload toward the hourly limit, however many files it carries. From isoharness 0.2.0 the limit is 600 uploads per address in an hour (`ISOHARNESS_READS_PER_HOUR`) for every POST endpoint but the workbook, and 60 for the workbook (`ISOHARNESS_UPLOADS_PER_HOUR`), the one LibreOffice computes.

The service converts one workbook at a time, because LibreOffice conversions that run together share a profile; in a test on 2026-10-11, one of three conversions started together failed, so the image runs a single service process. Three requests sent together came back in 5.1, 6.8, and 8.5 s, one after another. The image starts LibreOffice once while it is built, which in a test on 2026-10-11 brought the first workbook from a fresh container down from 3.42 s to 2.47 s (one trial each).

## Several runs at once (isoharness 0.2.0)

The EA-IRMS page takes several exports at once, chosen together or dropped together, up to 20 runs. Each is sent to `/api/process` on its own and corrected on its own, with the standards in that run; one lab configuration applies to all of them, and choosing another processes every run again, or none if any is refused. Each run keeps its own choices: the batch list on the upload screen and the run list under the run's name beside the steps show any run on the screen where you are, with the choices it was left with, and an export with the name of a run already in the batch replaces that run. The run date of each comes from its file name, and the date field sets the date of the run shown. On the Results screen, the downloads add the combined results, an Excel file, and the combined results table, a CSV file; a section below the samples lists every run with the corrections in force and a button for its own workbook, and the samples measured in more than one run, matched by name, with the mean, SD, and range of their δ¹⁵N and δ¹³C. The workbooks are still built one run per request.

The combined results come from `isoharness/batch.py`, which corrects nothing itself: it puts each run's Results rows, from the engine, one run after another with the run's file name in a Run column first; writes one row per run on the Runs sheet (file, run date, counts of positions, which corrections were in force, the calibration slopes, the samples marked for a rerun or a check, and the problems of the whole run); and lists on the Repeats sheet every analysis of a sample, or of a check standard, whose name appears in more than one run, with its difference from the mean of that name's analyses. A sample is matched by its identifier, ignoring case and surrounding spaces, and a check standard by its row of the standards table. The Results values are numbers, not formulas, at the 16 significant digits an Excel file keeps; the CSV file keeps every digit. Differences between runs are not part of any 1σ.

## Uncertainty of %N, %C, and C:N (ea_irms 1.4.3)

Each %N, %C, and C:N carries a 1σ. For %N and %C, 1σ combines five terms, each a fraction of the value: the instrument's repeatability, from the drift monitors' measured over certified ratios by successive differences in position, less the part their own weighing explains; the position's weighing, the balance's standard uncertainty (`balance_u_mg`, Setup B33) over its mass; the calibration, the interpolated elemental drift factor when that correction is in force and the k-factor line at the position's micrograms when it is not; the certificate uncertainty of the material that sets the scale (standards table columns 1σ(%C) and 1σ(%N)); and trueness, the anchors' RMS deviation from their certified values combined with their certificate uncertainty, as in the Nordtest TR 537 handbook. C:N takes the same terms without the weighing, which cancels in a ratio. Empty inputs count as 0. The workbook computes every term as a formula (Fits rows 300 to 335), and QC counts the anchors and linearity standards that fall within k × 1σ of their certified values. The page does not show these yet.

## Uncertainty as 1σ (ea_irms 1.4.5)

Every uncertainty that the workbook, the results table, and the page report is one standard deviation, 1σ, the combined standard uncertainty; up to 1.4.4 the workbook put the expanded uncertainty U = k·u beside each value and the u columns after Flags, and from 1.4.5 it puts 1σ beside each value and drops the separate block, so the elemental drift factors and the values as measured move from AA to AD to U to X. The checks still compare against k × 1σ, with k (`coverage_k`, Setup B32) 2 by default: the anchors and linearity standards within it on QC, the anchors outside it among the problems of the whole run on the Rerun sheet, and the page's counts of samples that reversing a correction would move by more than it. The results table heads its columns 1sigma(...) rather than 1σ(...), since Excel misreads a Greek letter in a CSV file that has no byte-order mark. The engine and the API keep their names, `u_*` for 1σ and `U_*` for k × 1σ.

## The lab configuration as an Excel file (ea_irms 1.4.6)

A lab configuration can be an Excel file as well as JSON. `ea_irms/labconfig.py` writes it with three sheets: README; Settings, one row per setting with its value, what it means, and a grey Key column that is how the file is read; and Standards, the standards table in the column order of the workbook's Setup sheet, at most 16 rows. The page offers the template, filled with the package defaults, on the upload screen, and reads a chosen .xlsx file through `/api/config/read`; the Configuration download after a run is the same Excel file with that run's choices, and a .json configuration still works. On the command line, `python -m ea_irms --write-config lab.xlsx` writes the template and `--config lab.xlsx` reads one. An empty cell means off or not given where the setting allows that, as on the Setup sheet; where it does not, an empty cell is refused rather than replaced by a default. Every cell that cannot be used is named in one message, such as `Settings B6 (blank_correction) is 'maybe'; it needs one of On, Off, Auto`. A row missing from the file, as in one written by an older version, takes the package default, and a key this version does not read is left out; both are reported as notes. The last rows of Settings (the first drift position and the positions left out) belong to one run. An Excel file keeps 16 significant digits, so a value computed to more digits, such as the drift monitors' SD saved as the sample repeatability, comes back changed in the 16th or 17th digit; over 200,000 random values the largest relative change was 5.6e-16.

## Rerun marks (ea_irms 1.4.4)

A sample is marked only for what measuring it again can fix: the amount, the peak, or the state of the instrument at its position, never its value. It is marked for a rerun of N, C, or both when its peak lies outside the range of the linearity standards or under the blank multiple (`min_blank_multiple`, 3 by default), when it follows an oversized sample and its background rose by more than `carryover_mv` (2.0 mV by default), or when the undesignated N2 or CO2 reference pulse at its position lies more than `ref_pulse_off_sd` robust SDs (4 by default) from the median of the run's pulses. It is marked for a check when it ran before the background settled or carries the peak of only one element. A δ value outside the anchors' range stays a flag and is not marked, because a rerun would extrapolate the same way.

For a rerun the workbook suggests a mass range that puts each peak inside the linearity range and above the blank multiple, from that analysis's own peak area per mg; when no one mass suits both N and C, it gives both ranges for two capsules. Problems of the whole run are listed once, at the top of a new Rerun sheet: a check standard further than `check_offset_permil` from its accepted value, anchors outside k × 1σ (2σ by default) for %N, %C, or C:N, and a reference-gas SD at or above `ref_gas_sd_limit` (0.1 permil by default, the zero-enrichment limit of the FIRMS Good Practice Guide, 2nd ed., section 4.3.2). Results carries the mark, its reason, and the suggested mass after Flags, with rows shaded orange (rerun) or pale orange (check), and the Rerun sheet lists the marked samples in the order they ran. Replicates are not compared yet; that needs a sample sheet saying which positions hold the same sample. The page does not show the marks yet.

## The preview port

`port/engine.js` recomputes blank, linearity, drift, calibration, elemental content, and the uncertainties in the browser, so the page redraws on every switch, while the API's values stay the ones downloaded. It reads the user's decisions (the switches, the drift model, exclusions, and the sample repeatability) and, from the block's `settings`, the configuration's settings that act inside those steps (the significance level, the monitor outlier limit, the check-standard offset, the blank multiple, the multiple k of 1σ that the checks use, the elemental gate, and the first drift position). Each position in the block also says whether the server set its CO2 peak aside as a reference pulse, which the port flags as the engine does. Settings that act before the block is built (the standards, the size limits, settling, carryover, and the reference-pulse windows) are applied by the server, so the page sends a newly loaded configuration back to `/api/process`.

On every upload the page also runs the port on the configuration's own decisions and compares each position's corrected values, uncertainties, %N, %C, C:N, elemental factors, and flags with the engine's from `/api/process`; the upload screen says whether they agree, and the steps warn when they do not.

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

`python scripts/e2e_web.py` checks the page end to end: it starts the service on a free port, opens the home page and follows its link to the EA-IRMS page, drives that page in Chromium through the upload, a lab configuration, all ten screens from the instrument step to the results, five decisions, a batch of three runs with the combined results, and the landing page's example run, and requires every value on the results screen to equal, to the shown precision, the results table the service returns for the same decisions built without the page, and the downloaded configuration to reproduce it. It fails on any console error or failed request, and saves screenshots (`--shots`). With `--url` it checks a service already running, such as the image or a deployment. It needs Playwright for Python (in `requirements-dev.txt`) with its Chromium (`python -m playwright install chromium`), and LibreOffice.

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
