# Putting isoharness on Google Cloud Run

These steps take the repository from GitHub to a public address on Cloud Run, rebuilt from the `main` branch on every push; a later section puts the site on a domain of its own through Firebase Hosting. They follow Google's own pages, read on 2026-10-11 unless marked: [continuous deployment](https://docs.cloud.google.com/run/docs/continuous-deployment), [billing settings](https://docs.cloud.google.com/run/docs/configuring/billing-settings), [maximum instances](https://docs.cloud.google.com/run/docs/configuring/max-instances), [public access](https://docs.cloud.google.com/run/docs/authenticating/public), [service identity](https://docs.cloud.google.com/run/docs/securing/service-identity) (2026-10-12), [budgets](https://docs.cloud.google.com/billing/docs/how-to/budgets), [spend caps](https://docs.cloud.google.com/billing/docs/how-to/budgets-spend-caps), [Cloud Run pricing](https://cloud.google.com/run/pricing), and [network pricing](https://cloud.google.com/vpc/network-pricing). Button names change from time to time; if one here is missing, the linked page has the current one.

The service stores nothing, so this phase needs no database and no secrets. Sign-in and saved runs come in phase three, with the Neon database; the last section, The database (phase three), puts its schema in place beforehand.

## What you need

- A Google Cloud project with a billing account attached. The OAuth client for Google sign-in will live in a Google Cloud project too (phase three), so one project can hold both.
- Owner rights on the project. Google's continuous-deployment page lists the roles a narrower account would need: Artifact Registry Administrator, Cloud Build Editor, Cloud Run Developer, Service Account User, and Service Usage Admin. Creating the service's own account in step 5 also needs Create Service Accounts.
- The repository on GitHub, here `xingchentonywang/isoharness`.

## Choose the region

Pick the Cloud Run region nearest the Neon project's region, which the Neon console shows, so that phase three's database calls stay short. Neon's `aws-us-east-1` is in Northern Virginia, for example, where Cloud Run's `us-east4` is. Until phase three any region works; the US regions on Cloud Run's Tier 1 price list include `us-east1`, `us-east4`, `us-east5`, `us-central1`, and `us-west1` (prices in the architecture note's Hosting section, read 2026-10-11).

## Create the service from the repository

1. In the Google Cloud console, open **Cloud Run** and click **Connect repository**. The **Create service** form opens.
2. Select **Cloud Build**, then open **Set up with Cloud Build**. Enable the Cloud Build API if the console asks.
3. Choose GitHub as the provider and click **Authenticate** if asked. If the panel says the GitHub App is not installed, click **Install Google Cloud Build** and install it on this repository alone; GitHub may first ask you to confirm your identity. Pick `xingchentonywang/isoharness`, tick the box that acknowledges that the repository's content will be transferred to this Google Cloud project, and click **Next**. If the repository is not listed, the **Manage connected repositories** link adds it.
4. Under **Build Configuration**, enter `^main$` as the **Branch**. For **Build Type** choose **Dockerfile**, with `/Dockerfile` as the **Source location**. Click **Save**.
5. Back in **Create service**, set:
   - **Service name**: `isoharness`.
   - **Region**: the one chosen above.
   - **Authentication**: **Allow public access**, so anyone can open the page. A project under an organization policy that restricts sharing to a domain needs the other route on Google's public-access page.
   - **Billing**: **Request-based**, the default, so the service is charged only while it handles a request.
   - **Service scaling**: **Minimum number of instances** 0, so an idle site costs nothing, and **Maximum number of instances** 1. The default maximum is 100; one instance limits how fast a burst of traffic can run up costs, and the service converts one workbook at a time anyway.
   - **Container**: leave the port at 8080 (Cloud Run hands it to the container as `PORT`, which the image listens on), and set **Memory** to 1 GiB and **CPU** to 1.
   - **Variables**: none are needed. `ISOHARNESS_UPLOADS_PER_HOUR` (default 60) sets the workbooks each address may ask for in an hour, `ISOHARNESS_READS_PER_HOUR` (default 600) its other uploads, `ISOHARNESS_MAX_UPLOAD_MB` (default 10) the largest file and the largest request of the combined results, and `ISOHARNESS_MAX_BATCH_RUNS` (default 20) the most runs combined in one request; from isoharness 0.5.0 a request body larger than `ISOHARNESS_MAX_UPLOAD_MB` and 1 MB more is refused (413) as it arrives. From isoharness 0.4.1, `ISOHARNESS_PROXY_HOPS` (default 2) is the number of proxies in front of the service (see The upload limit), and `ISOHARNESS_WORKBOOK_QUEUE` (default 3) and `ISOHARNESS_WORKBOOK_WAIT_S` (default 45) bound the workbook requests that wait for LibreOffice (see Workbook requests). `ISOHARNESS_API_DOCS` is not set by default, and then `/api/docs` and `/api/openapi.json` answer 404 (isoharness 0.5.0); set it to 1 to serve them. Leave it unset on isoharness.org, since the documentation page loads its scripts from a CDN and is served without the site's content policy.
   - **Service account**, on the **Security** tab of **Containers, Networking, Security**: choose **Create new service account**, name it `isoharness-run`, click **Create**, skip the roles, and click **Done**. On 2026-10-12 the form left this field empty and required it. The other choice, the Compute Engine default service account, held the Editor role on this project, which a service that opens files uploaded by anyone should not carry; Google's service identity page recommends an account of your own with the fewest permissions, and this phase needs none.
   - Leave the other settings as they are.
6. Click **Create**. The service page shows the build, then the address, which ends in `run.app`.

Every later push to `main` builds the image again and deploys it. **Edit repo settings** on the service page opens the trigger in Cloud Build, and the build step's **logs** link shows each build. The trigger's docker step builds without BuildKit, so the `Dockerfile` uses no BuildKit features; the first build, on 2026-10-12, stopped at the first `RUN --mount` line, and the next commit removed both. The trigger made this way builds, pushes, and deploys without running a test; Deploy only after the tests pass, below, has it run the repository's `cloudbuild.yaml` instead.

## Check the deployment

Open the address: the home page should list the instruments, and the **Open** button for EA-IRMS should lead to the EA-IRMS page. From a computer with the repository,

```
python scripts/smoke.py https://YOUR-ADDRESS.run.app tests/data/20260915_raw.csv
python scripts/e2e_web.py --url https://YOUR-ADDRESS.run.app
```

check the service and the page as continuous integration does. `/api/version` names the isoharness and ea_irms versions the deployment runs and, from isoharness 0.5.0, its LibreOffice (`libreoffice_version`). The first request after an idle spell waits while an instance starts; that wait has not been timed on Cloud Run.

## Deploy only after the tests pass

The trigger that Cloud Run made in Create the service runs three steps, build, push, and deploy, so a push to `main` reaches the service a few minutes later whether or not its tests pass; GitHub's continuous integration runs the tests at the same time but cannot stop the deployment. From isoharness 0.5.0 the repository holds `cloudbuild.yaml`, which runs the tests first:

| Step | What it does |
| --- | --- |
| Build | Builds the image from the `Dockerfile`, without BuildKit, as the trigger did |
| Database | Starts Postgres 16 in a container named `pg` on the build's network |
| Tests | Runs the whole test suite with LibreOffice and Node.js against that database, and fails if a test is skipped |
| Port, Port in node | Runs the preview port against the engine five ways: the three of continuous integration, the constructed carbon blank, and two positions without a reference background |
| Start the image, Smoke | Starts the new image, checks its home page and `engine.js`, and runs `scripts/smoke.py` against it |
| Push | Pushes the image to Artifact Registry |
| Deploy | Deploys it as a new revision of the service |

Push starts only once every step before it has passed, so a commit whose tests fail is not deployed and the service keeps the revision it has; the build's page in Cloud Build shows which step failed, with its log. The page's browser test, `scripts/e2e_web.py`, still runs only in GitHub's continuous integration. Images from Docker Hub come through `mirror.gcr.io`, Google's cache of them, which avoids Docker Hub's limits on pulls without an account, and the whole build stops after 40 minutes (`timeout` in the file). The file was rehearsed on 2026-10-16 with a local Docker daemon in place of Cloud Build, each step in its own image with its own arguments (and the rehearsing machine's proxy certificate passed to the image build), on the code of isoharness 0.5.0: the tests, 430 passed and none skipped (397 s for the suite, 486 s for the step); the port, 1,161,374 comparisons in each of the five runs, 0 mismatches; the image, built in 89 s; and the smoke test against it, 687 values, 0 problems. Push and Deploy were not run, and the file has not yet run on Cloud Build itself.

To have the trigger run it:

1. On the service's page in Cloud Run, click **Edit repo settings**, which opens the trigger in Cloud Build. **Triggers** in Cloud Build lists it too.
2. Under **Configuration**, choose the type **Cloud Build configuration file (yaml or json)** and the location **Repository**, and enter `cloudbuild.yaml` as the file's location.
3. Keep the substitution variables the trigger lists, such as `_SERVICE_NAME` and `_DEPLOY_REGION`, and its service account, which already pushes the image and deploys the service. A value set on the trigger takes the place of the one in `cloudbuild.yaml`, which carries this service's values for a trigger that sets none.
4. Click **Save**.

Switch the trigger before the push that should be tested first, since a push runs whatever configuration the trigger has when it arrives. To try the switch without a push, click **Run** on the trigger's row in **Triggers** and run it on the branch `main`. These steps follow Google's page on [triggers](https://docs.cloud.google.com/build/docs/automating-builds/create-manage-triggers), read 2026-10-16. How long a build takes on Cloud Build has not been measured. Cloud Build gives each billing account 2,500 build-minutes a month at no charge on its default machine, e2-standard-2, and charges $0.006 a minute beyond that ([Cloud Build pricing](https://cloud.google.com/build/pricing), read 2026-10-16), so at 10 minutes a build (an assumption) the free minutes cover 250 pushes a month.

To serve an earlier revision again, open the service's **Revision history** tab, click the menu at the right of that revision, then **Manage traffic**, choose **Send all traffic to one revision** and the revision, and click **Save**; or run

```
gcloud run services update-traffic isoharness --to-revisions REVISION=100 --region us-east5
```

with the revision's name. Google's page says that later deployments keep that traffic split, so a revision deployed after it serves nothing until traffic goes to the latest revision again, through **Manage traffic** or `gcloud run services update-traffic isoharness --to-latest --region us-east5` ([rollouts, rollbacks, and traffic migration](https://docs.cloud.google.com/run/docs/rollouts-rollbacks-traffic-migration), read 2026-10-16).

## Building the same image again

The image is built from pinned parts where they can be pinned, so a rebuild computes with what was checked.

The base image is `python:3.13-slim-trixie@sha256:70729b46c69b4f1e97c4822c1af3df53a1476cf5ddc6c087c0c10bc3a5678c2f`, the digest of the multi-architecture index the tag named on 2026-10-15. Its linux/amd64 image, the one Cloud Run runs, is `sha256:8fb4cfa1a2616d7b8e0c2175cc6ad68f5729c34ea8488c0b360d2934b7be9024`, with Python 3.13.16 on Debian trixie, created 2026-10-06. `cloudbuild.yaml` runs its tests and the port in the same image by the same digest (`_PYTHON`), and `tests/test_hardening.py` fails when the two differ. To move to a newer base image, read the tag's digest again and put it in the `FROM` line, with the date in the comment above it, and in `_PYTHON`:

```
docker buildx imagetools inspect python:3.13-slim-trixie
```

prints it on its `Digest` line. Docker Hub sometimes answers 429 Too Many Requests; [mirror.gcr.io](https://mirror.gcr.io) serves the same images (`mirror.gcr.io/library/python:3.13-slim-trixie`), and the registry's own answer gives the digest without Docker:

```
curl -sI -H "Accept: application/vnd.oci.image.index.v1+json, application/vnd.docker.distribution.manifest.list.v2+json" https://mirror.gcr.io/v2/library/python/manifests/3.13-slim-trixie | grep -i docker-content-digest
```

Of the Python libraries, `requirements.txt` is the list to edit, and continuous integration and `cloudbuild.yaml` install it (through `requirements-dev.txt`). The image installs `requirements.lock` instead, with `pip install --require-hashes --no-deps --only-binary=:all: -r requirements.lock` and then `pip check`, so it gets each library at the version the lock names, the libraries they need included, from wheels whose hashes match. After a change to `requirements.txt`, make the lock again from the repository's root with [uv](https://docs.astral.sh/uv/) (`pip install uv`):

```
uv pip compile requirements.txt --generate-hashes --python-version 3.13 --python-platform x86_64-manylinux_2_28 --output-file requirements.lock
```

It keeps the versions the lock already pins unless `requirements.txt` asks for others; add `--upgrade` to move every library to its newest version. `tests/test_hardening.py` fails while the lock does not pin what `requirements.txt` names. The lock carries the hashes of every file of each version, so it also holds the wheels for Linux on arm64, though an arm64 build has not been tried; it is made for Linux, so it leaves out what a library needs only on Windows.

LibreOffice and the Liberation fonts come from Debian's archive and are not pinned, because the archive keeps only the current version of a package, so a pinned version would stop the build once Debian replaced it, as its security updates and point releases do. The build log lists the versions installed (the `dpkg-query` line of the LibreOffice step), and `GET /api/version` reports the LibreOffice of a running service in `libreoffice_version`. The image built on 2026-10-15 got `libreoffice-calc-nogui` 4:25.2.3-2+deb13u8, which reports "LibreOffice 25.2.3.2 520(Build:2)", and `fonts-liberation` 1:2.1.5-3.

## Serve it at a domain of its own

Cloud Run's own address keeps working. Firebase Hosting can also serve the site at a domain such as `isoharness.org`, answering on the domain, holding its certificate, and passing every request to the Cloud Run service. These steps follow Firebase's pages on [custom domains](https://firebase.google.com/docs/hosting/custom-domain), [caching](https://firebase.google.com/docs/hosting/manage-cache), [serving dynamic content](https://firebase.google.com/docs/hosting/serverless-overview), and [pricing](https://firebase.google.com/pricing), read on 2026-10-12, the day `isoharness.org`, bought at Porkbun, was set up this way.

1. In the [Firebase console](https://console.firebase.google.com), add Firebase to the existing Google Cloud project rather than creating a new one. Because the project has a billing account, it is on Firebase's pay-as-you-go Blaze plan; Google Analytics is not needed and can stay off.
2. Deploy Hosting's configuration, `firebase.json` in this repository. It sends every path to the Cloud Run service `isoharness` in `us-east5`, and `public/` holds only a placeholder that `ignore` keeps out of the upload, so Hosting serves no files of its own. From the repository's root, or from any folder holding a copy of `firebase.json` and an empty `public` folder, run

   ```
   firebase deploy --only hosting --project isoharness
   ```

   in a shell where the Firebase CLI is signed in to your Google account. On 2026-10-12 the terminal that the Firebase console's **Cloud Shell** button opens had the CLI installed and signed in, once it was authorized. The site then answers at `https://isoharness.web.app`. A service with another name or region needs those changed in `firebase.json`.
3. In the Firebase console, open **Hosting** and click **Add custom domain**. Enter the domain, here `isoharness.org`, and then add `www.isoharness.org` the same way, with the option that redirects all its requests to `isoharness.org`.
4. At the registrar, enter the records the console lists. For `isoharness.org` they were a TXT record on the domain itself with the value `hosting-site=isoharness`, an A record on the domain pointing to `199.36.158.100`, and a CNAME record for `www` pointing to `isoharness.web.app`. Remove every other A, AAAA, ALIAS, or CNAME record for those two names, such as a registrar's parking records, since Firebase cannot issue a certificate while records point elsewhere. Keep the TXT record afterwards; Firebase's page says Hosting needs it continually to issue and renew certificates.
5. Firebase checks the records and then issues the certificate. Its page says either can take up to 24 hours; on 2026-10-12 both names showed **Connected** the same evening.
6. Check the site as in **Check the deployment**, with `https://isoharness.org` in place of the Cloud Run address, and check that `https://www.isoharness.org` lands on it. If the domain's root shows Firebase's **Site Not Found** page while other paths work, Hosting's cache still holds the page it served before the domain was connected. On 2026-10-12 that happened at Hosting's servers near Washington, D.C., and running the deploy in step 2 again cleared it, as Firebase's caching page says any deploy does.

Requests through Hosting stop after 60 seconds with a 504 response, while the service itself allows 300 s; on 2026-10-12 one workbook took 3.85 s through the domain. Hosting caches none of the service's responses unless their `Cache-Control` header allows it, though by default it keeps a 404 for 10 minutes. Firebase recommends `us-west1`, `us-central1`, `us-east1`, `europe-west1`, or `asia-east1` for a service behind Hosting, to sit near Hosting's own servers; this service runs in `us-east5`, near its database, and works through Hosting all the same.

## Limit the spending

Google Cloud sets no hard limit on a whole billing account. Two budgets come close, a spend cap on Cloud Run that pauses the service and an alert over the whole account that only sends email. Set both before the service goes live; the amounts below, $15 and $20, are examples for spending of about $20 a month.

### A spend cap on Cloud Run

Cloud Run is one of the services a spend cap budget can pause. Each such budget covers a single project and a single service, and only first-party Google Cloud customers can set one, not accounts held through a reseller. Once the month's estimated cost for the service passes the budget's target, the service stays paused until the first day of the next month or until someone lifts the cap. Cloud Build and Artifact Registry cannot be capped, but only pushes to `main` add to their costs. The one-instance limit alone does not hold a monthly figure, since an instance with 1 vCPU and 1 GiB kept busy through a 730-hour month would cost about $64 after the free tier (calculated from the prices of 2026-10-11).

1. In the console, open **Billing**, then **Budgets & caps** (**Budgets & alerts** in Google's pages of 2026-10-11) under **Cost management**, and click **Create budget** (Google's spend cap page calls it **Create new budget**).
2. Under **Define**, choose **Spend cap enforcement**, enter a name, and click **Next**.
3. Under **Scope**, the time range is fixed at Monthly. Choose this project and the service **Cloud Run**, and click **Next**.
4. Under **Amount**, enter a **Target amount** such as $15, and click **Next**.
5. Under **Actions**, emails at 50, 80, and 100 percent are already set. Click **Finish**.

When the cap is enforced, Cloud Run takes no new requests in the project; requests already running finish, and Google deletes no resources or data. To lift the cap before the month ends, open **Budgets & caps**, click **View spend cap details** on the spend caps banner, click the budget's name, then **Lift spend cap** and **Confirm**. Google warns that the service may take up to an hour to return to normal.

Set the target below the real limit. Google enforces the cap on estimated costs that arrive with a delay, and it bills whatever is spent in the meantime. The estimates use list prices without savings or credits, so the cap pauses the service even while free credit remains; at list price, $15 buys about 157 hours of active time at 1 vCPU and 1 GiB (calculated). Google's page does not say whether the data the service sends to visitors counts toward the cap. That transfer costs $0.12 per GiB to North America after the first GiB each month, which is free, and a first visit to the EA-IRMS page gets 348,725 bytes from the service, about 349 KB, for the page, its four scripts, and two short API answers (measured on isoharness 0.5.0 on 2026-10-16). They travel uncompressed, as isoharness.org sent its page and `logic.js` that day, and the page's fonts come from Google Fonts, not from the service.

### An alert on the whole account

A second budget covers the whole billing account and only sends email; it does not stop the service or the charges.

1. In **Budgets & caps**, click **Create budget**, choose **Alerts only**, enter a name, and click **Next**.
2. Under **Scope**, choose **Select all** for **Projects** and for **Services**, and clear every **Savings** option, so the budget counts usage before discounts and credits. Click **Next**.
3. Under **Amount**, choose **Specified amount**, enter a **Target amount** such as $20, and click **Next**.
4. Under **Actions**, keep the default thresholds of 50, 90, and 100 percent and the emails to the billing account's administrators and users, and click **Finish**.

Firebase Hosting, when the site uses it, is a service of its own, so the spend cap on Cloud Run does not cover it; this alert counts it with everything else. Firebase's pricing page lists 360 MB of Hosting transfer a day at no cost on the Blaze plan and $0.15 per GB beyond that, and it does not say how responses passed on from Cloud Run are counted (read 2026-10-12).

## Alerts when the site fails

Two alerts from Cloud Monitoring send email when something goes wrong: an uptime check that asks for the service's health from three places every 5 minutes, and an alert on the server errors the service logs. Cloud Monitoring's pricing page lists 1 million uptime-check executions a month at no charge in each project and $0.30 per 1,000 beyond that, and says that alerting will be charged starting no sooner than September 1, 2027, with alerting policies on uptime metrics not charged; it does not say how an alert on log entries will be charged ([Cloud Monitoring pricing](https://cloud.google.com/stackdriver/pricing), read 2026-10-16). A check every 5 minutes from three checkers makes 26,784 executions in a 31-day month. Each is a request to the service, and billed one by one at Cloud Run's minimum of 100 ms on a running instance they would cost about $0.08 a month at list price, before the free tier (both calculated, from Cloud Run's prices as read on 2026-10-16, unchanged since 2026-10-11).

The uptime check, from Google's page on [uptime checks](https://docs.cloud.google.com/monitoring/uptime-checks), read 2026-10-16:

1. In the console, open **Monitoring**, then **Uptime checks**, and click **Create Uptime Check**.
2. Set **Protocol** to HTTPS, **Resource Type** to URL, **Hostname** to `isoharness.org`, and **Path** to `/api/health`, and set the check frequency to 5 minutes. Under **More target options**, set **Regions** to USA, whose three checkers (Oregon, Iowa, and Virginia) are the fewest a check may have, and keep **Validate SSL certificates**. Click **Continue**.
3. Turn content matching on, choose the match type that requires the response to contain a text, and enter `"status":"ok"`, which `/api/health` returns. Click **Continue**.
4. Choose to create an alert, open **Notification channels**, and choose your email address; **Manage notification channels** adds one if none is listed. Click **Continue**.
5. Enter a title such as `isoharness.org health`, click **Test**, and then **Create**.

The alert on server errors, from Google's page on [log-based alerting policies](https://docs.cloud.google.com/logging/docs/alerting/log-based-alerts), read 2026-10-16: open **Logs Explorer**, enter

```
resource.type="cloud_run_revision"
resource.labels.service_name="isoharness"
(httpRequest.status>=500 OR severity>=ERROR)
```

click **Run query**, then **Actions** and **Create log alert**. Give it a name, keep the query, choose 5 minutes or more as the time between notifications, choose the same email channel, and click **Save**. A refused upload (413 or 429) is not a server error and sends nothing; a workbook request told to try again in a minute (503) is one, and sends email.

## The upload limit

The service counts each address's uploads (every POST endpoint: those that take a file, the sample-sheet template, and, from isoharness 0.5.0, the configuration written as Excel, `/api/config.xlsx`; a request of the combined results counts once) over the last hour, workbooks and the rest apart, and refuses the next one with a message once it reaches `ISOHARNESS_UPLOADS_PER_HOUR` for workbooks, which LibreOffice computes (a workbook request for the test run took 8.96 s on isoharness.org on 2026-10-13), or `ISOHARNESS_READS_PER_HOUR` for the rest, which the engine answers alone (`/api/process` reported 0.011 s of processing for the test run, measured locally the same day). From isoharness 0.2.0 the two are apart, because a batch sends one request per run, and a new configuration one more per run. The count is kept in the running instance, so it starts again when an instance starts, and from isoharness 0.4.1 the service keeps at most 20,000 counters, forgetting the least recently used first; an address takes one for its uploads and one for its workbooks.

The address comes from `X-Forwarded-For`. Each proxy adds the address it received the request from at the right-hand end of that header, and whatever a client writes there itself stays at the left, so from isoharness 0.4.1 the service counts the address `ISOHARNESS_PROXY_HOPS` places from the right (default 2, for Firebase Hosting and Cloud Run's own front end); up to 0.4.0 it counted the left-most address, which any client can write, so a made-up header reset the limit (finding R-05 of the review of 2026-10-14). A request that reaches the service through fewer proxies, such as one sent to the `run.app` address itself, is read one place too far to the left, where a client can write, so there the limit still holds only for clients that do not try to get around it. The default of 2 was checked on isoharness.org on 2026-10-14, when `/api/client` counted 2 addresses in the header from two different connections. To check a deployment, open `https://isoharness.org/api/client` from a connection without a proxy of its own: `forwarded_for_entries` is how many addresses the header held, which on such a connection is the number of proxies, and `address` should be your own public address. If `forwarded_for_entries` is not 2, set `ISOHARNESS_PROXY_HOPS` to it in the service's variables and deploy again. Turning the `run.app` address off would close that gap, but Cloud Run's page on [ingress](https://docs.cloud.google.com/run/docs/securing/ingress) lists Firebase Hosting among the services that use the default `run.app` URL to invoke Cloud Run and warns that disabling it stops them from working (read 2026-10-16, the page updated 2026-10-09), so while Hosting serves isoharness.org the address stays on. The gap does not raise what the service can cost, since the cap of one instance limits how fast costs can grow and the spend cap limits how far.

## Workbook requests

The service computes one workbook at a time with LibreOffice, because conversions that run together share a profile. From isoharness 0.4.1 a workbook request waits at most `ISOHARNESS_WORKBOOK_WAIT_S` seconds (default 45) for its turn, and once `ISOHARNESS_WORKBOOK_QUEUE` requests (default 3) are waiting, the next is told to try again in a minute (503, with `Retry-After: 60`). Firebase Hosting stops waiting for an answer after 60 seconds, and 13 workbook requests to isoharness.org on 2026-10-13 took 4.8 to 9.7 s each, so a wait of 45 s leaves time for the workbook itself, where one of 60 s would have computed some for no one. LibreOffice runs in a process group of its own and is stopped, with everything it started, after 120 seconds (`ea_irms.workbook.BAKE_TIMEOUT_S`).

## The database (phase three)

isoharness 0.5.0 does not connect to a database. When phase three does, it will use the schema in `db/`, which these steps put in place on Neon beforehand. Run them from a shell where `DATABASE_URL` holds the connection string of the role that owns the tables, Neon's direct connection rather than the pooled one, typed or pasted into that shell only, never into a chat or a file in the repository. First check the role, and which migrations the database already has:

```
psql "$DATABASE_URL" -c "SELECT rolname, rolsuper, rolbypassrls FROM pg_roles WHERE rolname = current_user;"
psql "$DATABASE_URL" -Atc "SELECT to_regclass('public.lab_members') IS NOT NULL, EXISTS (SELECT 1 FROM information_schema.columns WHERE table_schema = 'public' AND table_name = 'runs' AND column_name = 'u_sample_n'), to_regprocedure('app.is_owner(uuid)') IS NOT NULL;"
```

The first must show `t` under `rolsuper` or `rolbypassrls`. The functions the policies call read `lab_members` past its own policy, which only such a role can do; under any other role every query of the application fails with "stack depth limit exceeded" (checked on Postgres 16), and 003 refuses to run, with "run 003 as a superuser or a role with BYPASSRLS". Whether Neon's owner role shows `t` has not been checked, so run these steps on a Neon branch before the main one. The second prints one value for each of 001, 002, and 003, `t` where the migration is in place: `t|t|f` means only 003 is missing. Apply what is missing, in order:

```
psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -f db/001_schema.sql
psql "$DATABASE_URL" -v ON_ERROR_STOP=1 --single-transaction -f db/002_decisions_and_uncertainty.sql
psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -f db/003_owner_rules.sql
```

001 and 003 are each one transaction, and `--single-transaction` makes 002 one, so each applies whole or not at all, and running one a second time stops with "already exists" and changes nothing (each checked on Postgres 16 on 2026-10-16). After 003 the two queries

```
psql "$DATABASE_URL" -Atc "SELECT count(*) FROM pg_policies WHERE schemaname = 'public';"
psql "$DATABASE_URL" -Atc "SELECT tgname FROM pg_trigger WHERE tgrelid = 'lab_members'::regclass AND NOT tgisinternal;"
```

print 29, the 28 policies of 003 and `app_users_self` of 001, and `keep_an_owner`. 003 replaces each table's one policy with one per command: any member of a lab reads all it holds, adds runs, adds and edits instruments and configurations, and may leave the lab, while only an owner adds or removes members, changes a role, changes or deletes a saved run, deletes an instrument or a configuration, or renames or deletes the lab, and a lab always keeps one owner. A refused change or delete affects no rows rather than raising an error, so the application has to check row counts. There is no migration back: returning to 001's rules means dropping 003's policies, its trigger, and the functions `app.keep_an_owner` and `app.is_owner`, and creating 001's seven policies again. `ea_irms.store.delete_lab_data` calls `app.is_owner`, so 003 has to be in the database before a version of the service that calls it is deployed.
