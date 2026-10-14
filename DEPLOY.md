# Putting isoharness on Google Cloud Run

These steps take the repository from GitHub to a public address on Cloud Run, rebuilt from the `main` branch on every push; a later section puts the site on a domain of its own through Firebase Hosting. They follow Google's own pages, read on 2026-10-11 unless marked: [continuous deployment](https://docs.cloud.google.com/run/docs/continuous-deployment), [billing settings](https://docs.cloud.google.com/run/docs/configuring/billing-settings), [maximum instances](https://docs.cloud.google.com/run/docs/configuring/max-instances), [public access](https://docs.cloud.google.com/run/docs/authenticating/public), [service identity](https://docs.cloud.google.com/run/docs/securing/service-identity) (2026-10-12), [budgets](https://docs.cloud.google.com/billing/docs/how-to/budgets), [spend caps](https://docs.cloud.google.com/billing/docs/how-to/budgets-spend-caps), [Cloud Run pricing](https://cloud.google.com/run/pricing), and [network pricing](https://cloud.google.com/vpc/network-pricing). Button names change from time to time; if one here is missing, the linked page has the current one.

The service stores nothing, so this phase needs no database and no secrets. Sign-in and saved runs come in phase three, with the Neon database.

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
   - **Variables**: none are needed. `ISOHARNESS_UPLOADS_PER_HOUR` (default 60) sets the workbooks each address may ask for in an hour, `ISOHARNESS_READS_PER_HOUR` (default 600) its other uploads, `ISOHARNESS_MAX_UPLOAD_MB` (default 10) the largest file and the largest request of the combined results, and `ISOHARNESS_MAX_BATCH_RUNS` (default 20) the most runs combined in one request.
   - **Service account**, on the **Security** tab of **Containers, Networking, Security**: choose **Create new service account**, name it `isoharness-run`, click **Create**, skip the roles, and click **Done**. On 2026-10-12 the form left this field empty and required it. The other choice, the Compute Engine default service account, held the Editor role on this project, which a service that opens files uploaded by anyone should not carry; Google's service identity page recommends an account of your own with the fewest permissions, and this phase needs none.
   - Leave the other settings as they are.
6. Click **Create**. The service page shows the build, then the address, which ends in `run.app`.

Every later push to `main` builds the image again and deploys it. **Edit repo settings** on the service page opens the trigger in Cloud Build, and the build step's **logs** link shows each build. The trigger's docker step builds without BuildKit, so the `Dockerfile` uses no BuildKit features; the first build, on 2026-10-12, stopped at the first `RUN --mount` line, and the next commit removed both.

## Check the deployment

Open the address: the home page should list the instruments, and the **Open** button for EA-IRMS should lead to the EA-IRMS page. From a computer with the repository,

```
python scripts/smoke.py https://YOUR-ADDRESS.run.app tests/data/20260915_raw.csv
python scripts/e2e_web.py --url https://YOUR-ADDRESS.run.app
```

check the service and the page as continuous integration does. The first request after an idle spell waits while an instance starts; that wait has not been timed on Cloud Run.

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

Set the target below the real limit. Google enforces the cap on estimated costs that arrive with a delay, and it bills whatever is spent in the meantime. The estimates use list prices without savings or credits, so the cap pauses the service even while free credit remains; at list price, $15 buys about 157 hours of active time at 1 vCPU and 1 GiB (calculated). Google's page does not say whether the data the service sends to visitors counts toward the cap. That transfer costs $0.12 per GiB to North America after the first GiB each month, which is free, and a first visit to the EA-IRMS page sends about 226 KB, the size of its files.

### An alert on the whole account

A second budget covers the whole billing account and only sends email; it does not stop the service or the charges.

1. In **Budgets & caps**, click **Create budget**, choose **Alerts only**, enter a name, and click **Next**.
2. Under **Scope**, choose **Select all** for **Projects** and for **Services**, and clear every **Savings** option, so the budget counts usage before discounts and credits. Click **Next**.
3. Under **Amount**, choose **Specified amount**, enter a **Target amount** such as $20, and click **Next**.
4. Under **Actions**, keep the default thresholds of 50, 90, and 100 percent and the emails to the billing account's administrators and users, and click **Finish**.

Firebase Hosting, when the site uses it, is a service of its own, so the spend cap on Cloud Run does not cover it; this alert counts it with everything else. Firebase's pricing page lists 360 MB of Hosting transfer a day at no cost on the Blaze plan and $0.15 per GB beyond that, and it does not say how responses passed on from Cloud Run are counted (read 2026-10-12).

## The upload limit

The service counts each address's uploads (the POST endpoints that take a file, and the sample-sheet template, a request of the combined results counting once) over the last hour, workbooks and the rest apart, and refuses the next one with a message once it reaches `ISOHARNESS_UPLOADS_PER_HOUR` for workbooks, which LibreOffice computes (a workbook request for the test run took 8.96 s on isoharness.org on 2026-10-13), or `ISOHARNESS_READS_PER_HOUR` for the rest, which the engine answers alone (`/api/process` reported 0.011 s of processing for the test run, measured locally the same day). From isoharness 0.2.0 the two are apart, because a batch sends one request per run, and a new configuration one more per run. The count is kept in the running instance, so it starts again when an instance starts. The address is the one uvicorn takes from `X-Forwarded-For`, which a determined client can set, so the limit is a courtesy against accidents; the cap of one instance limits how fast costs can grow, and the spend cap limits how far. Behind Firebase Hosting, requests reach the service through Hosting's servers, which add addresses of their own to `X-Forwarded-For`. uvicorn counts the first address in that header, and whether Hosting leaves each visitor's own address first has not been tested.
