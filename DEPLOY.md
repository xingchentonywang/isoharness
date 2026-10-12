# Putting isoharness on Google Cloud Run

These steps take the repository from GitHub to a public address on Cloud Run, rebuilt from the `main` branch on every push. They follow Google's own pages as read on 2026-10-11: [continuous deployment](https://docs.cloud.google.com/run/docs/continuous-deployment), [billing settings](https://docs.cloud.google.com/run/docs/configuring/billing-settings), [maximum instances](https://docs.cloud.google.com/run/docs/configuring/max-instances), [public access](https://docs.cloud.google.com/run/docs/authenticating/public), and [budgets](https://docs.cloud.google.com/billing/docs/how-to/budgets). Button names change from time to time; if one here is missing, the linked page has the current one.

The service stores nothing, so this phase needs no database and no secrets. Sign-in and saved runs come in phase three, with the Neon database.

## What you need

- A Google Cloud project with a billing account attached. The OAuth client for Google sign-in will live in a Google Cloud project too (phase three), so one project can hold both.
- Owner rights on the project. Google's continuous-deployment page lists the roles a narrower account would need: Artifact Registry Administrator, Cloud Build Editor, Cloud Run Developer, Service Account User, and Service Usage Admin.
- The repository on GitHub, here `xingchentonywang/isoharness`.

## Choose the region

Pick the Cloud Run region nearest the Neon project's region, which the Neon console shows, so that phase three's database calls stay short. Neon's `aws-us-east-1` is in Northern Virginia, for example, where Cloud Run's `us-east4` is. Until phase three any region works; the US regions on Cloud Run's Tier 1 price list include `us-east1`, `us-east4`, `us-east5`, `us-central1`, and `us-west1` (prices in the architecture note's Hosting section, read 2026-10-11).

## Create the service from the repository

1. In the Google Cloud console, open **Cloud Run** and click **Connect repository**. The **Create service** form opens.
2. Select **Cloud Build**, then open **Set up with Cloud Build**. Enable the Cloud Build API if the console asks.
3. Choose GitHub as the provider, click **Authenticate** if asked, pick `xingchentonywang/isoharness`, and click **Next**. If the repository is not listed, the **Manage connected repositories** link adds it.
4. Under **Build Configuration**, enter `^main$` as the **Branch**. For **Build Type** choose **Dockerfile**, with `/Dockerfile` as the **Source location**. Click **Save**.
5. Back in **Create service**, set:
   - **Service name**: `isoharness`.
   - **Region**: the one chosen above.
   - **Authentication**: **Allow public access**, so anyone can open the page. A project under an organization policy that restricts sharing to a domain needs the other route on Google's public-access page.
   - **Billing**: **Request-based**, the default, so the service is charged only while it handles a request.
   - **Service scaling**: **Minimum number of instances** 0, so an idle site costs nothing, and **Maximum number of instances** 1. The default maximum is 100; one instance caps the cost of a burst of traffic, and the service converts one workbook at a time anyway.
   - **Container**: leave the port at 8080 (Cloud Run hands it to the container as `PORT`, which the image listens on), and set **Memory** to 1 GiB and **CPU** to 1.
   - **Variables**: none are needed. `ISOHARNESS_UPLOADS_PER_HOUR` (default 60) sets the uploads each address may send in an hour, and `ISOHARNESS_MAX_UPLOAD_MB` (default 10) the largest file.
6. Click **Create**. The service page shows the build, then the address, which ends in `run.app`.

Every later push to `main` builds the image again and deploys it. **Edit repo settings** on the service page opens the trigger in Cloud Build, and the build step's **logs** link shows each build.

## Check the deployment

Open the address: the page should offer **Correct a run**. From a computer with the repository,

```
python scripts/smoke.py https://YOUR-ADDRESS.run.app tests/data/20260915_raw.csv
python scripts/e2e_web.py --url https://YOUR-ADDRESS.run.app
```

check the service and the page as continuous integration does. The first request after an idle spell waits while an instance starts; that wait has not been timed on Cloud Run.

## Set a budget alert

In the console, open **Billing**, then **Budgets & alerts** under **Cost management**, and click **Create budget**. Choose **Alerts only**, name it, keep the scope on this project, choose **Specified amount**, and enter a **Target amount** such as $10 a month. Google's default thresholds are 50, 90, and 100 percent of it, and the alerts go by email to the billing account's administrators. A budget of this kind only warns; it does not stop the service or the charges.

## The upload limit

The service counts each address's uploads (the three POST endpoints) over the last hour and refuses the next one with a message once it reaches `ISOHARNESS_UPLOADS_PER_HOUR`. The count is kept in the running instance, so it starts again when an instance starts. The address is the one uvicorn takes from `X-Forwarded-For`, which a determined client can set, so the limit is a courtesy against accidents; the cap of one instance and the budget alert are what bound the cost.
