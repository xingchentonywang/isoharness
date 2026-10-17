"""Server hardening from the code review, each test named by what it holds in place.

R-25: a request body is counted as it arrives and refused (413) once it passes the upload limit, with or
without a Content-Length. R-28: the API's documentation page and its OpenAPI description are served only
with ISOHARNESS_API_DOCS=1. R-35: X-Frame-Options on every response, and no source maps. R-32:
/api/config.xlsx counts toward the reads allowance, as every other POST endpoint but the workbook does.
R-26: /api/version names the LibreOffice the service found, read once, and requirements.lock, which the
image installs, pins what requirements.txt names. cloudbuild.yaml deploys a commit only after its tests, the
port, and a smoke test of its image pass, testing in the Dockerfile's base image.
"""
import asyncio
import hashlib
import importlib.util
import re
import shutil
import time
from pathlib import Path

import httpx
import pytest
from fastapi.testclient import TestClient

import isoharness.app as service
from ea_irms.workbook import VERSION
from isoharness import __version__
from isoharness.app import LIMITED, app, limiter

ROOT = Path(__file__).resolve().parents[1]
RUN = ROOT / "tests" / "data" / "20260915_raw.csv"
BLOCK = 64 * 1024
BOUNDARY = "hardening-test"
client = TestClient(app)


def form_head(field="file", name="big.csv"):
    return (f'--{BOUNDARY}\r\nContent-Disposition: form-data; name="{field}"; filename="{name}"\r\n'
            "Content-Type: text/csv\r\n\r\n").encode()


FORM_TAIL = f"\r\n--{BOUNDARY}--\r\n".encode()
FORM_TYPE = {"Content-Type": f"multipart/form-data; boundary={BOUNDARY}"}


def stream(path, blocks, *, field="file", json_body=False, declared=None):
    """POST a body of ``blocks`` blocks of 64 KiB through httpx's ASGI transport, which hands the app the
    body one block at a time, as a server does, and without a Content-Length unless ``declared`` gives
    one. Returns the response and how many of the blocks (the form's head and tail included) the app
    asked for."""
    pulled = 0

    async def body():
        nonlocal pulled
        pulled += 1
        yield b'{"config": {}, "name": "' if json_body else form_head(field)
        for _ in range(blocks):
            pulled += 1
            yield b"x" * BLOCK
        pulled += 1
        yield b'"}' if json_body else FORM_TAIL

    headers = {"Content-Type": "application/json"} if json_body else dict(FORM_TYPE)
    if declared is not None:
        headers["Content-Length"] = str(declared)

    async def send():
        async with httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://test") as c:
            return await c.post(path, content=body(), headers=headers)

    return asyncio.run(send()), pulled


# R-25: the body limit


@pytest.mark.parametrize("path, field, json_body, detail", [
    ("/api/process", "file", False, "the upload is larger than the 1 MB this service accepts"),
    ("/api/batch/results.csv", "files", False,
     "the exports together are larger than the 1 MB this service accepts in one request"),
    ("/api/config.xlsx", None, True, "the upload is larger than the 1 MB this service accepts"),
])
def test_a_chunked_body_over_the_limit_is_refused_before_it_is_read(monkeypatch, path, field, json_body, detail):
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "1")      # with the form's 1 MB of room, 32 blocks
    r, pulled = stream(path, 160, field=field, json_body=json_body)
    assert r.request.headers["transfer-encoding"] == "chunked" and "content-length" not in r.request.headers
    assert r.status_code == 413 and r.json() == {"detail": detail}
    # refused at the block that passed 2 MB, of the 10 MB offered
    assert pulled == 33, pulled
    assert r.headers["connection"] == "close"                # a chunked body could go on without end
    assert r.headers["x-frame-options"] == "DENY" and r.headers["content-type"] == "application/json"


def test_a_chunked_body_the_server_hands_over_whole_is_refused_too(monkeypatch):
    """The test client hands the app the whole body in one message; it is counted all the same."""
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "1")

    def body():
        yield form_head()
        for _ in range(40):
            yield b"x" * BLOCK
        yield FORM_TAIL

    r = client.post("/api/process", content=body(), headers=FORM_TYPE)
    assert r.request.headers["transfer-encoding"] == "chunked"
    assert r.status_code == 413 and r.json() == {"detail": "the upload is larger than the 1 MB this service accepts"}


def test_a_declared_length_over_the_limit_is_refused_at_once(monkeypatch):
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "1")
    r, pulled = stream("/api/process", 160, declared=50 * 1024 * 1024)
    assert r.status_code == 413 and r.json() == {"detail": "the upload is larger than the 1 MB this service accepts"}
    assert pulled == 0                                       # not one block of the body was asked for
    assert "connection" not in r.headers and r.headers["x-frame-options"] == "DENY"

    # At the default limit, 10 MB and 1 MB of room: a declared length one byte over is refused, and one
    # of exactly the limit is not (this body, one byte, then fails as a form).
    monkeypatch.delenv("ISOHARNESS_MAX_UPLOAD_MB")
    room = 11 * 1024 * 1024
    r = client.post("/api/batch/results.xlsx", content=b"x", headers={**FORM_TYPE, "Content-Length": str(room + 1)})
    assert r.status_code == 413
    assert r.json() == {"detail": "the exports together are larger than the 10 MB this service accepts in one request"}
    r = client.post("/api/batch/results.xlsx", content=b"x", headers={**FORM_TYPE, "Content-Length": str(room)})
    assert r.status_code != 413


def test_a_refused_body_does_not_count_toward_the_hours_allowance(monkeypatch):
    monkeypatch.setenv("ISOHARNESS_READS_PER_HOUR", "1")
    limiter.reset()
    try:
        r = client.post("/api/process", content=b"x", headers={**FORM_TYPE, "Content-Length": str(50 * 1024 * 1024)})
        assert r.status_code == 413
        assert client.post("/api/process", files={"file": (RUN.name, RUN.read_bytes())}).status_code == 200
        assert client.post("/api/process", files={"file": (RUN.name, RUN.read_bytes())}).status_code == 429
    finally:
        limiter.reset()


def test_a_normal_upload_is_still_accepted():
    r = client.post("/api/process", files={"file": (RUN.name, RUN.read_bytes())})
    assert r.status_code == 200, r.text
    assert r.json()["file"] == {"name": RUN.name, "bytes": RUN.stat().st_size,
                                "sha256": hashlib.sha256(RUN.read_bytes()).hexdigest()}

    # in chunks, without a Content-Length: whole, through the test client ...
    form = form_head(name=RUN.name) + RUN.read_bytes() + FORM_TAIL
    r2 = client.post("/api/process", content=(form[i:i + 4096] for i in range(0, len(form), 4096)), headers=FORM_TYPE)
    assert r2.request.headers["transfer-encoding"] == "chunked"
    assert r2.status_code == 200 and r2.json()["positions"] == r.json()["positions"]

    # ... and block by block, as a server hands it over, with a body just under the limit
    async def blocks():
        for i in range(0, len(form), 4096):
            yield form[i:i + 4096]

    async def send():
        async with httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://test") as c:
            return await c.post("/api/process", content=blocks(), headers=FORM_TYPE)

    r3 = asyncio.run(send())
    assert r3.status_code == 200 and r3.json()["positions"] == r.json()["positions"]


def test_the_endpoints_still_check_each_file(monkeypatch):
    """A file over the upload limit in a body under max_body_bytes() is refused by the endpoint, with the
    message that names it, as before."""
    monkeypatch.setenv("ISOHARNESS_MAX_UPLOAD_MB", "1")
    big = b"x" * (1024 * 1024 + 1)
    r = client.post("/api/process", files={"file": ("big.csv", big)})
    assert r.status_code == 413 and r.json() == {"detail": "big.csv is larger than the 1 MB this service accepts"}
    r, pulled = stream("/api/results.csv", 20)               # 1.25 MB, under the 2 MB of the body limit
    assert r.status_code == 413 and r.json() == {"detail": "big.csv is larger than the 1 MB this service accepts"}
    assert pulled == 22


# R-28: the API documentation


def fresh_app(monkeypatch, docs):
    """The app as the service builds it when it starts, with ISOHARNESS_API_DOCS set to ``docs``
    (None: not set), loaded as a module of its own so the app the other tests use stays as it is."""
    if docs is None:
        monkeypatch.delenv("ISOHARNESS_API_DOCS", raising=False)
    else:
        monkeypatch.setenv("ISOHARNESS_API_DOCS", docs)
    spec = importlib.util.spec_from_file_location("isoharness._app_for_tests", ROOT / "isoharness" / "app.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.app


@pytest.mark.parametrize("docs", [None, "", "0", "true"])
def test_the_api_documentation_is_off_by_default(monkeypatch, docs):
    c = TestClient(fresh_app(monkeypatch, docs))
    for path in ("/api/docs", "/api/openapi.json", "/api/docs/oauth2-redirect"):
        r = c.get(path)
        assert r.status_code == 404 and r.json() == {"detail": "Not Found"}, path
    assert c.get("/api/health").json() == {"status": "ok"}


def test_the_api_documentation_with_isoharness_api_docs_1(monkeypatch):
    c = TestClient(fresh_app(monkeypatch, "1"))
    page = c.get("/api/docs")
    assert page.status_code == 200 and page.headers["content-type"].startswith("text/html")
    assert "/api/openapi.json" in page.text
    assert "content-security-policy" not in page.headers     # as before: the page loads its scripts from a CDN
    spec = c.get("/api/openapi.json")
    assert spec.status_code == 200 and spec.json()["info"]["version"] == __version__
    assert {"/api/process", "/api/workbook", "/api/config.xlsx"} <= set(spec.json()["paths"])


# R-35: X-Frame-Options and source maps


def test_x_frame_options_on_every_response(monkeypatch):
    responses = {
        "page": client.get("/ea-irms"),
        "home": client.get("/"),
        "script": client.get("/engine.js"),
        "api": client.get("/api/health"),
        "file": client.get("/api/config-template"),
        "not found": client.get("/no-such-page"),
        "refused upload": client.post("/api/process", files={"file": ("run.pdf", b"x")}),
    }
    monkeypatch.setenv("ISOHARNESS_READS_PER_HOUR", "1")
    limiter.reset()
    try:
        client.post("/api/config.xlsx", json={})
        responses["too many"] = client.post("/api/config.xlsx", json={})
    finally:
        limiter.reset()
    assert responses["too many"].status_code == 429 and responses["refused upload"].status_code == 415
    for what, r in responses.items():
        assert r.headers["x-frame-options"] == "DENY", what
        assert r.headers["x-content-type-options"] == "nosniff", what
    assert "frame-ancestors 'none'" in responses["page"].headers["content-security-policy"]
    assert "content-security-policy" not in responses["api"].headers


def test_x_frame_options_on_an_unexpected_error(monkeypatch):
    def fails(*args, **kwargs):
        raise RuntimeError("an error the service did not expect")

    monkeypatch.setattr(service, "config_to_dict", fails)          # what /api/defaults calls
    r = TestClient(app, raise_server_exceptions=False).get("/api/defaults")
    assert r.status_code == 500 and r.text == "Internal Server Error"
    assert r.headers["x-frame-options"] == "DENY" and r.headers["x-content-type-options"] == "nosniff"


def test_source_maps_are_not_served():
    # Preact's map is in the repository, beside the script that names it; .dockerignore keeps it out of the image.
    assert (ROOT / "web" / "vendor" / "preact" / "preact.min.js.map").is_file()
    for path in ("/vendor/preact/preact.min.js.map", "/vendor/preact/preact.min.js.MAP",
                 "/vendor/preact/./preact.min.js.map", "/vendor/preact/preact.min.js%2Emap",
                 "/vendor/preact/preact.min.js.map/"):
        r = client.get(path)
        assert r.status_code == 404 and r.json() == {"detail": "Not Found"}, path
        assert r.headers["x-frame-options"] == "DENY"
    assert client.head("/vendor/preact/preact.min.js.map").status_code == 404
    js = client.get("/vendor/preact/preact.min.js")
    assert js.status_code == 200 and js.headers["content-type"].startswith("text/javascript")


# R-32: /api/config.xlsx and the reads allowance


def test_every_post_endpoint_is_limited():
    posts = {route.path for route in app.routes if "POST" in (getattr(route, "methods", None) or ())}
    assert {"/api/process", "/api/workbook", "/api/config.xlsx"} <= posts
    assert posts <= set(LIMITED), sorted(posts - set(LIMITED))


def test_config_xlsx_counts_toward_the_reads_allowance(monkeypatch):
    monkeypatch.setenv("ISOHARNESS_READS_PER_HOUR", "2")
    monkeypatch.setenv("ISOHARNESS_UPLOADS_PER_HOUR", "5")
    limiter.reset()
    try:
        body = {"config": {"coverage_k": 3.0}, "name": "run"}
        assert client.post("/api/config.xlsx", json=body).status_code == 200
        assert client.post("/api/config.xlsx", json=body).status_code == 200
        r = client.post("/api/config.xlsx", json=body)
        assert r.status_code == 429 and "2 uploads in the last hour" in r.json()["detail"]
        assert int(r.headers["retry-after"]) > 0
        # one allowance with the other requests the engine answers alone; the workbook's is its own
        assert client.post("/api/process", files={"file": (RUN.name, RUN.read_bytes())}).status_code == 429
        assert client.post("/api/workbook", files={"file": (RUN.name, RUN.read_bytes())},
                           data={"values": "false"}).status_code == 200
    finally:
        limiter.reset()


# R-26: the LibreOffice version, and the lock file


def test_version_keeps_its_fields_and_names_libreoffice(monkeypatch):
    v = client.get("/api/version").json()
    present = bool(shutil.which("soffice") or shutil.which("libreoffice"))
    assert v["isoharness"] == __version__ and v["ea_irms"] == VERSION and v["libreoffice"] is present
    assert set(v) == {"isoharness", "ea_irms", "libreoffice", "libreoffice_version"}
    if present:
        assert re.match(r"LibreOffice \d+\.\d+", v["libreoffice_version"]), v
    else:
        assert v["libreoffice_version"] is None

    def started_again():
        raise AssertionError("LibreOffice was asked for its version again")

    # read once: later requests start no LibreOffice
    monkeypatch.setattr(service, "_read_libreoffice_version", started_again)
    assert client.get("/api/version").json() == v


def test_the_libreoffice_version_is_read_when_the_service_starts(monkeypatch):
    reads = []

    def read():
        reads.append(time.monotonic())
        return "LibreOffice 1.2.3.4"

    monkeypatch.setattr(service, "_libreoffice_version", {})         # as in a service that has just started
    monkeypatch.setattr(service, "_read_libreoffice_version", read)
    with TestClient(app) as c:                                       # runs the startup, as uvicorn does
        for _ in range(100):
            if reads:
                break
            time.sleep(0.02)
        assert len(reads) == 1                                       # read before any request asked
        assert c.get("/api/version").json()["libreoffice_version"] == "LibreOffice 1.2.3.4"
        assert c.get("/api/version").json()["libreoffice_version"] == "LibreOffice 1.2.3.4"
    assert len(reads) == 1


def test_the_libreoffice_version_is_null_without_libreoffice(monkeypatch):
    monkeypatch.setattr(service.shutil, "which", lambda name: None)
    assert service._read_libreoffice_version() is None


def _name(text):
    return re.sub(r"[-_.]+", "-", text).lower()


def test_the_lock_file_pins_what_requirements_txt_names():
    """requirements.lock, which the image installs with --require-hashes --no-deps, holds each library of
    requirements.txt at the version it names, and every library is pinned with its hashes. After a change
    to requirements.txt, make the lock again with the command in its first lines."""
    wanted = {}
    for line in (ROOT / "requirements.txt").read_text().splitlines():
        line = line.split("#")[0].strip()
        if line:
            name, version = line.split("==")
            wanted[_name(name)] = version.strip()
    text = (ROOT / "requirements.lock").read_text()
    pins = {}
    for entry in re.split(r"\n(?=[A-Za-z0-9])", text):
        m = re.match(r"([A-Za-z0-9][A-Za-z0-9._-]*)==([^\s\\;]+)", entry)
        if m:
            pins[_name(m.group(1))] = (m.group(2), entry.count("--hash=sha256:"))
    for name, version in wanted.items():
        assert pins.get(name, (None,))[0] == version, (
            f"requirements.txt names {name}=={version} and requirements.lock pins {pins.get(name, ('nothing',))[0]}; "
            "make the lock again with the command in its first lines")
    assert {"starlette", "pydantic", "pydantic-core", "anyio", "h11", "et-xmlfile"} <= set(pins)   # and what they need
    assert all(count > 0 for _, count in pins.values()), pins
    assert "uv pip compile requirements.txt --generate-hashes" in text


# cloudbuild.yaml, the Cloud Build steps that test a commit before it is deployed (isoharness 0.5.0). Read as
# text, since the tests install no YAML library.

def _cloudbuild():
    text = (ROOT / "cloudbuild.yaml").read_text()
    body = text.split("\nsteps:\n", 1)[1].split("\nimages:", 1)[0]
    steps = {}
    for block in ("\n" + body).split("\n  - id: ")[1:]:
        m = re.search(r"^    waitFor: \[(.*)\]$", block, re.M)
        steps[block.splitlines()[0].strip()] = (
            [w.strip().strip("'") for w in m.group(1).split(",")] if m else None, block)
    return text, steps


def test_cloudbuild_deploys_only_after_every_check():
    """Push waits for the tests, the port in node, and the smoke test, and Deploy for Push, so a commit whose
    checks fail is never deployed; every step names what it waits for."""
    _, steps = _cloudbuild()
    assert list(steps) == ["Build", "Database", "Tests", "Port", "Port in node", "Start the image", "Smoke",
                           "Push", "Deploy"]
    wait = {name: w for name, (w, _) in steps.items()}
    assert all(w is not None for w in wait.values()), wait
    assert all(d == "-" or d in steps for w in wait.values() for d in w), wait
    assert wait["Tests"] == ["Database"] and wait["Port in node"] == ["Port"]
    assert wait["Start the image"] == ["Build"] and wait["Smoke"] == ["Start the image"]
    assert set(wait["Push"]) == {"Tests", "Port in node", "Smoke"}
    assert wait["Deploy"] == ["Push"]
    assert 'grep -q " skipped" build_pytest.log' in steps["Tests"][1]    # a skipped test fails the build


def test_cloudbuild_runs_the_port_on_every_case_it_builds():
    _, steps = _cloudbuild()
    built = re.findall(r"--out build/(\w+)", steps["Port"][1])
    looped = re.search(r"for d in ([\w ]+); do", steps["Port in node"][1]).group(1).split()
    assert len(built) == 5 and built == looped


def test_cloudbuild_tests_in_the_dockerfile_base_image_and_deploys_the_hosted_service():
    """The test and port steps run in the Dockerfile's base image, by the same digest, and the service and
    region it deploys are the ones Firebase Hosting sends isoharness.org to."""
    text, _ = _cloudbuild()
    base = re.search(r"^FROM python:3\.13-slim-trixie@(sha256:[0-9a-f]{64})$", (ROOT / "Dockerfile").read_text(),
                     re.M).group(1)
    assert re.search(r"^  _PYTHON: mirror\.gcr\.io/library/python:3\.13-slim-trixie@(sha256:[0-9a-f]{64})$",
                     text, re.M).group(1) == base, "change _PYTHON in cloudbuild.yaml with the Dockerfile's FROM"
    hosting = (ROOT / "firebase.json").read_text()
    service = re.search(r"^  _SERVICE_NAME: (\S+)$", text, re.M).group(1)
    region = re.search(r"^  _DEPLOY_REGION: (\S+)$", text, re.M).group(1)
    assert f'"serviceId": "{service}"' in hosting and f'"region": "{region}"' in hosting


def test_cloudbuild_leaves_every_bash_dollar_to_bash():
    """Cloud Build reads $NAME and ${NAME} as its own substitutions, so a $ meant for bash is written $$; the
    only single $ left is a substitution such as ${_IMAGE} or ${COMMIT_SHA}."""
    text, _ = _cloudbuild()
    lines = [line for line in text.splitlines() if not line.lstrip().startswith("#")]
    stray = re.findall(r".{0,20}\$(?!\{_?[A-Z][A-Z0-9_]*\}).{0,20}", "\n".join(lines).replace("$$", ""))
    assert not stray, stray
