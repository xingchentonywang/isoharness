# isoharness: the API, ea_irms and LibreOffice in one image.
#
#   docker build -t isoharness .
#   docker run -p 8000:8000 isoharness        # then open http://localhost:8000
#
# The file uses no BuildKit features, because the docker step of a Cloud Build
# trigger builds without BuildKit. Behind a proxy that inspects TLS, pass the
# proxy's CA certificate as a build argument:
#   docker build --build-arg BUILD_CA="$(cat proxy-ca.pem)" -t isoharness .
# apt and pip then trust it alongside the system's certificates while they
# download. The copy they read is deleted in the same step, but Docker keeps
# build arguments in the image's history.
#
# A rebuilt image starts from the same base image and installs the same Python
# libraries (review finding R-26): the base image is pinned by digest, and the
# libraries by version and hash in requirements.lock. LibreOffice is not pinned
# (see its step below).

# Tag python:3.13-slim-trixie, pinned to the digest of the multi-architecture
# index the tag named on 2026-10-15 (Python 3.13.16 on Debian trixie, built
# 2026-10-06), read from Docker Hub and from mirror.gcr.io, which agreed. With
# a digest Docker ignores the tag, which stays for the reader. To move to a
# newer image, read the tag's digest again, for example with
#   docker buildx imagetools inspect python:3.13-slim-trixie
# (its Digest line), and put it here.
FROM python:3.13-slim-trixie@sha256:70729b46c69b4f1e97c4822c1af3df53a1476cf5ddc6c087c0c10bc3a5678c2f

ARG BUILD_CA=""

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# LibreOffice Calc without its GUI computes the workbook's formulas
# (ea_irms.workbook.bake_values). Liberation fonts have Arial's metrics.
# Both come from Debian's archive at whatever version it holds when the image
# is built. They are not pinned, because the archive keeps only the current
# version of a package, so a pinned one would stop the build as soon as Debian
# replaced it, as its security updates and point releases do. The build log
# lists the versions installed (dpkg-query), and the service reports its
# LibreOffice in /api/version (libreoffice_version).
RUN set -eu; \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources; \
    ca=""; \
    if [ -n "$BUILD_CA" ]; then \
        { cat /etc/ssl/certs/ca-certificates.crt; printf '%s\n' "$BUILD_CA"; } > /tmp/build_ca.pem; \
        ca="-o Acquire::https::CAInfo=/tmp/build_ca.pem"; \
    fi; \
    apt-get $ca update; \
    apt-get $ca install -y --no-install-recommends libreoffice-calc-nogui fonts-liberation; \
    dpkg-query -W libreoffice-calc-nogui fonts-liberation; \
    rm -rf /var/lib/apt/lists/* /tmp/build_ca.pem

RUN useradd --create-home --uid 10001 app
WORKDIR /srv/isoharness

# The Python libraries and the libraries they need, each at the version and
# with the hashes in requirements.lock, which is made from requirements.txt by
# the command in its first lines. pip installs only what the lock names
# (--no-deps), only wheels, and only files that match a hash; pip check then
# confirms that the lock left out nothing a library needs.
COPY requirements.lock .
RUN set -eu; \
    if [ -n "$BUILD_CA" ]; then \
        { cat /etc/ssl/certs/ca-certificates.crt; printf '%s\n' "$BUILD_CA"; } > /tmp/build_ca.pem; \
        export PIP_CERT=/tmp/build_ca.pem; \
    fi; \
    pip install --require-hashes --no-deps --only-binary=:all: -r requirements.lock; \
    pip check; \
    rm -f /tmp/build_ca.pem

COPY ea_irms ./ea_irms
COPY isoharness ./isoharness
# The site: the page's files and the preview port it runs (the file CI checks).
COPY web ./web
COPY port/engine.js ./port/engine.js
COPY tests/data/20260915_raw.csv ./warmup/20260915_raw.csv

USER app
ENV HOME=/home/app PORT=8000
# Start LibreOffice once while the image is built, so its first-run setup is not
# part of the first user's wait.
RUN python -m ea_irms warmup/20260915_raw.csv -o /tmp/warmup.xlsx && rm -f /tmp/warmup.xlsx

EXPOSE 8000
# One process: LibreOffice conversions are serialized inside it (isoharness.app). --proxy-headers keeps
# the https of the address in the URLs the app builds. The upload limit does not use the client address
# uvicorn derives from X-Forwarded-For, which with --forwarded-allow-ips='*' is whatever a client wrote at
# its left end: it reads the header itself and counts the address ISOHARNESS_PROXY_HOPS places from the
# right (isoharness 0.4.1, isoharness.app.client_address; GET /api/client shows what it reads).
CMD ["sh", "-c", "exec uvicorn isoharness.app:app --host 0.0.0.0 --port ${PORT} --proxy-headers --forwarded-allow-ips='*'"]
