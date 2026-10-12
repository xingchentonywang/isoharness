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
FROM python:3.13-slim-trixie

ARG BUILD_CA=""

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# LibreOffice Calc without its GUI computes the workbook's formulas
# (ea_irms.workbook.bake_values). Liberation fonts have Arial's metrics.
RUN set -eu; \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources; \
    ca=""; \
    if [ -n "$BUILD_CA" ]; then \
        { cat /etc/ssl/certs/ca-certificates.crt; printf '%s\n' "$BUILD_CA"; } > /tmp/build_ca.pem; \
        ca="-o Acquire::https::CAInfo=/tmp/build_ca.pem"; \
    fi; \
    apt-get $ca update; \
    apt-get $ca install -y --no-install-recommends libreoffice-calc-nogui fonts-liberation; \
    rm -rf /var/lib/apt/lists/* /tmp/build_ca.pem

RUN useradd --create-home --uid 10001 app
WORKDIR /srv/isoharness

COPY requirements.txt .
RUN set -eu; \
    if [ -n "$BUILD_CA" ]; then \
        { cat /etc/ssl/certs/ca-certificates.crt; printf '%s\n' "$BUILD_CA"; } > /tmp/build_ca.pem; \
        export PIP_CERT=/tmp/build_ca.pem; \
    fi; \
    pip install -r requirements.txt; \
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
# One process: LibreOffice conversions are serialized inside it (isoharness.app).
CMD ["sh", "-c", "exec uvicorn isoharness.app:app --host 0.0.0.0 --port ${PORT} --proxy-headers --forwarded-allow-ips='*'"]
