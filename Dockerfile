# syntax=docker/dockerfile:1
# isoharness: the API, ea_irms and LibreOffice in one image.
#
#   docker build -t isoharness .
#   docker run -p 8000:8000 isoharness        # then open http://localhost:8000/api/docs
#
# Behind a proxy that inspects TLS, pass its CA for the build only (it is not
# kept in the image):  docker build --secret id=build_ca,src=ca.pem .
FROM python:3.13-slim-trixie

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# LibreOffice Calc without its GUI computes the workbook's formulas
# (ea_irms.workbook.bake_values). Liberation fonts have Arial's metrics.
RUN --mount=type=secret,id=build_ca,required=false,mode=0444 \
    set -eu; \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources; \
    ca=""; if [ -s /run/secrets/build_ca ]; then ca="-o Acquire::https::CAInfo=/run/secrets/build_ca"; fi; \
    apt-get $ca update; \
    apt-get $ca install -y --no-install-recommends libreoffice-calc-nogui fonts-liberation; \
    rm -rf /var/lib/apt/lists/*

RUN useradd --create-home --uid 10001 app
WORKDIR /srv/isoharness

COPY requirements.txt .
RUN --mount=type=secret,id=build_ca,required=false,mode=0444 \
    set -eu; \
    if [ -s /run/secrets/build_ca ]; then export PIP_CERT=/run/secrets/build_ca; fi; \
    pip install -r requirements.txt

COPY ea_irms ./ea_irms
COPY isoharness ./isoharness
COPY tests/data/20260915_raw.csv ./warmup/20260915_raw.csv

USER app
ENV HOME=/home/app PORT=8000
# Start LibreOffice once while the image is built, so its first-run setup is not
# part of the first user's wait.
RUN python -m ea_irms warmup/20260915_raw.csv -o /tmp/warmup.xlsx && rm -f /tmp/warmup.xlsx

EXPOSE 8000
# One process: LibreOffice conversions are serialized inside it (isoharness.app).
CMD ["sh", "-c", "exec uvicorn isoharness.app:app --host 0.0.0.0 --port ${PORT} --proxy-headers --forwarded-allow-ips='*'"]
