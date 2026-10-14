"""Limits on the uploads ea_irms reads (1.4.10).

An .xlsx file is a zip archive, so a small upload can unpack to far more: the 2026-10-14 review built
a 5.0 MB file that unpacks to 121 MB of XML, and reading it as a whole workbook took 2,872 MB of
memory against the service's 1 GiB. ``check_xlsx`` reads only the archive's directory and refuses
such a file before anything opens it; the readers then stream the sheet (openpyxl's read-only mode)
and stop at ``MAX_READ_ROWS`` rows and ``MAX_READ_COLUMNS`` columns.
"""
from __future__ import annotations

import io
import zipfile

# An export of the 400 positions and 3,474 peak records the workbook was timed with unpacks to
# 1.5 MB as an .xlsx file, and a sample sheet of 5,000 rows to 1.6 MB (both measured 2026-10-14).
MAX_UNPACKED_MB = 25.0
MAX_CONFIG_UNPACKED_MB = 5.0       # a lab configuration unpacks to 46 KB
MAX_PARTS = 500                    # files inside the archive
MAX_READ_ROWS = 20000              # rows of an export read; the workbook holds 4,000 peak records
MAX_READ_COLUMNS = 256             # columns of an export read; an Isodat export has a few dozen


class UploadTooLarge(ValueError):
    """The upload is larger, once unpacked or as rows, than ea_irms reads."""


def check_xlsx(source, limit_mb: float = MAX_UNPACKED_MB, what: str = "the file") -> None:
    """Raise UploadTooLarge when an .xlsx file (a path, a file object, or bytes) unpacks to more than
    ``limit_mb`` or holds more than MAX_PARTS files. A file that is not a zip archive passes, and the
    reader then says what is wrong with it in its own words."""
    stream = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    start = stream.tell() if hasattr(stream, "tell") else None
    try:
        with zipfile.ZipFile(stream) as z:
            infos = z.infolist()
    except (zipfile.BadZipFile, OSError, ValueError):
        return
    finally:
        if start is not None:
            stream.seek(start)
    total = sum(i.file_size for i in infos)
    if len(infos) > MAX_PARTS:
        raise UploadTooLarge(f"{what} holds {len(infos)} files inside its archive, more than the {MAX_PARTS} "
                             "this service reads")
    if total > limit_mb * 1024 * 1024:
        raise UploadTooLarge(f"{what} unpacks to {total / 1048576:.0f} MB, more than the {limit_mb:.0f} MB "
                             "this service reads; save it again with only the sheet that is needed")
