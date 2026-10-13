#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import zipfile
from pathlib import Path


def download_url(public_url: str) -> str:
    parsed = urllib.parse.urlsplit(public_url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(key, value) for key, value in query if key.lower() != "download"]
    query.append(("download", "1"))
    return urllib.parse.urlunsplit(
        (parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), parsed.fragment)
    )


def validate_xlsx(path: Path) -> None:
    if not path.is_file() or path.stat().st_size < 1024:
        raise ValueError("Downloaded file is missing or unexpectedly small")
    with path.open("rb") as handle:
        if handle.read(4) != b"PK\x03\x04":
            raise ValueError("Downloaded response is not an XLSX/ZIP file")
    if not zipfile.is_zipfile(path):
        raise ValueError("Downloaded response is not a valid ZIP archive")
    with zipfile.ZipFile(path) as archive:
        required = {"[Content_Types].xml", "xl/workbook.xml"}
        missing = required - set(archive.namelist())
        if missing:
            raise ValueError(f"Downloaded XLSX is missing entries: {sorted(missing)}")


def download(public_url: str, output_path: str | Path, timeout: int = 90) -> None:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    curl = shutil.which("curl")
    if curl is None:
        raise RuntimeError("curl is required to keep SharePoint's anonymous session cookies")
    temporary: Path | None = None
    cookie_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", prefix="fleet-", suffix=".xlsx", dir=output.parent, delete=False
        ) as handle:
            temporary = Path(handle.name)
        with tempfile.NamedTemporaryFile(
            mode="w", prefix="fleet-cookies-", suffix=".txt", dir=output.parent, delete=False
        ) as handle:
            cookie_path = Path(handle.name)
        subprocess.run(
            [
                curl,
                "--fail",
                "--location",
                "--silent",
                "--show-error",
                "--retry", "3",
                "--retry-all-errors",
                "--max-time", str(timeout),
                "--user-agent", "Mozilla/5.0 (compatible; SKKU-Microsoft-Fleet-Sync/1.0)",
                "--header", "Accept: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*;q=0.8",
                "--cookie-jar", str(cookie_path),
                "--cookie", str(cookie_path),
                "--output", str(temporary),
                download_url(public_url),
            ],
            check=True,
        )
        validate_xlsx(temporary)
        temporary.replace(output)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
        if cookie_path is not None and cookie_path.exists():
            cookie_path.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description="Download a public SharePoint/OneDrive XLSX")
    parser.add_argument("public_url")
    parser.add_argument("output")
    args = parser.parse_args()
    try:
        download(args.public_url, args.output)
    except Exception as exc:  # Network, curl, and archive validation failures share one CLI path.
        print(f"Excel download failed: {exc}", file=sys.stderr)
        return 1
    print(f"Downloaded and validated {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
