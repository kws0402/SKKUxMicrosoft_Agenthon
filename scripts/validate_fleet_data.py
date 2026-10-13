#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys

try:
    from scripts.fleet_sync import FleetDataError, validate_payload
except ModuleNotFoundError:  # Supports `python scripts/validate_fleet_data.py`.
    from fleet_sync import FleetDataError, validate_payload


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate generated fleet JSON")
    parser.add_argument("json_file")
    args = parser.parse_args()
    try:
        with open(args.json_file, encoding="utf-8") as handle:
            payload = json.load(handle)
        validate_payload(payload)
    except (FleetDataError, OSError, json.JSONDecodeError) as exc:
        print(f"Fleet data validation failed: {exc}", file=sys.stderr)
        return 1
    print(f"Validated {args.json_file} ({len(payload['vessel'])} vessels)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
