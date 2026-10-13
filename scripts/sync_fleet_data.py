#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys

try:
    from scripts.fleet_sync import FleetDataError, build_payload, load_json, write_javascript_atomic, write_json_atomic
except ModuleNotFoundError:  # Supports `python scripts/sync_fleet_data.py`.
    from fleet_sync import FleetDataError, build_payload, load_json, write_javascript_atomic, write_json_atomic


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert the fleet Excel workbook to validated JSON")
    parser.add_argument("workbook", help="Downloaded .xlsx file")
    parser.add_argument("output", help="Destination JSON path")
    parser.add_argument("--javascript-output", help="Optional browser-ready window.FLEET_DATA JavaScript path")
    parser.add_argument("--previous", help="Existing JSON used to keep generatedAt stable when data is unchanged")
    parser.add_argument("--source-url", help="Public workbook URL recorded in metadata")
    args = parser.parse_args()

    try:
        previous = load_json(args.previous) if args.previous else None
        kwargs = {"previous_payload": previous}
        if args.source_url:
            kwargs["source_url"] = args.source_url
        payload = build_payload(args.workbook, **kwargs)
        write_json_atomic(payload, args.output)
        if args.javascript_output:
            write_javascript_atomic(payload, args.javascript_output)
    except (FleetDataError, OSError, ValueError) as exc:
        print(f"Fleet data sync failed: {exc}", file=sys.stderr)
        return 1

    print(
        f"Wrote {args.output}: {len(payload['vessel'])} vessels, "
        f"{len(payload['cargo'])} cargo orders, hash {payload['meta']['content_hash'][:12]}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
