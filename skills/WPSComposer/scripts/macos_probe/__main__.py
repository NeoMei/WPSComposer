from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .runner import run_phase0


def main() -> int:
    parser = argparse.ArgumentParser(description="macOS WPS JSAPI Phase 0 probe")
    parser.add_argument("--output-dir", default="build/macos-phase0", type=Path)
    parser.add_argument("--node", default=None)
    parser.add_argument("--timeout", default=90, type=float)
    args = parser.parse_args()

    try:
        report_path = run_phase0(args.output_dir, node=args.node, timeout=args.timeout)
    except Exception as exc:
        print(f"Phase 0 failed: {exc}", file=sys.stderr)
        print(f"Report: {args.output_dir / 'phase0-report.json'}", file=sys.stderr)
        return 1

    import json
    report = json.loads(report_path.read_text())
    if report.get("status") == "passed":
        print(str(report_path.resolve()))
        return 0
    print(f"Phase 0 status: {report.get('status')}", file=sys.stderr)
    print(f"Report: {report_path.resolve()}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
