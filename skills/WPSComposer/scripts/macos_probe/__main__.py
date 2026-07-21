from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

from .runner import Phase0Failed, run_phase0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the macOS WPS JSAPI Phase 0 feasibility probe."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("build/macos-phase0"),
    )
    parser.add_argument("--node")
    parser.add_argument("--timeout", type=float, default=90)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        report_path = run_phase0(args.output_dir, args.node, args.timeout)
    except Phase0Failed as exc:
        print(f"Mac WPS Phase 0 failed: {exc}", file=sys.stderr)
        print(f"Report: {exc.report_path}", file=sys.stderr)
        return 1
    print(report_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
