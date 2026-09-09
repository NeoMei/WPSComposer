"""Explicit command-line entry point for the dual-platform long-form M0 gate."""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import sys
from typing import Any, Optional

from .contracts import (
    REQUIRED_IDS,
    PlatformEvidence,
    merge_platform_evidence,
    validate_platform_evidence,
    write_canonical_json,
)
from .host_checks import prepare_evidence_directory
from .macos import MacosM0Failed, run_macos_probe
from .windows import WindowsM0Failed, run_windows_probe


def _positive_timeout(value: str) -> float:
    try:
        timeout = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("timeout must be a number") from exc
    if not math.isfinite(timeout) or timeout <= 0:
        raise argparse.ArgumentTypeError("timeout must be positive")
    return timeout


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run or merge the WPSComposer long-form M0 native gate."
    )
    parser.add_argument(
        "--platform",
        required=True,
        choices=("macos", "windows", "verify"),
    )
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--timeout", type=_positive_timeout, default=600.0)
    parser.add_argument("--windows-evidence", type=Path)
    parser.add_argument("--macos-evidence", type=Path)
    return parser


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = build_parser()
    args = parser.parse_args(argv)
    evidence_paths = (args.windows_evidence, args.macos_evidence)
    if args.platform == "verify":
        if any(path is None for path in evidence_paths):
            parser.error(
                "--platform verify requires --windows-evidence and --macos-evidence"
            )
    elif any(path is not None for path in evidence_paths):
        parser.error("evidence path arguments are valid only for --platform verify")
    return args


def _read_evidence(path: Path) -> PlatformEvidence:
    value: Any = json.loads(Path(path).read_text(encoding="utf-8"))
    return validate_platform_evidence(value)


def _platform_passed(evidence: PlatformEvidence) -> bool:
    return all(
        capability.status == "passed"
        for capability in evidence.capabilities
        if capability.id in REQUIRED_IDS
    )


def _run_native(args: argparse.Namespace) -> int:
    runner = run_macos_probe if args.platform == "macos" else run_windows_probe
    try:
        evidence_path = runner(args.output_dir, timeout=args.timeout)
    except (MacosM0Failed, WindowsM0Failed) as error:
        print(str(error), file=sys.stderr)
        print(f"Evidence: {error.evidence_path}", file=sys.stderr)
        return 1
    evidence = _read_evidence(evidence_path)
    if evidence.platform != args.platform:
        raise ValueError("native evidence platform does not match the request")
    print(evidence_path)
    return 0 if _platform_passed(evidence) else 1


def _verify(args: argparse.Namespace) -> int:
    output = prepare_evidence_directory(args.output_dir)
    windows = _read_evidence(args.windows_evidence)
    macos = _read_evidence(args.macos_evidence)
    matrix = merge_platform_evidence(windows, macos)
    matrix_path = output / "matrix-evidence.json"
    write_canonical_json(matrix_path, matrix)
    print(matrix_path)
    return 0 if matrix["decision"] == "go" else 1


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    try:
        if args.platform == "verify":
            return _verify(args)
        return _run_native(args)
    except (OSError, UnicodeError, json.JSONDecodeError, TypeError, ValueError) as error:
        print(f"Invalid M0 arguments or evidence: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
