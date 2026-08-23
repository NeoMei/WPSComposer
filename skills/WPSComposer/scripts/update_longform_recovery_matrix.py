"""Update only the generated recovery-matrix block in the macOS add-in."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence


REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from skills.WPSComposer.scripts.longform.degradation import (  # noqa: E402
    JS_RECOVERY_MATRIX_BEGIN,
    JS_RECOVERY_MATRIX_END,
    render_js_recovery_matrix,
)
from skills.WPSComposer.scripts.longform.privacy import (  # noqa: E402
    JS_PRIVACY_FILTER_BEGIN,
    JS_PRIVACY_FILTER_END,
    render_js_privacy_filter,
)


DEFAULT_ADDIN = (
    REPO_ROOT / "macos/wps-jsapi-probe/addin/writer-longform-v2.js"
)


def _replace_marked_block(
    source: str, begin: str, end: str, rendered: str, *, label: str
) -> str:
    if source.count(begin) != 1:
        raise ValueError(f"{label} begin marker must appear exactly once")
    if source.count(end) != 1:
        raise ValueError(f"{label} end marker must appear exactly once")
    start = source.index(begin)
    stop = source.index(end, start) + len(end)
    return source[:start] + rendered + source[stop:]


def _replace_generated_block(source: str) -> str:
    updated = _replace_marked_block(
        source,
        JS_RECOVERY_MATRIX_BEGIN,
        JS_RECOVERY_MATRIX_END,
        render_js_recovery_matrix(),
        label="recovery matrix",
    )
    return _replace_marked_block(
        updated,
        JS_PRIVACY_FILTER_BEGIN,
        JS_PRIVACY_FILTER_END,
        render_js_privacy_filter(),
        label="privacy filter",
    )


def update_recovery_matrix(path: Path, *, check: bool) -> bool:
    """Return whether *path* drifted; update it unless ``check`` is true."""

    source = path.read_text(encoding="utf-8")
    updated = _replace_generated_block(source)
    drifted = updated != source
    if drifted and not check:
        with path.open("w", encoding="utf-8", newline="\n") as stream:
            stream.write(updated)
    return drifted


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--path", type=Path, default=DEFAULT_ADDIN)
    args = parser.parse_args(argv)
    try:
        drifted = update_recovery_matrix(args.path, check=args.check)
    except (OSError, ValueError) as error:
        parser.error(str(error))
    if args.check and drifted:
        print("long-form recovery matrix is out of date", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
