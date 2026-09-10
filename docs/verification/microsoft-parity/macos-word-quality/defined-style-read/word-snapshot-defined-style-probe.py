"""Parent-owned native probe for the Word quality snapshot's style subset.

This source is inert unless invoked with ``--execute-native``.  It deliberately
does not call ``macos_word_quality._snapshot_commands()`` or read sections,
headers, footers, paragraph formatting, or character formatting.  A successful
run diagnoses only the style subset; it cannot certify the full quality
snapshot or any mutation/preservation method that depends on that snapshot.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import time
import traceback


STATUS_OK = "STYLE_SUBSET_DIAGNOSED"
STYLE_IDENTIFIERS = (
    "qualityStyleNames",
    "qualityStyleInUseFlags",
    "qualityStyleBuiltinFlags",
    "qualityStyleNamesAfter",
    "qualityStyleOrdinal",
    "qualityStyleCurrentName",
    "qualityStyleIsInUse",
    "qualityStyleIsBuiltin",
    "qualityCurrentStyle",
    "qualityResolvedStyleName",
    "qualityStyleDescription",
    "qualityStyleAutomaticallyUpdates",
    "qualityStyleRows",
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _assert_identifier_safety() -> None:
    """AppleScript identifiers are case-insensitive."""
    folded = [name.casefold() for name in STYLE_IDENTIFIERS]
    if len(folded) != len(set(folded)):
        raise AssertionError("case-insensitive AppleScript identifier collision")


def style_snapshot_lines() -> list[str]:
    """Return the candidate style-only AppleScript body.

    Every exposed style retains exact ordinal/name/in-use/built-in identity.
    Description and automatic-update values are read only for styles Word says
    are in use or non-built-in.  Prior evidence found those are the native
    definition-bearing candidates and that their individual getters preserve
    saved state and do not materialize list templates.  This probe must validate
    the combined and repeated form before production promotion.
    """
    _assert_identifier_safety()
    return [
        "set qualityStyleNames to get name local of every Word style of boundDoc",
        "set qualityStyleInUseFlags to get in use of every Word style of boundDoc",
        "set qualityStyleBuiltinFlags to get built in of every Word style of boundDoc",
        "if (count qualityStyleNames) is not (count qualityStyleInUseFlags) or (count qualityStyleNames) is not (count qualityStyleBuiltinFlags) then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\"",
        "set qualityStyleRows to {{\"style-catalog\",(count qualityStyleNames) as integer}}",
        "repeat with qualityStyleOrdinal from 1 to count qualityStyleNames",
        "set qualityStyleCurrentName to item qualityStyleOrdinal of qualityStyleNames as text",
        "set qualityStyleIsInUse to item qualityStyleOrdinal of qualityStyleInUseFlags",
        "set qualityStyleIsBuiltin to item qualityStyleOrdinal of qualityStyleBuiltinFlags",
        "if class of qualityStyleIsInUse is not boolean or class of qualityStyleIsBuiltin is not boolean then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\"",
        "set end of qualityStyleRows to {\"style-state\",qualityStyleOrdinal as integer,qualityStyleCurrentName,qualityStyleIsInUse,qualityStyleIsBuiltin}",
        "if qualityStyleIsInUse or not qualityStyleIsBuiltin then",
        "set qualityCurrentStyle to Word style qualityStyleOrdinal of boundDoc",
        "set qualityResolvedStyleName to name local of qualityCurrentStyle as text",
        "if not ((current application's NSString's stringWithString:qualityResolvedStyleName)'s isEqualToString:qualityStyleCurrentName) then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\"",
        "set qualityStyleDescription to description of qualityCurrentStyle as text",
        "set qualityStyleAutomaticallyUpdates to automatically update of qualityCurrentStyle",
        "set end of qualityStyleRows to {\"style-definition\",qualityStyleOrdinal as integer,qualityResolvedStyleName,qualityStyleDescription,qualityStyleAutomaticallyUpdates}",
        "end if",
        "end repeat",
        "set qualityStyleNamesAfter to get name local of every Word style of boundDoc",
        "if not ((current application's NSString's stringWithString:(my jsonRows({qualityStyleNamesAfter})))'s isEqualToString:(my jsonRows({qualityStyleNames}))) then error \"WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED\"",
    ]


def fixture_contract() -> dict[str, object]:
    """Facts the parent must establish for separate source-bound runs."""
    return {
        "all_runs": [
            "fresh task-owned DOCX copy opened read-only and hidden",
            "saved=true before the first style read and after every block",
            "body UTF-8 SHA-256 and native End unchanged after every block",
            "paragraph/field/table/bookmark/section counts unchanged",
            "list-template count unchanged; zero must remain zero",
            "private DOCX bytes equal source bytes before close",
            "exact URL-bound close with saving no and independent final inventory []",
            "two complete style snapshots have byte-identical typed JSON",
        ],
        "empty": [
            "minimal body and no custom style",
            "no materialized list template",
            "catalog still returns typed name/in-use/built-in vectors",
        ],
        "nonempty": [
            "applied paragraph and character styles",
            "one modified built-in style, one custom paragraph style, one custom character style",
            "one style linked to numbering and one deliberately unused custom style",
            "OOXML preflight records exact w:styleId/type/name/custom/autoRedefine and numbering relationships",
        ],
        "multisection": [
            "at least three sections",
            "style usage in first and last sections, including a table cell and field result",
            "same style catalog and definition rows repeat exactly without reading any page story",
        ],
        "independent_page_story_matrix_required": [
            "no header/footer parts or references",
            "nonempty primary header and footer with fields",
            "stored but inactive first/even page parts",
            "three sections covering linked-to-previous and unlinked page parts",
            "hidden text and field-code/result ranges",
            "exact story type, section ownership, native bounds, content hash, field identity, saved/body/topology/list-template preservation",
        ],
    }


def _validate_style_rows(rows: object) -> dict[str, int]:
    if not isinstance(rows, list) or not rows or rows[0][0] != "style-catalog":
        raise RuntimeError("malformed style catalog ACK")
    if len(rows[0]) != 2 or type(rows[0][1]) is not int or rows[0][1] <= 0:
        raise RuntimeError("invalid style catalog count")
    expected_count = rows[0][1]
    states: list[list[object]] = []
    definitions: dict[int, list[object]] = {}
    for row in rows[1:]:
        if not isinstance(row, list) or not row:
            raise RuntimeError("malformed style row")
        if row[0] == "style-state":
            if not (
                len(row) == 5
                and type(row[1]) is int
                and isinstance(row[2], str)
                and row[2]
                and type(row[3]) is bool
                and type(row[4]) is bool
            ):
                raise RuntimeError("malformed style-state row")
            states.append(row)
        elif row[0] == "style-definition":
            if not (
                len(row) == 5
                and type(row[1]) is int
                and isinstance(row[2], str)
                and row[2]
                and isinstance(row[3], str)
                and (type(row[4]) is bool or row[4] is None)
            ):
                raise RuntimeError("malformed style-definition row")
            if row[1] in definitions:
                raise RuntimeError("duplicate style-definition ordinal")
            definitions[row[1]] = row
        else:
            raise RuntimeError("unknown style row")
    if len(states) != expected_count:
        raise RuntimeError("style-state coverage mismatch")
    if [row[1] for row in states] != list(range(1, expected_count + 1)):
        raise RuntimeError("style-state ordinals are not complete and ordered")
    names = [row[2] for row in states]
    if len(names) != len(set(names)):
        raise RuntimeError("duplicate exact native style name")
    expected_defined = {row[1] for row in states if row[3] or not row[4]}
    if set(definitions) != expected_defined:
        raise RuntimeError("defined-style coverage mismatch")
    for state in states:
        definition = definitions.get(state[1])
        if definition is not None and definition[2] != state[2]:
            raise RuntimeError("style identity changed during definition read")
    return {"catalog": expected_count, "definitions": len(definitions)}


def _state_lines(label: str) -> list[str]:
    from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands

    return [
        f"log {json.dumps(label)}",
        "set probeSavedBefore to saved of boundDoc",
        *hash_commands("content of text object of boundDoc as text", "probeBodyHash"),
        "set nativeRows to {{\"state\",probeSavedBefore,saved of boundDoc,end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,(count list templates of boundDoc) as integer,(count paragraphs of boundDoc) as integer,(count fields of boundDoc) as integer,(count tables of boundDoc) as integer,(count bookmarks of boundDoc) as integer,(count sections of boundDoc) as integer}}",
    ]


def _validate_state(rows: object) -> list[object]:
    if not (
        isinstance(rows, list)
        and len(rows) == 1
        and isinstance(rows[0], list)
        and len(rows[0]) == 12
        and rows[0][0] == "state"
        and type(rows[0][1]) is bool
        and type(rows[0][2]) is bool
        and type(rows[0][3]) is int
        and isinstance(rows[0][4], str)
        and len(rows[0][4]) == 64
        and rows[0][5] is True
        and all(type(value) is int and value >= 0 for value in rows[0][6:])
    ):
        raise RuntimeError("malformed native state ACK")
    return rows[0]


def _stable(before: list[object], after: list[object]) -> bool:
    return before[1:] == after[1:] and before[1] is True and before[2] is True


def execute_native(source: Path, out: Path) -> int:
    """Execute only when the parent explicitly owns the Word native lease."""
    root = Path.cwd().resolve()
    sys.path.insert(0, str(root))
    from fixtures.microsoft_parity.macos_word_recovery import inventory
    from skills.WPSComposer.scripts.msoffice import macos_word_quality as quality
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession

    source = source.resolve(strict=True)
    out.mkdir(parents=True, exist_ok=False)
    probe_sources = [
        Path(__file__).resolve(),
        root / "skills/WPSComposer/scripts/msoffice/macos_word_quality.py",
        root / "skills/WPSComposer/scripts/msoffice/macos_word_session.py",
        root / "skills/WPSComposer/scripts/msoffice/macos_word_recovery.py",
        root / "skills/WPSComposer/scripts/msoffice/macos_runtime.py",
        root / "skills/WPSComposer/scripts/msoffice/macos_script.py",
    ]
    source_hashes = {str(path.relative_to(root)): _sha256(path) for path in probe_sources}
    report: dict[str, object] = {
        "status": "FAIL",
        "scope": "defined Word style subset only; no page-part/full-quality claim",
        "source": str(source),
        "source_sha256": _sha256(source),
        "probe_source_hashes": source_hashes,
        "fixture_contract": fixture_contract(),
        "steps": [],
    }
    session = None

    def flush() -> None:
        (out / "report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    try:
        report["inventory_before"] = inventory(out, "before")
        flush()
        if report["inventory_before"] != []:
            raise RuntimeError("nonempty Word inventory; probe did not open a document")
        session = MacWordSession.open_document(source, read_only=True, visible=False)
        session._retain_evidence = True
        original_execute = session._execute
        session._execute = lambda lines, **kwargs: original_execute(
            quality._guard(session) + list(lines), **kwargs
        )
        report["owned_path"] = session._bound_path
        report["runtime"] = str(session.staging_root)
        baseline = _validate_state(session._execute(_state_lines("baseline")))
        report["baseline"] = baseline
        flush()

        snapshots: list[list[object]] = []
        for ordinal in (1, 2):
            before = _validate_state(
                session._execute(_state_lines(f"style-snapshot-{ordinal}:before"))
            )
            started = time.monotonic()
            rows = session._execute(style_snapshot_lines() + [
                "set nativeRows to {{\"style-snapshot\",qualityStyleRows}}"
            ])
            seconds = round(time.monotonic() - started, 3)
            if not (
                isinstance(rows, list)
                and len(rows) == 1
                and isinstance(rows[0], list)
                and len(rows[0]) == 2
                and rows[0][0] == "style-snapshot"
            ):
                raise RuntimeError("malformed style snapshot envelope")
            counts = _validate_style_rows(rows[0][1])
            after = _validate_state(
                session._execute(_state_lines(f"style-snapshot-{ordinal}:after"))
            )
            step = {
                "ordinal": ordinal,
                "seconds": seconds,
                "before": before,
                "after": after,
                "stable": _stable(before, after),
                "counts": counts,
                "typed_json_sha256": hashlib.sha256(
                    json.dumps(rows[0][1], ensure_ascii=False, separators=(",", ":")).encode()
                ).hexdigest(),
            }
            report["steps"].append(step)
            flush()
            if not step["stable"]:
                raise RuntimeError("saved/body/topology/list-template state changed")
            snapshots.append(rows[0][1])

        report["checks"] = {
            "repeat_typed_rows_equal": snapshots[0] == snapshots[1],
            "private_bytes_equal_source": _sha256(Path(session._bound_path)) == _sha256(source),
            "baseline_still_exact": _validate_state(session._execute(_state_lines("final"))) == baseline,
            "probe_sources_unchanged": all(
                _sha256(root / relative) == digest
                for relative, digest in source_hashes.items()
            ),
        }
        if not all(report["checks"].values()):
            raise RuntimeError("style-only exact preservation check failed")
        session.close()
        report["owned_closed"] = session._closed
        report["inventory_final"] = inventory(out, "final")
        report["checks"]["inventory_empty"] = report["inventory_final"] == []
        if not report["checks"]["inventory_empty"]:
            raise RuntimeError("Word inventory is not empty after exact close")
        report["status"] = STATUS_OK
    except BaseException as exc:
        report["error"] = {"type": type(exc).__name__, "message": str(exc)}
        (out / "failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
        if session is not None and not session._closed:
            session._retain("Defined-style probe uncertainty; parent reconciliation required")
            report["remaining_owned_path"] = session._bound_path
            report["quarantined"] = session._quarantined
    finally:
        if session is not None and session.staging_root.exists():
            shutil.copytree(session.staging_root, out / "native-runtime", dirs_exist_ok=True)
        flush()
    return 0 if report["status"] == STATUS_OK else 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--emit-plan", action="store_true")
    parser.add_argument("--execute-native", action="store_true")
    parser.add_argument("--source", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    if args.emit_plan:
        print(json.dumps({
            "status": "SOURCE_ONLY",
            "style_snapshot_lines": style_snapshot_lines(),
            "fixture_contract": fixture_contract(),
        }, ensure_ascii=False, indent=2))
        return 0
    if not args.execute_native:
        parser.error("native execution requires explicit --execute-native")
    if args.source is None or args.out is None:
        parser.error("--execute-native requires --source and --out")
    return execute_native(args.source, args.out)


if __name__ == "__main__":
    raise SystemExit(main())
