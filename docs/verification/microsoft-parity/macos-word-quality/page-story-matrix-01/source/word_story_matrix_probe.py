"""Opt-in native diagnostic for Microsoft Word story-range reads.

The source is inert unless invoked with ``--execute-native``. It reads only
story ranges and their fields; it never reads a header/footer ``text object``.
A stable run is diagnostic evidence, not full-quality snapshot acceptance.
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


STATUS_OK = "STORY_MATRIX_DIAGNOSED"
STORY_SPECS = (
    ("primary-header", "primary header story"),
    ("first-header", "first page header story"),
    ("even-header", "even pages header story"),
    ("primary-footer", "primary footer story"),
    ("first-footer", "first page footer story"),
    ("even-footer", "even pages footer story"),
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def fixture_contract() -> dict[str, object]:
    return {
        "story_kinds": [label for label, _ in STORY_SPECS],
        "fixtures": {
            "absent": "no header/footer parts, relationships, or section references",
            "primary": "ordinary and hidden text plus locked PAGE field in primary header/footer",
            "inactive-first-even": "stored first/even header/footer while titlePg and evenAndOddHeaders are disabled",
            "linked-multisection": "three sections; all six explicit in section 1, inherited in section 2, all six replaced in section 3",
        },
        "per_story": [
            "root and each next-story node observed with explicit typed success/error/unresolved rows",
            "class, story type, native Start/End, exact content, field count",
            "field type, code content/Start/End, result content/Start/End, locked flag",
            "no missing/undefined value may be coerced to empty text or count zero",
        ],
        "preservation": [
            "saved/read-only/body End/body hash/list-template/paragraph/field/table/bookmark/section counts stable around every story block",
            "two complete story snapshots have byte-identical typed JSON",
            "private DOCX bytes equal source before close",
            "close exact URL-bound owned copy with saving no only after all stable ACKs",
            "independent final inventory is []",
        ],
    }


def _property_lines(
    label: str, node: str, prop: str, expression: str, *, field: str | None = None
) -> list[str]:
    """Emit a property read whose error/unresolved states remain explicit."""
    if field is None:
        error_row = f'{{"ordinary-property-error","{label}",{node},probePropertyName,probePropertyNumber as integer,probePropertyError as text}}'
        unresolved = lambda reason: f'{{"unresolved-property","{label}",{node},probePropertyName,"{reason}"}}'
        success = f'{{"property-ok","{label}",{node},probePropertyName,probeValueClass,probeSerializableValue}}'
    else:
        error_row = f'{{"field-property-error","{label}",{node},{field},probePropertyName,probePropertyNumber as integer,probePropertyError as text}}'
        unresolved = lambda reason: f'{{"unresolved-field-property","{label}",{node},{field},probePropertyName,"{reason}"}}'
        success = f'{{"field-property-ok","{label}",{node},{field},probePropertyName,probeValueClass,probeSerializableValue}}'
    return [
        f'set probePropertyName to "{prop}"',
        "set probeValue to missing value",
        "try",
        f"set probeValue to get {expression}",
        "set probeValueClass to (class of probeValue) as text",
        "if probeValueClass is \"\" then",
        f"set end of qualityStoryRows to {unresolved('empty-class')}",
        "else if probeValue is missing value then",
        f"set end of qualityStoryRows to {unresolved('native-missing')}",
        "else",
        "if probeValueClass is \"boolean\" or probeValueClass is \"integer\" or probeValueClass is \"real\" or probeValueClass is \"text\" then",
        "set probeSerializableValue to probeValue",
        "else",
        "set probeSerializableValue to probeValue as text",
        "end if",
        f"set end of qualityStoryRows to {success}",
        "end if",
        "on error probePropertyError number probePropertyNumber",
        f"set end of qualityStoryRows to {error_row}",
        "end try",
    ]


def story_snapshot_lines() -> list[str]:
    """Return one complete typed story-chain diagnostic snapshot."""
    lines = ["set qualityStoryRows to {}"]
    for label, story_type in STORY_SPECS:
        lines += [
            f'set end of qualityStoryRows to {{"story-block","{label}"}}',
            "set qualityStoryRange to missing value",
            "try",
            f"set qualityStoryRange to get story range boundDoc story type {story_type}",
            "set qualityRootClass to (class of qualityStoryRange) as text",
            "if qualityRootClass is \"\" then",
            f'set end of qualityStoryRows to {{"unresolved-property","{label}",0,"root","empty-class"}}',
            "else if qualityStoryRange is missing value then",
            f'set end of qualityStoryRows to {{"unresolved-property","{label}",0,"root","native-missing"}}',
            "else",
            "set qualityStoryNode to 1",
            "repeat",
            f'set end of qualityStoryRows to {{"story-node","{label}",qualityStoryNode as integer}}',
        ]
        lines += _property_lines(label, "qualityStoryNode", "class", "(class of qualityStoryRange) as text")
        lines += _property_lines(label, "qualityStoryNode", "story-type", "story type of qualityStoryRange")
        lines += _property_lines(label, "qualityStoryNode", "start", "start of content of qualityStoryRange")
        lines += _property_lines(label, "qualityStoryNode", "end", "end of content of qualityStoryRange")
        lines += _property_lines(label, "qualityStoryNode", "content", "content of qualityStoryRange")
        lines += _property_lines(label, "qualityStoryNode", "section-count", "(count sections of qualityStoryRange) as integer")
        lines += _property_lines(label, "qualityStoryNode", "section-first-index", "index of section 1 of qualityStoryRange")
        lines += [
            "set qualityFieldCountState to \"error\"",
            "set qualityStoryFieldCount to 0",
            "try",
            "set qualityStoryFieldCount to (count fields of qualityStoryRange) as integer",
            "set qualityFieldCountState to \"ok\"",
            "on error qualityFieldCountError number qualityFieldCountNumber",
            f'set end of qualityStoryRows to {{"ordinary-property-error","{label}",qualityStoryNode,"field-count",qualityFieldCountNumber as integer,qualityFieldCountError as text}}',
            "end try",
            "if qualityFieldCountState is \"ok\" then",
            f'set end of qualityStoryRows to {{"field-count","{label}",qualityStoryNode,qualityStoryFieldCount}}',
            "repeat with qualityFieldOrdinal from 1 to qualityStoryFieldCount",
            "set qualityStoryField to field qualityFieldOrdinal of qualityStoryRange",
            f'set end of qualityStoryRows to {{"field-node","{label}",qualityStoryNode,qualityFieldOrdinal as integer}}',
        ]
        lines += _property_lines(label, "qualityStoryNode", "field-type", "field type of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-code-content", "content of field code of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-code-start", "start of content of field code of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-code-end", "end of content of field code of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-result-content", "content of result range of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-result-start", "start of content of result range of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-result-end", "end of content of result range of qualityStoryField", field="qualityFieldOrdinal")
        lines += _property_lines(label, "qualityStoryNode", "field-locked", "locked of qualityStoryField", field="qualityFieldOrdinal")
        lines += [
            "end repeat",
            "end if",
            "set qualityNextStoryRange to missing value",
            "try",
            "set qualityNextStoryRange to get next story range of qualityStoryRange",
            "set qualityNextClass to (class of qualityNextStoryRange) as text",
            "if qualityNextClass is \"\" then",
            f'set end of qualityStoryRows to {{"unresolved-property","{label}",qualityStoryNode,"next-story","empty-class"}}',
            "exit repeat",
            "else if qualityNextStoryRange is missing value then",
            f'set end of qualityStoryRows to {{"chain-end","{label}",qualityStoryNode,"missing-value"}}',
            "exit repeat",
            "else",
            "set qualityStoryRange to qualityNextStoryRange",
            "set qualityStoryNode to qualityStoryNode + 1",
            "if qualityStoryNode > 64 then error \"WPSC_STORY_CHAIN_BOUND_EXCEEDED\"",
            "end if",
            "on error qualityNextError number qualityNextNumber",
            f'set end of qualityStoryRows to {{"ordinary-property-error","{label}",qualityStoryNode,"next-story",qualityNextNumber as integer,qualityNextError as text}}',
            "exit repeat",
            "end try",
            "end repeat",
            "end if",
            "on error qualityRootError number qualityRootNumber",
            f'set end of qualityStoryRows to {{"ordinary-property-error","{label}",0,"root",qualityRootNumber as integer,qualityRootError as text}}',
            "end try",
        ]
    return lines


def validate_story_rows(rows: object, *, expected_label: str | None = None) -> dict[str, int]:
    if not isinstance(rows, list) or not rows:
        raise RuntimeError("malformed story rows")
    labels = {label for label, _ in STORY_SPECS}
    expected_sequence = [expected_label] if expected_label else [label for label, _ in STORY_SPECS]
    blocks = [row[1] for row in rows if isinstance(row, list) and len(row) == 2 and row[0] == "story-block"]
    if blocks != expected_sequence:
        raise RuntimeError("story-block coverage/order mismatch")
    nodes = fields = observations = ordinary_errors = 0
    seen_nodes: dict[str, list[int]] = {label: [] for label in expected_sequence}
    allowed_properties = {
        "root", "class", "story-type", "start", "end", "content", "section-count",
        "section-first-index", "field-count",
        "field-type", "field-code-content", "field-code-start", "field-code-end",
        "field-result-content", "field-result-start", "field-result-end", "field-locked", "next-story",
    }
    expected_field_properties = {
        "field-type", "field-code-content", "field-code-start", "field-code-end",
        "field-result-content", "field-result-start", "field-result-end", "field-locked",
    }
    declared_field_counts: dict[tuple[str, int], int] = {}
    expected_story_properties = {"class", "story-type", "start", "end", "content", "section-count", "section-first-index"}
    story_properties: dict[tuple[str, int], set[str]] = {}
    field_count_errors: set[tuple[str, int]] = set()
    terminal_next: set[tuple[str, int]] = set()
    root_observations: dict[str, int] = {label: 0 for label in expected_sequence}
    field_nodes: set[tuple[str, int, int]] = set()
    field_properties: dict[tuple[str, int, int], set[str]] = {}
    for row in rows:
        if not isinstance(row, list) or not row or not isinstance(row[0], str):
            raise RuntimeError("malformed story row")
        tag = row[0]
        if tag == "story-block":
            continue
        if tag == "story-node":
            if len(row) != 3 or row[1] not in labels or type(row[2]) is not int or row[2] < 1:
                raise RuntimeError("malformed story-node")
            seen_nodes[row[1]].append(row[2])
            story_properties[(row[1], row[2])] = set()
            nodes += 1
            continue
        if len(row) < 4 or row[1] not in labels or type(row[2]) is not int or row[2] < 0:
            raise RuntimeError("malformed story observation identity")
        if tag == "field-node":
            if len(row) != 4 or type(row[3]) is not int or row[3] < 1:
                raise RuntimeError("malformed field-node")
            key = (row[1], row[2], row[3])
            if key in field_nodes:
                raise RuntimeError("duplicate field-node")
            field_nodes.add(key)
            field_properties[key] = set()
        elif tag == "field-count":
            if len(row) != 4 or type(row[3]) is not int or row[3] < 0:
                raise RuntimeError("malformed field-count")
            fields += row[3]
            declared_field_counts[(row[1], row[2])] = row[3]
            observations += 1
        elif tag == "field-property-ok":
            if (len(row) != 7 or type(row[3]) is not int or row[3] < 1
                    or row[4] not in expected_field_properties or not isinstance(row[5], str) or row[5] == ""):
                raise RuntimeError("malformed field-property-ok")
            key = (row[1], row[2], row[3])
            if key not in field_nodes:
                raise RuntimeError("field property precedes field-node")
            field_properties[key].add(row[4])
            observations += 1
        elif tag == "unresolved-field-property":
            if (len(row) != 6 or type(row[3]) is not int or row[3] < 1
                    or row[4] not in expected_field_properties or row[5] not in ("empty-class", "native-missing")):
                raise RuntimeError("malformed unresolved-field-property")
            key = (row[1], row[2], row[3])
            if key not in field_nodes:
                raise RuntimeError("field property precedes field-node")
            field_properties[key].add(row[4])
            observations += 1
        elif tag == "field-property-error":
            if (len(row) != 7 or type(row[3]) is not int or row[3] < 1
                    or row[4] not in expected_field_properties or type(row[5]) is not int or not isinstance(row[6], str)):
                raise RuntimeError("malformed field-property-error")
            key = (row[1], row[2], row[3])
            if key not in field_nodes:
                raise RuntimeError("field property precedes field-node")
            field_properties[key].add(row[4])
            observations += 1
            ordinary_errors += 1
        elif tag == "property-ok":
            if len(row) != 6 or row[3] not in allowed_properties or not isinstance(row[4], str):
                raise RuntimeError("malformed property-ok")
            if row[4] == "":
                raise RuntimeError("empty class cannot be property-ok")
            if row[3] in expected_story_properties:
                story_properties[(row[1], row[2])].add(row[3])
            observations += 1
        elif tag == "unresolved-property":
            if len(row) != 5 or row[3] not in allowed_properties or row[4] not in ("empty-class", "native-missing"):
                raise RuntimeError("malformed unresolved-property")
            if row[2] == 0 and row[3] == "root":
                root_observations[row[1]] += 1
            elif row[3] in expected_story_properties:
                story_properties[(row[1], row[2])].add(row[3])
            elif row[3] == "next-story":
                terminal_next.add((row[1], row[2]))
            observations += 1
        elif tag == "ordinary-property-error":
            if len(row) != 6 or row[3] not in allowed_properties or type(row[4]) is not int or not isinstance(row[5], str):
                raise RuntimeError("malformed ordinary-property-error")
            if row[2] == 0 and row[3] == "root":
                root_observations[row[1]] += 1
            elif row[3] in expected_story_properties:
                story_properties[(row[1], row[2])].add(row[3])
            elif row[3] == "field-count":
                field_count_errors.add((row[1], row[2]))
            elif row[3] == "next-story":
                terminal_next.add((row[1], row[2]))
            observations += 1
            ordinary_errors += 1
        elif tag == "chain-end":
            if len(row) != 4 or row[3] != "missing-value":
                raise RuntimeError("malformed chain-end")
            terminal_next.add((row[1], row[2]))
            observations += 1
        else:
            raise RuntimeError(f"unknown story row: {tag}")
    for label, ordinals in seen_nodes.items():
        if ordinals and ordinals != list(range(1, len(ordinals) + 1)):
            raise RuntimeError(f"noncontiguous story chain: {label}")
        if not ordinals:
            if root_observations[label] != 1:
                raise RuntimeError(f"root observation coverage mismatch: {label}")
            continue
        if root_observations[label]:
            raise RuntimeError(f"concrete and unresolved root mixed: {label}")
        for ordinal in ordinals:
            owner = (label, ordinal)
            if story_properties.get(owner) != expected_story_properties:
                raise RuntimeError("story property coverage mismatch")
            if owner not in declared_field_counts and owner not in field_count_errors:
                raise RuntimeError("field-count coverage mismatch")
        if (label, ordinals[-1]) not in terminal_next:
            raise RuntimeError("story chain termination missing")
    for owner, declared_count in declared_field_counts.items():
        actual = sorted(key[2] for key in field_nodes if key[:2] == owner)
        if actual != list(range(1, declared_count + 1)):
            raise RuntimeError("field-node coverage mismatch")
    if any(properties != expected_field_properties for properties in field_properties.values()):
        raise RuntimeError("field property coverage mismatch")
    return {"nodes": nodes, "fields": fields, "observations": observations,
            "ordinary_errors": ordinary_errors}


def _state_lines(label: str) -> list[str]:
    from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands

    return [
        f"log {json.dumps(label)}",
        "set probeSavedBefore to saved of boundDoc",
        *hash_commands("content of text object of boundDoc as text", "probeBodyHash"),
        "set nativeRows to {{\"state\",probeSavedBefore,saved of boundDoc,end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,(count list templates of boundDoc) as integer,(count paragraphs of boundDoc) as integer,(count fields of boundDoc) as integer,(count tables of boundDoc) as integer,(count bookmarks of boundDoc) as integer,(count sections of boundDoc) as integer}}",
    ]


def _validate_state(rows: object) -> list[object]:
    if not (isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list)
            and len(rows[0]) == 12 and rows[0][0] == "state"
            and type(rows[0][1]) is bool and type(rows[0][2]) is bool
            and type(rows[0][3]) is int and isinstance(rows[0][4], str) and len(rows[0][4]) == 64
            and rows[0][5] is True and all(type(value) is int and value >= 0 for value in rows[0][6:])):
        raise RuntimeError("malformed native state ACK")
    return rows[0]


def _stable(before: list[object], after: list[object]) -> bool:
    return before[1:] == after[1:] and before[1] is True and before[2] is True


def execute_native(source: Path, out: Path) -> int:
    root = Path.cwd().resolve()
    sys.path.insert(0, str(root))
    from fixtures.microsoft_parity.macos_word_recovery import inventory
    from skills.WPSComposer.scripts.msoffice import macos_word_quality as quality
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession

    source = source.resolve(strict=True)
    out.mkdir(parents=True, exist_ok=False)
    source_files = [Path(__file__).resolve(), root / "skills/WPSComposer/scripts/msoffice/macos_word_quality.py",
                    root / "skills/WPSComposer/scripts/msoffice/macos_word_session.py",
                    root / "skills/WPSComposer/scripts/msoffice/macos_word_recovery.py",
                    root / "skills/WPSComposer/scripts/msoffice/macos_runtime.py",
                    root / "skills/WPSComposer/scripts/msoffice/macos_script.py"]
    source_hashes = {str(path.relative_to(root)): _sha256(path) for path in source_files}
    report: dict[str, object] = {
        "status": "FAIL", "scope": "story-range diagnostic only; no full-quality or parity claim",
        "source": str(source), "source_sha256": _sha256(source),
        "probe_source_hashes": source_hashes, "fixture_contract": fixture_contract(), "steps": [],
    }
    session = None

    def flush() -> None:
        (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    try:
        report["inventory_before"] = inventory(out, "before")
        flush()
        if report["inventory_before"] != []:
            raise RuntimeError("nonempty Word inventory; probe did not open a document")
        session = MacWordSession.open_document(source, read_only=True, visible=False)
        session._retain_evidence = True
        original_execute = session._execute
        session._execute = lambda lines, **kwargs: original_execute(quality._guard(session) + list(lines), **kwargs)
        report["owned_path"] = session._bound_path
        report["runtime"] = str(session.staging_root)
        baseline = _validate_state(session._execute(_state_lines("baseline")))
        report["baseline"] = baseline
        snapshots = []
        for snapshot_ordinal in (1, 2):
            snapshot_rows = []
            for label, _ in STORY_SPECS:
                before = _validate_state(session._execute(_state_lines(f"snapshot-{snapshot_ordinal}:{label}:before")))
                # Each block is independently state-checked, while filtering the complete source keeps one implementation.
                all_lines = story_snapshot_lines()
                start = all_lines.index(f'set end of qualityStoryRows to {{"story-block","{label}"}}')
                next_starts = [i for i in range(start + 1, len(all_lines)) if all_lines[i].startswith('set end of qualityStoryRows to {"story-block"')]
                end = next_starts[0] if next_starts else len(all_lines)
                block_lines = ["set qualityStoryRows to {}"] + all_lines[start:end]
                started = time.monotonic()
                envelope = session._execute(block_lines + ['set nativeRows to {{"story-snapshot",qualityStoryRows}}'])
                seconds = round(time.monotonic() - started, 3)
                if not (isinstance(envelope, list) and len(envelope) == 1 and isinstance(envelope[0], list)
                        and len(envelope[0]) == 2 and envelope[0][0] == "story-snapshot"):
                    raise RuntimeError("malformed story snapshot envelope")
                rows = envelope[0][1]
                summary = validate_story_rows(rows, expected_label=label)
                after = _validate_state(session._execute(_state_lines(f"snapshot-{snapshot_ordinal}:{label}:after")))
                step = {"snapshot": snapshot_ordinal, "story": label, "seconds": seconds,
                        "before": before, "after": after, "stable": _stable(before, after),
                        "summary": summary, "rows": rows,
                        "typed_json_sha256": hashlib.sha256(json.dumps(
                            rows, ensure_ascii=False, separators=(",", ":")
                        ).encode()).hexdigest()}
                report["steps"].append(step)
                flush()
                if not step["stable"]:
                    raise RuntimeError(f"state changed during {label}; no further AppleEvents")
                snapshot_rows.extend(rows)
            snapshots.append(snapshot_rows)
        report["checks"] = {
            "repeat_typed_rows_equal": snapshots[0] == snapshots[1],
            "private_bytes_equal_source": _sha256(Path(session._bound_path)) == _sha256(source),
            "baseline_still_exact": _validate_state(session._execute(_state_lines("final"))) == baseline,
            "probe_sources_unchanged": all(_sha256(root / rel) == digest for rel, digest in source_hashes.items()),
        }
        if not all(report["checks"].values()):
            raise RuntimeError("story diagnostic exact preservation check failed")
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
            session._retain("Story-range diagnostic uncertainty; parent reconciliation required")
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
        print(json.dumps({"status": "SOURCE_ONLY", "story_snapshot_lines": story_snapshot_lines(),
                          "fixture_contract": fixture_contract()}, ensure_ascii=False, indent=2))
        return 0
    if not args.execute_native:
        parser.error("native execution requires explicit --execute-native")
    if args.source is None or args.out is None:
        parser.error("--execute-native requires --source and --out")
    return execute_native(args.source, args.out)


if __name__ == "__main__":
    raise SystemExit(main())
