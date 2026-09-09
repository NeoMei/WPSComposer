"""Native Mac Word checkpoint geometry probe; requires --execute."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_references import inventory_without_launch

SOURCES = [Path(__file__), ROOT / "skills/WPSComposer/scripts/writer.py",
           ROOT / "skills/WPSComposer/scripts/msoffice/macos_word_session.py",
           ROOT / "skills/WPSComposer/scripts/msoffice/macos_runtime.py"]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def snapshot(s):
    return s._execute([
        "set nativeRows to {}",
        "set bodyText to content of text object of boundDoc as text",
        "set terminalBound to (end of content of text object of boundDoc) - 1",
        'set end of nativeRows to {"body",bodyText,terminalBound,count characters of text object of boundDoc,count paragraphs of boundDoc,count tables of boundDoc,count fields of boundDoc}',
        "repeat with fi from 1 to count fields of boundDoc",
        "set ownField to field fi of boundDoc",
        'set end of nativeRows to {"field",fi,content of field code of ownField as text,content of result range of ownField as text,start of content of result range of ownField,end of content of result range of ownField}',
        "end repeat",
    ])


def set_text(s, text):
    s._execute([f"set content of text object of boundDoc to {apple_string(text)}", 'set nativeRows to {{"ok"}}'])


def checkpoint(s):
    row = s._execute(['set nativeRows to {{"bounds",(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}'])[0]
    assert row[0] == "bounds" and row[2] == row[1] + 1
    return int(row[1]), int(row[2])


def rollback(s, target):
    target = int(target)
    rows = s._execute([
        "set currentBound to (end of content of text object of boundDoc) - 1",
        f"set targetBound to {target}",
        'if targetBound < 0 or targetBound > currentBound then error "WPSC_INVALID_ROLLBACK_BOUND"',
        "if targetBound < currentBound then",
        "set doomedRange to create range boundDoc start targetBound end currentBound",
        "delete doomedRange",
        "end if",
        'set nativeRows to {{"rollback",targetBound,(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}',
    ])
    assert len(rows) == 1 and rows[0][0] == "rollback" and rows[0][1] == target
    return rows[0]


def append_text(s, text):
    units = len(text.encode("utf-16-le")) // 2
    s._execute([
        "set p to (end of content of text object of boundDoc) - 1",
        "set r to create range boundDoc start p end p",
        f"set content of r to {apple_string(text)}",
        f'if ((end of content of text object of boundDoc) - 1) is not p + {units} then error "WPSC_APPEND_BOUND_MISMATCH"',
        'set nativeRows to {{"ok"}}',
    ])


def run_case(s, name, prefix, mutator):
    set_text(s, prefix)
    before = snapshot(s)
    cp, content_end = checkpoint(s)
    mutator(s)
    changed = snapshot(s)
    rollback_rows = [rollback(s, cp)]
    while rollback_rows[-1][2] > cp and len(rollback_rows) < 3:
        rollback_rows.append(rollback(s, cp))
    assert rollback_rows[-1][2:] == [cp, cp + 1], repr(rollback_rows)
    after = snapshot(s)
    assert after == before
    rollback(s, str(cp))
    assert snapshot(s) == before
    return {"name": name, "checkpoint": cp, "content_end": content_end,
            "before": before, "changed": changed, "after": after, "rollback_acks": rollback_rows,
            "repeated_noop_acknowledged": True}


def run(output: Path):
    output.mkdir(parents=True, exist_ok=False)
    report = {"passed": False, "source_hashes": {str(p.relative_to(ROOT)): sha(p) for p in SOURCES}, "checks": {}, "cases": []}
    session = None
    sentinel_name = None
    starting = inventory_without_launch(output, "starting-inventory")
    report["starting_inventory"] = starting
    try:
        with create_document("writer", engine="msoffice", visible=False) as session:
            sentinel_text = "Checkpoint sentinel 中文😀 " + uuid4().hex
            sentinel_hash = hashlib.sha256((sentinel_text + "\r").encode()).hexdigest()
            sentinel_name = session._execute(["set sentinelDoc to make new document", f"set content of text object of sentinelDoc to {apple_string(sentinel_text)}", 'set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
            report["word_version"] = session._execute(['set nativeRows to {{version as text}}'])[0][0]

            report["cases"].append(run_case(session, "empty", "", lambda s: append_text(s, "A")))
            report["cases"].append(run_case(session, "terminal-paragraph", "prefix", lambda s: append_text(s, " appended")))
            report["cases"].append(run_case(session, "cjk-emoji", "前缀中文😀", lambda s: append_text(s, "追加🧪")))
            report["cases"].append(run_case(session, "appended-paragraph", "base", lambda s: append_text(s, "\rnew paragraph")))

            def partial_table(s):
                s._execute(["activate object boundWindow", "set p to (end of content of text object of boundDoc) - 1", "set insertionRange to create range boundDoc start p end p", "set insertedTable to make new table at boundDoc with properties {text object:insertionRange,number of rows:2,number of columns:2}", 'set content of text object of cell 1 of row 1 of insertedTable to "T1"', 'set content of text object of cell 2 of row 1 of insertedTable to "表😀"', 'set nativeRows to {{"ok"}}'])
            report["cases"].append(run_case(session, "partial-table", "table-prefix\r", partial_table))

            # A no-change failure leaves the checkpoint valid and retryable.
            set_text(session, "unchanged")
            before = snapshot(session); cp, _ = checkpoint(session)
            try:
                session._execute(['error "CONTROLLED_NO_MUTATION_FAILURE"'])
            except Exception:
                pass
            assert snapshot(session) == before
            rollback(session, cp); rollback(session, cp)
            report["checks"]["zero_mutation_failure_and_retry"] = snapshot(session) == before

            # Invalid bounds are rejected before mutation.
            for bad in (-1, cp + 1):
                try: rollback(session, bad)
                except Exception: pass
                else: raise AssertionError("invalid rollback bound accepted")
                assert snapshot(session) == before
            report["checks"]["invalid_bounds_preserve_document"] = True

            # Pre-existing REF survives; an appended REF is inventoried before deletion.
            set_text(session, "target\r")
            session._execute(["set targetRange to create range boundDoc start 0 end 6", 'make new bookmark at boundDoc with properties {name:"wpsc_checkpoint_target",text object:targetRange}', "set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", 'create new field text range r field type field ref field text "wpsc_checkpoint_target \\h" preserve formatting true', "set f to field 1 of boundDoc", 'if (update field f) is false then error "FIELD_UPDATE_FAILED"', 'set nativeRows to {{"ok"}}'])
            stable = snapshot(session); cp, _ = checkpoint(session)
            append_text(session, " tail ")
            session._execute(["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", 'create new field text range r field type field ref field text "wpsc_checkpoint_target \\h" preserve formatting true', "set f to field (count fields of boundDoc) of boundDoc", 'if (update field f) is false then error "FIELD_UPDATE_FAILED"', 'set nativeRows to {{"ok"}}'])
            appended_inventory = snapshot(session)
            assert len([r for r in appended_inventory if r[0] == "field"]) == 2
            rollback(session, cp)
            rolled = snapshot(session)
            assert rolled == stable and len([r for r in rolled if r[0] == "field"]) == 1
            report["field_tracking"] = {"before": stable, "pre_delete_inventory": appended_inventory, "after": rolled,
                "requirement": "Capture appended field identity before native delete; remove only tracking entries whose recorded range lies in the acknowledged reverted extent. Never inspect deleted field objects. Preserve pre-checkpoint REF/index identities."}
            report["checks"]["preexisting_ref_preserved_appended_ref_removed"] = True

            session.save_docx(output / "checkpoint-probe.docx")
            session.export_pdf(output / "checkpoint-probe.pdf")
            # Verify and close only the synthetic sentinel, explicitly discarding it.
            rows = session._execute([f"set sentinelDoc to document {apple_string(sentinel_name)}", "set sentinelText to content of text object of sentinelDoc as text", 'set nativeRows to {{sentinelText,saved of sentinelDoc}}'])
            assert hashlib.sha256(rows[0][0].encode()).hexdigest() == sentinel_hash and rows[0][1] is False
            session._execute([f"set sentinelDoc to document {apple_string(sentinel_name)}", "close sentinelDoc saving no", 'set nativeRows to {{"ok"}}'])
            sentinel_name = None
            shutil.copytree(session.staging_root, output / "native-runtime", dirs_exist_ok=True)
        report["after_owned_close"] = inventory_without_launch(output, "after-owned-close")
        assert report["after_owned_close"] == starting
        with open_document(output / "checkpoint-probe.docx", engine="msoffice", read_only=True, visible=False) as reopened:
            report["reopen"] = snapshot(reopened)
        report["checks"]["saved_native_artifact_reopened"] = True
        import pdfplumber
        with pdfplumber.open(output / "checkpoint-probe.pdf") as pdf:
            report["pdf_text"] = "\n".join(p.extract_text() or "" for p in pdf.pages)
        assert "target" in report["pdf_text"]
        report["checks"]["pdf_native_export_readable"] = True
        report["artifact_hashes"] = {p.name: sha(p) for p in (output / "checkpoint-probe.docx", output / "checkpoint-probe.pdf")}
        report["checks"]["all_geometry_cases_exact"] = all(c["before"] == c["after"] for c in report["cases"])
        report["passed"] = all(report["checks"].values())
    except BaseException as exc:
        report["error"] = {"type": type(exc).__name__, "message": str(exc)}
        (output / "failure.txt").write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root, output / "failed-runtime", dirs_exist_ok=True)
        if sentinel_name:
            report["recovery_targets"] = {"sentinel_name": sentinel_name, "owned_document": getattr(session, "_bound_path", None), "staging_root": str(getattr(session, "staging_root", "")), "note": "Classify from the exception/transport state; an ordinary acknowledged native error retains evidence but is not lock quarantine."}
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2))
    if not report["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--execute", action="store_true"); parser.add_argument("--output", required=True)
    args = parser.parse_args()
    if not args.execute: raise SystemExit("Refusing native Word mutation without --execute")
    run(Path(args.output).resolve())
