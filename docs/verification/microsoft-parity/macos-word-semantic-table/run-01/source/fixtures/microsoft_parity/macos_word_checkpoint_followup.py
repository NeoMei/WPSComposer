"""Table/REF-only native checkpoint follow-up; never runs without --execute."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer import create_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_references import inventory_without_launch

SOURCES = [Path(__file__), ROOT / "skills/WPSComposer/scripts/writer.py",
           ROOT / "skills/WPSComposer/scripts/msoffice/macos_word_session.py",
           ROOT / "skills/WPSComposer/scripts/msoffice/macos_runtime.py"]


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class Journal:
    def __init__(self, path: Path, mode: str):
        self.path = path
        self.data = {"passed": False, "mode": mode, "steps": [],
                     "source_hashes": {str(p.relative_to(ROOT)): digest(p) for p in SOURCES}}
        self.flush()

    def flush(self):
        pending = self.path.with_suffix(".pending")
        pending.write_text(json.dumps(self.data, ensure_ascii=False, indent=2))
        pending.replace(self.path)

    def add(self, label: str, rows):
        self.data["steps"].append({"label": label, "rows": rows})
        self.flush()  # Persist every acknowledged native result before assertions.
        return rows


def geometry(session):
    return session._execute([
        "set nativeRows to {}", "set bodyRange to text object of boundDoc",
        'set end of nativeRows to {"document",content of bodyRange as text,(end of content of bodyRange) - 1,end of content of bodyRange,count characters of bodyRange,count paragraphs of boundDoc,count tables of boundDoc,count fields of boundDoc}',
        "repeat with pi from 1 to count paragraphs of boundDoc", "set ownRange to text object of paragraph pi of boundDoc",
        'set end of nativeRows to {"paragraph",pi,content of ownRange as text,start of content of ownRange,end of content of ownRange}', "end repeat",
        "repeat with ti from 1 to count tables of boundDoc", "set ownTable to table ti of boundDoc",
        'set end of nativeRows to {"table",ti,start of content of text object of ownTable,end of content of text object of ownTable,count rows of ownTable,count columns of ownTable}', "end repeat",
        "repeat with bi from 1 to count bookmarks of boundDoc", "set ownBookmark to bookmark bi of boundDoc",
        'set end of nativeRows to {"bookmark",bi,name of ownBookmark as text,content of text object of ownBookmark as text,start of content of text object of ownBookmark,end of content of text object of ownBookmark}', "end repeat",
        "repeat with fi from 1 to count fields of boundDoc", "set ownField to field fi of boundDoc",
        'set end of nativeRows to {"field",fi,content of field code of ownField as text,content of result range of ownField as text,start of content of result range of ownField,end of content of result range of ownField}', "end repeat",
    ])


def delete_from(session, target: int):
    return session._execute([
        "set currentBound to (end of content of text object of boundDoc) - 1",
        f"set targetBound to {int(target)}",
        'if targetBound < 0 or targetBound > currentBound then error "WPSC_INVALID_ROLLBACK_BOUND"',
        "if targetBound < currentBound then delete (create range boundDoc start targetBound end currentBound)",
        'set nativeRows to {{"delete-ack",targetBound,(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}',
    ])


def table_only(session, journal: Journal):
    prefix = "table-prefix\r"
    journal.add("set-prefix-ack", session._execute([f"set content of text object of boundDoc to {apple_string(prefix)}", 'set nativeRows to {{"set-prefix-ack"}}']))
    before = journal.add("before", geometry(session))
    checkpoint = int(before[0][2])
    journal.data["checkpoint"] = checkpoint; journal.data["prefix"] = prefix; journal.flush()
    journal.add("table-insert-ack", session._execute([
        "activate object boundWindow", "set insertionPoint to (end of content of text object of boundDoc) - 1",
        "set insertionRange to create range boundDoc start insertionPoint end insertionPoint",
        "set insertedTable to make new table at boundDoc with properties {text object:insertionRange,number of rows:2,number of columns:2}",
        'set content of text object of cell 1 of row 1 of insertedTable to "T1"',
        'set content of text object of cell 2 of row 1 of insertedTable to "表😀"',
        'set nativeRows to {{"table-insert-ack",start of content of text object of insertedTable,end of content of text object of insertedTable}}']))
    journal.add("after-table", geometry(session))
    journal.add("first-delete-ack", delete_from(session, checkpoint))
    first = journal.add("after-first-delete", geometry(session))
    first_bound = int(first[0][2])
    journal.data["observations"] = {"before_document": before[0], "after_first_document": first[0], "checkpoint": checkpoint,
                                    "second_delete_legal": first_bound >= checkpoint}
    journal.flush()
    # Do not issue an invalid range merely to obtain a failure. If Word moved
    # left of the public checkpoint, preserve that result as the contract miss.
    if first_bound < checkpoint:
        journal.add("second-delete-skipped", [["contract-miss", "current-bound-below-checkpoint", first_bound, checkpoint]])
        second = first
    else:
        journal.add("second-delete-ack", delete_from(session, checkpoint))
        second = journal.add("after-second-delete", geometry(session))
    journal.data["observations"]["after_second_document"] = second[0]
    journal.flush()
    return before == second


def fields_only(session, journal: Journal):
    target = "target"
    target_units = len(target.encode("utf-16-le")) // 2
    journal.add("seed-text-ack", session._execute([f'set content of text object of boundDoc to {apple_string(target)}', 'set nativeRows to {{"seed-text-ack"}}']))
    bookmark_ack = journal.add("seed-bookmark-ack", session._execute([
        f"set targetRange to create range boundDoc start 0 end {target_units}",
        'make new bookmark at boundDoc with properties {name:"wpsc_checkpoint_preexisting",text object:targetRange}',
        'set ownBookmark to bookmark "wpsc_checkpoint_preexisting" of boundDoc',
        'set nativeRows to {{"seed-bookmark-ack",name of ownBookmark as text,content of text object of ownBookmark as text,start of content of text object of ownBookmark,end of content of text object of ownBookmark}}']))
    assert bookmark_ack == [["seed-bookmark-ack", "wpsc_checkpoint_preexisting", target, 0, target_units]]
    journal.add("seed-field-ack", session._execute([
        "set insertionPoint to (end of content of text object of boundDoc) - 1", "set ownRange to create range boundDoc start insertionPoint end insertionPoint",
        f'create new field text range ownRange field type field ref field text {apple_string("wpsc_checkpoint_preexisting " + chr(92) + "h")} preserve formatting true',
        "set ownField to field (count fields of boundDoc) of boundDoc", 'if (update field ownField) is false then error "FIELD_UPDATE_FAILED"',
        'set nativeRows to {{"seed-field-ack",count fields of boundDoc,content of field code of ownField as text,start of content of result range of ownField,end of content of result range of ownField}}']))
    before = journal.add("before", geometry(session)); checkpoint = int(before[0][2])
    journal.data["checkpoint"] = checkpoint; journal.flush()
    journal.add("append-prefix-ack", session._execute(["set insertionPoint to (end of content of text object of boundDoc) - 1", "set ownRange to create range boundDoc start insertionPoint end insertionPoint", 'set content of ownRange to " tail "', 'set nativeRows to {{"append-prefix-ack"}}']))
    journal.add("append-ref-ack", session._execute([
        "set insertionPoint to (end of content of text object of boundDoc) - 1", "set ownRange to create range boundDoc start insertionPoint end insertionPoint",
        f'create new field text range ownRange field type field ref field text {apple_string("wpsc_checkpoint_preexisting " + chr(92) + "h")} preserve formatting true',
        "set ownField to field (count fields of boundDoc) of boundDoc", 'if (update field ownField) is false then error "FIELD_UPDATE_FAILED"',
        'set nativeRows to {{"append-ref-ack",count fields of boundDoc,content of field code of ownField as text,start of content of result range of ownField,end of content of result range of ownField}}']))
    appended = journal.add("before-delete-field-inventory", geometry(session))
    journal.add("delete-ack", delete_from(session, checkpoint))
    after = journal.add("after-delete", geometry(session))
    journal.data["tracking_contract"] = {"preexisting_bookmark": [r for r in before if r[0] == "bookmark"], "preexisting_fields": [r for r in before if r[0] == "field"], "all_fields_before_delete": [r for r in appended if r[0] == "field"], "fields_after_delete": [r for r in after if r[0] == "field"]}
    journal.flush()
    bookmark = journal.data["tracking_contract"]["preexisting_bookmark"]
    assert len(bookmark) == 1 and bookmark[0][2] == "wpsc_checkpoint_preexisting" and bookmark[0][3:] == [target, 0, target_units]
    assert len(journal.data["tracking_contract"]["preexisting_fields"]) == 1
    assert len(journal.data["tracking_contract"]["all_fields_before_delete"]) == 2
    return before == after


def delete_new_objects(session, checkpoint: int, kind: str):
    collection = "tables" if kind == "table" else "fields"
    object_name = "table" if kind == "table" else "field"
    range_name = "text object" if kind == "table" else "result range"
    return session._execute([
        "set deletedRows to {}",
        f"repeat with oi from (count {collection} of boundDoc) to 1 by -1",
        f"set ownObject to {object_name} oi of boundDoc",
        f"set ownStart to start of content of {range_name} of ownObject",
        f"set ownEnd to end of content of {range_name} of ownObject",
        f"if ownStart >= {int(checkpoint)} then",
        f'set end of deletedRows to {{"{kind}",oi,ownStart,ownEnd}}',
        "delete ownObject", "end if", "end repeat",
        f'set nativeRows to {{{{"object-delete-ack",deletedRows,(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc,count {collection} of boundDoc}}}}',
    ])


def table_object_delete_only(session, journal: Journal):
    prefix = "table-prefix\r"
    journal.add("set-prefix-ack", session._execute([f"set content of text object of boundDoc to {apple_string(prefix)}", 'set nativeRows to {{"set-prefix-ack"}}']))
    before = journal.add("before", geometry(session)); checkpoint = int(before[0][2])
    journal.data["checkpoint"] = checkpoint; journal.data["pre_topology"] = {"tables": [r for r in before if r[0] == "table"]}; journal.flush()
    journal.add("table-insert-ack", session._execute(["activate object boundWindow", "set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", "set t to make new table at boundDoc with properties {text object:r,number of rows:2,number of columns:2}", 'set content of text object of cell 1 of row 1 of t to "T1"', 'set content of text object of cell 2 of row 1 of t to "表😀"', 'set nativeRows to {{"table-insert-ack",start of content of text object of t,end of content of text object of t}}']))
    journal.add("after-table", geometry(session))
    journal.add("object-delete-ack", delete_new_objects(session, checkpoint, "table"))
    journal.add("after-object-delete", geometry(session))
    journal.add("residual-delete-ack", delete_from(session, checkpoint))
    after = journal.add("after-residual-delete", geometry(session))
    # Exact repeated rollback no-op after the successful sequence.
    journal.add("repeat-object-delete-ack", delete_new_objects(session, checkpoint, "table"))
    journal.add("repeat-range-delete-ack", delete_from(session, checkpoint))
    repeated = journal.add("after-repeat", geometry(session))
    journal.data["observations"] = {"before": before, "after": after, "repeated": repeated}; journal.flush()
    return before == after == repeated


def field_object_delete_only(session, journal: Journal):
    target = "target"; units = len(target.encode("utf-16-le")) // 2
    journal.add("seed-text-ack", session._execute([f"set content of text object of boundDoc to {apple_string(target)}", 'set nativeRows to {{"seed-text-ack"}}']))
    bookmark_ack = journal.add("seed-bookmark-ack", session._execute([f"set r to create range boundDoc start 0 end {units}", 'make new bookmark at boundDoc with properties {name:"wpsc_checkpoint_preexisting",text object:r}', 'set b to bookmark "wpsc_checkpoint_preexisting" of boundDoc', 'set nativeRows to {{"seed-bookmark-ack",name of b as text,content of text object of b as text,start of content of text object of b,end of content of text object of b}}']))
    assert bookmark_ack == [["seed-bookmark-ack", "wpsc_checkpoint_preexisting", target, 0, units]]
    code = apple_string("wpsc_checkpoint_preexisting " + chr(92) + "h")
    journal.add("seed-field-ack", session._execute(["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", f"create new field text range r field type field ref field text {code} preserve formatting true", "set f to field (count fields of boundDoc) of boundDoc", 'if (update field f) is false then error "FIELD_UPDATE_FAILED"', 'set nativeRows to {{"seed-field-ack",content of field code of f as text,start of content of result range of f,end of content of result range of f}}']))
    before = journal.add("before", geometry(session)); checkpoint = int(before[0][2])
    journal.data["checkpoint"] = checkpoint; journal.data["pre_topology"] = {"bookmarks": [r for r in before if r[0] == "bookmark"], "fields": [r for r in before if r[0] == "field"]}; journal.flush()
    journal.add("append-text-ack", session._execute(["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", 'set content of r to " tail "', 'set nativeRows to {{"append-text-ack"}}']))
    journal.add("append-field-ack", session._execute(["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", f"create new field text range r field type field ref field text {code} preserve formatting true", "set f to field (count fields of boundDoc) of boundDoc", 'if (update field f) is false then error "FIELD_UPDATE_FAILED"', 'set nativeRows to {{"append-field-ack",content of field code of f as text,start of content of result range of f,end of content of result range of f}}']))
    journal.add("before-delete", geometry(session))
    journal.add("object-delete-ack", delete_new_objects(session, checkpoint, "field"))
    journal.add("after-object-delete", geometry(session))
    journal.add("residual-delete-ack", delete_from(session, checkpoint))
    after = journal.add("after-residual-delete", geometry(session))
    journal.add("repeat-object-delete-ack", delete_new_objects(session, checkpoint, "field"))
    journal.add("repeat-range-delete-ack", delete_from(session, checkpoint))
    repeated = journal.add("after-repeat", geometry(session))
    journal.data["observations"] = {"before": before, "after": after, "repeated": repeated}; journal.flush()
    return before == after == repeated


def field_object_content_clear(session, journal: Journal):
    target = "target"; units = len(target.encode("utf-16-le")) // 2
    journal.add("seed-text-ack", session._execute([f"set content of text object of boundDoc to {apple_string(target)}", 'set nativeRows to {{"seed-text-ack"}}']))
    bookmark_ack = journal.add("seed-bookmark-ack", session._execute([f"set r to create range boundDoc start 0 end {units}", 'make new bookmark at boundDoc with properties {name:"wpsc_checkpoint_preexisting",text object:r}', 'set b to bookmark "wpsc_checkpoint_preexisting" of boundDoc', 'set nativeRows to {{"seed-bookmark-ack",name of b as text,content of text object of b as text,start of content of text object of b,end of content of text object of b}}']))
    assert bookmark_ack == [["seed-bookmark-ack", "wpsc_checkpoint_preexisting", target, 0, units]]
    code = apple_string("wpsc_checkpoint_preexisting " + chr(92) + "h")
    insert_field = ["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", f"create new field text range r field type field ref field text {code} preserve formatting true", "set f to field (count fields of boundDoc) of boundDoc", 'if (update field f) is false then error "FIELD_UPDATE_FAILED"']
    journal.add("seed-field-ack", session._execute(insert_field + ['set nativeRows to {{"seed-field-ack",content of field code of f as text,start of content of result range of f,end of content of result range of f}}']))
    before = journal.add("before", geometry(session)); checkpoint = int(before[0][2]); journal.data["checkpoint"] = checkpoint; journal.flush()
    journal.add("append-text-ack", session._execute(["set p to (end of content of text object of boundDoc) - 1", "set r to create range boundDoc start p end p", 'set content of r to " tail "', 'set nativeRows to {{"append-text-ack"}}']))
    journal.add("append-field-ack", session._execute(insert_field + ['set nativeRows to {{"append-field-ack",content of field code of f as text,start of content of result range of f,end of content of result range of f}}']))
    journal.add("before-delete", geometry(session))
    journal.add("object-delete-ack", delete_new_objects(session, checkpoint, "field"))
    journal.add("after-object-delete", geometry(session))
    residual = journal.add("residual-range-readback", session._execute(["set currentBound to (end of content of text object of boundDoc) - 1", f"set rollbackRange to create range boundDoc start {checkpoint} end currentBound", 'set nativeRows to {{"residual-range",content of rollbackRange as text,start of content of rollbackRange,end of content of rollbackRange,currentBound}}']))
    journal.add("content-clear-ack", session._execute(["set currentBound to (end of content of text object of boundDoc) - 1", f"set rollbackRange to create range boundDoc start {checkpoint} end currentBound", 'set content of rollbackRange to ""', 'set nativeRows to {{"content-clear-ack",(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}']))
    after = journal.add("after-content-clear", geometry(session))
    journal.add("repeat-object-delete-ack", delete_new_objects(session, checkpoint, "field"))
    journal.add("repeat-content-clear-ack", session._execute(["set currentBound to (end of content of text object of boundDoc) - 1", f"set rollbackRange to create range boundDoc start {checkpoint} end currentBound", 'set content of rollbackRange to ""', 'set nativeRows to {{"repeat-content-clear-ack",(end of content of text object of boundDoc) - 1,end of content of text object of boundDoc}}']))
    repeated = journal.add("after-repeat", geometry(session))
    journal.data["observations"] = {"before": before, "residual_range": residual, "after": after, "repeated": repeated}; journal.flush()
    return before == after == repeated


def run(output: Path, mode: str):
    output.mkdir(parents=True, exist_ok=False)
    journal = Journal(output / "report.json", mode)
    session = None
    try:
        journal.data["starting_inventory"] = inventory_without_launch(output, "starting-inventory"); journal.flush()
        assert journal.data["starting_inventory"] == []
        with create_document("writer", engine="msoffice", visible=False) as session:
            journal.data["word_version"] = journal.add("word-version", session._execute(['set nativeRows to {{version as text}}']))[0][0]
            runners = {"table-only": table_only, "fields-only": fields_only,
                       "table-object-delete-only": table_object_delete_only,
                       "field-object-delete-only": field_object_delete_only,
                       "field-object-content-clear": field_object_content_clear}
            passed = runners[mode](session, journal)
            if session.staging_root.exists(): shutil.copytree(session.staging_root, output / "native-runtime", dirs_exist_ok=True)
        journal.data["after_owned_close"] = inventory_without_launch(output, "after-owned-close")
        journal.data["passed"] = bool(passed and journal.data["after_owned_close"] == journal.data["starting_inventory"])
        journal.flush()
    except BaseException as exc:
        journal.data["error"] = {"type": type(exc).__name__, "message": str(exc)}
        journal.data["recovery_targets"] = {"owned_document": getattr(session, "_bound_path", None), "staging_root": str(getattr(session, "staging_root", ""))}
        journal.flush(); (output / "failure.txt").write_text(traceback.format_exc())
        raise


if __name__ == "__main__":
    p = argparse.ArgumentParser(); p.add_argument("--execute", action="store_true"); p.add_argument("--mode", choices=("table-only", "fields-only", "table-object-delete-only", "field-object-delete-only", "field-object-content-clear"), required=True); p.add_argument("--output", required=True)
    a = p.parse_args()
    if not a.execute: raise SystemExit("Refusing native Word mutation without --execute")
    run(Path(a.output).resolve(), a.mode)
