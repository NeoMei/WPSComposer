"""Native Excel acceptance for recoverable close(save_changes=True) conflict."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import traceback
from uuid import uuid4

from skills.WPSComposer import open_document
from skills.WPSComposer.scripts.msoffice.macos_excel_session import _quote


ROOT = Path(__file__).resolve().parents[2]


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _osascript(output: Path, label: str, body: str) -> str:
    source = (
        'with timeout of 60 seconds\n'
        'tell application "/Applications/Microsoft Excel.app"\n'
        + body + '\nend tell\nend timeout\n'
    )
    (output / f"{label}.applescript").write_text(source)
    result = subprocess.run(
        ["/usr/bin/osascript", "-e", source],
        text=True, capture_output=True, timeout=60,
    )
    (output / f"{label}.stdout.log").write_text(result.stdout)
    (output / f"{label}.stderr.log").write_text(result.stderr)
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or f"{label} failed")
    return result.stdout.rstrip("\n")


def _inventory(output: Path, label: str) -> list[dict]:
    raw = _osascript(output, label, '''set inventoryRows to {}
repeat with workbookIndex from 1 to (count of workbooks)
 set candidateBook to workbook workbookIndex
 try
  set candidatePath to full name of candidateBook
 on error
  set candidatePath to "__UNSAVED__"
 end try
 set end of inventoryRows to ((name of candidateBook) & (character id 31) & ((saved of candidateBook) as text) & (character id 31) & candidatePath)
end repeat
set AppleScript's text item delimiters to character id 30
return inventoryRows as text''')
    rows = [] if not raw else raw.split(chr(30))
    inventory = []
    for row in rows:
        name, saved, path = row.split(chr(31))
        inventory.append({"name": name, "saved": saved == "true", "path": path})
    inventory.sort(key=lambda item: (item["name"], item["path"]))
    (output / f"{label}.json").write_text(
        json.dumps(inventory, ensure_ascii=False, indent=2) + "\n"
    )
    return inventory


def _cell(snapshot: dict, address: str):
    cells = {
        item["address"]: item["value"]
        for item in snapshot["sheets"][0]["cells"]
    }
    return cells.get(address)


def run(source: Path, output: Path) -> dict:
    source = Path(source).resolve()
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    logical = output / "logical.xlsx"
    bound = output / "bound.xlsx"
    concurrent = output / "concurrent.xlsx"
    recovery = output / "recovery.xlsx"
    recovery_pdf = output / "recovery.pdf"
    shutil.copyfile(source, logical)
    production = [
        ROOT / "skills/WPSComposer/scripts/artifact_transport.py",
        ROOT / "skills/WPSComposer/scripts/msoffice/macos_excel_session.py",
        Path(__file__).resolve(),
    ]
    report = {
        "passed": False,
        "source": str(source),
        "source_sha256_before": _sha(source),
        "source_hashes": {
            str(path.relative_to(ROOT)): _sha(path) for path in production
        },
        "checks": {},
    }
    for path in production:
        retained = output / "source" / path.relative_to(ROOT)
        retained.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, retained)
    session = None
    sentinel_name = None
    before_inventory = None
    marker = "integration-round2-sentinel-" + uuid4().hex
    try:
        before_inventory = _inventory(output, "inventory-before")
        sentinel_name = _osascript(output, "create-sentinel", f'''set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to {_quote(marker)}
return name of sentinelBook''')
        sentinel_inventory = _inventory(output, "inventory-with-sentinel")
        assert len(sentinel_inventory) == len(before_inventory) + 1

        session = open_document(logical, engine="msoffice")
        owned_name = session._native.name
        owned_path = str(session._native)
        report["owned_name"] = owned_name
        report["owned_path"] = owned_path
        report["job"] = str(session._job)
        session.apply_format_patch(
            "sheet:1/cell:A30", value="PUBLISHED BEFORE CONFLICT"
        )
        session.save(bound)
        session.save_copy(concurrent)
        session.apply_format_patch(
            "sheet:1/cell:A31", value="UNPUBLISHED RECOVERY EDIT"
        )
        shutil.copyfile(concurrent, bound)
        concurrent_hash = _sha(bound)
        calls_before_failed_close = session._counter
        try:
            session.close(save_changes=True)
        except Exception as error:
            assert "changed" in str(error).lower()
            report["failed_close"] = {
                "type": type(error).__name__, "message": str(error)
            }
        else:
            raise AssertionError("concurrent destination change was accepted")

        assert not session._closed and not session._failed
        assert session._lock.stream is not None
        assert _sha(bound) == concurrent_hash
        during = _inventory(output, "inventory-after-failed-close")
        matches = [
            item for item in during
            if item["name"] == owned_name and item["path"] == owned_path
        ]
        assert len(matches) == 1
        report["checks"]["failed_save_close_kept_exact_owned_workbook_open"] = True
        report["checks"]["failed_save_close_kept_exclusive_lock"] = True
        report["checks"]["concurrent_destination_preserved"] = True
        report["native_calls_added_by_failed_close"] = (
            session._counter - calls_before_failed_close
        )

        session.save(recovery)
        session.export_pdf(recovery_pdf)
        session.close(save_changes=False)
        assert session._closed and session._lock.stream is None
        close_scripts = [
            path for path in session._job.glob("step-*.applescript")
            if "close ownedBook saving no" in path.read_text()
        ]
        assert len(close_scripts) == 1
        report["checks"]["save_as_recovery_then_discard_closed_exactly_once"] = True
        report["checks"]["public_pdf_export_preserved_native_binding"] = (
            recovery_pdf.is_file()
        )
        report["close_script"] = close_scripts[0].name

        with open_document(bound, read_only=True, engine="msoffice") as reopened:
            bound_snapshot = reopened.inspect_document(max_cells=200)
        with open_document(recovery, read_only=True, engine="msoffice") as reopened:
            recovery_snapshot = reopened.inspect_document(max_cells=200)
        assert _cell(bound_snapshot, "$A$30") == "PUBLISHED BEFORE CONFLICT"
        assert _cell(bound_snapshot, "$A$31") is None
        assert _cell(recovery_snapshot, "$A$30") == "PUBLISHED BEFORE CONFLICT"
        assert _cell(recovery_snapshot, "$A$31") == "UNPUBLISHED RECOVERY EDIT"
        report["checks"]["native_reopen_separates_foreign_and_recovery_bytes"] = True

        after_owned = _inventory(output, "inventory-after-owned-close")
        assert after_owned == sentinel_inventory
        state = _osascript(output, "verify-sentinel", f'''set sentinelBook to workbook {_quote(sentinel_name)}
return ((saved of sentinelBook) as text) & (character id 31) & ((value of range "A1" of worksheet 1 of sentinelBook) as text)''').split(chr(31))
        assert state == ["false", marker]
        report["checks"]["unrelated_unsaved_sentinel_preserved"] = True
        report["artifact_hashes"] = {
            path.name: _sha(path)
            for path in (logical, bound, concurrent, recovery, recovery_pdf)
        }
        assert _sha(source) == report["source_sha256_before"]
        report["checks"]["original_source_preserved"] = True
        report["passed"] = all(report["checks"].values())
    except BaseException as error:
        report["error"] = {"type": type(error).__name__, "message": str(error)}
        (output / "failure.txt").write_text(traceback.format_exc())
    finally:
        if session is not None and not session._closed and not session._failed:
            try:
                session.close(save_changes=False)
            except BaseException as cleanup_error:
                report["session_cleanup_error"] = {
                    "type": type(cleanup_error).__name__,
                    "message": str(cleanup_error),
                }
                report["passed"] = False
        if report.get("job"):
            shutil.copytree(report["job"], output / "native-job", dirs_exist_ok=True)
        if sentinel_name is not None:
            try:
                state = _osascript(output, "verify-sentinel-before-close", f'''set sentinelBook to workbook {_quote(sentinel_name)}
return ((saved of sentinelBook) as text) & (character id 31) & ((value of range "A1" of worksheet 1 of sentinelBook) as text)''').split(chr(31))
                if state == ["false", marker]:
                    _osascript(output, "close-exact-sentinel", f'''close workbook {_quote(sentinel_name)} saving no
return "closed exact sentinel"''')
                    report["sentinel_closed_exactly"] = True
                else:
                    report["sentinel_closed_exactly"] = False
                    report["passed"] = False
            except BaseException as cleanup_error:
                report["sentinel_cleanup_error"] = {
                    "type": type(cleanup_error).__name__,
                    "message": str(cleanup_error),
                }
                report["passed"] = False
        if before_inventory is not None:
            try:
                report["inventory_restored"] = (
                    _inventory(output, "inventory-final") == before_inventory
                )
                if not report["inventory_restored"]:
                    report["passed"] = False
            except BaseException as cleanup_error:
                report["inventory_error"] = {
                    "type": type(cleanup_error).__name__,
                    "message": str(cleanup_error),
                }
                report["passed"] = False
        (output / "report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        )
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    if not args.execute:
        parser.error("--execute is required")
    result = run(args.source, args.output_dir)
    print(json.dumps({"passed": result["passed"], "error": result.get("error")}))
    raise SystemExit(0 if result["passed"] else 1)
