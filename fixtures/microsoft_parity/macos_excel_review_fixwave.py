"""Native Excel acceptance for clone destinations and logical save lifecycle."""
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


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _cells(snapshot: dict, sheet: int) -> dict:
    return {cell["address"]: cell["value"] for cell in snapshot["sheets"][sheet - 1]["cells"]}


def _osascript(output: Path, label: str, body: str) -> str:
    script = 'with timeout of 60 seconds\ntell application "/Applications/Microsoft Excel.app"\n' + body + '\nend tell\nend timeout\n'
    (output / f"{label}.applescript").write_text(script)
    result = subprocess.run(["/usr/bin/osascript", "-e", script], text=True, capture_output=True, timeout=60)
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
    (output / f"{label}.json").write_text(json.dumps(inventory, indent=2, ensure_ascii=False) + "\n")
    return inventory


def run(source: Path, output: Path) -> dict:
    source = Path(source).resolve(); output = Path(output).resolve(); output.mkdir(parents=True, exist_ok=False)
    logical = output / "logical-source.xlsx"; bound_output = output / "bound-output.xlsx"; copy_output = output / "copy-before-current.xlsx"
    shutil.copyfile(source, logical)
    implementation = Path(__file__).resolve().parents[2] / "skills/WPSComposer/scripts/msoffice/macos_excel_session.py"
    transport = Path(__file__).resolve().parents[2] / "skills/WPSComposer/scripts/artifact_transport.py"
    fixture = Path(__file__).resolve()
    report = {"verified": False, "source": str(source), "source_sha256_before": _sha(source), "logical_source": str(logical), "logical_source_sha256_before": _sha(logical), "session_source_sha256": _sha(implementation), "transport_source_sha256": _sha(transport), "fixture_sha256": _sha(fixture), "jobs": [], "owned_native_names": []}
    (output / "session-source.py").write_bytes(implementation.read_bytes()); (output / "artifact-transport-source.py").write_bytes(transport.read_bytes()); (output / "fixture-source.py").write_bytes(fixture.read_bytes())
    session = None; sentinel_name = None; marker = "wpscomposer-unsaved-sentinel-" + uuid4().hex
    try:
        before_inventory = _inventory(output, "inventory-before")
        sentinel_name = _osascript(output, "create-sentinel", f'''set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to {_quote(marker)}
return name of sentinelBook''')
        with_sentinel_inventory = _inventory(output, "inventory-with-sentinel")
        assert len(with_sentinel_inventory) == len(before_inventory) + 1

        session = open_document(logical, engine="msoffice"); report["public_session_type"] = type(session).__name__; report["jobs"].append(str(session._job)); report["owned_native_names"].append(session._native.name)
        before = session.inspect_document(max_cells=100); assert before["sheet_count"] >= 3
        source_name = before["sheets"][1]["name"]; source_a1 = _cells(before, 2)["$A$1"]; last_name = before["sheets"][-1]["name"]
        omitted = session.apply_structural_op({"op": "clone", "target": "sheet:2"}); after_omitted = session.inspect_document(max_cells=100); omitted_index = int(omitted["path"].split(":")[1])
        assert omitted_index == 3 and after_omitted["sheets"][omitted_index - 1]["name"].startswith(source_name) and _cells(after_omitted, omitted_index)["$A$1"] == source_a1 and after_omitted["sheets"][-1]["name"] == last_name
        explicit_end = session.apply_structural_op({"op": "clone", "target": "sheet:2", "to": "end"}); after_end = session.inspect_document(max_cells=100); end_index = int(explicit_end["path"].split(":")[1])
        assert end_index == after_end["sheet_count"] and after_end["sheets"][end_index - 1]["name"].startswith(source_name) and _cells(after_end, end_index)["$A$1"] == source_a1
        session.apply_format_patch("sheet:1/cell:A20", value="direct-save-current"); assert session.save_current() == str(logical); session.close(); session = None

        with open_document(logical, read_only=True, engine="msoffice") as reopened:
            report["jobs"].append(str(reopened._job)); report["owned_native_names"].append(reopened._native.name); direct_snapshot = reopened.inspect_document(max_cells=100)
            assert _cells(direct_snapshot, 1)["$A$20"] == "direct-save-current" and _cells(direct_snapshot, 3)["$A$1"] == source_a1 and _cells(direct_snapshot, direct_snapshot["sheet_count"])["$A$1"] == source_a1

        session = open_document(logical, engine="msoffice"); report["jobs"].append(str(session._job)); report["owned_native_names"].append(session._native.name)
        assert session.save(bound_output) == str(bound_output); assert session.save_copy(copy_output) == str(copy_output)
        session.apply_format_patch("sheet:1/cell:A21", value="after-save-copy-current"); assert session.save_current() == str(bound_output); session.close(); session = None

        reopen_snapshots = {}
        for label, path in (("logical", logical), ("bound", bound_output), ("copy", copy_output)):
            with open_document(path, read_only=True, engine="msoffice") as reopened:
                report["jobs"].append(str(reopened._job)); report["owned_native_names"].append(reopened._native.name); reopen_snapshots[label] = reopened.inspect_document(max_cells=100)
        assert "$A$21" not in _cells(reopen_snapshots["logical"], 1) and _cells(reopen_snapshots["bound"], 1)["$A$21"] == "after-save-copy-current" and "$A$21" not in _cells(reopen_snapshots["copy"], 1)
        (output / "reopened.json").write_text(json.dumps(reopen_snapshots, indent=2, ensure_ascii=False) + "\n")

        after_owned_close = _inventory(output, "inventory-after-owned-close"); open_names = {item["name"] for item in after_owned_close}
        assert not open_names.intersection(report["owned_native_names"]) and sentinel_name in open_names
        sentinel_state = _osascript(output, "verify-sentinel", f'''set sentinelBook to workbook {_quote(sentinel_name)}
return ((saved of sentinelBook) as text) & (character id 31) & ((value of range "A1" of worksheet 1 of sentinelBook) as text)''').split(chr(31))
        assert sentinel_state == ["false", marker]
        report.update(omitted_clone=omitted, explicit_end_clone=explicit_end, names_after=[sheet["name"] for sheet in after_end["sheets"]], direct_save_current_verified=True, save_save_copy_current_binding_verified=True, unsaved_sentinel={"name": sentinel_name, "marker": marker, "preserved": True}, source_sha256_after=_sha(source), logical_source_sha256_after=_sha(logical), bound_output_sha256=_sha(bound_output), copy_output_sha256=_sha(copy_output))
        assert report["source_sha256_after"] == report["source_sha256_before"] and report["logical_source_sha256_after"] != report["logical_source_sha256_before"]
        report["verified"] = True
    except BaseException as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"; (output / "error.log").write_text(traceback.format_exc())
    finally:
        if session is not None and not getattr(session, "_failed", False): session.close()
        for index, job in enumerate(report["jobs"], 1): shutil.copytree(job, output / f"native-logs-{index}")
        if sentinel_name is not None:
            try:
                _osascript(output, "close-exact-sentinel", f'''if {_quote(sentinel_name)} is in (name of every workbook) then close workbook {_quote(sentinel_name)} saving no
return "closed exact sentinel"''')
                after_sentinel_close = _inventory(output, "inventory-after-sentinel-close"); report["inventory_restored"] = after_sentinel_close == before_inventory
                if report["verified"]: assert report["inventory_restored"]
            except BaseException as cleanup_exc:
                report["verified"] = False; report["cleanup_error"] = f"{type(cleanup_exc).__name__}: {cleanup_exc}"
        (output / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n")
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--source", required=True, type=Path); parser.add_argument("--output-dir", required=True, type=Path); args = parser.parse_args(); result = run(args.source, args.output_dir); print(json.dumps(result, indent=2, ensure_ascii=False)); raise SystemExit(0 if result["verified"] else 1)
