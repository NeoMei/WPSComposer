#!/usr/bin/env python3
"""Standalone native Microsoft Word COM spike (requires Windows and pywin32)."""
from __future__ import annotations

import argparse
import hashlib
import json
import platform
import sys
import traceback
from pathlib import Path

from mac_word import CAPABILITIES, content


def snapshot(app):
    """Only inspect user documents; hashes avoid persisting their text."""
    rows = []
    for index in range(1, app.Documents.Count + 1):
        doc = app.Documents.Item(index)
        text = doc.Content.Text
        rows.append({"name": doc.Name, "full_name": doc.FullName, "saved": bool(doc.Saved), "characters": len(text), "text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest()})
    return sorted(rows, key=lambda row: (row["full_name"], row["name"]))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    out = args.output_dir.expanduser().resolve()
    out.mkdir(parents=True, exist_ok=False)
    result = {"engine": "Microsoft Word", "transport": "DispatchEx Word.Application COM", "platform": platform.platform(), "native_run": False, "capabilities": {key: {"status": "unrun"} for key in CAPABILITIES}, "events": [], "errors": [], "existing_document_scope": "GetActiveObject registered Word instance plus isolated DispatchEx instance; other unregistered Word instances are not enumerated"}
    app = doc = previous_app = None
    exclusive_instance = False
    before_previous = None
    core_ok = False

    def record(key, state, detail=None):
        event = {"capability": key, "status": state, "detail": detail}
        result["events"].append(event)
        if key in result["capabilities"]:
            result["capabilities"][key] = {"status": state, "detail": detail}
        (out / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    def attempt(key, action, required=False):
        record(key, "attempted")
        try:
            value = action()
            record(key, "succeeded", value)
            return value
        except Exception as exc:
            raw = traceback.format_exc()
            result["errors"].append({"capability": key, "raw_error": repr(exc), "traceback": raw})
            record(key, "failed", repr(exc))
            if required:
                raise
            return None

    try:
        if sys.platform != "win32":
            raise RuntimeError("Windows native execution is unrun: this runner requires Windows and installed Microsoft Word")
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            previous_app = win32com.client.GetActiveObject("Word.Application")
        except Exception as exc:
            # Only MK_E_UNAVAILABLE means no registered running object. Other
            # failures cannot establish a safe baseline and must stop the run.
            if getattr(exc, "hresult", None) != -2147221021:
                raise
            result["existing_registered_instance"] = "none returned by GetActiveObject"
            result["get_active_object_observation"] = repr(exc)
        if previous_app is not None:
            before_previous = snapshot(previous_app)
            result["existing_before"] = before_previous
            result["existing_application_hwnd"] = int(previous_app.Hwnd)
        else:
            result["existing_before"] = []
        app = win32com.client.DispatchEx("Word.Application")
        result["native_run"] = True
        result["version"] = str(app.Version)
        result["build"] = str(app.Build)
        result["owned_application_hwnd"] = int(app.Hwnd)
        if previous_app is not None and int(app.Hwnd) == int(previous_app.Hwnd):
            raise RuntimeError("DispatchEx returned the preexisting application; refusing mutation or Quit")
        result["isolated_existing_before"] = snapshot(app)
        if result["isolated_existing_before"]:
            raise RuntimeError("DispatchEx application already contains documents; refusing mutation or Quit")
        exclusive_instance = True
        record("snapshot_before", "succeeded", "Registered preexisting instance and isolated instance captured")
        app.Visible = True  # This setting belongs only to the isolated instance.
        record("create", "attempted")
        doc = app.Documents.Add()
        record("create", "succeeded", doc.Name)
        result["owned_document_initial_name"] = doc.Name
        doc.Content.Text = content()

        def body():
            rng = doc.Paragraphs.Item(4).Range
            rng.Font.Name = "FangSong"
            rng.Font.NameFarEast = "仿宋"
            rng.Font.Size = 12
            rng.ParagraphFormat.FirstLineIndent = 24
            rng.ParagraphFormat.CharacterUnitFirstLineIndent = 2
            return {"font": rng.Font.NameFarEast, "size": rng.Font.Size, "first_line_points": rng.ParagraphFormat.FirstLineIndent, "first_line_characters": rng.ParagraphFormat.CharacterUnitFirstLineIndent}

        attempt("chinese_body", body)

        def headings():
            for paragraph_index, style_id, size in ((3, -2, 16), (5, -3, 15), (7, -4, 15), (9, -2, 16)):
                rng = doc.Paragraphs.Item(paragraph_index).Range
                rng.Style = doc.Styles.Item(style_id)
                rng.Font.Size = size
            return "Built-in Heading 1/2/3 assigned; sizes 16/15/15 points"

        attempt("headings", headings)

        def numbering():
            template = doc.ListTemplates.Add(OutlineNumbered=True, Name="MSOfficeSpikeOutline")
            for level, style_id in ((1, -2), (2, -3), (3, -4)):
                item = template.ListLevels.Item(level)
                item.NumberStyle = 0  # wdListNumberStyleArabic
                item.NumberFormat = ".".join("%" + str(i) for i in range(1, level + 1))
                item.StartAt = 1
                item.ResetOnHigher = level - 1
                doc.Styles.Item(style_id).LinkToListTemplate(ListTemplate=template, ListLevelNumber=level)
            return "Document-owned outline list template linked to Heading 1/2/3"

        attempt("native_numbering", numbering)

        def table():
            table_range = doc.Range(doc.Paragraphs.Item(11).Range.Start, doc.Content.End - 1)
            native_table = table_range.ConvertToTable(Separator=1, NumColumns=3)  # wdSeparateByTabs
            native_table.Rows.Item(1).HeadingFormat = True
            native_table.Rows.AllowBreakAcrossPages = True
            native_table.Range.ParagraphFormat.FirstLineIndent = 0
            native_table.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            return {"rows": native_table.Rows.Count, "columns": native_table.Columns.Count}

        attempt("long_table", table)

        def toc():
            doc.TablesOfContents.Add(Range=doc.Range(0, 0), UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=3, IncludePageNumbers=True, RightAlignPageNumbers=True, UseHyperlinks=True)
            return {"toc_count": doc.TablesOfContents.Count}

        attempt("toc", toc)

        def fields():
            if doc.Fields.Count == 0:
                raise RuntimeError("No fields created")
            failed_index = doc.Fields.Update()
            if failed_index != 0:
                raise RuntimeError("Fields.Update returned failed field index: " + str(failed_index))
            for index in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents.Item(index).Update()
            return {"field_count": doc.Fields.Count}

        attempt("field_refresh", fields)

        def paginate():
            doc.Repaginate()
            result["page_count"] = doc.ComputeStatistics(2)  # wdStatisticPages
            return result["page_count"]

        attempt("repaginate", paginate)
        docx_path = str(out / "probe.docx")
        pdf_path = str(out / "probe.pdf")
        attempt("save_docx", lambda: doc.SaveAs2(FileName=docx_path, FileFormat=16, AddToRecentFiles=False), required=True)
        if Path(doc.FullName).resolve() != Path(docx_path):
            raise RuntimeError("Saved document identity does not match owned path")
        result["owned_document_path"] = doc.FullName
        doc.Close(SaveChanges=0)
        doc = None
        record("reopen", "attempted")
        doc = app.Documents.Open(FileName=docx_path, AddToRecentFiles=False, ReadOnly=False)
        if Path(doc.FullName).resolve() != Path(docx_path) or "MSOFFICE-SPIKE-END" not in doc.Content.Text:
            raise RuntimeError("Reopened document identity/content verification failed")
        record("reopen", "succeeded", doc.FullName)
        core_ok = True
        attempt("export_pdf", lambda: doc.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17, OpenAfterExport=False))
    except Exception as exc:
        result["errors"].append({"capability": "core", "raw_error": repr(exc), "traceback": traceback.format_exc()})
        record("core", "failed", repr(exc))
    finally:
        if doc is not None:
            attempt("cleanup", lambda: doc.Close(SaveChanges=0))
            doc = None
        elif app is not None and exclusive_instance:
            record("cleanup", "succeeded", "No task-owned document remains open")
        if app is not None and exclusive_instance:
            try:
                if app.Documents.Count == 0:
                    app.Quit(SaveChanges=0)
                    result["owned_application_cleanup"] = "Quit isolated DispatchEx instance after confirming zero documents"
                else:
                    record("cleanup", "failed", "Unexpected documents appeared; isolated application left open")
            except Exception as exc:
                record("cleanup", "failed", repr(exc))
        if previous_app is not None and before_previous is not None:
            try:
                result["existing_after"] = snapshot(previous_app)
                if result["existing_after"] != before_previous:
                    raise RuntimeError("Preexisting document names/paths/saved flags/content hashes changed")
                record("preservation", "succeeded", "Preexisting registered instance snapshots equal")
            except Exception as exc:
                record("preservation", "failed", repr(exc))
        elif exclusive_instance:
            record("preservation", "succeeded", "No registered preexisting instance; isolated instance started empty")
        for fmt in ("docx", "pdf"):
            path = out / ("probe." + fmt)
            result[fmt] = {"path": str(path), "exists": path.is_file(), "bytes": path.stat().st_size if path.is_file() else 0}
        core_ok = core_ok and result["docx"]["bytes"] > 0 and all(result["capabilities"][key]["status"] == "succeeded" for key in ("cleanup", "preservation"))
        all_ok = core_ok and result["pdf"]["bytes"] > 0 and all(row["status"] == "succeeded" for row in result["capabilities"].values())
        result["overall"] = "native_operations_succeeded_artifact_validation_pending" if all_ok else "partial_capability_gaps" if core_ok else "failed"
        (out / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(out / "result.json"))
    return 0 if all_ok else 2 if core_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
