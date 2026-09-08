#!/usr/bin/env python3
"""Native Word AppleScript feasibility probe; never controls WPS or quits Word."""
from __future__ import annotations

import argparse
import json
import platform
import hashlib
import shutil
import tempfile
import subprocess
import sys
from pathlib import Path

CAPABILITIES = ["snapshot_before", "create", "chinese_body", "headings", "native_numbering", "long_table", "toc", "field_refresh", "repaginate", "save_docx", "reopen", "export_pdf", "cleanup", "preservation"]


def block(name: str, source: str) -> str:
    return f'''\nmy eventLine("{name}", "attempted", "")
try
{source}
my eventLine("{name}", "succeeded", "")
on error errorText number errorNumber
my eventLine("{name}", "failed", (errorNumber as text) & ": " & errorText)
end try\n'''


def script() -> str:
    stages = block("chinese_body", '''set bodyRange to text object of paragraph 4 of ownedDoc
set name of font object of bodyRange to "FangSong"
set east asian name of font object of bodyRange to "仿宋"
set font size of font object of bodyRange to 12
set first line indent of paragraph format of bodyRange to 24
set character unit first line indent of paragraph format of bodyRange to 2''')
    stages += block("headings", '''set style of paragraph 3 of ownedDoc to style heading1
set style of paragraph 5 of ownedDoc to style heading2
set style of paragraph 7 of ownedDoc to style heading3
set style of paragraph 9 of ownedDoc to style heading1
set font size of font object of text object of paragraph 3 of ownedDoc to 16
set font size of font object of text object of paragraph 5 of ownedDoc to 15
set font size of font object of text object of paragraph 7 of ownedDoc to 15
set font size of font object of text object of paragraph 9 of ownedDoc to 16''')
    stages += block("native_numbering", '''set ownList to make new list template at ownedDoc with properties {name:"MSOfficeSpikeOutline", outline numbered:true}
repeat with levelIndex from 1 to 3
set lvl to list level levelIndex of ownList
set number style of lvl to list number style arabic
set start at of lvl to 1
set reset on higher of lvl to (levelIndex - 1)
if levelIndex is 1 then
set number format of lvl to "%1"
else if levelIndex is 2 then
set number format of lvl to "%1.%2"
else
set number format of lvl to "%1.%2.%3"
end if
end repeat
set headingOne to Word style (style heading1) of ownedDoc
set headingTwo to Word style (style heading2) of ownedDoc
set headingThree to Word style (style heading3) of ownedDoc
link to list template headingOne list template ownList list level number 1
link to list template headingTwo list template ownList list level number 2
link to list template headingThree list template ownList list level number 3''')
    stages += block("long_table", '''set firstTablePosition to start of content of text object of paragraph 11 of ownedDoc
set finalPosition to (end of content of text object of ownedDoc) - 1
set tableRange to create range ownedDoc start firstTablePosition end finalPosition
set ownedTable to make new table at ownedDoc with properties {text object:tableRange, number of rows:82, number of columns:3}
set rowData to paragraphs of inputText
repeat with rowIndex from 1 to 82
set AppleScript's text item delimiters to tab
set cellData to text items of item (rowIndex + 10) of rowData
set AppleScript's text item delimiters to ""
repeat with columnIndex from 1 to 3
set ownCell to get cell from table ownedTable row rowIndex column columnIndex
set content of text object of ownCell to item columnIndex of cellData
end repeat
end repeat
set allow page breaks of ownedTable to true
set heading format of row 1 of ownedTable to true
set character unit first line indent of paragraph format of text object of ownedTable to 0
my eventLine("table_rows", "observed", (number of rows of ownedTable) as text)''')
    stages += block("toc", '''set tocPosition to start of content of text object of paragraph 2 of ownedDoc
set tocRange to create range ownedDoc start tocPosition end tocPosition
create new field text range tocRange field type field toc field text "\\\\o \\"1-3\\" \\\\h \\\\z \\\\u" preserve formatting true''')
    stages += block("field_refresh", '''if (count of fields of ownedDoc) is 0 then error "No fields were created"
repeat with fieldIndex from 1 to (count of fields of ownedDoc)
set ownField to field fieldIndex of ownedDoc
if (update field ownField) is false then error "Word update field returned false"
end repeat''')
    stages += block("repaginate", '''repaginate ownedDoc
my eventLine("page_count", "observed", (compute statistics ownedDoc statistic statistic pages) as text)''')
    return '''on textLength(valueText)
return (count of characters of valueText) as text
end textLength

on eventLine(capabilityName, stateName, detailText)
set AppleScript's text item delimiters to " "
set detailText to (paragraphs of (detailText as text)) as text
set AppleScript's text item delimiters to ""
log "EVENT" & tab & capabilityName & tab & stateName & tab & detailText
end eventLine

on run argv
set docxPath to item 1 of argv
set pdfPath to item 2 of argv
set inputPath to item 3 of argv
set inputText to read POSIX file inputPath as «class utf8»
set ownedDoc to missing value
set beforeDocs to {}
set coreOK to false
set snapshotReady to false
with timeout of 600 seconds
 tell application "Microsoft Word"
  my eventLine("engine", "observed", version)
  try
   my eventLine("snapshot_before", "attempted", "")
   repeat with documentIndex from 1 to (count of documents)
    set userDoc to document documentIndex
    set end of beforeDocs to {name of userDoc, posix full name of userDoc, saved of userDoc, content of text object of userDoc}
    my eventLine("existing_before", "observed", (name of userDoc) & "|" & (posix full name of userDoc) & "|saved=" & (saved of userDoc as text) & "|characters=" & (my textLength(content of text object of userDoc)))
   end repeat
   set snapshotReady to true
   my eventLine("snapshot_before", "succeeded", (count beforeDocs) as text)
   my eventLine("create", "attempted", "")
   set ownedDoc to make new document
   set ownedName to name of ownedDoc
   repeat with prior in beforeDocs
    if item 1 of prior is ownedName then
     set ownedDoc to missing value
     error "New document name collided with a preexisting document; refusing mutation"
    end if
   end repeat
   set ownedDoc to document ownedName
   my eventLine("owned_identity", "observed", ownedName)
   set content of text object of ownedDoc to inputText
   my eventLine("create", "succeeded", ownedName)
''' + stages + '''
   my eventLine("save_docx", "attempted", "")
   save as ownedDoc file name docxPath file format format document default add to recent files false
   set ownedDoc to first document whose posix full name is docxPath
   my eventLine("save_docx", "succeeded", posix full name of ownedDoc)
   close ownedDoc saving no
   set ownedDoc to missing value
   my eventLine("reopen", "attempted", "")
   open file name docxPath add to recent files false
   set ownedDoc to first document whose posix full name is docxPath
   if (content of text object of ownedDoc) does not contain "MSOFFICE-SPIKE-END" then error "Reopened document missing end marker"
   my eventLine("reopen", "succeeded", posix full name of ownedDoc)
   set coreOK to true
''' + block("export_pdf", '''save as ownedDoc file name pdfPath file format format PDF add to recent files false
-- PDF SaveAs may change the document identity; reacquire only an exact owned path.
set ownedDoc to missing value
repeat with documentIndex from 1 to (count of documents)
set candidate to document documentIndex
if (posix full name of candidate is docxPath) or (posix full name of candidate is pdfPath) then set ownedDoc to candidate
end repeat''') + '''
  on error errorText number errorNumber
   my eventLine("core", "failed", (errorNumber as text) & ": " & errorText)
  end try
  my eventLine("cleanup", "attempted", "")
  try
   if ownedDoc is not missing value then close ownedDoc saving no
   set ownedDoc to missing value
   my eventLine("cleanup", "succeeded", "Only task-owned document; Word left running")
  on error errorText number errorNumber
   my eventLine("cleanup", "failed", (errorNumber as text) & ": " & errorText)
  end try
  my eventLine("preservation", "attempted", "")
  try
   if snapshotReady is false then error "Initial snapshot incomplete"
   if (count of documents) is not (count beforeDocs) then error "Open-document count changed"
   repeat with prior in beforeDocs
    set matchingDoc to document (item 1 of prior)
    if posix full name of matchingDoc is not item 2 of prior then error "Preexisting document path changed"
    if saved of matchingDoc is not item 3 of prior then error "Preexisting document saved state changed"
    if content of text object of matchingDoc is not item 4 of prior then error "Preexisting document text changed"
    my eventLine("existing_after", "observed", (name of matchingDoc) & "|" & (posix full name of matchingDoc) & "|saved=" & (saved of matchingDoc as text) & "|characters=" & (my textLength(content of text object of matchingDoc)))
   end repeat
   my eventLine("preservation", "succeeded", "Names, exact paths, saved flags and full in-memory text equal")
  on error errorText number errorNumber
   my eventLine("preservation", "failed", (errorNumber as text) & ": " & errorText)
  end try
 end tell
end timeout
if coreOK is false then error "Native core generation/save/reopen failed; see EVENT records"
return "Native probe ended; inspect capability gaps and artifacts separately"
end run
'''


def content() -> str:
    lines = ["目录", "", "原生办公文档验证", "这是中文正文，用于验证仿宋字体、十二磅字号和首行缩进两个字符。", "层级标题验证", "本节检查标题样式与原生多级编号。", "三级标题验证", "本段用于检查目录刷新后的层次结构和页码。", "长表格验证", "以下表格包含八十条中文记录，检查跨页排版和重复表头。", "编号\t检查内容\t结论"]
    lines += [f"{i:03d}\t原生排版验证第{i}条：中文内容自动换行，检查表格跨页和行高。\t待核验" for i in range(1, 81)]
    lines += ["MSOFFICE-SPIKE-END\t结束标记\t完成", ""]
    return "\r".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--compile-only", action="store_true", help="Compile AppleScript without executing Word events")
    args = parser.parse_args()
    out = args.output_dir.expanduser().resolve()
    out.mkdir(parents=True, exist_ok=False)
    result = {"engine": "Microsoft Word", "transport": "native AppleScript", "platform": platform.platform(), "native_run": False, "capabilities": {key: {"status": "unrun"} for key in CAPABILITIES}, "events": [], "errors": []}
    source = out / "probe.applescript"
    source.write_text(script(), encoding="utf-8")
    (out / "input.txt").write_text(content(), encoding="utf-8")
    code = 1
    try:
        if sys.platform != "darwin":
            raise RuntimeError("This runner requires macOS and Microsoft Word")
        compiled = subprocess.run(["osacompile", "-o", str(out / "probe.scpt"), str(source)], capture_output=True, text=True, timeout=60)
        (out / "compile.stderr.txt").write_text(compiled.stderr, encoding="utf-8")
        compiled.check_returncode()
        result["compiled"] = True
        if args.compile_only:
            result["overall"] = "unrun_compile_only"
            code = 0
        else:
            result["native_run"] = True
            # Word can natively save within its own container without granting
            # each unrelated caller output directory access in a dialog.
            sandbox_tmp = Path.home() / "Library/Containers/com.microsoft.Word/Data/tmp"
            if not sandbox_tmp.is_dir():
                raise RuntimeError("Word sandbox Data/tmp is absent; native save path not established")
            native_dir = Path(tempfile.mkdtemp(prefix="msoffice-spike-", dir=str(sandbox_tmp)))
            result["native_staging_directory"] = str(native_dir)
            result["native_save_strategy"] = "Fresh task-owned Word-container staging, then byte-identical Python copy to requested output directory"
            proc = subprocess.run(["osascript", str(out / "probe.scpt"), str(native_dir / "probe.docx"), str(native_dir / "probe.pdf"), str(out / "input.txt")], capture_output=True, text=True, timeout=660)
            (out / "stdout.txt").write_text(proc.stdout, encoding="utf-8")
            (out / "stderr.txt").write_text(proc.stderr, encoding="utf-8")
            for line in proc.stderr.splitlines():
                if not line.startswith("EVENT\t"):
                    continue
                _, key, status, detail = (line.split("\t", 3) + [""])[:4]
                result["events"].append({"capability": key, "status": status, "detail": detail})
                if key in result["capabilities"]:
                    result["capabilities"][key] = {"status": status, "detail": detail}
                if status == "failed":
                    result["errors"].append({"capability": key, "raw_error": detail})
                if key == "engine":
                    result["version"] = detail
            for fmt in ("docx", "pdf"):
                native_artifact = native_dir / ("probe." + fmt)
                if native_artifact.is_file():
                    shutil.copyfile(str(native_artifact), str(out / native_artifact.name))
                    result["native_" + fmt] = {"path": str(native_artifact), "sha256": hashlib.sha256(native_artifact.read_bytes()).hexdigest()}
                artifact = out / ("probe." + fmt)
                result[fmt] = {"path": str(artifact), "exists": artifact.is_file(), "bytes": artifact.stat().st_size if artifact.is_file() else 0}
            core = proc.returncode == 0 and result["docx"]["bytes"] > 0 and all(result["capabilities"][key]["status"] == "succeeded" for key in ("create", "save_docx", "reopen", "cleanup", "preservation"))
            all_ok = core and result["pdf"]["bytes"] > 0 and all(x["status"] == "succeeded" for x in result["capabilities"].values())
            result["overall"] = "native_operations_succeeded_artifact_validation_pending" if all_ok else "partial_capability_gaps" if core else "failed"
            result["process_returncode"] = proc.returncode
            code = 0 if all_ok else 2 if core else 1
    except Exception as exc:
        result["overall"] = "failed"
        result["errors"].append({"raw_error": repr(exc)})
        if isinstance(exc, subprocess.TimeoutExpired):
            for key in ("stdout", "stderr"):
                value = getattr(exc, key, None) or b""
                if isinstance(value, bytes):
                    value = value.decode("utf-8", errors="replace")
                (out / (key + ".txt")).write_text(value, encoding="utf-8")
            result["cleanup_warning"] = "Timed out: ownership cleanup is unverified; inspect Word, do not broadly quit/kill it"
    finally:
        (out / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(str(out / "result.json"))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
