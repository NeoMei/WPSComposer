# macOS WPS JSAPI Phase 0 Feasibility

Status: **PASSED 2026-07-20** — all 4 artifacts validated

## Setup

```bash
cd macos/wps-jsapi-probe
PATH=/Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin:$PATH npm ci
```

## Run

```bash
.venv/bin/python -m skills.WPSComposer.scripts.macos_probe \
  --node /Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  --output-dir build/macos-phase0 \
  --timeout 120
```

WPS must not be running before the probe starts (wpsjs registers add-ins via publish.xml at launch).

## Tested versions

| Item | Version |
|------|---------|
| macOS | Darwin (Apple Silicon) |
| WPS Office | 12.1.26035 |
| Node.js | v24.14.0 |
| wpsjs | 2.2.3 |
| wps-jsapi-declare | 2.2.0 |

## Artifacts

| Format | Size | WPS API | Validation |
|--------|------|---------|------------|
| DOCX | 11,181 B | Document.SaveAs2 format 16 (sandbox relay) | ✅ PK zip + word/document.xml |
| PPTX | 72,995 B | Presentation.SaveAs format 24 | ✅ PK zip + ppt/presentation.xml |
| XLSX | 13,181 B | Workbook.SaveAs format 51 | ✅ PK zip + xl/workbook.xml |
| PDF | 745,786 B | Document.ExportAsFixedFormat format 17 | ✅ %PDF- signature |

## Key findings

- **Writer SaveAs2 sandbox restriction**: `Document.SaveAs2` silently fails when targeting paths outside the WPS sandbox container. Workaround: save to `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/Documents/` then copy out from the host side. Classification: `mapped`.
- **Writer InlineShapes.AddPicture**: works with absolute paths (earlier null was a relative-path issue).
- **wpsjs --port flag**: broken in 2.2.3 (rest/spread parameter bug); probe uses portfinder sequential ports detected from server logs.
- **Presentation.SaveAs / Workbook.SaveAs**: work natively with absolute paths (no sandbox restriction).
- **ExportAsFixedFormat**: works natively for PDF export from Writer.

## Registration rollback

The probe snapshots `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml` before starting and restores it atomically on exit (including error and timeout paths). Verified: file was ABSENT before and ABSENT after the run.

Recovery directory is printed to stderr before any modification; if the process is killed, use it to restore manually:

```bash
recovery="<printed path>"
cp "$recovery/publish.xml.original" "$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
# or if registration.json says "existed": false:
rm "$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
```

## Phase 1 decision

**Proceed.** All four output formats (DOCX, PPTX, XLSX, PDF) are validated through Mac WPS JSAPI. DOCX requires a sandbox relay (save to container, copy out). Registration rollback is verified safe.
