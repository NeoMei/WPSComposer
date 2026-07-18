import json
from pathlib import Path
import subprocess

import pytest

ROOT = Path("macos/wps-jsapi-probe/addin")


def test_manifest_and_ribbon_are_minimal_and_load_probe():
    manifest = (ROOT / "manifest.xml").read_text()
    ribbon = (ROOT / "ribbon.xml").read_text()
    assert "<Name>WPSComposerPhase0</Name>" in manifest
    assert 'onLoad="OnAddinLoad"' in ribbon
    assert "<button" not in ribbon


def test_html_loads_shared_client_before_component():
    html = (ROOT / "index.html").read_text()
    assert html.index("bridge-client.js") < html.index("component.js")


def test_bridge_client_has_no_dynamic_code_execution():
    source = (ROOT / "bridge-client.js").read_text()
    assert "Authorization" in source
    assert "/v1/register" in source
    assert "/v1/next" in source
    assert "/v1/result" in source
    assert "eval(" not in source
    assert "new Function" not in source
    assert 'error.code === "string"' in source


def test_writer_uses_wps_save_and_pdf_export():
    source = (ROOT / "writer.js").read_text()
    assert "Application.Documents.Add" in source
    assert "Tables.Add" in source
    assert "InlineShapes.AddPicture" in source
    assert "SaveAs2(outputPath, 12)" in source
    assert "ExportAsFixedFormat(outputPath, 17" in source
    assert '"smoke_docx"' in source
    assert '"smoke_pdf"' in source


def test_writer_waits_for_wps_file_completion_event_before_close():
    source = (ROOT / "writer.js").read_text()
    assert 'AddApiEventListener("FileAfterSave"' in source
    assert 'RemoveApiEventListener("FileAfterSave")' in source
    assert "await waitForFileAfterSave(outputPath" in source
    assert "retainedDocuments" not in source


def test_writer_recovers_inserted_image_when_jsapi_returns_null():
    source = (ROOT / "writer.js").read_text()
    assert "document.InlineShapes.Count" in source
    assert "document.InlineShapes.Item(imageCount)" in source
    assert 'record("writer.image_return", "mapped"' in source


def test_writer_maps_image_to_document_shape_when_inline_insert_is_unsupported():
    source = (ROOT / "writer.js").read_text()
    assert "document.Shapes.AddPicture" in source
    assert "document.Shapes.Count" in source
    assert "document.Shapes.Item(shapeCount)" in source
    assert 'record("writer.image_api", "mapped"' in source


def test_writer_accepts_only_the_fixed_loopback_image_url():
    source = (ROOT / "writer.js").read_text()
    assert 'const WRITER_IMAGE_URL = "http://127.0.0.1:3889/fixture.png"' in source
    assert "params.imageUrl !== WRITER_IMAGE_URL" in source
    assert 'record("writer.image_source", "mapped"' in source


def test_writer_conversion_opens_exports_and_closes_without_saving():
    source = (ROOT / "writer.js").read_text()
    assert "Application.Documents.Open(sourcePath, false, true)" in source
    assert "document.ExportAsFixedFormat(outputPath, 17" in source
    assert "document.Close(0)" in source
    assert '"convert_writer_pdf": convertWriterPdf' in source


def test_presentation_uses_wps_shapes_and_save():
    source = (ROOT / "presentation.js").read_text()
    assert "Application.Presentations.Add" in source
    assert "Shapes.AddTextbox" in source
    assert "Shapes.AddPicture" in source
    assert "Shapes.AddTable" in source
    assert "SaveAs(outputPath, 24)" in source
    assert '"smoke_pptx"' in source


def test_presentation_conversion_opens_exports_and_closes_without_saving():
    source = (ROOT / "presentation.js").read_text()
    assert "Application.Presentations.Open(sourcePath, true, false, false)" in source
    assert "presentation.SaveAs(outputPath, 32)" in source
    assert "presentation.Close()" in source
    assert '"convert_presentation_pdf": convertPresentationPdf' in source


def test_spreadsheet_uses_wps_cells_chart_and_save():
    source = (ROOT / "spreadsheet.js").read_text()
    assert "Application.Workbooks.Add" in source
    assert '.Formula = "=B2*C2"' in source
    assert "ChartObjects().Add" in source
    assert "SetSourceData" in source
    assert "SaveAs(outputPath, 51)" in source
    assert '"smoke_xlsx"' in source


def test_spreadsheet_conversion_exports_the_workbook_and_closes_without_saving():
    source = (ROOT / "spreadsheet.js").read_text()
    assert "Application.Workbooks.Open(sourcePath, 0, true)" in source
    assert "workbook.ExportAsFixedFormat(0, outputPath)" in source
    assert "workbook.Close(false)" in source
    assert '"convert_workbook_pdf": convertWorkbookPdf' in source


def _run_component_failure_case(
    component_file: str,
    method: str,
    application_source: str,
    expected_code: str,
) -> None:
    path = json.dumps(str((ROOT / component_file).resolve()))
    script = f"""
const fs = require("fs");
global.window = {{}};
global.Application = {application_source};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: {json.dumps(method)},
      params: {{sourcePath: "/staged/source", outputPath: "/staged/output.pdf"}}
    }});
    console.error("handler unexpectedly succeeded");
    process.exit(2);
  }} catch (error) {{
    if (Application.DisplayAlerts !== 7) {{
      console.error(`DisplayAlerts leaked: ${{Application.DisplayAlerts}}`);
      process.exit(3);
    }}
    if (error.code !== {json.dumps(expected_code)}) {{
      console.error(`unexpected code: ${{error.code}}`);
      process.exit(4);
    }}
  }}
}}());
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


@pytest.mark.parametrize(
    ("component_file", "method", "application_source"),
    [
        (
            "writer.js",
            "convert_writer_pdf",
            '{DisplayAlerts: 7, Documents: {Open() {throw new Error("password required");}}}',
        ),
        (
            "spreadsheet.js",
            "convert_workbook_pdf",
            '{DisplayAlerts: 7, Workbooks: {Open() {throw new Error("password required");}}}',
        ),
        (
            "presentation.js",
            "convert_presentation_pdf",
            '{DisplayAlerts: 7, Presentations: {Open() {throw new Error("password required");}}}',
        ),
    ],
)
def test_conversion_open_failure_is_typed_and_restores_alerts(
    component_file: str, method: str, application_source: str
):
    _run_component_failure_case(
        component_file,
        method,
        application_source,
        "INTERACTIVE_INPUT_REQUIRED",
    )


def test_spreadsheet_rejects_workbook_without_visible_sheets():
    _run_component_failure_case(
        "spreadsheet.js",
        "convert_workbook_pdf",
        """{
          DisplayAlerts: 7,
          Workbooks: {Open() {return {
            Worksheets: {Count: 2, Item() {return {Visible: 0};}},
            ExportAsFixedFormat() {},
            Close() {}
          };}}
        }""",
        "NO_VISIBLE_WORKSHEETS",
    )


def test_spreadsheet_close_failure_still_restores_alerts():
    _run_component_failure_case(
        "spreadsheet.js",
        "convert_workbook_pdf",
        """{
          DisplayAlerts: 7,
          Workbooks: {Open() {return {
            Worksheets: {Count: 1, Item() {return {Visible: -1};}},
            ExportAsFixedFormat() {},
            Close() {throw new Error("close failed");}
          };}}
        }""",
        "CONVERSION_COMMAND_FAILED",
    )
