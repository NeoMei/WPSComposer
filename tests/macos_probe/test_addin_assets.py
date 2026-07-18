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
    assert "ERROR_CODES.indexOf(error.code)" in source


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


@pytest.mark.parametrize(
    ("component_file", "method", "application_source"),
    [
        (
            "writer.js",
            "convert_writer_pdf",
            """{DisplayAlerts: 7, Documents: {Open() {
              const error = new Error("vendor failure");
              error.code = "WPS_E_7001";
              throw error;
            }}}""",
        ),
        (
            "spreadsheet.js",
            "convert_workbook_pdf",
            """{DisplayAlerts: 7, Workbooks: {Open() {
              const error = new Error("vendor failure");
              error.code = "WPS_E_7001";
              throw error;
            }}}""",
        ),
        (
            "presentation.js",
            "convert_presentation_pdf",
            """{DisplayAlerts: 7, Presentations: {Open() {
              const error = new Error("vendor failure");
              error.code = "WPS_E_7001";
              throw error;
            }}}""",
        ),
    ],
)
def test_conversion_rejects_vendor_error_codes(
    component_file: str, method: str, application_source: str
):
    _run_component_failure_case(
        component_file,
        method,
        application_source,
        "CONVERSION_COMMAND_FAILED",
    )


GENERATION_CASES = [
    (
        "writer.js",
        "generate_writer_document",
        "writer",
        '[{op: "writer.reset", args: {}}, '
        '{op: "writer.add_paragraph", args: '
        '{text: "IN_PLACE_MARKER", style: "Body Text"}}]',
        """
const document = {
  saveCalls: 0, saveAsCalls: 0, closeArgument: null,
  Content: {Text: "template", InsertAfter(value) {this.Text += value;}},
  Save() {this.saveCalls += 1;},
  SaveAs() {this.saveAsCalls += 1;},
  SaveAs2() {this.saveAsCalls += 1;},
  Close(value) {this.closeArgument = value;}
};
global.Application = {
  DisplayAlerts: 7,
  Documents: {Open(path, confirm, readOnly) {
    assert.equal(path, "/staged/generated.docx");
    assert.equal(confirm, false);
    assert.equal(readOnly, false);
    return document;
  }}
};
""",
        "document",
        "assert.equal(document.closeArgument, 0);",
    ),
    (
        "spreadsheet.js",
        "generate_spreadsheet_workbook",
        "spreadsheet",
        '[{op: "sheet.reset", args: {}}, '
        '{op: "sheet.write_table", args: '
        '{startRow: 1, startCol: 1, values: [["IN_PLACE_MARKER"]]}}]',
        """
const cell = {Value2: null};
const sheet = {Cells: {Clear() {}, Item(row, column) {
  assert.equal(row, 1); assert.equal(column, 1); return cell;
}}};
const document = {
  saveCalls: 0, saveAsCalls: 0, closeArgument: null,
  Worksheets: {Item(index) {assert.equal(index, 1); return sheet;}},
  Save() {this.saveCalls += 1;},
  SaveAs() {this.saveAsCalls += 1;},
  Close(value) {this.closeArgument = value;}
};
global.Application = {
  DisplayAlerts: 7,
  Workbooks: {Open(path, updateLinks, readOnly) {
    assert.equal(path, "/staged/generated.xlsx");
    assert.equal(updateLinks, 0);
    assert.equal(readOnly, false);
    return document;
  }}
};
""",
        "document",
        "assert.equal(document.closeArgument, false); assert.equal(cell.Value2, \"IN_PLACE_MARKER\");",
    ),
    (
        "presentation.js",
        "generate_presentation_deck",
        "presentation",
        '[{op: "slide.reset", args: {}}, '
        '{op: "slide.add_title", args: {title: "IN_PLACE_MARKER"}}]',
        """
let slideCount = 1;
const textRange = {Text: ""};
const slides = {
  get Count() {return slideCount;},
  Item(index) {assert.equal(index, 1); return {Delete() {slideCount -= 1;}};},
  Add(index, layout) {
    assert.equal(index, 1); assert.equal(layout, 12); slideCount += 1;
    return {Shapes: {AddTextbox() {return {TextFrame: {TextRange: textRange}};}}};
  }
};
const document = {
  saveCalls: 0, saveAsCalls: 0, closeArgument: null,
  Slides: slides,
  Save() {this.saveCalls += 1;},
  SaveAs() {this.saveAsCalls += 1;},
  Close() {this.closeArgument = "closed";}
};
global.Application = {
  DisplayAlerts: 7,
  Presentations: {Open(path, readOnly, untitled, withWindow) {
    assert.equal(path, "/staged/generated.pptx");
    assert.equal(readOnly, false);
    assert.equal(untitled, false);
    assert.equal(withWindow, false);
    return document;
  }}
};
""",
        "document",
        "assert.equal(document.closeArgument, \"closed\"); assert.equal(textRange.Text, \"IN_PLACE_MARKER\");",
    ),
]


@pytest.mark.parametrize(
    (
        "component_file",
        "method",
        "component",
        "operations",
        "application_source",
        "document_name",
        "component_assertions",
    ),
    GENERATION_CASES,
)
def test_generation_saves_staged_template_in_place(
    component_file: str,
    method: str,
    component: str,
    operations: str,
    application_source: str,
    document_name: str,
    component_assertions: str,
):
    path = json.dumps(str((ROOT / component_file).resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
{application_source}
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  const result = await window.WPSComposerProbe.handleCommand({{
    method: {json.dumps(method)},
    params: {{
      stagedPath: "/staged/generated.{ {'writer': 'docx', 'spreadsheet': 'xlsx', 'presentation': 'pptx'}[component] }",
      plan: {{component: {json.dumps(component)}, operations: {operations}}}
    }}
  }});
  assert.equal(result.path, "/staged/generated.{ {'writer': 'docx', 'spreadsheet': 'xlsx', 'presentation': 'pptx'}[component] }");
  assert.equal(result.appliedOperations, 2);
  assert.equal({document_name}.saveCalls, 1);
  assert.equal({document_name}.saveAsCalls, 0);
  assert.equal(Application.DisplayAlerts, 7);
  {component_assertions}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def _generation_failure_application(component: str, stage: str) -> str:
    collection = {
        "writer": "Documents",
        "spreadsheet": "Workbooks",
        "presentation": "Presentations",
    }[component]
    if stage == "mutation":
        mutation = {
            "writer": 'Content: {Text: "", InsertAfter() {throw new Error("mutation failed");}},',
            "spreadsheet": "Worksheets: {Item() {return {Cells: {Clear() {}, Item() {throw new Error(\"mutation failed\");}}};}},",
            "presentation": "Slides: {Count: 0, Add() {throw new Error(\"mutation failed\");}},",
        }[component]
    else:
        mutation = {
            "writer": 'Content: {Text: "", InsertAfter() {}},',
            "spreadsheet": "Worksheets: {Item() {return {Cells: {Clear() {}, Item() {return {Value2: null};}}};}},",
            "presentation": "Slides: {Count: 0, Add() {return {Shapes: {AddTextbox() {return {TextFrame: {TextRange: {Text: \"\"}}};}}};}},",
        }[component]
    if stage == "open":
        return f'{{DisplayAlerts: 7, {collection}: {{Open() {{throw new Error("open failed");}}}}}}'
    save = 'Save() {throw new Error("save failed");}' if stage == "save" else "Save() {}"
    close = 'Close() {throw new Error("close failed");}' if stage == "close" else "Close() {}"
    return f"{{DisplayAlerts: 7, {collection}: {{Open() {{return {{{mutation} {save}, {close}}};}}}}}}"


@pytest.mark.parametrize("stage", ["open", "mutation", "save", "close"])
@pytest.mark.parametrize(
    ("component_file", "method", "component", "operations"),
    [tuple(case[:4]) for case in GENERATION_CASES],
)
def test_generation_failure_is_typed_and_restores_alerts(
    component_file: str,
    method: str,
    component: str,
    operations: str,
    stage: str,
):
    application_source = _generation_failure_application(component, stage)
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
      params: {{stagedPath: "/staged/generated", plan: {{
        component: {json.dumps(component)}, operations: {operations}
      }}}}
    }});
    process.exit(2);
  }} catch (error) {{
    if (error.code !== "GENERATION_COMMAND_FAILED") process.exit(3);
    if (Application.DisplayAlerts !== 7) process.exit(4);
    if (error.message !== {json.dumps(stage + " failed")}) process.exit(5);
  }}
}})();
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


@pytest.mark.parametrize(
    ("component_file", "method", "component", "application_source"),
    [
        (case[0], case[1], case[2], case[4])
        for case in GENERATION_CASES
    ],
)
def test_generation_rejects_unknown_operation(
    component_file: str,
    method: str,
    component: str,
    application_source: str,
):
    path = json.dumps(str((ROOT / component_file).resolve()))
    script = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
{application_source}
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: {json.dumps(method)},
      params: {{stagedPath: "/staged/generated", plan: {{
        component: {json.dumps(component)}, operations: [{{op: "rogue.eval", args: {{}}}}]
      }}}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)
