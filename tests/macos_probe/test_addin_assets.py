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


WRITER_OPERATION_NAMES = [
    "writer.reset",
    "writer.configure_page",
    "writer.ensure_styles",
    "writer.add_paragraph",
    "writer.add_heading",
    "writer.add_list",
    "writer.add_table",
    "writer.add_image",
    "writer.add_page_break",
    "writer.add_section",
    "writer.add_horizontal_line",
    "writer.insert_toc",
    "writer.set_page_number",
    "writer.update_fields",
]


def test_writer_generation_uses_a_literal_complete_operation_dispatch_table():
    source = (ROOT / "writer.js").read_text()
    assert "const writerOperations = {" in source
    for operation in WRITER_OPERATION_NAMES:
        assert json.dumps(operation) + ":" in source
    assert "eval(" not in source.split("const writerOperations = {", 1)[1].split(
        "const handlers =", 1
    )[0]
    assert "new Function" not in source


def _writer_generation_stub() -> str:
    return r"""
const state = {
  text: [], breaks: [], styleNames: [], tableValues: [], imagePaths: [],
  colors: [], tableCells: {}, tableBorders: [], imageParagraph: null,
  ranges: [], tableAutoFit: [], tocRange: null,
  fieldAdds: 0, fieldUpdates: 0, tocAdds: 0, tocUpdates: 0, rangeCalls: 0
};
function font() { return {set Color(value) {state.colors.push(value);}}; }
function paragraphFormat() {
  return {Borders() { return {}; }};
}
function range(start, end) {
  state.rangeCalls += 1;
  const result = {
    Start: start, End: end, Font: font(), ParagraphFormat: paragraphFormat(),
    InsertAfter(value) { state.text.push(String(value)); document.Content.End += String(value).length; },
    InsertBreak(kind) { state.breaks.push(kind); },
    set Text(value) { state.text.push(String(value)); document.Content.End += String(value).length; },
    get Text() { return ""; }
  };
  state.ranges.push(result);
  return result;
}
function style(name) { return {Name: name, Font: font(), ParagraphFormat: paragraphFormat()}; }
const styles = {
  Item(name) { state.styleNames.push(String(name)); return style(name); },
  Add(name) { state.styleNames.push(String(name)); return style(name); }
};
const tableColumns = [{}, {}];
const tableRows = {Item() { return {Range: {Font: font()}}; }};
const table = {
  AllowAutoFit: true,
  Cell(row, column) {
    const cell = {
      Range: {Font: font(), ParagraphFormat: paragraphFormat(), set Text(value) {
        state.tableValues.push([row, column, value]);
      }},
      Shading: {}
    };
    state.tableCells[`${row},${column}`] = cell;
    return cell;
  },
  Rows: tableRows,
  Columns: {Item(index) { return tableColumns[index - 1]; }},
  Borders(index) { const border = {}; state.tableBorders.push([index, border]); return border; },
  AutoFitBehavior(value) {state.tableAutoFit.push(value);}
};
const footerRange = {Font: font(), ParagraphFormat: paragraphFormat(), Text: ""};
const document = {
  saveCalls: 0, saveAsCalls: 0, closeArgument: null,
  Content: {Text: "template", End: 0},
  PageSetup: {TextColumns: {SetCount(value) {this.Count = value;}}},
  Styles: styles,
  Range(start, end) { return range(start, end); },
  Tables: {Add(ownerRange, rows, cols) {
    assert.ok(ownerRange); assert.equal(rows, 3); assert.equal(cols, 2); return table;
  }},
  InlineShapes: {
    AddPicture(path, link, save, ownerRange) {
      assert.ok(ownerRange); state.imagePaths.push(path);
      state.imageParagraph = paragraphFormat();
      return {Width: 640, Height: 480, Range: {End: 1, ParagraphFormat: state.imageParagraph}};
    }
  },
  Sections: {Item() {return {
    Headers: {Item() {return {Range: {Text: ""}};}},
    Footers: {Item() {return {Range: footerRange};}}
  };}},
  TablesOfContents: {
    Count: 1,
    Add(ownerRange) {state.tocAdds += 1; state.tocRange = ownerRange;},
    Item() {return {Update() {state.tocUpdates += 1;}};}
  },
  Fields: {
    Add() {state.fieldAdds += 1;},
    Update() {state.fieldUpdates += 1;}
  },
  Save() {this.saveCalls += 1;},
  SaveAs() {this.saveAsCalls += 1;},
  SaveAs2() {this.saveAsCalls += 1;},
  Close(value) {this.closeArgument = value;}
};
global.Application = {
  DisplayAlerts: 7,
  get Selection() {throw new Error("generation touched another document selection");},
  Documents: {Open(path, confirm, readOnly) {
    assert.equal(path, "/staged/generated.docx");
    assert.equal(confirm, false); assert.equal(readOnly, false); return document;
  }}
};
"""


def test_writer_generation_executes_all_operations_with_document_owned_ranges():
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    operations = [
        {"op": "writer.reset", "args": {}},
        {
            "op": "writer.configure_page",
            "args": {
                "marginTop": 72,
                "marginBottom": 73,
                "marginLeft": 90,
                "marginRight": 91,
                "pageWidth": 612,
                "pageHeight": 792,
                "landscape": False,
                "columns": 2,
                "header": "Header",
                "footer": "Footer",
            },
        },
        {
            "op": "writer.ensure_styles",
            "args": {"styles": [{"name": "Body Text", "fontSize": 12, "color": "#112233"}]},
        },
        {
            "op": "writer.add_paragraph",
            "args": {
                "text": "Rich text",
                "style": "Body Text",
                "spans": [
                    {"text": "Rich", "bold": True},
                    {"text": " text", "italic": True},
                ],
            },
        },
        {"op": "writer.add_heading", "args": {"text": "Heading", "level": 2}},
        {
            "op": "writer.add_list",
            "args": {"items": ["One", "Two"], "ordered": False, "glyph": "-", "indent": 20},
        },
        {
            "op": "writer.add_table",
            "args": {
                "rows": 3,
                "cols": 2,
                "data": [["Item", "Narrative"], ["A", "short"], ["B", "A much longer narrative value"]],
                "bandedRows": True,
                "autoFit": True,
                "borderColor": "#BFBFBF",
            },
        },
        {
            "op": "writer.add_image",
            "args": {"imageId": "image-1", "width": 240, "inline": True, "preserveAspect": True},
        },
        {"op": "writer.add_page_break", "args": {}},
        {"op": "writer.add_section", "args": {}},
        {"op": "writer.add_horizontal_line", "args": {}},
        {"op": "writer.insert_toc", "args": {"title": "Contents"}},
        {"op": "writer.set_page_number", "args": {}},
        {"op": "writer.update_fields", "args": {}},
    ]
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
{_writer_generation_stub()}
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  const result = await window.WPSComposerProbe.handleCommand({{
    method: "generate_writer_document",
    params: {{
      stagedPath: "/staged/generated.docx",
      resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
      plan: {{component: "writer", operations: {json.dumps(operations)}}}
    }}
  }});
  assert.equal(result.appliedOperations, 14);
  assert.equal(document.saveCalls, 1); assert.equal(document.saveAsCalls, 0);
  assert.equal(document.closeArgument, 0); assert.equal(Application.DisplayAlerts, 7);
  assert.ok(state.rangeCalls > 0);
  assert.equal(document.PageSetup.TopMargin, 72);
  assert.equal(document.PageSetup.TextColumns.Count, 2);
  assert.ok(state.colors.includes(0x332211));
  assert.ok(state.text.join("").includes("Rich text"));
  assert.ok(state.text.join("").includes("Heading"));
  assert.ok(state.text.join("").includes("-\\tOne"));
  assert.deepEqual(state.tableValues[5], [3, 2, "A much longer narrative value"]);
  assert.equal(state.tableCells["1,1"].Range.Style.Name, "Table Header");
  assert.equal(state.tableCells["2,1"].Range.Style.Name, "Table Body");
  assert.equal(state.tableCells["2,1"].Range.Font.Name, "\u4eff\u5b8b");
  assert.equal(state.tableCells["2,1"].Range.Font.NameAscii, "Times New Roman");
  assert.equal(state.tableCells["3,1"].Shading.BackgroundPatternColor, 0xF2F2F2);
  assert.equal(state.tableCells["3,1"].TopPadding, 1.5);
  assert.equal(state.tableCells["3,1"].Range.ParagraphFormat.FirstLineIndent, 0);
  assert.equal(state.tableBorders.length, 6);
  assert.equal(state.tableBorders[0][1].Color, 0xBFBFBF);
  assert.deepEqual(state.tableAutoFit, [0]);
  assert.equal(table.AllowAutoFit, false);
  assert.equal(tableRows.HeightRule, 0);
  assert.ok(tableColumns[1].Width > tableColumns[0].Width);
  assert.ok(Math.abs(tableColumns[0].Width + tableColumns[1].Width - 431) < 0.01);
  assert.deepEqual(state.imagePaths, ["http://127.0.0.1:3889/resource-image-1.png"]);
  assert.equal(state.imageParagraph.Alignment, 1);
  assert.equal(state.imageParagraph.KeepWithNext, -1);
  assert.equal(state.imageParagraph.SpaceAfter, 0);
  assert.ok(state.text.filter(function (value) {{return value === "\\r";}}).length >= 2);
  assert.equal(state.breaks.filter(function (value) {{return value === 7;}}).length, 2);
  assert.ok(state.breaks.includes(2));
  assert.equal(state.tocAdds, 1); assert.equal(state.tocUpdates, 1);
  const tocTitle = state.ranges.find(function (value) {{
    return value.Style && value.Style.Name === "Heading 1";
  }});
  assert.equal(tocTitle.Font.Size, 18); assert.equal(tocTitle.Font.Bold, 0);
  assert.equal(tocTitle.ParagraphFormat.Alignment, 1);
  assert.equal(tocTitle.ParagraphFormat.LineSpacing, 18);
  assert.equal(tocTitle.ParagraphFormat.SpaceAfter, 5);
  assert.equal(state.fieldAdds, 1); assert.equal(state.fieldUpdates, 1);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


@pytest.mark.parametrize(
    ("preserve_aspect", "maximums", "expected_width", "expected_height"),
    [
        (True, {"maxWidth": 150, "maxHeight": 100}, 150, 75),
        (False, {}, 300, 200),
    ],
)
def test_writer_floating_image_uses_natural_size_before_applying_requested_geometry(
    preserve_aspect, maximums, expected_width, expected_height
):
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    image_args = {
        "imageId": "image-1",
        "width": 300,
        "height": 200,
        "wrap": 3,
        "inline": False,
        "preserveAspect": preserve_aspect,
        "alt": "Diagram",
        **maximums,
    }
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
const anchor = {{owner: "staged-document"}};
const shape = {{
  Width: 600, Height: 300, WrapFormat: {{}},
  Range: {{ParagraphFormat: {{}}}}
}};
let pictureArgs = null;
const document = {{
  Content: {{Text: "template", End: 1, InsertAfter() {{}}}},
  Range(start, end) {{assert.equal(start, 0); assert.equal(end, 0); return anchor;}},
  Shapes: {{AddPicture() {{pictureArgs = Array.from(arguments); return shape;}}}},
  Save() {{}}, Close() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  get Selection() {{throw new Error("global selection accessed");}},
  Documents: {{Open() {{return document;}}}}
}};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{
    method: "generate_writer_document",
    params: {{
      stagedPath: "/staged/generated.docx",
      resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
      plan: {{component: "writer", operations: [{{
        op: "writer.add_image",
        args: {json.dumps(image_args)}
      }}]}}
    }}
  }});
  assert.equal(pictureArgs.length, 8);
  assert.equal(pictureArgs[5], -1); assert.equal(pictureArgs[6], -1);
  assert.equal(pictureArgs[7], anchor);
  assert.equal(shape.LockAspectRatio, {json.dumps(-1 if preserve_aspect else 0)});
  assert.equal(shape.WrapFormat.Type, 3);
  assert.equal(shape.Width, {expected_width}); assert.equal(shape.Height, {expected_height});
  assert.equal(shape.AlternativeText, "Diagram");
  assert.equal(Application.DisplayAlerts, 7);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def test_writer_generation_rejects_missing_dispatch_handler_before_open():
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Documents: {{Open() {{
  openCalls += 1;
  return {{Content: {{Text: ""}}, Save() {{}}, Close() {{}}}};
}}}}}};
let source = fs.readFileSync({path}, "utf8");
source = source.replace('"writer.update_fields": updateWriterFields', '"writer.update_fields": null');
eval(source);
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "generate_writer_document",
      params: {{stagedPath: "/staged/generated.docx", plan: {{component: "writer", operations: [
        {{op: "writer.reset", args: {{}}}}
      ]}}}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


@pytest.mark.parametrize(
    "resources",
    [
        {},
        {"image-1": "http://127.0.0.1:3889/../secret.png"},
        {"image-1": "http://127.0.0.1:3889/%2e%2e%2fsecret.png"},
        {"image-1": "https://example.test/image.png"},
    ],
)
def test_writer_generation_rejects_missing_or_unsafe_resources_before_open(resources):
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Documents: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "generate_writer_document",
      params: {{stagedPath: "/staged/generated.docx", resources: {json.dumps(resources)},
        plan: {{component: "writer", operations: [
          {{op: "writer.add_image", args: {{imageId: "image-1"}}}}
        ]}}}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


@pytest.mark.parametrize(
    "operation",
    [
        {"op": "writer.reset", "args": {"unexpected": True}},
        {
            "op": "writer.configure_page",
            "args": {
                "marginTop": 72,
                "marginBottom": 72,
                "marginLeft": 90,
                "marginRight": 90,
                "columns": True,
            },
        },
        {"op": "writer.ensure_styles", "args": {"styles": [{"name": 7}]}},
        {
            "op": "writer.add_paragraph",
            "args": {"text": "text", "spans": [{"text": "x", "bold": "yes"}]},
        },
        {"op": "writer.add_heading", "args": {"text": "Heading", "level": 1.5}},
        {"op": "writer.add_list", "args": {"items": ["ok", {"bad": True}], "ordered": False}},
        {
            "op": "writer.add_table",
            "args": {"rows": 2, "cols": 2, "data": [["a", "b"], ["ragged"]]},
        },
        {"op": "writer.add_table", "args": {"rows": 0, "cols": 1, "data": []}},
        {"op": "writer.add_table", "args": {"rows": 1, "cols": 2, "data": [["only one"]]}},
        {
            "op": "writer.add_table",
            "args": {
                "rows": 1,
                "cols": 1,
                "data": [["x"]],
                "columnWidths": [1e308],
            },
        },
        {"op": "writer.add_image", "args": {"imageId": "image-1", "inline": "true"}},
        {"op": "writer.insert_toc", "args": {"title": 7}},
        {"op": "writer.add_list", "args": {"items": ["x" * 100_001], "ordered": False}},
    ],
)
def test_writer_generation_revalidates_operation_arguments_before_open(operation):
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    resources = (
        {"image-1": "http://127.0.0.1:3889/resource-image-1.png"}
        if operation["op"] == "writer.add_image"
        else {}
    )
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Documents: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "generate_writer_document",
      params: {{stagedPath: "/staged/generated.docx", resources: {json.dumps(resources)},
        plan: {{component: "writer", operations: [{json.dumps(operation)}]}}}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def test_writer_generation_rejects_extra_plan_fields_before_open():
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Documents: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "generate_writer_document",
      params: {{stagedPath: "/staged/generated.docx", plan: {{
        component: "writer", operations: [{{op: "writer.reset", args: {{}}}}],
        dynamicHandler: "eval"
      }}}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def test_writer_generation_enforces_utf8_plan_byte_limit_before_open():
    path = json.dumps(str((ROOT / "writer.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Documents: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({path}, "utf8"));
(async function () {{
  const item = "汉".repeat(90000);
  const plan = {{component: "writer", operations: [{{
    op: "writer.add_list", args: {{items: Array(8).fill(item), ordered: false}}
  }}]}};
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "generate_writer_document",
      params: {{stagedPath: "/staged/generated.docx", plan}}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)
