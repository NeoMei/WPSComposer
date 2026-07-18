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
const cell = {row: 1, column: 1};
const range = {Font: {}, Interior: {}, Value2: null};
const sheet = {Cells: {Clear() {}, Item(row, column) {
  assert.equal(row, 1); assert.equal(column, 1); return cell;
}}, Range(start, end) {assert.equal(start, cell); assert.equal(end, cell); return range;}};
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
        "assert.equal(document.closeArgument, false); assert.deepEqual(range.Value2, [[\"IN_PLACE_MARKER\"]]);",
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
            "spreadsheet": "Worksheets: {Item() {return {Cells: {Clear() {}, Item() {return {};}} , Range() {return {Font: {}, Interior: {}, set Value2(value) {}};}};}},",
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


SHEET_OPERATION_NAMES = [
    "sheet.reset",
    "sheet.rename",
    "sheet.add",
    "sheet.select",
    "sheet.write_table",
    "sheet.set_column_width",
    "sheet.autofit",
]


def _run_sheet_script(body: str) -> None:
    path = json.dumps(str((ROOT / "spreadsheet.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
{body}
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def test_sheet_generation_uses_a_literal_complete_operation_dispatch_table():
    source = (ROOT / "spreadsheet.js").read_text()
    assert "const sheetOperations = {" in source
    for operation in SHEET_OPERATION_NAMES:
        assert json.dumps(operation) + ":" in source
    dispatch = source.split("const sheetOperations = {", 1)[1].split(
        "const handlers =", 1
    )[0]
    assert "eval(" not in dispatch
    assert "new Function" not in source


def test_sheet_generation_executes_all_operations_on_staged_workbook_objects():
    operations = [
        {"op": "sheet.reset", "args": {}},
        {"op": "sheet.rename", "args": {"index": 1, "name": "Summary"}},
        {
            "op": "sheet.write_table",
            "args": {
                "startRow": 1,
                "startCol": 1,
                "values": [["Item", "Amount"], ["A", 10]],
                "headerBold": True,
                "headerShade": "#123456",
                "headerFontColor": "#FFFFFF",
                "fontSize": 11,
            },
        },
        {"op": "sheet.set_column_width", "args": {"column": "A", "width": 15}},
        {"op": "sheet.add", "args": {"name": "Details"}},
        {
            "op": "sheet.write_table",
            "args": {"startRow": 2, "startCol": 2, "values": [["Code"], ["X"]]},
        },
        {"op": "sheet.set_column_width", "args": {"column": "B", "width": 12}},
        {"op": "sheet.select", "args": {"index": 1}},
        {"op": "sheet.autofit", "args": {}},
    ]
    body = f"""
const state = {{deletes: 0, clears: 0, writes: [], ranges: [], widths: [], activations: [], autofits: 0}};
function makeSheet(name, owner) {{
  const sheet = {{Name: name, owner}};
  sheet.Cells = {{
    Clear() {{assert.equal(sheet.owner, "staged"); state.clears += 1;}},
    Item(row, column) {{return {{row, column, owner: sheet.owner}};}}
  }};
  sheet.Range = function (start, end) {{
    assert.equal(start.owner, "staged"); assert.equal(end.owner, "staged");
    const range = {{Font: {{}}, Interior: {{}}, start, end}};
    Object.defineProperty(range, "Value2", {{set(value) {{state.writes.push(value);}}}});
    state.ranges.push(range); return range;
  }};
  sheet.Columns = {{Item(column) {{
    const target = {{}};
    Object.defineProperty(target, "ColumnWidth", {{set(value) {{state.widths.push([sheet.Name, column, value]);}}}});
    return target;
  }}}};
  sheet.UsedRange = {{Columns: {{AutoFit() {{assert.equal(sheet.owner, "staged"); state.autofits += 1;}}}}}};
  sheet.Activate = function () {{state.activations.push(sheet.Name);}};
  return sheet;
}}
const sheets = [makeSheet("Template", "staged"), makeSheet("Extra", "staged"), makeSheet("Third", "staged")];
const worksheets = {{
  get Count() {{return sheets.length;}},
  Item(index) {{assert.ok(index >= 1 && index <= sheets.length); return sheets[index - 1];}},
  Add(before, after) {{assert.equal(before, null); assert.equal(after, sheets[sheets.length - 1]); const sheet = makeSheet("New", "staged"); sheets.push(sheet); return sheet;}}
}};
sheets.forEach(function (sheet) {{sheet.Delete = function () {{assert.equal(sheet.owner, "staged"); sheets.splice(sheets.indexOf(sheet), 1); state.deletes += 1;}};}});
const otherSheets = [makeSheet("User workbook", "other")];
const workbook = {{Worksheets: worksheets, saveCalls: 0, closeArgument: null,
  Save() {{this.saveCalls += 1;}}, Close(value) {{this.closeArgument = value;}}}};
global.Application = {{DisplayAlerts: 7,
  get ActiveWorkbook() {{throw new Error("global workbook access");}},
  get ActiveSheet() {{throw new Error("global sheet access");}},
  Workbooks: {{Open() {{return workbook;}}}}
}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  const result = await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: {json.dumps(operations)}}}
  }}}});
  assert.equal(result.appliedOperations, 9); assert.equal(workbook.saveCalls, 1);
  assert.equal(workbook.closeArgument, false); assert.equal(Application.DisplayAlerts, 7);
  assert.equal(state.deletes, 2); assert.equal(state.clears, 1); assert.equal(sheets.length, 2);
  assert.equal(otherSheets.length, 1); assert.equal(otherSheets[0].Name, "User workbook");
  assert.deepEqual(state.writes, [
    [["Item", "Amount"], ["A", 10]], [["Code"], ["X"]]
  ]);
  assert.equal(state.ranges[0].Font.Size, 11); assert.equal(state.ranges[1].Font.Bold, true);
  assert.equal(state.ranges[1].Font.Color, 0xFFFFFF); assert.equal(state.ranges[1].Interior.Color, 0x563412);
  assert.deepEqual(state.widths, [["Summary", "A", 15], ["Details", "B", 12]]);
  assert.deepEqual(state.activations, ["Summary"]); assert.equal(state.autofits, 1);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_reset_keeps_one_staged_worksheet_and_ignores_other_workbooks():
    body = f"""
let deleteCalls = 0; let clearCalls = 0; let otherDeleteCalls = 0;
const sheet = {{Cells: {{Clear() {{clearCalls += 1;}}}}, Delete() {{deleteCalls += 1;}}}};
const workbook = {{Worksheets: {{Count: 1, Item(index) {{assert.equal(index, 1); return sheet;}}}}, Save() {{}}, Close() {{}}}};
const otherWorkbook = {{Worksheets: {{Count: 2, Item() {{return {{Delete() {{otherDeleteCalls += 1;}}}};}}}}}};
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{return workbook;}}}},
  get ActiveWorkbook() {{return otherWorkbook;}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [{{op: "sheet.reset", args: {{}}}}]}}
  }}}});
  assert.equal(deleteCalls, 0); assert.equal(clearCalls, 1); assert.equal(otherDeleteCalls, 0);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_generation_sanitizes_and_uniquifies_names_deterministically():
    long_name = "A" * 40
    operations = [
        {"op": "sheet.reset", "args": {}},
        {"op": "sheet.rename", "args": {"index": 1, "name": "   "}},
        {"op": "sheet.add", "args": {"name": "'"}},
        {"op": "sheet.add", "args": {"name": "Bad[]:*?/\\Name"}},
        {"op": "sheet.add", "args": {"name": long_name}},
        {"op": "sheet.add", "args": {"name": long_name}},
    ]
    body = f"""
const sheets = [];
function makeSheet(name) {{return {{Name: name, Cells: {{Clear() {{}}}}, Delete() {{sheets.splice(sheets.indexOf(this), 1);}}}};}}
sheets.push(makeSheet("Template"));
const workbook = {{Worksheets: {{
  get Count() {{return sheets.length;}}, Item(index) {{return sheets[index - 1];}},
  Add(before, after) {{const sheet = makeSheet("New"); sheets.push(sheet); return sheet;}}
}}, Save() {{}}, Close() {{}}}};
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{return workbook;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: {json.dumps(operations)}}}
  }}}});
  assert.deepEqual(sheets.map(function (sheet) {{return sheet.Name;}}), [
    "Sheet", "Sheet (2)", "Bad_______Name", {json.dumps('A' * 31)}, {json.dumps('A' * 27 + ' (2)')}
  ]);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_generation_strips_apostrophe_exposed_by_name_truncation():
    requested = "A" * 30 + "'B"
    body = f"""
let openCalls = 0;
const sheet = {{Name: "Template", Cells: {{Clear() {{}}}}, Delete() {{}}}};
const workbook = {{Worksheets: {{Count: 1, Item() {{return sheet;}}}}, Save() {{}}, Close() {{}}}};
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1; return workbook;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [
      {{op: "sheet.reset", args: {{}}}},
      {{op: "sheet.rename", args: {{index: 1, name: {json.dumps(requested)}}}}}
    ]}}
  }}}});
  assert.equal(openCalls, 1); assert.equal(sheet.Name, {json.dumps('A' * 30)});
  assert.equal(Application.DisplayAlerts, 7);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_generation_allows_object_prototype_name_when_not_used():
    body = f"""
const sheet = {{Name: "Template", Cells: {{Clear() {{}}}}, Delete() {{}}}};
const workbook = {{Worksheets: {{Count: 1, Item() {{return sheet;}}}}, Save() {{}}, Close() {{}}}};
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{return workbook;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [
      {{op: "sheet.reset", args: {{}}}},
      {{op: "sheet.rename", args: {{index: 1, name: "constructor"}}}}
    ]}}
  }}}});
  assert.equal(sheet.Name, "constructor"); assert.equal(Application.DisplayAlerts, 7);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize(
    "operation",
    [
        '{op: "sheet.reset", args: {unexpected: true}}',
        '{op: "sheet.rename", args: {index: 0, name: "Summary"}}',
        '{op: "sheet.rename", args: {index: 1, name: 7}}',
        '{op: "sheet.add", args: {name: {dynamic: true}}}',
        '{op: "sheet.select", args: {index: 1.5}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: [["a", "b"], ["ragged"]]}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: []}}',
        '{op: "sheet.write_table", args: {startRow: 0, startCol: 1, values: [["x"]]}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: [[Number.MAX_SAFE_INTEGER + 1]]}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: [[NaN]]}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: [["x"]], headerBold: "yes"}}',
        '{op: "sheet.write_table", args: {startRow: 1, startCol: 1, values: [["x"]], rogue: true}}',
        '{op: "sheet.set_column_width", args: {column: 7, width: 15}}',
        '{op: "sheet.set_column_width", args: {column: "A", width: Infinity}}',
        '{op: "sheet.autofit", args: {unexpected: true}}',
        '{op: "rogue.eval", args: {}}',
        '{op: "sheet.reset", args: {}, handler: "dynamic"}',
    ],
)
def test_sheet_generation_revalidates_operations_before_open(operation):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
      stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [{operation}]}}
    }}}});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize("operation_name", ["toString", "constructor", "valueOf"])
def test_sheet_generation_rejects_object_prototype_operation_before_open(
    operation_name,
):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
      stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [
        {{op: {json.dumps(operation_name)}, args: {{}}}}
      ]}}
    }}}});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize(
    "operations",
    [
        [
            {"op": "sheet.reset", "args": {}},
            {"op": "sheet.rename", "args": {"index": 2, "name": "Missing"}},
        ],
        [
            {"op": "sheet.reset", "args": {}},
            {"op": "sheet.add", "args": {"name": "Second"}},
            {"op": "sheet.select", "args": {"index": 3}},
        ],
        [
            {"op": "sheet.reset", "args": {}},
            {"op": "sheet.add", "args": {"name": "Second"}},
            {"op": "sheet.rename", "args": {"index": 3, "name": "Missing"}},
        ],
    ],
)
def test_sheet_generation_rejects_out_of_sequence_sheet_index_before_open(
    operations,
):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
      stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: {json.dumps(operations)}}}
    }}}});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize(
    "args",
    [
        {
            "startRow": 2**53 - 1,
            "startCol": 1,
            "values": [["a"], ["b"], ["c"]],
        },
        {
            "startRow": 1,
            "startCol": 2**53 - 1,
            "values": [["a", "b", "c"]],
        },
    ],
)
def test_sheet_generation_rejects_unsafe_derived_table_range_before_open(args):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
      stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [
        {{op: "sheet.write_table", args: {json.dumps(args)}}}
      ]}}
    }}}});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, "OPERATION_PLAN_INVALID");
    assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize(
    "plan",
    [
        '{component: "spreadsheet", operations: [{op: "sheet.reset", args: {}}], dynamicHandler: "eval"}',
        '{component: "writer", operations: [{op: "sheet.reset", args: {}}]}',
        '{component: "spreadsheet", operations: []}',
    ],
)
def test_sheet_generation_rejects_malformed_plan_before_open(plan):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{stagedPath: "/staged/generated.xlsx", plan: {plan}}}}});
    process.exit(2);
  }} catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_generation_enforces_string_and_utf8_plan_limits_before_open():
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
async function reject(plan) {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{stagedPath: "/staged/generated.xlsx", plan}}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}}
(async function () {{
  await reject({{component: "spreadsheet", operations: [{{op: "sheet.add", args: {{name: "x".repeat(100001)}}}}]}});
  const value = "汉".repeat(90000);
  await reject({{component: "spreadsheet", operations: [{{op: "sheet.write_table", args: {{startRow: 1, startCol: 1, values: Array(8).fill(0).map(function () {{return [value];}})}}}}]}});
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


@pytest.mark.parametrize(
    ("original", "replacement"),
    [
        ('"sheet.autofit": autofitSheet', '"sheet.autofit": null'),
        (
            '"sheet.autofit": {required: [], allowed: []}',
            '"sheet.autofit-missing": {required: [], allowed: []}',
        ),
    ],
)
def test_sheet_generation_rejects_incomplete_catalog_before_open(
    original, replacement
):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
let source = fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8");
source = source.replace({json.dumps(original)}, {json.dumps(replacement)});
eval(source);
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{
    stagedPath: "/staged/generated.xlsx", plan: {{component: "spreadsheet", operations: [{{op: "sheet.reset", args: {{}}}}]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


def test_sheet_generation_enforces_operation_and_table_cell_limits_before_open():
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Workbooks: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'spreadsheet.js').resolve()))}, "utf8"));
async function reject(plan) {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_spreadsheet_workbook", params: {{stagedPath: "/staged/generated.xlsx", plan}}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}}
(async function () {{
  await reject({{component: "spreadsheet", operations: Array(10001).fill(0).map(function () {{return {{op: "sheet.reset", args: {{}}}};}})}});
  const row = Array(101).fill("x");
  await reject({{component: "spreadsheet", operations: [{{op: "sheet.write_table", args: {{startRow: 1, startCol: 1, values: Array(100).fill(row)}}}}]}});
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_sheet_script(body)


SLIDE_OPERATION_NAMES = [
    "slide.reset",
    "slide.set_size",
    "slide.apply_preset",
    "slide.add_title",
    "slide.add_section",
    "slide.add_bullets",
    "slide.add_blank",
    "slide.add_image",
    "slide.add_table",
]


def _run_presentation_script(body: str):
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
{body}
"""
    subprocess.run(["node", "-e", script], check=True, capture_output=True, text=True)


def test_presentation_generation_uses_a_literal_complete_operation_dispatch_table():
    source = (ROOT / "presentation.js").read_text()
    assert "const slideOperations = {" in source
    for operation in SLIDE_OPERATION_NAMES:
        assert json.dumps(operation) + ":" in source
    dispatch = source.split("const slideOperations = {", 1)[1].split(
        "const handlers =", 1
    )[0]
    assert "eval(" not in dispatch
    assert "new Function" not in source


def test_presentation_generation_executes_all_operations_on_staged_presentation():
    preset = {
        "name": "Academic",
        "colors": {
            "primary": "#1A3C8B",
            "secondary": "#E67733",
            "accent": "#188050",
            "dark": "#222222",
            "light": "#F5F8FC",
            "background": "#FFFFFF",
        },
        "fonts": {
            "title": {"family": "Arial", "size": 40, "color": "#1A3C8B"},
            "subtitle": {"family": "Arial", "size": 22, "color": "#646464"},
            "body": {"family": "Arial", "size": 24, "color": "#222222"},
            "caption": {"family": "Arial", "size": 16, "color": "#808080"},
            "chinese": {"family": "Microsoft YaHei", "size": 24, "color": "#222222"},
        },
        "spacing": {"margin": 80, "gap": 24, "cardPadding": 20, "lineHeight": 1.5},
    }
    operations = [
        {"op": "slide.reset", "args": {}},
        {"op": "slide.set_size", "args": {"width": 960, "height": 540}},
        {"op": "slide.apply_preset", "args": {"preset": preset}},
        {
            "op": "slide.add_title",
            "args": {
                "title": "Quarterly review",
                "subtitle": "Ada | 2026-07-19",
                "titleSize": 40,
                "subtitleSize": 20,
                "titleColor": "#1A3C8B",
            },
        },
        {"op": "slide.add_section", "args": {"title": "Results"}},
        {
            "op": "slide.add_bullets",
            "args": {"title": "Highlights", "items": ["One", "Two"], "titleSize": 28, "bodySize": 18},
        },
        {"op": "slide.add_blank", "args": {}},
        {
            "op": "slide.add_image",
            "args": {"slide": 4, "imageId": "image-1", "left": 80, "top": 100, "width": 800, "height": 400},
        },
        {"op": "slide.add_blank", "args": {}},
        {
            "op": "slide.add_table",
            "args": {
                "slide": 5, "rows": 2, "cols": 2, "left": 60, "top": 120,
                "width": 840, "height": 380, "data": [["Item", "Amount"], ["A", 10]],
                "headerShade": "#4472C4", "headerFont": "#FFFFFF", "fontSize": 14,
            },
        },
    ]
    body = f"""
const state = {{deleted: 0, textboxes: [], pictures: [], tables: []}};
function textRange() {{return {{Text: "", Font: {{}}, ParagraphFormat: {{}}}};}}
function makeSlide(owner) {{
  const slide = {{owner, FollowMasterBackground: true, Background: {{Fill: {{Solid() {{}}}}}}}};
  slide.Shapes = {{
    AddTextbox(orientation, left, top, width, height) {{
      assert.equal(slide.owner, "staged");
      const shape = {{kind: "textbox", orientation, left, top, width, height, TextFrame: {{TextRange: textRange()}}}};
      state.textboxes.push(shape); return shape;
    }},
    AddPicture(path, link, save, left, top) {{
      assert.equal(slide.owner, "staged");
      const shape = {{kind: "picture", path, link, save, Left: left, Top: top, Width: 1600, Height: 800}};
      state.pictures.push(shape); return shape;
    }},
    AddTable(rows, cols, left, top, width, height) {{
      assert.equal(slide.owner, "staged");
      const cells = {{}};
      const shape = {{kind: "table", rows, cols, left, top, width, height, Table: {{Cell(row, col) {{
        const key = `${{row}},${{col}}`;
        if (!cells[key]) cells[key] = {{Shape: {{TextFrame: {{TextRange: textRange()}}, Fill: {{}}}}}};
        return cells[key];
      }}}}, cells}};
      state.tables.push(shape); return shape;
    }}
  }};
  return slide;
}}
const slides = [makeSlide("staged"), makeSlide("staged")];
slides.forEach(function (slide) {{slide.Delete = function () {{slides.splice(slides.indexOf(slide), 1); state.deleted += 1;}};}});
const collection = {{
  get Count() {{return slides.length;}},
  Item(index) {{return slides[index - 1];}},
  Add(index, layout) {{assert.equal(index, slides.length + 1); const slide = makeSlide("staged"); slides.push(slide); return slide;}}
}};
const presentation = {{
  Slides: collection, PageSetup: {{}}, SlideMaster: {{Background: {{Fill: {{Solid() {{}}}}}}}},
  saveCalls: 0, closeCalls: 0, Save() {{this.saveCalls += 1;}}, Close() {{this.closeCalls += 1;}}
}};
const otherPresentation = {{Slides: [makeSlide("other")]}};
global.Application = {{DisplayAlerts: 7,
  get ActivePresentation() {{throw new Error("global presentation access");}},
  Presentations: {{Open() {{return presentation;}}}}
}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  const result = await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx",
    resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
    plan: {{component: "presentation", operations: {json.dumps(operations)}}}
  }}}});
  assert.equal(result.appliedOperations, 10); assert.equal(presentation.saveCalls, 1);
  assert.equal(presentation.closeCalls, 1); assert.equal(Application.DisplayAlerts, 7);
  assert.equal(state.deleted, 2); assert.equal(slides.length, 5);
  assert.equal(presentation.PageSetup.SlideWidth, 960); assert.equal(presentation.PageSetup.SlideHeight, 540);
  assert.equal(state.textboxes[0].TextFrame.TextRange.Text, "Quarterly review");
  assert.equal(state.textboxes[0].TextFrame.TextRange.Font.Name, "Arial");
  assert.equal(state.textboxes[0].TextFrame.TextRange.Font.Color.RGB, 0x8B3C1A);
  assert.ok(state.textboxes.some(function (shape) {{return shape.TextFrame.TextRange.Text === "One\\rTwo";}}));
  assert.equal(state.pictures[0].path, "http://127.0.0.1:3889/resource-image-1.png");
  assert.equal(state.pictures[0].Width, 800); assert.equal(state.pictures[0].Height, 400);
  assert.equal(state.pictures[0].Left, 80); assert.equal(state.pictures[0].Top, 100);
  assert.equal(state.tables[0].cells["1,1"].Shape.TextFrame.TextRange.Text, "Item");
  assert.equal(state.tables[0].cells["1,1"].Shape.TextFrame.TextRange.Font.Bold, true);
  assert.equal(state.tables[0].cells["1,1"].Shape.Fill.ForeColor.RGB, 0xC47244);
  assert.equal(otherPresentation.Slides.length, 1);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize(
    "operation",
    [
        '{op: "slide.reset", args: {unexpected: true}}',
        '{op: "slide.set_size", args: {width: 960}}',
        '{op: "slide.set_size", args: {width: true, height: 540}}',
        '{op: "slide.apply_preset", args: {preset: {name: "x", colors: {primary: "#000000", dark: "#000000", background: "#FFFFFF"}, fonts: {title: {family: "Arial", size: 40, color: "#000000"}}}}}',
        '{op: "slide.add_title", args: {title: 7}}',
        '{op: "slide.add_section", args: {title: "x", rogue: true}}',
        '{op: "slide.add_bullets", args: {title: "x", items: ["ok", 7]}}',
        '{op: "slide.add_blank", args: {dynamic: true}}',
        '{op: "slide.add_image", args: {slide: 1, imageId: "image-1", left: 0, top: 0, source: "/tmp/x"}}',
        '{op: "slide.add_table", args: {slide: 1, rows: 2, cols: 2, left: 0, top: 0, width: 100, height: 100, data: [["x"]]}}',
        '{op: "rogue.eval", args: {}}',
        '{op: "slide.reset", args: {}, handler: "dynamic"}',
    ],
)
def test_presentation_generation_revalidates_closed_schemas_before_open(operation):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
    plan: {{component: "presentation", operations: [{operation}]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0); assert.equal(Application.DisplayAlerts, 7);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize("operation_name", ["toString", "constructor", "valueOf"])
def test_presentation_generation_rejects_prototype_operation_before_open(operation_name):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", plan: {{component: "presentation", operations: [{{op: {json.dumps(operation_name)}, args: {{}}}}]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize(
    "operations",
    [
        [{"op": "slide.add_image", "args": {"slide": 1, "imageId": "image-1", "left": 0, "top": 0}}],
        [
            {"op": "slide.reset", "args": {}},
            {"op": "slide.add_title", "args": {"title": "Only"}},
            {"op": "slide.add_table", "args": {"slide": 2, "rows": 1, "cols": 1, "left": 0, "top": 0, "width": 10, "height": 10, "data": [["x"]]}},
        ],
        [
            {"op": "slide.add_blank", "args": {}},
            {"op": "slide.reset", "args": {}},
            {"op": "slide.add_image", "args": {"slide": 1, "imageId": "image-1", "left": 0, "top": 0}},
        ],
    ],
)
def test_presentation_generation_rejects_out_of_sequence_slide_state_before_open(operations):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
    plan: {{component: "presentation", operations: {json.dumps(operations)}}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize(
    "operation",
    [
        '{op: "slide.set_size", args: {width: 0, height: 540}}',
        '{op: "slide.add_image", args: {slide: 1, imageId: "image-1", left: -1, top: 0}}',
        '{op: "slide.add_image", args: {slide: 1, imageId: "image-1", left: 300, top: 0, width: 200, height: 20}}',
        '{op: "slide.add_image", args: {slide: 1, imageId: "image-1", left: 0, top: 0, width: Infinity, height: 20}}',
        '{op: "slide.add_table", args: {slide: 1, rows: 1, cols: 1, left: 0, top: 150, width: 10, height: 100, data: [["x"]]}}',
        '{op: "slide.add_title", args: {title: "Too narrow"}}',
        '{op: "slide.add_bullets", args: {title: "Too short", items: ["x"]}}',
    ],
)
def test_presentation_generation_rejects_unsafe_geometry_before_open(operation):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  const testedOperation = {operation};
  const resources = testedOperation.op === "slide.add_image"
    ? {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}} : {{}};
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", resources,
    plan: {{component: "presentation", operations: [
      {{op: "slide.set_size", args: {{width: 100, height: 50}}}},
      {{op: "slide.add_blank", args: {{}}}}, testedOperation
    ]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize(
    "resources",
    [
        {},
        {"image-1": "file:///tmp/x.png"},
        {"image-1": "http://127.0.0.1:3889/a/b.png"},
        {"image-1": "http://127.0.0.1:3889/x.png", "unused": "http://127.0.0.1:3889/y.png"},
    ],
)
def test_presentation_generation_rejects_missing_or_unsafe_resources_before_open(resources):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", resources: {json.dumps(resources)},
    plan: {{component: "presentation", operations: [
      {{op: "slide.add_blank", args: {{}}}},
      {{op: "slide.add_image", args: {{slide: 1, imageId: "image-1", left: 0, top: 0}}}}
    ]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


def test_presentation_generation_accepts_safe_fractional_geometry():
    body = f"""
let slideCount = 0; let openCalls = 0; let stagedSlide = null;
const picture = {{Width: 1600, Height: 800}};
const slides = {{
  get Count() {{return slideCount;}},
  Add() {{slideCount += 1; stagedSlide = {{Shapes: {{AddPicture() {{return picture;}}}}}}; return stagedSlide;}},
  Item() {{return stagedSlide;}}
}};
const presentation = {{Slides: slides, PageSetup: {{}}, Save() {{}}, Close() {{}}}};
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1; return presentation;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx",
    resources: {{"image-1": "http://127.0.0.1:3889/resource-image-1.png"}},
    plan: {{component: "presentation", operations: [
      {{op: "slide.set_size", args: {{width: 960.1, height: 540.1}}}},
      {{op: "slide.add_blank", args: {{}}}},
      {{op: "slide.add_image", args: {{slide: 1, imageId: "image-1", left: 80.25, top: 100.25, width: 800.25, height: 400.25}}}}
    ]}}
  }}}});
  assert.equal(openCalls, 1); assert.equal(picture.Left, 80.25); assert.equal(picture.Top, 100.25);
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


def test_presentation_generation_rejects_non_object_resources_before_open():
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", resources: [],
    plan: {{component: "presentation", operations: [{{op: "slide.reset", args: {{}}}}]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


def test_presentation_generation_enforces_recursive_utf8_and_collection_limits_before_open():
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
eval(fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8"));
async function reject(plan) {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{stagedPath: "/staged/generated.pptx", plan}}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}}
(async function () {{
  await reject({{component: "presentation", operations: [{{op: "slide.add_title", args: {{title: "x".repeat(100001)}}}}]}});
  let nested = "x"; for (let i = 0; i < 65; i += 1) nested = [nested];
  await reject({{component: "presentation", operations: [{{op: "slide.add_bullets", args: {{title: "x", items: nested}}}}]}});
  const wide = "汉".repeat(90000);
  await reject({{component: "presentation", operations: Array(8).fill(0).map(function () {{return {{op: "slide.add_title", args: {{title: wide}}}};}})}});
  await reject({{component: "presentation", operations: Array(10001).fill(0).map(function () {{return {{op: "slide.reset", args: {{}}}};}})}});
  const row = Array(101).fill("x");
  await reject({{component: "presentation", operations: [{{op: "slide.add_blank", args: {{}}}}, {{op: "slide.add_table", args: {{slide: 1, rows: 100, cols: 101, left: 0, top: 0, width: 100, height: 100, data: Array(100).fill(row)}}}}]}});
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)


@pytest.mark.parametrize(
    ("original", "replacement"),
    [
        ('"slide.add_blank": addBlankSlide', '"slide.add_blank": null'),
        (
            '"slide.add_blank": {required: [], allowed: []}',
            '"slide.add_blank-missing": {required: [], allowed: []}',
        ),
    ],
)
def test_presentation_generation_rejects_incomplete_catalog_before_open(original, replacement):
    body = f"""
let openCalls = 0;
global.Application = {{DisplayAlerts: 7, Presentations: {{Open() {{openCalls += 1;}}}}}};
let source = fs.readFileSync({json.dumps(str((ROOT / 'presentation.js').resolve()))}, "utf8");
source = source.replace({json.dumps(original)}, {json.dumps(replacement)});
eval(source);
(async function () {{
  try {{await window.WPSComposerProbe.handleCommand({{method: "generate_presentation_deck", params: {{
    stagedPath: "/staged/generated.pptx", plan: {{component: "presentation", operations: [{{op: "slide.reset", args: {{}}}}]}}
  }}}}); process.exit(2);}}
  catch (error) {{assert.equal(error.code, "OPERATION_PLAN_INVALID"); assert.equal(openCalls, 0);}}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""
    _run_presentation_script(body)
