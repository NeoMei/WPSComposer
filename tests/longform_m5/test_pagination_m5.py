from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import tempfile
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.macos_probe.models import (
    ProtocolError,
    validate_longform_generation_value,
)
from skills.WPSComposer.scripts.writer import WriterComposer


ADDIN = Path("macos/wps-jsapi-probe/addin/writer-longform-v2.js")


def _m5_map():
    return {
        "version": "M5-v1",
        "nodes": [
            {
                "nodeId": "tab:one",
                "story": "main",
                "sections": ["front_matter"],
                "pageStart": 1,
                "pageEnd": 2,
                "range": "10:150",
                "fragments": [
                    {"page": 1, "bounds": [72, 90, 540, 720]},
                    {"page": 2, "bounds": [72, 72, 540, 200]},
                ],
            }
        ],
    }


def test_generation_result_accepts_complete_m5_range_and_fragment_map():
    value = validate_longform_generation_value(
        {
            "outputPath": "/private/output.docx",
            "appliedOperations": 1,
            "issueCodes": [],
            "fieldSnapshots": [],
            "childResults": [],
            "paginationMap": _m5_map(),
        }
    )
    assert value["paginationMap"] == _m5_map()


@pytest.mark.parametrize(
    "pagination",
    [
        {**_m5_map(), "version": "M5-v2"},
        {"version": "M5-v1", "nodes": [{**_m5_map()["nodes"][0], "range": "bad"}]},
        {"version": "M5-v1", "nodes": [{**_m5_map()["nodes"][0], "pageEnd": 0}]},
        {"version": "M5-v1", "nodes": [{**_m5_map()["nodes"][0], "sections": ["../private"]}]},
        {"version": "M5-v1", "nodes": [{**_m5_map()["nodes"][0], "fragments": [{"page": 1, "bounds": None, "private": True}]}]},
    ],
)
def test_generation_result_rejects_invalid_m5_pagination(pagination):
    with pytest.raises(ProtocolError, match="generation result is invalid"):
        validate_longform_generation_value(
            {
                "outputPath": "/private/output.docx",
                "appliedOperations": 1,
                "issueCodes": [],
                "fieldSnapshots": [],
                "childResults": [],
                "paginationMap": pagination,
            }
        )


class _PointRange:
    def __init__(self, start, end):
        self.Start = start
        self.End = end

    def Information(self, kind):
        position = self.Start
        if kind == 3:
            return 1 if position < 100 else 2
        if kind == 5:
            return 72
        if kind == 6:
            return 90 if position < 100 else 180
        raise AssertionError(kind)


class _Document:
    Content = SimpleNamespace(End=201)
    PageSetup = SimpleNamespace(
        PageWidth=612,
        PageHeight=792,
        LeftMargin=72,
        RightMargin=72,
        TopMargin=72,
        BottomMargin=72,
    )

    def __init__(self):
        self.repaginate_calls = 0

    def Repaginate(self):
        self.repaginate_calls += 1

    def Range(self, start, end):
        return _PointRange(start, end)


def test_windows_writer_builds_per_page_visual_fragments_and_text_null_bounds():
    composer = SimpleNamespace(_doc=_Document())
    tracked = (
        {"nodeId": "tab:one", "op": "writer.add_semantic_table", "role": "body", "range": _PointRange(10, 150)},
        {"nodeId": "p:one", "op": "writer.add_paragraph", "role": "body", "range": _PointRange(160, 170)},
        {"nodeId": "head:one", "op": "writer.add_heading", "role": "body", "range": _PointRange(175, 180)},
    )
    result = WriterComposer.pagination_map_for_ranges(composer, tracked)
    assert result["version"] == "M5-v1"
    assert result["nodes"][0]["pageStart"] == 1
    assert result["nodes"][0]["pageEnd"] == 2
    assert [item["page"] for item in result["nodes"][0]["fragments"]] == [1, 2]
    assert all("bounds" in item for item in result["nodes"][0]["fragments"])
    assert result["nodes"][1]["fragments"] == [{"page": 2}]
    assert "bounds" in result["nodes"][2]["fragments"][0]
    assert composer._doc.repaginate_calls == 1


def test_windows_writer_omits_vendor_sentinel_visual_bounds():
    class SentinelDocument(_Document):
        def Range(self, start, end):
            point = _PointRange(start, end)
            point.Information = lambda kind: (
                1 if kind == 3 else 9_999_999
            )
            return point

    composer = SimpleNamespace(_doc=SentinelDocument())
    result = WriterComposer.pagination_map_for_ranges(
        composer,
        ({
            "nodeId": "tab:sentinel",
            "op": "writer.add_semantic_table",
            "role": "landscape",
            "range": _PointRange(10, 20),
        },),
    )
    assert result["nodes"][0]["fragments"] == [{"page": 1}]


def test_macos_addin_builds_equivalent_m5_fragments():
    source = json.dumps(str(ADDIN.resolve()))
    script = f'''
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
eval(fs.readFileSync({source}, "utf8"));
const document = {{
  Content: {{End: 201}},
  PageSetup: {{PageWidth: 612, PageHeight: 792, LeftMargin: 72, RightMargin: 72, TopMargin: 72, BottomMargin: 72}},
  Range: function(start, end) {{ return {{Start: start, End: end, Information: function(kind) {{
    if (kind === 3) return start < 100 ? 1 : 2;
    if (kind === 5) return 72;
    if (kind === 6) return start < 100 ? 90 : 180;
    throw new Error("bad kind");
  }}}}; }}
}};
const result = window.WPSComposerLongformV2.__test.buildPaginationMap(document, [
  {{nodeId: "tab:one", op: "writer.add_semantic_table", role: "body", range: {{Start: 10, End: 150}}}},
  {{nodeId: "p:one", op: "writer.add_paragraph", role: "body", range: {{Start: 160, End: 170}}}},
  {{nodeId: "head:one", op: "writer.add_heading", role: "body", range: {{Start: 175, End: 180}}}}
]);
assert.equal(result.version, "M5-v1");
assert.deepEqual(result.nodes[0].fragments.map(x => x.page), [1, 2]);
assert.ok(result.nodes[0].fragments.every(x => Array.isArray(x.bounds)));
assert.deepEqual(result.nodes[1].fragments, [{{page: 2}}]);
assert.ok(Array.isArray(result.nodes[2].fragments[0].bounds));
assert.equal(window.WPSComposerLongformV2.__test.currentSectionIsLandscape({{
  Sections: {{Count: 1, Item: function() {{ return {{PageSetup: {{Orientation: 1}}}}; }}}},
  PageSetup: {{Orientation: 0}}
}}), true);
const sectionSetup = {{Orientation: 1}};
assert.strictEqual(window.WPSComposerLongformV2.__test.currentSectionPageSetup({{
  Sections: {{Count: 1, Item: function() {{ return {{PageSetup: sectionSetup}}; }}}},
  PageSetup: {{Orientation: 0}}
}}), sectionSetup);
const widths = window.WPSComposerLongformV2.__test.contentColumnWidths([
  ["指标", "2024", "2025", "2026", "2027", "说明"],
  ["生成耗时", "45 s", "80 s", "120 s", "160 s", "保持预算内"]
], 690);
assert.equal(widths.length, 6);
assert.ok(widths.every(x => x > 0));
assert.ok(Math.abs(widths.reduce((a, b) => a + b, 0) - 690) < 0.001);
const numberingCalls = [];
const template = {{name: "decimal"}};
window.WPSComposerLongformV2.__test.linkHeadingStyles({{
  Styles: {{Item: function(index) {{ return {{LinkToListTemplate: function(value, level) {{
    numberingCalls.push([index, value, level]);
  }}}}; }}}}
}}, template);
assert.deepEqual(numberingCalls, [
  [-2, template, 1], [-3, template, 2], [-4, template, 3], [-5, template, 4]
]);
const sentinelDocument = {{
  Content: {{End: 201}},
  PageSetup: document.PageSetup,
  Range: function(start, end) {{ return {{Start: start, End: end, Information: function(kind) {{
    return kind === 3 ? 1 : 9999999;
  }}}}; }}
}};
const sentinel = window.WPSComposerLongformV2.__test.buildPaginationMap(sentinelDocument, [
  {{nodeId: "tab:sentinel", op: "writer.add_semantic_table", role: "landscape", range: {{Start: 10, End: 20}}}}
]);
assert.deepEqual(sentinel.nodes[0].fragments, [{{page: 1}}]);
'''
    descriptor, name = tempfile.mkstemp(suffix=".js")
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(script)
        subprocess.run(["node", name], check=True, capture_output=True, text=True)
    finally:
        os.unlink(name)
