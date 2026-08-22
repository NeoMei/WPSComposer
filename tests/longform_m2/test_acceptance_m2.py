"""M2 acceptance fixtures and structural snapshots.

Tests exercise the six required acceptance categories restricted to M2 scope:
page skeleton/numbering/headers/TOC/headings only.  Recording-executor tests
compare deterministic structural snapshots of the generated plan and run the
plan through mocked COM/JSAPI primitives to verify header/footer/page-number
formatting.  macOS real-WPS tests inspect the unzipped DOCX OOXML and skip
cleanly when the bridge is unavailable.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
import types
import zipfile
from pathlib import Path
from typing import Any

import pytest

from skills.WPSComposer.scripts.longform.executor import (
    RecordingLongformExecutor,
    finalize_fields_with_convergence,
)
from skills.WPSComposer.scripts.longform.pipeline import (
    LongformBuild,
    _aggregate_issues,
    build_longform_generation,
)
from skills.WPSComposer.scripts.longform.plan import build_longform_plan
from skills.WPSComposer.scripts.longform.resources import preflight_resources
from skills.WPSComposer.scripts.longform.semantic import normalize_longform_document
from skills.WPSComposer.scripts.md_parser import parse_markdown


FIXTURES_DIR = Path(__file__).parent / "fixtures"
SNAPSHOTS_DIR = Path(__file__).parent / "snapshots"

_FIXTURE_NAMES = (
    "academic",
    "wide_figure",
    "hybrid_bid",
    "degradation",
    "plain_short",
    "protocol_edge",
)


# ---------------------------------------------------------------------------
# Fixture loading / plan building
# ---------------------------------------------------------------------------


def _load_fixture(name: str) -> str:
    return (FIXTURES_DIR / f"{name}.md").read_text(encoding="utf-8")


def build_fixture(name: str) -> LongformBuild:
    """Build a LongformBuild for an acceptance fixture."""
    md = _load_fixture(name)
    return build_longform_generation(md, base_dir=str(FIXTURES_DIR))


def _reduce_op(op: dict[str, Any]) -> dict[str, Any]:
    """Reduce a plan operation to its M2-structurally relevant fields."""
    reducers = {
        "writer.configure_front_matter": lambda a: {
            "title": a.get("title"),
            "author": a.get("author"),
            "date": a.get("date"),
            "header": a.get("header"),
            "titlePage": a.get("titlePage"),
            "shortTitle": a.get("shortTitle"),
        },
        "writer.configure_toc_styles": lambda a: {
            "tocTitle": a.get("tocTitle"),
            "minFontSizePt": a.get("minFontSizePt"),
            "minSpaceBeforePt": a.get("minSpaceBeforePt"),
            "minSpaceAfterPt": a.get("minSpaceAfterPt"),
        },
        "writer.configure_section": lambda a: {
            k: a.get(k)
            for k in (
                "role",
                "pageNumberFormat",
                "restartPageNumbering",
                "startPageNumber",
                "headerText",
                "footerText",
                "linkToPreviousHeader",
                "linkToPreviousFooter",
                "landscape",
            )
        },
        "writer.insert_toc": lambda a: {"title": a.get("title")},
        "writer.insert_figure_index": lambda a: {"title": a.get("title")},
        "writer.insert_table_index": lambda a: {"title": a.get("title")},
        "writer.add_heading": lambda a: {
            k: a.get(k)
            for k in ("text", "level", "numbering", "numberingScheme")
        },
        "writer.add_paragraph": lambda a: {
            k: a.get(k) for k in ("text", "style")
        },
        "writer.add_document_quality_notice": lambda a: {
            "notices": [
                {
                    "code": n.get("code"),
                    "placement": n.get("placement", "document"),
                }
                for n in a.get("notices", [])
            ]
        },
        "writer.finalize_fields": lambda a: {"maxRounds": a.get("maxRounds")},
    }
    fn = reducers.get(op["op"])
    return {"op": op["op"], "args": fn(op["args"]) if fn else {}}


def structural_snapshot(plan) -> list[dict[str, Any]]:
    return [_reduce_op(op) for op in plan.to_dict()["operations"]]


@pytest.fixture(params=_FIXTURE_NAMES)
def fixture_name(request: pytest.FixtureRequest) -> str:
    return request.param


@pytest.fixture
def fixture_build(fixture_name: str) -> LongformBuild:
    return build_fixture(fixture_name)


# ---------------------------------------------------------------------------
# Recording executor / plan snapshot tests
# ---------------------------------------------------------------------------


class TestRecordingStructuralSnapshots:
    """M2 Task 8: deterministic structural snapshots against the recording executor."""

    def test_fixture_matches_structural_snapshot(
        self, fixture_name: str, fixture_build: LongformBuild
    ) -> None:
        snapshot_path = SNAPSHOTS_DIR / f"{fixture_name}_ops.json"
        actual = structural_snapshot(fixture_build.plan)
        expected = json.loads(snapshot_path.read_text(encoding="utf-8"))
        assert actual == expected

    def test_fixture_executes_on_recording_executor(
        self, fixture_build: LongformBuild
    ) -> None:
        executor = RecordingLongformExecutor(artifact="staged.docx")
        outcome = executor.execute(
            fixture_build.plan, tuple(fixture_build.preflight.resources)
        )
        assert outcome.staged_artifact == "staged.docx"
        assert outcome.pagination_map.version == "M2-stub"

    def test_field_convergence_is_stable_on_recording_executor(
        self, fixture_build: LongformBuild
    ) -> None:
        executor = RecordingLongformExecutor()
        result = finalize_fields_with_convergence(executor, max_rounds=3)
        assert result.rounds == 2
        assert result.issues == ()


# ---------------------------------------------------------------------------
# Page-skeleton structure assertions at plan level
# ---------------------------------------------------------------------------


class TestPageSkeletonStructure:
    """Assert the M2 page-skeleton structure directly from the plan."""

    def _sections(self, build: LongformBuild) -> list[dict[str, Any]]:
        return [
            op
            for op in build.plan.to_dict()["operations"]
            if op["op"] == "writer.configure_section"
        ]

    def _first_body_section(self, build: LongformBuild) -> dict[str, Any]:
        for op in self._sections(build):
            if op["args"].get("role") == "body":
                return op
        pytest.fail("expected a body section")

    def test_cover_page_has_no_page_number(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        cover = [
            op for op in self._sections(fixture_build) if op["args"].get("role") == "cover"
        ]
        if not cover:
            if fixture_name in {"plain_short"}:
                return
            pytest.fail(f"expected cover section for {fixture_name}")
        assert len(cover) == 1
        assert cover[0]["args"]["pageNumberFormat"] == "none"
        assert cover[0]["args"].get("headerText") == ""
        assert cover[0]["args"].get("footerText") == ""

    def test_front_matter_roman_before_body(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        ops = fixture_build.plan.to_dict()["operations"]
        front = [
            op for op in self._sections(fixture_build) if op["args"].get("role") == "front_matter"
        ]
        body = [
            op for op in self._sections(fixture_build) if op["args"].get("role") == "body"
        ]
        if front:
            assert front[0]["args"]["pageNumberFormat"] == "roman"
            assert front[0]["args"]["restartPageNumbering"] is True
            assert front[0]["args"]["startPageNumber"] == 1
            assert front[0]["args"].get("headerText") == ""
            assert front[0]["args"].get("footerText") == ""
        if front and body:
            front_idx = next(i for i, op in enumerate(ops) if op == front[0])
            body_idx = next(i for i, op in enumerate(ops) if op == body[0])
            assert front_idx < body_idx

    def test_body_restarts_arabic(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        body = [
            op for op in self._sections(fixture_build) if op["args"].get("role") == "body"
        ]
        if not body:
            if fixture_name == "protocol_edge":
                pytest.fail("protocol_edge fixture must have a body section")
            return
        assert body[0]["args"]["pageNumberFormat"] == "arabic"
        assert body[0]["args"]["restartPageNumbering"] is True
        assert body[0]["args"]["startPageNumber"] == 1

    def test_centered_header_and_footer(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        """Body sections carry a non-empty header text and an empty footer text.

        The actual centering, bottom border, and page-number field are applied
        by the COM/JSAPI primitive layer; those properties are asserted in
        TestAddInM2Primitives.
        """
        body = [
            op for op in self._sections(fixture_build) if op["args"].get("role") == "body"
        ]
        if not body:
            return
        assert body[0]["args"].get("headerText") != ""
        assert body[0]["args"].get("footerText") == ""

    def test_compact_toc_styles_present(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        ops = fixture_build.plan.to_dict()["operations"]
        toc_style_ops = [op for op in ops if op["op"] == "writer.configure_toc_styles"]
        insert_toc_ops = [op for op in ops if op["op"] == "writer.insert_toc"]
        assert len(toc_style_ops) == 1
        args = toc_style_ops[0]["args"]
        assert args["minFontSizePt"]["toc1"] == 10.5
        assert args["minFontSizePt"]["toc2"] == 10.0
        assert args["minFontSizePt"]["toc3"] == 10.0
        assert args["minSpaceBeforePt"]["toc1"] == 0.0
        assert args["minSpaceAfterPt"]["toc3"] == 0.0
        if fixture_name in {"academic", "wide_figure", "hybrid_bid", "degradation"}:
            assert len(insert_toc_ops) == 1
        else:
            assert len(insert_toc_ops) == 0

    def test_heading_numbering_scheme(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        headings = [
            op
            for op in fixture_build.plan.to_dict()["operations"]
            if op["op"] == "writer.add_heading"
        ]
        if fixture_name == "plain_short":
            assert not any(h["args"].get("numbering") for h in headings)
            return
        assert headings, f"{fixture_name} should have headings"
        numbered = [h for h in headings if h["args"].get("numbering")]
        assert numbered, f"{fixture_name} should emit native numbering"
        if fixture_name == "hybrid_bid":
            assert all(
                h["args"].get("numberingScheme") == "hybrid-bid" for h in numbered
            )

    def test_no_empty_trailing_body_section(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        ops = fixture_build.plan.to_dict()["operations"]
        sections = self._sections(fixture_build)
        assert len(sections) >= 1
        last_role = sections[-1]["args"].get("role")
        assert last_role in {"body", "landscape", "front_matter", "cover"}
        if last_role == "body":
            last_sec_idx = next(
                i for i, op in enumerate(ops) if op == sections[-1]
            )
            following = [
                op
                for op in ops[last_sec_idx + 1 :]
                if op["op"] in {"writer.add_heading", "writer.add_paragraph"}
            ]
            assert following, f"{fixture_name} has an empty trailing body section"

    def test_temporary_landscape_continues_numbering(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        landscape = [
            op
            for op in self._sections(fixture_build)
            if op["args"].get("role") == "landscape"
        ]
        if fixture_name != "wide_figure":
            assert not landscape
            return
        assert len(landscape) == 1
        args = landscape[0]["args"]
        assert args["pageNumberFormat"] == "continue"
        assert args["restartPageNumbering"] is False
        assert args.get("landscape") is True
        assert args["linkToPreviousHeader"] is True
        assert args["linkToPreviousFooter"] is True

    def test_no_blank_pages_from_empty_front_matter(
        self, fixture_build: LongformBuild, fixture_name: str
    ) -> None:
        section_ops = self._sections(fixture_build)
        if fixture_name == "protocol_edge":
            assert not any(
                op["args"].get("role") == "front_matter" for op in section_ops
            )
        elif fixture_name == "plain_short":
            assert not any(
                op["args"].get("role") == "front_matter" for op in section_ops
            )


# ---------------------------------------------------------------------------
# Add-in primitive structural assertions (header/footer/page-number)
# ---------------------------------------------------------------------------


class TestAddInM2Primitives:
    """Verify the JSAPI add-in applies M2 header/footer/page-number formatting.

    These tests run Node against the generated add-in and mocked WPS JSAPI
    objects.  They do not require a running WPS instance.
    """

    def _addin_path(self) -> Path:
        return (
            Path(__file__).parents[2]
            / "macos"
            / "wps-jsapi-probe"
            / "addin"
            / "writer-longform-v2.js"
        )

    def _run_js(self, js: str) -> None:
        """Run a Node script that evals the add-in and asserts behavior."""
        path = Path(tempfile.mkdtemp()) / "addin_assert.js"
        path.write_text(js, encoding="utf-8")
        result = subprocess.run(
            ["node", str(path)],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, result.stderr or result.stdout

    def test_set_header_footer_centers_header_and_adds_bottom_border(self) -> None:
        v2_path = json.dumps(str(self._addin_path()))
        js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let headerAlignment = null;
let headerBorderStyle = null;
function makeHeader() {{
  const pf = {{}};
  Object.defineProperty(pf, "Alignment", {{
    get() {{ return pf._alignment; }},
    set(v) {{ pf._alignment = v; headerAlignment = v; }}
  }});
  Object.defineProperty(pf, "Borders", {{
    value: function(side) {{
      const border = {{}};
      Object.defineProperty(border, "LineStyle", {{
        get() {{ return border._lineStyle; }},
        set(v) {{ border._lineStyle = v; headerBorderStyle = v; }}
      }});
      return border;
    }}
  }});
  return {{
    LinkToPrevious: false,
    Range: {{ Text: "", ParagraphFormat: pf }}
  }};
}}
function makeFooter() {{
  return {{
    LinkToPrevious: false,
    Range: {{
      Text: "initial",
      Collapse: function() {{}},
      ParagraphFormat: {{ Alignment: 0 }},
      Fields: {{ Add: function() {{}} }}
    }},
    PageNumbers: {{ RestartNumberingAtSection: false, StartingNumber: 1, NumberStyle: 0 }}
  }};
}}
function makeSection() {{
  return {{
    Headers: function(i) {{ return makeHeader(); }},
    Footers: function(i) {{ return makeFooter(); }}
  }};
}}
const document = {{
  Content: {{ End: 1 }},
  _wpscFirstSectionConfigured: false,
  Sections: {{ Count: 1, Item: function(i) {{ return makeSection(); }} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.configure_section", args: {{role: "body", headerText: "Centered Header", footerText: "", linkToPreviousHeader: false, linkToPreviousFooter: false}}, nodeId: "doc:sec"}}
    ]
  }}
}});
assert.equal(headerAlignment, 1, "header must be centered");
assert.equal(headerBorderStyle, 1, "header must have bottom border");
"""
        self._run_js(js)
    def test_set_page_numbering_adds_page_field_to_footer(self) -> None:
        v2_path = json.dumps(str(self._addin_path()))
        js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let pageFieldAdded = false;
function makeFooter() {{
  return {{
    LinkToPrevious: false,
    Range: {{
      Text: "",
      Collapse: function() {{}},
      ParagraphFormat: {{ Alignment: 0 }},
      Fields: {{ Add: function(range, type) {{ if (type === 33) pageFieldAdded = true; }} }}
    }},
    PageNumbers: {{ RestartNumberingAtSection: false, StartingNumber: 1, NumberStyle: 0 }}
  }};
}}
function makeSection() {{
  return {{
    Headers: function(i) {{ return {{ LinkToPrevious: false, Range: {{ Text: "", ParagraphFormat: {{ Alignment: 0 }} }} }}; }},
    Footers: function(i) {{ return makeFooter(); }}
  }};
}}
const document = {{
  Content: {{ End: 1 }},
  _wpscFirstSectionConfigured: false,
  Sections: {{ Count: 1, Item: function(i) {{ return makeSection(); }} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.configure_section", args: {{role: "body", pageNumberFormat: "arabic", restartPageNumbering: true, startPageNumber: 1, headerText: "", footerText: ""}}, nodeId: "doc:sec"}}
    ]
  }}
}});
assert.ok(pageFieldAdded, "footer must contain a PAGE field");
"""
        self._run_js(js)

@pytest.fixture(scope="session")
def real_macos_bridge():
    """Yield a real LoopbackBridge if WPS is available; otherwise skip."""
    if sys.platform != "darwin":
        pytest.skip("real-WPS acceptance requires macOS")
    try:
        from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
    except Exception as exc:  # pragma: no cover
        pytest.skip(f"macOS bridge module unavailable: {exc}")

    try:
        bridge = LoopbackBridge(allowed_origins={"http://127.0.0.1"})
    except PermissionError as exc:
        pytest.skip(f"loopback socket permission denied: {exc}")
    except Exception as exc:  # pragma: no cover
        pytest.skip(f"loopback bridge unavailable: {exc}")

    with bridge:
        try:
            bridge.wait_registered({"writer"}, timeout=5.0)
        except TimeoutError:
            pytest.skip("WPS writer component did not register within timeout")
        except Exception as exc:  # pragma: no cover
            pytest.skip(f"WPS registration probe failed: {exc}")
        yield bridge

class TestMacOSRealWPS:
    """Generate each M2 fixture through the macOS WPS JSAPI executor.

    These tests always run against the recording executor offline; the macOS
    real-WPS path is exercised only when a LoopbackBridge can bind and WPS
    registers its writer component.
    """

    def _run_fixture(self, name: str, bridge) -> Path:
        from skills.WPSComposer.scripts.longform.macos_executor import (
            MacOSLongformExecutor,
        )

        build = build_fixture(name)
        staging = Path(tempfile.gettempdir()) / "wpsc-m2-acceptance"
        executor = MacOSLongformExecutor(
            bridge=bridge,
            staging_dir=str(staging),
        )
        outcome = executor.execute(
            build.plan, tuple(build.preflight.resources)
        )
        staged = Path(outcome.staged_artifact)
        assert staged.suffix == ".docx"
        assert staged.stat().st_size > 0
        return staged

    def _sectPr_roles(self, document_xml: str) -> dict[str, list[dict[str, str]]]:
        """Extract per-role section properties from document.xml.

        The JSAPI add-in stores the section role in a DocumentVariable field
        whose instruction text contains 'WpsComposerSectionRole_N' and whose
        displayed value is the role name.
        """
        roles: dict[str, list[dict[str, str]]] = {}
        for match in re.finditer(
            r'<w:sectPr[^>]*>.*?</w:sectPr>', document_xml, re.DOTALL
        ):
            sect = match.group(0)
            role_match = re.search(
                r'WpsComposerSectionRole_[0-9]+</w:instrText>.*?<w:t[^>]*>([^<]+)</w:t>',
                sect, re.DOTALL,
            )
            role = role_match.group(1) if role_match else "unknown"
            pg_fmt_match = re.search(r'<w:pgNumType[^>]*?w:fmt="([^"]+)"[^>]*>', sect)
            pg_start_match = re.search(r'<w:pgNumType[^>]*?w:start="([^"]+)"[^>]*>', sect)
            orient_match = re.search(r'<w:pgSz[^>]*?w:orient="([^"]+)"[^>]*>', sect)
            footer_ref = "w:footerReference" in sect
            header_ref = "w:headerReference" in sect
            roles.setdefault(role, []).append({
                "fmt": pg_fmt_match.group(1) if pg_fmt_match else "none",
                "start": pg_start_match.group(1) if pg_start_match else "",
                "orient": orient_match.group(1) if orient_match else "portrait",
                "hasFooterRef": footer_ref,
                "hasHeaderRef": header_ref,
            })
        return roles

    def _styles_xml(self, package: zipfile.ZipFile) -> dict[str, str]:
        """Map style names to their raw w:style XML fragments."""
        try:
            styles = package.read("word/styles.xml").decode("utf-8")
        except KeyError:
            return {}
        result: dict[str, str] = {}
        for match in re.finditer(r'<w:style[^>]*?</w:style>', styles, re.DOTALL):
            fragment = match.group(0)
            name_match = re.search(r'<w:name[^>]*?w:val="([^"]+)"', fragment)
            if name_match:
                result[name_match.group(1)] = fragment
        return result

    def _numbering_xml(self, package: zipfile.ZipFile) -> str:
        """Return numbering.xml contents if present, else empty string."""
        try:
            return package.read("word/numbering.xml").decode("utf-8")
        except KeyError:
            return ""

    def _header_footer_xml(self, package: zipfile.ZipFile) -> dict[str, str]:
        """Read all header/footer XML parts into a dict keyed by part name."""
        return {
            name: package.read(name).decode("utf-8")
            for name in package.namelist()
            if name.startswith(("word/header", "word/footer"))
        }

    def _assert_docx_has_page_skeleton(self, path: Path, fixture_name: str) -> None:
        with zipfile.ZipFile(path) as package:
            document_xml = package.read("word/document.xml").decode("utf-8")
            roles = self._sectPr_roles(document_xml)
            hf_parts = self._header_footer_xml(package)
            styles_xml = self._styles_xml(package)
            numbering_xml = self._numbering_xml(package)

        assert "<w:sectPr" in document_xml, "no section properties found"

        # Cover page must have no page-number footer and no numbering format.
        if fixture_name not in {"plain_short"}:
            assert "cover" in roles, "cover section missing"
            for sect in roles["cover"]:
                assert sect["fmt"] == "none", f"cover page must not be numbered: {sect}"
                assert not sect["hasFooterRef"], f"cover page must not reference a footer: {sect}"

        # Front matter must use lower-roman numbering and restart at 1.
        if "front_matter" in roles:
            for sect in roles["front_matter"]:
                assert sect["fmt"] == "lowerRoman", f"front matter must use roman: {sect}"
                assert sect["start"] == "1", f"front matter must restart at 1: {sect}"

        # Body must use decimal numbering and restart at 1.
        assert "body" in roles, "body section missing"
        for sect in roles["body"]:
            assert sect["fmt"] == "decimal", f"body must use arabic: {sect}"
            assert sect["start"] == "1", f"body must restart at 1: {sect}"
            assert sect["hasHeaderRef"], "body section must reference a header"
            assert sect["hasFooterRef"], "body section must reference a footer"

        # Landscape section must continue numbering and be landscape oriented.
        if fixture_name == "wide_figure":
            assert "landscape" in roles, "landscape section missing"
            for sect in roles["landscape"]:
                assert sect["fmt"] == "decimal", f"landscape must continue decimal: {sect}"
                assert sect["orient"] == "landscape", f"landscape section must be oriented landscape: {sect}"

        # Header must contain centered text and a bottom border.
        header_xml = " ".join(xml for name, xml in hf_parts.items() if name.startswith("word/header"))
        assert '<w:jc w:val="center"' in header_xml, "header text must be centered"
        assert re.search(r'<w:bottom[^>]*w:val="', header_xml), "header must have a bottom paragraph border"

        # Footer must contain a PAGE number field.
        footer_xml = " ".join(xml for name, xml in hf_parts.items() if name.startswith("word/footer"))
        assert re.search(r'<w:instrText[^>]*>\s*PAGE\s*</w:instrText>', footer_xml), "footer must contain a PAGE field"

        # TOC styles must be present with compact spacing.
        for key in ("toc1", "toc2", "toc3"):
            style_name = f"TOC {key[-1]}"
            style_xml = styles_xml.get(style_name, "")
            assert style_xml, f"{style_name} style must be defined"
            assert '<w:jc w:val="left"' in style_xml or '<w:jc' not in style_xml, f"{style_name} must not force center alignment"

        # Heading numbering definitions must exist for numbered fixtures.
        if fixture_name not in {"plain_short"}:
            assert numbering_xml, "numbering.xml must define heading numbering"
            assert "<w:ilvl" in numbering_xml, "numbering definitions must contain outline levels"

    @pytest.mark.parametrize("fixture_name", _FIXTURE_NAMES)
    def test_fixture_generates_docx(
        self, fixture_name: str, real_macos_bridge
    ) -> None:
        staged = self._run_fixture(fixture_name, real_macos_bridge)
        self._assert_docx_has_page_skeleton(staged, fixture_name)


# ---------------------------------------------------------------------------
# Protocol-edge and degradation fixtures
# ---------------------------------------------------------------------------


class TestProtocolEdge:
    """Explicit protocol-edge behaviour."""

    def test_protocol_edge_cover_no_frontmatter(self) -> None:
        build = build_fixture("protocol_edge")
        roles = [
            op["args"].get("role")
            for op in build.plan.to_dict()["operations"]
            if op["op"] == "writer.configure_section"
        ]
        assert roles == ["cover", "body"]

    def test_degradation_issues_emitted(self) -> None:
        build = build_fixture("degradation")
        codes = {issue.code for issue in build.issues}
        assert "ABSTRACT_CONTENT_DEGRADED" in codes
        assert "PAGE_BREAK_CONTENT_DEGRADED" in codes
