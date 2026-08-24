from __future__ import annotations

import json
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation


FIXTURES = Path(__file__).parent / "fixtures"
NAMES = (
    "academic",
    "toc_dense",
    "wide_objects",
    "degradation",
    "unicode",
    "plain_short",
)


def _build(name: str):
    return build_longform_generation(
        (FIXTURES / f"{name}.md").read_text(encoding="utf-8"),
        base_dir=str(FIXTURES),
    )


@pytest.mark.parametrize("name", NAMES)
def test_six_acceptance_fixtures_build_deterministically_and_privately(name):
    first = _build(name)
    second = _build(name)
    assert first.plan.protocol_version == 2
    assert first.semantic.config.layout_engine == "longform"
    assert first.plan.to_dict() == second.plan.to_dict()
    serialized = json.dumps(first.plan.to_dict(), ensure_ascii=False, sort_keys=True)
    assert str(FIXTURES) not in serialized
    assert "/Users/" not in serialized
    front = next(
        operation for operation in first.plan.operations
        if operation.op == "writer.configure_front_matter"
    )
    assert front.args["author"] == "WPSComposer"


def test_dense_toc_uses_the_compact_spacing_floor():
    build = _build("toc_dense")
    operation = next(
        item for item in build.plan.operations
        if item.op == "writer.configure_toc_styles"
    )
    assert set(operation.args["minSpaceBeforePt"].values()) == {0.0}
    assert set(operation.args["minSpaceAfterPt"].values()) == {0.0}
    assert min(operation.args["minFontSizePt"].values()) >= 10.0


def test_wide_fixture_contains_native_figure_and_landscape_table():
    build = _build("wide_objects")
    figure = next(
        item for item in build.plan.operations
        if item.op == "writer.add_captioned_figure"
    )
    table = next(
        item for item in build.plan.operations
        if item.op == "writer.add_semantic_table"
    )
    assert figure.args["widthMode"] == "full"
    assert figure.args["children"]
    assert table.args["orientation"] == "landscape"
    section_orientations = [
        item.args["landscape"]
        for item in build.plan.operations
        if item.op == "writer.configure_section" and "landscape" in item.args
    ]
    assert section_orientations == [True]


def test_degradation_fixture_keeps_visible_controlled_fallbacks():
    build = _build("degradation")
    codes = {issue.code for issue in build.issues}
    assert "REFERENCE_UNRESOLVED" in codes
    assert "FORMULA_MALFORMED" in codes
    assert any(
        item.op == "writer.reserve_document_quality_anchor"
        for item in build.plan.operations
    )


def test_every_quality_addressable_heading_has_a_stable_native_bookmark():
    build = _build("unicode")
    headings = [
        item for item in build.plan.operations if item.op == "writer.add_heading"
    ]
    bookmarks = [item.args["bookmarkName"] for item in headings]
    assert bookmarks
    assert all(name.startswith("wpsc_head_") and len(name) == 34 for name in bookmarks)
    assert len(bookmarks) == len(set(bookmarks))
