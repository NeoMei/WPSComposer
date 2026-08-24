from __future__ import annotations

import json
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation


HERE = Path(__file__).parent
FIXTURES = HERE / "fixtures"
SNAPSHOTS = HERE / "snapshots"
NAMES = ("auto_boundary", "chapter_native", "degradation", "global_media")


def _build(name: str):
    source = FIXTURES / f"{name}.md"
    return build_longform_generation(
        source.read_text(encoding="utf-8"), base_dir=str(FIXTURES)
    )


def _ops(build, name: str) -> list[dict]:
    return [
        operation
        for operation in build.plan.to_dict()["operations"]
        if operation["op"] == name
    ]


def _canonical_operations(build) -> bytes:
    return (
        json.dumps(
            build.plan.to_dict()["operations"],
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


@pytest.mark.parametrize("name", NAMES)
def test_canonical_plan_snapshot_is_byte_stable_and_private(name: str) -> None:
    first = _build(name)
    second = _build(name)
    actual = _canonical_operations(first)
    assert actual == _canonical_operations(second)
    assert actual == (SNAPSHOTS / f"{name}.json").read_bytes()
    lowered = actual.lower()
    assert b"payloadbytes" not in lowered
    assert b"sourcepath" not in lowered
    assert str(FIXTURES).encode() not in actual


def test_auto_numbering_is_object_local_around_numbered_h1() -> None:
    build = _build("auto_boundary")
    figures = _ops(build, "writer.add_captioned_figure")
    tables = _ops(build, "writer.add_semantic_table")
    assert [item["args"]["numbering"]["mode"] for item in figures] == [
        "global",
        "chapter",
        "chapter",
    ]
    assert [item["args"]["numbering"]["mode"] for item in tables] == [
        "global",
        "chapter",
        "chapter",
    ]
    assert figures[0]["args"]["numbering"]["resetLevel"] is None
    assert figures[1]["args"]["numbering"]["resetLevel"] == 1
    assert figures[2]["args"]["numbering"] == figures[1]["args"]["numbering"]
    assert tables[2]["args"]["numbering"] == tables[1]["args"]["numbering"]
    headings = _ops(build, "writer.add_heading")
    assert headings[0]["args"]["numbering"] is True
    assert headings[1]["args"]["numbering"] is False
    assert headings[1]["args"]["sequenceTransparent"] is True


def test_front_matter_indexes_precede_native_objects_and_are_populated() -> None:
    build = _build("chapter_native")
    operations = build.plan.to_dict()["operations"]
    names = [item["op"] for item in operations]
    first_native = min(
        names.index("writer.add_captioned_figure"),
        names.index("writer.add_semantic_table"),
    )
    figure_index = names.index("writer.insert_figure_index")
    table_index = names.index("writer.insert_table_index")
    assert figure_index < first_native
    assert table_index < first_native
    assert operations[figure_index]["args"]["sequenceId"] == "WPSC_FIG"
    assert operations[table_index]["args"]["sequenceId"] == "WPSC_TAB"


def test_supported_media_normalizers_and_two_image_layout_are_exact() -> None:
    chapter = _build("chapter_native")
    global_build = _build("global_media")
    children = [
        child
        for build in (chapter, global_build)
        for figure in _ops(build, "writer.add_captioned_figure")
        for child in figure["args"]["children"]
    ]
    assert {item["mediaType"] for item in children} == {
        "image/png",
        "image/jpeg",
        "image/bmp",
        "image/svg+xml",
    }
    assert {item["normalizerId"] for item in children} >= {
        "none-v1",
        "exif-transpose-png-v1",
        "gif-first-frame-png-v1",
        "tiff-first-page-png-v1",
        "svg-static-v1",
    }
    columns = [
        figure
        for build in (chapter, global_build)
        for figure in _ops(build, "writer.add_captioned_figure")
        if figure["args"]["layout"] == "columns"
    ]
    assert columns and all(len(item["args"]["children"]) == 2 for item in columns)


def test_table_borders_merges_and_recovery_ladders_are_exact() -> None:
    native = _build("chapter_native")
    degradation = _build("degradation")
    native_tables = _ops(native, "writer.add_semantic_table")
    assert native_tables[0]["args"]["borderSpec"] == {
        "top": 1.5,
        "bottom": 1.5,
        "headerBottom": 0.75,
        "left": 0.0,
        "right": 0.0,
        "insideHorizontal": 0.0,
        "insideVertical": 0.0,
    }
    assert native_tables[0]["args"]["merges"] == [
        {"top": 2, "left": 1, "bottom": 3, "right": 1}
    ]
    invalid = next(
        item
        for item in _ops(degradation, "writer.add_semantic_table")
        if item["nodeId"] == "tab:invalid-merge"
    )
    assert invalid["args"]["merges"] == []
    assert invalid["args"]["plannedDegradation"][0]["actions"] == [
        "discard-all-merges",
        "preserve-complete-grid",
    ]
    tall = next(
        item
        for item in _ops(degradation, "writer.add_semantic_table")
        if item["nodeId"] == "tab:tall-group"
    )
    assert tall["args"]["merges"] == [
        {"top": 2, "left": 1, "bottom": 17, "right": 1}
    ]
    assert tall["args"]["allowRowSplit"] is False
    assert tall["failurePolicy"]["fallback"] == "grid-then-text"


def test_references_equation_shell_and_caption_missing_omission() -> None:
    chapter = _build("chapter_native")
    references = _ops(chapter, "writer.add_cross_reference")
    assert len(references) == 1
    runs = references[0]["args"]["runs"]
    assert [run["type"] for run in runs].count("reference") == 3
    assert all(run["bookmarkName"].startswith("wpsc_") for run in runs if run["type"] == "reference")
    equation = _ops(chapter, "writer.add_equation")[0]
    assert equation["args"]["numbering"]["sequenceId"] == "WPSC_EQ"
    assert equation["args"]["numbering"]["prefix"] == "("
    assert equation["args"]["numbering"]["suffix"] == ")"
    assert "nativeMath" not in equation["args"]

    degraded = _build("degradation")
    missing_figure = next(
        item
        for item in _ops(degraded, "writer.add_captioned_figure")
        if item["nodeId"] == "fig:missing-caption"
    )
    missing_table = next(
        item
        for item in _ops(degraded, "writer.add_semantic_table")
        if item["nodeId"] == "tab:missing-caption"
    )
    for item in (missing_figure, missing_table):
        assert item["args"]["indexable"] is False
        assert "bookmarkName" not in item["args"]
        # The closed schema retains the controlled descriptor, but native
        # executors must not consume it when the caption is absent.
        assert item["args"]["numbering"]["sequenceId"] in {"WPSC_FIG", "WPSC_TAB"}
    assert not _ops(degraded, "writer.add_cross_reference")
