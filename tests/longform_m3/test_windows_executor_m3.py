from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.generation_plan import GenerationOperation, GenerationPlan
from skills.WPSComposer.scripts.longform.resources import ImageProfile, PreparedLongformResource
from skills.WPSComposer.scripts.longform.windows_executor import (
    WindowsLongformExecutor,
    WindowsLongformExecutorError,
)
from tests.longform_m3.fakes.windows_com import RecordingNativeComposer


BOOKMARK = "wpsc_fig_" + "a" * 24


def _numbering(sequence="WPSC_FIG", prefix="图 ", suffix=""):
    return {
        "mode": "chapter", "sequenceId": sequence, "chapterStyleLevel": 1,
        "resetLevel": 1, "prefix": prefix, "suffix": suffix,
    }


def _manifest(resource: PreparedLongformResource) -> str:
    envelope = {
        "version": "1",
        "entries": [{
            "resourceId": resource.id,
            "sourceSha256": resource.source_sha256,
            "payloadSha256": resource.payload_sha256,
            "byteLength": len(resource.payload_bytes),
            "mediaType": resource.media_type,
            "normalizerId": resource.normalizer_id,
        }],
    }
    raw = json.dumps(envelope, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()
    return "sha256:" + hashlib.sha256(raw).hexdigest()


def _resource(payload=b"normalized-png"):
    digest = hashlib.sha256(payload).hexdigest()
    return PreparedLongformResource(
        id="image-1", media_type="image/png", source_sha256="1" * 64,
        payload_sha256=digest, normalizer_id="none-v1", payload_bytes=payload,
        image_profile=ImageProfile(800, 400, 144.0, 144.0, 1, "PNG", False),
    )


def _figure():
    return GenerationOperation(
        "writer.add_captioned_figure",
        {
            "caption": "示例", "numbering": _numbering(), "bookmarkName": BOOKMARK,
            "indexable": True, "referenceable": True, "widthMode": "full",
            "orientation": "portrait", "kind": "diagram", "layout": "stack",
            "keepWithCaption": True,
            "children": [{
                "nodeId": "fig:one/image:1", "resourceId": "image-1",
                "displayWidthPt": 240.0, "displayHeightPt": 120.0,
                "effectiveDpi": 144.0, "mediaType": "image/png", "normalizerId": "none-v1",
            }],
        },
        node_id="fig:one",
        failure_policy={"mode": "degrade", "recoverableCodes": ["IMAGE_INSERT_FAILED"], "fallback": "figure-child-stack-then-notice"},
    )


def _table():
    return GenerationOperation(
        "writer.add_semantic_table",
        {
            "caption": "表格", "numbering": _numbering("WPSC_TAB", "表 "),
            "bookmarkName": "wpsc_tab_" + "b" * 24, "indexable": True,
            "referenceable": True, "headers": ["A", "B"], "rows": [["1", "2"]],
            "alignments": ["left", "right"], "style": "three-line",
            "orientation": "portrait",
            "borderSpec": {"top": 1.5, "bottom": 1.5, "headerBottom": .75, "left": 0.0, "right": 0.0, "insideHorizontal": 0.0, "insideVertical": 0.0},
            "merges": [], "repeatHeader": True, "allowRowSplit": False,
            "cellIndentPt": 0.0, "plannedDegradation": [],
            "keepCaptionWithFirstRow": True,
        },
        node_id="tab:one",
        failure_policy={"mode": "degrade", "recoverableCodes": ["TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED", "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"], "fallback": "grid-then-text"},
    )


def _plan(resource, *ops):
    finalize = GenerationOperation("writer.finalize_fields", {"maxRounds": 3}, node_id="doc:finalize")
    return GenerationPlan(
        component="writer", operations=(*ops, finalize), protocol_version=2,
        semantic_version="longform-1", resource_manifest_version=1,
        resource_manifest_digest=_manifest(resource) if resource else "sha256:" + "0" * 64,
    )


def test_native_operations_consume_closed_descriptors_and_private_locator(tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    reference = GenerationOperation(
        "writer.add_cross_reference", {"runs": [
            {"type": "text", "text": "见"},
            {"type": "reference", "targetNodeId": "fig:one", "targetKind": "figure", "bookmarkName": BOOKMARK, "prefix": "图 ", "suffix": "", "fallbackText": "[图]"},
        ]}, node_id="para:one",
        failure_policy={"mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"], "fallback": "inline-fallback"},
    )
    equation = GenerationOperation(
        "writer.add_equation", {"source": "E=mc^2", "numbering": _numbering("WPSC_EQ", "(", ")"), "bookmarkName": "wpsc_eq_" + "c" * 24, "fallbackText": "E=mc^2"}, node_id="eq:one", failure_policy={"mode": "fail"},
    )
    index = GenerationOperation("writer.insert_figure_index", {"title": "图目录", "sequenceId": "WPSC_FIG", "titleStyleId": "WPSC_INDEX_TITLE"}, node_id="doc:figure-index")
    front = GenerationOperation("writer.configure_section", {"role": "front_matter"}, node_id="doc:front")
    body = GenerationOperation("writer.configure_section", {"role": "body"}, node_id="doc:body")
    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=lambda: composer)

    executor.execute(_plan(resource, front, index, body, _figure(), _table(), equation, reference), (resource,))

    calls = {name: kwargs for name, kwargs in composer.calls if name in {"figure", "table", "equation", "reference", "index"}}
    assert calls["figure"]["numbering"] == _numbering()
    assert calls["figure"]["resource_locators"]["image-1"].endswith(".png")
    assert calls["table"]["borderSpec"]["insideVertical"] == 0.0
    assert calls["equation"]["source"] == "E=mc^2"
    assert calls["reference"]["runs"][1]["bookmarkName"] == BOOKMARK
    assert calls["index"] == {"title": "图目录", "sequence_id": "WPSC_FIG", "title_style_id": "WPSC_INDEX_TITLE", "owner_node_id": "doc:figure-index"}
    assert composer.resource_paths_seen
    assert all(not Path(path).exists() for path in composer.resource_paths_seen)


def test_resource_hash_mismatch_aborts_before_composer_acquisition(tmp_path: Path):
    resource = _resource()
    corrupt = PreparedLongformResource(**{**resource.__dict__, "payload_bytes": b"corrupt"})
    acquired = False
    def factory():
        nonlocal acquired
        acquired = True
        return RecordingNativeComposer()
    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=factory)
    with pytest.raises(WindowsLongformExecutorError, match="resource validation"):
        executor.execute(_plan(resource, _figure()), (corrupt,))
    assert acquired is False
    assert not list(tmp_path.glob("wpsc-resource-*"))


def test_unknown_native_error_and_save_error_are_fatal_and_cleanup(tmp_path: Path):
    resource = _resource()
    for failing_method in ("figure", "save_docx"):
        composer = RecordingNativeComposer()
        composer.failures[failing_method] = "COM_ERROR"
        executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=lambda: composer)
        with pytest.raises(WindowsLongformExecutorError):
            executor.execute(_plan(resource, _figure()), (resource,))
        assert all(not Path(path).exists() for path in composer.resource_paths_seen)


def test_native_field_adapter_runs_five_phase_order_twice_until_stable(tmp_path: Path):
    composer = RecordingNativeComposer()
    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=lambda: composer)
    executor.execute(_plan(None), ())
    phases = [name for name, _ in composer.calls if name.startswith("phase-")]
    assert phases == [
        "phase-numbering", "phase-references", "phase-indexes", "phase-pages", "phase-snapshot",
        "phase-numbering", "phase-references", "phase-indexes", "phase-pages", "phase-snapshot",
    ]


def test_writer_controlled_field_code_helpers_are_exact():
    from skills.WPSComposer.scripts.writer import _caption_field_codes, _reference_field_code

    assert _caption_field_codes(_numbering()) == ("STYLEREF 1 \\s", "SEQ WPSC_FIG \\* ARABIC \\s 1")
    assert _caption_field_codes({**_numbering(), "mode": "global", "chapterStyleLevel": None, "resetLevel": None}) == (None, "SEQ WPSC_FIG \\* ARABIC")
    assert _reference_field_code(BOOKMARK) == "REF " + BOOKMARK + " \\h"


class _Format:
    def __init__(self):
        self.Alignment = None
        self.KeepTogether = None
        self.KeepWithNext = None


class _Range:
    def __init__(self, start=0, end=None):
        self.Start = start
        self.End = start if end is None else end
        self.ParagraphFormat = _Format()
        self.Text = ""

    def Collapse(self, direction):
        assert direction == 0
        self.Start = self.End


class _Selection:
    def __init__(self):
        self.pos = 0
        self.ParagraphFormat = type("PF", (), {"TabStops": type("TS", (), {"Add": lambda self, value: None})()})()
        self.typed = []
        self.Style = None

    @property
    def Range(self):
        return _Range(self.pos)

    def SetRange(self, start, end):
        assert start == end
        self.pos = end

    def TypeText(self, text):
        self.typed.append(text)
        self.pos += len(text)

    def TypeParagraph(self):
        self.typed.append("\n")
        self.pos += 1

    def EndKey(self, unit):
        assert unit == 6


class _NativeField:
    def __init__(self, code, end):
        self.code = code
        self.Result = type("Result", (), {"End": end, "Text": code})()
        self.updates = 0

    def Update(self):
        self.updates += 1


class _Fields:
    def __init__(self):
        self.calls = []

    def Add(self, rng, field_type, code, preserve):
        self.calls.append((rng.Start, field_type, code, preserve))
        return _NativeField(code, rng.Start + 1)


class _Bookmarks:
    Count = 0

    def __init__(self):
        self.calls = []

    def Add(self, name, rng):
        self.calls.append((name, rng.Start, rng.End))


class _Indexes:
    def __init__(self):
        self.calls = []
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def Add(self, rng, label):
        self.calls.append((rng.Start, label))
        item = _NativeField("INDEX " + label, rng.Start + 1)
        self.items.append(item)
        return item

    def Item(self, index):
        return self.items[index - 1]


class _EmptyIndexes:
    Count = 0

    def Item(self, index):
        raise AssertionError(index)


class _Styles:
    def __call__(self, name):
        return name


class _NativeDocument:
    def __init__(self):
        self.Fields = _Fields()
        self.Bookmarks = _Bookmarks()
        self.TablesOfFigures = _Indexes()
        self.TablesOfContents = _EmptyIndexes()
        self.Styles = _Styles()
        self.repaginate_count = 0

    def Range(self, start, end):
        return _Range(start, end)

    def Repaginate(self):
        self.repaginate_count += 1

    def ComputeStatistics(self, kind):
        assert kind == 2
        return 3


def _writer_with_native_fakes():
    from skills.WPSComposer.scripts.writer import WriterComposer
    writer = WriterComposer.__new__(WriterComposer)
    writer._doc = _NativeDocument()
    writer._app = type("App", (), {})()
    writer._selection = _Selection()
    # WriterComposer.selection normally resolves app.Selection.
    writer._app.Selection = writer._selection
    return writer


def test_writer_chapter_caption_fields_and_bookmark_cover_number_only():
    writer = _writer_with_native_fakes()
    writer._add_native_caption("示例", _numbering(), BOOKMARK, "fig:one")

    assert [call[2] for call in writer.doc.Fields.calls] == [
        "STYLEREF 1 \\s", "SEQ WPSC_FIG \\* ARABIC \\s 1",
    ]
    # Prefix starts at 0 and is excluded; caption prose follows the bookmarked
    # chapter-separator-sequence range and is also excluded.
    assert writer.doc.Bookmarks.calls == [(BOOKMARK, 2, 5)]
    assert writer._selection.typed == ["图 ", "-", " 示例", "\n"]


@pytest.mark.parametrize(
    ("sequence", "prefix", "suffix", "bookmark", "kind"),
    [
        ("WPSC_FIG", "图 ", "", "wpsc_fig_" + "a" * 24, "SEQ_FIG"),
        ("WPSC_TAB", "表 ", "", "wpsc_tab_" + "b" * 24, "SEQ_TAB"),
        ("WPSC_EQ", "(", ")", "wpsc_eq_" + "c" * 24, "SEQ_EQ"),
    ],
)
def test_writer_global_caption_sequences_are_independent(sequence, prefix, suffix, bookmark, kind):
    writer = _writer_with_native_fakes()
    numbering = {
        "mode": "global", "sequenceId": sequence,
        "chapterStyleLevel": None, "resetLevel": None,
        "prefix": prefix, "suffix": suffix,
    }
    writer._add_native_number_shell(numbering, bookmark, "owner")
    assert [call[2] for call in writer.doc.Fields.calls] == [f"SEQ {sequence} \\* ARABIC"]
    assert writer._native_fields()[0][1] == kind
    assert writer.doc.Bookmarks.calls[0][1] == len(prefix)
    assert writer.doc.Bookmarks.calls[0][2] == len(prefix) + 1


def test_writer_reference_and_native_index_use_controlled_native_apis():
    writer = _writer_with_native_fakes()
    writer.add_cross_reference_paragraph(
        runs=[
            {"type": "text", "text": "见"},
            {"type": "reference", "bookmarkName": BOOKMARK, "prefix": "图 ", "suffix": "", "fallbackText": "[图]", "targetNodeId": "fig:one", "targetKind": "figure"},
            {"type": "text", "text": "。"},
        ],
        owner_node_id="para:one",
    )
    writer.insert_caption_index_native(
        title="图目录", sequence_id="WPSC_FIG",
        title_style_id="WPSC_INDEX_TITLE", owner_node_id="doc:figure-index",
    )

    assert [call[2] for call in writer.doc.Fields.calls] == ["REF " + BOOKMARK + " \\h"]
    assert writer.doc.TablesOfFigures.calls == [(10, "WPSC_FIG")]
    assert writer._selection.typed[:4] == ["见", "图 ", "。", "\n"]


def test_writer_refreshes_native_indexes_and_mutation_changes_snapshot_not_owner():
    writer = _writer_with_native_fakes()
    writer._add_native_number_shell(_numbering(), BOOKMARK, "fig:one")
    writer.add_cross_reference_paragraph(
        runs=[{"type": "reference", "bookmarkName": BOOKMARK, "prefix": "图 ", "suffix": "", "fallbackText": "[图]", "targetNodeId": "fig:one", "targetKind": "figure"}],
        owner_node_id="para:one",
    )
    writer.insert_caption_index_native(
        title="图目录", sequence_id="WPSC_FIG",
        title_style_id="WPSC_INDEX_TITLE", owner_node_id="doc:figure-index",
    )

    writer.repaginate_and_update_numbering()
    writer.refresh_bookmarks_and_references()
    writer.refresh_indexes()
    writer.repaginate_and_update_page_fields()
    first = writer.snapshot_fields()
    bookmark_before = tuple(writer.doc.Bookmarks.calls)

    # Model move/insert/delete: WPS changes visible SEQ/REF/index results while
    # the plan-owned field identity and bookmark owner remain fixed.
    for _owner, kind, native, _category in writer._native_fields():
        if kind in {"SEQ_FIG", "REF", "TOF_FIG"}:
            native.Result.Text += "-mutated"
    second = writer.snapshot_fields()

    assert [item.stable_key for item in first] == [item.stable_key for item in second]
    assert [item.result_hash for item in first] != [item.result_hash for item in second]
    assert tuple(writer.doc.Bookmarks.calls) == bookmark_before
    assert writer.doc.TablesOfFigures.items[0].updates == 1
    assert writer.doc.repaginate_count == 2


class _Border:
    def __init__(self):
        self.LineStyle = None
        self.LineWidth = None


class _Borders:
    def __init__(self):
        self.values = {}

    def __call__(self, key):
        return self.values.setdefault(key, _Border())


class _Rows:
    def __init__(self):
        self.header_borders = _Borders()

    def __call__(self, row):
        assert row == 1
        return type("Row", (), {"Borders": self.header_borders})()


def test_writer_applies_exact_three_line_border_contract():
    writer = _writer_with_native_fakes()
    table = type("Table", (), {"Borders": _Borders(), "Rows": _Rows()})()
    writer._apply_native_table_borders(table, {
        "top": 1.5, "bottom": 1.5, "headerBottom": .75,
        "left": 0.0, "right": 0.0,
        "insideHorizontal": 0.0, "insideVertical": 0.0,
    })
    assert table.Borders(-1).LineWidth == 12
    assert table.Borders(-3).LineWidth == 12
    assert table.Rows(1).Borders(-3).LineWidth == 6
    assert all(table.Borders(key).LineStyle == 0 for key in (-2, -4, -5, -6))


class _Column:
    def __init__(self):
        self.Width = None

    def SetWidth(self, width, rule):
        assert rule == 0
        self.Width = width


class _ColumnTable:
    def __init__(self):
        self._columns = {index: _Column() for index in (1, 2, 3)}
        self.Borders = _Borders()

    def Columns(self, index):
        return self._columns[index]


def test_writer_two_column_container_has_exact_12pt_gap_and_no_borders():
    writer = _writer_with_native_fakes()
    table = _ColumnTable()
    writer._doc.Tables = type("Tables", (), {"Add": lambda self, rng, rows, cols: table})()
    children = [{"displayWidthPt": 120.0}, {"displayWidthPt": 150.0}]

    assert writer._native_columns_container(children) is table
    assert [table.Columns(index).Width for index in (1, 2, 3)] == [120.0, 12.0, 150.0]
    assert all(table.Borders(index).LineStyle == 0 for index in range(-6, 0))


def test_writer_inline_figure_child_preserves_aspect_and_caption_cohesion():
    writer = _writer_with_native_fakes()
    calls = []
    shape = type("Shape", (), {"Range": _Range(0, 5)})()
    writer.add_image = lambda locator, **kwargs: calls.append((locator, kwargs)) or shape
    writer._native_insert_figure_child(
        {"displayWidthPt": 240.0, "displayHeightPt": 120.0},
        "/private/staged.png", "fig:one",
    )
    assert calls == [("/private/staged.png", {
        "width": 240.0, "height": 120.0, "inline": True,
        "preserve_aspect": True, "alt": "fig:one",
    })]
    assert shape.Range.ParagraphFormat.Alignment == 1
    assert shape.Range.ParagraphFormat.KeepTogether == -1
    assert shape.Range.ParagraphFormat.KeepWithNext == -1


@pytest.mark.parametrize(("orientation", "sections"), [("portrait", []), ("landscape", [True, False])])
def test_figure_caption_is_after_children_and_landscape_is_explicit_only(orientation, sections):
    from skills.WPSComposer.scripts.writer import WriterComposer

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = _Selection()
    writer._app = type("App", (), {"Selection": writer._selection})()
    events = []
    writer.add_section = lambda landscape=None: events.append(("section", landscape))
    writer._native_insert_figure_child = lambda child, locator, owner: events.append(("child", child["nodeId"]))
    writer._add_native_caption = lambda *args, **kwargs: events.append(("caption", args[0]))

    writer.add_captioned_figure_native(
        caption="图示", numbering=_numbering(), bookmarkName=BOOKMARK,
        indexable=True, referenceable=True, widthMode="full",
        orientation=orientation, kind="diagram",
        children=[{"nodeId": "img:1", "resourceId": "one"}],
        layout="stack", keepWithCaption=True, owner_node_id="fig:one",
        resource_locators={"one": "/private/one.png"},
    )
    assert [value for name, value in events if name == "section"] == sections
    assert next(index for index, item in enumerate(events) if item[0] == "child") < next(index for index, item in enumerate(events) if item[0] == "caption")


def test_one_figure_child_failure_rolls_back_that_child_then_retries_stack():
    from skills.WPSComposer.scripts.writer import NativeWriterObjectError, WriterComposer

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = _Selection()
    writer._app = type("App", (), {"Selection": writer._selection})()
    events = []
    attempts = {"img:1": 0, "img:2": 0}
    writer._native_position = lambda: 10 if attempts["img:2"] == 0 else 20

    def insert(child, locator, owner):
        node = child["nodeId"]
        attempts[node] += 1
        events.append(("insert", node, attempts[node]))
        if node == "img:2" and attempts[node] == 1:
            events.append(("rollback", node))
            raise NativeWriterObjectError("IMAGE_INSERT_FAILED")

    writer._native_insert_figure_child = insert
    writer._native_rollback = lambda start: events.append(("outer-rollback", start))
    writer._add_native_caption = lambda *args, **kwargs: events.append(("caption", args[0]))
    writer.add_degradation_notice = lambda *args: events.append(("notice", args[0]))

    outcome = writer.add_captioned_figure_native(
        caption="图示", numbering=_numbering(), bookmarkName=BOOKMARK,
        indexable=True, referenceable=True, widthMode="full", orientation="portrait",
        kind="diagram",
        children=[
            {"nodeId": "img:1", "resourceId": "one"},
            {"nodeId": "img:2", "resourceId": "two"},
        ],
        layout="stack", keepWithCaption=True, owner_node_id="fig:one",
        resource_locators={"one": "/private/one.png", "two": "/private/two.png"},
    )

    assert attempts == {"img:1": 1, "img:2": 2}
    assert ("rollback", "img:2") in events
    assert not [event for event in events if event[0] == "outer-rollback"]
    assert events[-1] == ("caption", "图示")
    assert outcome["issues"][0]["code"] == "IMAGE_INSERT_FAILED"


class _Cell:
    def __init__(self, row, col, merge_log):
        self.row = row
        self.col = col
        self.Range = _Range()
        self._merge_log = merge_log

    def Merge(self, other):
        self._merge_log.append((self.row, self.col, other.row, other.col))


class _TableRows:
    def __init__(self):
        self.AllowBreakAcrossPages = None
        self.header = type("Header", (), {"HeadingFormat": None, "Borders": _Borders()})()

    def __call__(self, index):
        assert index == 1
        return self.header


class _SemanticTableFake:
    def __init__(self, row_count, col_count):
        self.merge_log = []
        self.cells = {
            (row, col): _Cell(row, col, self.merge_log)
            for row in range(1, row_count + 1)
            for col in range(1, col_count + 1)
        }
        self.Rows = _TableRows()
        self.Borders = _Borders()
        self.Range = _Range(0, 80)

    def Cell(self, row, col):
        return self.cells[(row, col)]


def test_writer_native_table_preserves_alignment_row_and_merge_order():
    writer = _writer_with_native_fakes()
    table = _SemanticTableFake(3, 2)
    writer._doc.Tables = type("Tables", (), {"Add": lambda self, rng, rows, cols: table})()
    writer._apply_native_table_borders = lambda native, spec: None
    merges = [
        {"top": 2, "left": 1, "bottom": 2, "right": 2},
        {"top": 3, "left": 1, "bottom": 3, "right": 2},
    ]

    writer._create_native_table(
        ["A", "B"], [["1", ""], ["2", ""]], ["left", "right"],
        {"unused": 0.0}, True, False, 0.0, merges,
    )

    assert table.Rows.AllowBreakAcrossPages == 0
    assert table.Rows(1).HeadingFormat == -1
    assert table.Cell(1, 1).Range.ParagraphFormat.FirstLineIndent == 0.0
    assert table.Cell(2, 1).Range.ParagraphFormat.Alignment == 0
    assert table.Cell(2, 2).Range.ParagraphFormat.Alignment == 2
    assert table.merge_log == [(2, 1, 2, 2), (3, 1, 3, 2)]


def test_planned_invalid_merge_notice_is_caption_adjacent_before_grid():
    from skills.WPSComposer.scripts.writer import WriterComposer

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = _Selection()
    writer._app = type("App", (), {"Selection": writer._selection})()
    calls = []
    writer._add_native_caption = lambda *args, **kwargs: calls.append("caption")
    writer._add_native_table_notice = lambda *args: calls.append("notice")
    writer._native_position = lambda: 10
    writer._create_native_table = lambda *args: calls.append("table") or object()
    writer._table_overflow_group = lambda table, merges: None

    writer.add_semantic_table_native(
        caption="表格", numbering=_numbering("WPSC_TAB", "表 "), bookmarkName=None,
        indexable=False, referenceable=False, headers=["A"], rows=[["1"]],
        alignments=["left"], style="grid", orientation="portrait",
        borderSpec={"top": .75, "bottom": .75, "headerBottom": .75, "left": .75, "right": .75, "insideHorizontal": .75, "insideVertical": .75},
        merges=[], repeatHeader=True, allowRowSplit=False, cellIndentPt=0.0,
        plannedDegradation=[{
            "code": "TABLE_MERGE_INVALID", "message": "invalid", "placement": "block",
            "insertAfter": "caption", "trigger": "invalid-merge-declaration",
            "recoveryScope": "complete-table", "actions": ["discard-all-merges", "preserve-complete-grid"],
        }],
        keepCaptionWithFirstRow=True, owner_node_id="tab:one",
    )
    assert calls == ["caption", "notice", "table"]


def test_table_vertical_group_rolls_back_complete_table_to_splittable_grid():
    from skills.WPSComposer.scripts.writer import WriterComposer

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = _Selection()
    writer._app = type("App", (), {"Selection": writer._selection})()
    calls = []
    writer._add_native_caption = lambda *args, **kwargs: calls.append(("caption", kwargs))
    writer._native_position = lambda: 40
    writer._native_rollback = lambda start: calls.append(("rollback", start))
    writer._add_native_table_notice = lambda *args: calls.append(("notice", args))
    writer._table_overflow_group = lambda table, merges: (2, 3) if merges else None

    def create(*args):
        calls.append(("table", args))
        return object()
    writer._create_native_table = create

    outcome = writer.add_semantic_table_native(
        caption="表格", numbering=_numbering("WPSC_TAB", "表 "),
        bookmarkName="wpsc_tab_" + "b" * 24, indexable=True, referenceable=True,
        headers=["A", "B"], rows=[["1", ""]], alignments=["left", "right"],
        style="three-line", orientation="portrait",
        borderSpec={"top": 1.5, "bottom": 1.5, "headerBottom": .75, "left": 0.0, "right": 0.0, "insideHorizontal": 0.0, "insideVertical": 0.0},
        merges=[{"top": 2, "left": 1, "bottom": 3, "right": 1}],
        repeatHeader=True, allowRowSplit=False, cellIndentPt=0.0,
        plannedDegradation=[], keepCaptionWithFirstRow=True, owner_node_id="tab:one",
    )

    table_calls = [item for item in calls if item[0] == "table"]
    assert calls[0][0] == "caption"
    assert ("rollback", 40) in calls
    assert table_calls[0][1][-1] == [{"top": 2, "left": 1, "bottom": 3, "right": 1}]
    assert table_calls[1][1][-1] == ()
    assert table_calls[1][1][5] is True  # allowRowSplit
    assert table_calls[1][1][3]["insideVertical"] == .75
    assert outcome["issues"][0]["code"] == "TABLE_ROW_FORCED_SPLIT"


def test_table_named_style_failure_uses_grid_then_text_ladder():
    from skills.WPSComposer.scripts.writer import NativeWriterObjectError, WriterComposer

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = _Selection()
    writer._app = type("App", (), {"Selection": writer._selection})()
    calls = []
    writer._add_native_caption = lambda *args, **kwargs: calls.append("caption")
    writer._native_position = lambda: 20
    writer._native_rollback = lambda start: calls.append(("rollback", start))
    writer._add_native_table_notice = lambda *args: calls.append(("notice", args))
    writer._table_overflow_group = lambda table, merges: None
    writer._add_native_table_text_fallback = lambda headers, rows: calls.append(("text", headers, rows))

    attempts = iter(("TABLE_STYLE_APPLY_FAILED", "TABLE_INSERT_FAILED"))
    def create(*args):
        code = next(attempts)
        calls.append(("attempt", args[3], args[-1]))
        raise NativeWriterObjectError(code)
    writer._create_native_table = create

    outcome = writer.add_semantic_table_native(
        caption="表格", numbering=_numbering("WPSC_TAB", "表 "),
        bookmarkName="wpsc_tab_" + "b" * 24, indexable=True, referenceable=True,
        headers=["A"], rows=[["1"]], alignments=["left"], style="three-line",
        orientation="portrait",
        borderSpec={"top": 1.5, "bottom": 1.5, "headerBottom": .75, "left": 0.0, "right": 0.0, "insideHorizontal": 0.0, "insideVertical": 0.0},
        merges=[], repeatHeader=True, allowRowSplit=False, cellIndentPt=0.0,
        plannedDegradation=[], keepCaptionWithFirstRow=True, owner_node_id="tab:one",
    )

    attempts_seen = [item for item in calls if isinstance(item, tuple) and item[0] == "attempt"]
    assert attempts_seen[0][1]["insideVertical"] == 0.0
    assert attempts_seen[1][1]["insideVertical"] == .75
    assert attempts_seen[1][2] == ()
    assert calls.count(("rollback", 20)) == 2
    assert ("text", ["A"], [["1"]]) in calls
    assert outcome["issues"][-1]["code"] == "TABLE_INSERT_FAILED"


def test_executor_only_exact_recoverable_code_degrades_without_path_leak(tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    composer.failures["figure"] = "IMAGE_INSERT_FAILED"
    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=lambda: composer)

    outcome = executor.execute(_plan(resource, _figure()), (resource,))

    assert [issue.code for issue in outcome.issues] == ["IMAGE_INSERT_FAILED"]
    assert all(str(tmp_path) not in issue.message for issue in outcome.issues)
    assert any(name == "notice" for name, _ in composer.calls)


@pytest.mark.parametrize("code", ["BOOKMARK_FAILED", "FIELD_INSERT_FAILED", "LOCAL_MUTATION_ROLLBACK_FAILED"])
def test_executor_non_allowlisted_native_failures_abort(code: str, tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    composer.failures["figure"] = code
    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=lambda: composer)
    with pytest.raises(WindowsLongformExecutorError):
        executor.execute(_plan(resource, _figure()), (resource,))
