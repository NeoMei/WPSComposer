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
from skills.WPSComposer.scripts.writer import NativeWriterObjectError, WriterComposer
from tests.longform_m3.fakes.windows_com import RecordingNativeComposer


EMPTY_MANIFEST_DIGEST = (
    "sha256:dc7749a3af2a2bb77cad0700bddd3716d4b431bf91885cc20c6a1af68136f890"
)
EQ_BOOKMARK = "wpsc_eq_" + "e" * 24


def _numbering():
    return {
        "mode": "chapter",
        "sequenceId": "WPSC_EQ",
        "chapterStyleLevel": 1,
        "resetLevel": 1,
        "prefix": "(",
        "suffix": ")",
    }


def _manifest(resources) -> str:
    envelope = {
        "version": "1",
        "entries": [
            {
                "resourceId": resource.id,
                "sourceSha256": resource.source_sha256,
                "payloadSha256": resource.payload_sha256,
                "byteLength": len(resource.payload_bytes),
                "mediaType": resource.media_type,
                "normalizerId": resource.normalizer_id,
            }
            for resource in resources
        ],
    }
    raw = json.dumps(
        envelope, sort_keys=True, separators=(",", ":"), ensure_ascii=False
    ).encode()
    return "sha256:" + hashlib.sha256(raw).hexdigest()


def _resource(payload=b"formula-fallback"):
    return PreparedLongformResource(
        id="formula-image-1",
        media_type="image/png",
        source_sha256="1" * 64,
        payload_sha256=hashlib.sha256(payload).hexdigest(),
        normalizer_id="none-v1",
        payload_bytes=payload,
        image_profile=ImageProfile(300, 100, 144.0, 144.0, 1, "PNG", False),
    )


def _equation(*, fallback_resource=None, content=None):
    args = {
        "renderMode": "native-m4",
        "content": content
        or {
            "nativeMath": {
                "syntax": "wps-linear-v1",
                "linearText": "x+y",
                "sourceHash": "2" * 64,
            }
        },
        "numbering": _numbering(),
        "bookmarkName": EQ_BOOKMARK,
        "fallbackText": "x+y",
    }
    if fallback_resource is not None:
        args["fallbackResource"] = fallback_resource
    return GenerationOperation(
        "writer.add_equation",
        args,
        node_id="eq:one",
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
    )


def _anchor():
    return GenerationOperation(
        "writer.reserve_document_quality_anchor",
        {"title": "生成质量提示", "notices": []},
        node_id="doc:quality",
        failure_policy={"mode": "fail"},
    )


def _section(role: str, suffix: str):
    return GenerationOperation(
        "writer.configure_section", {"role": role}, node_id=f"doc:section:{suffix}"
    )


def _finalize():
    return GenerationOperation(
        "writer.finalize_fields", {"maxRounds": 3}, node_id="doc:finalize"
    )


def _plan(*body_ops, bibliography_ops=(), resources=()):
    operations = [_anchor(), _section("body", "body"), *body_ops]
    if bibliography_ops:
        operations.extend([_section("bibliography", "bibliography"), *bibliography_ops])
    operations.append(_finalize())
    return GenerationPlan(
        component="writer",
        operations=tuple(operations),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=(
            _manifest(resources) if resources else EMPTY_MANIFEST_DIGEST
        ),
    )


def _bibliography(entries=None):
    return GenerationOperation(
        "writer.add_bibliography",
        {
            "schemaVersion": 1,
            "entries": entries
            or [
                {
                    "id": "ref:a",
                    "nodeId": "ref:a",
                    "number": 1,
                    "text": "Alpha.",
                    "cited": True,
                },
                {
                    "id": "ref:b",
                    "nodeId": "ref:b",
                    "number": 2,
                    "text": "Beta.",
                    "cited": False,
                },
            ],
            "style": "numeric",
            "hangingIndentPt": 18.0,
            "leftIndentPt": 18.0,
            "spaceAfterPt": 6.0,
        },
        node_id="bib:block",
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
            "fallback": "notice",
        },
    )


def _citation():
    return GenerationOperation(
        "writer.add_cross_reference",
        {
            "runs": [
                {"type": "text", "text": "See "},
                {
                    "type": "citation",
                    "nodeId": "para:one/cite:1",
                    "targetId": "ref:a",
                    "targetNodeId": "ref:a",
                    "number": 1,
                    "fallbackText": "[1]",
                },
                {"type": "text", "text": " and again "},
                {
                    "type": "citation",
                    "nodeId": "para:one/cite:2",
                    "targetId": "ref:a",
                    "targetNodeId": "ref:a",
                    "number": 1,
                    "fallbackText": "[1]",
                },
            ]
        },
        node_id="para:one",
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    )


def test_executor_dispatches_native_formula_without_exposing_resource_locator(tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )

    executor.execute(
        _plan(
            _equation(fallback_resource={"fallbackResourceId": resource.id}),
            resources=(resource,),
        ),
        (resource,),
    )

    native = [kwargs for name, kwargs in composer.calls if name == "equation-native"]
    assert len(native) == 1
    assert native[0]["content"]["nativeMath"]["linearText"] == "x+y"
    assert "fallbackResource" not in native[0]
    assert "fallback_resource_locator" not in native[0]
    assert composer.resource_paths_seen == []


def test_formula_named_failure_rolls_back_then_uses_validated_image_once(tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    composer.failures["equation-native"] = "EQUATION_INSERT_FAILED"
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )

    outcome = executor.execute(
        _plan(
            _equation(fallback_resource={"fallbackResourceId": resource.id}),
            resources=(resource,),
        ),
        (resource,),
    )

    fallback = [kwargs for name, kwargs in composer.calls if name == "equation-fallback"]
    assert len(fallback) == 1
    assert fallback[0]["fallback_resource_locator"].endswith(".png")
    assert composer.rollback_tokens == composer.degradation_checkpoints
    assert [issue.code for issue in outcome.issues] == ["EQUATION_INSERT_FAILED"]
    assert all(str(tmp_path) not in issue.message for issue in outcome.issues)
    assert all(not Path(path).exists() for path in composer.resource_paths_seen)


def test_native_success_with_unavailable_optional_image_keeps_math_and_marks_once(tmp_path: Path):
    composer = RecordingNativeComposer()
    descriptor = {
        "code": "FORMULA_FALLBACK_IMAGE_UNAVAILABLE",
        "placement": "block",
        "objectLabel": "formula",
        "reason": "Fallback image is unavailable.",
        "fallbackText": "[FORMULA_FALLBACK_IMAGE_UNAVAILABLE 公式图像备选不可用]",
        "fallbackKind": "none",
    }
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )

    outcome = executor.execute(
        _plan(
            _equation(
                fallback_resource={"fallbackResourcePlannedDegradation": descriptor}
            )
        )
    )

    assert len([1 for name, _ in composer.calls if name == "equation-native"]) == 1
    assert not [1 for name, _ in composer.calls if name == "equation-fallback"]
    notices = [kwargs for name, kwargs in composer.calls if name == "notice"]
    assert [notice["code"] for notice in notices] == [
        "FORMULA_FALLBACK_IMAGE_UNAVAILABLE"
    ]
    assert [issue.code for issue in outcome.issues] == [
        "FORMULA_FALLBACK_IMAGE_UNAVAILABLE"
    ]
    assert composer.resource_paths_seen == []


def test_formula_without_image_uses_one_controller_owned_source_fallback(tmp_path: Path):
    composer = RecordingNativeComposer()
    composer.failures["equation-native"] = "EQUATION_INSERT_FAILED"

    outcome = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    ).execute(_plan(_equation()))

    fallback = [kwargs for name, kwargs in composer.calls if name == "equation-fallback"]
    assert len(fallback) == 1
    assert fallback[0]["fallback_resource_locator"] is None
    assert fallback[0]["fallbackText"] == "x+y"
    assert [issue.code for issue in outcome.issues] == ["EQUATION_INSERT_FAILED"]


def test_formula_resource_id_must_resolve_before_com_acquisition(tmp_path: Path):
    acquired = False

    def factory():
        nonlocal acquired
        acquired = True
        return RecordingNativeComposer()

    executor = WindowsLongformExecutor(staging_dir=str(tmp_path), composer_factory=factory)
    with pytest.raises(WindowsLongformExecutorError, match="resource validation"):
        executor.execute(
            _plan(
                _equation(
                    fallback_resource={"fallbackResourceId": "formula-image-1"}
                )
            )
        )
    assert acquired is False


def test_planned_formula_content_receives_its_validated_image_locator(tmp_path: Path):
    resource = _resource()
    composer = RecordingNativeComposer()
    planned = {
        "code": "FORMULA_MALFORMED",
        "placement": "block",
        "objectLabel": "formula",
        "reason": "Formula is malformed.",
        "fallbackText": "x+y",
        "fallbackKind": "source",
    }

    WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    ).execute(
        _plan(
            _equation(
                content={"plannedDegradation": planned},
                fallback_resource={"fallbackResourceId": resource.id},
            ),
            resources=(resource,),
        ),
        (resource,),
    )

    native = [kwargs for name, kwargs in composer.calls if name == "equation-native"]
    assert len(native) == 1
    assert native[0]["fallback_resource_locator"].endswith(".png")
    assert all(not Path(path).exists() for path in composer.resource_paths_seen)


def test_citations_reuse_static_number_in_one_owning_paragraph_and_bibliography_is_native(
    tmp_path: Path,
):
    composer = RecordingNativeComposer()
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )

    executor.execute(
        _plan(_citation(), bibliography_ops=(_bibliography(),))
    )

    citation = [kwargs for name, kwargs in composer.calls if name == "citation"]
    assert len(citation) == 1
    assert [run.get("fallbackText") for run in citation[0]["runs"] if run["type"] == "citation"] == ["[1]", "[1]"]
    bibliography = [kwargs for name, kwargs in composer.calls if name == "bibliography-native"]
    assert [entry["number"] for entry in bibliography[0]["entries"]] == [1, 2]
    assert not [name for name, _ in composer.calls if name == "notice"]


def test_structured_bibliography_named_failure_uses_one_notice_and_keeps_all_text(
    tmp_path: Path,
):
    composer = RecordingNativeComposer()
    composer.failures["bibliography-native"] = "BIBLIOGRAPHY_INSERT_FAILED"
    bibliography = _bibliography(entries=[{
        "id": "ref:a", "nodeId": "ref:a", "number": 1,
        "text": "Alpha.", "cited": False,
    }])

    outcome = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    ).execute(_plan(bibliography_ops=(bibliography,)))

    notices = [kwargs for name, kwargs in composer.calls if name == "notice"]
    assert len(notices) == 1
    assert notices[0]["fallback_text"] == "[1] Alpha."
    assert composer.rollback_tokens == composer.degradation_checkpoints
    assert [issue.code for issue in outcome.issues] == [
        "BIBLIOGRAPHY_INSERT_FAILED"
    ]


def test_m3_equation_and_legacy_bibliography_use_legacy_primitives(tmp_path: Path):
    composer = RecordingNativeComposer()
    equation = GenerationOperation(
        "writer.add_equation",
        {
            "source": "E=mc^2",
            "numbering": _numbering(),
            "bookmarkName": EQ_BOOKMARK,
            "fallbackText": "E=mc^2",
        },
        node_id="eq:legacy",
        failure_policy={"mode": "fail"},
    )
    bibliography = GenerationOperation(
        "writer.add_bibliography",
        {"entries": ["Legacy one.", "Legacy two."], "style": "numbered"},
        node_id="bib:legacy",
    )
    plan = GenerationPlan(
        component="writer",
        operations=(equation, bibliography, _finalize()),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=EMPTY_MANIFEST_DIGEST,
    )

    WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    ).execute(plan)

    assert len([1 for name, _ in composer.calls if name == "equation"]) == 1
    assert not [1 for name, _ in composer.calls if name == "equation-native"]
    legacy = [kwargs for name, kwargs in composer.calls if name == "bibliography-legacy"]
    assert legacy[0]["entries"] == ["Legacy one.", "Legacy two."]


class _PF:
    def __init__(self):
        self.Alignment = None
        self.KeepTogether = None
        self.KeepWithNext = None
        self.LeftIndent = None
        self.FirstLineIndent = None
        self.SpaceBefore = None
        self.SpaceAfter = None


class _Range:
    def __init__(self, start=0, end=0):
        self.Start = start
        self.End = end
        self.ParagraphFormat = _PF()
        self.Text = ""
        self.OMaths = _OMathView(()) if "_OMathView" in globals() else None

    def Collapse(self, direction):
        self.Start = self.End


class _Cell:
    def __init__(self, start, end):
        self.Range = _Range(start, end)
        self.VerticalAlignment = None


class _Borders:
    def __init__(self):
        self.items = {index: type("Border", (), {"LineStyle": None})() for index in range(-6, 0)}

    def __call__(self, index):
        return self.items[index]


class _Columns:
    def __init__(self):
        self.widths = {}

    def __call__(self, index):
        parent = self

        class Column:
            def SetWidth(self, width, rule):
                parent.widths[index] = (width, rule)

        return Column()


class _FormulaTable:
    def __init__(self):
        self.cells = {1: _Cell(10, 20), 2: _Cell(20, 80), 3: _Cell(80, 95)}
        self.Range = _Range(10, 96)
        self.Borders = _Borders()
        self.Columns = _Columns()
        self.Rows = type("Rows", (), {"AllowBreakAcrossPages": None})()

    def Cell(self, row, column):
        assert row == 1
        return self.cells[column]


class _Selection:
    def __init__(self):
        self.pos = 0
        self.typed = []
        self.ParagraphFormat = _PF()

    @property
    def Range(self):
        return _Range(self.pos, self.pos)

    @property
    def End(self):
        return self.pos

    def SetRange(self, start, end):
        assert start == end
        self.pos = end

    def TypeText(self, text):
        self.typed.append(text)
        self.pos += len(text)

    def TypeParagraph(self):
        self.typed.append("\n")
        self.pos += 1


class _OMath:
    def __init__(self, rng):
        self.Range = _Range(rng.Start, rng.End)
        self.built = 0

    def BuildUp(self):
        self.built += 1


class _OMathView:
    def __init__(self, items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        return self.items[index - 1]


class _OMaths:
    def __init__(self):
        self.items = []
        self.added_ranges = []

    @property
    def Count(self):
        return len(self.items)

    def Add(self, rng):
        self.added_ranges.append((rng.Start, rng.End, rng.Text))
        item = _OMath(rng)
        self.items.append(item)
        added_range = _Range(rng.Start, rng.End)
        added_range.Text = rng.Text
        added_range.OMaths = _OMathView((item,))
        return added_range

    def Item(self, index):
        return self.items[index - 1]


def _formula_writer():
    table = _FormulaTable()
    selection = _Selection()
    document = type("Doc", (), {})()
    document.OMaths = _OMaths()
    document.Tables = type("Tables", (), {"Add": lambda self, rng, rows, cols: table})()
    document.Content = type("Content", (), {"End": 96})()
    document.Range = lambda start, end: _Range(start, end)
    writer = WriterComposer.__new__(WriterComposer)
    writer._doc = document
    writer._selection = selection
    writer._app = type("App", (), {"Selection": selection})()
    writer._add_native_number_shell = lambda numbering, bookmark, owner: selection.TypeText("(1)")
    return writer, table, document


def test_writer_native_formula_uses_borderless_keep_together_omath_and_right_number_shell():
    writer, table, document = _formula_writer()

    outcome = writer.add_equation_native(
        renderMode="native-m4",
        content={
            "nativeMath": {
                "syntax": "wps-linear-v1",
                "linearText": "x+y",
                "sourceHash": "2" * 64,
            }
        },
        numbering=_numbering(),
        bookmarkName=EQ_BOOKMARK,
        fallbackText="x+y",
        owner_node_id="eq:one",
    )

    assert outcome == {"issues": []}
    assert document.OMaths.Count == 1
    assert document.OMaths.items[0].built == 1
    assert document.OMaths.added_ranges == [(20, 79, "x+y")]
    assert all(border.LineStyle == 0 for border in table.Borders.items.values())
    assert table.Range.ParagraphFormat.KeepTogether == -1
    assert table.Rows.AllowBreakAcrossPages == 0
    assert table.Cell(1, 2).Range.ParagraphFormat.Alignment == 1
    assert table.Cell(1, 3).Range.ParagraphFormat.Alignment == 2
    assert "(1)" in writer.selection.typed


def test_writer_planned_formula_content_skips_omath_and_keeps_numbered_fallback_shell():
    writer, _table, document = _formula_writer()
    calls = []
    writer.add_equation_native_fallback = lambda **kwargs: calls.append(kwargs) or {"issues": []}
    planned = {
        "code": "FORMULA_MALFORMED",
        "placement": "block",
        "objectLabel": "formula",
        "reason": "Formula is malformed.",
        "fallbackText": "x+y",
        "fallbackKind": "source",
    }

    outcome = writer.add_equation_native(
        renderMode="native-m4",
        content={"plannedDegradation": planned},
        numbering=_numbering(), bookmarkName=EQ_BOOKMARK,
        fallbackText="x+y", owner_node_id="eq:one", controller_owned=True,
        fallback_resource_locator="C:/validated/formula.png",
    )

    assert document.OMaths.Count == 0
    assert calls[0]["numbering"] == _numbering()
    assert calls[0]["bookmarkName"] == EQ_BOOKMARK
    assert calls[0]["fallback_resource_locator"] == "C:/validated/formula.png"
    assert outcome["issues"] == [{
        "code": "FORMULA_MALFORMED",
        "message": "Formula is malformed.",
        "placement": "block",
        "fallback": "source",
    }]


def test_writer_omath_count_verification_failure_is_exact_recoverable_code():
    writer, _table, document = _formula_writer()

    class BadCountOMaths(_OMaths):
        def Add(self, rng):
            item = super().Add(rng)
            self.items.append(_OMath(rng))
            return item

    document.OMaths = BadCountOMaths()
    with pytest.raises(NativeWriterObjectError) as caught:
        writer.add_equation_native(
            renderMode="native-m4",
            content={"nativeMath": {
                "syntax": "wps-linear-v1", "linearText": "x+y",
                "sourceHash": "2" * 64,
            }},
            numbering=_numbering(), bookmarkName=EQ_BOOKMARK,
            fallbackText="x+y", owner_node_id="eq:one", controller_owned=True,
        )
    assert caught.value.code == "EQUATION_INSERT_FAILED"


@pytest.mark.parametrize("boundary", ["count", "add", "buildup", "range"])
def test_writer_unknown_omath_boundary_errors_remain_fatal(boundary):
    writer, _table, document = _formula_writer()

    class UnknownComError(RuntimeError):
        pass

    if boundary == "count":
        class BadCount:
            @property
            def Count(self):
                raise UnknownComError("unknown count failure")
        document.OMaths = BadCount()
    elif boundary == "add":
        class BadAdd(_OMaths):
            def Add(self, rng):
                raise UnknownComError("unknown add failure")
        document.OMaths = BadAdd()
    elif boundary == "buildup":
        class BadBuild(_OMath):
            def BuildUp(self):
                raise UnknownComError("unknown buildup failure")

        class BadBuildCollection(_OMaths):
            def Add(self, rng):
                self.added_ranges.append((rng.Start, rng.End, rng.Text))
                item = BadBuild(rng)
                self.items.append(item)
                added_range = _Range(rng.Start, rng.End)
                added_range.OMaths = _OMathView((item,))
                return added_range
        document.OMaths = BadBuildCollection()
    else:
        class BadRangeOMath(_OMath):
            @property
            def Range(self):
                raise UnknownComError("unknown range failure")

            @Range.setter
            def Range(self, value):
                pass

        class BadRangeCollection(_OMaths):
            def Add(self, rng):
                item = BadRangeOMath(rng)
                self.items.append(item)
                added_range = _Range(rng.Start, rng.End)
                added_range.OMaths = _OMathView((item,))
                return added_range
        document.OMaths = BadRangeCollection()

    with pytest.raises(UnknownComError):
        writer.add_equation_native(
            renderMode="native-m4",
            content={"nativeMath": {
                "syntax": "wps-linear-v1", "linearText": "x+y",
                "sourceHash": "2" * 64,
            }},
            numbering=_numbering(), bookmarkName=EQ_BOOKMARK,
            fallbackText="x+y", owner_node_id="eq:one", controller_owned=True,
        )


def test_writer_formula_image_failure_is_attempted_once_then_source_is_inside_terminal_notice():
    writer, _table, _document = _formula_writer()
    attempts = []
    rollbacks = []
    inline = []

    def fail_image(locator, **kwargs):
        attempts.append(locator)
        raise RuntimeError("private path must not escape")

    writer.add_image = fail_image
    writer._native_rollback = lambda start, end: rollbacks.append((start, end))
    writer.add_inline_degradation = (
        lambda code, message, fallback_text: inline.append((code, fallback_text))
    )

    writer.add_equation_native_fallback(
        numbering=_numbering(),
        bookmarkName=EQ_BOOKMARK,
        fallbackText="x+y",
        fallback_resource_locator="C:/private/formula.png",
        owner_node_id="eq:one",
    )

    assert attempts == ["C:/private/formula.png"]
    assert len(rollbacks) == 1
    assert inline == [("EQUATION_INSERT_FAILED", "x+y")]
    assert "(1)" in writer.selection.typed


@pytest.mark.parametrize("boundary", ["none", "range", "alignment", "keep"])
def test_writer_formula_image_postprocess_failure_rolls_back_to_source_once(boundary):
    writer, _table, _document = _formula_writer()
    rollbacks = []
    inline = []

    class FailingFormat(_PF):
        def __setattr__(self, name, value):
            if name == "Alignment" and boundary == "alignment":
                raise RuntimeError("unknown alignment failure")
            if name == "KeepTogether" and boundary == "keep":
                raise RuntimeError("unknown keep failure")
            object.__setattr__(self, name, value)

    class FailingRangeShape:
        @property
        def Range(self):
            if boundary == "range":
                raise RuntimeError("unknown shape range failure")
            rng = _Range(20, 30)
            rng.ParagraphFormat = FailingFormat()
            return rng

    writer.add_image = lambda *args, **kwargs: (
        None if boundary == "none" else FailingRangeShape()
    )
    writer._native_rollback = lambda start, end: rollbacks.append((start, end))
    writer.add_inline_degradation = (
        lambda code, message, fallback_text: inline.append((code, fallback_text))
    )

    writer.add_equation_native_fallback(
        numbering=_numbering(), bookmarkName=EQ_BOOKMARK,
        fallbackText="x+y", fallback_resource_locator="C:/validated/formula.png",
        owner_node_id="eq:one",
    )

    assert len(rollbacks) == 1
    assert inline == [("EQUATION_INSERT_FAILED", "x+y")]


def test_writer_formula_image_rollback_failure_remains_fatal():
    writer, _table, _document = _formula_writer()
    writer.add_image = lambda *args, **kwargs: None

    def fail_rollback(start, end):
        raise NativeWriterObjectError("LOCAL_MUTATION_ROLLBACK_FAILED")

    writer._native_rollback = fail_rollback
    with pytest.raises(NativeWriterObjectError) as caught:
        writer.add_equation_native_fallback(
            numbering=_numbering(), bookmarkName=EQ_BOOKMARK,
            fallbackText="x+y", fallback_resource_locator="C:/validated/formula.png",
            owner_node_id="eq:one",
        )
    assert caught.value.code == "LOCAL_MUTATION_ROLLBACK_FAILED"


def test_writer_figure_fallback_initializes_empty_issues_and_keeps_planned_order():
    writer = WriterComposer.__new__(WriterComposer)
    notices = []
    writer.add_degradation_notice = lambda code, message, fallback, placement: notices.append(code)

    assert writer.add_captioned_figure_fallback(children=[]) == {"issues": []}
    outcome = writer.add_captioned_figure_fallback(children=[
        {"plannedDegradation": {
            "code": "RESOURCE_NOT_FOUND", "message": "first",
            "fallback": "[first]", "placement": "block",
        }},
        {"plannedDegradation": {
            "code": "RESOURCE_NORMALIZATION_FAILED", "message": "second",
            "fallback": "[second]", "placement": "block",
        }},
    ])

    assert notices == ["RESOURCE_NOT_FOUND", "RESOURCE_NORMALIZATION_FAILED"]
    assert [issue["code"] for issue in outcome["issues"]] == notices


def test_writer_citation_and_bibliography_preserve_runs_order_and_fixed_paragraph_geometry():
    selection = _Selection()
    ranges = []

    class Doc:
        def Range(self, start, end):
            rng = _Range(start, end)
            ranges.append(rng)
            return rng

    writer = WriterComposer.__new__(WriterComposer)
    writer._selection = selection
    writer._app = type("App", (), {"Selection": selection})()
    writer._doc = Doc()

    writer.add_citation_paragraph(
        runs=[
            {"type": "text", "text": "See "},
            {
                "type": "citation",
                "nodeId": "p/c:1",
                "targetId": "a",
                "targetNodeId": "ref:a",
                "number": 1,
                "fallbackText": "[1]",
            },
            {"type": "text", "text": "."},
        ],
        owner_node_id="p",
    )
    writer.add_bibliography_native(
        entries=[
            {"id": "a", "nodeId": "ref:a", "number": 1, "text": "Alpha.", "cited": True},
            {"id": "b", "nodeId": "ref:b", "number": 2, "text": "Beta.", "cited": False},
        ],
        style="numeric",
        hangingIndentPt=18.0,
        leftIndentPt=18.0,
        spaceAfterPt=6.0,
        owner_node_id="bib:block",
    )

    assert selection.typed == ["See ", "[1]", ".", "\n", "[1] Alpha.", "\n", "[2] Beta.", "\n"]
    bibliography_ranges = ranges[-2:]
    assert [rng.ParagraphFormat.Alignment for rng in bibliography_ranges] == [0, 0]
    assert [rng.ParagraphFormat.LeftIndent for rng in bibliography_ranges] == [18.0, 18.0]
    assert [rng.ParagraphFormat.FirstLineIndent for rng in bibliography_ranges] == [-18.0, -18.0]
    assert [rng.ParagraphFormat.SpaceAfter for rng in bibliography_ranges] == [6.0, 6.0]


@pytest.mark.parametrize(
    "code",
    ["COM_ERROR", "FIELD_INSERT_FAILED", "BOOKMARK_FAILED", "LOCAL_MUTATION_ROLLBACK_FAILED"],
)
def test_executor_formula_non_allowlisted_boundaries_are_fatal(code: str, tmp_path: Path):
    composer = RecordingNativeComposer()
    composer.failures["equation-native"] = code
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )
    with pytest.raises(WindowsLongformExecutorError):
        executor.execute(_plan(_equation()))
    assert not [name for name, _ in composer.calls if name == "equation-fallback"]
