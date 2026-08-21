"""Tests for the M1 long-form generation plan builder and protocol v2."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from skills.WPSComposer.scripts.document_model import (
    DegradationBlock,
    DocumentIssue,
    FigureBlock,
    FormulaBlock,
    ImageBlock,
    KeywordsBlock,
    ListBlock,
    Paragraph,
    ReferenceListBlock,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.generation_plan import (
    GenerationPlan,
    OperationPlanError,
    validate_generation_plan,
)
from skills.WPSComposer.scripts.longform.plan import build_longform_plan
from skills.WPSComposer.scripts.longform.policy import LongformPolicy, build_policy
from skills.WPSComposer.scripts.longform.resources import preflight_resources
from skills.WPSComposer.scripts.longform.semantic import (
    BookmarkMapResult,
    LongformConfig,
    SemanticResult,
)


def make_config(**kwargs) -> LongformConfig:
    defaults = {
        "title": "默认标题",
        "short_title": "默认",
        "author": "作者",
        "date": "2026-08-21",
        "header": "默认页眉",
        "title_page": False,
        "toc": False,
        "figure_index": False,
        "table_index": False,
        "bibliography_include_uncited": True,
        "caption_numbering": "global",
        "heading_numbering": "none",
        "layout_engine": "longform",
    }
    defaults.update(kwargs)
    return LongformConfig(**defaults)


def make_document(title="默认标题", sections=None, abstract=None, keywords=None):
    return StructuredDocument(
        title=title,
        metadata={},
        sections=sections or [],
        config={},
        longform=True,
        issues=[],
        abstract=abstract,
        keywords=keywords,
    )


def make_semantic(
    title="默认标题",
    sections=None,
    config=None,
    references=None,
    issues=(),
) -> SemanticResult:
    doc = make_document(title=title, sections=sections or [])
    return SemanticResult(
        document=doc,
        config=config or make_config(title=title),
        references=references or {},
        bookmarks=BookmarkMapResult(mapping={}, issues=()),
        issues=issues,
    )


def serialize_plan(plan: GenerationPlan) -> bytes:
    return json.dumps(plan.to_dict(), ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def test_build_policy_returns_stable_defaults():
    config = make_config()
    policy = build_policy(config)
    assert isinstance(policy, LongformPolicy)
    assert policy.page_size == "A4"
    assert policy.page_margins["left_mm"] == 30.0
    assert policy.body_font["cjk"] == "宋体"


def test_build_longform_plan_returns_protocol_v2_envelope():
    semantic = make_semantic(title="Test")
    preflight = preflight_resources([], ".")
    plan = build_longform_plan(semantic, preflight)
    envelope = plan.to_dict()
    assert envelope["protocolVersion"] == 2
    assert envelope["semanticVersion"] == "longform-1"
    assert envelope["component"] == "writer"
    assert envelope["resourceManifestVersion"] == 1
    assert "resourceManifestDigest" in envelope
    assert "operations" in envelope
    assert len(envelope["operations"]) > 0


def test_build_longform_plan_is_deterministic():
    semantic = make_semantic(
        title="稳定性测试",
        sections=[
            Section(
                level=1,
                heading="第一章 绪论",
                elements=[Paragraph(spans=[Span(text="正文段落。")])],
            ),
        ],
    )
    preflight = preflight_resources([], ".")
    first = serialize_plan(build_longform_plan(semantic, preflight))
    second = serialize_plan(build_longform_plan(semantic, preflight))
    assert first == second


def test_build_longform_plan_operation_order():
    semantic = make_semantic(
        title="操作顺序测试",
        config=make_config(toc=True, figure_index=True, table_index=True),
        sections=[
            Section(
                level=1,
                heading="章节",
                elements=[
                    Paragraph(spans=[Span(text="段落。")]),
                    FigureBlock(
                        identifier="fig:sample",
                        node_id="fig:sample",
                        caption="示例图",
                        images=[ImageBlock(path="sample.png", alt="样例")],
                        layout="stack",
                    ),
                    SemanticTableBlock(
                        identifier="tab:sample",
                        node_id="tab:sample",
                        caption="示例表",
                        headers=["列1"],
                        rows=[["行1"]],
                        alignments=["left"],
                    ),
                    FormulaBlock(
                        identifier="eq:sample",
                        node_id="eq:sample",
                        source="E=mc^2",
                        number="1",
                    ),
                    ReferenceListBlock(
                        node_id="ref:list",
                        entries=["- id: chen2025 | text: 陈. 示例[J]. 2025."],
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources(semantic.document.sections, ".")
    plan = build_longform_plan(semantic, preflight)
    ops = [op["op"] for op in plan.to_dict()["operations"]]

    assert ops[0] == "writer.reset"
    assert ops[-1] == "writer.finalize_fields"
    assert ops.index("writer.configure_page") < ops.index("writer.add_heading")
    assert ops.index("writer.ensure_styles") < ops.index("writer.add_heading")
    assert ops.index("writer.configure_front_matter") < ops.index("writer.add_heading")
    assert ops.index("writer.configure_toc_styles") < ops.index("writer.insert_toc")
    assert ops.index("writer.add_heading") < ops.index("writer.add_paragraph")
    assert ops.index("writer.add_paragraph") < ops.index("writer.add_captioned_figure")
    assert ops.index("writer.add_captioned_figure") < ops.index("writer.add_semantic_table")
    assert ops.index("writer.add_semantic_table") < ops.index("writer.add_equation")
    assert ops.index("writer.add_equation") < ops.index("writer.add_bibliography")
    assert ops.index("writer.insert_toc") < ops.index("writer.finalize_fields")
    assert ops.index("writer.insert_figure_index") < ops.index("writer.finalize_fields")
    assert ops.index("writer.insert_table_index") < ops.index("writer.finalize_fields")


def test_resource_manifest_digest_binding(tmp_path):
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake-png")
    semantic = make_semantic(
        title="资源绑定测试",
        sections=[
            Section(
                level=1,
                heading="章节",
                elements=[
                    FigureBlock(
                        identifier="fig:res",
                        node_id="fig:res",
                        caption="资源图",
                        images=[ImageBlock(path=str(image_path), alt="样例")],
                        layout="stack",
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    plan = build_longform_plan(semantic, preflight)
    envelope = plan.to_dict()
    assert envelope["resourceManifestDigest"] == preflight.manifest["digest"]
    assert envelope["resourceManifestVersion"] == int(preflight.manifest["version"])


def test_no_absolute_paths_in_plan_args(tmp_path):
    import re
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake-png")
    outside_path = "/etc/passwd"
    traversal_path = "../outside.png"
    missing_path = str(tmp_path / "missing.png")
    semantic = make_semantic(
        title="NoAbsPaths",
        sections=[
            Section(
                level=1,
                heading="Chapter",
                elements=[
                    FigureBlock(
                        identifier="fig:ok",
                        node_id="fig:ok",
                        caption="OK",
                        images=[ImageBlock(path=str(image_path), alt="sample")],
                        layout="stack",
                    ),
                    FigureBlock(
                        identifier="fig:outside",
                        node_id="fig:outside",
                        caption="Outside",
                        images=[ImageBlock(path=outside_path, alt="x")],
                        layout="stack",
                    ),
                    FigureBlock(
                        identifier="fig:traversal",
                        node_id="fig:traversal",
                        caption="Traversal",
                        images=[ImageBlock(path=traversal_path, alt="x")],
                        layout="stack",
                    ),
                    FigureBlock(
                        identifier="fig:missing",
                        node_id="fig:missing",
                        caption="Missing",
                        images=[ImageBlock(path=missing_path, alt="x")],
                        layout="stack",
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    plan = build_longform_plan(semantic, preflight)
    serialized = json.dumps(plan.to_dict(), ensure_ascii=False, separators=(",", ":"))

    # The absolute base_dir itself must never appear in the serialized plan.
    assert str(tmp_path) not in serialized

    # No Unix absolute filesystem path substring may leak.
    abs_unix = re.compile(r"/[A-Za-z0-9_.\-]+(?:/[A-Za-z0-9_.\-]+)+")
    for match in abs_unix.finditer(serialized):
        raise AssertionError(f"Unix absolute path leaked into plan: {match.group()!r}")

    # No Windows drive-letter absolute filesystem path substring may leak.
    abs_win = re.compile(r"[A-Za-z]:[\/][^\s\"'\\]+")
    for match in abs_win.finditer(serialized):
        raise AssertionError(f"Windows absolute path leaked into plan: {match.group()!r}")

    # Exact original path strings must not appear anywhere.
    for bad in (outside_path, traversal_path, missing_path):
        assert bad not in serialized


def test_stable_node_ids_for_explicit_and_internal_ids():
    semantic = make_semantic(
        title="节点 ID 测试",
        sections=[
            Section(
                level=1,
                heading="章节",
                elements=[
                    FigureBlock(
                        identifier="fig:explicit",
                        node_id="fig:explicit",
                        caption="显式图",
                        images=[ImageBlock(path="missing.png", alt="x")],
                        layout="stack",
                    ),
                    FigureBlock(
                        identifier=None,
                        node_id=None,
                        caption="内部图",
                        images=[ImageBlock(path="missing.png", alt="x")],
                        layout="stack",
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources(semantic.document.sections, ".")
    first = build_longform_plan(semantic, preflight)
    second = build_longform_plan(semantic, preflight)

    explicit_ops = [
        op for op in first.to_dict()["operations"]
        if op["op"] == "writer.add_captioned_figure"
    ]
    assert len(explicit_ops) == 2
    assert explicit_ops[0]["nodeId"] == "fig:explicit"
    internal_id_1 = explicit_ops[1]["nodeId"]
    internal_id_2 = [
        op for op in second.to_dict()["operations"]
        if op["op"] == "writer.add_captioned_figure"
    ][1]["nodeId"]
    assert internal_id_1 == internal_id_2
    assert not internal_id_1.startswith("__wpsc_")


def test_unknown_field_rejected_in_v2_operation():
    raw = {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": "sha256:" + "0" * 64,
        "operations": [
            {
                "op": "writer.configure_front_matter",
                "nodeId": "doc",
                "args": {"title": "T", "extraField": "bad"},
                "failurePolicy": {"mode": "fail"},
            },
        ],
    }
    with pytest.raises(OperationPlanError):
        validate_generation_plan(raw, "writer")


def test_missing_node_id_rejected_for_semantic_operation():
    raw = {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": "sha256:" + "0" * 64,
        "operations": [
            {
                "op": "writer.add_captioned_figure",
                "args": {
                    "caption": "图",
                    "children": [{"nodeId": "child", "resourceId": "r"}],
                    "layout": "stack",
                },
                "failurePolicy": {"mode": "degrade", "recoverableCodes": ["X"], "fallback": "notice"},
            },
        ],
    }
    with pytest.raises(OperationPlanError):
        validate_generation_plan(raw, "writer")


def test_failure_policy_degrade_requires_recoverable_codes_and_fallback():
    raw = {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": "sha256:" + "0" * 64,
        "operations": [
            {
                "op": "writer.add_captioned_figure",
                "nodeId": "fig",
                "args": {
                    "caption": "图",
                    "children": [{"nodeId": "child", "resourceId": "r"}],
                    "layout": "stack",
                },
                "failurePolicy": {"mode": "degrade"},
            },
        ],
    }
    with pytest.raises(OperationPlanError):
        validate_generation_plan(raw, "writer")


def test_recording_writer_composer_records_longform_operations():
    from skills.WPSComposer.scripts.recording_composers import RecordingWriterComposer

    composer = RecordingWriterComposer()
    with composer as writer:
        writer.configure_front_matter(title="标题", author="作者")
        writer.configure_section(landscape=True)
        writer.configure_toc_styles(tocTitle="目录", levels=3)
        writer.add_captioned_figure(
            node_id="fig:rec",
            caption="录制图",
            children=[{"nodeId": "fig:rec/image:1", "resourceId": "image-1"}],
            layout="stack",
            failure_policy={"mode": "degrade", "recoverableCodes": ["IMAGE_INSERT_FAILED"], "fallback": "notice"},
        )
        writer.add_semantic_table(
            node_id="tab:rec",
            caption="录制表",
            headers=["H"],
            rows=[["R"]],
        )
        writer.add_equation(
            node_id="eq:rec",
            source="E=mc^2",
            number="1",
            fallback_text="[公式]",
        )
        writer.add_cross_reference(
            node_id="ref:rec",
            target_id="fig:rec",
            kind="figure",
            fallback_text="[图]",
        )
        writer.insert_figure_index(title="图目录")
        writer.insert_table_index(title="表目录")
        writer.add_bibliography(node_id="bib:rec", entries=["[1] 文献"])
        writer.add_inline_degradation(
            node_id="deg:inline",
            code="INLINE_DEGRADED",
            message="inline degraded",
            fallback_text="[?]",
        )
        writer.add_degradation_notice(
            node_id="deg:block",
            code="BLOCK_DEGRADED",
            message="block degraded",
            fallback_text="[!]",
            placement="block",
        )
        writer.add_document_quality_notice(
            notices=[{"code": "QUALITY", "message": "q", "fallbackText": "[Q]", "placement": "document"}]
        )
        writer.finalize_fields()
        recorded = writer.save_docx("ignored.docx")

    names = [op.op for op in recorded.plan.operations]
    assert "writer.configure_front_matter" in names
    assert "writer.add_captioned_figure" in names
    assert "writer.add_document_quality_notice" in names
    assert "writer.finalize_fields" in names

    fig_op = [op for op in recorded.plan.operations if op.op == "writer.add_captioned_figure"][0]
    assert fig_op.node_id == "fig:rec"
    assert fig_op.failure_policy["mode"] == "degrade"


def test_build_longform_plan_emits_degradation_for_preflight_failure(tmp_path):
    semantic = make_semantic(
        title="降级测试",
        sections=[
            Section(
                level=1,
                heading="章节",
                elements=[
                    FigureBlock(
                        identifier="fig:missing",
                        node_id="fig:missing",
                        caption="缺失图",
                        images=[ImageBlock(path="/etc/passwd", alt="x")],
                        layout="stack",
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    plan = build_longform_plan(semantic, preflight)
    ops = plan.to_dict()["operations"]
    figure_op = [op for op in ops if op["op"] == "writer.add_captioned_figure"][0]
    assert figure_op["failurePolicy"]["mode"] == "degrade"
    children = figure_op["args"]["children"]
    assert len(children) == 1
    assert "plannedDegradation" in children[0]
    assert children[0]["plannedDegradation"]["code"] == "RESOURCE_ABSOLUTE_PATH_OUTSIDE"


def test_build_longform_plan_includes_document_quality_notices():
    issue = DocumentIssue(
        code="CONFIG_VALUE_INVALID",
        message="Invalid config.",
        placement="document",
    )
    semantic = make_semantic(title="质量提示测试", issues=(issue,))
    preflight = preflight_resources([], ".")
    plan = build_longform_plan(semantic, preflight)
    ops = plan.to_dict()["operations"]
    notice_ops = [op for op in ops if op["op"] == "writer.add_document_quality_notice"]
    assert len(notice_ops) == 1
    assert notice_ops[0]["args"]["notices"][0]["code"] == "CONFIG_VALUE_INVALID"


def test_build_longform_plan_includes_inline_degradation():
    issue = DocumentIssue(
        code="REFERENCE_UNRESOLVED",
        message="Missing ref.",
        placement="inline",
    )
    semantic = make_semantic(
        title="行内降级测试",
        sections=[
            Section(
                level=1,
                heading="章节",
                elements=[
                    DegradationBlock(
                        issue=issue,
                        node_id="deg:inline-1",
                        fallback_text="[?]",
                    ),
                ],
            ),
        ],
    )
    preflight = preflight_resources([], ".")
    plan = build_longform_plan(semantic, preflight)
    ops = plan.to_dict()["operations"]
    deg_ops = [op for op in ops if op["op"] == "writer.add_inline_degradation"]
    assert len(deg_ops) == 1
    assert deg_ops[0]["nodeId"] == "deg:inline-1"

