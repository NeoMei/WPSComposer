"""Build a closed, deterministic protocol v2 generation plan for a long-form document.

The plan builder converts a normalized semantic model and preflighted resource
manifest into a pure JSON operation plan.  It never reads files or launches WPS.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from ..document_model import (
    DegradationBlock,
    DocumentIssue,
    FigureBlock,
    FormulaBlock,
    ImageBlock,
    KeywordsBlock,
    ListBlock,
    MathBlock,
    PageBreakBlock,
    Paragraph,
    ReferenceListBlock,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
)
from ..generation_plan import GenerationOperation, GenerationPlan
from .page_policy import _group_sections, build_page_policy
from .policy import LongformPolicy, build_policy
from .resources import ResourcePreflight
from .semantic import SemanticResult


_LONGFORM_PROTOCOL_VERSION = 2
_LONGFORM_SEMANTIC_VERSION = "longform-1"
_RESOURCE_MANIFEST_VERSION = 1

_MM_TO_PT = 2.834645669


def _mm_to_pt(mm: float) -> float:
    return round(mm * _MM_TO_PT, 2)


def _span_text(spans: list[Span]) -> str:
    """Convert a list of spans to plain text."""
    return "".join(span.text for span in spans)


def _list_items(items: list[list[Span]]) -> list[str]:
    """Convert list item spans to plain text strings."""
    return [_span_text(spans) for spans in items]


def _keywords_text(keywords: KeywordsBlock) -> str:
    """Render keywords as a single prefixed paragraph."""
    parts = [str(k).strip() for k in keywords.keywords if str(k).strip()]
    if not parts:
        return ""
    return "关键词：" + "；".join(parts)


@dataclass
class _BuilderState:
    """Mutable builder state that is not part of the public plan."""

    operations: list[GenerationOperation] = field(default_factory=list)
    figure_count: int = 0
    table_count: int = 0
    equation_count: int = 0
    has_figures: bool = False
    has_tables: bool = False
    title_displayed: bool = False
    front_matter_figure_index_emitted: bool = False
    front_matter_table_index_emitted: bool = False

    def add(
        self,
        op: str,
        args: dict[str, Any],
        *,
        node_id: Optional[str] = None,
        failure_policy: Optional[dict[str, Any]] = None,
    ) -> None:
        self.operations.append(
            GenerationOperation(
                op=op,
                args=args,
                node_id=node_id,
                failure_policy=failure_policy,
            )
        )


def build_longform_plan(
    semantic: SemanticResult,
    preflight: ResourcePreflight,
) -> GenerationPlan:
    """Build a protocol v2 GenerationPlan from a semantic result and preflight data."""
    policy = build_policy(semantic.config)
    state = _BuilderState()

    _build_begin(state, semantic.document, policy)
    _build_page_skeleton(state, semantic.document, semantic.config, policy, preflight)
    _build_indexes(state, policy)
    _build_quality_notices(state, semantic.issues)
    _build_finalize(state)

    digest = preflight.manifest.get("digest", "sha256:" + "0" * 64)
    version = preflight.manifest.get("version", str(_RESOURCE_MANIFEST_VERSION))

    return GenerationPlan(
        component="writer",
        operations=tuple(state.operations),
        protocol_version=_LONGFORM_PROTOCOL_VERSION,
        semantic_version=_LONGFORM_SEMANTIC_VERSION,
        resource_manifest_version=int(version),
        resource_manifest_digest=digest,
    )


def _build_begin(state: _BuilderState, document: StructuredDocument, policy: LongformPolicy) -> None:
    state.add("writer.reset", {})
    state.add(
        "writer.configure_page",
        {
            "marginTop": _mm_to_pt(policy.page_margins["top_mm"]),
            "marginBottom": _mm_to_pt(policy.page_margins["bottom_mm"]),
            "marginLeft": _mm_to_pt(policy.page_margins["left_mm"]),
            "marginRight": _mm_to_pt(policy.page_margins["right_mm"]),
        },
    )
    state.add(
        "writer.ensure_styles",
        {
            "styles": [
                {
                    "name": "Body Text",
                    "type": "paragraph",
                    "fontName": policy.body_font["cjk"],
                    "fontNameAscii": policy.latin_font,
                    "fontSize": policy.body_size_pt,
                    "lineSpacing": policy.line_spacing,
                    "align": 3,
                },
                {
                    "name": "Title",
                    "type": "paragraph",
                    "fontName": policy.heading_font["cjk"],
                    "fontNameAscii": policy.latin_font,
                    "fontSize": max(policy.heading_size_pt + 4, 16),
                    "align": 1,
                    "bold": True,
                },
                *[
                    {
                        "name": f"Heading {level}",
                        "type": "paragraph",
                        "fontName": policy.heading_font["cjk"],
                        "fontNameAscii": policy.latin_font,
                        "fontSize": max(policy.heading_size_pt - (level - 1), 12),
                        "outlineLevel": level,
                    }
                    for level in range(1, 7)
                ],
            ],
        },
    )
    state.add(
        "writer.configure_front_matter",
        {
            "title": policy.title,
            "shortTitle": policy.short_title,
            "author": policy.author,
            "date": policy.date,
            "header": policy.header,
            "titlePage": policy.title_page,
        },
        node_id="doc:front-matter",
    )
    state.add(
        "writer.configure_toc_styles",
        {
            "tocTitle": policy.toc_title,
            "levels": policy.toc_levels,
            "includeFigureIndex": policy.figure_index,
            "includeTableIndex": policy.table_index,
            "figureIndexTitle": policy.figure_index_title,
            "tableIndexTitle": policy.table_index_title,
            "minFontSizePt": policy.toc_density["min_font_size_pt"],
            "minSpaceBeforePt": policy.toc_density["min_space_before_pt"],
            "minSpaceAfterPt": policy.toc_density["min_space_after_pt"],
        },
        node_id="doc:toc",
    )


def _emit_configure_section(
    state: _BuilderState,
    index: int,
    section_policy: Any,
    *,
    landscape: bool = False,
) -> None:
    args: dict[str, Any] = {
        "role": section_policy.role,
        "pageNumberFormat": section_policy.page_number_format,
        "restartPageNumbering": section_policy.restart_numbering,
        "headerText": section_policy.header_text,
        "footerText": section_policy.footer_text,
        "linkToPreviousHeader": section_policy.link_to_previous_header,
        "linkToPreviousFooter": section_policy.link_to_previous_footer,
    }
    if section_policy.start_page_number is not None:
        args["startPageNumber"] = section_policy.start_page_number
    if landscape:
        args["landscape"] = True
    state.add(
        "writer.configure_section",
        args,
        node_id=f"doc:section-{index}:{section_policy.role}",
    )




def _contains_figure(node: Any) -> bool:
    if isinstance(node, FigureBlock):
        return True
    if isinstance(node, Section):
        return any(_contains_figure(child) for child in node.elements)
    return False


def _contains_table(node: Any) -> bool:
    if isinstance(node, SemanticTableBlock):
        return True
    if isinstance(node, Section):
        return any(_contains_table(child) for child in node.elements)
    return False


def _scan_document_indexes(document: StructuredDocument) -> tuple[bool, bool]:
    has_figures = any(_contains_figure(section) for section in document.sections)
    has_tables = any(_contains_table(section) for section in document.sections)
    return has_figures, has_tables

def _build_page_skeleton(
    state: _BuilderState,
    document: StructuredDocument,
    config: Any,
    policy: LongformPolicy,
    preflight: ResourcePreflight,
) -> None:
    """Emit configure_section operations and render each role's content."""
    skeleton = build_page_policy(document, config, policy)
    groups = _group_sections(document.sections)
    group_iter = iter(groups)

    state.has_figures, state.has_tables = _scan_document_indexes(document)

    for index, section_policy in enumerate(skeleton.sections):
        if section_policy.role == "cover":
            _emit_configure_section(state, index, section_policy, landscape=False)
            continue

        if section_policy.role == "front_matter":
            _emit_configure_section(state, index, section_policy, landscape=False)
            _render_front_matter(state, document, section_policy, policy)
            continue

        # Body or landscape: consume the next grouped content section.
        role, sections = next(group_iter)
        if not sections:
            continue
        _emit_configure_section(state, index, section_policy, landscape=(role == "landscape"))

        if role == "body" and not state.title_displayed and document.title_display and not policy.title_page:
            state.add(
                "writer.add_paragraph",
                {"text": document.title_display.plain_text, "style": "Title"},
                node_id="doc:title-display",
            )
            state.title_displayed = True

        for section in sections:
            _render_section(state, section, policy, preflight)


def _render_front_matter(
    state: _BuilderState,
    document: StructuredDocument,
    section_policy: Any,
    policy: LongformPolicy,
) -> None:
    if section_policy.includes_abstract and document.abstract:
        for paragraph in document.abstract.paragraphs:
            text = paragraph.plain_text
            if text:
                state.add(
                    "writer.add_paragraph",
                    {"text": text, "style": "Body Text"},
                    node_id="doc:abstract",
                )
    if section_policy.includes_keywords and document.keywords:
        text = _keywords_text(document.keywords)
        if text:
            state.add(
                "writer.add_paragraph",
                {"text": text, "style": "Body Text"},
                node_id="doc:keywords",
            )
    if section_policy.includes_toc:
        state.add(
            "writer.insert_toc",
            {"title": policy.toc_title},
            node_id="doc:toc",
        )
    if section_policy.includes_figure_index and state.has_figures:
        state.add(
            "writer.insert_figure_index",
            {"title": policy.figure_index_title},
            node_id="doc:figure-index",
        )
        state.front_matter_figure_index_emitted = True
    if section_policy.includes_table_index and state.has_tables:
        state.add(
            "writer.insert_table_index",
            {"title": policy.table_index_title},
            node_id="doc:table-index",
        )
        state.front_matter_table_index_emitted = True


def _render_section(
    state: _BuilderState,
    section: Section,
    policy: LongformPolicy,
    preflight: ResourcePreflight,
) -> None:
    if not section.has_heading and not section.elements:
        return

    if section.has_heading:
        heading_args: dict[str, Any] = {
            "text": section.heading,
            "level": section.level,
        }
        if (
            section.level in {1, 2, 3, 4}
            and section.numbering not in ("none", "auto")
            and section.numbering_scheme
        ):
            heading_args["numbering"] = True
            heading_args["numberingScheme"] = section.numbering_scheme
        state.add(
            "writer.add_heading",
            heading_args,
            node_id=section.node_id,
        )

    for element in section.elements:
        _render_element(state, element, section.node_id, policy, preflight)


def _render_element(
    state: _BuilderState,
    node: Any,
    section_node_id: Optional[str],
    policy: LongformPolicy,
    preflight: ResourcePreflight,
) -> None:
    node_id = getattr(node, "node_id", None)

    if isinstance(node, Section):
        _render_section(state, node, policy, preflight)
        return

    if isinstance(node, Paragraph):
        state.add(
            "writer.add_paragraph",
            {"text": _span_text(node.spans), "style": "Body Text"},
            node_id=node_id,
        )
    elif isinstance(node, ListBlock):
        if node.ordered:
            state.add(
                "writer.add_list",
                {"items": _list_items(node.items), "ordered": True},
                node_id=node_id,
            )
        else:
            state.add(
                "writer.add_list",
                {"items": _list_items(node.items), "glyph": "•"},
                node_id=node_id,
            )
    elif isinstance(node, FigureBlock):
        _render_figure(state, node, node_id, preflight)
    elif isinstance(node, SemanticTableBlock):
        _render_semantic_table(state, node, node_id)
    elif isinstance(node, (FormulaBlock, MathBlock)):
        _render_equation(state, node, node_id)
    elif isinstance(node, ReferenceListBlock):
        _render_bibliography(state, node, node_id)
    elif isinstance(node, DegradationBlock):
        _render_degradation(state, node, node_id)
    elif isinstance(node, PageBreakBlock):
        state.add("writer.add_page_break", {}, node_id=node_id)


def _render_figure(
    state: _BuilderState,
    node: FigureBlock,
    node_id: Optional[str],
    preflight: ResourcePreflight,
) -> None:
    state.figure_count += 1
    if not node_id:
        node_id = f"wpsc-fig:{state.figure_count}"
    state.has_figures = True

    resource_by_path: dict[str, dict[str, Any]] = {}
    for resource in preflight.resources:
        resource_by_path[resource.source_path] = {"resourceId": resource.resource_id}

    degradation_by_path: dict[str, DocumentIssue] = {}
    for deg in preflight.degradations:
        if deg.source_path:
            degradation_by_path[deg.source_path] = deg

    children: list[dict[str, Any]] = []
    for index, image in enumerate(node.images, start=1):
        child_id = f"{node_id}/image:{index}"
        normalized_path = image.path.replace("\\", "/")
        if normalized_path in resource_by_path:
            children.append({"nodeId": child_id, "resourceId": resource_by_path[normalized_path]["resourceId"]})
        elif normalized_path in degradation_by_path:
            deg = degradation_by_path[normalized_path]
            children.append({
                "nodeId": child_id,
                "plannedDegradation": {
                    "code": deg.code,
                    "message": deg.message.replace(image.path, "<redacted>").replace(normalized_path, "<redacted>"),
                    "fallback": deg.fallback_text.replace(image.path, "<redacted>").replace(normalized_path, "<redacted>"),
                    "placement": "block",
                },
            })
        else:
            children.append({
                "nodeId": child_id,
                "plannedDegradation": {
                    "code": "RESOURCE_NOT_FOUND",
                    "message": "Resource not found in manifest",
                    "fallback": "[RESOURCE_NOT_FOUND]",
                    "placement": "block",
                },
            })

    state.add(
        "writer.add_captioned_figure",
        {
            "caption": node.caption,
            "children": children,
            "layout": node.layout or "stack",
        },
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["IMAGE_INSERT_FAILED"],
            "fallback": "notice",
        },
    )


def _render_semantic_table(
    state: _BuilderState,
    node: SemanticTableBlock,
    node_id: Optional[str],
) -> None:
    state.table_count += 1
    if not node_id:
        node_id = f"wpsc-tab:{state.table_count}"
    state.has_tables = True
    state.add(
        "writer.add_semantic_table",
        {
            "caption": node.caption,
            "headers": node.headers,
            "rows": node.rows,
            "alignments": node.alignments,
        },
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["TABLE_INSERT_FAILED"],
            "fallback": "notice",
        },
    )


def _render_equation(
    state: _BuilderState,
    node: Any,
    node_id: Optional[str],
) -> None:
    state.equation_count += 1
    if not node_id:
        node_id = f"wpsc-eq:{state.equation_count}"
    source = getattr(node, "source", None) or getattr(node, "latex", "")
    number = getattr(node, "number", None)
    state.add(
        "writer.add_equation",
        {
            "source": source,
            "number": number,
            "fallbackText": source,
        },
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "inline",
        },
    )


def _render_bibliography(
    state: _BuilderState,
    node: ReferenceListBlock,
    node_id: Optional[str],
) -> None:
    entries = [line.strip() for line in node.entries if line.strip()]
    if entries:
        state.add(
            "writer.add_bibliography",
            {"entries": entries, "style": "numbered"},
            node_id=node_id,
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                "fallback": "notice",
            },
        )


def _render_degradation(
    state: _BuilderState,
    node: DegradationBlock,
    node_id: Optional[str],
) -> None:
    placement = getattr(node.issue, "placement", "block") or "block"
    if placement == "inline":
        state.add(
            "writer.add_inline_degradation",
            {
                "code": node.issue.code,
                "message": node.issue.message,
                "fallbackText": node.fallback_text,
            },
            node_id=node_id,
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["DEGRADATION_INSERT_FAILED"],
                "fallback": "inline",
            },
        )
    else:
        state.add(
            "writer.add_degradation_notice",
            {
                "code": node.issue.code,
                "message": node.issue.message,
                "fallbackText": node.fallback_text,
                "placement": placement,
            },
            node_id=node_id,
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["DEGRADATION_INSERT_FAILED"],
                "fallback": "notice",
            },
        )


def _build_indexes(state: _BuilderState, policy: LongformPolicy) -> None:
    if (
        policy.figure_index
        and state.has_figures
        and not state.front_matter_figure_index_emitted
    ):
        state.add(
            "writer.insert_figure_index",
            {"title": policy.figure_index_title},
            node_id="doc:figure-index",
        )
    if (
        policy.table_index
        and state.has_tables
        and not state.front_matter_table_index_emitted
    ):
        state.add(
            "writer.insert_table_index",
            {"title": policy.table_index_title},
            node_id="doc:table-index",
        )


def _build_quality_notices(
    state: _BuilderState,
    issues: tuple[DocumentIssue, ...],
) -> None:
    notices = [
        {
            "code": issue.code,
            "message": issue.message,
            "fallbackText": issue.code,
            "placement": getattr(issue, "placement", "document") or "document",
        }
        for issue in issues
        if getattr(issue, "placement", "document") == "document"
    ]
    if notices:
        state.add(
            "writer.add_document_quality_notice",
            {"notices": notices},
            node_id="doc:quality",
        )


def _build_finalize(state: _BuilderState) -> None:
    state.add("writer.finalize_fields", {"maxRounds": 3}, node_id="doc:finalize")


__all__ = ["build_longform_plan"]
