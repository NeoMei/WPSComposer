"""Build a closed, deterministic protocol v2 generation plan for a long-form document.

The plan builder converts a normalized semantic model and preflighted resource
manifest into a pure JSON operation plan.  It never reads files or launches WPS.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
from types import SimpleNamespace
from typing import Any, Optional

from ..document_model import (
    BibliographyEntry,
    BlockQuote,
    CaptionBinding,
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
    TableBlock,
)
from ..generation_plan import GenerationOperation, GenerationPlan
from .page_policy import _group_sections, build_page_policy
from .image_policy import resolve_figure_layout
from .native_fields import caption_numbering_descriptor, cross_reference_descriptor
from .policy import LongformPolicy, build_policy, page_content_width_pt
from .resources import ResourcePreflight
from .semantic import SemanticResult
from .table_policy import resolve_table_policy
from .native_math import contains_forbidden_formula_command
from .privacy import contains_private_plan_text


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


def _redact_private_plan_text(text: str, fallback: str) -> str:
    if contains_private_plan_text(text):
        return fallback
    return text


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
    active_numbered_h1: bool = False
    table_preset: str = "academic"

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
    requested_preset = str(
        semantic.document.metadata.get("design", "academic")
    ).strip().lower()
    state = _BuilderState(
        table_preset=requested_preset
        if requested_preset in {"academic", "business", "consultant", "tech", "proposal"}
        else "academic"
    )

    _build_begin(state, semantic.document, policy)
    _build_page_skeleton(
        state,
        semantic.document,
        semantic.config,
        policy,
        preflight,
        semantic.issues,
    )
    _build_indexes(state, policy)
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
        binding = node.caption_binding
        return bool(node.caption) if binding is None else binding.indexable
    if isinstance(node, Section):
        return any(_contains_figure(child) for child in node.elements)
    return False


def _contains_table(node: Any) -> bool:
    if isinstance(node, SemanticTableBlock):
        binding = node.caption_binding
        return bool(node.caption) if binding is None else binding.indexable
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
    issues: tuple[DocumentIssue, ...],
) -> None:
    """Emit configure_section operations and render each role's content."""
    skeleton = build_page_policy(document, config, policy)
    groups = _group_sections(document.sections)
    group_iter = iter(groups)

    state.has_figures, state.has_tables = _scan_document_indexes(document)

    anchor_emitted = False
    for index, section_policy in enumerate(skeleton.sections):
        if section_policy.role == "cover":
            _emit_configure_section(state, index, section_policy, landscape=False)
            continue

        if not anchor_emitted:
            _build_quality_anchor(state, issues)
            anchor_emitted = True

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

    if not anchor_emitted:
        _build_quality_anchor(state, issues)


def _render_front_matter(
    state: _BuilderState,
    document: StructuredDocument,
    section_policy: Any,
    policy: LongformPolicy,
) -> None:
    if section_policy.includes_abstract and document.abstract:
        for index, paragraph in enumerate(document.abstract.paragraphs, start=1):
            text = paragraph.plain_text
            if text:
                paragraph_node_id = paragraph.node_id or f"doc:abstract:{index}"
                if any(_has_structured_inline_run(span) for span in paragraph.spans):
                    _render_cross_reference_paragraph(
                        state, paragraph, paragraph_node_id
                    )
                else:
                    state.add(
                        "writer.add_paragraph",
                        {"text": text, "style": "Body Text"},
                        node_id=paragraph_node_id,
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
            {
                "title": policy.figure_index_title,
                "sequenceId": "WPSC_FIG",
                "titleStyleId": "WPSC_INDEX_TITLE",
            },
            node_id="doc:figure-index",
        )
        state.front_matter_figure_index_emitted = True
    if section_policy.includes_table_index and state.has_tables:
        state.add(
            "writer.insert_table_index",
            {
                "title": policy.table_index_title,
                "sequenceId": "WPSC_TAB",
                "titleStyleId": "WPSC_INDEX_TITLE",
            },
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
        heading_digest = hashlib.sha256(
            f"heading\0{section.node_id}".encode("utf-8")
        ).hexdigest()[:24]
        heading_args: dict[str, Any] = {
            "text": section.heading,
            "level": section.level,
            "bookmarkName": f"wpsc_head_{heading_digest}",
        }
        if (
            section.level in {1, 2, 3, 4}
            and section.numbering not in ("none", "auto")
            and section.numbering_scheme
        ):
            heading_args["numbering"] = True
            heading_args["numberingScheme"] = section.numbering_scheme
            if section.level == 1:
                state.active_numbered_h1 = True
        elif section.level == 1 and state.active_numbered_h1:
            heading_args["numbering"] = False
            heading_args["sequenceTransparent"] = True
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
        if any(_has_structured_inline_run(span) for span in node.spans):
            _render_cross_reference_paragraph(state, node, node_id)
        else:
            state.add(
                "writer.add_paragraph",
                {"text": _span_text(node.spans), "style": "Body Text"},
                node_id=node_id,
            )
    elif isinstance(node, ListBlock):
        if any(
            _has_structured_inline_run(span)
            for item in node.items
            for span in item
        ):
            for index, spans in enumerate(node.items, start=1):
                item_node_id = (
                    node.item_node_ids[index - 1]
                    if index <= len(node.item_node_ids)
                    else f"{node_id or 'wpsc-list'}/item:{index}"
                )
                marker = f"{index}.\t" if node.ordered else "•\t"
                _render_cross_reference_paragraph(
                    state,
                    Paragraph(spans=list(spans), node_id=item_node_id),
                    item_node_id,
                    prefix_text=marker,
                    list_formatting={
                        "kind": "ordered" if node.ordered else "bullet",
                        "indentPt": 24.0,
                    },
                )
        elif node.ordered:
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
    elif isinstance(node, TableBlock):
        _render_semantic_table(
            state,
            SemanticTableBlock(
                node_id=node.node_id,
                headers=node.headers,
                rows=node.rows,
                alignments=node.alignments,
                style="grid",
                cell_degradations=node.cell_degradations,
                cell_citations=node.cell_citations,
            ),
            node_id,
        )
    elif isinstance(node, (FormulaBlock, MathBlock)):
        _render_equation(state, node, node_id, preflight)
    elif isinstance(node, ReferenceListBlock):
        _render_bibliography(state, node, node_id)
    elif isinstance(node, DegradationBlock):
        _render_degradation(state, node, node_id)
    elif isinstance(node, BlockQuote):
        for index, paragraph in enumerate(node.paragraphs, start=1):
            paragraph_node_id = paragraph.node_id or f"{node_id or 'blockquote'}/paragraph:{index}"
            if any(_has_structured_inline_run(span) for span in paragraph.spans):
                _render_cross_reference_paragraph(
                    state, paragraph, paragraph_node_id
                )
            else:
                state.add(
                    "writer.add_paragraph",
                    {"text": paragraph.plain_text, "style": "Body Text"},
                    node_id=paragraph_node_id,
                )
    elif isinstance(node, PageBreakBlock):
        state.add("writer.add_page_break", {}, node_id=node_id)
        for index, paragraph in enumerate(node.content, start=1):
            paragraph_node_id = paragraph.node_id or f"{node_id or 'page-break'}/paragraph:{index}"
            if any(_has_structured_inline_run(span) for span in paragraph.spans):
                _render_cross_reference_paragraph(
                    state, paragraph, paragraph_node_id
                )
            else:
                state.add(
                    "writer.add_paragraph",
                    {"text": paragraph.plain_text, "style": "Body Text"},
                    node_id=paragraph_node_id,
                )


def _render_figure(
    state: _BuilderState,
    node: FigureBlock,
    node_id: Optional[str],
    preflight: ResourcePreflight,
) -> None:
    state.figure_count += 1
    if not node_id:
        node_id = f"wpsc-fig:{state.figure_count}"

    resource_by_path: dict[str, dict[str, Any]] = {}
    for resource in preflight.resources:
        resource_by_path[resource.source_path] = {"resourceId": resource.resource_id}

    degradation_by_path: dict[str, DocumentIssue] = {}
    for deg in preflight.degradations:
        if deg.source_path:
            degradation_by_path[deg.source_path] = deg

    resolved_layouts = _resolve_available_figure_layouts(node, preflight)
    children: list[dict[str, Any]] = []
    for index, image in enumerate(node.images, start=1):
        child_id = f"{node_id}/image:{index}"
        normalized_path = image.path.replace("\\", "/")
        if normalized_path in resource_by_path:
            resource_id = resource_by_path[normalized_path]["resourceId"]
            layout = resolved_layouts[resource_id]
            children.append({
                "nodeId": child_id,
                "resourceId": resource_id,
                "displayWidthPt": round(layout.display_width_pt, 4),
                "displayHeightPt": round(layout.display_height_pt, 4),
                "effectiveDpi": round(layout.effective_dpi, 4),
                "mediaType": layout.media_type,
                "normalizerId": layout.normalizer_id,
            })
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

    binding = _binding_or_legacy(node.caption_binding, "figure", node_id, bool(node.caption))
    state.has_figures = state.has_figures or binding.indexable
    width_mode, explicit_width = _width_descriptor(node.width)
    args: dict[str, Any] = {
        "caption": node.caption,
        "numbering": caption_numbering_descriptor("figure", binding),
        "indexable": binding.indexable,
        "referenceable": binding.referenceable,
        "widthMode": width_mode,
        "orientation": node.orientation,
        "kind": node.kind,
        "children": children,
        "layout": "columns" if node.layout in {"columns", "side-by-side"} else "stack",
        "keepWithCaption": True,
    }
    if binding.bookmark_name is not None:
        args["bookmarkName"] = binding.bookmark_name
    if explicit_width is not None:
        args["explicitWidthPt"] = explicit_width
    if args["layout"] == "columns":
        args["columns"] = 2
    state.add(
        "writer.add_captioned_figure",
        args,
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["IMAGE_INSERT_FAILED"],
            "fallback": "figure-child-stack-then-notice",
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
    binding = _binding_or_legacy(node.caption_binding, "table", node_id, bool(node.caption))
    state.has_tables = state.has_tables or binding.indexable
    table_policy, table_issues = resolve_table_policy(node, state.table_preset)
    border_spec = {
        "top": table_policy.borders["top"],
        "bottom": table_policy.borders["bottom"],
        "headerBottom": table_policy.borders["header_bottom"],
        "left": table_policy.borders["left"],
        "right": table_policy.borders["right"],
        "insideHorizontal": table_policy.borders["inside_horizontal"],
        "insideVertical": table_policy.borders["inside_vertical"],
    }
    args: dict[str, Any] = {
        "caption": node.caption,
        "numbering": caption_numbering_descriptor("table", binding),
        "indexable": binding.indexable,
        "referenceable": binding.referenceable,
        "headers": node.headers,
        "rows": node.rows,
        "alignments": node.alignments,
        "style": table_policy.style,
        "orientation": node.orientation,
        "borderSpec": border_spec,
        "merges": [
            {"top": item.top, "left": item.left, "bottom": item.bottom, "right": item.right}
            for item in table_policy.merges
        ],
        "repeatHeader": table_policy.repeat_header,
        "allowRowSplit": table_policy.allow_row_split,
        "cellIndentPt": table_policy.cell_indent_pt,
        "plannedDegradation": [issue.to_dict() for issue in table_issues],
        "keepCaptionWithFirstRow": True,
    }
    if node.cell_degradations:
        args["cellDegradations"] = [
            {
                "row": item.row,
                "column": item.column,
                "code": item.code,
                "fallbackText": item.fallback_text,
            }
            for item in node.cell_degradations
        ]
    if node.cell_citations:
        args["cellCitations"] = [
            {
                "row": item.row,
                "column": item.column,
                "nodeId": item.node_id,
                "targetId": item.target_id,
                "targetNodeId": item.target_node_id,
                "number": item.number,
                "fallbackText": item.fallback_text,
            }
            for item in node.cell_citations
        ]
    if binding.bookmark_name is not None:
        args["bookmarkName"] = binding.bookmark_name
    state.add(
        "writer.add_semantic_table",
        args,
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": [
                "TABLE_STYLE_APPLY_FAILED",
                "TABLE_MERGE_APPLY_FAILED",
                "TABLE_ROW_FORCED_SPLIT",
                "TABLE_INSERT_FAILED",
            ],
            "fallback": "grid-then-text",
        },
    )


def _render_equation(
    state: _BuilderState,
    node: Any,
    node_id: Optional[str],
    preflight: ResourcePreflight,
) -> None:
    state.equation_count += 1
    if not node_id:
        node_id = f"wpsc-eq:{state.equation_count}"
    source = getattr(node, "source", None) or getattr(node, "latex", "")
    binding = _binding_or_legacy(
        getattr(node, "caption_binding", None), "equation", node_id, True
    )
    native_math = getattr(node, "native_math", None)
    content_issue = getattr(node, "content_degradation", None)
    if native_math is None and content_issue is None:
        state.add(
            "writer.add_equation",
            {
                "source": source,
                "numbering": caption_numbering_descriptor("equation", binding),
                "bookmarkName": binding.bookmark_name,
                "fallbackText": source,
            },
            node_id=node_id,
            failure_policy={"mode": "fail"},
        )
        return

    fallback_text = (
        "[FORMULA_SOURCE_REDACTED 公式源已脱敏]"
        if source and contains_forbidden_formula_command(source)
        else _redact_private_plan_text(
            source,
            "[FORMULA_SOURCE_REDACTED 公式源已脱敏]",
        )
        if source
        else "[FORMULA_MALFORMED 公式源不可用]"
    )
    if native_math is not None:
        content = {
            "nativeMath": {
                "syntax": native_math.syntax,
                "linearText": native_math.linear_text,
                "sourceHash": native_math.source_hash,
            }
        }
    else:
        code = content_issue.code
        content = {
            "plannedDegradation": _formula_degradation_descriptor(
                code=code,
                reason="Formula content was rejected during native-math preflight.",
                fallback_text=fallback_text,
                fallback_kind="source",
            )
        }

    args: dict[str, Any] = {
        "renderMode": "native-m4",
        "content": content,
        "numbering": caption_numbering_descriptor("equation", binding),
        "bookmarkName": binding.bookmark_name,
        "fallbackText": fallback_text,
    }
    resource_id = preflight.formula_bindings.get(node_id)
    if resource_id is not None:
        args["fallbackResource"] = {"fallbackResourceId": resource_id}
    else:
        resource_degradation = next(
            (
                item
                for item in preflight.degradations
                if item.node_id == node_id
                and item.code == "FORMULA_FALLBACK_IMAGE_UNAVAILABLE"
            ),
            None,
        )
        if resource_degradation is not None:
            args["fallbackResource"] = {
                "fallbackResourcePlannedDegradation": _formula_degradation_descriptor(
                    code=resource_degradation.code,
                    reason="Formula fallback image is unavailable.",
                    fallback_text=(
                        "[FORMULA_FALLBACK_IMAGE_UNAVAILABLE 公式图像备选不可用]"
                    ),
                    fallback_kind="none",
                )
            }
    state.add(
        "writer.add_equation",
        args,
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
    )


def _formula_degradation_descriptor(
    *,
    code: str,
    reason: str,
    fallback_text: str,
    fallback_kind: str,
) -> dict[str, Any]:
    return {
        "code": code,
        "placement": "block",
        "objectLabel": "formula",
        "reason": reason,
        "fallbackText": fallback_text,
        "fallbackKind": fallback_kind,
    }


def _render_cross_reference_paragraph(
    state: _BuilderState,
    node: Paragraph,
    node_id: Optional[str],
    *,
    prefix_text: str = "",
    list_formatting: Optional[dict[str, Any]] = None,
) -> None:
    runs: list[dict[str, Any]] = []
    if prefix_text:
        runs.append({"type": "text", "text": prefix_text})
    for span in node.spans:
        if span.citation is not None:
            citation = span.citation
            runs.append({
                "type": "citation",
                "nodeId": citation.node_id,
                "targetId": citation.target_id,
                "targetNodeId": citation.target_node_id,
                "number": citation.number,
                "fallbackText": citation.fallback_text,
            })
            continue
        if span.inline_degradation is not None:
            degradation = span.inline_degradation
            runs.append({
                "type": "degradation",
                "nodeId": degradation.node_id,
                "code": degradation.code,
                "fallbackText": degradation.fallback_text,
            })
            continue
        if span.cross_reference is None:
            if span.text:
                runs.append({"type": "text", "text": span.text})
            continue
        if not _is_resolved_reference(span):
            if span.text:
                runs.append({"type": "text", "text": span.text})
            continue
        descriptor = cross_reference_descriptor(span.cross_reference)
        runs.append({
            "type": "reference",
            **descriptor,
            "fallbackText": span.cross_reference.fallback_text,
        })
    args: dict[str, Any] = {"runs": runs}
    if list_formatting is not None:
        args["listFormatting"] = list_formatting
    state.add(
        "writer.add_cross_reference",
        args,
        node_id=node_id,
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    )


def _is_resolved_reference(span: Span) -> bool:
    run = span.cross_reference
    return bool(
        run is not None
        and run.target_node_id
        and run.target_kind in {"fig", "tab", "eq", "figure", "table", "equation"}
        and run.bookmark_name
    )


def _has_structured_inline_run(span: Span) -> bool:
    return bool(
        span.citation is not None
        or span.inline_degradation is not None
        or _is_resolved_reference(span)
    )


def _binding_or_legacy(
    binding: Optional[CaptionBinding],
    kind: str,
    node_id: str,
    has_caption: bool,
) -> CaptionBinding:
    if binding is not None:
        return binding
    short_kind = {"figure": "fig", "table": "tab", "equation": "eq"}[kind]
    digest = hashlib.sha256(f"{kind}\0{node_id}".encode("utf-8")).hexdigest()[:24]
    return CaptionBinding(
        mode="global",
        chapter_node_id=None,
        bookmark_name=f"wpsc_{short_kind}_{digest}" if has_caption else None,
        indexable=has_caption,
        referenceable=has_caption,
    )


def _resolve_available_figure_layouts(
    node: FigureBlock,
    preflight: ResourcePreflight,
) -> dict[str, Any]:
    """Resolve accepted children even when siblings have planned degradation."""
    by_path = {
        str(resource.source_path).replace("\\", "/"): resource
        for resource in preflight.resources
    }
    available = [
        (image, by_path[path])
        for image in node.images
        if (path := image.path.replace("\\", "/")) in by_path
    ]
    if not available:
        return {}
    content_width = page_content_width_pt(node.orientation)
    if len(available) == len(node.images):
        layouts = resolve_figure_layout(node, preflight.resources, content_width)
        return {layout.resource_id: layout for layout in layouts}

    slot_width = (
        (content_width - 12.0) / 2.0
        if node.layout in {"columns", "side-by-side"}
        else content_width
    )
    resolved: dict[str, Any] = {}
    for image, resource in available:
        partial = SimpleNamespace(
            images=[image],
            layout="stack",
            width="full" if node.width == "column" else node.width,
        )
        layout = resolve_figure_layout(partial, [resource], slot_width)[0]
        resolved[layout.resource_id] = layout
    return resolved


def _width_descriptor(width: Any) -> tuple[str, Optional[float]]:
    text = str(width or "auto").strip().lower()
    if text in {"auto", "column", "full"}:
        return text, None
    if text.endswith("pt"):
        return "explicit", float(text[:-2])
    raise ValueError(f"unsupported figure width: {width}")


def _render_bibliography(
    state: _BuilderState,
    node: ReferenceListBlock,
    node_id: Optional[str],
) -> None:
    if node.resolved_items:
        segment = 0
        batch: list[BibliographyEntry] = []

        def flush_batch() -> None:
            nonlocal segment
            if not batch:
                return
            segment += 1
            owner = (
                node_id
                if segment == 1
                else f"{node_id or 'bibliography'}/segment:{segment}"
            )
            state.add(
                "writer.add_bibliography",
                {
                    "schemaVersion": 1,
                    "entries": [
                        {
                            "id": item.identifier,
                            "nodeId": item.node_id,
                            "number": item.number,
                            "text": item.text,
                            "cited": item.cited,
                        }
                        for item in batch
                    ],
                    "style": "numeric",
                    "hangingIndentPt": 18.0,
                    "leftIndentPt": 18.0,
                    "spaceAfterPt": 6.0,
                },
                node_id=owner,
                failure_policy={
                    "mode": "degrade",
                    "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                    "fallback": "notice",
                },
            )
            batch.clear()

        for item in node.resolved_items:
            if isinstance(item, BibliographyEntry):
                batch.append(item)
            elif isinstance(item, DegradationBlock):
                flush_batch()
                _render_degradation(state, item, item.node_id)
        flush_batch()
        return

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
        raise ValueError(
            "figure index was not emitted in its front-matter section"
        )
    if (
        policy.table_index
        and state.has_tables
        and not state.front_matter_table_index_emitted
    ):
        raise ValueError(
            "table index was not emitted in its front-matter section"
        )


def _build_quality_anchor(
    state: _BuilderState,
    issues: tuple[DocumentIssue, ...],
) -> None:
    notices: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for issue in issues:
        if getattr(issue, "placement", "document") != "document":
            continue
        message = _redact_private_plan_text(
            issue.message,
            "Issue details were redacted.",
        )
        key = (issue.code, message)
        if key in seen:
            continue
        seen.add(key)
        notices.append({
            "code": issue.code,
            "message": message,
            "fallbackText": issue.code,
            "placement": "document",
        })
    state.add(
        "writer.reserve_document_quality_anchor",
        {"title": "生成质量提示", "notices": notices},
        node_id="doc:quality",
        failure_policy={"mode": "fail"},
    )


def _build_finalize(state: _BuilderState) -> None:
    state.add("writer.finalize_fields", {"maxRounds": 3}, node_id="doc:finalize")


__all__ = ["build_longform_plan"]
