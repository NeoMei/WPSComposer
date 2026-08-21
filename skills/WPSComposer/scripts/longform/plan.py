"""Build a closed, deterministic protocol v2 generation plan for a long-form document.

The plan builder converts a normalized semantic model and preflighted resource
manifest into a pure JSON operation plan.  It never reads files or launches WPS.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, Optional

from ..document_model import (
    DegradationBlock,
    DocumentIssue,
    FigureBlock,
    FormulaBlock,
    ImageBlock,
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


@dataclass
class _BuilderState:
    """Mutable builder state that is not part of the public plan."""

    operations: list[GenerationOperation] = field(default_factory=list)
    figure_count: int = 0
    table_count: int = 0
    equation_count: int = 0
    has_figures: bool = False
    has_tables: bool = False

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

    _build_begin(state, policy)
    _build_sections(state, semantic.document, policy, preflight)
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


def _build_begin(state: _BuilderState, policy: LongformPolicy) -> None:
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
        },
    )
    if policy.toc:
        state.add(
            "writer.insert_toc",
            {"title": policy.toc_title, "levels": policy.toc_levels},
        )


def _build_sections(
    state: _BuilderState,
    document: StructuredDocument,
    policy: LongformPolicy,
    preflight: ResourcePreflight,
) -> None:
    resource_by_path: dict[str, dict[str, Any]] = {}
    for resource in preflight.resources:
        resource_by_path[resource.source_path] = {
            "resourceId": resource.resource_id,
        }

    degradation_by_path: dict[str, DocumentIssue] = {}
    for deg in preflight.degradations:
        # Key by the normalized source path so lookup is independent of OS separators.
        if deg.source_path:
            degradation_by_path[deg.source_path] = deg

    def process_node(node: Any, section_node_id: Optional[str] = None) -> None:
        if isinstance(node, Section):
            node_id = node.node_id or section_node_id
            state.add(
                "writer.add_heading",
                {"text": node.heading, "level": node.level},
                node_id=node_id,
            )
            for element in node.elements:
                process_node(element, section_node_id=node_id)
            return

        node_id = getattr(node, "node_id", None)

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
            state.figure_count += 1
            if not node_id:
                node_id = f"wpsc-fig:{state.figure_count}"
            state.has_figures = True
            children: list[dict[str, Any]] = []
            for index, image in enumerate(node.images, start=1):
                child_id = f"{node_id}/image:{index}"
                normalized_path = image.path.replace("\\", "/")
                if normalized_path in resource_by_path:
                    children.append(
                        {
                            "nodeId": child_id,
                            "resourceId": resource_by_path[normalized_path]["resourceId"],
                        }
                    )
                elif normalized_path in degradation_by_path:
                    deg = degradation_by_path[normalized_path]
                    children.append(
                        {
                            "nodeId": child_id,
                            "plannedDegradation": {
                                "code": deg.code,
                                # Redact both the original and normalized path forms.
                                "message": deg.message.replace(image.path, "<redacted>").replace(normalized_path, "<redacted>"),
                                "fallback": deg.fallback_text.replace(image.path, "<redacted>").replace(normalized_path, "<redacted>"),
                                "placement": "block",
                            },
                        }
                    )
                else:
                    children.append(
                        {
                            "nodeId": child_id,
                            "plannedDegradation": {
                                "code": "RESOURCE_NOT_FOUND",
                                "message": "Resource not found in manifest",
                                "fallback": "[RESOURCE_NOT_FOUND]",
                                "placement": "block",
                            },
                        }
                    )
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
        elif isinstance(node, SemanticTableBlock):
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
        elif isinstance(node, (FormulaBlock, MathBlock)):
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
        elif isinstance(node, ReferenceListBlock):
            entries = [
                line.strip()
                for line in node.entries
                if line.strip()
            ]
            if entries:
                state.add(
                    "writer.add_bibliography",
                    {
                        "entries": entries,
                        "style": "numbered",
                    },
                    node_id=node_id,
                    failure_policy={
                        "mode": "degrade",
                        "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                        "fallback": "notice",
                    },
                )
        elif isinstance(node, DegradationBlock):
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
        elif isinstance(node, PageBreakBlock):
            state.add("writer.add_page_break", {}, node_id=node_id)

    for section in document.sections:
        process_node(section)


def _build_indexes(state: _BuilderState, policy: LongformPolicy) -> None:
    if policy.figure_index and state.has_figures:
        state.add(
            "writer.insert_figure_index",
            {"title": policy.figure_index_title},
        )
    if policy.table_index and state.has_tables:
        state.add(
            "writer.insert_table_index",
            {"title": policy.table_index_title},
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
        )


def _build_finalize(state: _BuilderState) -> None:
    state.add(
        "writer.finalize_fields",
        {"maxRounds": 3},
    )


__all__ = ["build_longform_plan"]
