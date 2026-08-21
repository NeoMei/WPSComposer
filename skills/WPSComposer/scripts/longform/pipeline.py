"""Offline long-form generation pipeline.

Composes the M1 platform-independent stages from Markdown through a closed,
deterministic generation plan.  This module never imports WPS/executor modules
and never writes files.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional, Tuple

from ..document_model import DocumentIssue, StructuredDocument
from ..generation_plan import GenerationPlan
from ..md_parser import parse_markdown
from .plan import build_longform_plan
from .resources import ResourcePreflight, preflight_resources
from .semantic import SemanticResult, normalize_longform_document


@dataclass(frozen=True)
class LongformBuild:
    """Result of building a long-form generation plan offline."""

    document: StructuredDocument
    semantic: SemanticResult
    preflight: ResourcePreflight
    plan: GenerationPlan
    issues: Tuple[DocumentIssue, ...]

    def to_json(self) -> dict[str, Any]:
        semantic_json = self.semantic.to_json()
        return {
            "document": semantic_json.get("document", {}),
            "semantic": semantic_json,
            "preflight": {
                "resources": [
                    {
                        "resourceId": r.resource_id,
                        "sourcePath": r.source_path,
                        "sourceSha256": r.source_sha256,
                        "payloadSha256": r.payload_sha256,
                        "byteLength": r.byte_length,
                        "mediaType": r.media_type,
                        "normalizerId": r.normalizer_id,
                    }
                    for r in self.preflight.resources
                ],
                "degradations": [
                    {
                        "nodeId": d.node_id,
                        "code": d.code,
                        "message": d.message,
                        "fallbackText": d.fallback_text,
                        "sourcePath": d.source_path,
                    }
                    for d in self.preflight.degradations
                ],
                "manifest": self.preflight.manifest,
            },
            "plan": self.plan.to_dict(),
            "issues": [
                {"code": i.code, "message": i.message, "placement": i.placement}
                for i in self.issues
            ],
        }


def _aggregate_issues(
    semantic: SemanticResult, preflight: ResourcePreflight
) -> Tuple[DocumentIssue, ...]:
    """Combine semantic and resource-preflight issues deterministically."""
    issues: list[DocumentIssue] = list(semantic.issues)
    for degradation in preflight.degradations:
        issues.append(
            DocumentIssue(
                code=degradation.code,
                message=degradation.message,
                placement="block",
            )
        )
    # Deterministic ordering: code, message, placement, then optional node id.
    issues.sort(key=lambda i: (i.code, i.message, i.placement))
    return tuple(issues)


def build_longform_generation(
    markdown: str,
    base_dir: str = "",
) -> LongformBuild:
    """Build an offline long-form generation plan from Markdown.

    Args:
        markdown: Raw Markdown input.
        base_dir: Directory for resolving relative resource paths.

    Returns:
        A LongformBuild containing the parsed document, normalized semantic
        result, resource preflight, protocol v2 generation plan, and merged
        deterministic issue list.
    """
    document = parse_markdown(markdown, base_dir=base_dir, longform=True)
    semantic = normalize_longform_document(document)
    preflight = preflight_resources(semantic.document.sections, base_dir)
    plan = build_longform_plan(semantic, preflight)
    issues = _aggregate_issues(semantic, preflight)
    return LongformBuild(
        document=document,
        semantic=semantic,
        preflight=preflight,
        plan=plan,
        issues=issues,
    )


__all__ = ["LongformBuild", "build_longform_generation"]
