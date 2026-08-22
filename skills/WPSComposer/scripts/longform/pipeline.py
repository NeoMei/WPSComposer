"""Offline long-form generation pipeline.

Composes the M1 platform-independent stages from Markdown through a closed,
deterministic generation plan.  This module never imports WPS/executor modules
and never writes files.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Tuple

from ..document_model import DocumentIssue, StructuredDocument
from ..generation_plan import (
    GenerationPlan,
    GenerationResource,
    validate_generation_plan,
)
from ..md_parser import parse_markdown
from .executor import ExecutionOutcome, LongformExecutor
from .plan import build_longform_plan
from .resources import (
    PreflightResource,
    ResourcePreflight,
    preflight_resources,
)
from .semantic import SemanticResult, normalize_longform_document


_REDACTED_PATH = "[REDACTED]"


def _is_absolute_path(value: str) -> bool:
    """Return True when value looks like an absolute filesystem path."""
    if not value:
        return False
    norm = value.replace("\\", "/")
    if norm.startswith("/"):
        return True
    if len(norm) >= 2 and norm[1] == ":" and norm[0].isalpha():
        return True
    return False


def _build_path_redaction_map(
    base_dir: str,
    preflight: ResourcePreflight,
) -> dict[str, str]:
    """Build a map of sensitive path strings to redaction placeholders.

    The map covers the absolute base_dir, every preflight source_path
    (relative or absolute), and both slash and backslash variants.
    Resource paths are replaced with their deterministic resource id;
    rejected/degraded paths are replaced with a generic placeholder.
    """
    redactions: dict[str, str] = {}
    if base_dir:
        try:
            abs_base = str(Path(base_dir).resolve())
        except (OSError, RuntimeError):
            abs_base = base_dir
        redactions[abs_base] = _REDACTED_PATH
        redactions[base_dir] = _REDACTED_PATH

    def _add_token(token: str, placeholder: str) -> None:
        if not token:
            return
        redactions[token] = placeholder
        redactions[token.replace("\\", "/")] = placeholder
        redactions[token.replace("/", "\\")] = placeholder

    for resource in preflight.resources:
        _add_token(resource.source_path, resource.resource_id)
    for degradation in preflight.degradations:
        _add_token(degradation.source_path, _REDACTED_PATH)
    return redactions


def _redact_path_values(value: Any) -> Any:
    """Recursively replace any dict value keyed "path" with a placeholder."""
    if isinstance(value, dict):
        return {
            k: (_REDACTED_PATH if k == "path" and isinstance(v, str) else _redact_path_values(v))
            for k, v in value.items()
        }
    if isinstance(value, (list, tuple)):
        return [_redact_path_values(v) for v in value]
    return value


def _apply_redactions(value: Any, redactions: dict[str, str]) -> Any:
    """Recursively apply redaction token substitution to string values."""
    if isinstance(value, dict):
        return {k: _apply_redactions(v, redactions) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_apply_redactions(v, redactions) for v in value]
    if isinstance(value, str):
        for token in sorted(redactions, key=len, reverse=True):
            value = value.replace(token, redactions[token])
        return value
    return value


def _build_executor_resources(
    base_dir: str,
    preflight: ResourcePreflight,
) -> tuple[GenerationResource, ...]:
    """Convert preflight resources into executor-ready GenerationResources.

    Relative source_paths are resolved against base_dir so that real
    executors receive absolute, staging-ready paths.
    """
    if base_dir:
        base = Path(base_dir).resolve()
    else:
        base = Path.cwd()

    resources: list[GenerationResource] = []
    for resource in preflight.resources:
        source = Path(resource.source_path)
        if source.is_absolute():
            resolved = source.resolve()
        else:
            resolved = (base / source).resolve()
        resources.append(
            GenerationResource(
                id=resource.resource_id,
                source_path=str(resolved),
                media_type=resource.media_type,
            )
        )
    return tuple(resources)


def execute_longform_plan(
    build: "LongformBuild",
    executor: LongformExecutor,
    deadline: Optional[float] = None,
) -> ExecutionOutcome:
    """Bind a validated long-form plan and resource manifest to an executor.

    This is an optional step after `build_longform_generation`.  It does not
    change the public `generate()` path and remains platform-pure: the caller
    supplies the executor instance; the pipeline never imports platform modules.

    The executor resources are resolved to absolute paths via the original
    base_dir stored on the build.  Callers that pass an absolute base_dir or
    absolute image paths receive fully resolved staging paths; callers using
    relative paths receive paths resolved against the recorded base_dir.
    """
    validate_generation_plan(build.plan.to_dict(), component="writer")
    resources = _build_executor_resources(build.base_dir, build.preflight)
    return executor.execute(build.plan, resources, deadline=deadline)


@dataclass(frozen=True)
class LongformBuild:
    """Result of building a long-form generation plan offline."""

    document: StructuredDocument
    semantic: SemanticResult
    preflight: ResourcePreflight
    plan: GenerationPlan
    issues: Tuple[DocumentIssue, ...]
    base_dir: str = ""

    def resource_source_map(self) -> dict[str, str]:
        """Private staging transport map: resource id -> source path.

        This map is intentionally absent from `to_json()` so that diagnostic
        output never leaks absolute or relative source paths.
        """
        return {
            r.resource_id: r.source_path
            for r in self.preflight.resources
        }

    def to_json(self) -> dict[str, Any]:
        redactions = _build_path_redaction_map(self.base_dir, self.preflight)
        semantic_json = self.semantic.to_json()
        document_json = _redact_path_values(
            _apply_redactions(semantic_json.get("document", {}), redactions)
        )
        semantic_redacted = _redact_path_values(
            _apply_redactions(semantic_json, redactions)
        )
        return {
            "document": document_json,
            "semantic": semantic_redacted,
            "preflight": {
                "resources": [
                    {
                        "resourceId": r.resource_id,
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
                        "message": _apply_redactions(d.message, redactions),
                        "fallbackText": _apply_redactions(d.fallback_text, redactions),
                    }
                    for d in self.preflight.degradations
                ],
                "manifest": self.preflight.manifest,
            },
            "plan": self.plan.to_dict(),
            "issues": [
                {
                    "code": i.code,
                    "message": _apply_redactions(i.message, redactions),
                    "placement": i.placement,
                }
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
        base_dir=base_dir,
    )


__all__ = [
    "LongformBuild",
    "build_longform_generation",
    "execute_longform_plan",
]
