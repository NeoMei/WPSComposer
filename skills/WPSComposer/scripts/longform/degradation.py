"""Closed, privacy-safe recovery contract shared by native executors.

Python is the canonical source for the allowlist below. The macOS add-in
contains a generated frozen representation and implements the same local
checkpoint/rollback state machine because native mutations cannot cross the
loopback bridge.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any, Callable, Mapping, Optional, Tuple

from .privacy import redact_private_text


_PLACEMENTS = frozenset({"inline", "block", "document"})
_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{2,63}$")
_CONTROLLED_TOKEN_RE = re.compile(r"^[a-z][a-z0-9-]{0,63}$")


@dataclass(frozen=True)
class RecoveryRule:
    fallback_kind: str
    placement: str
    fallback_attempts: int = 1


def _rules(*items: Tuple[str, str]) -> Tuple[RecoveryRule, ...]:
    return tuple(
        RecoveryRule(fallback_kind=fallback, placement=placement)
        for fallback, placement in items
    )


RECOVERY_MATRIX: Mapping[str, Tuple[RecoveryRule, ...]] = MappingProxyType({
    "BIBLIOGRAPHY_INSERT_FAILED": _rules(("notice", "block")),
    "CROSS_REFERENCE_FAILED": _rules(("inline-fallback", "inline")),
    "DEGRADATION_INSERT_FAILED": _rules(
        ("inline", "inline"),
        ("notice", "block"),
    ),
    "EQUATION_INSERT_FAILED": _rules(
        ("explicit-image-then-source-notice", "block")
    ),
    "FIELD_REFRESH_UNSTABLE": _rules(
        ("document-quality-notice", "document")
    ),
    "IMAGE_INSERT_FAILED": _rules(
        ("figure-child-stack-then-notice", "block")
    ),
    "TABLE_INSERT_FAILED": _rules(("grid-then-text", "block")),
    "TABLE_MERGE_APPLY_FAILED": _rules(("grid-then-text", "block")),
    "TABLE_ROW_FORCED_SPLIT": _rules(("grid-then-text", "block")),
    "TABLE_STYLE_APPLY_FAILED": _rules(("grid-then-text", "block")),
})


FATAL_BOUNDARY_CODES = frozenset({
    "COM_ERROR",
    "CONFIGURATION_INVALID",
    "DEGRADATION_FALLBACK_FAILED",
    "ENGINE_LOST",
    "EXECUTION_ABORTED",
    "EXECUTION_FAILED",
    "EXECUTOR_CAPABILITY_MISSING",
    "FIELD_REFRESH_CONTRACT_INVALID",
    "FIELD_REFRESH_FAILED",
    "GENERATION_COMMAND_FAILED",
    "INDEX_REFRESH_FAILED",
    "LOCAL_MUTATION_CHECKPOINT_FAILED",
    "LOCAL_MUTATION_ROLLBACK_FAILED",
    "MACOS_DEDICATED_HOST_UNAVAILABLE",
    "PRIVATE_RESOURCE_CLEANUP_FAILED",
    "PRIVATE_RESOURCE_STAGING_FAILED",
    "PROTOCOL_INVALID",
    "RESOURCE_HASH_MISMATCH",
    "SAVE_FAILED",
    "UNKNOWN_OPERATION",
    "WINDOWS_DEDICATED_HOST_UNAVAILABLE",
})


def controlled_token(value: Any) -> Optional[str]:
    """Return a bounded public token, or ``None`` for unsafe values."""

    if isinstance(value, str) and _CONTROLLED_TOKEN_RE.fullmatch(value):
        return value
    return None


@dataclass(frozen=True)
class DegradationDescriptor:
    code: str
    placement: str
    object_label: str
    reason: str
    fallback_text: str
    fallback_kind: str

    def __post_init__(self) -> None:
        if not isinstance(self.code, str) or _CODE_RE.fullmatch(self.code) is None:
            raise ValueError("degradation code must be a stable public code")
        if self.placement not in _PLACEMENTS:
            raise ValueError("degradation placement must be inline, block, or document")
        if controlled_token(self.fallback_kind) is None:
            raise ValueError("degradation fallback kind must be a controlled token")
        for value in (self.object_label, self.reason, self.fallback_text):
            if not isinstance(value, str) or len(value) > 4096:
                raise ValueError("degradation visible text is invalid")

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "placement": self.placement,
            "objectLabel": redact_private_text(self.object_label),
            "reason": redact_private_text(self.reason),
            "fallbackText": redact_private_text(self.fallback_text),
            "fallbackKind": self.fallback_kind,
        }

    def __repr__(self) -> str:
        return f"DegradationDescriptor({self.to_dict()!r})"


@dataclass(frozen=True)
class RecoveryDecision:
    code: str
    recoverable: bool
    placement: str
    fallback_kind: str
    fallback_attempts: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "recoverable": self.recoverable,
            "placement": self.placement,
            "fallbackKind": self.fallback_kind,
            "fallbackAttempts": self.fallback_attempts,
        }


class RecoveryFatalError(RuntimeError):
    """Sanitized failure for anything outside the closed local contract."""

    def __init__(self, code: str, reason: str = "outside recovery contract") -> None:
        self.code = (
            code
            if isinstance(code, str) and _CODE_RE.fullmatch(code)
            else "EXECUTION_FAILED"
        )
        super().__init__(f"{self.code}: {reason}")


def decide_recovery(
    code: str,
    fallback_kind: str,
    placement: str,
) -> RecoveryDecision:
    """Accept one exact allowlisted code/fallback/placement triple or fail."""

    rules = RECOVERY_MATRIX.get(code, ())
    for rule in rules:
        if rule.fallback_kind == fallback_kind and rule.placement == placement:
            return RecoveryDecision(
                code=code,
                recoverable=True,
                placement=placement,
                fallback_kind=fallback_kind,
                fallback_attempts=rule.fallback_attempts,
            )
    raise RecoveryFatalError(code, "outside recovery contract") from None


class LocalRecoveryController:
    """One-node, one-fallback local checkpoint state machine."""

    def __init__(self) -> None:
        self._by_identity: dict[Tuple[str, str], RecoveryDecision] = {}
        self._ordered: list[RecoveryDecision] = []

    @property
    def decisions(self) -> Tuple[RecoveryDecision, ...]:
        return tuple(self._ordered)

    def execute(
        self,
        *,
        node_id: str,
        descriptor: DegradationDescriptor,
        checkpoint: Callable[[], Any],
        native_attempt: Callable[[], Any],
        rollback: Callable[[Any], None],
        fallback_attempt: Callable[[DegradationDescriptor], None],
        insert_notice: Callable[[str, DegradationDescriptor], None],
    ) -> RecoveryDecision:
        identity = (str(node_id), descriptor.code)
        prior = self._by_identity.get(identity)
        if prior is not None:
            return prior

        checkpoint_failed = False
        token: Any = None
        try:
            token = checkpoint()
        except Exception:
            checkpoint_failed = True
        if checkpoint_failed:
            raise RecoveryFatalError(
                "LOCAL_MUTATION_CHECKPOINT_FAILED", "native checkpoint failed"
            ) from None

        native_failed = False
        native_code = ""
        try:
            native_attempt()
        except Exception as error:
            native_failed = True
            candidate = getattr(error, "code", None)
            if isinstance(candidate, str) and _CODE_RE.fullmatch(candidate):
                native_code = candidate
        if not native_failed:
            return RecoveryDecision(
                code=descriptor.code,
                recoverable=False,
                placement=descriptor.placement,
                fallback_kind=descriptor.fallback_kind,
                fallback_attempts=0,
            )
        if not native_code:
            raise RecoveryFatalError(
                "EXECUTION_FAILED", "native operation failed outside recovery contract"
            ) from None
        return self.recover_after_failure(
            node_id=node_id,
            descriptor=descriptor,
            checkpoint_token=token,
            native_code=native_code,
            rollback=rollback,
            fallback_attempt=fallback_attempt,
            insert_notice=insert_notice,
        )

    def recover_after_failure(
        self,
        *,
        node_id: str,
        descriptor: DegradationDescriptor,
        checkpoint_token: Any,
        native_code: str,
        rollback: Callable[[Any], None],
        fallback_attempt: Callable[[DegradationDescriptor], None],
        insert_notice: Callable[[str, DegradationDescriptor], None],
    ) -> RecoveryDecision:
        """Continue the same state machine when the caller owns native dispatch."""
        identity = (str(node_id), descriptor.code)
        prior = self._by_identity.get(identity)
        if prior is not None:
            return prior
        if native_code != descriptor.code:
            raise RecoveryFatalError(native_code, "outside recovery contract") from None
        decision = decide_recovery(
            native_code, descriptor.fallback_kind, descriptor.placement
        )

        if _callback_failed(rollback, checkpoint_token):
            raise RecoveryFatalError(
                "LOCAL_MUTATION_ROLLBACK_FAILED", "native rollback failed"
            ) from None
        if _callback_failed(fallback_attempt, descriptor):
            raise RecoveryFatalError(
                "DEGRADATION_FALLBACK_FAILED", "declared fallback failed"
            ) from None
        if _callback_failed(insert_notice, str(node_id), descriptor):
            raise RecoveryFatalError(
                "DEGRADATION_INSERT_FAILED", "visible notice insertion failed"
            ) from None

        self._by_identity[identity] = decision
        self._ordered.append(decision)
        return decision


def _callback_failed(callback: Callable[..., Any], *args: Any) -> bool:
    failed = False
    try:
        callback(*args)
    except Exception:
        failed = True
    return failed


JS_RECOVERY_MATRIX_BEGIN = "  // BEGIN WPSCOMPOSER GENERATED RECOVERY MATRIX\n"
JS_RECOVERY_MATRIX_END = "  // END WPSCOMPOSER GENERATED RECOVERY MATRIX\n"


def _js_matrix_payload() -> Mapping[str, Mapping[str, Mapping[str, Any]]]:
    return {
        code: {
            rule.fallback_kind: {
                "fallbackAttempts": rule.fallback_attempts,
                "placement": rule.placement,
            }
            for rule in rules
        }
        for code, rules in RECOVERY_MATRIX.items()
    }


def render_js_recovery_matrix() -> str:
    """Render the exact generated frozen-table block tracked by the add-in."""

    lines = [JS_RECOVERY_MATRIX_BEGIN.rstrip("\n")]
    payload = _js_matrix_payload()
    lines.append("  const WPSCOMPOSER_RECOVERY_MATRIX = Object.freeze({")
    codes = sorted(payload)
    for code_index, code in enumerate(codes):
        suffix = "," if code_index < len(codes) - 1 else ""
        lines.append(f"    {json.dumps(code)}: Object.freeze({{")
        branches = payload[code]
        fallback_kinds = sorted(branches)
        for branch_index, fallback_kind in enumerate(fallback_kinds):
            branch_suffix = "," if branch_index < len(fallback_kinds) - 1 else ""
            value = json.dumps(
                branches[fallback_kind],
                ensure_ascii=True,
                sort_keys=True,
                separators=(",", ":"),
            )
            lines.append(
                f"      {json.dumps(fallback_kind)}: Object.freeze({value})"
                f"{branch_suffix}"
            )
        lines.append(f"    }}){suffix}")
    lines.append("  });")
    lines.append(JS_RECOVERY_MATRIX_END.rstrip("\n"))
    return "\n".join(lines) + "\n"


__all__ = [
    "FATAL_BOUNDARY_CODES",
    "JS_RECOVERY_MATRIX_BEGIN",
    "JS_RECOVERY_MATRIX_END",
    "RECOVERY_MATRIX",
    "DegradationDescriptor",
    "LocalRecoveryController",
    "RecoveryDecision",
    "RecoveryFatalError",
    "RecoveryRule",
    "controlled_token",
    "decide_recovery",
    "render_js_recovery_matrix",
]
