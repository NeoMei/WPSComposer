"""Shared native-field refresh and convergence contract.

The module is deliberately platform-pure.  Windows COM and macOS JSAPI
implement :class:`NativeFieldAdapter`; this module owns ordering, bounded
convergence, canonical snapshots, and privacy-safe instability reporting.
"""

from __future__ import annotations

import hashlib
import json
import unicodedata
from typing import Any, Protocol, Tuple, runtime_checkable

from .executor import (
    FIELD_REFRESH_UNSTABLE,
    ConvergenceResult,
    ExecutionIssue,
    FieldSnapshot,
)


FIELD_KINDS = frozenset(
    {
        "STYLEREF",
        "SEQ_FIG",
        "SEQ_TAB",
        "SEQ_EQ",
        "REF",
        "TOC",
        "TOF_FIG",
        "TOF_TAB",
        "PAGE",
        "NUMPAGES",
    }
)


class NativeFieldContractError(RuntimeError):
    """Fatal failure of a required native-field phase.

    Exception text intentionally excludes the platform exception.  Native API
    errors can contain visible field values, bookmark names, or local paths;
    the original exception is deliberately discarded.  Contract helpers raise
    only after leaving exception handlers, yielding neither ``__cause__`` nor
    ``__context__`` and keeping formatted tracebacks privacy-safe.
    """

    def __init__(self, phase: str, reason: str = "failed") -> None:
        self.phase = phase
        self.reason = reason
        super().__init__(f"Required native field API {phase} {reason}")


@runtime_checkable
class NativeFieldAdapter(Protocol):
    """Five-phase native field adapter implemented by each WPS platform."""

    def repaginate_and_update_numbering(self) -> None:
        ...

    def refresh_bookmarks_and_references(self) -> None:
        ...

    def refresh_indexes(self) -> None:
        ...

    def repaginate_and_update_page_fields(self) -> None:
        ...

    def snapshot_fields(self) -> Tuple[FieldSnapshot, ...]:
        ...


def _normalize_visible_result(value: Any) -> str:
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return unicodedata.normalize("NFC", text)


def snapshot_visible_field(
    *,
    owner_node_id: str,
    field_kind: str,
    ordinal_within_node: int,
    visible_result: Any,
    field_category: str,
    toc_page_count: int = 0,
    figure_index_page_count: int = 0,
    table_index_page_count: int = 0,
    total_pages: int = 0,
) -> FieldSnapshot:
    """Hash one normalized visible field result without retaining its text."""

    normalized = _normalize_visible_result(visible_result)
    result_hash = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    return FieldSnapshot(
        stable_key=(owner_node_id, field_kind, ordinal_within_node),
        field_category=field_category,
        result_hash=result_hash,
        toc_page_count=toc_page_count,
        figure_index_page_count=figure_index_page_count,
        table_index_page_count=table_index_page_count,
        total_pages=total_pages,
    )


_PHASES = (
    "repaginate_and_update_numbering",
    "refresh_bookmarks_and_references",
    "refresh_indexes",
    "repaginate_and_update_page_fields",
)


def _require_adapter(adapter: Any) -> None:
    failed_phase = None
    for phase in (*_PHASES, "snapshot_fields"):
        try:
            available = callable(getattr(adapter, phase, None))
        except Exception:
            available = False
        if not available:
            failed_phase = phase
            break
    if failed_phase is not None:
        raise NativeFieldContractError(failed_phase, "is missing") from None


def _invoke(adapter: NativeFieldAdapter, phase: str) -> Any:
    failed = False
    result = None
    try:
        method = getattr(adapter, phase)
        result = method()
    except Exception:
        failed = True
    if failed:
        # Raise after leaving the handler so neither __cause__ nor __context__
        # retains platform exception text, paths, field values, or bookmarks.
        raise NativeFieldContractError(phase) from None
    return result


def _is_sha256(value: str) -> bool:
    return (
        len(value) == 64
        and value == value.lower()
        and all(character in "0123456789abcdef" for character in value)
    )


def _canonical_snapshot_unchecked(
    raw: Any,
    *,
    allow_legacy_hashes: bool,
) -> Tuple[FieldSnapshot, ...]:
    snapshot = tuple(raw or ())

    keys: set[Tuple[str, str, int]] = set()
    for item in snapshot:
        if type(item) is not FieldSnapshot:
            raise NativeFieldContractError("snapshot_fields", "returned an invalid field snapshot")
        key = item.stable_key
        if (
            not isinstance(key, tuple)
            or len(key) != 3
            or not isinstance(key[0], str)
            or not key[0]
            or key[1] not in FIELD_KINDS
            or not isinstance(key[2], int)
            or isinstance(key[2], bool)
            or key[2] < 0
            or key in keys
        ):
            raise NativeFieldContractError("snapshot_fields", "returned an invalid field snapshot")
        if not allow_legacy_hashes and not _is_sha256(item.result_hash):
            raise NativeFieldContractError("snapshot_fields", "returned an invalid field snapshot")
        counts = (
            item.toc_page_count,
            item.figure_index_page_count,
            item.table_index_page_count,
            item.total_pages,
        )
        if any(
            not isinstance(count, int) or isinstance(count, bool) or count < 0
            for count in counts
        ):
            raise NativeFieldContractError("snapshot_fields", "returned an invalid field snapshot")
        keys.add(key)
    return tuple(sorted(snapshot))


def _canonical_snapshot(
    raw: Any,
    *,
    allow_legacy_hashes: bool,
) -> Tuple[FieldSnapshot, ...]:
    failed = False
    snapshot: Tuple[FieldSnapshot, ...] = ()
    try:
        snapshot = _canonical_snapshot_unchecked(
            raw,
            allow_legacy_hashes=allow_legacy_hashes,
        )
    except Exception:
        failed = True
    if failed:
        raise NativeFieldContractError(
            "snapshot_fields", "returned an invalid field snapshot"
        ) from None
    return snapshot


def _snapshot_digest_unchecked(snapshot: Tuple[FieldSnapshot, ...]) -> str:
    canonical = json.dumps(
        [item.to_dict() for item in snapshot],
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest()


def _snapshot_digest(snapshot: Tuple[FieldSnapshot, ...]) -> str:
    failed = False
    digest = ""
    try:
        digest = _snapshot_digest_unchecked(snapshot)
    except Exception:
        failed = True
    if failed:
        raise NativeFieldContractError("snapshot_digest") from None
    return digest


def _unstable_issue(snapshot: Tuple[FieldSnapshot, ...], rounds: int) -> ExecutionIssue:
    if snapshot:
        representative = snapshot[-1]
        toc_pages = representative.toc_page_count
        figure_pages = representative.figure_index_page_count
        table_pages = representative.table_index_page_count
        total_pages = representative.total_pages
    else:
        toc_pages = figure_pages = table_pages = total_pages = 0
    message = (
        f"Field refresh did not converge after {rounds} rounds; "
        f"fields={len(snapshot)}, toc_pages={toc_pages}, "
        f"figure_index_pages={figure_pages}, table_index_pages={table_pages}, "
        f"total_pages={total_pages}"
    )
    return ExecutionIssue(
        code=FIELD_REFRESH_UNSTABLE,
        message=message,
        placement="document",
    )


def _finalize_native_fields(
    adapter: NativeFieldAdapter,
    max_rounds: int,
    *,
    allow_legacy_hashes: bool,
) -> ConvergenceResult:
    if type(max_rounds) is not int or not 1 <= max_rounds <= 3:
        raise NativeFieldContractError(
            "max_rounds", "must be an integer from 1 through 3"
        ) from None
    _require_adapter(adapter)
    rounds_limit = max_rounds
    previous_digest: str | None = None

    for round_index in range(rounds_limit):
        for phase in _PHASES:
            _invoke(adapter, phase)
        snapshot = _canonical_snapshot(
            _invoke(adapter, "snapshot_fields"),
            allow_legacy_hashes=allow_legacy_hashes,
        )
        digest = _snapshot_digest(snapshot)
        if previous_digest is not None and digest == previous_digest:
            return ConvergenceResult(snapshot=snapshot, issues=(), rounds=round_index + 1)
        previous_digest = digest

    # One frozen diagnostic snapshot: no native mutation phase may run here.
    frozen = _canonical_snapshot(
        _invoke(adapter, "snapshot_fields"),
        allow_legacy_hashes=allow_legacy_hashes,
    )
    return ConvergenceResult(
        snapshot=frozen,
        issues=(_unstable_issue(frozen, rounds_limit + 1),),
        rounds=rounds_limit + 1,
    )


def evaluate_field_snapshot_history(
    history: Any,
    max_rounds: int = 3,
    *,
    allow_legacy_hashes: bool = False,
) -> ConvergenceResult:
    """Validate convergence already executed by a remote native adapter.

    A remote executor must stop immediately after the second adjacent equal
    snapshot.  If all mutation rounds change, it appends exactly one read-only
    diagnostic snapshot.  This function never replays or mutates native state.
    """

    if type(max_rounds) is not int or not 1 <= max_rounds <= 3:
        raise NativeFieldContractError(
            "max_rounds", "must be an integer from 1 through 3"
        ) from None

    failed = False
    raw_rounds: Tuple[Any, ...] = ()
    try:
        raw_rounds = tuple(history)
    except Exception:
        failed = True
    if failed:
        raise NativeFieldContractError("field_history", "is invalid") from None

    rounds = tuple(
        _canonical_snapshot(raw, allow_legacy_hashes=allow_legacy_hashes)
        for raw in raw_rounds
    )
    digests = tuple(_snapshot_digest(snapshot) for snapshot in rounds)
    round_count = len(rounds)

    if 2 <= round_count <= max_rounds:
        equal_positions = tuple(
            index
            for index in range(1, round_count)
            if digests[index] == digests[index - 1]
        )
        if equal_positions == (round_count - 1,):
            return ConvergenceResult(
                snapshot=rounds[-1],
                issues=(),
                rounds=round_count,
            )
        raise NativeFieldContractError("field_history", "is inconsistent") from None

    if round_count == max_rounds + 1:
        converged_before_bound = any(
            digests[index] == digests[index - 1]
            for index in range(1, max_rounds)
        )
        if converged_before_bound:
            raise NativeFieldContractError("field_history", "is inconsistent") from None
        frozen = rounds[-1]
        return ConvergenceResult(
            snapshot=frozen,
            issues=(_unstable_issue(frozen, max_rounds + 1),),
            rounds=max_rounds + 1,
        )

    raise NativeFieldContractError("field_history", "is incomplete") from None


def finalize_native_fields(
    adapter: NativeFieldAdapter,
    max_rounds: int = 3,
) -> ConvergenceResult:
    """Run the single bounded five-phase native-field convergence algorithm."""

    return _finalize_native_fields(adapter, max_rounds, allow_legacy_hashes=False)


__all__ = [
    "FIELD_KINDS",
    "NativeFieldAdapter",
    "NativeFieldContractError",
    "evaluate_field_snapshot_history",
    "finalize_native_fields",
    "snapshot_visible_field",
]
