"""Deterministic WPS bookmark mapping from external semantic IDs.

External IDs are never used directly as WPS bookmark names.  Instead, each
valid external ID is salted with its kind and a collision counter, hashed
with SHA-256, and formatted as a fixed ASCII string.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from typing import Iterable


BOOKMARK_ID_INVALID = "BOOKMARK_ID_INVALID"
BOOKMARK_KIND_INVALID = "BOOKMARK_KIND_INVALID"
BOOKMARK_COLLISION_UNRESOLVED = "BOOKMARK_COLLISION_UNRESOLVED"

_VALID_KINDS = frozenset({"fig", "tab", "eq", "ref", "head", "para"})
_MAX_ATTEMPTS = 256
_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9._:-]{0,127}$")


@dataclass(frozen=True)
class BookmarkMapResult:
    """Immutable result of mapping external IDs to internal WPS bookmark names."""

    mapping: dict[str, str]
    issues: tuple[str, ...]

    def __iter__(self) -> Iterable[object]:
        yield self.mapping
        yield self.issues


def _bookmark_name(kind: str, external_id: str, attempt: int) -> str:
    """Build the internal bookmark name for one collision attempt."""
    data = f"{kind}\x00{external_id}\x00{attempt}".encode("utf-8")
    digest = hashlib.sha256(data).hexdigest()
    return f"wpsc_{kind}_{digest[:24]}"


def map_bookmarks(external_ids: Iterable[str], kind: str) -> BookmarkMapResult:
    """Map validated external IDs to deterministic WPS bookmark names.

    The mapping is ordered by UTF-8 byte sort of the external IDs, not by
    document position.  Collisions are resolved by incrementing the attempt
    counter up to 256 times before emitting BOOKMARK_COLLISION_UNRESOLVED.
    """
    if kind not in _VALID_KINDS:
        return BookmarkMapResult({}, (BOOKMARK_KIND_INVALID,))

    used: set[str] = set()
    mapping: dict[str, str] = {}
    issues: list[str] = []

    # Deterministic ordering by UTF-8 bytes; de-duplicate in that order.
    sorted_ids = sorted(set(external_ids), key=lambda item: item.encode("utf-8"))

    for external_id in sorted_ids:
        if _ID_RE.fullmatch(external_id) is None:
            issues.append(BOOKMARK_ID_INVALID)
            continue

        assigned = False
        for attempt in range(_MAX_ATTEMPTS):
            name = _bookmark_name(kind, external_id, attempt)
            if name not in used:
                used.add(name)
                mapping[external_id] = name
                assigned = True
                break

        if not assigned:
            issues.append(BOOKMARK_COLLISION_UNRESOLVED)

    return BookmarkMapResult(mapping, tuple(issues))


__all__ = [
    "BOOKMARK_COLLISION_UNRESOLVED",
    "BOOKMARK_ID_INVALID",
    "BOOKMARK_KIND_INVALID",
    "BookmarkMapResult",
    "map_bookmarks",
]
