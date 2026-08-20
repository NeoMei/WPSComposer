from __future__ import annotations

import hashlib
from unittest.mock import patch

import pytest

from skills.WPSComposer.scripts.longform.bookmark_ids import (
    BOOKMARK_COLLISION_UNRESOLVED,
    BOOKMARK_ID_INVALID,
    BOOKMARK_KIND_INVALID,
    BookmarkMapResult,
    map_bookmarks,
)


def _expected_name(kind: str, external_id: str, attempt: int = 0) -> str:
    data = f"{kind}\x00{external_id}\x00{attempt}".encode("utf-8")
    digest = hashlib.sha256(data).hexdigest()
    return f"wpsc_{kind}_{digest[:24]}"


def test_map_bookmarks_creates_ascii_bookmarks_from_external_ids() -> None:
    result = map_bookmarks(["intro", "method"], "fig")

    assert result.issues == ()
    assert result.mapping == {
        "intro": _expected_name("fig", "intro"),
        "method": _expected_name("fig", "method"),
    }


def test_map_bookmarks_sorts_by_utf8_bytes_not_input_order() -> None:
    ids = ["zeta", "alpha", "beta"]
    result = map_bookmarks(ids, "tab")

    # Deterministic sort by UTF-8 bytes: alpha, beta, zeta
    assert list(result.mapping.keys()) == ["alpha", "beta", "zeta"]
    assert result.issues == ()


def test_map_bookmarks_sorts_by_utf8_bytes_not_ascii_alphabetical() -> None:
    # Valid IDs are ASCII-only; UTF-8 byte sort is case-sensitive ASCII order.
    # "Beta" (0x42) sorts before "alpha" (0x61) by UTF-8 bytes.
    ids = ["alpha", "Beta", "gamma"]
    result = map_bookmarks(ids, "ref")

    assert list(result.mapping.keys()) == ["Beta", "alpha", "gamma"]


def test_map_bookmarks_removes_duplicate_external_ids() -> None:
    result = map_bookmarks(["a", "a", "b"], "eq")

    assert list(result.mapping.keys()) == ["a", "b"]
    assert result.issues == ()


def test_map_bookmarks_rejects_invalid_kind() -> None:
    result = map_bookmarks(["a"], "unknown")

    assert result.mapping == {}
    assert result.issues == (BOOKMARK_KIND_INVALID,)


def test_map_bookmarks_rejects_invalid_external_ids() -> None:
    result = map_bookmarks(["", "123", "a b", "a"], "para")

    assert result.mapping == {"a": _expected_name("para", "a")}
    assert result.issues == (BOOKMARK_ID_INVALID, BOOKMARK_ID_INVALID, BOOKMARK_ID_INVALID)


def test_map_bookmarks_resolves_collisions_by_incrementing_attempt() -> None:
    # Force attempt 0 to collide while letting attempt 1 diverge.
    # The hash input ends with the decimal attempt after a NUL byte.
    def fake_sha256(data: bytes) -> object:
        class Fake:
            def hexdigest(self) -> str:
                if data.endswith(bytes.fromhex("00") + b"0"):
                    return "a" * 64
                return "b" * 64
        return Fake()

    with patch.object(hashlib, "sha256", fake_sha256):
        result = map_bookmarks(["x", "y"], "fig")

    # x consumes attempt 0; y must use attempt 1.
    assert len(result.mapping) == 2
    assert result.issues == ()
    assert result.mapping["x"].endswith("a" * 24)
    assert result.mapping["y"].endswith("b" * 24)


def test_map_bookmarks_reports_unresolvable_collision() -> None:
    def fake_sha256(data: bytes) -> object:
        class Fake:
            def hexdigest(self) -> str:
                return "b" * 64
        return Fake()

    with patch.object(hashlib, "sha256", fake_sha256):
        result = map_bookmarks(["x", "y"], "fig")

    # Only 256 attempts available; second ID cannot be assigned.
    assert "x" in result.mapping
    assert "y" not in result.mapping
    assert result.issues == (BOOKMARK_COLLISION_UNRESOLVED,)


def test_bookmark_name_format_is_ascii_and_24_hex() -> None:
    result = map_bookmarks(["my-id"], "head")

    assert result.issues == ()
    name = result.mapping["my-id"]
    assert name.startswith("wpsc_head_")
    suffix = name[len("wpsc_head_") :]
    assert len(suffix) == 24
    assert suffix == suffix.lower()
    assert set(suffix).issubset(set("0123456789abcdef"))

