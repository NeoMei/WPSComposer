"""Canonical privacy classification for serialized long-form plan text."""

from __future__ import annotations

import re
from typing import Any


_HASH_RE = re.compile(r"(?i)(?<![0-9a-f])[0-9a-f]{64}(?![0-9a-f])")
_EXCEPTION_RE = re.compile(
    r"\b(?:Traceback|[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception))\s*(?:\(|\b)"
)
_URI_RE = re.compile(r"(?i)\b[A-Za-z][A-Za-z0-9+.-]*://")
_WINDOWS_RE = re.compile(r"(?i)(?<![A-Za-z0-9])[A-Za-z]:[\\/][^\s|,;]+")
_UNC_RE = re.compile(r"(?:\\\\|(?<!:)//)[A-Za-z0-9_.-]+[\\/][^\s|,;]+")
_TILDE_RE = re.compile(r"(?:^|[\s({=:\"'])~[\\/]")
_TRAVERSAL_RE = re.compile(r"(?:^|[\\/])\.\.[\\/]")
_DOT_RELATIVE_RE = re.compile(r"(?:^|[\s({=:\"'])\.\.?[\\/]")
_RELATIVE_FILE_RE = re.compile(
    r"(?<![A-Za-z0-9_:-])(?:[A-Za-z0-9_.-]+[\\/])+"
    r"[A-Za-z0-9_.-]+\.[A-Za-z][A-Za-z0-9]{0,15}"
    r"(?![A-Za-z0-9_.-])"
)
_BASE64_RE = re.compile(
    r"(?<![A-Za-z0-9+/])[A-Za-z0-9+/]{76,}={0,2}(?![A-Za-z0-9+/])"
)
_LABELLED_RELATIVE_RE = re.compile(
    r"(?i)\b(?:path|source|sourcePath|stagingPath|file)\s*[:=]\s*"
    r"(?:\.\.?[\\/]|[^\s|,;]+[\\/][^\s|,;]+)"
)


def _has_base64_payload(text: str) -> bool:
    return any(len(match.group(0)) % 4 == 0 for match in _BASE64_RE.finditer(text))


def _has_posix_absolute_path(text: str) -> bool:
    for match in re.finditer(r"/(?!/)(?=\S)", text):
        index = match.start()
        if index == 0:
            return True
        previous = text[index - 1]
        if previous.isspace() or previous in "({=\"'":
            return True
        if previous == ":" and text[index + 1:index + 2] != "/":
            return True
    return False


def contains_private_plan_text(value: Any) -> bool:
    """Return whether text contains private diagnostics or encoded payloads."""

    text = str(value or "")
    return bool(
        _HASH_RE.search(text)
        or _EXCEPTION_RE.search(text)
        or _URI_RE.search(text)
        or _WINDOWS_RE.search(text)
        or _UNC_RE.search(text)
        or _TILDE_RE.search(text)
        or _TRAVERSAL_RE.search(text)
        or _DOT_RELATIVE_RE.search(text)
        or _RELATIVE_FILE_RE.search(text)
        or _has_base64_payload(text)
        or _LABELLED_RELATIVE_RE.search(text)
        or re.search(r"(?i)data:[^,;\s]{0,80};base64,", text)
        or _has_posix_absolute_path(text)
    )


__all__ = ["contains_private_plan_text"]
