"""Canonical privacy classification for serialized long-form plan text."""

from __future__ import annotations

import json
import re
from typing import Any


_HASH_RE = re.compile(r"(?i)(?<![0-9a-f])[0-9a-f]{64}(?![0-9a-f])")
_EXCEPTION_RE = re.compile(
    r"\b(?:Traceback|[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception))\s*(?:\(|\b)"
)
_URI_RE = re.compile(r"(?i)\b[A-Za-z][A-Za-z0-9+.-]*://")
_OPAQUE_URI_RE = re.compile(r"(?i)\b(?:data|blob):")
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
        or _OPAQUE_URI_RE.search(text)
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


def redact_private_text(value: Any) -> str:
    """Return public text, replacing any private-bearing value atomically."""

    text = str(value or "")
    return "<redacted>" if contains_private_plan_text(text) else text


JS_PRIVACY_FILTER_BEGIN = "  // BEGIN WPSCOMPOSER GENERATED PRIVACY FILTER\n"
JS_PRIVACY_FILTER_END = "  // END WPSCOMPOSER GENERATED PRIVACY FILTER\n"

# JavaScript-compatible equivalents of the canonical classifier above. They
# avoid lookbehind for older WPS JSAPI engines and redact the whole value.
_JS_PRIVATE_PATTERNS = (
    (r"[0-9a-f]{64}", "i"),
    (r"(?:Traceback|[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception))\s*(?:\(|\b)", ""),
    (r"[A-Za-z][A-Za-z0-9+.-]*://", "i"),
    (r"(?:data|blob):", "i"),
    (r"[A-Za-z]:[\\/][^\s|,;]+", "i"),
    (r"(?:\\\\|//)[A-Za-z0-9_.-]+[\\/][^\s|,;]+", ""),
    (r"(?:^|[\s({=:\"'])~[\\/]", ""),
    (r"(?:^|[\\/])\.\.[\\/]", ""),
    (r"(?:^|[\s({=:\"'])\.\.?[\\/]", ""),
    (
        r"(?:^|[^A-Za-z0-9_:-])(?:[A-Za-z0-9_.-]+[\\/])+"
        r"[A-Za-z0-9_.-]+\.[A-Za-z][A-Za-z0-9]{0,15}(?:$|[^A-Za-z0-9_.-])",
        "",
    ),
    (r"(?:^|[^A-Za-z0-9+/])[A-Za-z0-9+/]{76,}={0,2}(?:$|[^A-Za-z0-9+/])", ""),
    (
        r"\b(?:path|source|sourcePath|stagingPath|file)\s*[:=]\s*"
        r"(?:\.\.?[\\/]|[^\s|,;]+[\\/][^\s|,;]+)",
        "i",
    ),
    (r"(?:^|[\s({=:\"']|:(?!/))/(?!/)(?=\S)", ""),
)


def render_js_privacy_filter() -> str:
    """Render the frozen JS classifier from this canonical Python module."""

    lines = [JS_PRIVACY_FILTER_BEGIN.rstrip("\n")]
    lines.append("  const WPSCOMPOSER_PRIVATE_PATTERNS = Object.freeze([")
    for index, (source, flags) in enumerate(_JS_PRIVATE_PATTERNS):
        suffix = "," if index < len(_JS_PRIVATE_PATTERNS) - 1 else ""
        payload = json.dumps(
            {"flags": flags, "source": source},
            ensure_ascii=True,
            sort_keys=True,
            separators=(",", ":"),
        )
        lines.append(f"    Object.freeze({payload}){suffix}")
    lines.extend((
        "  ]);",
        "",
        "  function safePublicText(value) {",
        "    const text = safeString(value);",
        "    const privateValue = WPSCOMPOSER_PRIVATE_PATTERNS.some(function (item) {",
        "      return new RegExp(item.source, item.flags).test(text);",
        "    });",
        '    return privateValue ? "<redacted>" : text;',
        "  }",
        JS_PRIVACY_FILTER_END.rstrip("\n"),
    ))
    return "\n".join(lines) + "\n"


__all__ = [
    "JS_PRIVACY_FILTER_BEGIN",
    "JS_PRIVACY_FILTER_END",
    "contains_private_plan_text",
    "redact_private_text",
    "render_js_privacy_filter",
]
