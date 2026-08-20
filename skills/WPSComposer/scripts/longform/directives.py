"""Shared lexer for top-level long-form Markdown block directives.

The scanner recognizes directive boundaries without interpreting directive
names.  Invalid local syntax is represented on an immutable result object;
readable block content is always retained and malformed input never reaches a
second, platform-specific parser.
"""

from __future__ import annotations

from dataclasses import dataclass
import json
import re
import unicodedata
from typing import Iterator, Optional, Union


DIRECTIVE_SYNTAX_INVALID = "DIRECTIVE_SYNTAX_INVALID"
DIRECTIVE_UNCLOSED = "DIRECTIVE_UNCLOSED"
NESTED_DIRECTIVE_UNSUPPORTED = "NESTED_DIRECTIVE_UNSUPPORTED"

_NAME_RE = re.compile(r"^[a-z][a-z0-9-]{0,31}$")
_KEY_RE = re.compile(r"[a-z][a-z0-9_]{0,31}")
_IDENTIFIER_RE = re.compile(r"[A-Za-z][A-Za-z0-9._:-]{0,127}")
_TOKEN_RE = re.compile(r"[A-Za-z0-9._:/+%-]+")
_OPENING_RE = re.compile(
    r"^:::([a-z][a-z0-9-]{0,31})(?:[ \t]+\{(.*)\})?[ \t]*$"
)
_CLOSING_RE = re.compile(r"^:::[ \t]*$")


class _InvalidDirective(ValueError):
    pass


@dataclass(frozen=True)
class BlockDirective:
    """One closed or degraded block directive found in source order."""

    name: str
    identifier: Optional[str]
    attributes: dict[str, str]
    body: str
    start_line: int
    issues: tuple[str, ...]

    def __iter__(self) -> Iterator[object]:
        yield self.name
        yield self.identifier
        yield self.attributes
        yield self.body
        yield self.start_line
        yield self.issues


DirectiveRegion = Union[str, BlockDirective]


@dataclass(frozen=True)
class _Fence:
    marker: str
    length: int


def _line_content(line: str) -> str:
    if line.endswith("\r\n"):
        return line[:-2]
    if line.endswith(("\n", "\r")):
        return line[:-1]
    return line


def _source_lines(markdown: str) -> list[str]:
    """Split only on Markdown line endings, not Unicode separator controls."""
    lines = []
    start = 0
    index = 0
    while index < len(markdown):
        character = markdown[index]
        if character == "\n":
            lines.append(markdown[start : index + 1])
            index += 1
            start = index
            continue
        if character == "\r":
            end = index + 2 if markdown.startswith("\r\n", index) else index + 1
            lines.append(markdown[start:end])
            index = end
            start = index
            continue
        index += 1
    if start < len(markdown):
        lines.append(markdown[start:])
    return lines


def _is_forbidden_character(character: str) -> bool:
    codepoint = ord(character)
    return (
        codepoint <= 0x1F
        or 0x7F <= codepoint <= 0x9F
        or 0xD800 <= codepoint <= 0xDFFF
    )


def _validated_text(value: str) -> str:
    if any(_is_forbidden_character(character) for character in value):
        raise _InvalidDirective("directive attribute contains unsafe Unicode")
    return unicodedata.normalize("NFC", value)


def _json_string_end(source: str, start: int) -> int:
    index = start + 1
    escaped = False
    while index < len(source):
        character = source[index]
        if escaped:
            escaped = False
        elif character == "\\":
            escaped = True
        elif character == '"':
            return index + 1
        index += 1
    raise _InvalidDirective("unterminated JSON string attribute")


def _decode_json_string(source: str) -> str:
    try:
        value = json.loads(source)
    except (UnicodeError, ValueError) as error:
        raise _InvalidDirective("invalid JSON string attribute") from error
    if not isinstance(value, str):
        raise _InvalidDirective("directive attribute must decode to a string")
    return _validated_text(value)


def _parse_attributes(source: str) -> tuple[Optional[str], dict[str, str]]:
    identifier: Optional[str] = None
    attributes: dict[str, str] = {}
    index = 0

    while True:
        while index < len(source) and source[index] in " \t":
            index += 1
        if index == len(source):
            return identifier, attributes

        if source[index] == "#":
            end = index + 1
            while end < len(source) and source[end] not in " \t":
                end += 1
            candidate = source[index + 1 : end]
            if identifier is not None or _IDENTIFIER_RE.fullmatch(candidate) is None:
                raise _InvalidDirective("invalid or duplicate directive identifier")
            identifier = candidate
            index = end
            continue

        key_match = _KEY_RE.match(source, index)
        if key_match is None:
            raise _InvalidDirective("invalid directive attribute key")
        key = key_match.group(0)
        index = key_match.end()
        if index >= len(source) or source[index] != "=":
            raise _InvalidDirective("directive attribute assignment is missing")
        index += 1
        if index >= len(source):
            raise _InvalidDirective("directive attribute value is missing")

        if source[index] == '"':
            end = _json_string_end(source, index)
            value = _decode_json_string(source[index:end])
        else:
            token_match = _TOKEN_RE.match(source, index)
            if token_match is None:
                raise _InvalidDirective("invalid unquoted directive token")
            end = token_match.end()
            value = _validated_text(token_match.group(0))

        if end < len(source) and source[end] not in " \t":
            raise _InvalidDirective("directive attributes must be whitespace-separated")
        if key in attributes:
            raise _InvalidDirective("duplicate directive attribute key")
        attributes[key] = value
        index = end


def _fallback_name(line: str) -> str:
    remainder = line[3:].lstrip(" \t")
    end = 0
    while end < len(remainder) and remainder[end] not in " \t{":
        end += 1
    return remainder[:end]


def _parse_opening_line(
    line: str,
) -> tuple[str, Optional[str], dict[str, str], tuple[str, ...]]:
    match = _OPENING_RE.fullmatch(line)
    if match is None:
        return _fallback_name(line), None, {}, (DIRECTIVE_SYNTAX_INVALID,)

    name = match.group(1)
    if _NAME_RE.fullmatch(name) is None:
        return name, None, {}, (DIRECTIVE_SYNTAX_INVALID,)
    attributes_source = match.group(2)
    if attributes_source is None:
        return name, None, {}, ()
    try:
        identifier, attributes = _parse_attributes(attributes_source)
    except (UnicodeError, ValueError):
        return name, None, {}, (DIRECTIVE_SYNTAX_INVALID,)
    return name, identifier, attributes, ()


def _opening_fence(line: str) -> Optional[_Fence]:
    indentation = len(line) - len(line.lstrip(" "))
    if indentation > 3:
        return None
    content = line[indentation:]
    if not content or content[0] not in {"`", "~"}:
        return None
    marker = content[0]
    length = len(content) - len(content.lstrip(marker))
    if length < 3:
        return None
    if marker == "`" and "`" in content[length:]:
        return None
    return _Fence(marker, length)


def _is_closing_fence(line: str, fence: _Fence) -> bool:
    indentation = len(line) - len(line.lstrip(" "))
    if indentation > 3:
        return False
    content = line[indentation:]
    length = len(content) - len(content.lstrip(fence.marker))
    return length >= fence.length and not content[length:].strip()


def _is_valid_opening_line(line: str) -> bool:
    if _CLOSING_RE.fullmatch(line) is not None:
        return False
    name, _, _, issues = _parse_opening_line(line)
    return bool(name) and not issues


def _append_issue(issues: list[str], issue: str) -> None:
    if issue not in issues:
        issues.append(issue)


def scan_block_directives(markdown: str) -> tuple[DirectiveRegion, ...]:
    """Split Markdown into literal regions and top-level block directives.

    Markdown regions are returned as exact source strings.  Directive bodies
    likewise preserve their original line endings, while parsed attribute
    values are normalized to NFC.
    """
    if not isinstance(markdown, str):
        raise TypeError("markdown must be a string")

    lines = _source_lines(markdown)
    regions: list[DirectiveRegion] = []
    pending_markdown: list[str] = []
    fence: Optional[_Fence] = None
    index = 0

    def flush_markdown() -> None:
        if pending_markdown:
            regions.append("".join(pending_markdown))
            pending_markdown.clear()

    while index < len(lines):
        line = lines[index]
        content = _line_content(line)

        if fence is not None:
            pending_markdown.append(line)
            if _is_closing_fence(content, fence):
                fence = None
            index += 1
            continue

        opening_fence = _opening_fence(content)
        if opening_fence is not None:
            pending_markdown.append(line)
            fence = opening_fence
            index += 1
            continue

        if not content.startswith(":::") or _CLOSING_RE.fullmatch(content):
            pending_markdown.append(line)
            index += 1
            continue

        flush_markdown()
        name, identifier, attributes, opening_issues = _parse_opening_line(content)
        issues = list(opening_issues)
        start_line = index + 1
        index += 1
        body_lines: list[str] = []
        body_fence: Optional[_Fence] = None
        nested_depth = 0
        closed = False

        while index < len(lines):
            body_line = lines[index]
            body_content = _line_content(body_line)

            if body_fence is not None:
                body_lines.append(body_line)
                if _is_closing_fence(body_content, body_fence):
                    body_fence = None
                index += 1
                continue

            opening_body_fence = _opening_fence(body_content)
            if opening_body_fence is not None:
                body_lines.append(body_line)
                body_fence = opening_body_fence
                index += 1
                continue

            if _CLOSING_RE.fullmatch(body_content):
                if nested_depth:
                    nested_depth -= 1
                    body_lines.append(body_line)
                    index += 1
                    continue
                closed = True
                index += 1
                break

            if body_content.startswith(":::") and _is_valid_opening_line(body_content):
                nested_depth += 1
                _append_issue(issues, NESTED_DIRECTIVE_UNSUPPORTED)

            body_lines.append(body_line)
            index += 1

        if not closed:
            _append_issue(issues, DIRECTIVE_UNCLOSED)

        regions.append(
            BlockDirective(
                name=name,
                identifier=identifier,
                attributes=attributes,
                body="".join(body_lines),
                start_line=start_line,
                issues=tuple(issues),
            )
        )

    flush_markdown()
    return tuple(regions)


__all__ = [
    "DIRECTIVE_SYNTAX_INVALID",
    "DIRECTIVE_UNCLOSED",
    "NESTED_DIRECTIVE_UNSUPPORTED",
    "BlockDirective",
    "DirectiveRegion",
    "scan_block_directives",
]
