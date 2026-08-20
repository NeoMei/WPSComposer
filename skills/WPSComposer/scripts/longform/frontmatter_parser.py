"""Bounded parser for the long-form frontmatter YAML data subset.

The parser deliberately composes only JSON-compatible scalar, list, and
mapping nodes.  It does not import a YAML object constructor or resolve tags.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
import unicodedata
from typing import Any, Iterator, Optional


FRONTMATTER_INVALID = "FRONTMATTER_INVALID"
FRONTMATTER_UNCLOSED = "FRONTMATTER_UNCLOSED"

MAX_BOUNDARY_BYTES = 64 * 1024
MAX_NODES = 256
MAX_DEPTH = 8

_INT_RE = re.compile(
    r"^[+-]?(?:0|[1-9][0-9_]*|0o[0-7_]+|0x[0-9a-fA-F_]+)$"
)
_FLOAT_RE = re.compile(
    r"^[+-]?(?:(?:[0-9][0-9_]*)?\.[0-9_]+"
    r"|[0-9][0-9_]*(?:\.[0-9_]*)?[eE][+-]?[0-9_]+)$"
)
_NONFINITE_RE = re.compile(r"^[+-]?\.(?:inf|nan)$", re.IGNORECASE)


class _InvalidFrontmatter(ValueError):
    pass


@dataclass(frozen=True)
class FrontmatterParseResult:
    """Safe frontmatter values plus the exact source span removed from body."""

    values: dict[str, Any]
    issues: tuple[str, ...]
    boundary: Optional[tuple[int, int]]
    body: str

    def __iter__(self) -> Iterator[Any]:
        yield self.values
        yield self.issues
        yield self.boundary
        yield self.body


@dataclass(frozen=True)
class _SourceLine:
    indent: int
    content: str
    number: int


class _NodeBudget:
    def __init__(self) -> None:
        self.count = 0

    def consume(self) -> None:
        self.count += 1
        if self.count > MAX_NODES:
            raise _InvalidFrontmatter("frontmatter exceeds the node limit")


def _line_content(line: str) -> str:
    if line.endswith("\r\n"):
        return line[:-2]
    if line.endswith(("\n", "\r")):
        return line[:-1]
    return line


def _utf8_length(value: str) -> int:
    return len(value.encode("utf-8", errors="surrogatepass"))


def _find_boundary(text: str) -> Optional[tuple[int, int, int]]:
    lines = text.splitlines(keepends=True)
    if not lines or _line_content(lines[0]) != "---":
        return None

    opener_end = len(lines[0])
    char_offset = opener_end
    byte_offset = 0
    for line in lines[1:]:
        line_bytes = _utf8_length(line)
        if _line_content(line) == "---":
            if byte_offset + len(b"---") > MAX_BOUNDARY_BYTES:
                return (opener_end, -1, -1)
            return (opener_end, char_offset, char_offset + len(line))
        if byte_offset + line_bytes > MAX_BOUNDARY_BYTES:
            return (opener_end, -1, -1)
        byte_offset += line_bytes
        char_offset += len(line)
    return (opener_end, -1, -1)


def _is_forbidden_character(character: str) -> bool:
    codepoint = ord(character)
    return (
        codepoint == 0
        or 0x01 <= codepoint <= 0x1F
        or 0x7F <= codepoint <= 0x9F
        or 0xD800 <= codepoint <= 0xDFFF
    )


def _validate_string(value: str) -> str:
    if any(_is_forbidden_character(character) for character in value):
        raise _InvalidFrontmatter("frontmatter string contains unsafe Unicode")
    return unicodedata.normalize("NFC", value)


def _strip_comment(value: str) -> str:
    quote: Optional[str] = None
    escaped = False
    index = 0
    while index < len(value):
        character = value[index]
        if quote == '"':
            if escaped:
                escaped = False
            elif character == "\\":
                escaped = True
            elif character == quote:
                quote = None
        elif quote == "'":
            if character == quote:
                if index + 1 < len(value) and value[index + 1] == quote:
                    index += 2
                    continue
                quote = None
        elif character in {'"', "'"}:
            quote = character
        elif character == "#" and (index == 0 or value[index - 1].isspace()):
            return value[:index]
        index += 1
    return value


def _prepare_lines(source: str) -> list[_SourceLine]:
    for character in source:
        if character not in {"\n", "\r"} and _is_forbidden_character(character):
            raise _InvalidFrontmatter("frontmatter contains unsafe Unicode")
    if "\t" in source:
        raise _InvalidFrontmatter("tabs are not accepted in frontmatter")

    result = []
    for number, raw_line in enumerate(source.splitlines(), start=2):
        indent = len(raw_line) - len(raw_line.lstrip(" "))
        content = _strip_comment(raw_line[indent:]).rstrip()
        if content:
            result.append(_SourceLine(indent, content, number))
    return result


def _split_mapping_entry(value: str) -> tuple[str, str]:
    quote: Optional[str] = None
    escaped = False
    nesting = 0
    index = 0
    while index < len(value):
        character = value[index]
        if quote == '"':
            if escaped:
                escaped = False
            elif character == "\\":
                escaped = True
            elif character == quote:
                quote = None
        elif quote == "'":
            if character == quote:
                if index + 1 < len(value) and value[index + 1] == quote:
                    index += 2
                    continue
                quote = None
        elif character in {'"', "'"}:
            quote = character
        elif character in "[{":
            nesting += 1
        elif character in "]}":
            nesting -= 1
            if nesting < 0:
                raise _InvalidFrontmatter("unbalanced flow collection")
        elif (
            character == ":"
            and nesting == 0
            and (index + 1 == len(value) or value[index + 1].isspace())
        ):
            key = value[:index].strip()
            if not key:
                raise _InvalidFrontmatter("mapping key is missing")
            return key, value[index + 1 :].strip()
        index += 1
    raise _InvalidFrontmatter("expected a mapping entry")


def _is_sequence_entry(value: str) -> bool:
    return value == "-" or value.startswith("- ")


def _looks_like_mapping(value: str) -> bool:
    try:
        _split_mapping_entry(value)
    except _InvalidFrontmatter:
        return False
    return True


def _decode_quoted(value: str, start: int = 0) -> tuple[str, int]:
    quote = value[start]
    index = start + 1
    output = []
    while index < len(value):
        character = value[index]
        if character == quote:
            if quote == "'" and index + 1 < len(value) and value[index + 1] == "'":
                output.append("'")
                index += 2
                continue
            return _validate_string("".join(output)), index + 1
        if quote == "'" or character != "\\":
            output.append(character)
            index += 1
            continue

        index += 1
        if index >= len(value):
            raise _InvalidFrontmatter("unfinished quoted escape")
        escape = value[index]
        simple = {
            '"': '"',
            "/": "/",
            "\\": "\\",
            "0": "\0",
            "a": "\a",
            "b": "\b",
            "t": "\t",
            "n": "\n",
            "v": "\v",
            "f": "\f",
            "r": "\r",
            "e": "\x1b",
            " ": " ",
            "_": "\xa0",
            "N": "\x85",
            "L": "\u2028",
            "P": "\u2029",
        }
        if escape in simple:
            output.append(simple[escape])
            index += 1
            continue
        widths = {"x": 2, "u": 4, "U": 8}
        width = widths.get(escape)
        if width is None:
            raise _InvalidFrontmatter("unsupported quoted escape")
        digits = value[index + 1 : index + 1 + width]
        if len(digits) != width or not all(
            character in "0123456789abcdefABCDEF" for character in digits
        ):
            raise _InvalidFrontmatter("invalid Unicode escape")
        try:
            output.append(chr(int(digits, 16)))
        except ValueError as error:
            raise _InvalidFrontmatter("invalid Unicode scalar value") from error
        index += width + 1
    raise _InvalidFrontmatter("unterminated quoted scalar")


def _resolve_plain(value: str) -> Any:
    lowered = value.lower()
    if value == "~" or lowered == "null":
        return None
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    compact = value.replace("_", "")
    if _INT_RE.fullmatch(value):
        sign = -1 if compact.startswith("-") else 1
        unsigned = compact.lstrip("+-")
        if unsigned.startswith("0o"):
            return sign * int(unsigned[2:], 8)
        if unsigned.startswith("0x"):
            return sign * int(unsigned[2:], 16)
        return int(compact, 10)
    if _NONFINITE_RE.fullmatch(value):
        raise _InvalidFrontmatter("non-finite numbers are not JSON-compatible")
    if _FLOAT_RE.fullmatch(value):
        number = float(compact)
        if not math.isfinite(number):
            raise _InvalidFrontmatter("non-finite numbers are not JSON-compatible")
        return number
    return _validate_string(value)


def _parse_scalar(value: str, *, key: bool = False) -> Any:
    value = value.strip()
    if not value:
        raise _InvalidFrontmatter("scalar is missing")
    if value[0] in {'"', "'"}:
        parsed, end = _decode_quoted(value)
        if value[end:].strip():
            raise _InvalidFrontmatter("unexpected text after quoted scalar")
        return parsed
    if value[0] in "!&*|>@`" or value.startswith("?"):
        raise _InvalidFrontmatter("unsupported YAML node feature")
    if re.search(r"(?:^|\s)[!&*](?:\S|$)", value):
        raise _InvalidFrontmatter("unsupported YAML node feature")
    if any(character in value for character in "[]{}"):
        raise _InvalidFrontmatter("unexpected flow collection marker")
    resolved = _resolve_plain(value)
    if key and not isinstance(resolved, str):
        raise _InvalidFrontmatter("mapping keys must be strings")
    if key and resolved == "<<":
        raise _InvalidFrontmatter("merge keys are not supported")
    return resolved


class _FlowParser:
    def __init__(self, source: str, budget: _NodeBudget) -> None:
        self.source = source
        self.budget = budget
        self.index = 0

    def parse(self, depth: int) -> Any:
        value = self._parse_value(depth)
        self._skip_space()
        if self.index != len(self.source):
            raise _InvalidFrontmatter("unexpected text after flow value")
        return value

    def _skip_space(self) -> None:
        while self.index < len(self.source) and self.source[self.index].isspace():
            self.index += 1

    def _parse_value(self, depth: int) -> Any:
        self._skip_space()
        if self.index >= len(self.source):
            raise _InvalidFrontmatter("flow value is missing")
        character = self.source[self.index]
        if character == "[":
            return self._parse_list(depth)
        if character == "{":
            return self._parse_mapping(depth)
        if character in {'"', "'"}:
            value, self.index = _decode_quoted(self.source, self.index)
            return value
        start = self.index
        while (
            self.index < len(self.source)
            and self.source[self.index] not in ",]}"
        ):
            self.index += 1
        return _parse_scalar(self.source[start : self.index])

    def _parse_list(self, depth: int) -> list[Any]:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        self.index += 1
        result = []
        self._skip_space()
        if self.index < len(self.source) and self.source[self.index] == "]":
            self.index += 1
            return result
        while True:
            self.budget.consume()
            result.append(self._parse_value(depth + 1))
            self._skip_space()
            if self.index >= len(self.source):
                raise _InvalidFrontmatter("unterminated flow list")
            marker = self.source[self.index]
            self.index += 1
            if marker == "]":
                return result
            if marker != ",":
                raise _InvalidFrontmatter("expected a flow list separator")

    def _parse_mapping_key(self) -> str:
        self._skip_space()
        if self.index >= len(self.source):
            raise _InvalidFrontmatter("flow mapping key is missing")
        if self.source[self.index] in {'"', "'"}:
            value, self.index = _decode_quoted(self.source, self.index)
        else:
            start = self.index
            while self.index < len(self.source) and self.source[self.index] != ":":
                if self.source[self.index] in "{},[]":
                    raise _InvalidFrontmatter("compound mapping keys are unsupported")
                self.index += 1
            value = _parse_scalar(self.source[start : self.index], key=True)
        self._skip_space()
        if self.index >= len(self.source) or self.source[self.index] != ":":
            raise _InvalidFrontmatter("flow mapping separator is missing")
        self.index += 1
        if value == "<<":
            raise _InvalidFrontmatter("merge keys are not supported")
        return value

    def _parse_mapping(self, depth: int) -> dict[str, Any]:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        self.index += 1
        result: dict[str, Any] = {}
        self._skip_space()
        if self.index < len(self.source) and self.source[self.index] == "}":
            self.index += 1
            return result
        while True:
            key = self._parse_mapping_key()
            if key in result:
                raise _InvalidFrontmatter("duplicate mapping key")
            self.budget.consume()
            result[key] = self._parse_value(depth + 1)
            self._skip_space()
            if self.index >= len(self.source):
                raise _InvalidFrontmatter("unterminated flow mapping")
            marker = self.source[self.index]
            self.index += 1
            if marker == "}":
                return result
            if marker != ",":
                raise _InvalidFrontmatter("expected a flow mapping separator")


def _parse_inline(value: str, budget: _NodeBudget, depth: int) -> Any:
    value = value.strip()
    if value.startswith(("[", "{")):
        return _FlowParser(value, budget).parse(depth)
    return _parse_scalar(value)


class _BlockParser:
    def __init__(self, source: str) -> None:
        self.lines = _prepare_lines(source)
        self.budget = _NodeBudget()

    def parse(self) -> dict[str, Any]:
        if not self.lines or self.lines[0].indent != 0:
            raise _InvalidFrontmatter("frontmatter root must be a mapping")
        if len(self.lines) == 1 and self.lines[0].content.startswith(("{", "[")):
            value = _parse_inline(self.lines[0].content, self.budget, 1)
            if not isinstance(value, dict):
                raise _InvalidFrontmatter("frontmatter root must be a mapping")
            return value
        value, index = self._parse_node(0, 0, 1)
        if index != len(self.lines) or not isinstance(value, dict):
            raise _InvalidFrontmatter("frontmatter root must be a mapping")
        return value

    def _parse_node(self, index: int, indent: int, depth: int) -> tuple[Any, int]:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        if index >= len(self.lines) or self.lines[index].indent != indent:
            raise _InvalidFrontmatter("invalid indentation")
        if _is_sequence_entry(self.lines[index].content):
            return self._parse_sequence(index, indent, depth)
        return self._parse_mapping(index, indent, depth)

    def _mapping_value(
        self, value_text: str, index: int, indent: int, depth: int
    ) -> tuple[Any, int]:
        if value_text:
            return _parse_inline(value_text, self.budget, depth + 1), index
        if index < len(self.lines) and self.lines[index].indent > indent:
            nested_indent = self.lines[index].indent
            return self._parse_node(index, nested_indent, depth + 1)
        return None, index

    def _add_mapping_entry(
        self,
        result: dict[str, Any],
        content: str,
        index: int,
        indent: int,
        depth: int,
    ) -> int:
        key_text, value_text = _split_mapping_entry(content)
        key = _parse_scalar(key_text, key=True)
        if key in result:
            raise _InvalidFrontmatter("duplicate mapping key")
        self.budget.consume()
        value, index = self._mapping_value(
            value_text, index, indent, depth
        )
        result[key] = value
        return index

    def _parse_mapping(
        self, index: int, indent: int, depth: int
    ) -> tuple[dict[str, Any], int]:
        result: dict[str, Any] = {}
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < indent:
                break
            if line.indent > indent or _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid mapping indentation")
            index += 1
            index = self._add_mapping_entry(
                result, line.content, index, indent, depth
            )
        return result, index

    def _parse_sequence_mapping(
        self, content: str, index: int, sequence_indent: int, depth: int
    ) -> tuple[dict[str, Any], int]:
        mapping_indent = sequence_indent + 2
        result: dict[str, Any] = {}
        index = self._add_mapping_entry(
            result, content, index, mapping_indent, depth
        )
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < mapping_indent:
                break
            if line.indent > mapping_indent or _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid sequence mapping indentation")
            index += 1
            index = self._add_mapping_entry(
                result, line.content, index, mapping_indent, depth
            )
        return result, index

    def _parse_sequence(
        self, index: int, indent: int, depth: int
    ) -> tuple[list[Any], int]:
        result = []
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < indent:
                break
            if line.indent > indent or not _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid sequence indentation")
            item_text = line.content[1:].strip()
            index += 1
            self.budget.consume()
            if not item_text:
                if index < len(self.lines) and self.lines[index].indent > indent:
                    value, index = self._parse_node(
                        index, self.lines[index].indent, depth + 1
                    )
                else:
                    value = None
            elif _looks_like_mapping(item_text):
                if depth + 1 > MAX_DEPTH:
                    raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
                value, index = self._parse_sequence_mapping(
                    item_text, index, indent, depth + 1
                )
            else:
                value = _parse_inline(item_text, self.budget, depth + 1)
                if index < len(self.lines) and self.lines[index].indent > indent:
                    raise _InvalidFrontmatter("scalar sequence item has nested content")
            result.append(value)
        return result, index


def parse_frontmatter_document(text: str) -> FrontmatterParseResult:
    """Parse a leading restricted-YAML frontmatter block without side effects."""
    if not isinstance(text, str):
        raise TypeError("text must be a string")

    located = _find_boundary(text)
    if located is None:
        return FrontmatterParseResult({}, (), None, text)
    opener_end, closing_start, boundary_end = located
    if closing_start < 0:
        return FrontmatterParseResult(
            {}, (FRONTMATTER_UNCLOSED,), None, text
        )

    boundary = (0, boundary_end)
    body = text[boundary_end:]
    try:
        values = _BlockParser(text[opener_end:closing_start]).parse()
    except (OverflowError, UnicodeError, _InvalidFrontmatter):
        return FrontmatterParseResult(
            {}, (FRONTMATTER_INVALID,), boundary, body
        )
    return FrontmatterParseResult(values, (), boundary, body)


__all__ = [
    "FRONTMATTER_INVALID",
    "FRONTMATTER_UNCLOSED",
    "FrontmatterParseResult",
    "parse_frontmatter_document",
]
