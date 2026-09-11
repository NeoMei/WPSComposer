"""Bounded parser for the long-form frontmatter YAML data subset.

Parsing has three explicit phases: compose a syntax-only node tree, validate
that complete tree, and only then construct JSON-compatible Python values.
No YAML object constructor, tag resolver, or arbitrary loader is used.
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
MAX_NUMERIC_DIGITS = 1_024

_DECIMAL_DIGITS = r"[0-9](?:_?[0-9])*"
_OCTAL_DIGITS = r"[0-7](?:_?[0-7])*"
_HEX_DIGITS = r"[0-9a-fA-F](?:_?[0-9a-fA-F])*"
_INT_RE = re.compile(
    rf"^[+-]?(?:{_DECIMAL_DIGITS}|0[oO]{_OCTAL_DIGITS}|0[xX]{_HEX_DIGITS})$"
)
_FLOAT_RE = re.compile(
    rf"^[+-]?(?:(?:{_DECIMAL_DIGITS}\.(?:{_DECIMAL_DIGITS})?"
    rf"|\.{_DECIMAL_DIGITS})(?:[eE][+-]?{_DECIMAL_DIGITS})?"
    rf"|{_DECIMAL_DIGITS}[eE][+-]?{_DECIMAL_DIGITS})$"
)
_NUMERIC_CANDIDATE_RE = re.compile(
    r"^[+-]?(?:0[xXoO][0-9a-fA-F_]*"
    r"|[0-9_]+(?:\.[0-9_]*)?(?:[eE][+-]?[0-9_]*)?"
    r"|\.[0-9_]*(?:[eE][+-]?[0-9_]*)?)$"
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


# Syntax-only node AST. Scalar source is intentionally unresolved here.


@dataclass(frozen=True)
class _ScalarNode:
    source: str
    style: str = "plain"


@dataclass(frozen=True)
class _SequenceNode:
    items: tuple[Any, ...]


@dataclass(frozen=True)
class _MappingEntry:
    key: Any
    value: Any


@dataclass(frozen=True)
class _MappingNode:
    entries: tuple[_MappingEntry, ...]


@dataclass(frozen=True)
class _TaggedNode:
    tag: str
    value: Any


@dataclass(frozen=True)
class _AnchoredNode:
    anchor: str
    value: Any


@dataclass(frozen=True)
class _AliasNode:
    alias: str


@dataclass(frozen=True)
class _CompoundKeyNode:
    source: str


@dataclass(frozen=True)
class _SourceLine:
    indent: int
    content: str


class _NodeBudget:
    def __init__(self) -> None:
        self.count = 0

    def consume(self) -> None:
        self.count += 1
        if self.count > MAX_NODES:
            raise _InvalidFrontmatter("frontmatter exceeds the node limit")


def _opening_boundary_end(text: str) -> Optional[int]:
    if text.startswith("---\r\n"):
        return 5
    if text.startswith(("---\n", "---\r")):
        return 4
    if text == "---":
        return 3
    return None


def _utf8_codepoint_length(character: str) -> int:
    codepoint = ord(character)
    if codepoint <= 0x7F:
        return 1
    if codepoint <= 0x7FF:
        return 2
    if codepoint <= 0xFFFF:
        return 3
    return 4


def _find_boundary(text: str) -> Optional[tuple[int, int, int]]:
    """Locate the delimiter without scanning beyond the UTF-8 byte budget."""
    opener_end = _opening_boundary_end(text)
    if opener_end is None:
        return None

    index = opener_end
    completed_bytes = 0
    line_start = index
    line_bytes = 0
    while index < len(text):
        character = text[index]
        if character in {"\n", "\r"}:
            line_is_boundary = (
                index - line_start == 3
                and text.startswith("---", line_start)
                and completed_bytes + line_bytes <= MAX_BOUNDARY_BYTES
            )
            if character == "\r" and index + 1 < len(text) and text[index + 1] == "\n":
                newline_end = index + 2
                newline_bytes = 2
            else:
                newline_end = index + 1
                newline_bytes = 1
            if line_is_boundary:
                return (opener_end, line_start, newline_end)
            if completed_bytes + line_bytes + newline_bytes > MAX_BOUNDARY_BYTES:
                return (opener_end, -1, -1)
            completed_bytes += line_bytes + newline_bytes
            index = newline_end
            line_start = index
            line_bytes = 0
            continue

        character_bytes = _utf8_codepoint_length(character)
        if completed_bytes + line_bytes + character_bytes > MAX_BOUNDARY_BYTES:
            return (opener_end, -1, -1)
        line_bytes += character_bytes
        index += 1

    if (
        index - line_start == 3
        and text.startswith("---", line_start)
        and completed_bytes + line_bytes <= MAX_BOUNDARY_BYTES
    ):
        return (opener_end, line_start, index)
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
    for raw_line in source.splitlines():
        indent = len(raw_line) - len(raw_line.lstrip(" "))
        content = _strip_comment(raw_line[indent:]).rstrip()
        if content:
            result.append(_SourceLine(indent, content))
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


def _scan_quoted_end(value: str, start: int = 0) -> int:
    quote = value[start]
    index = start + 1
    while index < len(value):
        character = value[index]
        if character == quote:
            if quote == "'" and index + 1 < len(value) and value[index + 1] == "'":
                index += 2
                continue
            return index + 1
        if quote == '"' and character == "\\":
            index += 2
        else:
            index += 1
    raise _InvalidFrontmatter("unterminated quoted scalar")


def _decode_quoted(value: str) -> str:
    quote = value[0]
    index = 1
    output = []
    while index < len(value):
        character = value[index]
        if character == quote:
            if quote == "'" and index + 1 < len(value) and value[index + 1] == "'":
                output.append("'")
                index += 2
                continue
            if value[index + 1 :].strip():
                raise _InvalidFrontmatter("unexpected text after quoted scalar")
            return _validate_string("".join(output))
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
        width = {"x": 2, "u": 4, "U": 8}.get(escape)
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


def _split_node_property(source: str) -> tuple[str, str]:
    index = 1
    while index < len(source) and not source[index].isspace():
        index += 1
    return source[1:index], source[index:].strip()


def _new_scalar(source: str, budget: _NodeBudget, style: str = "plain") -> _ScalarNode:
    budget.consume()
    return _ScalarNode(source, style)


def _compose_scalar_node(source: str, budget: _NodeBudget) -> Any:
    source = source.strip()
    if not source:
        return _new_scalar("", budget, "empty")
    if source[0] in {'"', "'"}:
        end = _scan_quoted_end(source)
        if source[end:].strip():
            raise _InvalidFrontmatter("unexpected text after quoted scalar")
        style = "double" if source[0] == '"' else "single"
        return _new_scalar(source, budget, style)
    if source[0] == "!":
        tag, remainder = _split_node_property(source)
        budget.consume()
        return _TaggedNode(tag, _compose_inline_node(remainder, budget, 1))
    if source[0] == "&":
        anchor, remainder = _split_node_property(source)
        budget.consume()
        return _AnchoredNode(anchor, _compose_inline_node(remainder, budget, 1))
    if source[0] == "*":
        budget.consume()
        return _AliasNode(source[1:].strip())
    if source.startswith("?"):
        budget.consume()
        return _CompoundKeyNode(source)
    if source[0] in "|>@`" or re.search(r"(?:^|\s)[!&*](?:\S|$)", source):
        raise _InvalidFrontmatter("unsupported YAML node feature")
    if any(character in source for character in "[]{}"):
        raise _InvalidFrontmatter("unexpected flow collection marker")
    return _new_scalar(source, budget)


class _FlowNodeComposer:
    def __init__(self, source: str, budget: _NodeBudget) -> None:
        self.source = source
        self.budget = budget
        self.index = 0

    def compose(self, depth: int) -> Any:
        node = self._compose_value(depth)
        self._skip_space()
        if self.index != len(self.source):
            raise _InvalidFrontmatter("unexpected text after flow value")
        return node

    def _skip_space(self) -> None:
        while self.index < len(self.source) and self.source[self.index].isspace():
            self.index += 1

    def _compose_value(self, depth: int) -> Any:
        self._skip_space()
        if self.index >= len(self.source):
            return _new_scalar("", self.budget, "empty")
        character = self.source[self.index]
        if character == "[":
            return self._compose_sequence(depth)
        if character == "{":
            return self._compose_mapping(depth)
        if character in {'"', "'"}:
            start = self.index
            self.index = _scan_quoted_end(self.source, start)
            style = "double" if character == '"' else "single"
            return _new_scalar(self.source[start:self.index], self.budget, style)
        start = self.index
        while self.index < len(self.source) and self.source[self.index] not in ",]}":
            self.index += 1
        return _compose_scalar_node(self.source[start:self.index], self.budget)

    def _compose_sequence(self, depth: int) -> _SequenceNode:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        self.budget.consume()
        self.index += 1
        items = []
        self._skip_space()
        if self.index < len(self.source) and self.source[self.index] == "]":
            self.index += 1
            return _SequenceNode(())
        while True:
            items.append(self._compose_value(depth + 1))
            self._skip_space()
            if self.index >= len(self.source):
                raise _InvalidFrontmatter("unterminated flow list")
            marker = self.source[self.index]
            self.index += 1
            if marker == "]":
                return _SequenceNode(tuple(items))
            if marker != ",":
                raise _InvalidFrontmatter("expected a flow list separator")

    def _compose_mapping_key(self, depth: int) -> Any:
        self._skip_space()
        if self.index >= len(self.source):
            raise _InvalidFrontmatter("flow mapping key is missing")
        if self.source[self.index] in "[{":
            key = self._compose_value(depth)
        elif self.source[self.index] in {'"', "'"}:
            start = self.index
            self.index = _scan_quoted_end(self.source, start)
            style = "double" if self.source[start] == '"' else "single"
            key = _new_scalar(self.source[start:self.index], self.budget, style)
        else:
            start = self.index
            while self.index < len(self.source) and self.source[self.index] != ":":
                if self.source[self.index] in "{},[]":
                    raise _InvalidFrontmatter("invalid flow mapping key")
                self.index += 1
            key = _compose_scalar_node(self.source[start:self.index], self.budget)
        self._skip_space()
        if self.index >= len(self.source) or self.source[self.index] != ":":
            raise _InvalidFrontmatter("flow mapping separator is missing")
        self.index += 1
        return key

    def _compose_mapping(self, depth: int) -> _MappingNode:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        self.budget.consume()
        self.index += 1
        entries = []
        self._skip_space()
        if self.index < len(self.source) and self.source[self.index] == "}":
            self.index += 1
            return _MappingNode(())
        while True:
            key = self._compose_mapping_key(depth + 1)
            value = self._compose_value(depth + 1)
            entries.append(_MappingEntry(key, value))
            self._skip_space()
            if self.index >= len(self.source):
                raise _InvalidFrontmatter("unterminated flow mapping")
            marker = self.source[self.index]
            self.index += 1
            if marker == "}":
                return _MappingNode(tuple(entries))
            if marker != ",":
                raise _InvalidFrontmatter("expected a flow mapping separator")


def _compose_inline_node(source: str, budget: _NodeBudget, depth: int) -> Any:
    source = source.strip()
    if source.startswith(("[", "{")):
        return _FlowNodeComposer(source, budget).compose(depth)
    return _compose_scalar_node(source, budget)


class _BlockNodeComposer:
    def __init__(self, source: str) -> None:
        self.lines = _prepare_lines(source)
        self.budget = _NodeBudget()

    def compose(self) -> Any:
        if not self.lines or self.lines[0].indent != 0:
            raise _InvalidFrontmatter("frontmatter root must be a mapping")
        if len(self.lines) == 1 and self.lines[0].content.startswith(("{", "[")):
            return _compose_inline_node(self.lines[0].content, self.budget, 1)
        if self.lines[0].content.startswith("?"):
            self.budget.consume()
            return _CompoundKeyNode(self.lines[0].content)
        node, index = self._compose_node(0, 0, 1)
        if index != len(self.lines):
            raise _InvalidFrontmatter("frontmatter contains trailing nodes")
        return node

    def _compose_node(self, index: int, indent: int, depth: int) -> tuple[Any, int]:
        if depth > MAX_DEPTH:
            raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
        if index >= len(self.lines) or self.lines[index].indent != indent:
            raise _InvalidFrontmatter("invalid indentation")
        if _is_sequence_entry(self.lines[index].content):
            return self._compose_sequence(index, indent, depth)
        return self._compose_mapping(index, indent, depth)

    def _compose_mapping_value(
        self, source: str, index: int, indent: int, depth: int
    ) -> tuple[Any, int]:
        if source:
            return _compose_inline_node(source, self.budget, depth + 1), index
        if index < len(self.lines) and self.lines[index].indent > indent:
            return self._compose_node(index, self.lines[index].indent, depth + 1)
        return _new_scalar("", self.budget, "empty"), index

    def _compose_mapping_entry(
        self,
        content: str,
        index: int,
        indent: int,
        depth: int,
    ) -> tuple[_MappingEntry, int]:
        key_source, value_source = _split_mapping_entry(content)
        key = _compose_scalar_node(key_source, self.budget)
        value, index = self._compose_mapping_value(
            value_source, index, indent, depth
        )
        return _MappingEntry(key, value), index

    def _compose_mapping(
        self, index: int, indent: int, depth: int
    ) -> tuple[_MappingNode, int]:
        self.budget.consume()
        entries = []
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < indent:
                break
            if line.indent > indent or _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid mapping indentation")
            if line.content.startswith("?"):
                self.budget.consume()
                entries.append(
                    _MappingEntry(
                        _CompoundKeyNode(line.content),
                        _new_scalar("", self.budget, "empty"),
                    )
                )
                index += 1
                continue
            index += 1
            entry, index = self._compose_mapping_entry(
                line.content, index, indent, depth
            )
            entries.append(entry)
        return _MappingNode(tuple(entries)), index

    def _compose_sequence_mapping(
        self, content: str, index: int, sequence_indent: int, depth: int
    ) -> tuple[_MappingNode, int]:
        self.budget.consume()
        mapping_indent = sequence_indent + 2
        entries = []
        entry, index = self._compose_mapping_entry(
            content, index, mapping_indent, depth
        )
        entries.append(entry)
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < mapping_indent:
                break
            if line.indent > mapping_indent or _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid sequence mapping indentation")
            index += 1
            entry, index = self._compose_mapping_entry(
                line.content, index, mapping_indent, depth
            )
            entries.append(entry)
        return _MappingNode(tuple(entries)), index

    def _compose_sequence(
        self, index: int, indent: int, depth: int
    ) -> tuple[_SequenceNode, int]:
        self.budget.consume()
        items = []
        while index < len(self.lines):
            line = self.lines[index]
            if line.indent < indent:
                break
            if line.indent > indent or not _is_sequence_entry(line.content):
                raise _InvalidFrontmatter("invalid sequence indentation")
            item_source = line.content[1:].strip()
            index += 1
            if not item_source:
                if index < len(self.lines) and self.lines[index].indent > indent:
                    item, index = self._compose_node(
                        index, self.lines[index].indent, depth + 1
                    )
                else:
                    item = _new_scalar("", self.budget, "empty")
            elif _looks_like_mapping(item_source):
                if depth + 1 > MAX_DEPTH:
                    raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
                item, index = self._compose_sequence_mapping(
                    item_source, index, indent, depth + 1
                )
            else:
                item = _compose_inline_node(item_source, self.budget, depth + 1)
                if index < len(self.lines) and self.lines[index].indent > indent:
                    raise _InvalidFrontmatter("scalar sequence item has nested content")
            items.append(item)
        return _SequenceNode(tuple(items)), index


def _compose_frontmatter_nodes(source: str) -> Any:
    """Compose syntax nodes without resolving any scalar to a Python value."""
    return _BlockNodeComposer(source).compose()


def _plain_scalar_kind(source: str) -> str:
    lowered = source.lower()
    if source == "~" or lowered == "null":
        return "null"
    if lowered in {"true", "false"}:
        return "bool"
    if _NONFINITE_RE.fullmatch(source):
        raise _InvalidFrontmatter("non-finite numbers are not JSON-compatible")

    numeric_candidate = bool(_NUMERIC_CANDIDATE_RE.fullmatch(source))
    if _INT_RE.fullmatch(source):
        kind = "int"
    elif _FLOAT_RE.fullmatch(source):
        kind = "float"
    elif numeric_candidate or source.lower().startswith(
        ("0x", "0o", "+0x", "-0x", "+0o", "-0o")
    ):
        raise _InvalidFrontmatter("malformed numeric scalar")
    else:
        return "string"

    digit_count = sum(character.isdigit() for character in source)
    if digit_count > MAX_NUMERIC_DIGITS:
        raise _InvalidFrontmatter("numeric scalar exceeds the digit limit")
    if kind == "float":
        try:
            number = float(source.replace("_", ""))
        except (OverflowError, ValueError) as error:
            raise _InvalidFrontmatter("malformed numeric scalar") from error
        if not math.isfinite(number):
            raise _InvalidFrontmatter("non-finite numbers are not JSON-compatible")
    return kind


def _validate_scalar_node(node: _ScalarNode) -> None:
    if node.style == "empty":
        return
    if node.style in {"single", "double"}:
        _decode_quoted(node.source)
        return
    if node.style != "plain":
        raise _InvalidFrontmatter("unknown scalar style")
    kind = _plain_scalar_kind(node.source)
    if kind == "string":
        _validate_string(node.source)


def _validated_mapping_key(node: Any) -> str:
    if not isinstance(node, _ScalarNode) or node.style == "empty":
        raise _InvalidFrontmatter("mapping keys must be scalar strings")
    if node.style in {"single", "double"}:
        key = _decode_quoted(node.source)
    elif node.style == "plain":
        if _plain_scalar_kind(node.source) != "string":
            raise _InvalidFrontmatter("mapping keys must be strings")
        key = _validate_string(node.source)
    else:
        raise _InvalidFrontmatter("unknown mapping key style")
    if key == "<<":
        raise _InvalidFrontmatter("merge keys are not supported")
    return key


def _validate_composed_node(root: Any) -> None:
    """Validate the complete AST before value construction is permitted."""
    if not isinstance(root, _MappingNode):
        raise _InvalidFrontmatter("frontmatter root must be a mapping")

    count = 0

    def visit(node: Any, depth: int) -> None:
        nonlocal count
        count += 1
        if count > MAX_NODES:
            raise _InvalidFrontmatter("frontmatter exceeds the node limit")
        if isinstance(node, (_TaggedNode, _AnchoredNode, _AliasNode)):
            raise _InvalidFrontmatter("tags, anchors, and aliases are unsupported")
        if isinstance(node, _CompoundKeyNode):
            raise _InvalidFrontmatter("compound mapping keys are unsupported")
        if isinstance(node, _ScalarNode):
            _validate_scalar_node(node)
            return
        if isinstance(node, _SequenceNode):
            if depth > MAX_DEPTH:
                raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
            for item in node.items:
                visit(item, depth + 1)
            return
        if isinstance(node, _MappingNode):
            if depth > MAX_DEPTH:
                raise _InvalidFrontmatter("frontmatter exceeds the depth limit")
            keys = set()
            for entry in node.entries:
                key = _validated_mapping_key(entry.key)
                if key in keys:
                    raise _InvalidFrontmatter("duplicate mapping key")
                keys.add(key)
                visit(entry.key, depth + 1)
                visit(entry.value, depth + 1)
            return
        raise _InvalidFrontmatter("unknown composed node type")

    visit(root, 1)


def _construct_scalar(node: _ScalarNode) -> Any:
    if node.style == "empty":
        return None
    if node.style in {"single", "double"}:
        return _decode_quoted(node.source)
    kind = _plain_scalar_kind(node.source)
    if kind == "null":
        return None
    if kind == "bool":
        return node.source.lower() == "true"
    if kind == "string":
        return _validate_string(node.source)

    compact = node.source.replace("_", "")
    try:
        if kind == "float":
            return float(compact)
        sign = -1 if compact.startswith("-") else 1
        unsigned = compact.lstrip("+-")
        if unsigned.lower().startswith("0o"):
            return sign * int(unsigned[2:], 8)
        if unsigned.lower().startswith("0x"):
            return sign * int(unsigned[2:], 16)
        return int(compact, 10)
    except (OverflowError, SystemError, ValueError) as error:
        raise _InvalidFrontmatter("malformed numeric scalar") from error


def _construct_node(node: Any) -> Any:
    if isinstance(node, _ScalarNode):
        return _construct_scalar(node)
    if isinstance(node, _SequenceNode):
        return [_construct_node(item) for item in node.items]
    if isinstance(node, _MappingNode):
        return {
            _validated_mapping_key(entry.key): _construct_node(entry.value)
            for entry in node.entries
        }
    raise _InvalidFrontmatter("unsafe node reached value construction")


def _construct_frontmatter_values(root: Any) -> dict[str, Any]:
    """Construct permitted data after `_validate_composed_node` succeeds."""
    value = _construct_node(root)
    if not isinstance(value, dict):
        raise _InvalidFrontmatter("frontmatter root must be a mapping")
    return value


def parse_frontmatter_document(text: str) -> FrontmatterParseResult:
    """Parse a leading restricted-YAML frontmatter block without side effects."""
    if not isinstance(text, str):
        raise TypeError("text must be a string")

    located = _find_boundary(text)
    if located is None:
        return FrontmatterParseResult({}, (), None, text)
    opener_end, closing_start, boundary_end = located
    if closing_start < 0:
        return FrontmatterParseResult({}, (FRONTMATTER_UNCLOSED,), None, text)

    boundary = (0, boundary_end)
    body = text[boundary_end:]
    try:
        root = _compose_frontmatter_nodes(text[opener_end:closing_start])
        _validate_composed_node(root)
        values = _construct_frontmatter_values(root)
    except (OverflowError, SystemError, UnicodeError, ValueError):
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
