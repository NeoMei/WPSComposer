"""Bounded restricted-LaTeX parser and WPS linear-math serializer.

This module deliberately implements a small common subset instead of invoking a
TeX engine. Parsing and conversion are one operation: callers only receive a
descriptor after the complete source has been accepted and serialized.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass
from typing import FrozenSet, List, Optional, Sequence, Tuple


FORMULA_TOO_LONG = "FORMULA_TOO_LONG"
FORMULA_NESTING_TOO_DEEP = "FORMULA_NESTING_TOO_DEEP"
FORMULA_FORBIDDEN_PRIMITIVE = "FORMULA_FORBIDDEN_PRIMITIVE"
FORMULA_UNKNOWN_COMMAND = "FORMULA_UNKNOWN_COMMAND"
FORMULA_MALFORMED = "FORMULA_MALFORMED"
FORMULA_TOO_COMPLEX = "FORMULA_TOO_COMPLEX"

_SYNTAX = "wps-linear-v1"
_MAX_CODE_POINTS = 10_000
_MAX_BRACE_DEPTH = 64
_MAX_TOKENS = 4_096
_MAX_MATRIX_ROWS = 64
_MAX_MATRIX_COLUMNS = 64

_EXTERNAL_REFERENCE_RE = re.compile(
    r"(?:https?://|ftp://|file://|www\.)", re.IGNORECASE
)

_FORBIDDEN_COMMANDS: FrozenSet[str] = frozenset({
    "input", "include", "write", "write18", "openout", "openin", "immediate",
    "csname", "catcode", "escapechar", "endcsname", "noexpand", "expandafter",
    "def", "gdef", "edef", "xdef", "let", "futurelet", "global", "newcommand",
    "renewcommand", "newenvironment", "renewenvironment", "DeclareMathOperator",
    "usepackage", "documentclass", "RequirePackage", "LoadClass", "href", "url",
    "path", "includegraphics", "graphicspath", "label", "ref", "pageref", "cite",
    "bibliography", "bibliographystyle", "batchmode", "nonstopmode", "scrollmode",
    "errorstopmode", "shell", "system", "exec", "open", "file", "read",
    "directlua", "luadirect", "writefile",
})

_SYMBOL_COMMANDS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
    "epsilon": "ϵ", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π",
    "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ", "sigma": "σ",
    "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "ϕ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    "times": "×", "div": "÷", "pm": "±", "mp": "∓", "cdot": "⋅",
    "ast": "∗", "star": "⋆", "circ": "∘", "bullet": "∙",
    "oplus": "⊕", "ominus": "⊖", "otimes": "⊗", "oslash": "⊘",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠",
    "ne": "≠", "approx": "≈", "sim": "∼", "simeq": "≃",
    "equiv": "≡", "cong": "≅", "propto": "∝", "in": "∈",
    "notin": "∉", "subset": "⊂", "supset": "⊃", "subseteq": "⊆",
    "supseteq": "⊇", "cup": "∪", "cap": "∩", "setminus": "∖",
    "emptyset": "∅", "forall": "∀", "exists": "∃", "nexists": "∄",
    "neg": "¬", "land": "∧", "lor": "∨", "wedge": "∧", "vee": "∨",
    "to": "→", "gets": "←", "rightarrow": "→", "leftarrow": "←",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "leftrightarrow": "↔",
    "Leftrightarrow": "⇔", "mapsto": "↦", "iff": "⇔", "infty": "∞",
    "nabla": "∇", "partial": "∂", "prime": "′", "hbar": "ℏ",
    "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "ldots": "…",
    "cdots": "⋯", "vdots": "⋮", "ddots": "⋱", "dots": "…",
}

_LARGE_OPERATORS = {
    "sum": "∑", "prod": "∏", "int": "∫", "oint": "∮",
    "iint": "∬", "iiint": "∭", "iiiint": "⨌",
}
_NAMED_FUNCTIONS: FrozenSet[str] = frozenset({
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "exp", "lim", "sup", "inf", "max",
    "min", "arg", "dim", "det", "ker", "mod", "gcd", "lcm", "Pr",
})
_GROUP_WRAPPERS: FrozenSet[str] = frozenset({
    "text", "mbox", "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "mathcal",
    "mathbb", "mathfrak", "mathscr", "boldsymbol", "bm",
})
_SPACING_COMMANDS: FrozenSet[str] = frozenset({
    ",", ":", ";", "!", "quad", "qquad", "space", "thinspace", "medspace",
    "thickspace", "enspace",
})
_STYLE_COMMANDS: FrozenSet[str] = frozenset({
    "displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle",
})
_MATRIX_DELIMITERS = {
    "matrix": ("", ""), "smallmatrix": ("", ""), "pmatrix": ("(", ")"),
    "bmatrix": ("[", "]"), "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"),
    "Vmatrix": ("‖", "‖"), "cases": ("{", ""),
}
_DELIMITER_COMMANDS = {
    "lbrace": "{", "rbrace": "}", "langle": "⟨", "rangle": "⟩",
    "vert": "|", "Vert": "‖", "lfloor": "⌊", "rfloor": "⌋",
    "lceil": "⌈", "rceil": "⌉", "{": "{", "}": "}", "|": "|",
}
_LITERAL_COMMANDS = {
    "%": "%", "#": "#", "$": "$", "_": "_", "&": "&",
    "{": "{", "}": "}", "|": "|", " ": " ",
}


class NativeMathConversionError(ValueError):
    """Stable preflight failure raised before a native-math descriptor exists."""

    def __init__(self, code: str, detail: str = "") -> None:
        self.code = code
        self.detail = detail
        super().__init__(code if not detail else f"{code}: {detail}")


@dataclass(frozen=True)
class NativeMathDescriptor:
    """Trusted, immutable WPS linear-math content."""

    syntax: str
    linear_text: str
    source_hash: str

    def __post_init__(self) -> None:
        if self.syntax != _SYNTAX:
            raise ValueError("unsupported native-math syntax")
        if not isinstance(self.linear_text, str):
            raise TypeError("linear_text must be str")
        if not re.fullmatch(r"[0-9a-f]{64}", self.source_hash):
            raise ValueError("source_hash must be a lowercase SHA-256 digest")

    def __repr__(self) -> str:
        return (
            "NativeMathDescriptor(syntax='wps-linear-v1', "
            f"linear_text=<{len(self.linear_text)} chars>, source_hash='{self.source_hash}')"
        )


@dataclass(frozen=True)
class _Token:
    kind: str
    value: str


def _raise(code: str, detail: str = "") -> None:
    raise NativeMathConversionError(code, detail)


def _check_braces(source: str) -> None:
    depth = 0
    index = 0
    while index < len(source):
        char = source[index]
        if char == "\\":
            index += 2
            continue
        if char == "{":
            depth += 1
            if depth > _MAX_BRACE_DEPTH:
                _raise(FORMULA_NESTING_TOO_DEEP)
        elif char == "}":
            depth -= 1
            if depth < 0:
                _raise(FORMULA_MALFORMED, "unbalanced group")
        index += 1
    if depth:
        _raise(FORMULA_MALFORMED, "unbalanced group")


def _tokenize(source: str) -> Tuple[_Token, ...]:
    tokens: List[_Token] = []
    index = 0
    punctuation = frozenset("{}[]^_&+-=*/<>(),|!:;@")
    while index < len(source):
        char = source[index]
        if char.isspace():
            index += 1
            continue
        if char == "\\":
            if index + 1 >= len(source):
                _raise(FORMULA_MALFORMED, "trailing escape")
            following = source[index + 1]
            if following == "\\":
                tokens.append(_Token("ROW", "\\\\"))
                index += 2
                continue
            if following.isalpha():
                end = index + 2
                while end < len(source) and source[end].isalpha():
                    end += 1
                tokens.append(_Token("COMMAND", source[index + 1:end]))
                index = end
                continue
            tokens.append(_Token("COMMAND", following))
            index += 2
            continue
        kind = {
            "{": "LBRACE", "}": "RBRACE", "[": "LBRACKET", "]": "RBRACKET",
            "^": "SUP", "_": "SUB", "&": "AMP",
        }.get(char)
        if kind is not None:
            tokens.append(_Token(kind, char))
            index += 1
            continue
        if char in punctuation:
            tokens.append(_Token("TEXT", char))
            index += 1
            continue
        end = index + 1
        while (
            end < len(source)
            and not source[end].isspace()
            and source[end] != "\\"
            and source[end] not in punctuation
        ):
            end += 1
        tokens.append(_Token("TEXT", source[index:end]))
        index = end
        if len(tokens) > _MAX_TOKENS:
            _raise(FORMULA_TOO_COMPLEX, "token limit exceeded")
    if len(tokens) > _MAX_TOKENS:
        _raise(FORMULA_TOO_COMPLEX, "token limit exceeded")
    return tuple(tokens)


class _Parser:
    def __init__(self, tokens: Sequence[_Token]) -> None:
        self._tokens = tokens
        self._index = 0
        self._depth = 0

    def parse(self) -> str:
        rendered = self._parse_sequence()
        if self._index != len(self._tokens):
            _raise(FORMULA_MALFORMED, "unexpected trailing token")
        return rendered.strip()

    def _peek(self, offset: int = 0) -> Optional[_Token]:
        index = self._index + offset
        return None if index >= len(self._tokens) else self._tokens[index]

    def _take(self, kind: Optional[str] = None) -> _Token:
        token = self._peek()
        if token is None:
            _raise(FORMULA_MALFORMED, "unexpected end of formula")
        if kind is not None and token.kind != kind:
            _raise(FORMULA_MALFORMED, f"expected {kind}")
        self._index += 1
        return token

    def _at_command(self, name: str) -> bool:
        token = self._peek()
        return token is not None and token.kind == "COMMAND" and token.value == name

    def _parse_sequence(
        self,
        stop_kinds: FrozenSet[str] = frozenset(),
        stop_on_right: bool = False,
        stop_on_end: bool = False,
    ) -> str:
        parts: List[str] = []
        while self._peek() is not None:
            token = self._peek()
            assert token is not None
            if token.kind in stop_kinds:
                break
            if stop_on_right and self._at_command("right"):
                break
            if stop_on_end and self._at_command("end"):
                break
            if token.kind in {"RBRACE", "RBRACKET", "AMP", "ROW"}:
                _raise(FORMULA_MALFORMED, f"unexpected {token.value}")
            atom = self._parse_scripts(self._parse_atom())
            parts.append(atom)
        return "".join(parts)

    def _parse_atom(self) -> str:
        token = self._take()
        if token.kind == "TEXT":
            if token.value in {"#", "$", "%", "@"}:
                _raise(FORMULA_MALFORMED, f"unescaped special character {token.value}")
            return token.value
        if token.kind == "LBRACE":
            return self._parse_group_after_open()
        if token.kind == "COMMAND":
            return self._parse_command(token.value)
        if token.kind in {"SUP", "SUB"}:
            _raise(FORMULA_MALFORMED, "script has no base")
        _raise(FORMULA_MALFORMED, f"unexpected {token.value}")
        return ""  # pragma: no cover

    def _parse_group_after_open(self) -> str:
        self._enter_nested()
        value = self._parse_sequence(stop_kinds=frozenset({"RBRACE"}))
        self._take("RBRACE")
        self._leave_nested()
        return value

    def _enter_nested(self) -> None:
        self._depth += 1
        if self._depth > _MAX_BRACE_DEPTH:
            _raise(FORMULA_NESTING_TOO_DEEP)

    def _leave_nested(self) -> None:
        self._depth -= 1

    def _parse_required_group(self) -> str:
        self._take("LBRACE")
        return self._parse_group_after_open()

    def _parse_optional_group(self) -> Optional[str]:
        token = self._peek()
        if token is None or token.kind != "LBRACKET":
            return None
        self._take("LBRACKET")
        value = self._parse_sequence(stop_kinds=frozenset({"RBRACKET"}))
        self._take("RBRACKET")
        if not value:
            _raise(FORMULA_MALFORMED, "empty optional group")
        return value

    def _parse_script_value(self) -> str:
        token = self._peek()
        if token is None:
            _raise(FORMULA_MALFORMED, "missing script value")
        if token.kind == "LBRACE":
            self._take("LBRACE")
            value = self._parse_group_after_open()
        else:
            value = self._parse_atom()
        if not value:
            _raise(FORMULA_MALFORMED, "empty script")
        return value

    def _parse_scripts(self, base: str) -> str:
        subscript: Optional[str] = None
        superscript: Optional[str] = None
        is_large_operator = bool(base and base[0] in _LARGE_OPERATORS.values())
        if is_large_operator and (
            self._at_command("limits") or self._at_command("nolimits")
        ):
            self._take("COMMAND")
        while self._peek() is not None and self._peek().kind in {"SUB", "SUP"}:
            kind = self._take().kind
            value = self._parse_script_value()
            if kind == "SUB":
                if subscript is not None:
                    _raise(FORMULA_MALFORMED, "duplicate subscript")
                subscript = value
            else:
                if superscript is not None:
                    _raise(FORMULA_MALFORMED, "duplicate superscript")
                superscript = value
        if subscript is not None:
            base += "_(" + subscript + ")" if len(subscript) > 1 else "_" + subscript
        if superscript is not None:
            base += "^(" + superscript + ")" if len(superscript) > 1 else "^" + superscript
        if is_large_operator:
            base += " "
        return base

    def _parse_command(self, command: str) -> str:
        if command in _FORBIDDEN_COMMANDS:
            _raise(FORMULA_FORBIDDEN_PRIMITIVE, "\\" + command)
        if command in _SYMBOL_COMMANDS:
            return _SYMBOL_COMMANDS[command]
        if command in _LARGE_OPERATORS:
            return _LARGE_OPERATORS[command]
        if command in _NAMED_FUNCTIONS:
            return command + " "
        if command in _SPACING_COMMANDS:
            return " "
        if command in _STYLE_COMMANDS:
            return ""
        if command in _LITERAL_COMMANDS:
            return _LITERAL_COMMANDS[command]
        if command in _GROUP_WRAPPERS:
            return self._parse_required_group()
        if command in {"frac", "dfrac", "tfrac"}:
            numerator = self._parse_required_group()
            denominator = self._parse_required_group()
            if not numerator or not denominator:
                _raise(FORMULA_MALFORMED, "empty fraction operand")
            return f"({numerator})/({denominator})"
        if command == "sqrt":
            degree = self._parse_optional_group()
            radicand = self._parse_required_group()
            if not radicand:
                _raise(FORMULA_MALFORMED, "empty root")
            return f"√({radicand})" if degree is None else f"√({degree}&{radicand})"
        if command == "left":
            return self._parse_scalable_delimiters()
        if command == "begin":
            return self._parse_environment()
        if command in {"right", "end"}:
            _raise(FORMULA_MALFORMED, "unmatched \\" + command)
        _raise(FORMULA_UNKNOWN_COMMAND, "\\" + command)
        return ""  # pragma: no cover

    def _parse_delimiter(self) -> str:
        token = self._take()
        if token.kind == "TEXT" and token.value in {"(", ")", "[", "]", "|", "."}:
            return "" if token.value == "." else token.value
        if token.kind == "LBRACKET":
            return "["
        if token.kind == "RBRACKET":
            return "]"
        if token.kind == "COMMAND" and token.value in _DELIMITER_COMMANDS:
            return _DELIMITER_COMMANDS[token.value]
        _raise(FORMULA_MALFORMED, "invalid scalable delimiter")
        return ""  # pragma: no cover

    def _parse_scalable_delimiters(self) -> str:
        self._enter_nested()
        try:
            left = self._parse_delimiter()
            body = self._parse_sequence(stop_on_right=True)
            if not self._at_command("right"):
                _raise(FORMULA_MALFORMED, "missing \\right")
            self._take("COMMAND")
            right = self._parse_delimiter()
            return left + body + right
        finally:
            self._leave_nested()

    def _read_environment_name(self) -> str:
        self._take("LBRACE")
        token = self._take("TEXT")
        self._take("RBRACE")
        if not token.value.isalpha():
            _raise(FORMULA_MALFORMED, "invalid environment name")
        return token.value

    def _parse_environment(self) -> str:
        self._enter_nested()
        try:
            name = self._read_environment_name()
            if name not in _MATRIX_DELIMITERS:
                _raise(FORMULA_UNKNOWN_COMMAND, "environment " + name)
            rows: List[List[str]] = []
            row: List[str] = []
            while True:
                if self._at_command("end"):
                    if row:
                        rows.append(row)
                    break
                cell = self._parse_sequence(
                    stop_kinds=frozenset({"AMP", "ROW"}), stop_on_end=True
                )
                if not cell:
                    _raise(FORMULA_MALFORMED, "empty matrix cell")
                row.append(cell)
                if len(row) > _MAX_MATRIX_COLUMNS:
                    _raise(FORMULA_TOO_COMPLEX, "matrix column limit exceeded")
                token = self._peek()
                if token is None:
                    _raise(FORMULA_MALFORMED, "unterminated environment")
                if token.kind == "AMP":
                    self._take("AMP")
                    continue
                if token.kind == "ROW":
                    self._take("ROW")
                    rows.append(row)
                    if len(rows) > _MAX_MATRIX_ROWS:
                        _raise(FORMULA_TOO_COMPLEX, "matrix row limit exceeded")
                    row = []
                    if self._at_command("end"):
                        break
                    continue
                if self._at_command("end"):
                    rows.append(row)
                    break
                _raise(FORMULA_MALFORMED, "invalid matrix separator")
            self._take("COMMAND")
            closing_name = self._read_environment_name()
            if closing_name != name:
                _raise(FORMULA_MALFORMED, "mismatched environment")
            if not rows or len(rows) > _MAX_MATRIX_ROWS:
                _raise(FORMULA_TOO_COMPLEX, "matrix row limit exceeded")
            width = len(rows[0])
            if width == 0 or any(len(candidate) != width for candidate in rows):
                _raise(FORMULA_MALFORMED, "ragged matrix")
            if name == "cases" and width > 2:
                _raise(FORMULA_TOO_COMPLEX, "cases column limit exceeded")
            matrix = "■(" + "@".join("&".join(candidate) for candidate in rows) + ")"
            left, right = _MATRIX_DELIMITERS[name]
            return left + matrix + right
        finally:
            self._leave_nested()


def convert_restricted_latex(source: str) -> NativeMathDescriptor:
    """Validate and convert restricted LaTeX to trusted WPS linear math."""
    if not isinstance(source, str):
        raise TypeError("formula source must be str")
    if len(source) > _MAX_CODE_POINTS:
        _raise(FORMULA_TOO_LONG)
    normalized = unicodedata.normalize("NFC", source)
    for char in normalized:
        category = unicodedata.category(char)
        if category in {"Cc", "Cf", "Cs"} and char not in {"\t", "\n", "\r"}:
            _raise(FORMULA_MALFORMED, "control character")
    if _EXTERNAL_REFERENCE_RE.search(normalized):
        _raise(FORMULA_FORBIDDEN_PRIMITIVE, "external reference")
    _check_braces(normalized)
    linear_text = _Parser(_tokenize(normalized)).parse()
    linear_text = " ".join(linear_text.split())
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    return NativeMathDescriptor(_SYNTAX, linear_text, digest)


__all__ = [
    "FORMULA_FORBIDDEN_PRIMITIVE", "FORMULA_MALFORMED",
    "FORMULA_NESTING_TOO_DEEP", "FORMULA_TOO_COMPLEX", "FORMULA_TOO_LONG",
    "FORMULA_UNKNOWN_COMMAND", "NativeMathConversionError", "NativeMathDescriptor",
    "convert_restricted_latex",
]
