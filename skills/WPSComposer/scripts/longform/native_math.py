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
_MAX_LINEAR_CODE_POINTS = 10_000

_EXTERNAL_REFERENCE_RE = re.compile(
    r"(?:https?://|ftp://|file://|www\.)", re.IGNORECASE
)

_FORBIDDEN_COMMANDS: FrozenSet[str] = frozenset({
    "input", "include", "write", "write18", "openout", "openin", "immediate",
    "csname", "catcode", "escapechar", "endcsname", "noexpand", "expandafter",
    "def", "gdef", "edef", "xdef", "let", "futurelet", "global", "newcommand",
    "renewcommand", "newenvironment", "renewenvironment", "DeclareMathOperator",
    "usepackage", "documentclass", "RequirePackage", "LoadClass", "href", "url",
    "path", "includegraphics", "graphicspath", "ref", "pageref", "cite",
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
    "ell": "ℓ", "wp": "℘", "Re": "ℜ", "Im": "ℑ", "ldots": "…",
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
_EQUATION_ARRAY_ENVIRONMENTS: FrozenSet[str] = frozenset({
    "align", "aligned", "alignedat", "flalign", "split", "gather", "gathered",
    "multline", "equation", "array",
})
_BINOMIAL_COMMANDS: FrozenSet[str] = frozenset({"binom", "tbinom", "dbinom"})
_ACCENT_COMMANDS = {
    "hat": "\u0302", "widehat": "\u0302", "tilde": "\u0303",
    "widetilde": "\u0303", "vec": "\u20d7", "bar": "\u0305",
    "dot": "\u0307", "ddot": "\u0308", "overline": "\u0305",
    "underline": "\u0332",
    "acute": "\u0301", "grave": "\u0300", "check": "\u030c",
    "breve": "\u0306", "overbrace": "⏞", "underbrace": "⏟",
}
_STACK_COMMANDS: FrozenSet[str] = frozenset({
    "overset", "underset", "stackrel", "buildrel",
})
_CHOICE_COMMANDS: FrozenSet[str] = frozenset({"atop", "choose", "brack", "brace"})
_DIMENSION_COMMANDS: FrozenSet[str] = frozenset({
    "hspace", "hskip", "vspace", "vskip", "kern", "mskip", "mkern",
})
_CONTENT_BOX_COMMANDS: FrozenSet[str] = frozenset({"raisebox", "lower"})
_SIMPLE_BOX_COMMANDS: FrozenSet[str] = frozenset({"box"})
_PHANTOM_COMMANDS: FrozenSet[str] = frozenset({
    "phantom", "vphantom", "hphantom",
})
_DELIMITER_SIZE_COMMANDS: FrozenSet[str] = frozenset({
    "big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr",
    "biggl", "biggr", "Biggl", "Biggr",
})
_METADATA_COMMANDS: FrozenSet[str] = frozenset({"nonumber", "tag", "label"})
_DELIMITER_COMMANDS = {
    "lbrace": "{", "rbrace": "}", "langle": "⟨", "rangle": "⟩",
    "vert": "|", "Vert": "‖", "lfloor": "⌊", "rfloor": "⌋",
    "lceil": "⌈", "rceil": "⌉", "{": "{", "}": "}", "|": "|",
}
_LITERAL_COMMANDS = {
    "%": "%", "#": "#", "$": "$", "_": "_", "&": "&",
    "{": "{", "}": "}", "|": "|", " ": " ",
}

_LEGACY_ENVIRONMENT_NAMES: FrozenSet[str] = frozenset({
    "matrix", "pmatrix", "bmatrix", "vmatrix", "Vmatrix", "Bmatrix",
    "smallmatrix", "cases", "align", "aligned", "alignedat", "gather",
    "multline", "equation", "array", "split", "flalign",
})
_LEGACY_STRUCTURAL_COMMANDS: FrozenSet[str] = frozenset({
    "limits", "nolimits", "left", "right", "begin", "end",
})
_FRACTION_COMMANDS: FrozenSet[str] = frozenset({"frac", "dfrac", "tfrac"})
_SPECIAL_FORM_COMMANDS: FrozenSet[str] = frozenset({
    "sqrt", "operatorname", "mathop", "bmod", "pmod", "pod", "genfrac", "not",
})

LEGACY_ALLOWED_COMMANDS: FrozenSet[str] = frozenset().union(
    _SYMBOL_COMMANDS,
    _LARGE_OPERATORS,
    _NAMED_FUNCTIONS,
    _GROUP_WRAPPERS,
    _SPACING_COMMANDS,
    _STYLE_COMMANDS,
    _LEGACY_ENVIRONMENT_NAMES,
    _LEGACY_STRUCTURAL_COMMANDS,
    _FRACTION_COMMANDS,
    _SPECIAL_FORM_COMMANDS,
    _BINOMIAL_COMMANDS,
    _ACCENT_COMMANDS,
    _STACK_COMMANDS,
    _CHOICE_COMMANDS,
    _DIMENSION_COMMANDS,
    _CONTENT_BOX_COMMANDS,
    _SIMPLE_BOX_COMMANDS,
    _PHANTOM_COMMANDS,
    _DELIMITER_SIZE_COMMANDS,
    _METADATA_COMMANDS,
)


def legacy_command_category(command: str) -> Optional[str]:
    """Return the explicit grammar category for a legacy-accepted command."""
    categories = (
        (_SYMBOL_COMMANDS, "symbol"),
        (_LARGE_OPERATORS, "large-operator"),
        (_NAMED_FUNCTIONS, "named-function"),
        (_GROUP_WRAPPERS, "group-wrapper"),
        (_SPACING_COMMANDS, "spacing"),
        (_STYLE_COMMANDS, "style"),
        (_LEGACY_ENVIRONMENT_NAMES, "environment"),
        (_LEGACY_STRUCTURAL_COMMANDS, "structural"),
        (_FRACTION_COMMANDS, "fraction"),
        (_SPECIAL_FORM_COMMANDS, "special-form"),
        (_BINOMIAL_COMMANDS, "binomial"),
        (_ACCENT_COMMANDS, "accent"),
        (_STACK_COMMANDS, "stack"),
        (_CHOICE_COMMANDS, "choice"),
        (_DIMENSION_COMMANDS, "dimension"),
        (_CONTENT_BOX_COMMANDS, "content-box"),
        (_SIMPLE_BOX_COMMANDS, "box"),
        (_PHANTOM_COMMANDS, "phantom"),
        (_DELIMITER_SIZE_COMMANDS, "delimiter-size"),
        (_METADATA_COMMANDS, "metadata"),
    )
    for commands, category in categories:
        if command in commands:
            return category
    return None


class NativeMathConversionError(ValueError):
    """Stable preflight failure raised before a native-math descriptor exists."""

    def __init__(self, code: str, detail: str = "") -> None:
        self.code = code
        self.detail = detail
        super().__init__(code if not detail else f"{code}: {detail}")


_DESCRIPTOR_FACTORY_TOKEN = object()


@dataclass(frozen=True, init=False)
class NativeMathDescriptor:
    """Trusted, immutable WPS linear-math content."""

    syntax: str
    linear_text: str
    source_hash: str

    def __init__(
        self,
        *,
        _factory_token: object,
        syntax: str,
        linear_text: str,
        source_hash: str,
    ) -> None:
        if _factory_token is not _DESCRIPTOR_FACTORY_TOKEN:
            raise TypeError("NativeMathDescriptor must be created by the converter")
        object.__setattr__(self, "syntax", syntax)
        object.__setattr__(self, "linear_text", linear_text)
        object.__setattr__(self, "source_hash", source_hash)
        self.__post_init__()

    def __post_init__(self) -> None:
        if self.syntax != _SYNTAX:
            raise ValueError("unsupported native-math syntax")
        if not isinstance(self.linear_text, str):
            raise TypeError("linear_text must be str")
        if not self.linear_text or len(self.linear_text) > _MAX_LINEAR_CODE_POINTS:
            raise ValueError("linear_text length is outside the trusted bound")
        if any(
            unicodedata.category(char) in {"Cc", "Cf", "Cs"}
            for char in self.linear_text
        ):
            raise ValueError("linear_text contains a control character")
        if not re.fullmatch(r"[0-9a-f]{64}", self.source_hash):
            raise ValueError("source_hash must be a lowercase SHA-256 digest")

    def __repr__(self) -> str:
        return (
            "NativeMathDescriptor(syntax='wps-linear-v1', "
            f"linear_text=<{len(self.linear_text)} chars>, source_hash='{self.source_hash}')"
        )


def _create_descriptor(linear_text: str, source_hash: str) -> NativeMathDescriptor:
    return NativeMathDescriptor(
        _factory_token=_DESCRIPTOR_FACTORY_TOKEN,
        syntax=_SYNTAX,
        linear_text=linear_text,
        source_hash=source_hash,
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
            tokens.append(_Token("SPACE", " "))
            index += 1
            while index < len(source) and source[index].isspace():
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

    def _skip_spaces(self) -> None:
        while self._peek() is not None and self._peek().kind == "SPACE":
            self._take("SPACE")

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
            if token.kind == "SPACE":
                self._take("SPACE")
                continue
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
        if not value.strip():
            _raise(FORMULA_MALFORMED, "empty group")
        return value

    def _enter_nested(self) -> None:
        self._depth += 1
        if self._depth > _MAX_BRACE_DEPTH:
            _raise(FORMULA_NESTING_TOO_DEEP)

    def _leave_nested(self) -> None:
        self._depth -= 1

    def _parse_required_group(self) -> str:
        self._skip_spaces()
        self._take("LBRACE")
        return self._parse_group_after_open()

    def _parse_required_group_allow_empty(self) -> str:
        self._skip_spaces()
        self._take("LBRACE")
        self._enter_nested()
        value = self._parse_sequence(stop_kinds=frozenset({"RBRACE"}))
        self._take("RBRACE")
        self._leave_nested()
        return value

    def _parse_dimension_group(self) -> str:
        value = self._parse_required_group()
        if re.fullmatch(
            r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:pt|em|ex|mu|cm|mm|in)",
            value,
        ) is None:
            _raise(FORMULA_MALFORMED, "invalid dimension")
        return value

    def _parse_required_text_group(self) -> str:
        """Parse a bounded text/mbox/operator-name argument with visible spaces."""
        self._skip_spaces()
        self._take("LBRACE")
        self._enter_nested()
        parts: List[str] = []
        while True:
            token = self._peek()
            if token is None:
                _raise(FORMULA_MALFORMED, "unterminated text group")
            if token.kind == "RBRACE":
                self._take("RBRACE")
                break
            token = self._take()
            if token.kind == "SPACE":
                if parts and parts[-1] != " ":
                    parts.append(" ")
                continue
            if token.kind == "TEXT":
                if token.value in {"#", "$", "%", "@", '"'}:
                    _raise(FORMULA_MALFORMED, "unsafe text character")
                parts.append(token.value)
                continue
            if token.kind == "COMMAND":
                if token.value in _LITERAL_COMMANDS:
                    parts.append(_LITERAL_COMMANDS[token.value])
                    continue
                if token.value in _SYMBOL_COMMANDS:
                    parts.append(_SYMBOL_COMMANDS[token.value])
                    continue
                _raise(FORMULA_UNKNOWN_COMMAND, "text command \\" + token.value)
            if token.kind == "LBRACE":
                _raise(FORMULA_MALFORMED, "nested text group")
            _raise(FORMULA_MALFORMED, "invalid text token")
        self._leave_nested()
        value = "".join(parts).strip()
        if not value:
            _raise(FORMULA_MALFORMED, "empty text group")
        return value

    def _parse_optional_group(self) -> Optional[str]:
        self._skip_spaces()
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
        self._skip_spaces()
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
        is_operator = bool(
            (base and base[0] in _LARGE_OPERATORS.values()) or base.endswith(" ")
        )
        if is_operator:
            base = base.rstrip()
        self._skip_spaces()
        if (
            not base.strip()
            and self._peek() is not None
            and self._peek().kind in {"SUB", "SUP"}
        ):
            _raise(FORMULA_MALFORMED, "script has no visible base")
        if is_operator and (
            self._at_command("limits") or self._at_command("nolimits")
        ):
            self._take("COMMAND")
            self._skip_spaces()
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
            self._skip_spaces()
        if subscript is not None:
            base += "_(" + subscript + ")" if len(subscript) > 1 else "_" + subscript
        if superscript is not None:
            base += "^(" + superscript + ")" if len(superscript) > 1 else "^" + superscript
        if is_operator:
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
        if command in {"text", "mbox"}:
            return '"' + self._parse_required_text_group() + '"'
        if command in _GROUP_WRAPPERS:
            return self._parse_required_group()
        if command in _FRACTION_COMMANDS:
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
        if command in _BINOMIAL_COMMANDS:
            upper = self._parse_required_group()
            lower = self._parse_required_group()
            return f"(({upper})¦({lower}))"
        if command in _ACCENT_COMMANDS:
            accented = self._parse_required_group()
            return f"({accented}){_ACCENT_COMMANDS[command]}"
        if command in _STACK_COMMANDS:
            annotation = self._parse_required_group()
            base = self._parse_required_group()
            if command == "underset":
                return f"({base})_({annotation})"
            return f"({base})^({annotation})"
        if command in _CHOICE_COMMANDS:
            upper = self._parse_required_group()
            lower = self._parse_required_group()
            stack = f"({upper})¦({lower})"
            if command == "choose":
                return "(" + stack + ")"
            if command == "brack":
                return "[" + stack + "]"
            if command == "brace":
                return "{" + stack + "}"
            return stack
        if command in _DIMENSION_COMMANDS:
            self._parse_dimension_group()
            return " " if command in {"hspace", "hskip", "kern", "mskip", "mkern"} else ""
        if command in _CONTENT_BOX_COMMANDS:
            self._parse_dimension_group()
            return self._parse_required_group()
        if command in _SIMPLE_BOX_COMMANDS:
            return self._parse_required_group()
        if command in _PHANTOM_COMMANDS:
            self._parse_required_group()
            return " "
        if command == "bmod":
            return " mod "
        if command == "pmod":
            return "(mod " + self._parse_required_group() + ")"
        if command == "pod":
            return "(" + self._parse_required_group() + ")"
        if command == "genfrac":
            return self._parse_generalized_fraction()
        if command in _DELIMITER_SIZE_COMMANDS:
            return self._parse_delimiter()
        if command == "nonumber":
            return ""
        if command in {"tag", "label"}:
            self._parse_required_group()
            return ""
        if command == "not":
            return self._parse_negated_atom()
        if command == "operatorname":
            self._skip_spaces()
            if (
                self._peek() is not None
                and self._peek().kind == "TEXT"
                and self._peek().value == "*"
            ):
                self._take("TEXT")
            return '"' + self._parse_required_text_group() + '" '
        if command == "mathop":
            return self._parse_required_group() + " "
        if command == "left":
            return self._parse_scalable_delimiters()
        if command == "begin":
            return self._parse_environment()
        if command in {"right", "end", "limits", "nolimits"}:
            _raise(FORMULA_MALFORMED, "unmatched \\" + command)
        if command in _LEGACY_ENVIRONMENT_NAMES:
            _raise(FORMULA_MALFORMED, "environment name outside \\begin")
        _raise(FORMULA_UNKNOWN_COMMAND, "\\" + command)
        return ""  # pragma: no cover

    def _parse_generalized_fraction(self) -> str:
        left = self._parse_required_group_allow_empty()
        right = self._parse_required_group_allow_empty()
        thickness = self._parse_required_group_allow_empty()
        style = self._parse_required_group_allow_empty()
        numerator = self._parse_required_group()
        denominator = self._parse_required_group()
        if left not in {"", "(", "[", "{", "|"}:
            _raise(FORMULA_MALFORMED, "invalid generalized-fraction delimiter")
        if right not in {"", ")", "]", "}", "|"}:
            _raise(FORMULA_MALFORMED, "invalid generalized-fraction delimiter")
        if thickness and re.fullmatch(
            r"(?:0|[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:pt|em|ex|mu))",
            thickness,
        ) is None:
            _raise(FORMULA_MALFORMED, "invalid generalized-fraction thickness")
        if style not in {"", "0", "1", "2", "3"}:
            _raise(FORMULA_MALFORMED, "invalid generalized-fraction style")
        return left + f"({numerator})/({denominator})" + right

    def _parse_negated_atom(self) -> str:
        self._skip_spaces()
        token = self._peek()
        if token is None:
            _raise(FORMULA_MALFORMED, "missing negated relation")
        direct_text = {"=": "≠", "<": "≮", ">": "≯"}
        direct_command = {
            "in": "∉", "le": "≰", "leq": "≰", "ge": "≱",
            "geq": "≱", "equiv": "≢", "approx": "≉",
            "subset": "⊄", "supset": "⊅", "subseteq": "⊈",
            "supseteq": "⊉",
        }
        if token.kind == "TEXT" and token.value in direct_text:
            self._take("TEXT")
            return direct_text[token.value]
        if token.kind == "COMMAND" and token.value in direct_command:
            self._take("COMMAND")
            return direct_command[token.value]
        atom = self._parse_atom()
        if not atom.strip():
            _raise(FORMULA_MALFORMED, "missing negated relation")
        return atom + "\u0338"

    def _parse_delimiter(self) -> str:
        self._skip_spaces()
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
        self._skip_spaces()
        self._take("LBRACE")
        parts: List[str] = []
        while self._peek() is not None and self._peek().kind != "RBRACE":
            token = self._take()
            if token.kind != "TEXT":
                _raise(FORMULA_MALFORMED, "invalid environment name")
            parts.append(token.value)
        self._take("RBRACE")
        name = "".join(parts)
        if re.fullmatch(r"[A-Za-z]+\*?", name) is None:
            _raise(FORMULA_MALFORMED, "invalid environment name")
        return name

    def _parse_environment(self) -> str:
        self._enter_nested()
        try:
            name = self._read_environment_name()
            base_name = name[:-1] if name.endswith("*") else name
            is_equation_array = base_name in _EQUATION_ARRAY_ENVIRONMENTS
            if base_name not in _MATRIX_DELIMITERS and not is_equation_array:
                _raise(FORMULA_UNKNOWN_COMMAND, "environment " + base_name)
            if name.endswith("*") and not is_equation_array:
                _raise(FORMULA_MALFORMED, "starred matrix environment")
            expected_columns: Optional[int] = None
            if base_name == "alignedat":
                pair_count = self._parse_required_group()
                if not pair_count.isdigit() or not 1 <= int(pair_count) <= 32:
                    _raise(FORMULA_TOO_COMPLEX, "alignedat pair limit exceeded")
                expected_columns = 2 * int(pair_count)
            if base_name == "array":
                column_spec = self._parse_required_text_group().replace(" ", "")
                if (
                    not column_spec
                    or re.fullmatch(r"[lcr|]{1,64}", column_spec) is None
                ):
                    _raise(FORMULA_MALFORMED, "unsupported array column specification")
                expected_columns = sum(char in "lcr" for char in column_spec)
            rows: List[List[str]] = []
            row: List[str] = []
            pending_column = False
            while True:
                self._skip_spaces()
                if self._at_command("end"):
                    if pending_column:
                        if not is_equation_array:
                            _raise(FORMULA_MALFORMED, "empty matrix cell")
                        row.append("")
                    if row:
                        rows.append(row)
                    break
                cell = self._parse_sequence(
                    stop_kinds=frozenset({"AMP", "ROW"}), stop_on_end=True
                )
                if not cell and not is_equation_array:
                    _raise(FORMULA_MALFORMED, "empty matrix cell")
                row.append(cell)
                pending_column = False
                if len(row) > _MAX_MATRIX_COLUMNS:
                    _raise(FORMULA_TOO_COMPLEX, "matrix column limit exceeded")
                token = self._peek()
                if token is None:
                    _raise(FORMULA_MALFORMED, "unterminated environment")
                if token.kind == "AMP":
                    self._take("AMP")
                    pending_column = True
                    continue
                if token.kind == "ROW":
                    self._take("ROW")
                    rows.append(row)
                    if len(rows) > _MAX_MATRIX_ROWS:
                        _raise(FORMULA_TOO_COMPLEX, "matrix row limit exceeded")
                    row = []
                    self._skip_spaces()
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
            if any(not any(cell.strip() for cell in candidate) for candidate in rows):
                _raise(FORMULA_MALFORMED, "empty equation row")
            if expected_columns is not None and any(
                len(candidate) != expected_columns for candidate in rows
            ):
                _raise(FORMULA_MALFORMED, "declared column count mismatch")
            if base_name == "cases" and width > 2:
                _raise(FORMULA_TOO_COMPLEX, "cases column limit exceeded")
            matrix = "■(" + "@".join("&".join(candidate) for candidate in rows) + ")"
            if is_equation_array:
                return matrix
            left, right = _MATRIX_DELIMITERS[base_name]
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
        if category in {"Cc", "Cf", "Cs"}:
            _raise(FORMULA_MALFORMED, "control character")
    if _EXTERNAL_REFERENCE_RE.search(normalized):
        _raise(FORMULA_FORBIDDEN_PRIMITIVE, "external reference")
    _check_braces(normalized)
    linear_text = _Parser(_tokenize(normalized)).parse()
    linear_text = " ".join(linear_text.split())
    if not linear_text:
        _raise(FORMULA_MALFORMED, "empty formula")
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    return _create_descriptor(linear_text, digest)


__all__ = [
    "FORMULA_FORBIDDEN_PRIMITIVE", "FORMULA_MALFORMED",
    "FORMULA_NESTING_TOO_DEEP", "FORMULA_TOO_COMPLEX", "FORMULA_TOO_LONG",
    "FORMULA_UNKNOWN_COMMAND", "LEGACY_ALLOWED_COMMANDS", "NativeMathConversionError",
    "NativeMathDescriptor", "convert_restricted_latex", "legacy_command_category",
]
