"""Conservative LaTeX formula validation for the long-form engine."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

FORMULA_TOO_LONG = "FORMULA_TOO_LONG"
FORMULA_NESTING_TOO_DEEP = "FORMULA_NESTING_TOO_DEEP"
FORMULA_FORBIDDEN_PRIMITIVE = "FORMULA_FORBIDDEN_PRIMITIVE"
FORMULA_UNKNOWN_COMMAND = "FORMULA_UNKNOWN_COMMAND"

_MAX_CODE_POINTS = 10_000
_MAX_BRACE_DEPTH = 64

_ALLOWED_COMMANDS = frozenset({
    # Greek letters
    "alpha", "beta", "gamma", "delta", "epsilon", "varepsilon", "zeta", "eta",
    "theta", "vartheta", "iota", "kappa", "lambda", "mu", "nu", "xi", "pi",
    "varpi", "rho", "varrho", "sigma", "varsigma", "tau", "upsilon", "phi",
    "varphi", "chi", "psi", "omega", "Gamma", "Delta", "Theta", "Lambda",
    "Xi", "Pi", "Sigma", "Upsilon", "Phi", "Psi", "Omega",
    # Common operators and relations
    "times", "div", "pm", "mp", "cdot", "ast", "star", "circ", "bullet",
    "oplus", "ominus", "otimes", "oslash", "leq", "geq", "le", "ge", "neq",
    "ne", "approx", "sim", "simeq", "equiv", "cong", "propto", "in", "notin",
    "subset", "supset", "subseteq", "supseteq", "cup", "cap", "setminus",
    "emptyset", "forall", "exists", "nexists", "neg", "land", "lor", "wedge",
    "vee", "to", "gets", "rightarrow", "leftarrow", "Rightarrow", "Leftarrow",
    "leftrightarrow", "Leftrightarrow", "mapsto", "iff", "infty", "nabla",
    "partial", "prime", "hbar", "ell", "wp", "Re", "Im",
    # Named functions
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "exp", "lim", "sup", "inf", "max",
    "min", "arg", "dim", "det", "ker", "mod", "gcd", "lcm", "Pr",
    # Fractions, roots, large operators
    "frac", "dfrac", "tfrac", "sqrt", "sum", "prod", "int", "oint", "iint",
    "iiint", "iiiint", "limits", "nolimits", "displaystyle", "textstyle",
    "scriptstyle", "scriptscriptstyle",
    # Delimiters and matrices
    "left", "right", "begin", "end", "matrix", "pmatrix", "bmatrix", "vmatrix",
    "Vmatrix", "Bmatrix", "smallmatrix", "cases", "align", "aligned",
    "alignedat", "gather", "multline", "equation", "array", "split",
    "flalign",
    # Text and font commands
    "text", "mbox", "mathrm", "mathbf", "mathit", "mathsf", "mathtt",
    "mathcal", "mathbb", "mathfrak", "mathscr", "boldsymbol", "bm",
    "operatorname", "mathop",
    # Accents and decorations
    "overline", "underline", "hat", "widehat", "tilde", "widetilde", "vec",
    "bar", "dot", "ddot", "acute", "grave", "check", "breve", "overbrace",
    "underbrace", "overset", "underset", "stackrel", "buildrel", "atop",
    "choose", "brack", "brace",
    # Spacing
    ",", ":", ";", "!", "quad", "qquad", "space", "thinspace", "medspace",
    "thickspace", "enspace", "hspace", "hskip", "vspace", "vskip", "kern",
    "mskip", "mkern", "raisebox", "lower", "box", "phantom", "vphantom",
    "hphantom",
    # Dots and binomials
    "ldots", "cdots", "vdots", "ddots", "dots", "binom", "tbinom", "dbinom",
    "bmod", "pmod", "pod", "genfrac",
    # Math spacing/sizing
    "big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr", "biggl",
    "biggr", "Biggl", "Biggr",
    # Misc
    "nonumber", "tag", "label", "not",
})

_FORBIDDEN_COMMANDS = frozenset({
    # File I/O and shell escape
    "input", "include", "write", "write18", "openout", "openin", "immediate",
    "csname", "catcode", "escapechar", "endcsname", "noexpand", "expandafter",
    # Macro / package definitions
    "def", "gdef", "edef", "xdef", "let", "futurelet", "global", "newcommand",
    "renewcommand", "newenvironment", "renewenvironment", "DeclareMathOperator",
    "usepackage", "documentclass", "RequirePackage", "LoadClass",
    # External references / hyperlinks / graphics
    "href", "url", "path", "includegraphics", "graphicspath",
    # Engine modes that could alter execution
    "batchmode", "nonstopmode", "scrollmode", "errorstopmode",
    # System / shell primitives
    "shell", "system", "exec", "open", "file", "read",
    "directlua", "luadirect", "writefile",
})

_COMMAND_RE = re.compile(r"\\([A-Za-z]+|.)")


@dataclass(frozen=True)
class FormulaValidation:
    """Result of validating a formula source string."""

    valid: bool
    issues: tuple[str, ...]
    fallback_text: str


def _brace_depth(source: str) -> tuple[int, bool]:
    r"""Return (maximum_depth, well_formed).

    Only structural braces are counted; escaped braces (\{, \}) are skipped
    because they are commands, not grouping tokens.
    """
    depth = 0
    max_depth = 0
    i = 0
    n = len(source)
    while i < n:
        ch = source[i]
        if ch == "\\":
            i += 1
            if i < n and source[i].isalpha():
                while i < n and source[i].isalpha():
                    i += 1
            else:
                i += 1
            continue
        if ch == "{":
            depth += 1
            if depth > max_depth:
                max_depth = depth
        elif ch == "}":
            depth -= 1
            if depth < 0:
                return max_depth, False
        i += 1
    return max_depth, depth == 0


def validate_formula_source(source: str) -> FormulaValidation:
    """Validate a restricted LaTeX formula subset.

    Returns a FormulaValidation without raising for malformed input.
    """
    if not isinstance(source, str):
        source = str(source)

    code_point_count = len(source)
    if code_point_count > _MAX_CODE_POINTS:
        return FormulaValidation(
            valid=False,
            issues=(FORMULA_TOO_LONG,),
            fallback_text=source,
        )

    max_depth, well_formed = _brace_depth(source)
    if not well_formed or max_depth > _MAX_BRACE_DEPTH:
        return FormulaValidation(
            valid=False,
            issues=(FORMULA_NESTING_TOO_DEEP,),
            fallback_text=source,
        )

    issues: list[str] = []
    for match in _COMMAND_RE.finditer(source):
        cmd = match.group(1)
        if cmd in _FORBIDDEN_COMMANDS:
            issues.append(f"{FORMULA_FORBIDDEN_PRIMITIVE}: \\{cmd}")
        elif cmd.isalpha() and cmd not in _ALLOWED_COMMANDS:
            issues.append(f"{FORMULA_UNKNOWN_COMMAND}: \\{cmd}")

    if issues:
        return FormulaValidation(
            valid=False,
            issues=tuple(issues),
            fallback_text=source,
        )

    return FormulaValidation(
        valid=True,
        issues=(),
        fallback_text=source,
    )


__all__ = [
    "FORMULA_FORBIDDEN_PRIMITIVE",
    "FORMULA_NESTING_TOO_DEEP",
    "FORMULA_TOO_LONG",
    "FORMULA_UNKNOWN_COMMAND",
    "FormulaValidation",
    "validate_formula_source",
]