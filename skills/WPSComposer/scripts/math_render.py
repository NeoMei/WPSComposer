"""LaTeX math helpers — Unicode conversion (inline) and PNG rendering (display).

Inline math ($...$)  → readable Unicode approximation (no external deps)
Display math ($$...$$) → PNG image via matplotlib mathtext
"""

from __future__ import annotations

import hashlib
import os
import re
import tempfile

# ---------------------------------------------------------------------------
# Unicode conversion for inline math
# ---------------------------------------------------------------------------

_GREEK = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ε", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\vartheta": "ϑ", r"\iota": "ι", r"\kappa": "κ",
    r"\lambda": "λ", r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ",
    r"\pi": "π", r"\varpi": "ϖ", r"\rho": "ρ", r"\varrho": "ϱ",
    r"\sigma": "σ", r"\varsigma": "ς", r"\tau": "τ", r"\upsilon": "υ",
    r"\phi": "φ", r"\varphi": "φ", r"\chi": "χ", r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ", r"\Lambda": "Λ",
    r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ", r"\Upsilon": "Υ",
    r"\Phi": "Φ", r"\Psi": "Ψ", r"\Omega": "Ω",
}

_OPERATORS = {
    r"\leq": "≤", r"\le": "≤", r"\geq": "≥", r"\ge": "≥",
    r"\neq": "≠", r"\ne": "≠", r"\approx": "≈", r"\equiv": "≡",
    r"\sim": "∼", r"\propto": "∝",
    r"\in": "∈", r"\notin": "∉", r"\ni": "∋",
    r"\subset": "⊂", r"\supset": "⊃", r"\subseteq": "⊆", r"\supseteq": "⊇",
    r"\cup": "∪", r"\cap": "∩", r"\setminus": "∖", r"\emptyset": "∅", r"\varnothing": "∅",
    r"\forall": "∀", r"\exists": "∃", r"\nexists": "∄",
    r"\rightarrow": "→", r"\to": "→", r"\Rightarrow": "⇒",
    r"\leftarrow": "←", r"\gets": "←", r"\Leftarrow": "⇐",
    r"\leftrightarrow": "↔", r"\Leftrightarrow": "⇔",
    r"\mapsto": "↦", r"\hookrightarrow": "↪",
    r"\uparrow": "↑", r"\downarrow": "↓",
    r"\cdot": "·",     r"\cdots": "⋯", r"\ldots": "…", r"\dots": "…", r"\vdots": "⋮", r"\ddots": "⋱",
    r"\pm": "±", r"\mp": "∓", r"\times": "×", r"\div": "÷",
    r"\partial": "∂", r"\nabla": "∇", r"\infty": "∞",
    r"\sum": "∑", r"\prod": "∏", r"\int": "∫", r"\oint": "∮",
    r"\bigcup": "∪", r"\bigcap": "∩",
    r"\sqrt": "√",
    r"\star": "⋆", r"\ast": "∗", r"\dagger": "†", r"\ddagger": "‡",
    r"\bullet": "•", r"\circ": "∘", r"\diamond": "⋄",
    r"\succ": "≻", r"\prec": "≺",
    r"\succeq": "⪰", r"\preceq": "⪯",
    r"\hat": "", r"\bar": "", r"\tilde": "", r"\vec": "",
    r"\text": "", r"\mathrm": "", r"\mathbf": "", r"\mathit": "",
    r"\mathcal": "", r"\operatorname": "", r"\displaystyle": "",
    r"\limits": "", r"\nolimits": "",
    r"\min": "min", r"\max": "max", r"\arg": "arg",
    r"\sup": "sup", r"\inf": "inf",
    r"\dim": "dim", r"\det": "det", r"\ker": "ker",
    r"\deg": "deg", r"\log": "log", r"\ln": "ln", r"\exp": "exp",
    r"\sin": "sin", r"\cos": "cos", r"\tan": "tan",
    r"\Re": "ℜ", r"\Im": "ℑ", r"\aleph": "ℵ", r"\hbar": "ℏ",
    r"\ell": "ℓ", r"\imath": "ı", r"\jmath": "ȷ",
    r"\|": "‖", r"\Vert": "‖",
    r"\langle": "⟨", r"\rangle": "⟩",
    r"\lceil": "⌈", r"\rceil": "⌉", r"\lfloor": "⌊", r"\rfloor": "⌋",
    r"\big": "", r"\Big": "", r"\bigg": "", r"\Bigg": "",
    r"\frac": "/",  # fallback for \frac not caught by _convert_frac
    r"\left": "", r"\right": "",
    r"\quad": "  ", r"\qquad": "    ",
    r"\,": " ", r"\;": " ", r"\:": " ", r"\!": "",
    r"\%": "%", r"\#": "#", r"\$": "$", r"\&": "&", r"\_": "_",
    r"\{": "{", r"\}": "}", r"\backslash": "\\",
}

_SUPER_MAP = str.maketrans("0123456789+-=()nivx*", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱⱽˣ∗")
# Unicode has subscripts for: 0-9 + - = a e h i k l m n o p r s t u x
# No subscript exists for: b c d f g j q w y z — those fall back to _(x)
_SUB_MAP = str.maketrans(
    "0123456789+-=aehiklmnoprstux",
    "₀₁₂₃₄₅₆₇₈₉₊₋₌ₐₑₕᵢₖₗₘₙₒₚᵣₛₜᵤₓ",
)

# Pre-compiled: matches \command (backslash + letters only)
_LETTER_CMD_RE = re.compile(r"\\([a-zA-Z]+)")


def _convert_scripts(latex: str) -> str:
    """Convert x^{...} and x_{...} to Unicode super/subscripts where possible."""
    _SUP_CHARS = frozenset("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱⱽˣ∗")
    _SUB_CHARS = frozenset("₀₁₂₃₄₅₆₇₈₉₊₋₌ₐₑₕᵢₖₗₘₙₒₚᵣₛₜᵤₓ")

    # Superscript: ^{...} or ^c (single char including * + -)
    def _sup_repl(m):
        g1, g2 = m.group(1), m.group(2)
        content = g1 if g1 is not None else g2
        if content is None or content == "":
            return ""
        converted = content.translate(_SUPER_MAP)
        if not all(c in _SUP_CHARS for c in converted):
            return f"^({content})"
        return converted

    latex = re.sub(r"\^\{([^}]*)\}|\^([a-zA-Z0-9*\-+])", _sup_repl, latex)

    # Subscript: _{...} or _c (single char)
    def _sub_repl(m):
        g1, g2 = m.group(1), m.group(2)
        content = g1 if g1 is not None else g2
        if content is None or content == "":
            return ""
        converted = content.translate(_SUB_MAP)
        if not all(c in _SUB_CHARS for c in converted):
            return f"_({content})"
        return converted

    latex = re.sub(r"_\{([^}]*)\}|_([a-zA-Z0-9*\-+])", _sub_repl, latex)
    return latex


def _convert_frac(latex: str) -> str:
    """Convert \\frac{a}{b} → (a)/(b) for inline readability."""
    def _frac_repl(m):
        return f"({m.group(1)})/({m.group(2)})"
    # handle nested fractions by iterating
    prev = None
    while prev != latex:
        prev = latex
        latex = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", _frac_repl, latex)
    return latex


def latex_to_unicode(latex: str) -> str:
    """Convert a LaTeX math snippet to a readable Unicode approximation."""
    result = latex.strip()
    # strip surrounding $ if present
    result = result.strip("$")

    # Remove \text{...}, \mathrm{...} wrappers but keep content
    result = re.sub(r"\\(?:text|mathrm|mathbf|mathit|mathcal|operatorname)\{([^{}]*)\}", r"\1", result)

    # Convert fractions first
    result = _convert_frac(result)

    # Handle \sqrt[n]{x} — cube root etc. (before command regex strips \sqrt)
    result = re.sub(
        r"\\sqrt\[(\d+)\]\{([^{}]*)\}",
        lambda m: {3: "∛", 4: "∜"}.get(int(m.group(1)), f"√[{m.group(1)}]") + m.group(2),
        result,
    )

    # Replace commands: letter-commands via regex (word-boundary safe),
    # symbol-commands via str.replace (no collision possible).
    all_cmds = {**_GREEK, **_OPERATORS}
    _letter_cmds = {k: v for k, v in all_cmds.items() if len(k) > 1 and k[1].isalpha()}
    _symbol_cmds = {k: v for k, v in all_cmds.items() if k not in _letter_cmds}

    # Symbol commands first (\|, \,, \;, etc.) — safe with str.replace
    for cmd in sorted(_symbol_cmds, key=len, reverse=True):
        result = result.replace(cmd, _symbol_cmds[cmd])

    # Letter commands via single regex pass — \cmd only matches as a
    # complete command, never as a prefix of \cmdXXX
    _letter_pattern = _LETTER_CMD_RE
    def _letter_repl(m):
        full = "\\" + m.group(1)
        return _letter_cmds.get(full, full)
    result = _letter_pattern.sub(_letter_repl, result)

    # Convert super/subscripts
    result = _convert_scripts(result)

    # Remaining braces → just remove them
    result = result.replace("{", "").replace("}", "")

    # Clean up multiple spaces
    result = re.sub(r"  +", " ", result).strip()

    return result


# ---------------------------------------------------------------------------
# PNG rendering for display math (matplotlib mathtext)
# ---------------------------------------------------------------------------

_png_cache: dict[str, str] = {}


def latex_to_png(latex: str, dpi: int = 200) -> str:
    """Render a LaTeX math expression to a tightly-cropped PNG via matplotlib.

    Returns the path to the generated PNG file.  Results are cached by
    expression hash so repeated formulas don't re-render.
    """
    key = hashlib.md5(latex.encode()).hexdigest()
    if key in _png_cache and os.path.exists(_png_cache[key]):
        return _png_cache[key]

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig = plt.figure(figsize=(0.1, 0.1))
    fig.text(0, 0, f"${latex}$", fontsize=14, color="black")
    fig.canvas.draw()
    bbox = fig.get_tightbbox(fig.canvas.get_renderer())

    # Re-render with correct size
    plt.close(fig)
    w = max(bbox.width + 0.3, 0.5)
    h = max(bbox.height + 0.2, 0.3)
    fig = plt.figure(figsize=(w, h), dpi=dpi)
    fig.patch.set_alpha(0)  # transparent background
    fig.text(0.5, 0.5, f"${latex}$", fontsize=14, color="black",
             ha="center", va="center")
    out_path = os.path.join(tempfile.gettempdir(), f"math_{key}.png")
    fig.savefig(out_path, transparent=True, dpi=dpi,
                bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    _png_cache[key] = out_path
    return out_path
