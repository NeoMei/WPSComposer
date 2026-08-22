from __future__ import annotations

import dataclasses
import hashlib
import json

import pytest

from skills.WPSComposer.scripts.longform.formula import (
    FORMULA_FORBIDDEN_PRIMITIVE,
    FORMULA_NESTING_TOO_DEEP,
    FORMULA_TOO_COMPLEX,
    FORMULA_UNKNOWN_COMMAND,
    validate_formula_source,
)
from skills.WPSComposer.scripts.longform.native_math import (
    NativeMathConversionError,
    NativeMathDescriptor,
    convert_restricted_latex,
)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("E = mc^2", "E=mc^2"),
        ("a_i^2+b^{n}", "a_i^2+b^n"),
        (r"\frac{a+b}{c}", "(a+b)/(c)"),
        (r"\sqrt{x}+\sqrt[3]{y}", "√(x)+√(3&y)"),
        (r"\sum\limits_{i=1}^{n}i+\prod_{j=1}^{m}j", "∑_(i=1)^n i+∏_(j=1)^m j"),
        (r"\int_{0}^{1}x\,dx", "∫_0^1 x dx"),
        (r"\left(\frac{x}{y}\right)", "((x)/(y))"),
        (r"\alpha+\Gamma\leq\infty", "α+Γ≤∞"),
        (r"\begin{pmatrix}a&b\\c&d\end{pmatrix}", "(■(a&b@c&d))"),
        (r"\begin{cases}x&x>0\\-x&x\leq0\end{cases}", "{■(x&x>0@-x&x≤0)"),
        (r"\frac{1}{1+\sqrt{x_i^2}}", "(1)/(1+√(x_i^2))"),
        ("α + 中文变量 = é", "α+中文变量=é"),
    ],
)
def test_supported_formula_families_convert_to_wps_linear(
    source: str,
    expected: str,
) -> None:
    descriptor = convert_restricted_latex(source)

    assert descriptor.syntax == "wps-linear-v1"
    assert descriptor.linear_text == expected
    assert len(descriptor.source_hash) == 64


def test_whitespace_is_normalized_deterministically() -> None:
    compact = convert_restricted_latex(r"\sum_{i=1}^{n} i")
    spaced = convert_restricted_latex("  \\sum _ { i = 1 } ^ { n }   i  ")

    assert compact.linear_text == spaced.linear_text


@pytest.mark.parametrize(
    "source",
    [
        r"\input{private.tex}",
        r"\href{https://example.test}{x}",
        r"\includegraphics{plot.png}",
        r"\newcommand{\x}{y}",
        r"\label{external}",
    ],
)
def test_forbidden_or_external_commands_are_rejected(source: str) -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)

    assert exc_info.value.code == FORMULA_FORBIDDEN_PRIMITIVE


def test_unknown_command_is_rejected_before_descriptor_exists() -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(r"\madeup{x}")

    assert exc_info.value.code == FORMULA_UNKNOWN_COMMAND


@pytest.mark.parametrize(
    "source",
    [
        r"\frac{x}",
        r"\sqrt[3{x}",
        r"\begin{matrix}a&b\\c\end{matrix}",
        r"\begin{matrix}a&b\end{pmatrix}",
        r"\begin{matrix}a&&b\end{matrix}",
        "{x",
        "x}",
        "x\x00+y",
    ],
)
def test_malformed_groups_environments_and_controls_are_rejected(source: str) -> None:
    with pytest.raises(NativeMathConversionError):
        convert_restricted_latex(source)


def test_excess_depth_is_rejected_by_converter_and_validator() -> None:
    source = "{" * 65 + "x" + "}" * 65

    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)
    validation = validate_formula_source(source)

    assert exc_info.value.code == FORMULA_NESTING_TOO_DEEP
    assert validation.valid is False
    assert validation.issues == (FORMULA_NESTING_TOO_DEEP,)


def test_excess_non_brace_parser_nesting_is_also_rejected() -> None:
    source = r"\left(" * 65 + "x" + r"\right)" * 65

    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)

    assert exc_info.value.code == FORMULA_NESTING_TOO_DEEP


def test_excess_token_count_is_rejected() -> None:
    source = "+".join("x" for _ in range(2050))

    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)

    assert exc_info.value.code == FORMULA_TOO_COMPLEX


@pytest.mark.parametrize(
    "source",
    [
        r"\begin{matrix}" + r"\\".join("x" for _ in range(65)) + r"\end{matrix}",
        r"\begin{matrix}" + "&".join("x" for _ in range(65)) + r"\end{matrix}",
    ],
)
def test_excess_matrix_rows_or_columns_are_rejected(source: str) -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)

    assert exc_info.value.code == FORMULA_TOO_COMPLEX


def test_nfc_normalization_controls_hash_and_output() -> None:
    decomposed = convert_restricted_latex("e\u0301+x")
    composed = convert_restricted_latex("é+x")

    assert decomposed == composed
    assert decomposed.linear_text == "é+x"


def test_source_hash_is_sha256_of_nfc_normalized_source() -> None:
    descriptor = convert_restricted_latex("e\u0301+x")

    assert descriptor.source_hash == (
        "2fc60ef46e5ee65fec777cce7ceb054bd99bc84b88b99410bbdca31bfc5820d1"
    )


def test_descriptor_is_immutable_json_safe_and_repr_does_not_retain_source() -> None:
    source = "  x   +   y  "
    descriptor = convert_restricted_latex(source)

    assert isinstance(descriptor, NativeMathDescriptor)
    with pytest.raises(dataclasses.FrozenInstanceError):
        descriptor.linear_text = "changed"  # type: ignore[misc]
    assert json.loads(json.dumps(dataclasses.asdict(descriptor))) == {
        "linear_text": "x+y",
        "source_hash": descriptor.source_hash,
        "syntax": "wps-linear-v1",
    }
    assert source not in repr(descriptor)
    assert "source=" not in repr(descriptor)


def test_repeated_conversion_is_byte_stable() -> None:
    first = convert_restricted_latex(r"\frac{\alpha_i}{\sqrt{1+x^2}}")
    second = convert_restricted_latex(r"\frac{\alpha_i}{\sqrt{1+x^2}}")

    assert first == second
    assert dataclasses.asdict(first) == dataclasses.asdict(second)


@pytest.mark.parametrize("source", ["", "   ", r"\displaystyle", r"\,", "{}"])
def test_blank_and_zero_width_formulas_are_rejected(source: str) -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(source)

    assert exc_info.value.code == "FORMULA_MALFORMED"
    assert validate_formula_source(source).valid is False


def test_script_cannot_attach_to_zero_width_command() -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex(r"\displaystyle^2")

    assert exc_info.value.code == "FORMULA_MALFORMED"


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        (r"\text{arg   max}", '"arg max"'),
        (r"\mbox{x  if  y}", '"x if y"'),
        (r"x + y", "x+y"),
        (r"x\,y", "x y"),
    ],
)
def test_text_commands_preserve_spaces_but_math_spaces_are_ignored(
    source: str,
    expected: str,
) -> None:
    assert convert_restricted_latex(source).linear_text == expected


@pytest.mark.parametrize("control", ["\t", "\n", "\r", "\x00", "\x1f", "\x7f"])
def test_all_control_characters_are_rejected(control: str) -> None:
    with pytest.raises(NativeMathConversionError) as exc_info:
        convert_restricted_latex("x" + control + "y")

    assert exc_info.value.code == "FORMULA_MALFORMED"


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        (r"\binom{n}{k}", "((n)¦(k))"),
        (r"\tbinom{n}{k}", "((n)¦(k))"),
        (r"\dbinom{n}{k}", "((n)¦(k))"),
        (r"\hat{x}", "(x)\u0302"),
        (r"\widehat{x+y}", "(x+y)\u0302"),
        (r"\tilde{x}", "(x)\u0303"),
        (r"\widetilde{x+y}", "(x+y)\u0303"),
        (r"\vec{x}", "(x)\u20d7"),
        (r"\bar{x}", "(x)\u0305"),
        (r"\dot{x}", "(x)\u0307"),
        (r"\ddot{x}", "(x)\u0308"),
        (r"\overline{x+y}", "(x+y)\u0305"),
        (r"\underline{x+y}", "(x+y)\u0332"),
        (r"\operatorname{arg max}_{x}f(x)", '"arg max"_x f(x)'),
        (r"\operatorname*{arg max}_{x}f(x)", '"arg max"_x f(x)'),
        (r"\mathop{lim}_{x\to0}f(x)", "lim_(x→0) f(x)"),
        (r"\wp(z)", "℘(z)"),
        (
            r"\begin{aligned}a&=b\\c&=d\end{aligned}",
            "■(a&=b@c&=d)",
        ),
        (
            r"\begin{split}a&=b+c\\&=d\end{split}",
            "■(a&=b+c@&=d)",
        ),
        (
            r"\begin{gathered}a=b\\c=d\end{gathered}",
            "■(a=b@c=d)",
        ),
        (r"\begin{align*}a&=b\end{align*}", "■(a&=b)"),
        (r"\begin{alignedat}{1}a&=b\end{alignedat}", "■(a&=b)"),
        (r"\begin{gather}a=b\\c=d\end{gather}", "■(a=b@c=d)"),
        (r"\begin{multline}a+b\\=c\end{multline}", "■(a+b@=c)"),
        (r"\begin{equation}a=b\end{equation}", "■(a=b)"),
        (r"\begin{array}{cc}a&b\\c&d\end{array}", "■(a&b@c&d)"),
    ],
)
def test_legacy_common_constructs_have_controlled_wps_linear_mappings(
    source: str,
    expected: str,
) -> None:
    assert convert_restricted_latex(source).linear_text == expected


@pytest.mark.parametrize(
    ("syntax", "linear_text", "source_hash"),
    [
        ("wps-linear-v1", "", hashlib.sha256(b"").hexdigest()),
        ("wps-linear-v1", "x" * 10_001, hashlib.sha256(b"x").hexdigest()),
        ("wps-linear-v1", "x\ny", hashlib.sha256(b"x").hexdigest()),
    ],
)
def test_descriptor_public_constructor_rejects_untrusted_content(
    syntax: str,
    linear_text: str,
    source_hash: str,
) -> None:
    with pytest.raises(ValueError):
        NativeMathDescriptor(syntax, linear_text, source_hash)


def test_parser_nested_failure_does_not_affect_later_conversion() -> None:
    with pytest.raises(NativeMathConversionError):
        convert_restricted_latex(r"\left(\left[x\right)")

    assert convert_restricted_latex(r"\left(x\right)").linear_text == "(x)"


@pytest.mark.parametrize(
    "source",
    [
        "E=mc^2",
        r"\frac{x}{y}",
        r"\sqrt[3]{x}",
        r"\begin{bmatrix}a&b\\c&d\end{bmatrix}",
    ],
)
def test_validator_accepts_exactly_what_converter_accepts(source: str) -> None:
    descriptor = convert_restricted_latex(source)
    validation = validate_formula_source(source)

    assert descriptor.linear_text
    assert validation.valid is True
    assert validation.issues == ()


def test_operatorname_star_is_accepted_by_converter_and_validator() -> None:
    source = r"\operatorname*{arg max}"

    assert convert_restricted_latex(source).linear_text == '"arg max"'
    assert validate_formula_source(source).valid is True
