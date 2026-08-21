from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.longform.formula import (
    FORMULA_FORBIDDEN_PRIMITIVE,
    FORMULA_NESTING_TOO_DEEP,
    FORMULA_TOO_LONG,
    FORMULA_UNKNOWN_COMMAND,
    FormulaValidation,
    validate_formula_source,
)


def test_valid_simple_formula_is_accepted() -> None:
    result = validate_formula_source("E = mc^2")
    assert isinstance(result, FormulaValidation)
    assert result.valid is True
    assert result.issues == ()
    assert result.fallback_text == "E = mc^2"


def test_valid_fraction_sqrt_sum_are_accepted() -> None:
    result = validate_formula_source(
        r"\sum_{i=1}^{n} \sqrt{\frac{a_i}{b_i}}"
    )
    assert result.valid is True
    assert result.issues == ()


def test_valid_greek_letters_are_accepted() -> None:
    result = validate_formula_source(r"\alpha \beta + \Gamma \Delta")
    assert result.valid is True


def test_valid_matrix_is_accepted() -> None:
    result = validate_formula_source(
        r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}"
    )
    assert result.valid is True


def test_input_is_rejected() -> None:
    result = validate_formula_source(r"\input{file.tex}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)
    assert r"\input" in result.fallback_text


def test_include_is_rejected() -> None:
    result = validate_formula_source(r"\include{file.tex}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_write_is_rejected() -> None:
    result = validate_formula_source(r"\write18{rm -rf /}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_shell_escape_is_rejected() -> None:
    result = validate_formula_source(r"\immediate\write18{id}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_def_macro_is_rejected() -> None:
    result = validate_formula_source(r"\def\foo{bar}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_newcommand_is_rejected() -> None:
    result = validate_formula_source(r"\newcommand{\foo}{bar}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_usepackage_is_rejected() -> None:
    result = validate_formula_source(r"\usepackage{amsmath}")
    assert result.valid is False
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_unknown_command_is_rejected() -> None:
    result = validate_formula_source(r"\custommacro{x}")
    assert result.valid is False
    assert any(FORMULA_UNKNOWN_COMMAND in issue for issue in result.issues)
    assert any(r"\custommacro" in issue for issue in result.issues)


def test_too_long_formula_is_rejected() -> None:
    source = "x" * 10001
    result = validate_formula_source(source)
    assert result.valid is False
    assert any(FORMULA_TOO_LONG in issue for issue in result.issues)


def test_too_deep_brace_nesting_is_rejected() -> None:
    source = "x" + "{" * 65 + "y" + "}" * 65
    result = validate_formula_source(source)
    assert result.valid is False
    assert any(FORMULA_NESTING_TOO_DEEP in issue for issue in result.issues)


def test_deep_but_allowed_nesting_is_accepted() -> None:
    source = "x" + "{" * 64 + "y" + "}" * 64
    result = validate_formula_source(source)
    assert result.valid is True


def test_visible_fallback_text_for_invalid_formula() -> None:
    source = r"\input{secrets}"
    result = validate_formula_source(source)
    assert result.valid is False
    assert result.fallback_text == source
    assert any(FORMULA_FORBIDDEN_PRIMITIVE in issue for issue in result.issues)


def test_empty_formula_is_accepted() -> None:
    result = validate_formula_source("")
    assert result.valid is True
    assert result.issues == ()
    assert result.fallback_text == ""