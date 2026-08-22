"""Restricted formula validation backed by the native-math converter."""

from __future__ import annotations

from dataclasses import dataclass

from .native_math import (
    FORMULA_FORBIDDEN_PRIMITIVE,
    FORMULA_MALFORMED,
    FORMULA_NESTING_TOO_DEEP,
    FORMULA_TOO_COMPLEX,
    FORMULA_TOO_LONG,
    FORMULA_UNKNOWN_COMMAND,
    NativeMathConversionError,
    convert_restricted_latex,
)


@dataclass(frozen=True)
class FormulaValidation:
    """Result of validating a formula source string."""

    valid: bool
    issues: tuple[str, ...]
    fallback_text: str


def validate_formula_source(source: str) -> FormulaValidation:
    """Validate exactly the subset accepted by native-math conversion.

    The legacy API remains non-raising. Conversion itself is the validator, so
    a successful result cannot fail later because of a second grammar.
    """
    if not isinstance(source, str):
        source = str(source)
    try:
        convert_restricted_latex(source)
    except NativeMathConversionError as exc:
        return FormulaValidation(
            valid=False,
            issues=(str(exc),),
            fallback_text=source,
        )
    return FormulaValidation(valid=True, issues=(), fallback_text=source)


__all__ = [
    "FORMULA_FORBIDDEN_PRIMITIVE",
    "FORMULA_MALFORMED",
    "FORMULA_NESTING_TOO_DEEP",
    "FORMULA_TOO_COMPLEX",
    "FORMULA_TOO_LONG",
    "FORMULA_UNKNOWN_COMMAND",
    "FormulaValidation",
    "validate_formula_source",
]
