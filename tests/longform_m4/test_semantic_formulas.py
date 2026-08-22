from __future__ import annotations

import json

import pytest

from skills.WPSComposer.scripts.document_model import CaptionBinding, FormulaBlock
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.native_math import (
    FORMULA_FORBIDDEN_PRIMITIVE,
    FORMULA_MALFORMED,
    FORMULA_NATIVE_EQUIVALENCE_UNSUPPORTED,
    FORMULA_NESTING_TOO_DEEP,
    FORMULA_TOO_LONG,
    FORMULA_UNKNOWN_COMMAND,
)
from skills.WPSComposer.scripts.longform.semantic import normalize_longform_document


def _formula_result(source: str, fallback_image: str = "private/formula.png"):
    markdown = (
        f':::equation {{#eq:test fallback_image="{fallback_image}"}}\n'
        f"{source}\n"
        ":::\n\n"
        "See {{ref:eq:test}}.\n"
    )
    return normalize_longform_document(parse_markdown(markdown, longform=True))


def _formula(result) -> FormulaBlock:
    return next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, FormulaBlock)
    )


def test_valid_formula_owns_only_a_trusted_native_descriptor() -> None:
    result = _formula_result(r"E = mc^2")
    formula = _formula(result)

    assert formula.native_math is not None
    assert formula.native_math.syntax == "wps-linear-v1"
    assert formula.content_degradation is None
    assert formula.identifier == "eq:test"
    assert formula.caption_binding is not None
    assert formula.caption_binding.referenceable is True
    assert result.references["eq:test"]["kind"] == "eq"


@pytest.mark.parametrize(
    ("source", "code"),
    [
        (r"\input{secret}", FORMULA_FORBIDDEN_PRIMITIVE),
        (r"\unknown{x}", FORMULA_UNKNOWN_COMMAND),
        ("x_{", FORMULA_MALFORMED),
        ("{" * 65 + "x" + "}" * 65, FORMULA_NESTING_TOO_DEEP),
        ("x" * 10_001, FORMULA_TOO_LONG),
        (r"\hspace{1em}x", FORMULA_NATIVE_EQUIVALENCE_UNSUPPORTED),
    ],
)
def test_invalid_formula_is_a_same_node_planned_degradation_but_keeps_target(
    source: str,
    code: str,
) -> None:
    result = _formula_result(source)
    formula = _formula(result)

    assert formula.native_math is None
    assert formula.content_degradation is not None
    assert formula.content_degradation.code == code
    assert formula.source == source
    assert formula.node_id == "eq:test"
    assert formula.caption_binding is not None
    assert formula.caption_binding.referenceable is True
    assert result.references["eq:test"]["node_id"] == "eq:test"
    assert any(issue.code == code and issue.placement == "block" for issue in result.issues)


def test_formula_fallback_path_is_private_in_semantic_json_and_repr() -> None:
    result = _formula_result("x + y", "private/secret-formula.png")
    formula = _formula(result)
    serialized = json.dumps(result.to_json(), ensure_ascii=False, sort_keys=True)

    assert "private/secret-formula.png" not in serialized
    assert "private/secret-formula.png" not in repr(formula)
    assert "fallback_image" not in serialized
    assert "sourceSha256" not in serialized
    assert "payloadSha256" not in serialized


def test_formula_semantic_json_is_byte_stable() -> None:
    markdown = ":::equation {#eq:a}\n\\frac{a}{b}\n:::\n"
    first = normalize_longform_document(parse_markdown(markdown, longform=True))
    second = normalize_longform_document(parse_markdown(markdown, longform=True))
    assert json.dumps(first.to_json(), ensure_ascii=False, separators=(",", ":")) == json.dumps(
        second.to_json(), ensure_ascii=False, separators=(",", ":")
    )


def test_formula_block_legacy_positional_constructor_remains_compatible() -> None:
    binding = CaptionBinding("global", None, "wpsc_eq_stable", True, True)
    formula = FormulaBlock("eq:legacy", "eq:legacy", "x", "(1)", binding)
    assert formula.identifier == "eq:legacy"
    assert formula.node_id == "eq:legacy"
    assert formula.source == "x"
    assert formula.number == "(1)"
    assert formula.caption_binding is binding
    assert formula.fallback_image is None


def test_formula_block_rejects_untrusted_native_descriptor() -> None:
    with pytest.raises(TypeError, match="converter-issued"):
        FormulaBlock(source="x", native_math={"linear_text": "x"})
