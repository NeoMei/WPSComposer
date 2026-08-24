from __future__ import annotations

import io
import json
from dataclasses import replace

from PIL import Image
import pytest

from skills.WPSComposer.scripts.document_model import FigureBlock, FormulaBlock, ImageBlock
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.pipeline import (
    _build_executor_resources,
    build_longform_generation,
    execute_longform_plan,
)
from skills.WPSComposer.scripts.longform.resources import (
    FORMULA_FALLBACK_IMAGE_UNAVAILABLE,
    FORMULA_RESOURCE_BINDING_INVALID,
    PreparedLongformResource,
    preflight_resources,
    validate_formula_resource_bindings,
)
from skills.WPSComposer.scripts.longform.semantic import normalize_longform_document


def _png_bytes() -> bytes:
    output = io.BytesIO()
    Image.new("RGB", (3, 2), (20, 40, 60)).save(output, format="PNG")
    return output.getvalue()


def _semantic(markdown: str, base_dir: str = ""):
    return normalize_longform_document(
        parse_markdown(markdown, base_dir=base_dir, longform=True)
    )


def test_explicit_formula_fallback_is_preflighted_and_bound_exactly_once(tmp_path) -> None:
    (tmp_path / "fallback.png").write_bytes(_png_bytes())
    semantic = _semantic(
        ':::equation {#eq:a fallback_image="fallback.png"}\nx + y\n:::\n'
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))

    assert len(preflight.resources) == 1
    resource = preflight.resources[0]
    assert preflight.formula_bindings == {"eq:a": resource.resource_id}
    assert preflight.formula_resource_ids == frozenset({resource.resource_id})
    assert preflight.formula_node_ids == frozenset({"eq:a"})
    assert preflight.degradations == []
    prepared = _build_executor_resources(str(tmp_path), preflight)
    assert len(prepared) == 1
    assert isinstance(prepared[0], PreparedLongformResource)


def test_formula_without_explicit_fallback_scans_no_formula_resource(tmp_path) -> None:
    (tmp_path / "unmentioned.png").write_bytes(_png_bytes())
    semantic = _semantic(":::equation {#eq:a}\nunmentioned.png + x\n:::\n")
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    assert preflight.resources == []
    assert preflight.formula_bindings == {}


def test_missing_or_invalid_formula_fallback_degrades_locally_without_suppressing_math(
    tmp_path,
) -> None:
    (tmp_path / "invalid.png").write_text("not an image", encoding="utf-8")
    for name in ("missing.png", "invalid.png"):
        semantic = _semantic(
            f':::equation {{#eq:a fallback_image="{name}"}}\nx + y\n:::\n'
        )
        formula = next(
            element
            for section in semantic.document.sections
            for element in section.elements
            if isinstance(element, FormulaBlock)
        )
        preflight = preflight_resources(semantic.document.sections, str(tmp_path))
        assert formula.native_math is not None
        assert preflight.formula_bindings == {}
        assert len(preflight.degradations) == 1
        degradation = preflight.degradations[0]
        assert degradation.node_id == "eq:a"
        assert degradation.code == FORMULA_FALLBACK_IMAGE_UNAVAILABLE
        assert name not in degradation.message
        assert name not in degradation.fallback_text


def test_one_fallback_resource_cannot_bind_two_formula_nodes(tmp_path) -> None:
    (tmp_path / "shared.png").write_bytes(_png_bytes())
    semantic = _semantic(
        """:::equation {#eq:a fallback_image="shared.png"}
x
:::
:::equation {#eq:b fallback_image="shared.png"}
y
:::
"""
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    assert preflight.formula_bindings == {}
    assert len(preflight.degradations) == 2
    assert {item.node_id for item in preflight.degradations} == {"eq:a", "eq:b"}
    assert {item.code for item in preflight.degradations} == {
        FORMULA_FALLBACK_IMAGE_UNAVAILABLE
    }


def test_equivalent_relative_paths_cannot_bypass_single_formula_binding(tmp_path) -> None:
    (tmp_path / "shared.png").write_bytes(_png_bytes())
    semantic = _semantic(
        """:::equation {#eq:a fallback_image="shared.png"}
x
:::
:::equation {#eq:b fallback_image="./shared.png"}
y
:::
"""
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    assert preflight.resources == []
    assert preflight.formula_bindings == {}
    assert {item.node_id for item in preflight.degradations} == {"eq:a", "eq:b"}


def test_figure_resource_does_not_become_an_implicit_formula_binding(tmp_path) -> None:
    (tmp_path / "figure.png").write_bytes(_png_bytes())
    semantic = _semantic(
        """:::figure {#fig:a caption="Figure"}
![A](figure.png)
:::
:::equation {#eq:a}
x
:::
""",
        base_dir=str(tmp_path),
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    assert len(preflight.resources) == 1
    assert preflight.formula_bindings == {}


def test_formula_resource_paths_hashes_bytes_and_bindings_are_private(tmp_path) -> None:
    secret = "private-secret.png"
    (tmp_path / secret).write_bytes(_png_bytes())
    semantic = _semantic(
        f':::equation {{#eq:a fallback_image="{secret}"}}\nx\n:::\n'
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    formula = next(
        element
        for section in semantic.document.sections
        for element in section.elements
        if isinstance(element, FormulaBlock)
    )

    semantic_json = json.dumps(semantic.to_json(), ensure_ascii=False, sort_keys=True)
    assert secret not in semantic_json
    assert secret not in repr(formula)
    assert secret not in repr(preflight)
    assert "payload_bytes" not in repr(preflight)
    assert preflight.resources[0].source_sha256 not in repr(preflight)
    assert preflight.resources[0].payload_sha256 not in repr(preflight)


def test_orphaned_or_multiply_bound_formula_resources_fail_before_preparation(
    tmp_path,
) -> None:
    (tmp_path / "fallback.png").write_bytes(_png_bytes())
    semantic = _semantic(
        ':::equation {#eq:a fallback_image="fallback.png"}\nx\n:::\n'
    )
    preflight = preflight_resources(semantic.document.sections, str(tmp_path))
    resource_id = preflight.resources[0].resource_id

    invalid_cases = (
        replace(preflight, formula_bindings={}),
        replace(preflight, formula_bindings={"eq:fake": resource_id}),
        replace(
            preflight,
            formula_bindings={"eq:a": resource_id, "eq:b": resource_id},
        ),
        replace(
            preflight,
            resources=[preflight.resources[0], preflight.resources[0]],
        ),
    )
    for invalid in invalid_cases:
        try:
            validate_formula_resource_bindings(invalid)
        except ValueError as error:
            assert str(error) == FORMULA_RESOURCE_BINDING_INVALID
        else:
            raise AssertionError("invalid formula binding was accepted")
        try:
            _build_executor_resources(str(tmp_path), invalid)
        except ValueError as error:
            assert str(error) == FORMULA_RESOURCE_BINDING_INVALID
        else:
            raise AssertionError("invalid binding reached prepared resources")


def test_invalid_formula_binding_is_rejected_before_executor_start(tmp_path) -> None:
    (tmp_path / "fallback.png").write_bytes(_png_bytes())
    build = build_longform_generation(
        ':::equation {#eq:a fallback_image="fallback.png"}\nx\n:::\n',
        base_dir=str(tmp_path),
    )
    resource_id = build.preflight.resources[0].resource_id
    invalid = replace(
        build,
        preflight=replace(
            build.preflight,
            formula_bindings={"eq:fake": resource_id},
        ),
    )

    class NeverStartedExecutor:
        called = False

        def execute(self, plan, resources, deadline=None):
            self.called = True
            raise AssertionError("executor must not start")

    executor = NeverStartedExecutor()
    with pytest.raises(ValueError, match=FORMULA_RESOURCE_BINDING_INVALID):
        execute_longform_plan(invalid, executor)
    assert executor.called is False


def test_ordinary_resource_sharing_remains_valid(tmp_path) -> None:
    (tmp_path / "shared.png").write_bytes(_png_bytes())
    preflight = preflight_resources(
        [
            ImageBlock(path="shared.png", alt="first"),
            ImageBlock(path="shared.png", alt="second"),
        ],
        str(tmp_path),
    )
    assert len(preflight.resources) == 1
    assert preflight.formula_bindings == {}
    assert preflight.formula_node_ids == frozenset()
    validate_formula_resource_bindings(preflight)


def test_build_repr_semantic_and_plan_surfaces_hide_formula_resource_details(
    tmp_path,
) -> None:
    secret = "private-fallback.png"
    (tmp_path / secret).write_bytes(_png_bytes())
    build = build_longform_generation(
        f':::equation {{#eq:a fallback_image="{secret}"}}\nx\n:::\n',
        base_dir=str(tmp_path),
    )
    diagnostic = json.dumps(build.to_json(), ensure_ascii=False, sort_keys=True)
    combined = repr(build) + diagnostic
    assert secret not in combined
    assert str(tmp_path) not in combined
    assert "formula_bindings" not in combined
    assert "formula_resource_ids" not in combined
    assert build.preflight.resources[0].source_sha256 not in combined
    assert build.preflight.resources[0].payload_sha256 not in combined
