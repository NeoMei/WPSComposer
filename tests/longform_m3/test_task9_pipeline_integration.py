"""M3 Task 9 integration, privacy, and compatibility contracts."""

from __future__ import annotations

import base64
import gc
import hashlib
import json
import sys
import weakref
from pathlib import Path

import pytest
from PIL import Image

from skills.WPSComposer.scripts.longform.executor import ExecutionOutcome
from skills.WPSComposer.scripts.generation_plan import validate_generation_plan
from skills.WPSComposer.scripts.longform.pipeline import (
    build_longform_generation,
    execute_longform_plan,
)
from skills.WPSComposer.scripts.longform.resources import PreparedLongformResource
from tests.longform_m2.test_acceptance_m2 import (
    FIXTURES_DIR as M2_FIXTURES,
    SNAPSHOTS_DIR as M2_SNAPSHOTS,
    build_fixture as build_m2_fixture,
    structural_snapshot as m2_structural_snapshot,
)


def _write_png(path: Path) -> bytes:
    Image.new("RGB", (3, 2), (31, 63, 127)).save(path, format="PNG")
    return path.read_bytes()


def _image_build(tmp_path: Path):
    _write_png(tmp_path / "private.png")
    return build_longform_generation(
        """---
title: Task 9 privacy
---

# Chapter

:::figure {#fig:private caption="Private"}
![private](private.png)
:::

See {{ref:fig:private}}.
""",
        base_dir=str(tmp_path),
    )


def test_longform_package_exports_only_platform_pure_m3_contracts() -> None:
    import skills.WPSComposer.scripts.longform as longform

    intended = {
        "CaptionBinding",
        "CrossReferenceRun",
        "FigureImageLayout",
        "ImageProfile",
        "LongformBuild",
        "NativeFieldAdapter",
        "NativeFieldContractError",
        "PreparedLongformResource",
        "TableMerge",
        "TablePolicy",
        "build_longform_generation",
        "execute_longform_plan",
    }
    assert intended <= set(longform.__all__)
    assert not hasattr(longform, "WindowsLongformExecutor")
    assert not hasattr(longform, "MacOSLongformExecutor")


def test_build_stays_platform_pure_and_does_not_load_wps_modules(tmp_path: Path) -> None:
    forbidden = (
        "skills.WPSComposer.scripts.writer",
        "skills.WPSComposer.scripts._dispatch",
        "skills.WPSComposer.scripts.macos_probe",
        "skills.WPSComposer.scripts.longform.windows_executor",
        "skills.WPSComposer.scripts.longform.macos_executor",
        "win32com",
        "pythoncom",
    )
    before = set(sys.modules)
    _image_build(tmp_path)
    newly_loaded = set(sys.modules) - before
    assert not [name for name in newly_loaded if name.startswith(forbidden)]


def test_diagnostic_json_redacts_private_payload_hashes_paths_and_bytes(
    tmp_path: Path,
) -> None:
    payload = _write_png(tmp_path / "private.png")
    build = _image_build(tmp_path)
    resource = build.preflight.resources[0]
    serialized = json.dumps(build.to_json(), ensure_ascii=False, sort_keys=True)

    assert str(tmp_path) not in serialized
    assert "private.png" not in serialized
    assert resource.source_sha256 not in serialized
    assert resource.payload_sha256 not in serialized
    assert base64.b64encode(payload).decode("ascii") not in serialized
    assert "payload_bytes" not in serialized
    assert "bookmarkNames" not in serialized
    assert "fieldResults" not in serialized
    assert "resultHash" not in serialized


class _EphemeralExecutor:
    def __init__(self, failure: BaseException | None = None) -> None:
        self.failure = failure
        self.resources_type = None
        self.resource_refs: tuple[weakref.ReferenceType[PreparedLongformResource], ...] = ()

    def execute(self, plan, resources, deadline=None):
        self.resources_type = type(resources)
        self.resource_refs = tuple(weakref.ref(resource) for resource in resources)
        assert all(isinstance(resource, PreparedLongformResource) for resource in resources)
        if self.failure is not None:
            raise self.failure
        return ExecutionOutcome(staged_artifact="staged.docx")


@pytest.mark.parametrize(
    "failure",
    [None, RuntimeError("executor failed"), TimeoutError("executor timed out")],
    ids=("success", "error", "timeout"),
)
def test_execution_boundary_releases_prepared_resource_references(
    tmp_path: Path, failure: BaseException | None
) -> None:
    build = _image_build(tmp_path)
    executor = _EphemeralExecutor(failure)

    if failure is None:
        execute_longform_plan(build, executor, deadline=123.0)
    else:
        with pytest.raises(type(failure)):
            execute_longform_plan(build, executor, deadline=123.0)

    gc.collect()
    assert executor.resources_type is tuple
    assert executor.resource_refs
    assert all(reference() is None for reference in executor.resource_refs)


@pytest.mark.parametrize(
    "name",
    (
        "academic",
        "wide_figure",
        "hybrid_bid",
        "degradation",
        "plain_short",
        "protocol_edge",
    ),
)
def test_m2_structural_snapshots_remain_byte_compatible(name: str) -> None:
    build = build_m2_fixture(name)
    actual = json.dumps(
        m2_structural_snapshot(build.plan),
        ensure_ascii=False,
        indent=2,
    ) + "\n"
    assert actual.encode("utf-8") == (M2_SNAPSHOTS / f"{name}_ops.json").read_bytes()
    assert M2_FIXTURES.is_dir()


def test_plan_determinism_keeps_only_the_required_manifest_digest(tmp_path: Path) -> None:
    first = _image_build(tmp_path)
    second = _image_build(tmp_path)
    first_plan = first.plan.to_dict()
    second_plan = second.plan.to_dict()
    assert first_plan == second_plan
    digest = first_plan["resourceManifestDigest"]
    assert digest.startswith("sha256:")
    assert len(digest) == len("sha256:") + hashlib.sha256().digest_size * 2
    plan_json = json.dumps(first_plan, ensure_ascii=False)
    assert first.preflight.resources[0].source_sha256 not in plan_json
    assert first.preflight.resources[0].payload_sha256 not in plan_json


def test_document_design_selects_default_table_style() -> None:
    markdown = """---
title: Business table
design: business
---

# Chapter

:::table {#tab:business caption="Business default"}
| Name | Value |
|---|---|
| A | 1 |
:::
"""
    build = build_longform_generation(markdown)
    table = next(
        operation
        for operation in build.plan.to_dict()["operations"]
        if operation["op"] == "writer.add_semantic_table"
    )
    assert table["args"]["style"] == "grid"
    assert table["args"]["borderSpec"]["insideVertical"] == 0.75


def test_abstract_and_list_references_reach_native_plan_operations() -> None:
    markdown = """---
title: Nested references
---

:::abstract
Abstract sees {{ref:fig:native}}.
:::

# Chapter

:::figure {#fig:native caption="Native target"}
![missing](missing.png)
:::

- First sees {{ref:fig:native}}.
- Plain companion.

1. Ordered sees {{ref:fig:native}}.
2. Ordered plain companion.
"""
    build = build_longform_generation(markdown)
    validate_generation_plan(build.plan.to_dict(), component="writer")
    references = [
        operation
        for operation in build.plan.to_dict()["operations"]
        if operation["op"] == "writer.add_cross_reference"
    ]
    by_node = {operation["nodeId"]: operation for operation in references}
    assert "__wpsc_para:0:1" in by_node
    assert "__wpsc_para:1:1" in by_node
    assert "__wpsc_para:1:2" in by_node
    assert "__wpsc_para:1:3" in by_node
    assert "__wpsc_para:1:4" in by_node
    for node_id in ("__wpsc_para:0:1", "__wpsc_para:1:1"):
        runs = by_node[node_id]["args"]["runs"]
        assert [run["type"] for run in runs].count("reference") == 1
        assert next(run for run in runs if run["type"] == "reference")[
            "targetNodeId"
        ] == "fig:native"
    assert by_node["__wpsc_para:1:1"]["args"]["listFormatting"] == {
        "kind": "bullet",
        "indentPt": 24.0,
    }
    assert by_node["__wpsc_para:1:1"]["args"]["runs"][0] == {
        "type": "text",
        "text": "•\t",
    }
    assert by_node["__wpsc_para:1:2"]["args"]["listFormatting"]["kind"] == "bullet"
    assert by_node["__wpsc_para:1:2"]["args"]["runs"][0]["text"] == "•\t"
    assert by_node["__wpsc_para:1:3"]["args"]["listFormatting"]["kind"] == "ordered"
    assert by_node["__wpsc_para:1:3"]["args"]["runs"][0]["text"] == "1.\t"
    assert by_node["__wpsc_para:1:4"]["args"]["runs"][0]["text"] == "2.\t"
