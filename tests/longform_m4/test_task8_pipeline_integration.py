"""M4 Task 8 offline integration, cleanup, and compatibility contracts."""

from __future__ import annotations

import base64
import gc
import json
import os
import subprocess
import sys
import weakref
from pathlib import Path
from typing import Optional

import pytest
from PIL import Image

from skills.WPSComposer.scripts.longform.executor import ExecutionOutcome
from skills.WPSComposer.scripts.longform.pipeline import (
    build_longform_generation,
    execute_longform_plan,
)
from skills.WPSComposer.scripts.longform.resources import PreparedLongformResource
from tests.longform_m2.test_acceptance_m2 import (
    SNAPSHOTS_DIR as M2_SNAPSHOTS,
    build_fixture as build_m2_fixture,
    structural_snapshot as m2_structural_snapshot,
)
from tests.longform_m3.test_acceptance_m3 import (
    NAMES as M3_NAMES,
    SNAPSHOTS as M3_SNAPSHOTS,
    _build as build_m3_fixture,
    _canonical_operations as m3_canonical_operations,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _write_png(path: Path) -> bytes:
    Image.new("RGB", (4, 3), (22, 44, 88)).save(path, format="PNG")
    return path.read_bytes()


def _formula_build(tmp_path: Path):
    return build_longform_generation(
        """---
title: M4 private formula
---

# Body

:::equation {#eq:private fallback_image="formula-private.png"}
x^2 + y_1
:::
""",
        base_dir=str(tmp_path),
    )


def test_offline_formula_build_never_starts_wps_or_writes_files(tmp_path: Path) -> None:
    _write_png(tmp_path / "formula-private.png")
    script = r"""
import json
import os
import sys

project_root, base_dir = sys.argv[1:]
sys.path.insert(0, project_root)

write_flags = os.O_WRONLY | os.O_RDWR | os.O_CREAT | os.O_TRUNC | os.O_APPEND

def audit(event, args):
    if event == "open":
        mode = args[1] if len(args) > 1 else None
        flags = args[2] if len(args) > 2 else 0
        if (isinstance(mode, str) and any(mark in mode for mark in "wax+")) or (
            isinstance(flags, int) and flags & write_flags
        ):
            raise RuntimeError("OFFLINE_WRITE_ATTEMPT")
    if event in {"subprocess.Popen", "os.system", "os.posix_spawn"}:
        raise RuntimeError("OFFLINE_PROCESS_ATTEMPT")

sys.addaudithook(audit)
from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation

build = build_longform_generation(
    '''---
title: M4 private formula
---

# Body

:::equation {#eq:private fallback_image="formula-private.png"}
x^2 + y_1
:::
''',
    base_dir=base_dir,
)
equation = next(
    operation
    for operation in build.plan.to_dict()["operations"]
    if operation["op"] == "writer.add_equation"
)
forbidden = (
    "skills.WPSComposer.scripts.writer",
    "skills.WPSComposer.scripts._dispatch",
    "skills.WPSComposer.scripts.macos_probe",
    "skills.WPSComposer.scripts.longform.windows_executor",
    "skills.WPSComposer.scripts.longform.macos_executor",
    "win32com",
    "pythoncom",
)
print(json.dumps({
    "forbiddenModules": [
        name for name in sys.modules if name.startswith(forbidden)
    ],
    "renderMode": equation["args"]["renderMode"],
    "fallbackResource": equation["args"]["fallbackResource"],
    "resourceId": build.preflight.resources[0].resource_id,
}, sort_keys=True))
"""
    environment = dict(os.environ)
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    completed = subprocess.run(
        [sys.executable, "-c", script, str(PROJECT_ROOT), str(tmp_path)],
        cwd=str(PROJECT_ROOT),
        env=environment,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    result = json.loads(completed.stdout)
    assert result["forbiddenModules"] == []
    assert result["renderMode"] == "native-m4"
    assert result["fallbackResource"] == {
        "fallbackResourceId": result["resourceId"]
    }


class _FormulaResourceExecutor:
    def __init__(self, failure: Optional[BaseException] = None) -> None:
        self.failure = failure
        self.deadline = None
        self.references: tuple[
            weakref.ReferenceType[PreparedLongformResource], ...
        ] = ()

    def execute(self, plan, resources, deadline=None):
        self.deadline = deadline
        self.references = tuple(weakref.ref(resource) for resource in resources)
        assert len(resources) == 1
        assert resources[0].payload_bytes.startswith(b"\x89PNG\r\n\x1a\n")
        if self.failure is not None:
            raise self.failure
        return ExecutionOutcome(staged_artifact="private-staged.docx")


@pytest.mark.parametrize(
    "failure",
    (None, RuntimeError("closed execution error"), TimeoutError("closed timeout")),
    ids=("success", "error", "timeout"),
)
def test_formula_private_transport_is_released_on_every_exit(
    tmp_path: Path, failure: Optional[BaseException]
) -> None:
    _write_png(tmp_path / "formula-private.png")
    build = _formula_build(tmp_path)
    executor = _FormulaResourceExecutor(failure)

    if failure is None:
        execute_longform_plan(build, executor, deadline=17.0)
    else:
        with pytest.raises(type(failure)):
            execute_longform_plan(build, executor, deadline=17.0)

    gc.collect()
    assert executor.deadline == 17.0
    assert executor.references
    assert all(reference() is None for reference in executor.references)


def test_formula_diagnostics_are_private_and_path_free(tmp_path: Path) -> None:
    payload = _write_png(tmp_path / "formula-private.png")
    build = _formula_build(tmp_path)
    resource = build.preflight.resources[0]
    diagnostic = json.dumps(build.to_json(), ensure_ascii=False, sort_keys=True)

    for forbidden in (
        str(tmp_path),
        "formula-private.png",
        resource.source_sha256,
        resource.payload_sha256,
        base64.b64encode(payload).decode("ascii"),
        "payload_bytes",
        "formula_bindings",
    ):
        assert forbidden not in diagnostic


@pytest.mark.parametrize(
    "name",
    ("academic", "wide_figure", "hybrid_bid", "degradation", "plain_short", "protocol_edge"),
)
def test_m1_m2_structural_snapshots_remain_exact(name: str) -> None:
    actual = json.dumps(
        m2_structural_snapshot(build_m2_fixture(name).plan),
        ensure_ascii=False,
        indent=2,
    ) + "\n"
    assert actual.encode("utf-8") == (M2_SNAPSHOTS / f"{name}_ops.json").read_bytes()


@pytest.mark.parametrize("name", M3_NAMES)
def test_m3_plan_snapshots_remain_exact(name: str) -> None:
    assert m3_canonical_operations(build_m3_fixture(name)) == (
        M3_SNAPSHOTS / f"{name}.json"
    ).read_bytes()


@pytest.mark.parametrize("markdown", ("", "# Title\n\nPlain body only.\n"))
def test_absent_optional_content_keeps_the_quality_anchor_invisible(markdown: str) -> None:
    build = build_longform_generation(markdown)
    operations = build.plan.to_dict()["operations"]
    anchors = [
        operation
        for operation in operations
        if operation["op"] == "writer.reserve_document_quality_anchor"
    ]

    assert build.issues == ()
    assert len(anchors) == 1
    assert anchors[0]["args"]["notices"] == []
    assert not [
        operation
        for operation in operations
        if "degradation" in operation["op"].lower()
        or "notice" in operation["op"].lower()
    ]


def test_public_docs_describe_the_closed_m4_boundary() -> None:
    skill = (PROJECT_ROOT / "skills/WPSComposer/SKILL.md").read_text(encoding="utf-8")
    api = (PROJECT_ROOT / "skills/WPSComposer/references/api.md").read_text(
        encoding="utf-8"
    )
    required = (
        "M4",
        "fallback_image",
        "bibliography_include_uncited",
        "EQUATION_INSERT_FAILED",
        "ENGINE_LOST",
        "M5",
        "macOS",
    )
    for document in (skill, api):
        for token in required:
            assert token in document
