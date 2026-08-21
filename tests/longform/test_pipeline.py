"""Integration tests for the M1 offline long-form pipeline."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform.pipeline import (
    LongformBuild,
    build_longform_generation,
)


def _snapshot_plan(plan) -> bytes:
    return json.dumps(plan.to_dict(), ensure_ascii=False, separators=(",", ":")).encode(
        "utf-8"
    )


def _snapshot_semantic(semantic) -> bytes:
    return json.dumps(semantic.to_json(), ensure_ascii=False, separators=(",", ":")).encode(
        "utf-8"
    )


@pytest.fixture
def project_root() -> Path:
    return Path(__file__).parents[2]


def test_pipeline_import_is_pure_in_subprocess(project_root: Path):
    """Importing the pipeline module in a clean process must not load executors."""
    script = """
import sys
import skills.WPSComposer.scripts.longform.pipeline

forbidden = [
    "skills.WPSComposer.scripts.writer",
    "skills.WPSComposer.scripts.slide",
    "skills.WPSComposer.scripts.sheet",
    "skills.WPSComposer.scripts.wps_engine",
    "skills.WPSComposer.scripts.conversion",
    "skills.WPSComposer.scripts.macos_probe",
    "skills.WPSComposer.scripts._dispatch",
    "subprocess",
]
loaded = [m for m in sys.modules if any(m.startswith(f) for f in forbidden)]
if loaded:
    print(loaded)
    sys.exit(1)
print("pure")
"""
    env = {"PYTHONPATH": str(project_root)}
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        env=env,
        cwd=str(project_root),
    )
    assert result.returncode == 0, f"Import-time purity failed: {result.stdout} {result.stderr}"
    assert "pure" in result.stdout


def test_pipeline_module_does_not_import_platform_executors(project_root: Path):
    """The pipeline source must not directly import WPS/executor modules."""
    pipeline_path = project_root / "skills" / "WPSComposer" / "scripts" / "longform" / "pipeline.py"
    source = pipeline_path.read_text("utf-8")
    forbidden = (
        "from ..writer",
        "from ..slide",
        "from ..sheet",
        "from .._dispatch",
        "from ..conversion",
        "from ..macos_probe",
        "from ..windows_writer_worker",
        "import skills.WPSComposer.scripts.writer",
        "import skills.WPSComposer.scripts.slide",
        "import skills.WPSComposer.scripts.sheet",
        "import skills.WPSComposer.scripts._dispatch",
        "import skills.WPSComposer.scripts.conversion",
        "import skills.WPSComposer.scripts.macos_probe",
        "import skills.WPSComposer.scripts.windows_writer_worker",
    )
    for fragment in forbidden:
        assert fragment not in source, f"pipeline.py imports forbidden module via {fragment!r}"


def test_build_longform_generation_returns_longform_build():
    markdown = "# Hello\n\nWorld.\n"
    build = build_longform_generation(markdown)
    assert isinstance(build, LongformBuild)
    assert build.document.longform is True
    assert build.document.title == "Hello"
    assert build.semantic.config.title == "Hello"
    assert build.preflight is not None
    assert build.plan is not None
    assert build.plan.component == "writer"
    assert build.plan.protocol_version == 2


@pytest.mark.parametrize(
    "description,markdown",
    [
        ("empty input", ""),
        (
            "unclosed directive",
            "# T\n\n:::figure{caption=\"x\"}\n![alt](missing.png)\n",
        ),
        (
            "nested directives",
            "# T\n\n:::figure\n:::table\n|a|\n|--|\n|b|\n:::\n:::\n",
        ),
        (
            "invalid frontmatter values",
            "---\ntoc: [a, b]\n---\n# T\n",
        ),
        (
            "bad figure attributes",
            "# T\n\n:::figure{caption=has space without quotes}\n![alt](x.png)\n:::\n",
        ),
        (
            "control characters in directive",
            "# T\n\n:::figure{caption=\"\x01bad\"}\n![alt](x.png)\n:::\n",
        ),
        (
            "deep heading nesting",
            "# H1\n## H2\n### H3\n#### H4\n##### H5\n###### H6\nparagraph\n",
        ),
        (
            "path-traversal image path",
            "# T\n\n:::figure\n![alt](../traversal.png)\n:::\n",
        ),
    ],
)
def test_build_longform_generation_adversarial_inputs(description, markdown):
    """Adversarial inputs must not raise and must produce deterministic degradation issues."""
    first = build_longform_generation(markdown)
    second = build_longform_generation(markdown)
    assert isinstance(first, LongformBuild)
    assert isinstance(second, LongformBuild)
    assert _snapshot_semantic(first.semantic) == _snapshot_semantic(second.semantic)
    assert _snapshot_plan(first.plan) == _snapshot_plan(second.plan)


def test_build_longform_generation_missing_image_becomes_planned_degradation(tmp_path):
    missing_path = "missing.png"
    markdown = f"""---
title: Missing Image Test
---

# Introduction

:::figure {{caption="A missing figure"}}
![alt text]({missing_path})
:::
"""
    build = build_longform_generation(markdown, base_dir=str(tmp_path))
    ops = build.plan.to_dict()["operations"]
    figure_ops = [op for op in ops if op["op"] == "writer.add_captioned_figure"]
    assert len(figure_ops) == 1
    children = figure_ops[0]["args"]["children"]
    assert len(children) == 1
    assert "plannedDegradation" in children[0]
    assert children[0]["plannedDegradation"]["code"] == "RESOURCE_NOT_FOUND"
    assert children[0]["plannedDegradation"]["placement"] == "block"

    issue_codes = {issue.code for issue in build.issues}
    assert "RESOURCE_NOT_FOUND" in issue_codes


def test_build_longform_generation_semantic_and_plan_snapshots_are_stable(tmp_path):
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake-png-data")

    markdown = f"""---
title: Snapshot Test
author: Tester
toc: true
figure_index: true
---

# Chapter One

Some body text for the snapshot.\n\nRepeated to reach deterministic length.

:::figure {{caption="Stable figure", layout="stack"}}
![A sample image]({image_path.name})
:::

:::table {{caption="Stable table"}}
| Name | Value |
|------|-------|
| A    | 1     |
| B    | 2     |
:::

:::formula {{identifier="eq:stable"}}
E = mc^2
:::

## Section 1.1

A nested section to exercise heading numbering.

:::references {{identifier="ref:stable"}}
- id: chen2025 | text: 陈. 示例[J]. 2025.
:::
"""
    first = build_longform_generation(markdown, base_dir=str(tmp_path))
    second = build_longform_generation(markdown, base_dir=str(tmp_path))

    assert _snapshot_semantic(first.semantic) == _snapshot_semantic(second.semantic)
    assert _snapshot_plan(first.plan) == _snapshot_plan(second.plan)
    assert first.plan.to_dict()["resourceManifestDigest"] == second.plan.to_dict()[
        "resourceManifestDigest"
    ]


def test_build_longform_generation_issue_placement(tmp_path):
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake-png-data")

    markdown = f"""---
title: Placement Test
toc: not_a_boolean
figure_index: true
---

# Chapter

:::figure {{caption="Good figure"}}
![ok]({image_path.name})
:::

:::figure {{caption="Missing figure"}}
![missing](does_not_exist.png)
:::
"""
    build = build_longform_generation(markdown, base_dir=str(tmp_path))

    by_code = {issue.code: issue for issue in build.issues}
    assert "CONFIG_VALUE_INVALID" in by_code
    assert by_code["CONFIG_VALUE_INVALID"].placement == "document"
    assert "RESOURCE_NOT_FOUND" in by_code
    assert by_code["RESOURCE_NOT_FOUND"].placement == "block"


def test_build_longform_generation_writes_no_files_and_cwd_unchanged(tmp_path, monkeypatch):
    """The pipeline must not write files and must not modify the working directory."""
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"fake-png-data")

    markdown = f"""---
title: Side Effect Test
---

# Chapter

:::figure {{caption="Sample"}}
![sample]({image_path.name})
:::
"""
    before_tmp_files = {p for p in tmp_path.rglob("*") if p.is_file()}
    cwd_before = {p for p in Path.cwd().rglob("*") if p.is_file()}

    build = build_longform_generation(markdown, base_dir=str(tmp_path))
    assert build is not None

    after_tmp_files = {p for p in tmp_path.rglob("*") if p.is_file()}
    cwd_after = {p for p in Path.cwd().rglob("*") if p.is_file()}

    assert after_tmp_files == before_tmp_files, f"Pipeline wrote files in base_dir: {after_tmp_files - before_tmp_files}"
    assert cwd_after == cwd_before, f"Pipeline wrote files in cwd: {cwd_after - cwd_before}"


def test_build_longform_generation_empty_markdown():
    build = build_longform_generation("")
    assert isinstance(build, LongformBuild)
    assert build.plan.component == "writer"
