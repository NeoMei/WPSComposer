"""M2 Task 9: pipeline-to-executor integration and diagnostic hygiene.

Tests verify that:
- LongformBuild.to_json() redacts absolute source paths.
- A private staging transport map keeps resource source paths.
- execute_longform_plan binds a validated plan + resource manifest to an executor.
- Importing the pipeline module remains pure (no platform/WPS modules loaded).
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from dataclasses import replace
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationPlan,
    GenerationResource,
    OperationPlanError,
)
from skills.WPSComposer.scripts.longform.executor import (
    ExecutionOutcome,
    RecordingLongformExecutor,
)
from skills.WPSComposer.scripts.longform.pipeline import (
    LongformBuild,
    build_longform_generation,
)


@pytest.fixture
def project_root() -> Path:
    return Path(__file__).parents[2]


@pytest.fixture
def sample_build(tmp_path: Path) -> LongformBuild:
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"fake-png-data")
    markdown = f"""---
title: Task 9 Test
---

# Chapter

:::figure {{caption="Sample"}}
![sample]({image_path.name})
:::
"""
    return build_longform_generation(markdown, base_dir=str(tmp_path))


class TestDiagnosticHygiene:
    """LongformBuild.to_json must not leak sourcePath or base_dir."""

    def _assert_no_path_leak(self, build: LongformBuild, base_dir: str) -> None:
        data = build.to_json()
        json_text = json.dumps(data, ensure_ascii=False, sort_keys=True)
        assert "sourcePath" not in json_text
        assert base_dir not in json_text
        if _is_unix(base_dir):
            assert "/private" not in json_text or base_dir.startswith("/private")

    def test_to_json_has_no_source_path(self, sample_build: LongformBuild) -> None:
        self._assert_no_path_leak(sample_build, str(Path.cwd()))

    def test_to_json_resource_entry_has_required_fields(self, sample_build: LongformBuild) -> None:
        resources = sample_build.to_json()["preflight"]["resources"]
        assert resources
        for entry in resources:
            assert "sourcePath" not in entry
            assert "resourceId" in entry
            assert "sourceSha256" in entry
            assert "payloadSha256" in entry
            assert "byteLength" in entry
            assert "mediaType" in entry
            assert "normalizerId" in entry

    def test_to_json_degradation_has_no_source_path(self) -> None:
        # Force a missing resource degradation by referencing a non-existent file.
        markdown = """# T

:::figure
![missing](absent.png)
:::
"""
        build = build_longform_generation(markdown, base_dir="/tmp")
        self._assert_no_path_leak(build, "/tmp")

    def test_to_json_with_absolute_base_dir_redacts_everything(self, tmp_path: Path) -> None:
        abs_base = str(tmp_path.resolve())
        image_path = tmp_path / "sample.png"
        image_path.write_bytes(b"fake-png-data")
        # Use an absolute path to the image itself as well.
        markdown = f"""---
title: Absolute Path Test
---

# Chapter

:::figure {{caption="Sample"}}
![sample]({image_path})
:::
"""
        build = build_longform_generation(markdown, base_dir=abs_base)
        self._assert_no_path_leak(build, abs_base)

    def test_to_json_with_path_traversal_redacts_outside_paths(self, tmp_path: Path) -> None:
        abs_base = str(tmp_path.resolve())
        markdown = """# T

:::figure
![outside](../outside.png)
:::
"""
        build = build_longform_generation(markdown, base_dir=abs_base)
        data = build.to_json()
        json_text = json.dumps(data, ensure_ascii=False, sort_keys=True)
        assert "sourcePath" not in json_text
        assert abs_base not in json_text
        assert "../outside.png" not in json_text

    def test_private_transport_map_keeps_source_paths(self, sample_build: LongformBuild) -> None:
        transport = sample_build.resource_source_map()
        for resource in sample_build.preflight.resources:
            assert resource.resource_id in transport
            assert transport[resource.resource_id] == resource.source_path


class TestExecutorBinding:
    """execute_longform_plan binds a LongformBuild to a LongformExecutor."""

    def test_execute_longform_plan_exists_and_returns_execution_outcome(
        self, sample_build: LongformBuild
    ) -> None:
        from skills.WPSComposer.scripts.longform.pipeline import execute_longform_plan

        executor = RecordingLongformExecutor(artifact="task9-staged.docx")
        outcome = execute_longform_plan(sample_build, executor)
        assert isinstance(outcome, ExecutionOutcome)
        assert outcome.staged_artifact == "task9-staged.docx"

    def test_execute_longform_plan_passes_plan_and_resources(
        self, sample_build: LongformBuild
    ) -> None:
        from skills.WPSComposer.scripts.longform.pipeline import execute_longform_plan

        executor = RecordingLongformExecutor()
        execute_longform_plan(sample_build, executor, deadline=42.0)
        assert len(executor.calls) == 1
        plan, resources, deadline = executor.calls[0]
        assert plan is sample_build.plan
        assert len(resources) == len(sample_build.preflight.resources)
        for resource in resources:
            assert isinstance(resource, GenerationResource)
        assert deadline == 42.0

    def test_execute_longform_plan_resolves_resources_to_absolute_paths(
        self, tmp_path: Path
    ) -> None:
        from skills.WPSComposer.scripts.longform.pipeline import execute_longform_plan

        abs_base = str(tmp_path.resolve())
        image_path = tmp_path / "sample.png"
        image_path.write_bytes(b"fake-png-data")
        markdown = f"""---
title: Resolve Test
---

# Chapter

:::figure {{caption="Sample"}}
![sample](sample.png)
:::
"""
        build = build_longform_generation(markdown, base_dir=abs_base)
        executor = RecordingLongformExecutor()
        execute_longform_plan(build, executor)
        assert len(executor.calls) == 1
        _plan, resources, _deadline = executor.calls[0]
        assert len(resources) == 1
        assert Path(resources[0].source_path).is_absolute()

    def test_execute_longform_plan_validates_plan_before_executor(
        self, sample_build: LongformBuild
    ) -> None:
        from skills.WPSComposer.scripts.longform.pipeline import execute_longform_plan

        executor = RecordingLongformExecutor()
        # Validate that a corrupt plan raises before touching the executor.
        corrupt_build = replace(
            sample_build, plan=GenerationPlan("writer", ())
        )

        with pytest.raises(OperationPlanError):
            execute_longform_plan(corrupt_build, executor)
        assert len(executor.calls) == 0


class TestPipelineImportPurity:
    """The pipeline module must stay platform-pure at import time."""

    def test_pipeline_import_is_pure_in_subprocess(self, project_root: Path) -> None:
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
    "skills.WPSComposer.scripts.longform.windows_executor",
    "skills.WPSComposer.scripts.longform.macos_executor",
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

    def test_pipeline_source_does_not_import_platform_modules(self, project_root: Path) -> None:
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
            "from .windows_executor",
            "from .macos_executor",
            "import skills.WPSComposer.scripts.writer",
            "import skills.WPSComposer.scripts.slide",
            "import skills.WPSComposer.scripts.sheet",
            "import skills.WPSComposer.scripts._dispatch",
            "import skills.WPSComposer.scripts.conversion",
            "import skills.WPSComposer.scripts.macos_probe",
            "import skills.WPSComposer.scripts.windows_writer_worker",
            "import skills.WPSComposer.scripts.longform.windows_executor",
            "import skills.WPSComposer.scripts.longform.macos_executor",
            "import subprocess",
        )
        for fragment in forbidden:
            assert fragment not in source, f"pipeline.py imports forbidden module via {fragment!r}"


def _is_unix(path: str) -> bool:
    return os.name == "posix" and not path.startswith("\\")

