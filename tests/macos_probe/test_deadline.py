from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import artifact_transport
from skills.WPSComposer.scripts.generation_plan import (
    GenerationResource,
    GenerationOperation,
    GenerationPlan,
    RecordedGeneration,
)
from skills.WPSComposer.scripts.macos_probe import (
    conversion,
    generation,
    inspection,
    runtime,
)
from skills.WPSComposer.scripts.macos_probe.generation import GenerationRequest
from tests._pdf_fixture import write_minimal_pdf


class FakeClock:
    def __init__(self, now: float = 100.0):
        self.now = now

    def monotonic(self) -> float:
        return self.now

    def advance(self, seconds: float) -> None:
        self.now += seconds

    def sleep(self, seconds: float) -> None:
        self.advance(seconds)


def test_runtime_lock_uses_caller_absolute_deadline(monkeypatch, tmp_path: Path):
    clock = FakeClock()
    attempts = 0

    def blocked(*_args):
        nonlocal attempts
        attempts += 1
        raise BlockingIOError

    monkeypatch.setattr("fcntl.flock", blocked)
    monkeypatch.setattr(runtime.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(runtime.time, "sleep", clock.sleep)

    with pytest.raises(TimeoutError, match="another WPSComposer session"):
        with runtime.wps_runtime_lock(tmp_path / "runtime.lock", deadline=100.12):
            pass

    assert clock.now == pytest.approx(100.12)
    assert attempts >= 2


def test_three_server_readiness_checks_share_one_deadline(monkeypatch, tmp_path: Path):
    clock = FakeClock()
    calls: list[tuple[str, float]] = []

    class Server:
        def __init__(self, _profile, port):
            self.port = port

        def start(self):
            clock.advance(0.34)

        def close(self, timeout=0):
            pass

    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        publish_xml=tmp_path / "publish.xml",
        staging_root=tmp_path / "stable" / "WPSComposer",
    )
    for component in runtime.COMPONENT_CONFIG:
        profile = tmp_path / component
        profile.mkdir()
        probe.profiles[component] = profile

    monkeypatch.setattr(runtime, "StaticProfileServer", Server)
    monkeypatch.setattr(
        runtime.RegistrationSnapshot,
        "capture",
        lambda *args: SimpleNamespace(restore=lambda: None),
    )
    monkeypatch.setattr(runtime, "install_registration_entries", lambda *a, **k: None)
    monkeypatch.setattr(runtime.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(
        probe,
        "_wait_for_server",
        lambda component, port, deadline: calls.append((component, deadline)),
    )

    with pytest.raises(TimeoutError, match="deadline"):
        probe.start_servers(deadline=101.0)

    assert calls == [
        ("writer", 101.0),
        ("presentation", 101.0),
    ]


def test_activation_subprocess_receives_only_deadline_remaining(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock()
    probe_root = tmp_path / "probe"
    resource_dir = probe_root / "node_modules/wpsjs/src/lib/res"
    resource_dir.mkdir(parents=True)
    (resource_dir / "wpsDemo.docx").write_bytes(b"fixture")
    probe = runtime.ProbeRuntime(
        probe_root,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    probe.staging_dir = tmp_path / "staging"
    probe.staging_dir.mkdir()
    observed = []
    monkeypatch.setattr(runtime.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(
        runtime.subprocess,
        "run",
        lambda command, **kwargs: observed.append(kwargs["timeout"]),
    )

    probe.activate_component("writer", deadline=100.4)

    assert observed == [pytest.approx(0.4)]


def test_registration_retry_reuses_deadline_for_activation(monkeypatch):
    clock = FakeClock()
    waits = []
    activations = []

    class Bridge:
        def wait_registered(self, expected, timeout):
            waits.append(timeout)
            clock.advance(timeout)
            raise TimeoutError

    class Runtime:
        def activate_component(self, component, *, deadline):
            activations.append((component, deadline))

    monkeypatch.setattr(conversion.time, "monotonic", clock.monotonic)

    with pytest.raises(TimeoutError):
        conversion._wait_for_registration(
            Bridge(), Runtime(), "writer", deadline=100.25
        )

    assert waits == [pytest.approx(0.25), 0]
    assert activations == []


def test_public_conversion_deadline_precedes_bridge_and_reaches_runtime(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock()
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    request = SimpleNamespace(
        source=source.resolve(),
        output=(tmp_path / "result.pdf").resolve(),
        component="writer",
        overwrite=False,
    )
    seen = {}

    class Bridge:
        url = "http://127.0.0.1:45678"
        token = "token"

        def __enter__(self):
            clock.advance(0.2)
            return self

        def __exit__(self, *_args):
            pass

    class Runtime:
        registration_restored = True

        def __init__(self, *_args, deadline):
            seen["runtime_deadline"] = deadline

        def __enter__(self):
            clock.advance(0.2)
            return self

        def __exit__(self, *_args):
            pass

    monkeypatch.setattr(conversion.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(
        conversion,
        "_run_conversion",
        lambda request, method, bridge, runtime, deadline: seen.setdefault(
            "run_deadline", deadline
        ),
    )

    result = conversion.convert_macos(
        request,
        enabled=True,
        bridge_factory=lambda origins: Bridge(),
        runtime_factory=Runtime,
        timeout=1.0,
    )

    assert result == 101.0
    assert seen == {"runtime_deadline": 101.0, "run_deadline": 101.0}


def test_public_generation_deadline_precedes_bridge_and_reaches_runtime(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock()
    seen = {}
    request = GenerationRequest(tmp_path / "result.docx", "writer", "docx")
    recorded = RecordedGeneration(
        GenerationPlan("writer", (GenerationOperation("writer.reset", {}),)), ()
    )

    class Bridge:
        url = "http://127.0.0.1:45678"
        token = "token"

        def __enter__(self):
            clock.advance(0.2)
            return self

        def __exit__(self, *_args):
            pass

    class Runtime:
        registration_restored = True

        def __init__(self, *_args, deadline):
            seen["runtime_deadline"] = deadline

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            pass

    monkeypatch.setattr(generation.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(
        generation,
        "_run_generation",
        lambda request, recorded, method, bridge, runtime, probe_root,
        deadline, feasibility: seen.setdefault("run_deadline", deadline),
    )

    result = generation.execute_generation_plan(
        request,
        recorded,
        enabled={"docx": True},
        bridge_factory=lambda origins: Bridge(),
        runtime_factory=Runtime,
        timeout=1.0,
    )

    assert result == 101.0
    assert seen == {"runtime_deadline": 101.0, "run_deadline": 101.0}


@pytest.mark.parametrize("operation", ["inspect", "edit"])
def test_public_inspection_paths_create_deadline_before_bridge(
    monkeypatch, tmp_path: Path, operation: str
):
    clock = FakeClock()
    source = tmp_path / "source.pptx"
    source.write_bytes(b"source")
    seen = {}

    class Bridge:
        url = "http://127.0.0.1:45678"
        token = "token"

        def __enter__(self):
            clock.advance(0.2)
            return self

        def __exit__(self, *_args):
            pass

    class Runtime:
        registration_restored = True

        def __init__(self, *_args, deadline):
            seen["runtime_deadline"] = deadline

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            pass

    monkeypatch.setattr(inspection.time, "monotonic", clock.monotonic)
    if operation == "inspect":
        monkeypatch.setattr(
            inspection,
            "_run_inspection",
            lambda source, component, method, bridge, runtime, deadline,
            **kwargs: seen.setdefault("run_deadline", deadline),
        )
        result = inspection.inspect_macos(
            source,
            bridge_factory=lambda origins: Bridge(),
            runtime_factory=Runtime,
            timeout=1.0,
        )
    else:
        monkeypatch.setattr(
            inspection,
            "_run_edit",
            lambda source, output, component, method, patches, bridge, runtime,
            deadline, **kwargs: seen.setdefault("run_deadline", deadline),
        )
        result = inspection.edit_macos(
            source,
            [{"target": "slide:1", "name": "Updated"}],
            bridge_factory=lambda origins: Bridge(),
            runtime_factory=Runtime,
            timeout=1.0,
        )

    assert result == 101.0
    assert seen == {"runtime_deadline": 101.0, "run_deadline": 101.0}


def test_pdf_generation_and_final_validators_reuse_public_deadline(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock()
    deadline_calls = []
    staged = tmp_path / "staged.docx"
    staged_pdf = tmp_path / "staged.pdf"
    staged.write_bytes(b"template")
    staged_pdf.write_bytes(b"pdf")
    request = GenerationRequest(tmp_path / "result.pdf", "writer", "pdf")
    recorded = RecordedGeneration(
        GenerationPlan("writer", (GenerationOperation("writer.reset", {}),)), ()
    )

    class Runtime:
        staging_dir = tmp_path

        def prepare_profiles(self):
            pass

        def start_servers(self, *, deadline):
            deadline_calls.append(("servers", deadline))

        def activate_component(self, component, *, deadline):
            deadline_calls.append(("activation", deadline))

    class Bridge:
        state = SimpleNamespace(cancel=lambda command_id: None)

        def wait_registered(self, expected, timeout):
            pass

        def issue(self, component, method, params):
            return SimpleNamespace(id="command-1")

        def wait_result(self, command_id, timeout):
            return SimpleNamespace(
                ok=True,
                value={
                    "path": str(staged_pdf),
                    "sourcePath": str(staged),
                    "appliedOperations": 1,
                },
                error=None,
            )

    monkeypatch.setattr(generation.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(generation, "clone_template", lambda *args: staged)
    monkeypatch.setattr(generation, "_sha256", lambda path: "digest")
    monkeypatch.setattr(
        generation, "stage_generation_resources", lambda *args, **kwargs: {}
    )
    monkeypatch.setattr(
        generation,
        "_wait_for_generated_package",
        lambda *args, deadline: deadline_calls.append(("ooxml", deadline)),
    )
    monkeypatch.setattr(
        generation,
        "_wait_for_pdf",
        lambda path, *, deadline: deadline_calls.append(("pdf", deadline)),
    )

    def publish(_staged, destination, *, overwrite, validator, deadline):
        deadline_calls.append(("publish-deadline", deadline))
        deadline_calls.append(("publish-before", clock.now))
        validator(staged_pdf)
        return destination

    monkeypatch.setattr(generation, "publish_artifact", publish)
    monkeypatch.setattr(generation, "validate_pdf", lambda path: None)
    monkeypatch.setattr(
        generation,
        "validate_before_deadline",
        lambda validator, path, deadline: validator(path),
    )

    result = generation._run_generation(
        request,
        recorded,
        "generate_writer_document",
        Bridge(),
        Runtime(),
        tmp_path,
        deadline=101.0,
        feasibility=False,
    )

    assert result == request.output
    assert [value for _stage, value in deadline_calls[:5]] == [101.0] * 5


def test_pdf_validator_does_not_open_a_new_timeout_window(monkeypatch, tmp_path):
    clock = FakeClock(100.9)
    attempts = 0

    def invalid(_path):
        nonlocal attempts
        attempts += 1
        clock.advance(0.11)
        raise generation.ArtifactValidationError("not ready")

    monkeypatch.setattr(generation.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(generation.time, "sleep", clock.sleep)
    monkeypatch.setattr(generation, "validate_pdf", invalid)
    monkeypatch.setattr(
        generation,
        "validate_before_deadline",
        lambda validator, path, deadline: validator(path),
    )

    with pytest.raises(generation.ArtifactValidationError, match="not ready"):
        generation._wait_for_pdf(tmp_path / "artifact.pdf", deadline=101.0)

    assert attempts == 1
    assert clock.now == pytest.approx(101.01)


def test_conversion_pdf_readiness_uses_existing_deadline(monkeypatch, tmp_path):
    clock = FakeClock(100.9)
    attempts = 0

    def invalid(_path):
        nonlocal attempts
        attempts += 1
        clock.advance(0.11)
        raise conversion.ArtifactValidationError("not ready")

    monkeypatch.setattr(conversion.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(conversion.time, "sleep", clock.sleep)
    monkeypatch.setattr(conversion, "validate_pdf", invalid)
    monkeypatch.setattr(
        conversion,
        "validate_before_deadline",
        lambda validator, path, deadline: validator(path),
    )

    with pytest.raises(conversion.ArtifactValidationError, match="not ready"):
        conversion._wait_for_pdf_artifact(
            tmp_path / "artifact.pdf", deadline=101.0
        )

    assert attempts == 1
    assert clock.now == pytest.approx(101.01)


def test_conversion_end_to_end_consumes_one_cumulative_budget(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock()
    stages = []
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    request = SimpleNamespace(
        source=source.resolve(),
        output=(tmp_path / "result.pdf").resolve(),
        component="writer",
        overwrite=False,
    )

    def spend(stage: str, seconds: float, deadline: float) -> None:
        assert deadline == 101.0
        runtime.require_remaining(deadline)
        stages.append(stage)
        clock.advance(seconds)
        runtime.require_remaining(deadline)

    class Runtime:
        registration_restored = True

        def __init__(self, *_args, deadline):
            self.deadline = deadline
            self.staging_dir = tmp_path / "staging"

        def __enter__(self):
            self.staging_dir.mkdir()
            spend("lock", 0.06, self.deadline)
            return self

        def __exit__(self, *_args):
            stages.append("cleanup-grace")

        def prepare_profiles(self):
            spend("profiles", 0.03, self.deadline)

        def start_servers(self, *, deadline):
            for component in ("writer", "presentation", "spreadsheet"):
                spend(f"server:{component}", 0.04, deadline)

        def activate_component(self, component, *, deadline):
            spend("activation", 0.05, deadline)

    class Bridge:
        url = "http://127.0.0.1:45678"
        token = "token"
        state = SimpleNamespace(cancel=lambda command_id: None)

        def __init__(self):
            self.registration_attempts = 0
            self.command = None

        def __enter__(self):
            spend("bridge", 0.04, 101.0)
            return self

        def __exit__(self, *_args):
            pass

        def wait_registered(self, expected, timeout):
            self.registration_attempts += 1
            if self.registration_attempts == 1:
                spend("registration:retry", 0.05, 101.0)
                raise TimeoutError
            spend("registration:ready", 0.03, 101.0)

        def issue(self, component, method, params):
            spend("issue", 0.02, 101.0)
            self.command = SimpleNamespace(id="command-1", params=params)
            return self.command

        def wait_result(self, command_id, timeout):
            assert timeout == pytest.approx(101.0 - clock.now)
            spend("command", 0.08, 101.0)
            output = Path(self.command.params["outputPath"])
            write_minimal_pdf(output)
            return SimpleNamespace(ok=True, value={"path": str(output)}, error=None)

    original_validate_pdf = conversion.validate_pdf

    def validating_pdf(path):
        spend("pdf-validation", 0.02, 101.0)
        original_validate_pdf(path)

    monkeypatch.setattr(runtime.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(conversion.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(artifact_transport.time, "monotonic", clock.monotonic)
    monkeypatch.setattr(conversion, "validate_pdf", validating_pdf)
    monkeypatch.setattr(
        conversion,
        "validate_before_deadline",
        lambda validator, path, deadline: validator(path),
    )

    result = conversion.convert_macos(
        request,
        enabled=True,
        bridge_factory=lambda origins: Bridge(),
        runtime_factory=Runtime,
        timeout=1.0,
    )

    assert result == request.output
    assert clock.now < 101.0
    assert stages[:10] == [
        "bridge",
        "lock",
        "profiles",
        "server:writer",
        "server:presentation",
        "server:spreadsheet",
        "activation",
        "registration:retry",
        "activation",
        "registration:ready",
    ]
    assert stages.count("pdf-validation") == 4
    assert stages[-1] == "cleanup-grace"


def test_generation_resource_staging_stops_at_deadline_and_removes_partial(
    monkeypatch, tmp_path: Path
):
    clock = FakeClock(10.0)
    source = tmp_path / "large.png"
    source.write_bytes(b"x" * (3 * 1024 * 1024))
    request = GenerationRequest(tmp_path / "result.pptx", "presentation", "pptx")
    recorded = RecordedGeneration(
        GenerationPlan(
            "presentation",
            (
                GenerationOperation("slide.reset", {}),
                GenerationOperation("slide.add_blank", {}),
                GenerationOperation(
                    "slide.add_image", {"slide": 1, "imageId": "image-1"}
                ),
            ),
        ),
        (GenerationResource("image-1", source, "image/png"),),
    )
    staging = tmp_path / "staging"
    staging.mkdir()
    fake_runtime = SimpleNamespace(staging_dir=staging, profiles={})

    def ticking_clock():
        value = clock.now
        clock.advance(0.004)
        return value

    monkeypatch.setattr(runtime.time, "monotonic", ticking_clock)
    monkeypatch.setattr(artifact_transport.time, "monotonic", ticking_clock)

    with pytest.raises(TimeoutError):
        generation.stage_generation_resources(
            request, recorded, fake_runtime, deadline=10.02
        )

    assert source.stat().st_size == 3 * 1024 * 1024
    assert not list((staging / "resources").glob("*"))


@pytest.mark.parametrize("operation", ["conversion", "inspection", "edit"])
def test_source_staging_stops_at_deadline_and_preserves_files(
    monkeypatch, tmp_path: Path, operation: str
):
    clock = FakeClock(20.0)
    suffix = "docx" if operation != "edit" else "pptx"
    source = tmp_path / f"source.{suffix}"
    source_bytes = b"s" * (3 * 1024 * 1024)
    source.write_bytes(source_bytes)
    output = tmp_path / ("result.pdf" if operation == "conversion" else "result.pptx")
    output.write_bytes(b"approved-old-target")
    staging = tmp_path / "staging"
    staging.mkdir()

    class Runtime:
        staging_dir = staging

        def prepare_profiles(self):
            pass

        def start_servers(self, *, deadline):
            pass

        def activate_component(self, component, *, deadline):
            pass

    class Bridge:
        state = SimpleNamespace(cancel=lambda command_id: None)

        def __init__(self):
            self.commands = []

        def wait_registered(self, expected, timeout):
            pass

        def issue(self, component, method, params):
            self.commands.append(params)
            return SimpleNamespace(id="unexpected")

    bridge = Bridge()

    def ticking_clock():
        value = clock.now
        clock.advance(0.003)
        return value

    monkeypatch.setattr(runtime.time, "monotonic", ticking_clock)
    monkeypatch.setattr(artifact_transport.time, "monotonic", ticking_clock)

    with pytest.raises(TimeoutError):
        if operation == "conversion":
            request = SimpleNamespace(
                source=source,
                output=output,
                component="writer",
                overwrite=True,
            )
            conversion._run_conversion(
                request,
                "convert_writer_pdf",
                bridge,
                Runtime(),
                deadline=20.04,
            )
        elif operation == "inspection":
            inspection._run_inspection(
                source,
                "writer",
                "inspect_document",
                bridge,
                Runtime(),
                deadline=20.04,
            )
        else:
            inspection._run_edit(
                source,
                output,
                "presentation",
                "edit_presentation",
                [{"target": "slide:1", "name": "Updated"}],
                bridge,
                Runtime(),
                deadline=20.04,
                overwrite=True,
            )

    assert source.read_bytes() == source_bytes
    assert output.read_bytes() == b"approved-old-target"
    assert bridge.commands == []
    assert not list(staging.glob("source.*"))
