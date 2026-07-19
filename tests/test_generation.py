from __future__ import annotations

import builtins
import json
from pathlib import Path
import shutil
from types import SimpleNamespace
import zipfile
from xml.etree import ElementTree

import pytest

from skills.WPSComposer.scripts import orchestrator
from skills.WPSComposer.scripts.artifact_transport import ArtifactTransportError
from skills.WPSComposer.scripts.document_model import (
    ImageBlock,
    Paragraph,
    Section,
    Span,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    RecordedGeneration,
)
from skills.WPSComposer.scripts.macos_probe.generation import (
    GenerationError,
    GenerationRequest,
    execute_generation_plan,
    generate_macos,
    normalize_generation_error_code,
)
from skills.WPSComposer.scripts.macos_probe import generation as mac_generation
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


def _operation_text(plan):
    values = []
    for operation in plan["operations"]:
        args = operation["args"]
        for key in ("text", "title", "subtitle"):
            if isinstance(args.get(key), str) and args[key]:
                values.append(args[key])
        for key in ("items",):
            values.extend(item for item in args.get(key, []) if isinstance(item, str))
        for key in ("data", "values"):
            for row in args.get(key, []):
                values.extend(str(item) for item in row if item is not None)
    return " ".join(values) or "GENERATED_CONTENT"


def _rewrite_package(path, transform):
    with zipfile.ZipFile(path) as source:
        contents = {name: source.read(name) for name in source.namelist()}
    transform(contents)
    with zipfile.ZipFile(path, "w") as destination:
        for name, data in contents.items():
            destination.writestr(name, data)


def _write_semantic_artifact(path, format_name, plan, *, include_structure=True):
    text = _operation_text(plan)

    def transform(contents):
        if format_name == "docx":
            name = "word/document.xml"
            root = ElementTree.fromstring(contents[name])
            node = next(
                root.iter(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
                )
            )
            node.text = text
            if include_structure:
                body = root.find(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body"
                )
                for operation in plan["operations"]:
                    if operation["op"] == "writer.add_table":
                        ElementTree.SubElement(
                            body,
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl",
                        )
            contents[name] = ElementTree.tostring(root, encoding="utf-8")
            if include_structure and any(
                operation["op"] == "writer.add_image"
                for operation in plan["operations"]
            ):
                rel_name = "word/_rels/document.xml.rels"
                rel_root = ElementTree.fromstring(contents[rel_name])
                ElementTree.SubElement(
                    rel_root,
                    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                    {
                        "Id": "rIdWPSComposerImage",
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "Target": "media/wpscomposer.png",
                    },
                )
                contents[rel_name] = ElementTree.tostring(rel_root, encoding="utf-8")
                contents["word/media/wpscomposer.png"] = b"image"
        elif format_name == "xlsx":
            name = "xl/sharedStrings.xml"
            root = ElementTree.fromstring(contents[name])
            node = next(
                root.iter(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
                )
            )
            node.text = text
            contents[name] = ElementTree.tostring(root, encoding="utf-8")
        else:
            slide_ops = {
                "slide.add_title",
                "slide.add_section",
                "slide.add_bullets",
                "slide.add_blank",
            }
            count = sum(op["op"] in slide_ops for op in plan["operations"])
            original = contents["ppt/slides/slide1.xml"]
            for name in list(contents):
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    del contents[name]
            for index in range(1, count + 1):
                root = ElementTree.fromstring(original)
                node = next(
                    root.iter(
                        "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
                    )
                )
                node.text = text
                if include_structure and index == 1:
                    for operation in plan["operations"]:
                        if operation["op"] == "slide.add_table":
                            ElementTree.SubElement(
                                root,
                                "{http://schemas.openxmlformats.org/drawingml/2006/main}tbl",
                            )
                contents[f"ppt/slides/slide{index}.xml"] = ElementTree.tostring(
                    root, encoding="utf-8"
                )
            if include_structure and any(
                operation["op"] == "slide.add_image"
                for operation in plan["operations"]
            ):
                rel_name = "ppt/slides/_rels/slide1.xml.rels"
                if rel_name in contents:
                    rel_root = ElementTree.fromstring(contents[rel_name])
                else:
                    rel_root = ElementTree.Element(
                        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
                    )
                ElementTree.SubElement(
                    rel_root,
                    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                    {
                        "Id": "rIdWPSComposerImage",
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "Target": "../media/wpscomposer.png",
                    },
                )
                contents[rel_name] = ElementTree.tostring(rel_root, encoding="utf-8")
                contents["ppt/media/wpscomposer.png"] = b"image"

    _rewrite_package(path, transform)


class GenerationBridge:
    def __init__(self, mode="success"):
        self.url = "http://127.0.0.1:45678"
        self.token = "token"
        self.mode = mode
        self.commands = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return None

    def wait_registered(self, expected, timeout):
        return None

    def issue(self, component, method, params):
        command = SimpleNamespace(
            id=f"command-{len(self.commands) + 1}",
            component=component,
            method=method,
            params=dict(params),
        )
        self.commands.append(command)
        return command

    def wait_result(self, command_id, timeout):
        command = self.commands[-1]
        staged = Path(command.params["stagedPath"])
        if self.mode == "timeout":
            raise TimeoutError(f"private timeout at {staged}")
        if self.mode == "remote":
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": "WPS_E_PRIVATE", "message": f"secret vendor {staged}"},
            )
        if self.mode == "invalid":
            staged.write_bytes(b"not an office package")
        elif self.mode != "unchanged":
            plan = command.params["plan"]
            if self.mode == "partial":
                plan = {
                    "component": plan["component"],
                    "operations": plan["operations"][:2],
                }
            _write_semantic_artifact(
                staged,
                command.params["formatName"],
                plan,
                include_structure=self.mode != "no-structure",
            )
            if self.mode == "malformed":
                _rewrite_package(
                    staged,
                    lambda contents: contents.__setitem__(
                        "docProps/core.xml", b"<broken>"
                    ),
                )
        returned = staged
        if self.mode == "outside":
            returned = Path("/tmp/wpscomposer-outside.docx")
        count = len(command.params["plan"]["operations"])
        if self.mode == "count":
            count += 1
        return ProbeResult(
            command_id,
            True,
            {"path": str(returned), "appliedOperations": count},
            None,
        )


class GenerationRuntime:
    def __init__(self, runtime_dir, calls):
        self.runtime_dir = Path(runtime_dir).resolve()
        self.staging_dir = self.runtime_dir.parent / "session"
        self.profiles = {}
        self.calls = calls
        self.registration_restored = True

    def __enter__(self):
        self.staging_dir.mkdir(parents=True)
        self.calls.append("enter")
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.calls.append("close")
        shutil.rmtree(self.staging_dir, ignore_errors=True)

    def prepare_profiles(self):
        for component in ("writer", "spreadsheet", "presentation"):
            profile = self.runtime_dir / "profiles" / component
            profile.mkdir(parents=True, exist_ok=True)
            self.profiles[component] = profile
        self.calls.append("prepare_profiles")
        return dict(self.profiles)

    def start_servers(self):
        self.calls.append("start_servers")

    def activate_component(self, component):
        self.calls.append(("activate", component))


def _runtime_factory(captured):
    def factory(probe_root, runtime_dir, bridge_url, token):
        runtime = GenerationRuntime(runtime_dir, [])
        captured.append(runtime)
        return runtime

    return factory


def _writer_recording(text="Report"):
    return RecordedGeneration(
        GenerationPlan(
            "writer",
            (
                GenerationOperation("writer.reset", {}),
                GenerationOperation(
                    "writer.add_paragraph", {"text": text, "style": "Body Text"}
                ),
            ),
        ),
        (),
    )


def test_generate_routes_darwin_without_importing_pywin32(monkeypatch, tmp_path):
    calls = []
    real_import = builtins.__import__

    def guarded_import(name, *args, **kwargs):
        if name in {"pythoncom", "win32com", "win32com.client"}:
            raise AssertionError(f"Darwin route imported {name}")
        return real_import(name, *args, **kwargs)

    def fake_generate(doc, format_name, output, preset):
        calls.append((doc, format_name, output, preset))
        return output

    monkeypatch.setattr(builtins, "__import__", guarded_import)
    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(orchestrator, "generate_macos", fake_generate)

    output = tmp_path / "report.docx"
    result = orchestrator.generate(
        "# Report", format="docx", output=str(output), source_is_text=True
    )

    assert result == str(output.resolve())
    assert len(calls) == 1
    doc, format_name, routed_output, preset = calls[0]
    assert isinstance(doc, StructuredDocument)
    assert format_name == "docx"
    assert routed_output == output.resolve()
    assert preset is None


def test_generate_uses_resolved_safe_default_output(monkeypatch, tmp_path):
    calls = []
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(
        orchestrator,
        "generate_macos",
        lambda doc, format_name, output, preset: calls.append(output) or output,
    )

    result = orchestrator.generate(
        "# Report", format="pptx", source_is_text=True
    )

    assert result == str((tmp_path / "document.pptx").resolve())
    assert calls == [(tmp_path / "document.pptx").resolve()]


def test_generate_refuses_existing_output_before_backend(monkeypatch, tmp_path):
    output = tmp_path / "report.docx"
    output.write_bytes(b"keep")
    backend_called = False

    def fake_generate(*args):
        nonlocal backend_called
        backend_called = True
        raise AssertionError("backend must not run")

    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(orchestrator, "generate_macos", fake_generate)

    with pytest.raises(FileExistsError, match="Output already exists"):
        orchestrator.generate(
            "# Report", format="docx", output=str(output), source_is_text=True
        )

    assert not backend_called
    assert output.read_bytes() == b"keep"


def test_generate_rejects_output_extension_mismatch_before_backend(
    monkeypatch, tmp_path
):
    backend_called = False

    def fake_generate(*args):
        nonlocal backend_called
        backend_called = True

    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(orchestrator, "generate_macos", fake_generate)

    with pytest.raises(ValueError, match="extension"):
        orchestrator.generate(
            "# Report",
            format="docx",
            output=str(tmp_path / "report.pptx"),
            source_is_text=True,
        )

    assert not backend_called


def test_generate_rejects_unsupported_platform_before_renderer_import(
    monkeypatch, tmp_path
):
    real_import = builtins.__import__

    def guarded_import(name, *args, **kwargs):
        if ".renderers." in name or name.endswith(".writer"):
            raise AssertionError(f"unsupported route imported {name}")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", guarded_import)
    monkeypatch.setattr(orchestrator.sys, "platform", "linux")

    with pytest.raises(GenerationError) as caught:
        orchestrator.generate(
            "# Report",
            format="docx",
            output=str(tmp_path / "report.docx"),
            source_is_text=True,
        )

    assert caught.value.code == "MACOS_CAPABILITY_UNAVAILABLE"
    assert not (tmp_path / "report.docx").exists()


def test_generate_keeps_windows_writer_renderer_call_behavior(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.renderers import writer_renderer

    calls = []
    monkeypatch.setattr(orchestrator.sys, "platform", "win32")
    monkeypatch.setattr(
        writer_renderer,
        "render",
        lambda doc, output, preset=None: calls.append((doc, output, preset))
        or output,
    )
    monkeypatch.setattr(
        orchestrator,
        "generate_macos",
        lambda *args: (_ for _ in ()).throw(
            AssertionError("Windows route must not call the Mac backend")
        ),
    )

    output = tmp_path / "windows.docx"
    result = orchestrator.generate(
        "# Windows", format="docx", output=str(output), source_is_text=True
    )

    assert result == str(output.resolve())
    assert len(calls) == 1
    assert isinstance(calls[0][0], StructuredDocument)
    assert calls[0][1] == str(output.resolve())


def test_generate_macos_checks_production_gate_before_recording(tmp_path):
    renderer_called = False

    def forbidden_renderer(*args, **kwargs):
        nonlocal renderer_called
        renderer_called = True
        raise AssertionError("renderer must not run while gate is disabled")

    with pytest.raises(GenerationError) as caught:
        generate_macos(
            StructuredDocument(),
            "docx",
            tmp_path / "report.docx",
            None,
            renderer_factory=forbidden_renderer,
        )

    assert caught.value.code == "MACOS_GENERATION_GATE_NOT_PASSED"
    assert not renderer_called


def test_pdf_enabled_injection_still_fails_unsupported_before_runtime(tmp_path):
    bridge_called = False
    request = GenerationRequest(tmp_path / "report.pdf", "writer", "pdf")
    recorded = RecordedGeneration(
        GenerationPlan("writer", (GenerationOperation("writer.reset", {}),)),
        (),
    )

    def bridge_factory(origins):
        nonlocal bridge_called
        bridge_called = True
        raise AssertionError("Task 7 must not start WPS for PDF")

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            recorded,
            enabled={"pdf": True},
            bridge_factory=bridge_factory,
        )

    assert caught.value.code == "MACOS_CAPABILITY_UNAVAILABLE"
    assert not bridge_called
    assert not request.output.exists()


def test_unknown_vendor_generation_code_is_normalized():
    assert (
        normalize_generation_error_code("WPS_E_9001")
        == "GENERATION_COMMAND_FAILED"
    )
    assert normalize_generation_error_code(None) == "GENERATION_COMMAND_FAILED"


@pytest.mark.parametrize(
    ("format_name", "component", "method", "document"),
    [
        (
            "docx",
            "writer",
            "generate_writer_document",
            StructuredDocument(
                title="Report",
                sections=[
                    Section(1, "Summary", [Paragraph([Span("Writer body")])])
                ],
            ),
        ),
        (
            "xlsx",
            "spreadsheet",
            "generate_spreadsheet_workbook",
            StructuredDocument(
                sections=[
                    Section(
                        2,
                        "Data",
                        [TableBlock(["Item", "Amount"], [["A", "10"]])],
                    )
                ]
            ),
        ),
        (
            "pptx",
            "presentation",
            "generate_presentation_deck",
            StructuredDocument(
                title="Deck",
                sections=[
                    Section(2, "Finding", [Paragraph([Span("Slide body")])])
                ],
            ),
        ),
    ],
)
def test_generate_macos_records_executes_validates_and_publishes_all_formats(
    tmp_path, format_name, component, method, document
):
    bridge = GenerationBridge()
    runtimes = []
    output = tmp_path / f"final.{format_name}"

    result = generate_macos(
        document,
        format_name,
        output,
        None,
        enabled={format_name: True},
        bridge_factory=lambda origins: bridge,
        runtime_factory=_runtime_factory(runtimes),
        timeout=2,
    )

    assert result == output.resolve()
    assert output.is_file()
    assert len(bridge.commands) == 1
    command = bridge.commands[0]
    assert command.component == component
    assert command.method == method
    assert command.params["plan"]["component"] == component
    assert command.params["plan"]["operations"][0]["op"].endswith(".reset")
    assert command.params["resources"] == {}
    assert str(output.resolve()) not in json.dumps(command.params)
    assert runtimes[0].calls[-1] == "close"
    assert not runtimes[0].staging_dir.exists()


@pytest.mark.parametrize("format_name", ["docx", "pptx"])
def test_generate_macos_stages_resources_with_component_contract(
    tmp_path, format_name
):
    image = tmp_path / "figure.png"
    image.write_bytes(b"png resource bytes")
    if format_name == "docx":
        document = StructuredDocument(
            title="Image report",
            sections=[Section(1, "Image", [ImageBlock(str(image), "Figure")])],
        )
    else:
        document = StructuredDocument(
            title="Image deck",
            sections=[Section(2, "Image", [ImageBlock(str(image), "Figure")])],
        )
    bridge = GenerationBridge()
    runtimes = []

    generate_macos(
        document,
        format_name,
        tmp_path / f"final.{format_name}",
        None,
        enabled={format_name: True},
        bridge_factory=lambda origins: bridge,
        runtime_factory=_runtime_factory(runtimes),
        timeout=2,
    )

    resource = bridge.commands[0].params["resources"]["image-1"]
    if format_name == "docx":
        assert resource == "http://127.0.0.1:3889/resource-image-1.png"
    else:
        resource_path = Path(resource)
        assert resource_path.name == "resource-image-1.png"
        assert resource_path.parent.name == "resources"
        assert runtimes[0].staging_dir in resource_path.parents
    assert str(image.resolve()) not in json.dumps(bridge.commands[0].params)


@pytest.mark.parametrize(
    "recorded",
    [
        RecordedGeneration(
            GenerationPlan(
                "writer",
                (
                    GenerationOperation("writer.reset", {}),
                    GenerationOperation(
                        "writer.add_image", {"imageId": "image-1"}
                    ),
                ),
            ),
            (),
        ),
        RecordedGeneration(
            GenerationPlan(
                "writer",
                (
                    GenerationOperation("writer.reset", {}),
                    GenerationOperation(
                        "writer.add_paragraph", {"text": "Report"}
                    ),
                ),
            ),
            (),
        ),
    ],
)
def test_invalid_resource_usage_fails_before_runtime(tmp_path, recorded):
    source = tmp_path / "unused.png"
    source.write_bytes(b"unused")
    if not recorded.resources and recorded.plan.operations[-1].op != "writer.add_image":
        recorded = RecordedGeneration(
            recorded.plan,
            (GenerationResource("image-1", source, "image/png"),),
        )
    request = GenerationRequest(tmp_path / "final.docx", "writer", "docx")
    factory_called = False

    def bridge_factory(origins):
        nonlocal factory_called
        factory_called = True
        raise AssertionError("runtime must not start")

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            recorded,
            enabled={"docx": True},
            bridge_factory=bridge_factory,
        )

    assert caught.value.code == "OPERATION_PLAN_INVALID"
    assert not factory_called
    assert str(source.resolve()) not in caught.value.message


def test_generation_rejects_duplicate_mismatched_and_oversized_resources(
    tmp_path
):
    plan = GenerationPlan(
        "writer",
        (
            GenerationOperation("writer.reset", {}),
            GenerationOperation("writer.add_image", {"imageId": "image-1"}),
        ),
    )
    png = tmp_path / "figure.png"
    png.write_bytes(b"data")
    jpg = tmp_path / "figure.jpg"
    jpg.write_bytes(b"data")
    huge = tmp_path / "huge.png"
    with huge.open("wb") as stream:
        stream.truncate(50 * 1024 * 1024 + 1)
    missing = tmp_path / "missing.png"
    directory = tmp_path / "directory.png"
    directory.mkdir()
    cases = [
        (
            GenerationResource("image-1", png, "image/png"),
            GenerationResource("image-1", png, "image/png"),
        ),
        (GenerationResource("image-1", jpg, "image/png"),),
        (GenerationResource("image-1", huge, "image/png"),),
        (GenerationResource("image-1", missing, "image/png"),),
        (GenerationResource("image-1", directory, "image/png"),),
    ]

    for index, resources in enumerate(cases):
        request = GenerationRequest(
            tmp_path / f"final-{index}.docx", "writer", "docx"
        )
        with pytest.raises(GenerationError) as caught:
            execute_generation_plan(
                request,
                RecordedGeneration(plan, resources),
                enabled={"docx": True},
                bridge_factory=lambda origins: (_ for _ in ()).throw(
                    AssertionError("bridge must not start")
                ),
            )
        assert caught.value.code == "OPERATION_PLAN_INVALID"
        assert not request.output.exists()


@pytest.mark.parametrize(
    "plan",
    [
        GenerationPlan(
            "writer",
            (
                GenerationOperation("writer.add_paragraph", {"text": "x"}),
                GenerationOperation("writer.reset", {}),
            ),
        ),
        GenerationPlan(
            "spreadsheet",
            (
                GenerationOperation("sheet.reset", {}),
                GenerationOperation("sheet.select", {"index": 2}),
            ),
        ),
        GenerationPlan(
            "presentation",
            (
                GenerationOperation("slide.reset", {}),
                GenerationOperation("slide.reset", {}),
            ),
        ),
    ],
)
def test_generation_rejects_baseline_invariant_violations_before_runtime(
    tmp_path, plan
):
    format_name = {
        "writer": "docx",
        "spreadsheet": "xlsx",
        "presentation": "pptx",
    }[plan.component]
    request = GenerationRequest(
        tmp_path / f"final.{format_name}", plan.component, format_name
    )
    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            RecordedGeneration(plan, ()),
            enabled={format_name: True},
            bridge_factory=lambda origins: (_ for _ in ()).throw(
                AssertionError("bridge must not start")
            ),
        )
    assert caught.value.code == "OPERATION_PLAN_INVALID"


def test_execute_generation_refuses_existing_output_before_runtime(tmp_path):
    output = tmp_path / "final.docx"
    output.write_bytes(b"keep")
    request = GenerationRequest(output, "writer", "docx")

    with pytest.raises(FileExistsError, match="Output already exists"):
        execute_generation_plan(
            request,
            _writer_recording(),
            enabled={"docx": True},
            bridge_factory=lambda origins: (_ for _ in ()).throw(
                AssertionError("bridge must not start")
            ),
        )
    assert output.read_bytes() == b"keep"


@pytest.mark.parametrize(
    ("mode", "code"),
    [
        ("timeout", "GENERATION_COMMAND_FAILED"),
        ("remote", "GENERATION_COMMAND_FAILED"),
        ("outside", "PROTOCOL_ERROR"),
        ("count", "PROTOCOL_ERROR"),
        ("invalid", "STAGED_ARTIFACT_INVALID"),
        ("malformed", "STAGED_ARTIFACT_INVALID"),
        ("unchanged", "STAGED_ARTIFACT_INVALID"),
    ],
)
def test_production_generation_failures_cleanup_without_partial_output(
    tmp_path, mode, code
):
    bridge = GenerationBridge(mode)
    runtimes = []
    request = GenerationRequest(tmp_path / "final.docx", "writer", "docx")

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            _writer_recording("Expected generated text"),
            enabled={"docx": True},
            bridge_factory=lambda origins: bridge,
            runtime_factory=_runtime_factory(runtimes),
            timeout=0.02,
        )

    assert caught.value.code == code
    assert not request.output.exists()
    assert runtimes[0].calls[-1] == "close"
    assert not runtimes[0].staging_dir.exists()
    assert "secret vendor" not in caught.value.message
    assert "private timeout" not in caught.value.message
    assert "/tmp/wpscomposer-outside.docx" not in caught.value.message
    assert str(runtimes[0].staging_dir) not in caught.value.message


def test_artifact_transport_internal_path_is_redacted(monkeypatch, tmp_path):
    bridge = GenerationBridge()
    runtimes = []
    request = GenerationRequest(tmp_path / "final.docx", "writer", "docx")

    def fail_publish(*args, **kwargs):
        raise ArtifactTransportError(
            "ARTIFACT_PUBLISH_FAILED", "/private/internal/publish.tmp failed"
        )

    monkeypatch.setattr(mac_generation, "publish_artifact", fail_publish)

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            _writer_recording(),
            enabled={"docx": True},
            bridge_factory=lambda origins: bridge,
            runtime_factory=_runtime_factory(runtimes),
            timeout=2,
        )

    assert caught.value.code == "ARTIFACT_PUBLISH_FAILED"
    assert "/private/internal" not in caught.value.message
    assert not request.output.exists()


def test_semantic_validation_requires_all_representative_renderer_text(tmp_path):
    bridge = GenerationBridge("partial")
    request = GenerationRequest(tmp_path / "final.docx", "writer", "docx")
    recorded = RecordedGeneration(
        GenerationPlan(
            "writer",
            (
                GenerationOperation("writer.reset", {}),
                GenerationOperation("writer.add_paragraph", {"text": "First"}),
                GenerationOperation("writer.add_paragraph", {"text": "Second"}),
            ),
        ),
        (),
    )

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            recorded,
            enabled={"docx": True},
            bridge_factory=lambda origins: bridge,
            runtime_factory=_runtime_factory([]),
            timeout=0.02,
        )

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not request.output.exists()


def test_bridge_startup_failure_is_typed_and_redacted(tmp_path):
    request = GenerationRequest(tmp_path / "final.docx", "writer", "docx")

    def bridge_factory(origins):
        raise RuntimeError("secret bridge startup /private/internal")

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            _writer_recording(),
            enabled={"docx": True},
            bridge_factory=bridge_factory,
        )

    assert caught.value.code == "GENERATION_COMMAND_FAILED"
    assert "secret bridge" not in caught.value.message
    assert not request.output.exists()


@pytest.mark.parametrize("kind", ["table", "image"])
def test_writer_semantics_require_planned_structure(tmp_path, kind):
    operations = [GenerationOperation("writer.reset", {})]
    resources = ()
    if kind == "table":
        operations.append(
            GenerationOperation(
                "writer.add_table",
                {"rows": 1, "cols": 1, "data": [["Cell"]]},
            )
        )
    else:
        image = tmp_path / "figure.png"
        image.write_bytes(b"image")
        operations.extend(
            [
                GenerationOperation("writer.add_paragraph", {"text": "Image"}),
                GenerationOperation("writer.add_image", {"imageId": "image-1"}),
            ]
        )
        resources = (GenerationResource("image-1", image, "image/png"),)
    recorded = RecordedGeneration(
        GenerationPlan("writer", tuple(operations)), resources
    )
    request = GenerationRequest(tmp_path / f"{kind}.docx", "writer", "docx")

    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            recorded,
            enabled={"docx": True},
            bridge_factory=lambda origins: GenerationBridge("no-structure"),
            runtime_factory=_runtime_factory([]),
            timeout=0.02,
        )

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not request.output.exists()
