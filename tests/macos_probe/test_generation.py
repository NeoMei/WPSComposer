from __future__ import annotations

from pathlib import Path
import shutil
from types import SimpleNamespace
import zipfile
from xml.etree import ElementTree

import pytest

from skills.WPSComposer.scripts.artifact_transport import ArtifactValidationError
from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    RecordedGeneration,
)
from skills.WPSComposer.scripts.macos_probe import generation as mac_generation
from skills.WPSComposer.scripts.macos_probe.generation import (
    GenerationError,
    GenerationRequest,
    execute_feasibility_plan,
)
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


MARKER = "IN_PLACE_MARKER"
WRITER_MARKER_PLAN = GenerationPlan(
    "writer",
    (
        GenerationOperation("writer.reset", {}),
        GenerationOperation(
            "writer.add_paragraph",
            {"text": MARKER, "style": "Body Text"},
        ),
    ),
)
SHEET_MARKER_PLAN = GenerationPlan(
    "spreadsheet",
    (
        GenerationOperation("sheet.reset", {}),
        GenerationOperation(
            "sheet.write_table",
            {"startRow": 1, "startCol": 1, "values": [[MARKER]]},
        ),
    ),
)
SLIDE_MARKER_PLAN = GenerationPlan(
    "presentation",
    (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": MARKER}),
    ),
)
MARKER_PLANS = {
    "writer": ("docx", WRITER_MARKER_PLAN),
    "spreadsheet": ("xlsx", SHEET_MARKER_PLAN),
    "presentation": ("pptx", SLIDE_MARKER_PLAN),
}
CONTENT_MEMBERS = {
    "docx": "word/document.xml",
    "xlsx": "xl/sharedStrings.xml",
    "pptx": "ppt/slides/slide1.xml",
}
VISIBLE_TEXT_QNAMES = {
    "docx": (
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    ),
    "xlsx": (
        "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
    ),
    "pptx": (
        "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
    ),
}


def _rewrite_member(path: Path, member: str, transform) -> None:
    with zipfile.ZipFile(path) as source:
        contents = {
            name: source.read(name) for name in source.namelist()
        }
    contents[member] = transform(contents[member])
    with zipfile.ZipFile(path, "w") as destination:
        for name, data in contents.items():
            destination.writestr(name, data)


def _write_semantic_package(path: Path, members: dict[str, str]) -> Path:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        for name, content in members.items():
            package.writestr(name, content)
    return path


def _presentation_members(
    slide_texts: list[str],
    *,
    order: list[str] | None = None,
    targets: dict[str, str] | None = None,
) -> dict[str, str]:
    presentation = (
        "http://schemas.openxmlformats.org/presentationml/2006/main"
    )
    relationships = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    package_relationships = (
        "http://schemas.openxmlformats.org/package/2006/relationships"
    )
    drawing = "http://schemas.openxmlformats.org/drawingml/2006/main"
    order = order or [f"rId{index}" for index in range(1, len(slide_texts) + 1)]
    targets = targets or {
        f"rId{index}": f"slides/slide{index}.xml"
        for index in range(1, len(slide_texts) + 1)
    }
    slide_ids = "".join(
        f'<p:sldId id="{255 + index}" r:id="{relationship_id}"/>'
        for index, relationship_id in enumerate(order, start=1)
    )
    relation_nodes = "".join(
        '<Relationship '
        f'Id="{relationship_id}" '
        f'Type="{relationships}/slide" Target="{target}"/>'
        for relationship_id, target in targets.items()
    )
    members = {
        "ppt/presentation.xml": (
            f'<p:presentation xmlns:p="{presentation}" xmlns:r="{relationships}">'
            f"<p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>"
        ),
        "ppt/_rels/presentation.xml.rels": (
            f'<Relationships xmlns="{package_relationships}">'
            f"{relation_nodes}</Relationships>"
        ),
    }
    for index, text in enumerate(slide_texts, start=1):
        members[f"ppt/slides/slide{index}.xml"] = (
            f'<p:sld xmlns:p="{presentation}" xmlns:a="{drawing}">'
            f"<a:t>{text}</a:t></p:sld>"
        )
    return members


def _xml_with_root_marker(data: bytes) -> bytes:
    root = ElementTree.fromstring(data)
    root.text = (root.text or "") + MARKER
    return ElementTree.tostring(
        root, encoding="utf-8", xml_declaration=True
    )


def _xml_with_visible_marker(data: bytes, format_name: str) -> bytes:
    root = ElementTree.fromstring(data)
    qname = VISIBLE_TEXT_QNAMES[format_name]
    if format_name == "xlsx":
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        shared_string = next(root.iter(f"{{{namespace}}}si"))
        for child in list(shared_string):
            shared_string.remove(child)
        rich_text = ElementTree.SubElement(shared_string, f"{{{namespace}}}r")
        text = ElementTree.SubElement(rich_text, qname)
        text.text = MARKER
    else:
        text = next(root.iter(qname))
        text.text = (text.text or "") + MARKER
    return ElementTree.tostring(
        root, encoding="utf-8", xml_declaration=True
    )


def _add_marker(path: Path, format_name: str) -> None:
    _rewrite_member(
        path,
        CONTENT_MEMBERS[format_name],
        lambda data: _xml_with_visible_marker(data, format_name),
    )


def _add_root_marker(path: Path, format_name: str) -> None:
    _rewrite_member(
        path, CONTENT_MEMBERS[format_name], _xml_with_root_marker
    )


def _invalid_feasibility_plans():
    for component, (format_name, plan) in MARKER_PLANS.items():
        reset, marker = plan.operations
        yield component, format_name, "rogue", GenerationPlan(
            component,
            (
                reset,
                GenerationOperation("rogue.eval", {}),
            ),
        )
        yield component, format_name, "missing-reset", GenerationPlan(
            component, (marker,)
        )
        wrong_args = marker.to_dict()["args"]
        if component == "writer":
            wrong_args["text"] = "WRONG_MARKER"
        elif component == "spreadsheet":
            wrong_args["values"] = [["WRONG_MARKER"]]
        else:
            wrong_args["title"] = "WRONG_MARKER"
        yield component, format_name, "wrong-marker", GenerationPlan(
            component,
            (reset, GenerationOperation(marker.op, wrong_args)),
        )
        yield component, format_name, "wrong-order", GenerationPlan(
            component, (marker, reset)
        )
        yield component, format_name, "extra-operation", GenerationPlan(
            component, (reset, marker, reset)
        )


class FakeBridge:
    def __init__(
        self,
        *,
        result_error=None,
        error_code="GENERATION_COMMAND_FAILED",
        returned_path=None,
        applied_operations=2,
        write_marker=True,
        marker_writer=_add_marker,
    ):
        self.url = "http://127.0.0.1:45678"
        self.token = "token"
        self.result_error = result_error
        self.error_code = error_code
        self.returned_path = returned_path
        self.applied_operations = applied_operations
        self.write_marker = write_marker
        self.marker_writer = marker_writer
        self.commands = []
        self.registrations = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return None

    def wait_registered(self, expected, timeout):
        self.registrations.append((set(expected), timeout))

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
        command = next(item for item in self.commands if item.id == command_id)
        staged = Path(command.params["stagedPath"])
        if self.result_error is not None:
            message = self.result_error.replace("{staging}", str(staged.parent))
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": self.error_code, "message": message},
            )
        if self.write_marker:
            self.marker_writer(staged, command.params["formatName"])
        returned = staged if self.returned_path is None else self.returned_path
        return ProbeResult(
            command_id,
            True,
            {
                "path": str(returned),
                "appliedOperations": self.applied_operations,
            },
            None,
        )


class FakeRuntime:
    def __init__(self, staging_dir: Path, calls: list):
        self.staging_dir = staging_dir.resolve()
        self.calls = calls
        self.registration_restored = True

    def __enter__(self):
        self.staging_dir.mkdir(parents=True)
        self.calls.append(("enter", self.staging_dir))
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.calls.append(("close",))
        shutil.rmtree(self.staging_dir, ignore_errors=True)

    def prepare_profiles(self):
        self.calls.append(("prepare_profiles",))

    def start_servers(self, *, deadline):
        self.calls.append(("start_servers",))

    def activate_component(self, component, *, deadline):
        self.calls.append(("activate_component", component))


def _run_with_fakes(
    tmp_path: Path,
    component: str,
    format_name: str,
    plan: GenerationPlan,
    *,
    bridge: FakeBridge | None = None,
    timeout: float = 2,
):
    calls = []
    fake_bridge = bridge or FakeBridge()
    runtime = FakeRuntime(tmp_path / "container" / "session-1", calls)
    request = GenerationRequest(
        output=tmp_path / f"final.{format_name}",
        component=component,
        format_name=format_name,
        overwrite=False,
    )
    result = execute_feasibility_plan(
        request,
        RecordedGeneration(plan, ()),
        enabled={format_name: True},
        bridge_factory=lambda origins: fake_bridge,
        runtime_factory=lambda *args, **kwargs: runtime,
        timeout=timeout,
    )
    return result, request, fake_bridge, runtime, calls


@pytest.mark.parametrize(
    ("component", "format_name", "method", "plan"),
    [
        ("writer", "docx", "generate_writer_document", WRITER_MARKER_PLAN),
        (
            "spreadsheet",
            "xlsx",
            "generate_spreadsheet_workbook",
            SHEET_MARKER_PLAN,
        ),
        (
            "presentation",
            "pptx",
            "generate_presentation_deck",
            SLIDE_MARKER_PLAN,
        ),
    ],
)
def test_generation_uses_only_staged_path_and_publishes_valid_package(
    tmp_path: Path,
    component: str,
    format_name: str,
    method: str,
    plan: GenerationPlan,
):
    result, request, bridge, runtime, calls = _run_with_fakes(
        tmp_path, component, format_name, plan
    )

    command = bridge.commands[0]
    assert command.component == component
    assert command.method == method
    assert set(command.params) == {"stagedPath", "formatName", "plan", "resources"}
    assert Path(command.params["stagedPath"]).parent == runtime.staging_dir
    assert command.params["formatName"] == format_name
    assert command.params["plan"] == plan.to_dict()
    assert command.params["resources"] == {}
    assert "output" not in command.params
    assert result == request.output.resolve()
    assert result.is_file()
    assert not runtime.staging_dir.exists()
    assert ("activate_component", component) in calls


def test_generation_production_gates_now_enabled(tmp_path: Path):
    # All four formats were verified to serialize on macOS WPS 12.1.26035 via
    # the gated JSAPI backend, so the production gates are open. The earlier
    # NO-GO decision in docs/macos-phase0.md no longer reproduces.
    assert mac_generation.MACOS_GENERATION_ENABLED == {
        "docx": True,
        "xlsx": True,
        "pptx": True,
        "pdf": True,
    }
    request = GenerationRequest(
        tmp_path / "final.docx", "writer", "docx", False
    )
    # The gate logic itself still fires when a format is explicitly disabled.
    with pytest.raises(GenerationError) as caught:
        execute_feasibility_plan(
            request, RecordedGeneration(WRITER_MARKER_PLAN, ()),
            enabled={"docx": False},
        )
    assert caught.value.code == "MACOS_GENERATION_GATE_NOT_PASSED"
    assert caught.value.backend == "mac-wps-jsapi"


def test_generation_rejects_component_format_and_plan_mismatch(tmp_path: Path):
    request = GenerationRequest(
        tmp_path / "final.xlsx", "spreadsheet", "xlsx", False
    )
    with pytest.raises(GenerationError) as caught:
        execute_feasibility_plan(
            request,
            RecordedGeneration(WRITER_MARKER_PLAN, ()),
            enabled={"xlsx": True},
        )
    assert caught.value.code == "OPERATION_PLAN_INVALID"


def test_generation_rejects_host_resources_in_feasibility_backend(tmp_path: Path):
    resource_path = tmp_path / "image.png"
    resource_path.write_bytes(b"not staged")
    resource = GenerationResource("image-1", resource_path, "image/png")
    request = GenerationRequest(
        tmp_path / "final.docx", "writer", "docx", False
    )
    with pytest.raises(GenerationError) as caught:
        execute_feasibility_plan(
            request,
            RecordedGeneration(WRITER_MARKER_PLAN, (resource,)),
            enabled={"docx": True},
        )
    assert caught.value.code == "OPERATION_PLAN_INVALID"


@pytest.mark.parametrize(
    ("component", "format_name", "case", "plan"),
    list(_invalid_feasibility_plans()),
)
def test_generation_rejects_non_feasibility_plan_before_runtime_creation(
    tmp_path: Path,
    component: str,
    format_name: str,
    case: str,
    plan: GenerationPlan,
):
    factory_calls = []

    def bridge_factory(origins):
        factory_calls.append(("bridge", origins))
        raise AssertionError("bridge must not be created for an invalid plan")

    def runtime_factory(*args, **kwargs):
        factory_calls.append(("runtime", args, kwargs))
        raise AssertionError("runtime must not be created for an invalid plan")

    request = GenerationRequest(
        tmp_path / f"final.{format_name}",
        component,
        format_name,
        False,
    )
    with pytest.raises(GenerationError) as caught:
        execute_feasibility_plan(
            request,
            RecordedGeneration(plan, ()),
            enabled={format_name: True},
            bridge_factory=bridge_factory,
            runtime_factory=runtime_factory,
        )

    assert caught.value.code == "OPERATION_PLAN_INVALID", case
    assert factory_calls == [], case
    assert not request.output.exists()


def _copy_template(tmp_path: Path, format_name: str) -> Path:
    filename = {
        "docx": "wpsDemo.docx",
        "xlsx": "etDemo.xlsx",
        "pptx": "wppDemo.pptx",
    }[format_name]
    source = (
        Path("macos/wps-jsapi-probe/node_modules/wpsjs/src/lib/res")
        / filename
    )
    target = tmp_path / f"package.{format_name}"
    shutil.copy2(source, target)
    return target


@pytest.mark.parametrize("format_name", ["docx", "xlsx", "pptx"])
def test_marker_validator_accepts_well_formed_marker_in_content_member(
    tmp_path: Path, format_name: str
):
    package = _copy_template(tmp_path, format_name)
    _add_marker(package, format_name)

    mac_generation._validate_marker_package(package, format_name)


def test_marker_validator_rejects_malformed_content_xml_with_marker_bytes(
    tmp_path: Path,
):
    package = _copy_template(tmp_path, "docx")
    _rewrite_member(
        package,
        CONTENT_MEMBERS["docx"],
        lambda data: b"<document>IN_PLACE_MARKER",
    )

    with pytest.raises(ArtifactValidationError, match="Invalid DOCX XML"):
        mac_generation._validate_marker_package(package, "docx")


def test_marker_validator_rejects_marker_only_in_unrelated_xml(tmp_path: Path):
    package = _copy_template(tmp_path, "docx")
    _rewrite_member(package, "docProps/core.xml", _xml_with_root_marker)

    with pytest.raises(ArtifactValidationError, match="missing IN_PLACE_MARKER"):
        mac_generation._validate_marker_package(package, "docx")


def test_marker_validator_rejects_malformed_unrelated_xml(
    tmp_path: Path,
):
    package = _copy_template(tmp_path, "docx")
    _add_marker(package, "docx")
    _rewrite_member(
        package,
        "docProps/core.xml",
        lambda data: b"<broken>",
    )

    with pytest.raises(ArtifactValidationError, match="Invalid DOCX XML"):
        mac_generation._validate_marker_package(package, "docx")


def test_marker_validator_rejects_content_member_without_marker(tmp_path: Path):
    package = _copy_template(tmp_path, "docx")

    with pytest.raises(ArtifactValidationError, match="missing IN_PLACE_MARKER"):
        mac_generation._validate_marker_package(package, "docx")


@pytest.mark.parametrize(
    ("component", "format_name", "plan"),
    [
        ("writer", "docx", WRITER_MARKER_PLAN),
        ("spreadsheet", "xlsx", SHEET_MARKER_PLAN),
        ("presentation", "pptx", SLIDE_MARKER_PLAN),
    ],
)
def test_generation_rejects_marker_only_in_content_root_text_without_publishing(
    tmp_path: Path,
    component: str,
    format_name: str,
    plan: GenerationPlan,
):
    bridge = FakeBridge(marker_writer=_add_root_marker)

    with pytest.raises(GenerationError) as caught:
        _run_with_fakes(
            tmp_path,
            component,
            format_name,
            plan,
            bridge=bridge,
            timeout=0.02,
        )

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not (tmp_path / f"final.{format_name}").exists()


@pytest.mark.parametrize(
    ("bridge", "expected_code"),
    [
        (
            FakeBridge(
                result_error="save failed at {staging}/generated.docx",
                error_code="GENERATION_COMMAND_FAILED",
            ),
            "GENERATION_COMMAND_FAILED",
        ),
        (
            FakeBridge(result_error="vendor failure", error_code="WPS_E_9001"),
            "GENERATION_COMMAND_FAILED",
        ),
        (FakeBridge(returned_path=Path("/tmp/outside.docx")), "PROTOCOL_ERROR"),
        (FakeBridge(applied_operations=1), "PROTOCOL_ERROR"),
        (FakeBridge(write_marker=False), "STAGED_ARTIFACT_INVALID"),
    ],
)
def test_generation_rejects_failed_or_invalid_results(
    tmp_path: Path, bridge: FakeBridge, expected_code: str
):
    with pytest.raises(GenerationError) as caught:
        _run_with_fakes(
            tmp_path,
            "writer",
            "docx",
            WRITER_MARKER_PLAN,
            bridge=bridge,
            timeout=0.02,
        )
    assert caught.value.code == expected_code
    assert not (tmp_path / "final.docx").exists()
    assert str((tmp_path / "container").resolve()) not in caught.value.message


def test_semantic_validator_counts_repeated_writer_content(tmp_path: Path):
    package = _write_semantic_package(tmp_path / "repeated.docx", {
        "word/document.xml": (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"><w:body><w:p><w:r>'
            '<w:t>REPEATED</w:t></w:r></w:p></w:body></w:document>'
        ),
    })
    plan = GenerationPlan("writer", (
        GenerationOperation("writer.reset", {}),
        GenerationOperation("writer.add_paragraph", {"text": "REPEATED"}),
        GenerationOperation("writer.add_paragraph", {"text": "REPEATED"}),
    ))

    with pytest.raises(ArtifactValidationError, match="content count"):
        mac_generation._validate_generated_package(
            package, "docx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_does_not_reuse_overlapping_writer_text(tmp_path: Path):
    package = _write_semantic_package(tmp_path / "overlap.docx", {
        "word/document.xml": (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"><w:body><w:p><w:r>'
            '<w:t>AA</w:t></w:r></w:p></w:body></w:document>'
        ),
    })
    plan = GenerationPlan("writer", (
        GenerationOperation("writer.reset", {}),
        GenerationOperation("writer.add_paragraph", {"text": "A"}),
        GenerationOperation("writer.add_paragraph", {"text": "AA"}),
    ))

    with pytest.raises(ArtifactValidationError, match="content count"):
        mac_generation._validate_generated_package(
            package, "docx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_requires_every_planned_writer_image(tmp_path: Path):
    package = _write_semantic_package(tmp_path / "images.docx", {
        "word/document.xml": (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.'
            'org/drawingml/2006/main"><w:body><w:drawing><a:blip/>'
            '</w:drawing></w:body></w:document>'
        ),
        "word/_rels/document.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="http://schemas.'
            'openxmlformats.org/officeDocument/2006/relationships/image" '
            'Target="media/image1.png"/></Relationships>'
        ),
    })
    plan = GenerationPlan("writer", (
        GenerationOperation("writer.reset", {}),
        GenerationOperation("writer.add_image", {"imageId": "image-1"}),
        GenerationOperation("writer.add_image", {"imageId": "image-2"}),
    ))

    with pytest.raises(ArtifactValidationError, match="image structure"):
        mac_generation._validate_generated_package(
            package, "docx", RecordedGeneration(plan, ()), "not-the-digest"
        )


@pytest.mark.parametrize(
    ("blip", "relationship", "media"),
    [
        ("<a:blip/>", "", {}),
        (
            '<a:blip r:embed="rId1"/>',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/hyperlink" '
            'Target="media/image1.png"/>',
            {"word/media/image1.png": b"PNG"},
        ),
        (
            '<a:blip r:embed="rId1"/>',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/image" '
            'Target="media/missing.png"/>',
            {},
        ),
    ],
)
def test_semantic_validator_rejects_unresolved_writer_blips(
    tmp_path: Path, blip: str, relationship: str, media: dict[str, bytes]
):
    members = {
        "word/document.xml": (
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.'
            'org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships"><w:body><w:drawing>'
            f"{blip}</w:drawing></w:body></w:document>"
        ),
        "word/_rels/document.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            f'2006/relationships">{relationship}</Relationships>'
        ),
        **media,
    }
    package = _write_semantic_package(tmp_path / "unresolved.docx", members)
    plan = GenerationPlan("writer", (
        GenerationOperation("writer.reset", {}),
        GenerationOperation("writer.add_image", {"imageId": "image-1"}),
    ))

    with pytest.raises(ArtifactValidationError, match="image structure"):
        mac_generation._validate_generated_package(
            package, "docx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_matches_presentation_images_to_logical_slide(
    tmp_path: Path,
):
    members = _presentation_members(["FIRST", "SECOND"])
    members["ppt/slides/slide2.xml"] = members["ppt/slides/slide2.xml"].replace(
        "</p:sld>",
        '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships" r:embed="rIdImage"/></p:sld>',
    )
    members["ppt/slides/_rels/slide2.xml.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships"><Relationship Id="rIdImage" Type="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="../media/image1.png"/></Relationships>'
    )
    members["ppt/media/image1.png"] = b"PNG"
    package = _write_semantic_package(tmp_path / "wrong-slide.pptx", members)
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": "FIRST"}),
        GenerationOperation("slide.add_title", {"title": "SECOND"}),
        GenerationOperation("slide.add_image", {
            "slide": 1,
            "imageId": "image-1",
            "left": 0,
            "top": 0,
        }),
    ))

    with pytest.raises(ArtifactValidationError, match="slide 1 image structure"):
        mac_generation._validate_generated_package(
            package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_aligns_presentation_content_by_slide(tmp_path: Path):
    package = _write_semantic_package(
        tmp_path / "swapped.pptx",
        _presentation_members(["SECOND", "FIRST"]),
    )
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": "FIRST"}),
        GenerationOperation("slide.add_title", {"title": "SECOND"}),
    ))

    with pytest.raises(ArtifactValidationError, match="slide 1 content"):
        mac_generation._validate_generated_package(
            package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_uses_presentation_relationship_playback_order(
    tmp_path: Path,
):
    package = _write_semantic_package(
        tmp_path / "relationship-order.pptx",
        _presentation_members(["FIRST", "SECOND"], order=["rId2", "rId1"]),
    )
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": "FIRST"}),
        GenerationOperation("slide.add_title", {"title": "SECOND"}),
    ))

    with pytest.raises(ArtifactValidationError, match="slide 1 content"):
        mac_generation._validate_generated_package(
            package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_accepts_logical_order_independent_of_part_name(
    tmp_path: Path,
):
    package = _write_semantic_package(
        tmp_path / "logical-order.pptx",
        _presentation_members(["SECOND", "FIRST"], order=["rId2", "rId1"]),
    )
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": "FIRST"}),
        GenerationOperation("slide.add_title", {"title": "SECOND"}),
    ))

    mac_generation._validate_generated_package(
        package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
    )


@pytest.mark.parametrize(
    ("order", "targets"),
    [
        (["rId1", "rId1"], {"rId1": "slides/slide1.xml"}),
        (["rId1", "rIdMissing"], {"rId1": "slides/slide1.xml"}),
        (["rId1", "rId2"], {
            "rId1": "slides/slide1.xml",
            "rId2": "slides/missing.xml",
        }),
        (["rId1", "rId2"], {
            "rId1": "slides/slide1.xml",
            "rId2": "../../outside.xml",
        }),
    ],
)
def test_semantic_validator_rejects_invalid_presentation_slide_relationships(
    tmp_path: Path, order: list[str], targets: dict[str, str]
):
    package = _write_semantic_package(
        tmp_path / "invalid-relationships.pptx",
        _presentation_members(["FIRST", "SECOND"], order=order, targets=targets),
    )
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": "FIRST"}),
        GenerationOperation("slide.add_title", {"title": "SECOND"}),
    ))

    with pytest.raises(ArtifactValidationError, match="slide relationship"):
        mac_generation._validate_generated_package(
            package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
        )


def test_semantic_validator_does_not_reuse_overlapping_slide_text(tmp_path: Path):
    package = _write_semantic_package(
        tmp_path / "overlap.pptx", _presentation_members(["AA"])
    )
    plan = GenerationPlan("presentation", (
        GenerationOperation("slide.reset", {}),
        GenerationOperation(
            "slide.add_title", {"title": "A", "subtitle": "AA"}
        ),
    ))

    with pytest.raises(ArtifactValidationError, match="slide 1 content"):
        mac_generation._validate_generated_package(
            package, "pptx", RecordedGeneration(plan, ()), "not-the-digest"
        )
