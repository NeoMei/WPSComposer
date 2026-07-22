"""Gated macOS WPS JSAPI generation through private native templates."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import re
import shutil
import stat
import tempfile
import time
from typing import Callable, Mapping, Optional
import zipfile
from xml.etree import ElementTree

from ..artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    publish_artifact,
    validate_office_package,
    validate_pdf,
)
from ..generation_plan import (
    GenerationResource,
    OperationPlanError,
    RecordedGeneration,
    validate_generation_plan,
)
from .bridge import LoopbackBridge
from .models import PathPolicy, ProtocolError
from .runtime import ProbeRuntime
from .templates import TemplateError, clone_template


ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
METHODS = {
    "writer": "generate_writer_document",
    "spreadsheet": "generate_spreadsheet_workbook",
    "presentation": "generate_presentation_deck",
}
FORMAT_COMPONENTS = {
    "docx": "writer",
    "pdf": "writer",
    "xlsx": "spreadsheet",
    "pptx": "presentation",
}
MAX_RESOURCE_BYTES = 50 * 1024 * 1024
RESOURCE_EXTENSIONS = {
    "image/bmp": frozenset({".bmp"}),
    "image/gif": frozenset({".gif"}),
    "image/jpeg": frozenset({".jpg", ".jpeg"}),
    "image/png": frozenset({".png"}),
    "image/tiff": frozenset({".tif", ".tiff"}),
}
RESOURCE_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
REMOTE_GENERATION_ERROR_CODES = frozenset(
    {
        "GENERATION_COMMAND_FAILED",
        "INTERACTIVE_INPUT_REQUIRED",
        "MACOS_CAPABILITY_UNAVAILABLE",
        "OPERATION_PLAN_INVALID",
    }
)
MACOS_GENERATION_ENABLED = {
    "docx": False,
    "xlsx": False,
    "pptx": False,
    "pdf": False,
}
FEASIBILITY_MARKER = "IN_PLACE_MARKER"
FEASIBILITY_PLANS = {
    "writer": {
        "component": "writer",
        "operations": [
            {"op": "writer.reset", "args": {}},
            {
                "op": "writer.add_paragraph",
                "args": {
                    "text": FEASIBILITY_MARKER,
                    "style": "Body Text",
                },
            },
        ],
    },
    "spreadsheet": {
        "component": "spreadsheet",
        "operations": [
            {"op": "sheet.reset", "args": {}},
            {
                "op": "sheet.write_table",
                "args": {
                    "startRow": 1,
                    "startCol": 1,
                    "values": [[FEASIBILITY_MARKER]],
                },
            },
        ],
    },
    "presentation": {
        "component": "presentation",
        "operations": [
            {"op": "slide.reset", "args": {}},
            {
                "op": "slide.add_title",
                "args": {"title": FEASIBILITY_MARKER},
            },
        ],
    },
}
MARKER_CONTENT_MEMBERS = {
    "docx": frozenset({"word/document.xml"}),
    "xlsx": frozenset({"xl/sharedStrings.xml"}),
    "pptx": frozenset({"ppt/slides/slide1.xml"}),
}
MARKER_TEXT_QNAMES = {
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


@dataclass(frozen=True)
class GenerationRequest:
    output: Path
    component: str
    format_name: str
    overwrite: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "output", Path(self.output).expanduser().resolve())


class GenerationError(RuntimeError):
    def __init__(
        self, *, code, output, component, backend, message
    ):
        super().__init__(message)
        self.code = code
        self.output = output
        self.component = component
        self.backend = backend
        self.message = message


def _error(
    request: GenerationRequest, code: str, message: str
) -> GenerationError:
    return GenerationError(
        code=code,
        output=str(request.output),
        component=request.component,
        backend="mac-wps-jsapi",
        message=message,
    )


def _normalize_remote_error_code(code: object) -> str:
    value = str(code) if isinstance(code, str) else ""
    if value in REMOTE_GENERATION_ERROR_CODES:
        return value
    return "GENERATION_COMMAND_FAILED"


def normalize_generation_error_code(code: object) -> str:
    """Return the stable public code for a vendor generation failure."""
    return _normalize_remote_error_code(code)


def _redact_staging(message: str, staging_dir: Path) -> str:
    return str(message).replace(str(staging_dir), "<wps-staging>")


def _validate_feasibility_recording(
    request: GenerationRequest, recorded: RecordedGeneration
) -> RecordedGeneration:
    try:
        validated = validate_generation_plan(
            recorded.plan.to_dict(), request.component
        )
    except (AttributeError, OperationPlanError, TypeError) as exc:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            str(exc) or "Generation plan is invalid",
        ) from exc
    if validated.to_dict() != FEASIBILITY_PLANS[request.component]:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Generation plan is not the exact in-place feasibility plan",
        )
    if recorded.resources:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Feasibility generation does not accept host resources",
        )
    return RecordedGeneration(validated, ())


def _validate_baseline_invariants(
    request: GenerationRequest, recorded: RecordedGeneration
) -> None:
    operations = recorded.plan.operations
    reset_name = {
        "writer": "writer.reset",
        "spreadsheet": "sheet.reset",
        "presentation": "slide.reset",
    }[request.component]
    if operations[0].op != reset_name or sum(
        operation.op == reset_name for operation in operations
    ) != 1:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            f"{request.component.capitalize()} generation requires one leading reset",
        )

    if request.component == "writer":
        for operation in operations:
            if operation.op != "writer.add_paragraph":
                continue
            spans = operation.args.get("spans")
            if spans is None:
                continue
            joined = "".join(span["text"] for span in spans)
            if joined != operation.args["text"]:
                raise _error(
                    request,
                    "OPERATION_PLAN_INVALID",
                    "Writer paragraph spans must concatenate to the paragraph text",
                )
    if request.component == "spreadsheet":
        sheet_count = 1
        for operation in operations:
            if operation.op == "sheet.add":
                sheet_count += 1
            elif operation.op in {"sheet.rename", "sheet.select"}:
                index = operation.args["index"]
                if index < 1 or index > sheet_count:
                    raise _error(
                        request,
                        "OPERATION_PLAN_INVALID",
                        "Spreadsheet operation references unavailable sheet state",
                    )
    elif request.component == "presentation":
        slide_count = 0
        add_operations = {
            "slide.add_title",
            "slide.add_section",
            "slide.add_bullets",
            "slide.add_blank",
        }
        for operation in operations:
            if operation.op in add_operations:
                slide_count += 1
            elif operation.op in {"slide.add_image", "slide.add_table"}:
                index = operation.args["slide"]
                if index < 1 or index > slide_count:
                    raise _error(
                        request,
                        "OPERATION_PLAN_INVALID",
                        "Presentation operation references unavailable slide state",
                    )


def _resource_ids_in_plan(recorded: RecordedGeneration) -> set[str]:
    return {
        str(operation.args["imageId"])
        for operation in recorded.plan.operations
        if operation.op.endswith(".add_image")
    }


def _validate_generation_resources(
    request: GenerationRequest, recorded: RecordedGeneration
) -> tuple[GenerationResource, ...]:
    resources = tuple(recorded.resources)
    seen: set[str] = set()
    for resource in resources:
        if not isinstance(resource, GenerationResource):
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                "Generation resource metadata is invalid",
            )
        if not RESOURCE_ID.fullmatch(resource.id) or resource.id in {".", ".."}:
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                "Generation resource identifier is invalid",
            )
        if resource.id in seen:
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                f"Duplicate generation resource: {resource.id}",
            )
        seen.add(resource.id)
        suffix = resource.source_path.suffix.lower()
        if suffix not in RESOURCE_EXTENSIONS.get(resource.media_type, frozenset()):
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                f"Generation resource {resource.id} has an invalid extension",
            )
        try:
            source_stat = resource.source_path.stat()
        except OSError as exc:
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                f"Generation resource {resource.id} is unavailable",
            ) from exc
        if not stat.S_ISREG(source_stat.st_mode):
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                f"Generation resource {resource.id} is not a regular file",
            )
        if source_stat.st_size > MAX_RESOURCE_BYTES:
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                f"Generation resource {resource.id} exceeds 50 MiB",
            )
    used = _resource_ids_in_plan(recorded)
    if seen != used:
        missing = sorted(used - seen)
        unused = sorted(seen - used)
        detail = "missing" if missing else "unused"
        identifiers = missing or unused
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            f"Generation plan has {detail} resources: {', '.join(identifiers)}",
        )
    if request.component == "spreadsheet" and resources:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Spreadsheet generation does not accept resources",
        )
    return resources


def _validate_production_recording(
    request: GenerationRequest, recorded: RecordedGeneration
) -> RecordedGeneration:
    try:
        validated = validate_generation_plan(
            recorded.plan.to_dict(), request.component
        )
    except (AttributeError, OperationPlanError, TypeError) as exc:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            str(exc) or "Generation plan is invalid",
        ) from exc
    normalized = RecordedGeneration(validated, tuple(recorded.resources))
    _validate_baseline_invariants(request, normalized)
    resources = _validate_generation_resources(request, normalized)
    return RecordedGeneration(validated, resources)


def _wait_for_registration(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    component: str,
    deadline: float,
) -> None:
    for attempt in range(4):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            bridge.wait_registered({component}, 0)
            return
        try:
            bridge.wait_registered({component}, min(10, remaining))
            return
        except TimeoutError:
            if attempt == 3:
                raise
            runtime.activate_component(component)


def _validate_marker_package(path: Path, format_name: str) -> None:
    validate_office_package(path, format_name)
    normalized = str(format_name).lower().lstrip(".")
    content_members = MARKER_CONTENT_MEMBERS[normalized]
    text_qname = MARKER_TEXT_QNAMES[normalized]
    found = False
    try:
        with zipfile.ZipFile(path) as package:
            for name in package.namelist():
                lowered = name.lower()
                if not (lowered.endswith(".xml") or lowered.endswith(".rels")):
                    continue
                root = ElementTree.fromstring(package.read(name))
                if name in content_members:
                    text = "".join(
                        element.text or ""
                        for element in root.iter(text_qname)
                    )
                    if FEASIBILITY_MARKER in text:
                        found = True
    except (
        ElementTree.ParseError,
        KeyError,
        OSError,
        RuntimeError,
        zipfile.BadZipFile,
    ) as exc:
        raise ArtifactValidationError(
            f"Invalid {normalized.upper()} XML package: {path}"
        ) from exc
    if not found:
        raise ArtifactValidationError(
            f"{normalized.upper()} package is missing {FEASIBILITY_MARKER} "
            f"in its generation content member: {path}"
        )


def _wait_for_marker(path: Path, format_name: str, timeout: float) -> None:
    deadline = time.monotonic() + timeout
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            _validate_marker_package(path, format_name)
            return
        except ArtifactValidationError as exc:
            failure = exc
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            assert failure is not None
            raise failure
        time.sleep(min(0.05, remaining))


def _wait_for_pdf(path: Path, timeout: float) -> None:
    deadline = time.monotonic() + timeout
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            validate_pdf(path)
            return
        except ArtifactValidationError as exc:
            failure = exc
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            assert failure is not None
            raise failure
        time.sleep(min(0.05, remaining))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except FileNotFoundError as exc:
        # async WPS flush replaced the file mid-read; let callers retry
        raise ArtifactValidationError(f"Artifact vanished: {path}") from exc
    return digest.hexdigest()


def _copy_generation_resource(
    request: GenerationRequest, resource: GenerationResource, target: Path
) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary: Optional[Path] = None
    descriptor: Optional[int] = None
    try:
        flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
        descriptor = os.open(resource.source_path, flags)
        source_stat = os.fstat(descriptor)
        if (
            not stat.S_ISREG(source_stat.st_mode)
            or source_stat.st_size > MAX_RESOURCE_BYTES
        ):
            raise OSError("resource changed during staging")
        with tempfile.NamedTemporaryFile(
            mode="wb",
            dir=target.parent,
            prefix=".wpscomposer-resource-",
            suffix=".tmp",
            delete=False,
        ) as outgoing:
            temporary = Path(outgoing.name)
            with os.fdopen(descriptor, "rb") as incoming:
                descriptor = None
                shutil.copyfileobj(incoming, outgoing)
            outgoing.flush()
            os.fsync(outgoing.fileno())
        if temporary.stat().st_size != source_stat.st_size:
            raise OSError("resource changed during staging")
        os.chmod(temporary, 0o600)
        os.link(temporary, target)
        temporary.unlink()
        temporary = None
    except (OSError, ValueError) as exc:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            f"Generation resource {resource.id} could not be staged",
        ) from exc
    finally:
        if descriptor is not None:
            os.close(descriptor)
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def stage_generation_resources(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    runtime: ProbeRuntime,
) -> dict[str, str]:
    """Copy validated host resources into component-private locations."""
    manifest: dict[str, str] = {}
    for resource in recorded.resources:
        safe_name = f"resource-{resource.id}{resource.source_path.suffix.lower()}"
        if request.component == "writer":
            try:
                target = runtime.profiles["writer"] / safe_name
            except KeyError as exc:
                raise _error(
                    request,
                    "GENERATION_COMMAND_FAILED",
                    "Writer resource profile is unavailable",
                ) from exc
            _copy_generation_resource(request, resource, target)
            manifest[resource.id] = f"http://127.0.0.1:3889/{safe_name}"
        elif request.component == "presentation":
            if runtime.staging_dir is None:
                raise _error(
                    request,
                    "GENERATION_COMMAND_FAILED",
                    "WPS staging session is unavailable",
                )
            target = runtime.staging_dir / "resources" / safe_name
            _copy_generation_resource(request, resource, target)
            manifest[resource.id] = str(target.resolve())
        else:
            raise _error(
                request,
                "OPERATION_PLAN_INVALID",
                "Spreadsheet generation does not accept resources",
            )
    return manifest


def _package_xml(path: Path) -> dict[str, ElementTree.Element]:
    parsed: dict[str, ElementTree.Element] = {}
    try:
        with zipfile.ZipFile(path) as package:
            for name in package.namelist():
                lowered = name.lower()
                if lowered.endswith(".xml") or lowered.endswith(".rels"):
                    parsed[name] = ElementTree.fromstring(package.read(name))
    except (
        ElementTree.ParseError,
        KeyError,
        OSError,
        RuntimeError,
        zipfile.BadZipFile,
    ) as exc:
        raise ArtifactValidationError("Generated Office package XML is invalid") from exc
    return parsed


def _expected_cell_text(value: object, component: str) -> Optional[str]:
    """Canonical visible text for a table cell, matching renderer semantics.

    Writers render cells with JS ``String(value)`` (null → ""), spreadsheets
    store typed values (bool → 1/0, not searchable text).
    """
    if value is None:
        return None
    if isinstance(value, bool):
        if component == "spreadsheet":
            return None
        return "true" if value else "false"
    if isinstance(value, float):
        if value.is_integer():
            value = int(value)
        elif "e" in str(value).lower():
            # spreadsheet XML serializes exponent floats differently
            # ("1.0E-05" vs Python "1e-05"); cannot substring-match reliably
            return None
    if isinstance(value, int) and abs(value) >= 10**15 and component == "spreadsheet":
        # Excel caps at 15 significant digits and switches to E-notation
        return None
    text = str(value)
    return text if text else None


def _expected_text(recorded: RecordedGeneration) -> list[str]:
    values: list[str] = []
    for operation in recorded.plan.operations:
        args = operation.args
        for key in ("text", "title", "subtitle"):
            value = args.get(key)
            if isinstance(value, str) and value:
                values.append(value)
        items = args.get("items")
        if isinstance(items, tuple):
            values.extend(str(item) for item in items if str(item))
        for key in ("data", "values"):
            rows = args.get(key)
            if isinstance(rows, tuple):
                for row in rows:
                    if isinstance(row, tuple):
                        values.extend(
                            text
                            for item in row
                            if (text := _expected_cell_text(
                                item, recorded.plan.component
                            ))
                            is not None
                        )
    return list(dict.fromkeys(values))


def _visible_package_text(
    xml: Mapping[str, ElementTree.Element], format_name: str
) -> str:
    if format_name == "docx":
        roots = [xml.get("word/document.xml")]
        qname = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    elif format_name == "xlsx":
        roots = [
            root
            for name, root in xml.items()
            if name == "xl/sharedStrings.xml"
            or name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        ]
        qname = None
    else:
        roots = [
            root
            for name, root in xml.items()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        ]
        qname = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
    chunks: list[str] = []
    for root in roots:
        if root is None:
            continue
        if qname is None:
            chunks.extend(text for text in root.itertext() if text)
        else:
            chunks.extend(element.text or "" for element in root.iter(qname))
    return " ".join(chunks)


def _image_relationship_count(
    xml: Mapping[str, ElementTree.Element], prefix: str
) -> int:
    relationship = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    count = 0
    for name, root in xml.items():
        if not name.startswith(prefix) or not name.endswith(".rels"):
            continue
        count += sum(
            str(element.attrib.get("Type", "")).endswith("/image")
            for element in root.iter(relationship)
        )
    return count


def _validate_generated_package(
    path: Path,
    format_name: str,
    recorded: RecordedGeneration,
    template_digest: str,
) -> None:
    validate_office_package(path, format_name)
    if _sha256(path) == template_digest:
        raise ArtifactValidationError("Generated artifact is an unchanged template")
    xml = _package_xml(path)
    visible_text = _visible_package_text(xml, format_name)
    expected = _expected_text(recorded)
    if expected and not all(value in visible_text for value in expected):
        raise ArtifactValidationError(
            "Generated artifact is missing representative renderer content"
        )
    if format_name == "docx":
        document = xml.get("word/document.xml")
        table_name = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"
        )
        actual_tables = (
            0 if document is None else sum(1 for _ in document.iter(table_name))
        )
        planned_tables = sum(
            operation.op == "writer.add_table"
            for operation in recorded.plan.operations
        )
        if actual_tables < planned_tables:
            raise ArtifactValidationError(
                "Generated document is missing planned table structure"
            )
        planned_images = any(
            operation.op == "writer.add_image"
            for operation in recorded.plan.operations
        )
        if planned_images and _image_relationship_count(xml, "word/_rels/") < 1:
            raise ArtifactValidationError(
                "Generated document is missing planned image structure"
            )
    elif format_name == "xlsx":
        workbook = xml.get("xl/workbook.xml")
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"
        actual = 0 if workbook is None else sum(1 for _ in workbook.iter(namespace))
        planned = 1 + sum(
            operation.op == "sheet.add" for operation in recorded.plan.operations
        )
        if actual != planned:
            raise ArtifactValidationError(
                "Generated workbook has an unexpected worksheet count"
            )
    elif format_name == "pptx":
        actual = sum(
            name.startswith("ppt/slides/slide") and name.endswith(".xml")
            for name in xml
        )
        planned = sum(
            operation.op
            in {
                "slide.add_title",
                "slide.add_section",
                "slide.add_bullets",
                "slide.add_blank",
            }
            for operation in recorded.plan.operations
        )
        if actual != planned:
            raise ArtifactValidationError(
                "Generated presentation has an unexpected slide count"
            )
        table_name = (
            "{http://schemas.openxmlformats.org/drawingml/2006/main}tbl"
        )
        actual_tables = sum(
            sum(1 for _ in root.iter(table_name))
            for name, root in xml.items()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        planned_tables = sum(
            operation.op == "slide.add_table"
            for operation in recorded.plan.operations
        )
        if actual_tables < planned_tables:
            raise ArtifactValidationError(
                "Generated presentation is missing planned table structure"
            )
        planned_images = any(
            operation.op == "slide.add_image"
            for operation in recorded.plan.operations
        )
        if planned_images and _image_relationship_count(
            xml, "ppt/slides/_rels/"
        ) < 1:
            raise ArtifactValidationError(
                "Generated presentation is missing planned image structure"
            )


def _wait_for_generated_package(
    path: Path,
    format_name: str,
    recorded: RecordedGeneration,
    template_digest: str,
    timeout: float,
) -> None:
    deadline = time.monotonic() + timeout
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            _validate_generated_package(
                path, format_name, recorded, template_digest
            )
            return
        except ArtifactValidationError as exc:
            failure = exc
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            assert failure is not None
            raise failure
        time.sleep(min(0.05, remaining))


def _run_generation(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    method: str,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    probe_root: Path,
    timeout: float,
    feasibility: bool,
) -> Path:
    is_pdf = request.format_name == "pdf"
    if runtime.staging_dir is None:
        raise _error(
            request,
            "STAGING_UNAVAILABLE",
            "WPS container staging session was not created",
        )
    runtime.prepare_profiles()
    policy = PathPolicy((runtime.staging_dir,))
    try:
        staged = policy.require_allowed(
            clone_template(probe_root, runtime.staging_dir, request.component)
        )
    except (ArtifactValidationError, OSError, TemplateError) as exc:
        raise _error(
            request,
            "STAGING_SAVE_FAILED",
            "Pinned WPS generation template could not be staged",
        ) from exc
    template_digest = _sha256(staged)
    staged_pdf = policy.require_allowed(staged.with_suffix(".pdf")) if is_pdf else None
    resources = stage_generation_resources(request, recorded, runtime)
    runtime.start_servers()
    runtime.activate_component(request.component)
    deadline = time.monotonic() + timeout
    try:
        _wait_for_registration(bridge, runtime, request.component, deadline)
    except TimeoutError as exc:
        raise _error(
            request,
            "GENERATION_COMMAND_FAILED",
            "Timed out waiting for the WPS add-in to register",
        ) from exc

    command_params = {
        "stagedPath": str(staged),
        "formatName": request.format_name,
        "plan": recorded.plan.to_dict(),
        "resources": resources,
    }
    if is_pdf:
        command_params["outputFormat"] = "pdf"
        command_params["stagedPdfPath"] = str(staged_pdf)
    command = bridge.issue(
        request.component,
        method,
        command_params,
    )
    try:
        result = bridge.wait_result(
            command.id, max(0.0, deadline - time.monotonic())
        )
    except TimeoutError as exc:
        bridge.state.cancel(command.id)
        raise _error(
            request,
            "GENERATION_COMMAND_FAILED",
            "Timed out waiting for WPS generation",
        ) from exc
    if not result.ok:
        details = dict(result.error or {})
        remote_code = _normalize_remote_error_code(details.get("code"))
        raise _error(
            request,
            remote_code,
            "WPS generation command failed",
        )
    expected_returned = staged_pdf if is_pdf else staged
    try:
        reported = policy.require_allowed(str(result.value.get("path", "")))
    except ProtocolError as exc:
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an invalid staged generation path",
        ) from exc
    if reported != expected_returned:
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an unexpected staged generation path",
        )
    if is_pdf:
        reported_source = result.value.get("sourcePath", "")
        try:
            source_path = policy.require_allowed(str(reported_source))
        except ProtocolError as exc:
            raise _error(
                request,
                "PROTOCOL_ERROR",
                "WPS returned an invalid private source path",
            ) from exc
        if source_path != staged:
            raise _error(
                request,
                "PROTOCOL_ERROR",
                "WPS returned an unexpected private source path",
            )
    applied = result.value.get("appliedOperations")
    if (
        isinstance(applied, bool)
        or not isinstance(applied, int)
        or applied != len(recorded.plan.operations)
    ):
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an unexpected applied operation count",
        )
    try:
        if feasibility:
            _wait_for_marker(staged, request.format_name, timeout)
        elif is_pdf:
            _wait_for_generated_package(
                staged,
                "docx",
                recorded,
                template_digest,
                timeout,
            )
            _wait_for_pdf(staged_pdf, timeout)
        else:
            _wait_for_generated_package(
                staged,
                request.format_name,
                recorded,
                template_digest,
                timeout,
            )
    except ArtifactValidationError as exc:
        raise _error(
            request,
            "STAGED_ARTIFACT_INVALID",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    try:
        if feasibility:
            validator = lambda path: _validate_marker_package(
                path, request.format_name
            )
        elif is_pdf:
            validator = lambda path: validate_pdf(path)
        else:
            validator = lambda path: _validate_generated_package(
                path,
                request.format_name,
                recorded,
                template_digest,
            )
        published = publish_artifact(
            staged_pdf if is_pdf else staged,
            request.output,
            overwrite=request.overwrite,
            validator=validator,
        )
        if is_pdf:
            staged.unlink(missing_ok=True)
            staged_pdf.unlink(missing_ok=True)
        return published
    except ArtifactTransportError as exc:
        if exc.code == "FINAL_ARTIFACT_INVALID":
            # the published artifact failed validation; never leave it behind
            # (with overwrite=True it has already replaced the user's file)
            request.output.unlink(missing_ok=True)
        message = {
            "STAGED_ARTIFACT_INVALID": "Staged WPS artifact is invalid",
            "ARTIFACT_PUBLISH_FAILED": "Generated artifact could not be published",
            "FINAL_ARTIFACT_INVALID": "Published generation artifact is invalid",
        }.get(exc.code, "Generated artifact transport failed")
        raise _error(
            request,
            exc.code,
            message,
        ) from exc


def _execute_generation_plan(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    *,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
    feasibility: bool = False,
) -> Path:
    """Execute one host-validated plan against a private template clone."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    gates = MACOS_GENERATION_ENABLED if enabled is None else enabled
    if not bool(gates.get(request.format_name, False)):
        raise _error(
            request,
            "MACOS_GENERATION_GATE_NOT_PASSED",
            "Mac WPS generation is disabled until the real acceptance gate passes",
        )
    expected_component = FORMAT_COMPONENTS.get(request.format_name)
    if expected_component is None:
        raise _error(
            request,
            "MACOS_CAPABILITY_UNAVAILABLE",
            "Requested Mac WPS generation format is unavailable",
        )
    if expected_component != request.component:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Generation component and format do not match",
        )
    if request.output.suffix.lower() != f".{request.format_name}":
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Output extension does not match the generation format",
        )
    if request.output.exists() and not request.overwrite:
        raise FileExistsError(f"Output already exists: {request.output}")
    recorded = _validate_production_recording(request, recorded)
    method = METHODS[request.component]

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-generate-")
    ).resolve()
    runtime = None
    try:
        with bridge_factory(ORIGINS) as bridge:
            runtime = runtime_factory(
                probe_root,
                runtime_root / "runtime",
                bridge.url,
                bridge.token,
            )
            try:
                with runtime:
                    return _run_generation(
                        request,
                        recorded,
                        method,
                        bridge,
                        runtime,
                        probe_root,
                        timeout,
                        feasibility,
                    )
            except GenerationError:
                raise
            except FileExistsError:
                raise
            except Exception as exc:
                if not getattr(runtime, "registration_restored", True):
                    raise _error(
                        request,
                        "REGISTRATION_RESTORE_FAILED",
                        "WPS registration restore failed; recovery evidence was retained",
                    ) from exc
                raise _error(
                    request,
                    "GENERATION_COMMAND_FAILED",
                    "Mac WPS generation command failed",
                ) from exc
    except (GenerationError, FileExistsError):
        raise
    except Exception as exc:
        if runtime is not None and not getattr(
            runtime, "registration_restored", True
        ):
            raise _error(
                request,
                "REGISTRATION_RESTORE_FAILED",
                "WPS registration restore failed; recovery evidence was retained",
            ) from exc
        raise _error(
            request,
            "GENERATION_COMMAND_FAILED",
            "Mac WPS generation command failed",
        ) from exc
    finally:
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(runtime_root, ignore_errors=True)


def execute_generation_plan(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    *,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> Path:
    """Execute a complete renderer plan through one private WPS runtime."""
    return _execute_generation_plan(
        request,
        recorded,
        enabled=enabled,
        bridge_factory=bridge_factory,
        runtime_factory=runtime_factory,
        timeout=timeout,
        feasibility=False,
    )


def execute_feasibility_plan(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    *,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> Path:
    """Retain Task 3's exact marker-only feasibility path without weakening it."""
    if request.format_name == "pdf":
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "The marker-only feasibility path does not support PDF output",
        )
    exact = _validate_feasibility_recording(request, recorded)
    return _execute_generation_plan(
        request,
        exact,
        enabled=enabled,
        bridge_factory=bridge_factory,
        runtime_factory=runtime_factory,
        timeout=timeout,
        feasibility=True,
    )


def generate_macos(
    doc,
    format_name: str,
    output: Path,
    preset,
    *,
    renderer_factory=None,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> Path:
    """Record a public renderer and generate through the gated macOS backend."""
    normalized_format = str(format_name).lower().lstrip(".")
    component = FORMAT_COMPONENTS.get(normalized_format, "writer")
    request = GenerationRequest(Path(output), component, normalized_format)
    gates = MACOS_GENERATION_ENABLED if enabled is None else enabled
    if not bool(gates.get(normalized_format, False)):
        raise _error(
            request,
            "MACOS_GENERATION_GATE_NOT_PASSED",
            "Mac WPS generation is disabled until the real acceptance gate passes",
        )
    if normalized_format not in FORMAT_COMPONENTS:
        raise _error(
            request,
            "MACOS_CAPABILITY_UNAVAILABLE",
            "Requested Mac WPS generation format is unavailable",
        )

    from ..recording_composers import (
        RecordingSheetComposer,
        RecordingSlideComposer,
        RecordingWriterComposer,
    )
    from ..renderers import sheet_renderer, slide_renderer, writer_renderer

    renderer, composer_factory = {
        "docx": (writer_renderer.render, RecordingWriterComposer),
        "pdf": (writer_renderer.render, RecordingWriterComposer),
        "xlsx": (sheet_renderer.render, RecordingSheetComposer),
        "pptx": (slide_renderer.render, RecordingSlideComposer),
    }[normalized_format]
    if renderer_factory is not None:
        renderer = renderer_factory
    try:
        recorded = renderer(
            doc,
            f"recording.{normalized_format}",
            preset=preset,
            composer_factory=composer_factory,
        )
    except GenerationError:
        raise
    except (OperationPlanError, TypeError, ValueError) as exc:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Renderer could not produce a valid generation plan",
        ) from exc
    if not isinstance(recorded, RecordedGeneration):
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Renderer did not return a recorded generation plan",
        )
    return execute_generation_plan(
        request,
        recorded,
        enabled=gates,
        bridge_factory=bridge_factory,
        runtime_factory=runtime_factory,
        timeout=timeout,
    )
