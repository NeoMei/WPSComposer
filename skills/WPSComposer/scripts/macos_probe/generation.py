"""Gated macOS WPS JSAPI generation through private native templates."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from functools import partial
import hashlib
import os
from pathlib import Path
import posixpath
import re
import shutil
import stat
import tempfile
import time
from typing import Callable, Mapping, Optional
from urllib.parse import unquote, urlsplit
import zipfile
from xml.etree import ElementTree

from ..artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    copy_stream_before_deadline,
    publish_artifact,
    validate_before_deadline,
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
from .runtime import ProbeRuntime, remaining, require_remaining
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
IMAGE_RELATIONSHIP_TYPES = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
})
SLIDE_RELATIONSHIP_TYPES = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slide",
})
REMOTE_GENERATION_ERROR_CODES = frozenset(
    {
        "GENERATION_COMMAND_FAILED",
        "INTERACTIVE_INPUT_REQUIRED",
        "MACOS_CAPABILITY_UNAVAILABLE",
        "OPERATION_PLAN_INVALID",
    }
)
MACOS_GENERATION_ENABLED = {
    # All four formats were verified to serialize on macOS WPS 12.1.26035 via
    # the gated JSAPI backend (docx/xlsx/pptx/pdf each produced a valid package
    # through generate_macos). Phase 0's Writer SaveAs2 failure no longer
    # reproduces; re-disable here only if the acceptance gate regresses.
    "docx": True,
    "xlsx": True,
    "pptx": True,
    "pdf": True,
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
        budget = remaining(deadline)
        if budget <= 0:
            bridge.wait_registered({component}, 0)
            return
        try:
            bridge.wait_registered({component}, min(10, budget))
            return
        except TimeoutError:
            if attempt == 3 or remaining(deadline) <= 0:
                bridge.wait_registered({component}, 0)
                raise
            runtime.activate_component(component, deadline=deadline)


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


def _wait_for_marker(
    path: Path, format_name: str, *, deadline: float
) -> None:
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            validate_before_deadline(
                partial(_validate_marker_package, format_name=format_name),
                path,
                deadline,
            )
            return
        except ArtifactValidationError as exc:
            failure = exc
        except TimeoutError:
            if failure is not None:
                raise failure
            raise
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            assert failure is not None
            raise failure
        time.sleep(min(0.05, remaining))


def _wait_for_pdf(path: Path, *, deadline: float) -> None:
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            validate_before_deadline(validate_pdf, path, deadline)
            return
        except ArtifactValidationError as exc:
            failure = exc
        except TimeoutError:
            if failure is not None:
                raise failure
            raise
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
    request: GenerationRequest,
    resource: GenerationResource,
    target: Path,
    *,
    deadline: float,
) -> None:
    require_remaining(deadline)
    target.parent.mkdir(parents=True, exist_ok=True)
    require_remaining(deadline)
    temporary: Optional[Path] = None
    descriptor: Optional[int] = None
    try:
        flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0)
        descriptor = os.open(resource.source_path, flags)
        source_stat = os.fstat(descriptor)
        require_remaining(deadline)
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
                copy_stream_before_deadline(incoming, outgoing, deadline)
            require_remaining(deadline)
            outgoing.flush()
            require_remaining(deadline)
            os.fsync(outgoing.fileno())
            require_remaining(deadline)
        if temporary.stat().st_size != source_stat.st_size:
            raise OSError("resource changed during staging")
        require_remaining(deadline)
        os.chmod(temporary, 0o600)
        require_remaining(deadline)
        os.link(temporary, target)
        require_remaining(deadline)
        temporary.unlink()
        temporary = None
    except TimeoutError:
        raise
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
    *,
    deadline: float,
) -> dict[str, str]:
    """Copy validated host resources into component-private locations."""
    manifest: dict[str, str] = {}
    for resource in recorded.resources:
        require_remaining(deadline)
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
            _copy_generation_resource(
                request, resource, target, deadline=deadline
            )
            manifest[resource.id] = f"http://127.0.0.1:3889/{safe_name}"
        elif request.component == "presentation":
            if runtime.staging_dir is None:
                raise _error(
                    request,
                    "GENERATION_COMMAND_FAILED",
                    "WPS staging session is unavailable",
                )
            target = runtime.staging_dir / "resources" / safe_name
            _copy_generation_resource(
                request, resource, target, deadline=deadline
            )
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


def _package_members(path: Path) -> set[str]:
    try:
        with zipfile.ZipFile(path) as package:
            return set(package.namelist())
    except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
        raise ArtifactValidationError(
            "Generated Office package members are invalid"
        ) from exc


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
        values.extend(
            _operation_expected_text(operation, recorded.plan.component)
        )
    return values


def _operation_expected_text(operation, component: str) -> list[str]:
    values: list[str] = []
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
                        if (text := _expected_cell_text(item, component)) is not None
                    )
    return values


def _relationship_part_name(source_part: str) -> str:
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _relationships_for_part(
    xml: Mapping[str, ElementTree.Element],
    source_part: str,
    context: str,
) -> dict[str, ElementTree.Element]:
    relationship_name = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    root = xml.get(_relationship_part_name(source_part))
    if root is None:
        raise ArtifactValidationError(
            f"Generated artifact is missing {context} relationships"
        )
    relationships = {}
    for relationship in root.iter(relationship_name):
        relationship_id = relationship.attrib.get("Id")
        if not relationship_id or relationship_id in relationships:
            raise ArtifactValidationError(
                f"Generated artifact has an invalid {context} relationship ID"
            )
        relationships[relationship_id] = relationship
    return relationships


def _relationship_target_member(
    source_part: str,
    relationship: ElementTree.Element,
    members: set[str],
    context: str,
    *,
    required_prefix: str,
) -> str:
    if relationship.attrib.get("TargetMode", "Internal").lower() == "external":
        raise ArtifactValidationError(
            f"Generated artifact has an external {context} relationship"
        )
    raw_target = relationship.attrib.get("Target", "").replace("\\", "/")
    parsed = urlsplit(raw_target)
    if not raw_target or parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise ArtifactValidationError(
            f"Generated artifact has an invalid {context} relationship target"
        )
    target_path = unquote(parsed.path)
    if target_path.startswith("/"):
        normalized = posixpath.normpath(target_path.lstrip("/"))
    else:
        normalized = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), target_path)
        )
    if (
        normalized in {"", ".", ".."}
        or normalized.startswith("../")
        or not normalized.startswith(required_prefix)
        or normalized not in members
    ):
        raise ArtifactValidationError(
            f"Generated artifact has an invalid {context} relationship target"
        )
    return normalized


def _valid_image_reference_count(
    xml: Mapping[str, ElementTree.Element],
    members: set[str],
    source_part: str,
    context: str,
    *,
    media_prefix: str,
) -> int:
    root = xml.get(source_part)
    if root is None:
        return 0
    blip_name = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    embed_name = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    )
    blips = list(root.iter(blip_name))
    if not blips:
        return 0
    structure_context = f"{context} image structure"
    relationships = _relationships_for_part(xml, source_part, structure_context)
    for blip in blips:
        relationship_id = blip.attrib.get(embed_name)
        relationship = relationships.get(relationship_id or "")
        if (
            not relationship_id
            or relationship is None
            or relationship.attrib.get("Type") not in IMAGE_RELATIONSHIP_TYPES
        ):
            raise ArtifactValidationError(
                f"Generated artifact has an invalid {structure_context}"
            )
        _relationship_target_member(
            source_part,
            relationship,
            members,
            structure_context,
            required_prefix=media_prefix,
        )
    return len(blips)


def _is_presentation_slide_part(name: str) -> bool:
    return (
        name.startswith("ppt/slides/")
        and name.endswith(".xml")
        and "/_rels/" not in name
        and name.count("/") == 2
    )


def _presentation_slide_parts(
    xml: Mapping[str, ElementTree.Element], members: set[str]
) -> list[str]:
    presentation_name = "ppt/presentation.xml"
    presentation = xml.get(presentation_name)
    if presentation is None:
        raise ArtifactValidationError(
            "Generated presentation is missing slide relationship metadata"
        )
    slide_id_name = (
        "{http://schemas.openxmlformats.org/presentationml/2006/main}sldId"
    )
    relationship_id_name = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    relationships = _relationships_for_part(
        xml, presentation_name, "presentation slide"
    )
    ordered = []
    seen_ids = set()
    seen_parts = set()
    for slide_id in presentation.iter(slide_id_name):
        relationship_id = slide_id.attrib.get(relationship_id_name)
        if not relationship_id or relationship_id in seen_ids:
            raise ArtifactValidationError(
                "Generated presentation has an invalid slide relationship ID"
            )
        seen_ids.add(relationship_id)
        relationship = relationships.get(relationship_id)
        if (
            relationship is None
            or relationship.attrib.get("Type") not in SLIDE_RELATIONSHIP_TYPES
        ):
            raise ArtifactValidationError(
                "Generated presentation has an invalid slide relationship"
            )
        part = _relationship_target_member(
            presentation_name,
            relationship,
            members,
            "presentation slide",
            required_prefix="ppt/slides/",
        )
        if not _is_presentation_slide_part(part) or part in seen_parts:
            raise ArtifactValidationError(
                "Generated presentation has an invalid slide relationship target"
            )
        seen_parts.add(part)
        ordered.append(part)
    return ordered


def _whitespace_stripped(text: str) -> str:
    """Normalize insignificant whitespace around OOXML run boundaries."""
    return re.sub(r"\s+", "", text)


def _normalized_expected_counts(values: list[str]) -> Counter:
    return Counter(
        normalized
        for value in values
        if (normalized := _whitespace_stripped(value))
    )


def _structured_text_items(
    root: ElementTree.Element | None,
    container_name: str,
    text_name: str,
) -> list[str]:
    if root is None:
        return []
    containers = list(root.iter(container_name))
    if containers:
        return [
            "".join(node.text or "" for node in container.iter(text_name))
            for container in containers
        ]
    # Small synthetic packages and malformed vendor output may omit the normal
    # paragraph wrapper. Keep each text node distinct rather than joining the
    # whole document and allowing one node to satisfy several expectations.
    return [node.text or "" for node in root.iter(text_name)]


def _require_structured_text_counts(
    values: list[str], actual_items: list[str], context: str
) -> None:
    expected = [
        normalized
        for value in values
        if (normalized := _whitespace_stripped(value))
    ]
    actual = [
        normalized
        for value in actual_items
        if (normalized := _whitespace_stripped(value))
    ]
    candidates = [
        [index for index, item in enumerate(actual) if value in item]
        for value in expected
    ]
    matched_expected = [-1] * len(actual)

    def assign(expected_index: int, visited: set[int]) -> bool:
        for actual_index in candidates[expected_index]:
            if actual_index in visited:
                continue
            visited.add(actual_index)
            previous = matched_expected[actual_index]
            if previous == -1 or assign(previous, visited):
                matched_expected[actual_index] = expected_index
                return True
        return False

    order = sorted(
        range(len(expected)),
        key=lambda index: (len(candidates[index]), -len(expected[index])),
    )
    if any(not assign(index, set()) for index in order):
        raise ArtifactValidationError(
            f"Generated artifact is missing {context} content count"
        )


def _spreadsheet_cell_text(xml: Mapping[str, ElementTree.Element]) -> list[str]:
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    shared = []
    shared_root = xml.get("xl/sharedStrings.xml")
    if shared_root is not None:
        for item in shared_root.iter(f"{{{namespace}}}si"):
            shared.append("".join(
                text.text or "" for text in item.iter(f"{{{namespace}}}t")
            ))
    values = []
    for name, root in sorted(xml.items()):
        if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
            continue
        for cell in root.iter(f"{{{namespace}}}c"):
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                value = "".join(
                    text.text or "" for text in cell.iter(f"{{{namespace}}}t")
                )
            else:
                value_node = cell.find(f"{{{namespace}}}v")
                value = "" if value_node is None else value_node.text or ""
                if cell_type == "s" and value:
                    try:
                        value = shared[int(value)]
                    except (IndexError, ValueError):
                        value = ""
            if value:
                values.append(value)
    return values


def _presentation_expected_by_slide(recorded: RecordedGeneration) -> list[list[str]]:
    slides: list[list[str]] = []
    creates = {
        "slide.add_title", "slide.add_section", "slide.add_bullets", "slide.add_blank"
    }
    for operation in recorded.plan.operations:
        if operation.op in creates:
            slides.append(_operation_expected_text(
                operation, recorded.plan.component
            ))
        elif operation.op == "slide.add_table":
            slide_index = int(operation.args.get("slide", 0) or 0)
            if 1 <= slide_index <= len(slides):
                slides[slide_index - 1].extend(_operation_expected_text(
                    operation, recorded.plan.component
                ))
    return slides


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
    members = _package_members(path)
    expected = _expected_text(recorded)
    if format_name == "docx":
        word_paragraph = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
        )
        word_text = (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
        )
        _require_structured_text_counts(
            expected,
            _structured_text_items(
                xml.get("word/document.xml"), word_paragraph, word_text
            ),
            "writer",
        )
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
        planned_images = sum(
            operation.op == "writer.add_image"
            for operation in recorded.plan.operations
        )
        actual_images = _valid_image_reference_count(
            xml,
            members,
            "word/document.xml",
            "writer",
            media_prefix="word/media/",
        )
        if actual_images != planned_images:
            raise ArtifactValidationError(
                "Generated document is missing planned image structure"
            )
    elif format_name == "xlsx":
        actual_values = Counter(
            normalized
            for value in _spreadsheet_cell_text(xml)
            if (normalized := _whitespace_stripped(value))
        )
        for value, required in _normalized_expected_counts(expected).items():
            if actual_values[value] < required:
                raise ArtifactValidationError(
                    "Generated workbook is missing spreadsheet content count"
                )
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
        actual = sum(_is_presentation_slide_part(name) for name in xml)
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
        expected_slides = _presentation_expected_by_slide(recorded)
        slide_names = _presentation_slide_parts(xml, members)
        if len(slide_names) != planned:
            raise ArtifactValidationError(
                "Generated presentation has an unexpected slide relationship count"
            )
        drawing_text = (
            "{http://schemas.openxmlformats.org/drawingml/2006/main}t"
        )
        drawing_paragraph = (
            "{http://schemas.openxmlformats.org/drawingml/2006/main}p"
        )
        for index, (name, slide_expected) in enumerate(
            zip(slide_names, expected_slides), start=1
        ):
            _require_structured_text_counts(
                slide_expected,
                _structured_text_items(
                    xml[name], drawing_paragraph, drawing_text
                ),
                f"slide {index}",
            )
        table_name = (
            "{http://schemas.openxmlformats.org/drawingml/2006/main}tbl"
        )
        actual_tables = sum(
            sum(1 for _ in root.iter(table_name))
            for name, root in xml.items()
            if _is_presentation_slide_part(name)
        )
        planned_tables = sum(
            operation.op == "slide.add_table"
            for operation in recorded.plan.operations
        )
        if actual_tables < planned_tables:
            raise ArtifactValidationError(
                "Generated presentation is missing planned table structure"
            )
        planned_images = Counter(
            int(operation.args["slide"])
            for operation in recorded.plan.operations
            if operation.op == "slide.add_image"
        )
        for slide_index, slide_name in enumerate(slide_names, start=1):
            actual_images = _valid_image_reference_count(
                xml,
                members,
                slide_name,
                f"slide {slide_index}",
                media_prefix="ppt/media/",
            )
            if actual_images != planned_images[slide_index]:
                raise ArtifactValidationError(
                    "Generated presentation is missing planned "
                    f"slide {slide_index} image structure"
                )


def _wait_for_generated_package(
    path: Path,
    format_name: str,
    recorded: RecordedGeneration,
    template_digest: str,
    *,
    deadline: float,
) -> None:
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            validate_before_deadline(
                partial(
                    _validate_generated_package,
                    format_name=format_name,
                    recorded=recorded,
                    template_digest=template_digest,
                ),
                path,
                deadline,
            )
            return
        except ArtifactValidationError as exc:
            failure = exc
        except TimeoutError:
            if failure is not None:
                raise failure
            raise
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
    deadline: float,
    feasibility: bool,
) -> Path:
    is_pdf = request.format_name == "pdf"
    if runtime.staging_dir is None:
        raise _error(
            request,
            "STAGING_UNAVAILABLE",
            "WPS container staging session was not created",
        )
    require_remaining(deadline)
    runtime.prepare_profiles()
    require_remaining(deadline)
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
    require_remaining(deadline)
    staged_pdf = policy.require_allowed(staged.with_suffix(".pdf")) if is_pdf else None
    resources = stage_generation_resources(
        request, recorded, runtime, deadline=deadline
    )
    require_remaining(deadline)
    runtime.start_servers(deadline=deadline)
    runtime.activate_component(request.component, deadline=deadline)
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
    require_remaining(deadline)
    command = bridge.issue(
        request.component,
        method,
        command_params,
    )
    try:
        result = bridge.wait_result(command.id, remaining(deadline))
    except TimeoutError as exc:
        bridge.state.cancel(command.id)
        raise _error(
            request,
            "GENERATION_COMMAND_FAILED",
            "Timed out waiting for WPS generation",
        ) from exc
    require_remaining(deadline)
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
            _wait_for_marker(staged, request.format_name, deadline=deadline)
        elif is_pdf:
            _wait_for_generated_package(
                staged,
                "docx",
                recorded,
                template_digest,
                deadline=deadline,
            )
            _wait_for_pdf(staged_pdf, deadline=deadline)
        else:
            _wait_for_generated_package(
                staged,
                request.format_name,
                recorded,
                template_digest,
                deadline=deadline,
            )
    except ArtifactValidationError as exc:
        raise _error(
            request,
            "STAGED_ARTIFACT_INVALID",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    try:
        if feasibility:
            artifact_validator = partial(
                _validate_marker_package, format_name=request.format_name
            )
        elif is_pdf:
            artifact_validator = validate_pdf
        else:
            artifact_validator = partial(
                _validate_generated_package,
                format_name=request.format_name,
                recorded=recorded,
                template_digest=template_digest,
            )

        def validator(path: Path) -> None:
            try:
                require_remaining(deadline)
                validate_before_deadline(artifact_validator, path, deadline)
                require_remaining(deadline)
            except TimeoutError as exc:
                raise ArtifactValidationError(
                    "WPS generation deadline expired during artifact validation"
                ) from exc

        published = publish_artifact(
            staged_pdf if is_pdf else staged,
            request.output,
            overwrite=request.overwrite,
            validator=validator,
            deadline=deadline,
        )
        if is_pdf:
            staged.unlink(missing_ok=True)
            staged_pdf.unlink(missing_ok=True)
        return published
    except ArtifactTransportError as exc:
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
    timeout: float = 600,
    feasibility: bool = False,
    deadline: Optional[float] = None,
) -> Path:
    """Execute one host-validated plan against a private template clone."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    if deadline is None:
        deadline = time.monotonic() + timeout
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
    require_remaining(deadline)
    method = METHODS[request.component]

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-generate-")
    ).resolve()
    runtime = None
    try:
        with bridge_factory(ORIGINS) as bridge:
            require_remaining(deadline)
            runtime = runtime_factory(
                probe_root,
                runtime_root / "runtime",
                bridge.url,
                bridge.token,
                deadline=deadline,
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
                        deadline,
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
    timeout: float = 600,
    _deadline: Optional[float] = None,
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
        deadline=_deadline,
    )


def execute_feasibility_plan(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    *,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 600,
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
    timeout: float = 600,
    overwrite: bool = False,
) -> Path:
    """Record a public renderer and generate through the gated macOS backend."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    deadline = time.monotonic() + timeout
    normalized_format = str(format_name).lower().lstrip(".")
    component = FORMAT_COMPONENTS.get(normalized_format, "writer")
    request = GenerationRequest(
        Path(output), component, normalized_format, overwrite=bool(overwrite)
    )
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
        _deadline=deadline,
    )
