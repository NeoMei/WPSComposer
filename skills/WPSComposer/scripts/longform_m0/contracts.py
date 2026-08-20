"""Closed evidence contract for the dual-platform long-form M0 gate."""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path, PurePosixPath
import re
from types import MappingProxyType
from typing import Any, Mapping

PROBE_VERSION = "0.8.0-m0.1"
SCHEMA_VERSION = 1
PROTOCOL_VERSION = 2
RESOURCE_MANIFEST_VERSION = 1
REQUIRED_IDS = frozenset(range(1, 15))
OPTIONAL_IDS = frozenset({15})
ALL_IDS = REQUIRED_IDS | OPTIONAL_IDS
STATUSES = frozenset({"passed", "failed", "unsupported", "not-run"})

_ROOT_FIELDS = frozenset(
    {
        "schemaVersion",
        "probeVersion",
        "platform",
        "wpsVersion",
        "protocolVersion",
        "resourceManifestVersion",
        "capabilities",
        "artifacts",
        "failures",
    }
)
_CAPABILITY_FIELDS = frozenset(
    {"id", "status", "checks", "artifacts", "metrics"}
)
_ARTIFACT_FIELDS = frozenset({"name", "sha256"})
_ARTIFACT_KINDS = frozenset({"docx", "pdf"})
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_WINDOWS_ABSOLUTE = re.compile(r"^[A-Za-z]:[\\/]")


@dataclass(frozen=True)
class CapabilityEvidence:
    id: int
    status: str
    checks: tuple[str, ...]
    artifacts: tuple[str, ...]
    metrics: Mapping[str, Any]


@dataclass(frozen=True)
class PlatformEvidence:
    schema_version: int
    probe_version: str
    platform: str
    wps_version: str
    protocol_version: int
    resource_manifest_version: int
    capabilities: tuple[CapabilityEvidence, ...]
    artifacts: Mapping[str, Mapping[str, str]]
    failures: tuple[Mapping[str, Any], ...]


def _require_exact_fields(
    raw: Mapping[str, Any], allowed: frozenset[str], label: str
) -> None:
    unknown = set(raw) - allowed
    missing = allowed - set(raw)
    if unknown or missing:
        details = []
        if unknown:
            details.append("unknown fields: " + ", ".join(sorted(unknown)))
        if missing:
            details.append("missing fields: " + ", ".join(sorted(missing)))
        raise ValueError(f"{label} has " + "; ".join(details))


def _require_version(value: Any, expected: int, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value != expected:
        raise ValueError(f"{label} must be {expected}")
    return value


def _require_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{label} must be a non-empty string")
    return value


def _require_relative_name(value: Any, label: str) -> str:
    name = _require_string(value, label)
    normalized = name.replace("\\", "/")
    path = PurePosixPath(normalized)
    if (
        path.is_absolute()
        or _WINDOWS_ABSOLUTE.match(name)
        or any(part in {"", ".", ".."} for part in path.parts)
    ):
        raise ValueError(f"{label} must be a safe relative filename")
    return name


def _freeze_json(value: Any, label: str) -> Any:
    if value is None or isinstance(value, (str, bool)):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value or value in {float("inf"), float("-inf")}:
            raise ValueError(f"{label} must contain finite JSON values")
        return value
    if isinstance(value, list):
        return tuple(
            _freeze_json(item, f"{label}[{index}]")
            for index, item in enumerate(value)
        )
    if isinstance(value, Mapping):
        frozen = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise ValueError(f"{label} keys must be strings")
            frozen[key] = _freeze_json(item, f"{label}.{key}")
        return MappingProxyType(frozen)
    raise ValueError(f"{label} must contain JSON-compatible values")


def _validate_capability(raw: Any, platform: str) -> CapabilityEvidence:
    if not isinstance(raw, Mapping):
        raise ValueError("capability must be an object")
    _require_exact_fields(raw, _CAPABILITY_FIELDS, "capability")
    capability_id = raw["id"]
    if (
        isinstance(capability_id, bool)
        or not isinstance(capability_id, int)
        or capability_id not in ALL_IDS
    ):
        raise ValueError("capability id must be in 1-15")
    status = raw["status"]
    if not isinstance(status, str) or status not in STATUSES:
        raise ValueError(f"capability {capability_id} has invalid status")
    if capability_id in REQUIRED_IDS and status == "unsupported":
        raise ValueError(
            f"required capability {capability_id} cannot be unsupported"
        )
    checks_raw = raw["checks"]
    if not isinstance(checks_raw, list) or any(
        not isinstance(check, str) or not check for check in checks_raw
    ):
        raise ValueError(f"capability {capability_id} checks must be strings")
    checks = tuple(checks_raw)
    if len(set(checks)) != len(checks):
        raise ValueError(f"capability {capability_id} checks must be unique")
    if capability_id == 2 and platform == "macos":
        if status != "passed" or checks != ("not-applicable-macos",):
            raise ValueError(
                "macos capability 2 must be passed as not-applicable-macos"
            )
    elif status == "passed" and "native" not in checks:
        raise ValueError(
            f"passed capability {capability_id} must include native check"
        )
    if status == "passed" and capability_id >= 3 and not {
        "reopened",
        "refreshed",
    }.issubset(checks):
        raise ValueError(
            f"passed capability {capability_id} must be reopened and refreshed"
        )
    artifacts_raw = raw["artifacts"]
    if not isinstance(artifacts_raw, list):
        raise ValueError(f"capability {capability_id} artifacts must be a list")
    artifacts = tuple(
        _require_relative_name(item, f"capability {capability_id} artifact")
        for item in artifacts_raw
    )
    metrics_raw = raw["metrics"]
    if not isinstance(metrics_raw, Mapping):
        raise ValueError(f"capability {capability_id} metrics must be an object")
    metrics = _freeze_json(metrics_raw, f"capability {capability_id} metrics")
    return CapabilityEvidence(
        id=capability_id,
        status=status,
        checks=checks,
        artifacts=artifacts,
        metrics=metrics,
    )


def _validate_artifacts(raw: Any) -> Mapping[str, Mapping[str, str]]:
    if not isinstance(raw, Mapping):
        raise ValueError("artifacts must be an object")
    unknown = set(raw) - _ARTIFACT_KINDS
    if unknown:
        raise ValueError(
            "artifacts contain unknown fields: " + ", ".join(sorted(unknown))
        )
    artifacts = {}
    for kind in sorted(raw):
        entry = raw[kind]
        if not isinstance(entry, Mapping):
            raise ValueError(f"{kind} artifact must be an object")
        _require_exact_fields(entry, _ARTIFACT_FIELDS, f"{kind} artifact")
        name = _require_relative_name(entry["name"], f"{kind} artifact name")
        digest = entry["sha256"]
        if not isinstance(digest, str) or not _SHA256.fullmatch(digest):
            raise ValueError(f"{kind} artifact SHA-256 is invalid")
        artifacts[kind] = MappingProxyType(
            {"name": name, "sha256": digest}
        )
    return MappingProxyType(artifacts)


def validate_platform_evidence(raw: Any) -> PlatformEvidence:
    """Validate untrusted native evidence and return an immutable snapshot."""
    if not isinstance(raw, Mapping):
        raise ValueError("platform evidence must be an object")
    _require_exact_fields(raw, _ROOT_FIELDS, "platform evidence")
    schema_version = _require_version(
        raw["schemaVersion"], SCHEMA_VERSION, "schemaVersion"
    )
    probe_version = _require_string(raw["probeVersion"], "probeVersion")
    if probe_version != PROBE_VERSION:
        raise ValueError(f"probeVersion must be {PROBE_VERSION}")
    platform = raw["platform"]
    if platform not in {"windows", "macos"}:
        raise ValueError("platform must be windows or macos")
    wps_version = _require_string(raw["wpsVersion"], "wpsVersion")
    protocol_version = _require_version(
        raw["protocolVersion"], PROTOCOL_VERSION, "protocolVersion"
    )
    manifest_version = _require_version(
        raw["resourceManifestVersion"],
        RESOURCE_MANIFEST_VERSION,
        "resourceManifestVersion",
    )
    capabilities_raw = raw["capabilities"]
    if not isinstance(capabilities_raw, list):
        raise ValueError("capabilities must be a list")
    capabilities = tuple(
        sorted(
            (
                _validate_capability(item, platform)
                for item in capabilities_raw
            ),
            key=lambda item: item.id,
        )
    )
    capability_ids = [item.id for item in capabilities]
    if len(capability_ids) != len(ALL_IDS) or set(capability_ids) != ALL_IDS:
        raise ValueError("capability ids must contain each id 1-15 exactly once")
    failures_raw = raw["failures"]
    if not isinstance(failures_raw, list) or any(
        not isinstance(failure, Mapping) for failure in failures_raw
    ):
        raise ValueError("failures must be a list of objects")
    failures = tuple(
        _freeze_json(failure, f"failures[{index}]")
        for index, failure in enumerate(failures_raw)
    )
    artifacts = _validate_artifacts(raw["artifacts"])
    platform_passed = all(
        capability.status == "passed"
        for capability in capabilities
        if capability.id in REQUIRED_IDS
    )
    if platform_passed and set(artifacts) != _ARTIFACT_KINDS:
        raise ValueError("passed platform must contain docx and pdf artifacts")
    artifact_names = {entry["name"] for entry in artifacts.values()}
    for capability in capabilities:
        unregistered = set(capability.artifacts) - artifact_names
        if unregistered:
            raise ValueError(
                f"capability {capability.id} references unregistered artifact: "
                + ", ".join(sorted(unregistered))
            )
    return PlatformEvidence(
        schema_version=schema_version,
        probe_version=probe_version,
        platform=platform,
        wps_version=wps_version,
        protocol_version=protocol_version,
        resource_manifest_version=manifest_version,
        capabilities=capabilities,
        artifacts=artifacts,
        failures=failures,
    )


def merge_platform_evidence(
    windows: PlatformEvidence, macos: PlatformEvidence
) -> dict[str, Any]:
    """Merge validated native evidence into the release gate decision."""
    if windows.platform != "windows" or macos.platform != "macos":
        raise ValueError("merge requires windows and macos evidence")
    windows_by_id = {item.id: item for item in windows.capabilities}
    macos_by_id = {item.id: item for item in macos.capabilities}
    blocking = [
        capability_id
        for capability_id in sorted(REQUIRED_IDS)
        if windows_by_id[capability_id].status != "passed"
        or macos_by_id[capability_id].status != "passed"
    ]
    svg_included = (
        windows_by_id[15].status == "passed"
        and macos_by_id[15].status == "passed"
    )
    return {
        "schemaVersion": SCHEMA_VERSION,
        "probeVersion": PROBE_VERSION,
        "decision": "no-go" if blocking else "go",
        "svg": "included" if svg_included else "excluded",
        "blockingCapabilities": blocking,
        "platforms": {
            "windows": {"wpsVersion": windows.wps_version},
            "macos": {"wpsVersion": macos.wps_version},
        },
        "capabilities": [
            {
                "id": capability_id,
                "windows": windows_by_id[capability_id].status,
                "macos": macos_by_id[capability_id].status,
            }
            for capability_id in sorted(ALL_IDS)
        ],
    }


def write_canonical_json(path: Path, value: Any) -> None:
    """Write stable UTF-8 JSON suitable for hashing and evidence review."""
    serialized = json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    Path(path).write_text(serialized + "\n", encoding="utf-8")
