from __future__ import annotations

import json
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform_m0.contracts import (
    ALL_IDS,
    OPTIONAL_IDS,
    PROBE_VERSION,
    REQUIRED_IDS,
    CapabilityEvidence,
    PlatformEvidence,
    merge_platform_evidence,
    validate_platform_evidence,
    write_canonical_json,
)


def passing_evidence(platform: str) -> dict:
    capabilities = []
    for capability_id in range(1, 16):
        checks = ["native"]
        if capability_id >= 3:
            checks.extend(("reopened", "refreshed"))
        if capability_id == 2 and platform == "macos":
            checks = ["not-applicable-macos"]
        capabilities.append(
            {
                "id": capability_id,
                "status": "passed",
                "checks": checks,
                "artifacts": ["probe.docx", "probe.pdf"],
                "metrics": {},
            }
        )
    return {
        "schemaVersion": 1,
        "probeVersion": PROBE_VERSION,
        "platform": platform,
        "wpsVersion": "12.1.test",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "capabilities": capabilities,
        "artifacts": {
            "docx": {"name": "probe.docx", "sha256": "a" * 64},
            "pdf": {"name": "probe.pdf", "sha256": "b" * 64},
        },
        "failures": [],
    }


def failed_evidence(platform: str) -> dict:
    raw = passing_evidence(platform)
    for capability in raw["capabilities"]:
        capability["artifacts"] = []
        if capability["id"] == 2 and platform == "macos":
            continue
        capability["status"] = "not-run"
        capability["checks"] = []
    raw["artifacts"] = {}
    raw["failures"] = [
        {"code": "ENGINE_UNAVAILABLE", "message": "WPS is unavailable"}
    ]
    return raw


def test_capability_id_sets_are_closed():
    assert REQUIRED_IDS == frozenset(range(1, 15))
    assert OPTIONAL_IDS == frozenset({15})
    assert ALL_IDS == frozenset(range(1, 16))


def test_valid_platform_evidence_is_immutable_and_sorted():
    raw = passing_evidence("macos")
    raw["capabilities"].reverse()

    evidence = validate_platform_evidence(raw)

    assert isinstance(evidence, PlatformEvidence)
    assert all(isinstance(item, CapabilityEvidence) for item in evidence.capabilities)
    assert [item.id for item in evidence.capabilities] == list(range(1, 16))
    with pytest.raises(TypeError):
        evidence.artifacts["docx"] = {}  # type: ignore[index]


def test_rejects_duplicate_or_missing_capability_ids():
    raw = passing_evidence("macos")
    raw["capabilities"].pop()
    raw["capabilities"].append(dict(raw["capabilities"][0]))

    with pytest.raises(ValueError, match="capability ids"):
        validate_platform_evidence(raw)


@pytest.mark.parametrize(
    ("location", "key"),
    [
        ("root", "unknownRoot"),
        ("capability", "unknownCapability"),
    ],
)
def test_rejects_unknown_fields(location: str, key: str):
    raw = passing_evidence("windows")
    if location == "root":
        raw[key] = True
    else:
        raw["capabilities"][0][key] = True

    with pytest.raises(ValueError, match="unknown fields"):
        validate_platform_evidence(raw)


def test_rejects_unsupported_required_capability():
    raw = passing_evidence("windows")
    raw["capabilities"][8]["status"] = "unsupported"

    with pytest.raises(ValueError, match="required capability 9"):
        validate_platform_evidence(raw)


@pytest.mark.parametrize("capability_id", range(3, 16))
def test_passed_content_capability_requires_reopen_and_refresh(capability_id: int):
    raw = passing_evidence("windows")
    raw["capabilities"][capability_id - 1]["checks"].remove("refreshed")

    with pytest.raises(ValueError, match="reopened and refreshed"):
        validate_platform_evidence(raw)


def test_passed_capability_requires_native_check_except_macos_ownership():
    raw = passing_evidence("windows")
    raw["capabilities"][0]["checks"] = []

    with pytest.raises(ValueError, match="native"):
        validate_platform_evidence(raw)

    macos = validate_platform_evidence(passing_evidence("macos"))
    assert macos.capabilities[1].checks == ("not-applicable-macos",)


@pytest.mark.parametrize(
    ("name", "digest", "message"),
    [
        ("/private/probe.docx", "a" * 64, "relative filename"),
        ("../probe.docx", "a" * 64, "relative filename"),
        ("probe.docx", "not-a-digest", "SHA-256"),
    ],
)
def test_rejects_unsafe_artifact_entries(name: str, digest: str, message: str):
    raw = passing_evidence("macos")
    raw["artifacts"]["docx"] = {"name": name, "sha256": digest}

    with pytest.raises(ValueError, match=message):
        validate_platform_evidence(raw)


def test_fatal_platform_failure_can_be_recorded_without_artifacts():
    evidence = validate_platform_evidence(failed_evidence("macos"))

    assert evidence.artifacts == {}
    assert evidence.failures[0]["code"] == "ENGINE_UNAVAILABLE"


def test_passing_platform_requires_docx_and_pdf_artifacts():
    raw = passing_evidence("windows")
    raw["artifacts"].pop("pdf")

    with pytest.raises(ValueError, match="passed platform.*docx and pdf"):
        validate_platform_evidence(raw)


def test_capability_cannot_reference_unregistered_artifact():
    raw = passing_evidence("macos")
    raw["capabilities"][4]["artifacts"].append("unregistered.pdf")

    with pytest.raises(ValueError, match="unregistered artifact"):
        validate_platform_evidence(raw)


def test_optional_svg_does_not_block_go_decision():
    windows = passing_evidence("windows")
    macos = passing_evidence("macos")
    macos["capabilities"][14] = {
        "id": 15,
        "status": "unsupported",
        "checks": ["native-insertion-unavailable"],
        "artifacts": ["probe.docx", "probe.pdf"],
        "metrics": {},
    }

    matrix = merge_platform_evidence(
        validate_platform_evidence(windows),
        validate_platform_evidence(macos),
    )

    assert matrix["decision"] == "go"
    assert matrix["svg"] == "excluded"
    assert matrix["blockingCapabilities"] == []


def test_required_failure_forces_no_go():
    windows = passing_evidence("windows")
    macos = passing_evidence("macos")
    macos["capabilities"][9]["status"] = "failed"

    matrix = merge_platform_evidence(
        validate_platform_evidence(windows),
        validate_platform_evidence(macos),
    )

    assert matrix["decision"] == "no-go"
    assert matrix["blockingCapabilities"] == [10]


def test_merge_requires_one_evidence_record_per_platform():
    macos = validate_platform_evidence(passing_evidence("macos"))

    with pytest.raises(ValueError, match="windows and macos"):
        merge_platform_evidence(macos, macos)


def test_canonical_json_is_sorted_compact_and_newline_terminated(tmp_path: Path):
    path = tmp_path / "evidence.json"

    write_canonical_json(path, {"z": 1, "a": {"d": 2, "c": 3}})

    assert path.read_text(encoding="utf-8") == (
        '{"a":{"c":3,"d":2},"z":1}\n'
    )
    assert json.loads(path.read_text(encoding="utf-8"))["z"] == 1
