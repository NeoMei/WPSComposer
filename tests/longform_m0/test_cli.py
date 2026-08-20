from __future__ import annotations

import json
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform_m0 import __main__ as cli
from skills.WPSComposer.scripts.longform_m0.contracts import write_canonical_json


def evidence(platform: str, *, failed_id: int | None = None) -> dict:
    capabilities = []
    for capability_id in range(1, 16):
        status = "failed" if capability_id == failed_id else "passed"
        checks = ["native", "reopened", "refreshed"]
        if capability_id < 3:
            checks = ["native"]
        if platform == "macos" and capability_id == 2:
            checks = ["not-applicable-macos"]
        capabilities.append(
            {
                "id": capability_id,
                "status": status,
                "checks": checks,
                "artifacts": (
                    ["probe.docx", "probe.pdf"]
                    if status == "passed" and capability_id >= 3
                    else []
                ),
                "metrics": {},
            }
        )
    return {
        "schemaVersion": 1,
        "probeVersion": "0.8.0-m0.1",
        "platform": platform,
        "wpsVersion": "12.1.fake",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "capabilities": capabilities,
        "artifacts": {
            "docx": {"name": "probe.docx", "sha256": "a" * 64},
            "pdf": {"name": "probe.pdf", "sha256": "b" * 64},
        },
        "failures": [],
    }


@pytest.mark.parametrize(
    "argv",
    (
        [],
        ["--platform", "macos"],
        ["--platform", "macos", "--output-dir", "out", "--timeout", "0"],
        ["--platform", "windows", "--output-dir", "out", "--timeout", "-1"],
        ["--platform", "windows", "--output-dir", "out", "--timeout", "nan"],
        ["--platform", "windows", "--output-dir", "out", "--timeout", "inf"],
        ["--platform", "verify", "--output-dir", "out"],
        [
            "--platform",
            "verify",
            "--output-dir",
            "out",
            "--windows-evidence",
            "windows.json",
        ],
    ),
)
def test_parser_requires_closed_platform_arguments(argv):
    with pytest.raises(SystemExit) as caught:
        cli.parse_args(argv)
    assert caught.value.code == 2


def test_parser_rejects_evidence_arguments_for_native_platform():
    with pytest.raises(SystemExit) as caught:
        cli.parse_args(
            [
                "--platform",
                "macos",
                "--output-dir",
                "out",
                "--windows-evidence",
                "windows.json",
            ]
        )
    assert caught.value.code == 2


def test_verify_writes_canonical_go_matrix(tmp_path: Path):
    windows = tmp_path / "windows.json"
    macos = tmp_path / "macos.json"
    write_canonical_json(windows, evidence("windows"))
    write_canonical_json(macos, evidence("macos"))
    output = tmp_path / "matrix"

    exit_code = cli.main(
        [
            "--platform",
            "verify",
            "--output-dir",
            str(output),
            "--windows-evidence",
            str(windows),
            "--macos-evidence",
            str(macos),
        ]
    )

    assert exit_code == 0
    matrix_path = output / "matrix-evidence.json"
    matrix = json.loads(matrix_path.read_text(encoding="utf-8"))
    assert matrix["decision"] == "go"
    assert matrix["blockingCapabilities"] == []


def test_verify_returns_one_for_no_go_and_two_for_invalid_evidence(tmp_path: Path):
    windows = tmp_path / "windows.json"
    macos = tmp_path / "macos.json"
    write_canonical_json(windows, evidence("windows", failed_id=10))
    write_canonical_json(macos, evidence("macos"))
    common = [
        "--platform",
        "verify",
        "--windows-evidence",
        str(windows),
        "--macos-evidence",
        str(macos),
    ]

    assert cli.main([*common, "--output-dir", str(tmp_path / "no-go")]) == 1
    windows.write_text("not json", encoding="utf-8")
    assert cli.main([*common, "--output-dir", str(tmp_path / "invalid")]) == 2


def test_native_platform_exit_code_uses_validated_required_capabilities(
    monkeypatch, tmp_path: Path
):
    output = tmp_path / "macos"
    output.mkdir()
    evidence_path = output / "platform-evidence.json"
    write_canonical_json(evidence_path, evidence("macos", failed_id=5))
    monkeypatch.setattr(cli, "run_macos_probe", lambda *args, **kwargs: evidence_path)

    assert (
        cli.main(
            [
                "--platform",
                "macos",
                "--output-dir",
                str(output),
                "--timeout",
                "30",
            ]
        )
        == 1
    )
