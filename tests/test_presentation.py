from __future__ import annotations

import importlib
from pathlib import Path

import pytest


def _presentation_module():
    return importlib.import_module("skills.WPSComposer.scripts.presentation")


@pytest.mark.parametrize("suffix", ["docx", "pdf", "pptx", "xlsx"])
def test_present_artifact_passes_every_output_format_as_one_literal_argument(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    suffix: str,
):
    presentation = _presentation_module()
    artifact = tmp_path / f"result $HOME;$(touch nope).{suffix}"
    artifact.write_bytes(b"published")
    calls = []

    monkeypatch.setattr(presentation.sys, "platform", "darwin")
    monkeypatch.setattr(
        presentation.subprocess,
        "run",
        lambda argv, **kwargs: calls.append((argv, kwargs)),
    )
    presentation.present_artifact(artifact)

    assert calls[0][0] == ["open", str(artifact.resolve())]
    assert calls[0][1]["shell"] is False
    assert calls[0][1]["check"] is True
    assert calls[0][1]["timeout"] == 10


@pytest.mark.parametrize(
    ("platform", "command"),
    [
        ("darwin", "open"),
        ("win32", "explorer.exe"),
        ("linux", "xdg-open"),
    ],
)
def test_present_artifact_uses_shell_free_platform_launcher(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    platform: str,
    command: str,
):
    presentation = _presentation_module()
    artifact = tmp_path / "final artifact.pdf"
    artifact.write_bytes(b"published")
    calls = []

    monkeypatch.setattr(presentation.sys, "platform", platform)
    monkeypatch.setattr(
        presentation.subprocess,
        "run",
        lambda argv, **kwargs: calls.append((argv, kwargs)),
    )
    presentation.present_artifact(artifact)

    assert calls[0][0] == [command, str(artifact.resolve())]
    assert calls[0][1]["shell"] is False
    assert calls[0][1]["check"] is True


def test_present_artifact_rejects_missing_or_non_file_path_before_launch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    presentation = _presentation_module()
    calls = []
    monkeypatch.setattr(
        presentation.subprocess,
        "run",
        lambda argv, **kwargs: calls.append((argv, kwargs)),
    )

    with pytest.raises(FileNotFoundError, match="Final artifact is not a file"):
        presentation.present_artifact(tmp_path / "missing.pdf")
    with pytest.raises(FileNotFoundError, match="Final artifact is not a file"):
        presentation.present_artifact(tmp_path)

    assert calls == []


def test_present_artifact_rejects_unknown_platform_before_launch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    presentation = _presentation_module()
    artifact = tmp_path / "final.pdf"
    artifact.write_bytes(b"published")
    calls = []
    monkeypatch.setattr(presentation.sys, "platform", "plan9")
    monkeypatch.setattr(
        presentation.subprocess,
        "run",
        lambda argv, **kwargs: calls.append((argv, kwargs)),
    )

    with pytest.raises(OSError, match="Unsupported platform"):
        presentation.present_artifact(artifact)

    assert calls == []


def test_present_artifact_surfaces_nonzero_launcher_exit(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    presentation = _presentation_module()
    artifact = tmp_path / "final.pdf"
    artifact.write_bytes(b"published")
    monkeypatch.setattr(presentation.sys, "platform", "darwin")

    def fail(argv, **kwargs):
        raise presentation.subprocess.CalledProcessError(3, argv)

    monkeypatch.setattr(presentation.subprocess, "run", fail)
    with pytest.raises(presentation.subprocess.CalledProcessError) as caught:
        presentation.present_artifact(artifact)

    assert caught.value.returncode == 3
