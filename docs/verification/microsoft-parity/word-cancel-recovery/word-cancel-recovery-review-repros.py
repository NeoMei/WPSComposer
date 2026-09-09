"""Independent pure reproductions for interrupted Mac Word AppleEvents."""

from __future__ import annotations

import importlib
import json
from pathlib import Path
import subprocess

import pytest


@pytest.fixture
def word_host(monkeypatch, tmp_path):
    module = importlib.import_module(
        "skills.WPSComposer.scripts.msoffice.macos_word_session"
    )
    monkeypatch.setattr(module.sys, "platform", "darwin")
    app = tmp_path / "Word.app"
    app.mkdir()
    root = tmp_path / "container"
    root.mkdir()
    source = tmp_path / "source.docx"
    source.write_bytes(b"original source")
    monkeypatch.setattr(module, "WORD_APP", app)
    monkeypatch.setattr(module, "_temporary_root", lambda: root)
    monkeypatch.setattr(module, "validate_native_input", lambda *a, **k: None)
    calls = []

    def success(command, **kwargs):
        calls.append(tuple(command))
        script = Path(command[-1]).read_text()
        stage = Path(command[-1]).parent
        rows = (
            [["binding", 901, str(next(stage.glob("document-*.docx")))]]
            if "WPSC_BIND_OPEN" in script
            else [["ok"]]
        )
        return subprocess.CompletedProcess(
            command,
            0,
            json.dumps(["WPSCOMPOSER_WORD_SESSION_OK", rows]),
            "",
        )

    monkeypatch.setattr(module.subprocess, "run", success)
    return module, root, source, calls


@pytest.mark.parametrize("cancel_type", [KeyboardInterrupt, SystemExit])
def test_bound_cancel_quarantines_and_context_exit_sends_no_close_event(
    word_host, monkeypatch, cancel_type
):
    module, root, source, calls = word_host
    session = module.MacWordSession.open_document(source)
    stage = session.staging_root
    prior_calls = len(calls)

    def cancel(command, **kwargs):
        calls.append(tuple(command))
        raise cancel_type("cancelled native checkpoint")

    monkeypatch.setattr(module.subprocess, "run", cancel)
    with pytest.raises(cancel_type, match="cancelled native checkpoint"):
        with session:
            session._execute_structural(["set nativeRows to {{\"unreachable\"}}"])

    assert len(calls) == prior_calls + 1
    assert session._quarantined
    assert session._retain_evidence
    assert session.lock.file is None
    assert stage.is_dir()
    quarantine = session.lock.quarantine_path
    assert quarantine.is_file()
    detail = json.loads(quarantine.read_text())
    assert detail["stagingRoot"] == str(stage)
    assert detail["documentPath"] == session._bound_path
    logs = list(stage.glob("*.log"))
    assert logs
    assert any(cancel_type.__name__ in log.read_text() for log in logs)


@pytest.mark.parametrize("cancel_type", [KeyboardInterrupt, SystemExit])
def test_open_cancel_retains_script_and_diagnostic_before_releasing_lock(
    word_host, monkeypatch, cancel_type
):
    module, root, source, calls = word_host

    def cancel(command, **kwargs):
        calls.append(tuple(command))
        raise cancel_type("cancelled native open")

    monkeypatch.setattr(module.subprocess, "run", cancel)
    with pytest.raises(cancel_type, match="cancelled native open"):
        module.MacWordSession.open_document(source)

    stages = list(root.glob("wpscomposer-session-*"))
    assert len(calls) == 1
    assert len(stages) == 1
    assert list(stages[0].glob("*.applescript"))
    logs = list(stages[0].glob("*.log"))
    assert logs
    assert any(cancel_type.__name__ in log.read_text() for log in logs)
    quarantine = root / "wpscomposer-native-word.lock.quarantine"
    assert quarantine.is_file()
    detail = json.loads(quarantine.read_text())
    assert detail["stagingRoot"] == str(stages[0])
