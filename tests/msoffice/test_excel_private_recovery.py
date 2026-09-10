from __future__ import annotations

from dataclasses import asdict, replace
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice import macos_office_runtime as runtime
from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity


def _identity(**changes):
    value = ExcelProcessIdentity(
        pid=4242,
        start_seconds=100,
        start_microseconds=250,
        executable="/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel",
        bundle_id="com.microsoft.Excel",
    )
    return replace(value, **changes)


def _quarantine(tmp_path: Path, monkeypatch, *, private_process=...):
    root = tmp_path / "spreadsheet"
    stage = root / "session-private"
    stage.mkdir(parents=True)
    owned = stage / "owned.xlsx"
    owned.write_bytes(b"retained workbook")
    recovery = stage / "recovery.json"
    recovery.write_text('{"retained":true}\n', encoding="utf-8")
    detail = {
        "component": "spreadsheet",
        "staging_path": str(stage),
        "owned_path": str(owned),
        "closed": False,
    }
    if private_process is ...:
        detail["private_process"] = {
            "state": "owned",
            "code": "OSA_TIMEOUT",
            "identity": asdict(_identity()),
            "launch_date": 100.00025,
            "launch_candidate": {"pid": 4242, "verified": False},
            "owned_path": str(owned),
            "inventory": [],
            "outcome_uncertain": True,
        }
    else:
        detail["private_process"] = private_process
    quarantine = root / "native-office.quarantine.json"
    quarantine.write_text(json.dumps(detail) + "\n", encoding="utf-8")
    monkeypatch.setattr(runtime, "_container_root", lambda component: root)
    return root, stage, owned, recovery, quarantine


def _process_listing(monkeypatch, text=""):
    calls = []

    def run(argv, **kwargs):
        calls.append(argv)
        if argv[0] != "/bin/ps":
            pytest.fail("private-process recovery must not query Excel with AppleScript")
        return SimpleNamespace(returncode=0, stdout=text, stderr="")

    monkeypatch.setattr(runtime.subprocess, "run", run)
    return calls


def test_private_recovery_refuses_while_exact_owned_process_is_alive(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    calls = _process_listing(monkeypatch)
    monkeypatch.setattr(runtime, "_process_exists", lambda pid: True, raising=False)
    monkeypatch.setattr(runtime, "_read_excel_process_identity", lambda pid: _identity(), raising=False)

    with pytest.raises(RuntimeError, match="private Excel process is still running"):
        runtime.recover_quarantine("spreadsheet")

    assert quarantine.is_file()
    assert stage.is_dir() and owned.read_bytes() == b"retained workbook"
    assert recovery.read_text(encoding="utf-8") == '{"retained":true}\n'
    assert len(calls) == 1


def test_private_recovery_clears_only_marker_after_owned_process_is_dead(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    calls = _process_listing(
        monkeypatch,
        "9212 /Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel\n",
    )
    signals = []

    def exists(pid):
        signals.append((pid, 0))
        return False

    monkeypatch.setattr(runtime, "_process_exists", exists, raising=False)
    monkeypatch.setattr(
        runtime,
        "_read_excel_process_identity",
        lambda pid: pytest.fail("a dead PID has no identity to read"),
        raising=False,
    )

    assert runtime.recover_quarantine("spreadsheet") is True

    assert not quarantine.exists()
    assert stage.is_dir() and owned.read_bytes() == b"retained workbook"
    assert recovery.read_text(encoding="utf-8") == '{"retained":true}\n'
    assert signals == [(4242, 0)]
    assert len(calls) == 1


def test_private_recovery_treats_reused_pid_as_old_process_gone(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    _process_listing(monkeypatch)
    replacement = _identity(start_seconds=101, start_microseconds=0)
    reads = []
    monkeypatch.setattr(runtime, "_process_exists", lambda pid: True, raising=False)
    monkeypatch.setattr(
        runtime,
        "_read_excel_process_identity",
        lambda pid: reads.append(pid) or replacement,
        raising=False,
    )

    assert runtime.recover_quarantine("spreadsheet") is True

    assert reads == [4242]
    assert not quarantine.exists()
    assert stage.is_dir() and owned.read_bytes() == b"retained workbook"
    assert recovery.is_file()


def test_private_recovery_refuses_same_birth_with_wrong_executable(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    _process_listing(monkeypatch)
    monkeypatch.setattr(runtime, "_process_exists", lambda pid: True, raising=False)
    monkeypatch.setattr(
        runtime,
        "_read_excel_process_identity",
        lambda pid: _identity(executable="/tmp/Microsoft Excel"),
        raising=False,
    )

    with pytest.raises(RuntimeError, match="identity is not verified"):
        runtime.recover_quarantine("spreadsheet")

    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


def test_private_recovery_propagates_cancellation_from_identity_reader(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    _process_listing(monkeypatch)
    monkeypatch.setattr(runtime, "_process_exists", lambda pid: True, raising=False)

    def cancel(pid):
        raise KeyboardInterrupt

    monkeypatch.setattr(runtime, "_read_excel_process_identity", cancel, raising=False)

    with pytest.raises(KeyboardInterrupt):
        runtime.recover_quarantine("spreadsheet")

    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


@pytest.mark.parametrize("state", [None, "permission", "unreadable"])
def test_private_recovery_refuses_unknown_process_state(tmp_path, monkeypatch, state):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    _process_listing(monkeypatch)
    monkeypatch.setattr(
        runtime,
        "_process_exists",
        (lambda pid: None) if state is None else (lambda pid: True),
        raising=False,
    )

    def read(pid):
        if state == "permission":
            raise PermissionError("denied")
        return None

    monkeypatch.setattr(runtime, "_read_excel_process_identity", read, raising=False)

    with pytest.raises(RuntimeError, match="identity is not verified"):
        runtime.recover_quarantine("spreadsheet")

    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


@pytest.mark.parametrize(
    "private_process",
    [
        None,
        {"identity": None, "launch_candidate": {"pid": 4242}},
        {"identity": {"pid": 4242, "start_seconds": 100}},
    ],
)
def test_private_recovery_rejects_missing_or_partial_owner_identity(
    tmp_path, monkeypatch, private_process
):
    _, stage, owned, recovery, quarantine = _quarantine(
        tmp_path, monkeypatch, private_process=private_process
    )
    calls = _process_listing(monkeypatch)
    monkeypatch.setattr(
        runtime,
        "_process_exists",
        lambda pid: pytest.fail("invalid owner identity must be rejected first"),
        raising=False,
    )

    with pytest.raises(ValueError, match="private process identity is invalid"):
        runtime.recover_quarantine("spreadsheet")

    assert calls == []
    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


def test_private_recovery_refuses_active_pid_bound_helper_for_stage(tmp_path, monkeypatch):
    _, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    _process_listing(
        monkeypatch,
        f"777 python3 -m skills.WPSComposer.scripts.msoffice.macos_osa_transport "
        f"--helper --request-file {stage / 'request.json'}\n",
    )
    monkeypatch.setattr(
        runtime,
        "_process_exists",
        lambda pid: pytest.fail("active helper must be rejected before PID recovery"),
        raising=False,
    )

    with pytest.raises(RuntimeError, match="native script is still running"):
        runtime.recover_quarantine("spreadsheet")

    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


def test_private_recovery_rejects_stage_mismatch_before_process_checks(tmp_path, monkeypatch):
    root, stage, owned, recovery, quarantine = _quarantine(tmp_path, monkeypatch)
    detail = json.loads(quarantine.read_text(encoding="utf-8"))
    detail["private_process"]["owned_path"] = str(root / "other-session" / "owned.xlsx")
    quarantine.write_text(json.dumps(detail) + "\n", encoding="utf-8")
    calls = _process_listing(monkeypatch)
    monkeypatch.setattr(
        runtime,
        "_process_exists",
        lambda pid: pytest.fail("stage mismatch must be rejected first"),
        raising=False,
    )

    with pytest.raises(ValueError, match="private process staging identity is invalid"):
        runtime.recover_quarantine("spreadsheet")

    assert calls == []
    assert quarantine.is_file()
    assert stage.is_dir() and owned.is_file() and recovery.is_file()


def test_kernel_process_existence_probe_uses_signal_zero_only():
    calls = []

    def signal(pid, value):
        calls.append((pid, value))

    assert runtime._process_exists(4242, _signal=signal) is True
    assert calls == [(4242, 0)]
