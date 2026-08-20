from __future__ import annotations

import importlib
import json
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest


def test_windows_module_import_is_safe_without_pywin32(monkeypatch):
    for name in ("pythoncom", "win32api", "win32process", "win32com.client"):
        monkeypatch.delitem(sys.modules, name, raising=False)

    module = importlib.import_module(
        "skills.WPSComposer.scripts.longform_m0.windows"
    )

    assert "win32com.client" not in sys.modules
    assert module.PROBE_VERSION == "0.8.0-m0.1"


def test_dispatch_owned_application_requires_new_dispatch_ex_identity():
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _dispatch_owned_application,
    )

    calls = []
    app = SimpleNamespace(Hwnd=9001)
    client = SimpleNamespace(
        DispatchEx=lambda progid: calls.append(("DispatchEx", progid)) or app,
        Dispatch=lambda progid: pytest.fail("shared Dispatch must never run"),
    )
    identity = ProcessIdentity(
        pid=202,
        executable=r"C:\Program Files\Kingsoft\WPS Office\office6\wps.exe",
        created_ns=123456789,
        parent_pid=101,
    )

    owned = _dispatch_owned_application(
        client,
        hwnd_pid=lambda hwnd: 202 if hwnd == 9001 else 0,
        read_identity=lambda pid: identity if pid == 202 else None,
        preexisting_pids={101},
    )

    assert owned.app is app
    assert owned.identity == identity
    assert calls == [("DispatchEx", "kwps.Application")]


@pytest.mark.parametrize("identity", [None, "preexisting"])
def test_dispatch_owned_application_fails_without_provable_new_process(identity):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        WindowsOwnershipError,
        _dispatch_owned_application,
    )

    actual = (
        None
        if identity is None
        else ProcessIdentity(202, r"C:\wps.exe", 1, 100)
    )
    client = SimpleNamespace(DispatchEx=lambda progid: SimpleNamespace(Hwnd=7))

    with pytest.raises(WindowsOwnershipError):
        _dispatch_owned_application(
            client,
            hwnd_pid=lambda hwnd: 202,
            read_identity=lambda pid: actual,
            preexisting_pids={202} if identity == "preexisting" else set(),
        )


def test_timeout_cancels_then_terminates_only_matching_owned_identity(tmp_path: Path):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _cancel_then_terminate,
    )

    identity = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    worker = SimpleNamespace()
    waits = []

    def wait(timeout):
        waits.append(timeout)
        raise TimeoutError("still running")

    worker.wait = wait
    terminated = []
    cancel_path = tmp_path / "cancel.json"

    _cancel_then_terminate(
        worker,
        cancel_path,
        identity,
        read_identity=lambda pid: identity,
        terminate_tree=lambda current: terminated.append(current),
        grace_seconds=5,
    )

    assert cancel_path.read_text(encoding="utf-8") == '{"cancel":true}\n'
    assert waits == [5]
    assert terminated == [identity]


def test_timeout_never_terminates_reused_or_unowned_pid(tmp_path: Path):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _cancel_then_terminate,
    )

    owned = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    reused = ProcessIdentity(202, r"C:\wps.exe", 100, 101)
    worker = SimpleNamespace(wait=lambda timeout: (_ for _ in ()).throw(TimeoutError()))
    terminated = []

    _cancel_then_terminate(
        worker,
        tmp_path / "cancel.json",
        owned,
        read_identity=lambda pid: reused,
        terminate_tree=lambda current: terminated.append(current),
        grace_seconds=5,
    )

    assert terminated == []


def test_worker_request_and_progress_reject_unknown_or_private_fields():
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _validate_progress,
        _validate_worker_request,
    )

    request = {
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "workerMode": "probe",
        "ownershipTimeoutVerified": True,
        "stagedDocxPath": r"C:\private\probe.docx",
        "stagedPdfPath": r"C:\private\probe.pdf",
        "stagedSvgPath": r"C:\private\probe.svg",
        "expectedWpsVersion": "12.1.fake",
        "preexistingPids": [101],
        "cancelPath": r"C:\private\cancel.json",
        "ackPath": r"C:\private\ack.json",
    }
    validated = _validate_worker_request(request)
    assert validated["preexistingPids"] == (101,)
    with pytest.raises(ValueError, match="fields"):
        _validate_worker_request({**request, "unexpected": True})
    with pytest.raises(ValueError, match="timeout verification"):
        _validate_worker_request({**request, "ownershipTimeoutVerified": False})
    fixture = _validate_worker_request(
        {
            **request,
            "workerMode": "ownership-timeout",
            "ownershipTimeoutVerified": False,
        }
    )
    assert fixture["workerMode"] == "ownership-timeout"

    identity = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    assert _validate_progress(identity.to_json(), {101}) == identity
    with pytest.raises(ValueError, match="pre-existing"):
        _validate_progress(identity.to_json(), {101, 202})


def test_supervisor_accepts_result_only_after_live_identity_and_worker_exit(
    tmp_path: Path,
):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _supervise_worker,
    )

    identity = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    progress = tmp_path / "progress.json"
    result = tmp_path / "result.json"
    cancel = tmp_path / "cancel.json"
    ack = tmp_path / "ack.json"
    progress.write_text(json.dumps(identity.to_json()), encoding="utf-8")
    result.write_text(json.dumps({"ok": True}), encoding="utf-8")
    waits = []
    worker = SimpleNamespace(
        poll=lambda: None,
        wait=lambda timeout: waits.append(timeout) or 0,
    )

    value, owned = _supervise_worker(
        worker,
        progress,
        result,
        cancel,
        ack,
        preexisting_pids={101},
        read_identity=lambda pid: identity,
        terminate_tree=lambda current: pytest.fail("must not terminate"),
        timeout=30,
        sleep=lambda seconds: None,
    )

    assert value == {"ok": True}
    assert owned == identity
    assert waits and 0 < waits[0] <= 5
    assert not cancel.exists()
    assert ack.read_text(encoding="utf-8") == '{"accepted":true}\n'


def test_supervisor_rejects_identity_change_before_consuming_result(tmp_path: Path):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        WindowsOwnershipError,
        _supervise_worker,
    )

    claimed = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    reused = ProcessIdentity(202, r"C:\wps.exe", 100, 101)
    progress = tmp_path / "progress.json"
    result = tmp_path / "result.json"
    progress.write_text(json.dumps(claimed.to_json()), encoding="utf-8")
    result.write_text(json.dumps({"ok": True}), encoding="utf-8")

    with pytest.raises(WindowsOwnershipError, match="changed"):
        _supervise_worker(
            SimpleNamespace(poll=lambda: None, wait=lambda timeout: 0),
            progress,
            result,
            tmp_path / "cancel.json",
            tmp_path / "ack.json",
            preexisting_pids=set(),
            read_identity=lambda pid: reused,
            terminate_tree=lambda current: pytest.fail("must not terminate"),
            timeout=30,
            sleep=lambda seconds: None,
        )


def test_supervisor_rejects_worker_exit_without_result(tmp_path: Path):
    from skills.WPSComposer.scripts.longform_m0.windows import _supervise_worker

    with pytest.raises(RuntimeError, match="without a result"):
        _supervise_worker(
            SimpleNamespace(poll=lambda: 7, wait=lambda timeout: 7),
            tmp_path / "progress.json",
            tmp_path / "result.json",
            tmp_path / "cancel.json",
            tmp_path / "ack.json",
            preexisting_pids=set(),
            read_identity=lambda pid: None,
            terminate_tree=lambda current: None,
            timeout=30,
            sleep=lambda seconds: None,
        )


def test_worker_balances_com_reports_ownership_and_quits_only_owned_app(
    tmp_path: Path,
):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _worker_execute,
    )

    calls = []

    class App:
        Hwnd = 44
        Visible = -1
        DisplayAlerts = -1

        def Quit(self):
            calls.append("quit")

    app = App()
    identity = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    client = SimpleNamespace(
        DispatchEx=lambda progid: calls.append(("dispatch_ex", progid)) or app,
        Dispatch=lambda progid: pytest.fail("Dispatch must not be used"),
    )
    native = SimpleNamespace(
        pythoncom=SimpleNamespace(
            CoInitialize=lambda: calls.append("init"),
            CoUninitialize=lambda: calls.append("uninit"),
        ),
        client=client,
        hwnd_pid=lambda hwnd: 202,
        read_identity=lambda pid: identity,
    )
    request = {
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "workerMode": "probe",
        "ownershipTimeoutVerified": True,
        "stagedDocxPath": r"C:\private\probe.docx",
        "stagedPdfPath": r"C:\private\probe.pdf",
        "stagedSvgPath": r"C:\private\probe.svg",
        "expectedWpsVersion": "12.1.fake",
        "preexistingPids": [101],
        "cancelPath": r"C:\private\cancel.json",
        "ackPath": r"C:\private\ack.json",
    }

    value = _worker_execute(
        request,
        tmp_path / "progress.json",
        native=native,
        assay_runner=lambda owned_app, validated, canceled: {
            "ok": owned_app is app,
            "canceled": canceled(),
        },
    )

    assert value == {"ok": True, "canceled": False}
    assert json.loads((tmp_path / "progress.json").read_text()) == identity.to_json()
    assert calls == [
        "init",
        ("dispatch_ex", "kwps.Application"),
        "quit",
        "uninit",
    ]
    assert app.Visible == 0
    assert app.DisplayAlerts == 0


def test_worker_interrupt_still_quits_owned_app_and_uninitializes(tmp_path: Path):
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _worker_execute,
    )

    calls = []
    app = SimpleNamespace(
        Hwnd=44,
        Quit=lambda: calls.append("quit"),
        Visible=-1,
        DisplayAlerts=-1,
    )
    identity = ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    native = SimpleNamespace(
        pythoncom=SimpleNamespace(
            CoInitialize=lambda: calls.append("init"),
            CoUninitialize=lambda: calls.append("uninit"),
        ),
        client=SimpleNamespace(DispatchEx=lambda progid: app),
        hwnd_pid=lambda hwnd: 202,
        read_identity=lambda pid: identity,
    )
    request = {
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "workerMode": "probe",
        "ownershipTimeoutVerified": True,
        "stagedDocxPath": r"C:\private\probe.docx",
        "stagedPdfPath": r"C:\private\probe.pdf",
        "stagedSvgPath": r"C:\private\probe.svg",
        "expectedWpsVersion": "12.1.fake",
        "preexistingPids": [101],
        "cancelPath": r"C:\private\cancel.json",
        "ackPath": r"C:\private\ack.json",
    }

    with pytest.raises(KeyboardInterrupt):
        _worker_execute(
            request,
            tmp_path / "progress.json",
            native=native,
            assay_runner=lambda *args: (_ for _ in ()).throw(KeyboardInterrupt()),
        )

    assert calls == ["init", "quit", "uninit"]


def test_windows_native_assay_contains_closed_com_capabilities_and_lifecycle():
    module_path = Path(
        "skills/WPSComposer/scripts/longform_m0/windows.py"
    )
    source = module_path.read_text(encoding="utf-8")

    for required in (
        "Documents.Add()",
        "SaveAs2(request[\"stagedDocxPath\"], 12)",
        "Documents.Open(request[\"stagedDocxPath\"]",
        "document.Fields.Update()",
        "ExportAsFixedFormat(request[\"stagedPdfPath\"], 17",
        "ListTemplates.Add",
        "ListLevelNumber",
        "TablesOfContents.Add",
        "TablesOfFigures.Add",
        "Bookmarks.Add",
        "OMaths.Add",
        "InlineShapes.AddPicture",
        "Repaginate()",
    ):
        assert required in source
    assert source.count("_com_assay(document, capabilities,") == 13
    assert ".Dispatch(" not in source


def test_worker_publishes_result_before_quit_and_waits_for_supervisor_ack(
    tmp_path: Path,
):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    calls = []
    app = SimpleNamespace(
        Hwnd=44,
        Quit=lambda: calls.append("quit"),
        Visible=-1,
        DisplayAlerts=-1,
    )
    identity = windows.ProcessIdentity(202, r"C:\wps.exe", 99, 101)
    native = SimpleNamespace(
        pythoncom=SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
        client=SimpleNamespace(DispatchEx=lambda progid: app),
        hwnd_pid=lambda hwnd: 202,
        read_identity=lambda pid: identity,
    )
    request = {
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "workerMode": "probe",
        "ownershipTimeoutVerified": True,
        "stagedDocxPath": r"C:\private\probe.docx",
        "stagedPdfPath": r"C:\private\probe.pdf",
        "stagedSvgPath": r"C:\private\probe.svg",
        "expectedWpsVersion": "12.1.fake",
        "preexistingPids": [101],
        "cancelPath": r"C:\private\cancel.json",
        "ackPath": r"C:\private\ack.json",
    }
    result_path = tmp_path / "result.json"
    observed = []

    def sleep(_seconds):
        observed.append(json.loads(result_path.read_text(encoding="utf-8")))
        (tmp_path / "ack.json").write_text('{"accepted":true}\n', encoding="utf-8")

    value = windows._worker_execute(
        request,
        tmp_path / "progress.json",
        result_path=result_path,
        ack_path=tmp_path / "ack.json",
        cancel_path=tmp_path / "cancel.json",
        native=native,
        assay_runner=lambda *args: {"ok": True},
        sleep=sleep,
    )

    assert value == {"ok": True}
    assert observed == [{"ok": True}]
    assert calls == ["quit"]


def test_timeout_fixture_requires_owned_progress_and_preserves_preexisting_processes(
    tmp_path: Path,
):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    preexisting = windows.ProcessIdentity(101, r"c:\wps\wps.exe", 1, 10)
    fixture = windows.ProcessIdentity(202, r"c:\wps\wps.exe", 2, 20)
    progress = tmp_path / "progress.json"
    progress.write_text(json.dumps(fixture.to_json()), encoding="utf-8")
    worker = SimpleNamespace(poll=lambda: None, wait=lambda timeout: 0)
    ticks = iter((0.0, 0.0, 2.0, 2.0))

    with pytest.raises(windows.WindowsWorkerTimeout) as caught:
        windows._supervise_worker(
            worker,
            progress,
            tmp_path / "result.json",
            tmp_path / "cancel.json",
            tmp_path / "ack.json",
            preexisting_pids={101},
            read_identity=lambda pid: fixture if pid == 202 else preexisting,
            terminate_tree=lambda current: pytest.fail("cooperative exit expected"),
            timeout=1,
            sleep=lambda seconds: None,
            monotonic=lambda: next(ticks),
        )

    assert caught.value.identity == fixture
    windows._verify_preexisting_survived(
        {101: preexisting},
        read_identity=lambda pid: preexisting,
    )
    with pytest.raises(windows.WindowsOwnershipError, match="pre-existing"):
        windows._verify_preexisting_survived(
            {101: preexisting},
            read_identity=lambda pid: None,
        )


def test_read_windows_identity_uses_executable_creation_and_parent(monkeypatch):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    closed = []

    class Created:
        def timestamp(self):
            return 12.5

    win32api = SimpleNamespace(
        OpenProcess=lambda access, inherit, pid: "handle",
        CloseHandle=lambda handle: closed.append(handle),
    )
    win32process = SimpleNamespace(
        GetModuleFileNameEx=lambda handle, module: r"C:\WPS\wps.exe",
        GetProcessTimes=lambda handle: (Created(), None, None, None),
    )
    monkeypatch.setitem(sys.modules, "win32api", win32api)
    monkeypatch.setitem(sys.modules, "win32process", win32process)
    monkeypatch.setattr(windows, "_read_parent_pid", lambda pid: 101)

    identity = windows._read_windows_identity(202)

    assert identity == windows.ProcessIdentity(
        202, r"c:\wps\wps.exe", 12_500_000_000, 101
    )
    assert closed == ["handle"]


def test_snapshot_and_termination_are_closed_to_matching_identity():
    from skills.WPSComposer.scripts.longform_m0.windows import (
        ProcessIdentity,
        _snapshot_wps_pids,
        _terminate_windows_tree,
    )

    snapshot_result = SimpleNamespace(stdout="101\n202\n", returncode=0)
    assert _snapshot_wps_pids(
        runner=lambda *args, **kwargs: snapshot_result
    ) == {101, 202}

    owned = ProcessIdentity(202, r"c:\wps\wps.exe", 99, 101)
    reused = ProcessIdentity(202, r"c:\wps\wps.exe", 100, 101)
    calls = []
    _terminate_windows_tree(
        owned,
        read_identity=lambda pid: reused,
        runner=lambda *args, **kwargs: calls.append(args),
    )
    assert calls == []

    _terminate_windows_tree(
        owned,
        read_identity=lambda pid: owned,
        runner=lambda *args, **kwargs: calls.append((args, kwargs)),
    )
    assert calls[0][0][0] == ["taskkill.exe", "/PID", "202", "/T", "/F"]


def test_worker_result_is_closed_and_bound_to_staging(tmp_path: Path):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    docx = (tmp_path / "probe.docx").resolve()
    pdf = (tmp_path / "probe.pdf").resolve()
    raw = {
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "platform": "windows",
        "wpsVersion": "12.1.fake",
        "capabilities": [],
        "failures": [],
        "docxPath": str(docx),
        "pdfPath": str(pdf),
    }

    capabilities, failures, version = windows._validate_worker_result(
        raw, docx, pdf
    )
    assert capabilities == []
    assert failures == []
    assert version == "12.1.fake"
    with pytest.raises(ValueError, match="fields"):
        windows._validate_worker_result({**raw, "private": True}, docx, pdf)
    with pytest.raises(ValueError, match="staging"):
        windows._validate_worker_result(
            {**raw, "pdfPath": str(tmp_path / "other.pdf")}, docx, pdf
        )


def test_non_windows_run_writes_redacted_failure_and_no_artifacts(
    monkeypatch, tmp_path: Path
):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    monkeypatch.setattr(windows.platform, "system", lambda: "Darwin")
    output = tmp_path / "windows-evidence"

    with pytest.raises(windows.WindowsM0Failed) as caught:
        windows.run_windows_probe(output)

    assert caught.value.evidence_path == output / "platform-evidence.json"
    evidence = json.loads(caught.value.evidence_path.read_text(encoding="utf-8"))
    assert evidence["platform"] == "windows"
    assert evidence["failures"][0]["code"] == "WINDOWS_REQUIRED"
    assert evidence["artifacts"] == {}
    assert not (output / "probe.docx").exists()
    assert not (output / "probe.pdf").exists()


def test_private_worker_entrypoint_requires_exact_arguments(monkeypatch, tmp_path: Path):
    import skills.WPSComposer.scripts.longform_m0.windows as windows

    assert windows._module_main(["windows.py"]) == 2
    observed = []
    monkeypatch.setattr(
        windows,
        "_worker_main",
        lambda request, progress, result: observed.append(
            (request, progress, result)
        )
        or 0,
    )
    args = [
        "windows.py",
        "--worker",
        str(tmp_path / "request.json"),
        str(tmp_path / "progress.json"),
        str(tmp_path / "result.json"),
    ]
    assert windows._module_main(args) == 0
    assert observed == [tuple(Path(value) for value in args[2:])]
