from __future__ import annotations

import json
from pathlib import Path
import subprocess
import ctypes as C

import pytest


def api():
    from skills.WPSComposer.scripts.msoffice import macos_osa_transport

    return macos_osa_transport


def identity(**changes):
    values = {
        "pid": 4242,
        "start_seconds": 100,
        "start_microseconds": 250,
        "executable": "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel",
        "bundle_id": "com.microsoft.Excel",
    }
    values.update(changes)
    return api().ExcelProcessIdentity(**values)


class FakeProcessLookup:
    def __init__(self, snapshots):
        self.snapshots = snapshots

    def snapshot(self, pid):
        return self.snapshots.get(pid)


class FakeNativeAdapter:
    def __init__(self, target, *, invoke_result=0, fail_at=None, fail_dispose=None):
        self.target = target
        self.invoke_result = invoke_result
        self.fail_at = fail_at
        self.fail_dispose = fail_dispose
        self.disposed = []
        self.replaced_pid = None
        self.invoked_event = None

    def get_address(self, event):
        if self.fail_at == "get_address":
            raise api().OSATransportError("OSA_ADDRESS_READ_FAILED")
        return "address"

    def decode_target(self, descriptor):
        if self.fail_at == "decode_target":
            raise api().OSATransportError("OSA_ADDRESS_READ_FAILED")
        return self.target

    def duplicate_event(self, event):
        if self.fail_at == "duplicate":
            raise api().OSATransportError("OSA_EVENT_COPY_FAILED")
        return "copy"

    def replace_address_with_pid(self, event, pid):
        if self.fail_at == "replace":
            raise api().OSATransportError("OSA_ADDRESS_WRITE_FAILED")
        self.replaced_pid = pid

    def invoke_original(self, event, reply, mode, priority, timeout, idle, filter_):
        self.invoked_event = event
        self.invoked_timeout = timeout
        return self.invoke_result

    def dispose(self, descriptor):
        self.disposed.append(descriptor)
        if descriptor == self.fail_dispose:
            raise api().OSATransportError("OSA_DESCRIPTOR_DISPOSE_FAILED")


def snapshots(*, owned=None, foreground=None):
    owned = owned or identity()
    foreground = foreground or identity(
        pid=9000, start_seconds=90, start_microseconds=0
    )
    return {
        owned.pid: owned,
        foreground.pid: foreground,
    }


def test_identity_rejects_pid_reuse_and_executable_mismatch():
    mod = api()
    expected = identity()

    with pytest.raises(mod.OSATransportError) as reused:
        mod.require_same_process(
            expected, identity(start_seconds=101, start_microseconds=0)
        )
    assert reused.value.code == "OSA_PROCESS_IDENTITY_CHANGED"

    with pytest.raises(mod.OSATransportError) as wrong_path:
        mod.require_same_process(expected, identity(executable="/tmp/Microsoft Excel"))
    assert wrong_path.value.code == "OSA_PROCESS_EXECUTABLE_CHANGED"


def test_darwin_ctypes_structures_match_64_bit_sdk_layout():
    mod = api()

    assert C.sizeof(mod._AEDesc) == 16
    assert C.sizeof(mod._ProcessSerialNumber) == 8
    assert C.sizeof(mod._ProcBSDInfo) == 136
    assert mod._ProcBSDInfo.start_seconds.offset == 120
    assert mod._ProcBSDInfo.start_microseconds.offset == 128


def test_carbon_function_bindings_use_sdk_return_and_size_types():
    mod = api()

    class Function:
        def __call__(self, *args):
            return 0

    class Library:
        def __getattr__(self, name):
            function = Function()
            setattr(self, name, function)
            return function

    runtime = mod._DarwinRuntime.__new__(mod._DarwinRuntime)
    runtime.carbon = Library()
    runtime._configure_carbon()

    assert runtime.ae_get_attribute.restype is C.c_int16
    assert runtime.ae_get_data.argtypes[-1] is C.c_long
    assert runtime.ae_put_attribute.argtypes[-1] is C.c_long
    assert runtime.ae_get_size.restype is C.c_long
    assert runtime.get_process_pid.restype is C.c_int32
    assert runtime.osa_get_send.restype is C.c_int32


@pytest.mark.parametrize("requested", [-2, -1, 6000])
def test_event_timeout_is_capped_by_absolute_deadline(requested):
    assert api().finite_timeout_ticks(requested, deadline=12.5, now=10.0) == 150


def test_expired_event_deadline_is_structured_timeout():
    with pytest.raises(api().OSATransportError) as caught:
        api().finite_timeout_ticks(-2, deadline=10.0, now=10.0)
    assert caught.value.code == "OSA_TIMEOUT"
    assert caught.value.outcome_uncertain is True


def test_current_helper_event_is_forwarded_without_retargeting():
    mod = api()
    adapter = FakeNativeAdapter(mod.AppleEventTarget.current_process())
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    assert router.send(adapter, "event", "reply", 3, 0, -2, None, None, 12.5) == 0
    assert adapter.invoked_event == "event"
    assert adapter.replaced_pid is None
    assert adapter.disposed == ["address"]
    assert adapter.invoked_timeout == 150


def test_excel_pid_target_is_rewritten_only_after_both_processes_validate():
    mod = api()
    adapter = FakeNativeAdapter(mod.AppleEventTarget.kernel_pid(9000))
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    assert router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5) == 0
    assert adapter.replaced_pid == 4242
    assert adapter.invoked_event == "copy"
    assert adapter.disposed == ["address", "copy"]


def test_excel_bundle_target_is_rewritten_to_validated_owned_process():
    mod = api()
    adapter = FakeNativeAdapter(mod.AppleEventTarget.bundle("com.microsoft.Excel"))
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    assert router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5) == 0
    assert adapter.replaced_pid == 4242


@pytest.mark.parametrize(
    "target",
    [
        pytest.param(lambda m: m.AppleEventTarget.bundle("com.apple.systemevents"), id="foreign-bundle"),
        pytest.param(lambda m: m.AppleEventTarget.kernel_pid(12345), id="missing-pid"),
        pytest.param(lambda m: m.AppleEventTarget.unknown("sign", b"XCEL"), id="unknown-kind"),
    ],
)
def test_unknown_or_foreign_external_target_is_rejected_and_disposed(target):
    mod = api()
    adapter = FakeNativeAdapter(target(mod))
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    with pytest.raises(mod.OSATransportError) as caught:
        router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5)
    assert caught.value.code == "OSA_TARGET_REJECTED"
    assert adapter.invoked_event is None
    assert adapter.disposed == ["address"]


def test_owned_pid_reuse_is_rejected_before_event_copy():
    mod = api()
    current = snapshots()
    current[4242] = identity(start_microseconds=251)
    adapter = FakeNativeAdapter(mod.AppleEventTarget.bundle("com.microsoft.Excel"))
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(current), helper_pid=777, clock=lambda: 10.0
    )

    with pytest.raises(mod.OSATransportError) as caught:
        router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5)
    assert caught.value.code == "OSA_PROCESS_IDENTITY_CHANGED"
    assert adapter.replaced_pid is None
    assert adapter.disposed == ["address"]


def test_native_process_lookup_rejects_reuse_during_identity_read():
    mod = api()

    class Runtime:
        def __init__(self):
            self.starts = iter([(100, 250), (101, 0)])

        def proc_pidinfo(self, pid, flavor, arg, output, size):
            seconds, micros = next(self.starts)
            info = C.cast(output, C.POINTER(mod._ProcBSDInfo)).contents
            info.pid = pid
            info.start_seconds = seconds
            info.start_microseconds = micros
            return C.sizeof(mod._ProcBSDInfo)

        def proc_pidpath(self, pid, output, size):
            value = identity().executable.encode() + b"\0"
            C.memmove(output, value, len(value))
            return len(value) - 1

        def objc_class(self, name):
            return "NSRunningApplication"

        def message(self, receiver, name, args=(), types=(), result=None):
            return {
                "runningApplicationWithProcessIdentifier:": "app",
                "bundleIdentifier": "bundle",
                "executableURL": "url",
                "path": "path",
            }[name]

        def object_text(self, value):
            return {
                "bundle": "com.microsoft.Excel",
                "path": identity().executable,
            }[value]

    with pytest.raises(api().OSATransportError) as caught:
        api()._DarwinProcessLookup(Runtime()).snapshot(4242)
    assert caught.value.code == "OSA_PROCESS_IDENTITY_CHANGED"


def test_native_process_lookup_uses_supplied_retained_application_handle():
    mod = api()

    class Runtime:
        def __init__(self):
            self.process_reads = 0
            self.messages = []

        def proc_pidinfo(self, pid, flavor, arg, output, size):
            self.process_reads += 1
            info = C.cast(output, C.POINTER(mod._ProcBSDInfo)).contents
            info.pid = pid
            info.start_seconds = 100
            info.start_microseconds = 250
            return C.sizeof(mod._ProcBSDInfo)

        def proc_pidpath(self, pid, output, size):
            value = identity().executable.encode() + b"\0"
            C.memmove(output, value, len(value))
            return len(value) - 1

        def objc_class(self, name):
            pytest.fail("a retained application must not be rediscovered by PID")

        def message(self, receiver, name, args=(), types=(), result=None):
            self.messages.append((receiver, name))
            return {
                ("retained-app", "processIdentifier"): 4242,
                ("retained-app", "bundleIdentifier"): "bundle",
                ("retained-app", "executableURL"): "url",
                ("url", "path"): "path",
            }[(receiver, name)]

        def object_text(self, value):
            return {
                "bundle": "com.microsoft.Excel",
                "path": identity().executable,
            }[value]

    runtime = Runtime()
    observed = mod._DarwinProcessLookup(runtime).snapshot(
        4242, application="retained-app"
    )

    assert observed == identity()
    assert runtime.process_reads == 2
    assert runtime.messages[0] == ("retained-app", "processIdentifier")


def test_native_process_lookup_rejects_retained_application_for_another_pid():
    mod = api()

    class Runtime:
        def proc_pidinfo(self, pid, flavor, arg, output, size):
            info = C.cast(output, C.POINTER(mod._ProcBSDInfo)).contents
            info.pid = pid
            info.start_seconds = 100
            info.start_microseconds = 250
            return C.sizeof(mod._ProcBSDInfo)

        def proc_pidpath(self, pid, output, size):
            value = identity().executable.encode() + b"\0"
            C.memmove(output, value, len(value))
            return len(value) - 1

        def message(self, receiver, name, args=(), types=(), result=None):
            assert (receiver, name) == ("retained-app", "processIdentifier")
            return 9000

    with pytest.raises(mod.OSATransportError) as caught:
        mod._DarwinProcessLookup(Runtime()).snapshot(
            4242, application="retained-app"
        )

    assert caught.value.code == "OSA_PROCESS_IDENTITY_CHANGED"
    assert caught.value.pid == 4242


@pytest.mark.parametrize("fail_at", ["decode_target", "duplicate", "replace"])
def test_descriptor_cleanup_survives_adapter_errors(fail_at):
    mod = api()
    adapter = FakeNativeAdapter(
        mod.AppleEventTarget.bundle("com.microsoft.Excel"), fail_at=fail_at
    )
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    with pytest.raises(mod.OSATransportError):
        router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5)
    assert adapter.disposed == (["address", "copy"] if fail_at == "replace" else ["address"])


def test_descriptor_cleanup_failure_does_not_mask_primary_routing_error():
    mod = api()
    adapter = FakeNativeAdapter(
        mod.AppleEventTarget.bundle("com.apple.systemevents"),
        fail_dispose="address",
    )
    router = mod.BoundSendRouter(
        identity(), FakeProcessLookup(snapshots()), helper_pid=777, clock=lambda: 10.0
    )

    with pytest.raises(mod.OSATransportError) as caught:
        router.send(adapter, "event", "reply", 3, 0, 500, None, None, 12.5)
    assert caught.value.code == "OSA_TARGET_REJECTED"
    assert [error.code for error in caught.value.cleanup_errors] == [
        "OSA_DESCRIPTOR_DISPOSE_FAILED"
    ]


def test_original_send_callback_is_restored_when_execution_raises():
    mod = api()
    calls = []

    def set_send(callback, refcon):
        calls.append((callback, refcon))

    lease = mod.SendCallbackLease(set_send, original="original", original_refcon=73)
    with pytest.raises(RuntimeError, match="script failed"):
        with lease.installed("bound"):
            raise RuntimeError("script failed")

    assert calls == [("bound", 0), ("original", 73)]


def test_callback_bridge_preserves_structured_router_failure():
    mod = api()

    class Router:
        def send(self, *args):
            raise mod.OSATransportError(
                "OSA_PROCESS_IDENTITY_CHANGED", pid=4242
            )

    bridge = mod.SendCallbackBridge(Router(), adapter=object(), deadline=12.5)

    assert bridge("event", "reply", 3, 0, 500, None, None, 0) == -600
    assert bridge.error.code == "OSA_PROCESS_IDENTITY_CHANGED"
    assert bridge.error.pid == 4242


def test_callback_bridge_retains_nonzero_native_send_status_and_stops_later_sends():
    mod = api()

    class Router:
        identity = identity()

        def __init__(self):
            self.calls = []

        def send(self, *args):
            self.calls.append(args)
            return -1712

    router = Router()
    bridge = mod.SendCallbackBridge(router, adapter=object(), deadline=12.5)

    first = bridge("first", "reply", 3, 0, 500, None, None, 0)
    second = bridge("second", "reply", 3, 0, 500, None, None, 0)

    assert first == second == -1712
    assert [call[1] for call in router.calls] == ["first"]
    assert isinstance(bridge.error, mod.OSATransportError)
    assert bridge.error.code == "OSA_SEND_FAILED"
    assert bridge.error.pid == 4242
    assert bridge.error.outcome_uncertain is True
    assert bridge.error.detail == "-1712"


def test_callback_bridge_leaves_application_reply_errors_to_native_semantics():
    mod = api()

    class Router:
        identity = identity()

        def __init__(self):
            self.calls = 0

        def send(self, *args):
            self.calls += 1
            return 0

    router = Router()
    bridge = mod.SendCallbackBridge(router, adapter=object(), deadline=12.5)
    reply = {"keyErrorNumber": -1728}

    assert bridge("event", reply, 3, 0, 500, None, None, 0) == 0
    assert router.calls == 1
    assert bridge.error is None


def test_run_translates_caller_deadline_to_shared_helper_clock_without_extension(tmp_path):
    mod = api()
    script = tmp_path / "step.applescript"
    script.write_text('return "{}"')
    seen = {}

    def runner(command, **kwargs):
        seen.update(command=command, kwargs=kwargs)
        return subprocess.CompletedProcess(command, 7, "stdout", "stderr")

    caller_now = [10.0]

    def shared_clock():
        caller_now[0] = 10.5
        return 1_000.0

    transport = mod.BoundOSAKitTransport(
        identity(),
        runner=runner,
        clock=lambda: caller_now[0],
        shared_clock=shared_clock,
        helper_path=Path("/helper.py"),
    )
    result = transport.run(script, deadline=12.5)

    assert result.returncode == 7 and result.stdout == "stdout" and result.stderr == "stderr"
    assert seen["kwargs"] == {"capture_output": True, "text": True, "timeout": 2.0}
    request = json.loads(seen["command"][2])
    assert request == {
        "pid": 4242,
        "start_seconds": 100,
        "start_microseconds": 250,
        "executable": "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel",
        "bundle_id": "com.microsoft.Excel",
        "shared_deadline": 1_002.0,
    }
    assert seen["command"][3:] == ["--helper", str(script.resolve())]


def test_run_converts_outer_timeout_to_structured_uncertain_error(tmp_path):
    mod = api()
    script = tmp_path / "step.applescript"
    script.write_text('return "{}"')

    def timeout(command, **kwargs):
        raise subprocess.TimeoutExpired(command, kwargs["timeout"], output="partial", stderr="late")

    transport = mod.BoundOSAKitTransport(
        identity(),
        runner=timeout,
        clock=lambda: 10.0,
        shared_clock=lambda: 1_000.0,
        helper_path=Path("/helper.py"),
    )
    with pytest.raises(mod.OSATransportError) as caught:
        transport.run(script, deadline=12.5)
    assert caught.value.code == "OSA_TIMEOUT"
    assert caught.value.stdout == "partial" and caught.value.stderr == "late"
    assert caught.value.pid == 4242 and caught.value.outcome_uncertain is True


def test_run_converts_helper_spawn_failure_to_structured_error(tmp_path):
    mod = api()
    script = tmp_path / "step.applescript"
    script.write_text('return "{}"')

    def unavailable(*args, **kwargs):
        raise OSError("interpreter unavailable")

    transport = mod.BoundOSAKitTransport(
        identity(), runner=unavailable, clock=lambda: 10.0, shared_clock=lambda: 1_000.0
    )
    with pytest.raises(mod.OSATransportError) as caught:
        transport.run(script, deadline=12.5)
    assert caught.value.code == "OSA_HELPER_LAUNCH_FAILED"
    assert caught.value.pid == 4242 and caught.value.outcome_uncertain is False


def test_transport_rejects_non_excel_identity_before_spawning():
    mod = api()
    with pytest.raises(mod.OSATransportError) as caught:
        mod.BoundOSAKitTransport(identity(bundle_id="com.apple.TextEdit"))
    assert caught.value.code == "OSA_NOT_EXCEL_PROCESS"


def test_run_rejects_missing_script_and_elapsed_deadline_without_spawning(tmp_path):
    mod = api()
    transport = mod.BoundOSAKitTransport(
        identity(),
        runner=lambda *a, **k: pytest.fail("must not spawn"),
        clock=lambda: 10.0,
        shared_clock=lambda: 1_000.0,
    )
    with pytest.raises(FileNotFoundError):
        transport.run(tmp_path / "missing.applescript", deadline=12.5)
    script = tmp_path / "step.applescript"
    script.write_text('return "{}"')
    with pytest.raises(mod.OSATransportError) as caught:
        transport.run(script, deadline=10.0)
    assert caught.value.code == "OSA_TIMEOUT"


def test_helper_request_round_trip_preserves_shared_absolute_deadline():
    mod = api()
    payload = identity().to_request(shared_deadline=1_003.25)
    restored = mod.HelperRequest.from_json(json.dumps(payload))

    assert restored.identity == identity()
    assert restored.shared_deadline == 1_003.25


def test_helper_router_uses_shared_clock_even_when_process_monotonic_epochs_differ(
    monkeypatch,
):
    mod = api()
    monkeypatch.setattr(mod, "_shared_monotonic", lambda: 1_000.0)
    router = mod._helper_router(identity(), object(), helper_pid=777)

    assert router.clock() == 1_000.0


def test_cli_argument_validation_does_not_load_native_runtime(capsys):
    assert api().main([]) == 2
    assert "usage:" in capsys.readouterr().err
