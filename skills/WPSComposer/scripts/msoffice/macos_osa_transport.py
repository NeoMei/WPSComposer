"""PID-bound, one-shot OSAKit execution for an owned Microsoft Excel process.

Importing this internal module is platform independent. Cocoa, Carbon, OSAKit,
and libproc are loaded only by the ``--helper`` CLI on macOS. Session wiring and
application launch ownership deliberately live outside this transport slice.
"""
from __future__ import annotations

import ctypes as C
from contextlib import contextmanager
from dataclasses import dataclass
import json
import math
import os
import posixpath
from pathlib import Path, PurePosixPath
import subprocess
import sys
import time
from typing import Any, Callable, Optional


_EXCEL_BUNDLE_ID = "com.microsoft.Excel"
_TICKS_PER_SECOND = 60
_MAX_TIMEOUT_TICKS = 2**31 - 1
_K_AE_DEFAULT_TIMEOUT = -1
_K_NO_TIMEOUT = -2
_ERR_AE_TIMEOUT = -1712
_ERR_AE_EVENT_NOT_PERMITTED = -1743
_PROC_NOT_FOUND = -600
_PARAM_ERR = -50


def _shared_monotonic() -> float:
    """Return the OS-wide monotonic clock shared by independent processes."""
    return time.clock_gettime(time.CLOCK_MONOTONIC)


class OSATransportError(RuntimeError):
    """A structured transport failure suitable for later quarantine wiring."""

    def __init__(
        self,
        code: str,
        *,
        stdout: str = "",
        stderr: str = "",
        pid: Optional[int] = None,
        outcome_uncertain: bool = False,
        detail: Optional[str] = None,
    ) -> None:
        super().__init__(code if detail is None else f"{code}: {detail}")
        self.code = code
        self.stdout = stdout
        self.stderr = stderr
        self.pid = pid
        self.outcome_uncertain = outcome_uncertain
        self.detail = detail

    def as_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "pid": self.pid,
            "outcome_uncertain": self.outcome_uncertain,
            "detail": self.detail,
        }


def _strict_int(value: Any, name: str, *, minimum: int = 0) -> int:
    if type(value) is not int or value < minimum:
        raise ValueError(f"Invalid {name}")
    return value


@dataclass(frozen=True)
class ExcelProcessIdentity:
    pid: int
    start_seconds: int
    start_microseconds: int
    executable: str
    bundle_id: str

    def __post_init__(self) -> None:
        _strict_int(self.pid, "pid", minimum=1)
        _strict_int(self.start_seconds, "process start seconds")
        micros = _strict_int(self.start_microseconds, "process start microseconds")
        if micros >= 1_000_000:
            raise ValueError("Invalid process start microseconds")
        if not isinstance(self.executable, str) or not PurePosixPath(self.executable).is_absolute():
            raise ValueError("Process executable must be absolute")
        if not isinstance(self.bundle_id, str) or not self.bundle_id:
            raise ValueError("Process bundle identifier is required")

    def to_request(self, *, shared_deadline: float) -> dict[str, Any]:
        if (
            isinstance(shared_deadline, bool)
            or not isinstance(shared_deadline, (int, float))
            or not math.isfinite(shared_deadline)
            or shared_deadline <= 0
        ):
            raise ValueError("Invalid helper deadline")
        return {
            "pid": self.pid,
            "start_seconds": self.start_seconds,
            "start_microseconds": self.start_microseconds,
            "executable": self.executable,
            "bundle_id": self.bundle_id,
            "shared_deadline": float(shared_deadline),
        }


@dataclass(frozen=True)
class HelperRequest:
    identity: ExcelProcessIdentity
    shared_deadline: float

    @classmethod
    def from_json(cls, source: str) -> "HelperRequest":
        try:
            value = json.loads(source)
        except (TypeError, json.JSONDecodeError) as exc:
            raise ValueError("Invalid helper request JSON") from exc
        required = {
            "pid",
            "start_seconds",
            "start_microseconds",
            "executable",
            "bundle_id",
            "shared_deadline",
        }
        if not isinstance(value, dict) or set(value) != required:
            raise ValueError("Invalid helper request fields")
        identity = ExcelProcessIdentity(
            pid=value["pid"],
            start_seconds=value["start_seconds"],
            start_microseconds=value["start_microseconds"],
            executable=value["executable"],
            bundle_id=value["bundle_id"],
        )
        deadline = value["shared_deadline"]
        if (
            isinstance(deadline, bool)
            or not isinstance(deadline, (int, float))
            or not math.isfinite(deadline)
            or deadline <= 0
        ):
            raise ValueError("Invalid helper deadline")
        return cls(identity=identity, shared_deadline=float(deadline))


def require_excel_process(identity: ExcelProcessIdentity) -> None:
    if (
        identity.bundle_id != _EXCEL_BUNDLE_ID
        or PurePosixPath(identity.executable).name != "Microsoft Excel"
    ):
        raise OSATransportError("OSA_NOT_EXCEL_PROCESS", pid=identity.pid)


def require_same_process(
    expected: ExcelProcessIdentity, current: Optional[ExcelProcessIdentity]
) -> None:
    if current is None:
        raise OSATransportError("OSA_PROCESS_NOT_FOUND", pid=expected.pid)
    if current.pid != expected.pid or (
        current.start_seconds,
        current.start_microseconds,
    ) != (expected.start_seconds, expected.start_microseconds):
        raise OSATransportError("OSA_PROCESS_IDENTITY_CHANGED", pid=expected.pid)
    if posixpath.realpath(current.executable) != posixpath.realpath(expected.executable):
        raise OSATransportError("OSA_PROCESS_EXECUTABLE_CHANGED", pid=expected.pid)
    if current.bundle_id != expected.bundle_id:
        raise OSATransportError("OSA_PROCESS_BUNDLE_CHANGED", pid=expected.pid)


def finite_timeout_ticks(requested: int, *, deadline: float, now: float) -> int:
    """Return a positive finite Apple Event timeout bounded by a deadline."""
    remaining = deadline - now
    if not math.isfinite(remaining) or remaining <= 0:
        raise OSATransportError("OSA_TIMEOUT", outcome_uncertain=True)
    remaining_ticks = min(
        _MAX_TIMEOUT_TICKS, max(1, math.floor(remaining * _TICKS_PER_SECOND))
    )
    if requested in (_K_AE_DEFAULT_TIMEOUT, _K_NO_TIMEOUT) or requested <= 0:
        return remaining_ticks
    return min(requested, remaining_ticks)


@dataclass(frozen=True)
class AppleEventTarget:
    kind: str
    value: Any = None

    @classmethod
    def current_process(cls) -> "AppleEventTarget":
        return cls("current")

    @classmethod
    def kernel_pid(cls, pid: int) -> "AppleEventTarget":
        return cls("pid", pid)

    @classmethod
    def bundle(cls, bundle_id: str) -> "AppleEventTarget":
        return cls("bundle", bundle_id)

    @classmethod
    def application_path(cls, path: str) -> "AppleEventTarget":
        return cls("path", path)

    @classmethod
    def unknown(cls, descriptor_type: str, data: bytes) -> "AppleEventTarget":
        return cls("unknown", (descriptor_type, data))


class BoundSendRouter:
    """Pure routing policy around a small native Apple Event adapter."""

    def __init__(
        self,
        identity: ExcelProcessIdentity,
        process_lookup: Any,
        *,
        helper_pid: int,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.identity = identity
        self.process_lookup = process_lookup
        self.helper_pid = helper_pid
        self.clock = clock

    def _is_verified_excel_target(self, target: AppleEventTarget) -> bool:
        if target.kind == "bundle":
            return target.value == self.identity.bundle_id == _EXCEL_BUNDLE_ID
        if target.kind == "path":
            target_path = posixpath.realpath(str(target.value))
            executable = posixpath.realpath(self.identity.executable)
            app = str(PurePosixPath(executable).parents[2])
            return target_path in (executable, posixpath.realpath(app))
        if target.kind == "pid":
            if target.value == self.helper_pid:
                return False
            addressed = self.process_lookup.snapshot(target.value)
            return (
                addressed is not None
                and addressed.bundle_id == _EXCEL_BUNDLE_ID
                and posixpath.realpath(addressed.executable)
                == posixpath.realpath(self.identity.executable)
            )
        return False

    def send(
        self,
        adapter: Any,
        event: Any,
        reply: Any,
        mode: int,
        priority: int,
        requested_timeout: int,
        idle: Any,
        filter_: Any,
        deadline: float,
    ) -> int:
        address = None
        event_copy = None
        primary = None
        try:
            address = adapter.get_address(event)
            target = adapter.decode_target(address)
            timeout = finite_timeout_ticks(
                requested_timeout, deadline=deadline, now=self.clock()
            )
            if target.kind == "current" or (
                target.kind == "pid" and target.value == self.helper_pid
            ):
                return adapter.invoke_original(
                    event, reply, mode, priority, timeout, idle, filter_
                )
            if not self._is_verified_excel_target(target):
                raise OSATransportError("OSA_TARGET_REJECTED", pid=self.identity.pid)
            require_same_process(
                self.identity, self.process_lookup.snapshot(self.identity.pid)
            )
            event_copy = adapter.duplicate_event(event)
            adapter.replace_address_with_pid(event_copy, self.identity.pid)
            return adapter.invoke_original(
                event_copy, reply, mode, priority, timeout, idle, filter_
            )
        except BaseException as exc:
            primary = exc
            raise
        finally:
            cleanup_errors = []
            for descriptor in (address, event_copy):
                if descriptor is None:
                    continue
                try:
                    adapter.dispose(descriptor)
                except BaseException as exc:
                    cleanup_errors.append(exc)
            if cleanup_errors:
                if primary is None:
                    raise cleanup_errors[0]
                setattr(primary, "cleanup_errors", tuple(cleanup_errors))
                if hasattr(primary, "add_note"):
                    primary.add_note(
                        "Descriptor dispose failed: "
                        + ", ".join(str(error) for error in cleanup_errors)
                    )


class SendCallbackLease:
    """Install a callback temporarily and restore the exact prior callback."""

    def __init__(
        self,
        setter: Callable[[Any, int], None],
        *,
        original: Any,
        original_refcon: int,
    ) -> None:
        self._setter = setter
        self._original = original
        self._original_refcon = original_refcon

    @contextmanager
    def installed(self, callback: Any):
        self._setter(callback, 0)
        try:
            yield
        except BaseException as primary:
            try:
                self._setter(self._original, self._original_refcon)
            except BaseException as restore_error:
                if hasattr(primary, "add_note"):
                    primary.add_note(
                        f"OSA send callback restoration failed: {restore_error}"
                    )
            raise
        else:
            self._setter(self._original, self._original_refcon)


class SendCallbackBridge:
    """Convert Python routing failures to OSErr while retaining their detail."""

    def __init__(self, router: BoundSendRouter, *, adapter: Any, deadline: float):
        self.router = router
        self.adapter = adapter
        self.deadline = deadline
        self.error: Optional[BaseException] = None
        self._failed_status: Optional[int] = None

    def __call__(
        self,
        event: Any,
        reply: Any,
        mode: int,
        priority: int,
        timeout: int,
        idle: Any,
        filter_: Any,
        refcon: int,
    ) -> int:
        if self._failed_status is not None:
            return self._failed_status
        try:
            status = self.router.send(
                self.adapter,
                event,
                reply,
                mode,
                priority,
                timeout,
                idle,
                filter_,
                self.deadline,
            )
            if status:
                self.error = OSATransportError(
                    "OSA_SEND_FAILED",
                    pid=self.router.identity.pid,
                    outcome_uncertain=True,
                    detail=str(status),
                )
                self._failed_status = status
            return status
        except BaseException as exc:
            if self.error is None:
                self.error = exc
            self._failed_status = _callback_status(exc)
            return self._failed_status


class BoundOSAKitTransport:
    """Spawn the one-shot helper with an immutable process identity request."""

    def __init__(
        self,
        identity: ExcelProcessIdentity,
        *,
        runner: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run,
        clock: Callable[[], float] = time.monotonic,
        shared_clock: Callable[[], float] = _shared_monotonic,
        helper_path: Optional[Path] = None,
    ) -> None:
        require_excel_process(identity)
        self.identity = identity
        self._runner = runner
        self._clock = clock
        self._shared_clock = shared_clock
        self._helper_path = Path(helper_path or __file__).resolve()

    def run(self, script_path: Path, deadline: float) -> subprocess.CompletedProcess[str]:
        script = Path(script_path).expanduser().resolve()
        if not script.is_file():
            raise FileNotFoundError(script)
        remaining = deadline - self._clock()
        if not math.isfinite(remaining) or remaining <= 0:
            raise OSATransportError(
                "OSA_TIMEOUT", pid=self.identity.pid, outcome_uncertain=False
            )
        shared_now = self._shared_clock()
        remaining = deadline - self._clock()
        if not math.isfinite(shared_now) or not math.isfinite(remaining) or remaining <= 0:
            raise OSATransportError(
                "OSA_TIMEOUT", pid=self.identity.pid, outcome_uncertain=False
            )
        shared_deadline = shared_now + remaining
        request = self.identity.to_request(shared_deadline=shared_deadline)
        command = [
            sys.executable,
            str(self._helper_path),
            json.dumps(request, separators=(",", ":")),
            "--helper",
            str(script),
        ]
        try:
            remaining = deadline - self._clock()
            if not math.isfinite(remaining) or remaining <= 0:
                raise OSATransportError(
                    "OSA_TIMEOUT", pid=self.identity.pid, outcome_uncertain=False
                )
            return self._runner(
                command, capture_output=True, text=True, timeout=remaining
            )
        except subprocess.TimeoutExpired as exc:
            raise OSATransportError(
                "OSA_TIMEOUT",
                stdout=exc.stdout or "",
                stderr=exc.stderr or "",
                pid=self.identity.pid,
                outcome_uncertain=True,
            ) from None
        except OSError as exc:
            raise OSATransportError(
                "OSA_HELPER_LAUNCH_FAILED",
                stderr=str(exc),
                pid=self.identity.pid,
                outcome_uncertain=False,
                detail=type(exc).__name__,
            ) from None


class _AEDesc(C.Structure):
    _fields_ = [("descriptor_type", C.c_uint32), ("data_handle", C.c_void_p)]


class _ProcessSerialNumber(C.Structure):
    _fields_ = [("high", C.c_uint32), ("low", C.c_uint32)]


def _four_char(value: str) -> int:
    return int.from_bytes(value.encode("ascii"), "big")


class _DarwinRuntime:
    """Exact ctypes bridge, instantiated only by the macOS helper CLI."""

    def __init__(self) -> None:
        if sys.platform != "darwin":
            raise OSATransportError("OSA_PLATFORM_UNAVAILABLE")
        C.CDLL("/System/Library/Frameworks/Foundation.framework/Foundation")
        C.CDLL("/System/Library/Frameworks/AppKit.framework/AppKit")
        C.CDLL("/System/Library/Frameworks/OSAKit.framework/OSAKit")
        self.objc = C.CDLL("/usr/lib/libobjc.A.dylib")
        self.carbon = C.CDLL("/System/Library/Frameworks/Carbon.framework/Carbon")
        self.libproc = C.CDLL("/usr/lib/libproc.dylib")
        self._configure_objc()
        self._configure_carbon()
        self._configure_libproc()

    def _configure_objc(self) -> None:
        self.objc.objc_getClass.argtypes = [C.c_char_p]
        self.objc.objc_getClass.restype = C.c_void_p
        self.objc.sel_registerName.argtypes = [C.c_char_p]
        self.objc.sel_registerName.restype = C.c_void_p
        self.objc.objc_msgSend.argtypes = [C.c_void_p, C.c_void_p]
        self.objc.objc_msgSend.restype = C.c_void_p

    @staticmethod
    def _bind(library: Any, name: str, argtypes: list[Any], restype: Any) -> Any:
        function = getattr(library, name)
        function.argtypes = argtypes
        function.restype = restype
        return function

    def _configure_carbon(self) -> None:
        desc = C.POINTER(_AEDesc)
        self.ae_get_attribute = self._bind(
            self.carbon,
            "AEGetAttributeDesc",
            [desc, C.c_uint32, C.c_uint32, desc],
            C.c_int16,
        )
        self.ae_get_size = self._bind(
            self.carbon, "AEGetDescDataSize", [desc], C.c_long
        )
        self.ae_get_data = self._bind(
            self.carbon, "AEGetDescData", [desc, C.c_void_p, C.c_long], C.c_int16
        )
        self.ae_duplicate = self._bind(
            self.carbon, "AEDuplicateDesc", [desc, desc], C.c_int16
        )
        self.ae_put_attribute = self._bind(
            self.carbon,
            "AEPutAttributePtr",
            [desc, C.c_uint32, C.c_uint32, C.c_void_p, C.c_long],
            C.c_int16,
        )
        self.ae_dispose = self._bind(
            self.carbon, "AEDisposeDesc", [desc], C.c_int16
        )
        self.get_process_pid = self._bind(
            self.carbon,
            "GetProcessPID",
            [C.POINTER(_ProcessSerialNumber), C.POINTER(C.c_int32)],
            C.c_int32,
        )
        callback_args = [
            desc,
            desc,
            C.c_int32,
            C.c_int16,
            C.c_int32,
            C.c_void_p,
            C.c_void_p,
            C.c_ssize_t,
        ]
        self.callback_type = C.CFUNCTYPE(C.c_int16, *callback_args)
        self.osa_get_send = self._bind(
            self.carbon,
            "OSAGetSendProc",
            [C.c_void_p, C.POINTER(C.c_void_p), C.POINTER(C.c_ssize_t)],
            C.c_int32,
        )
        self.osa_set_send = self._bind(
            self.carbon,
            "OSASetSendProc",
            [C.c_void_p, self.callback_type, C.c_ssize_t],
            C.c_int32,
        )

    def _configure_libproc(self) -> None:
        self.proc_pidinfo = self._bind(
            self.libproc,
            "proc_pidinfo",
            [C.c_int, C.c_int, C.c_uint64, C.c_void_p, C.c_int],
            C.c_int,
        )
        self.proc_pidpath = self._bind(
            self.libproc,
            "proc_pidpath",
            [C.c_int, C.c_void_p, C.c_uint32],
            C.c_int,
        )

    def objc_class(self, name: str) -> int:
        value = self.objc.objc_getClass(name.encode())
        if not value:
            raise OSATransportError("OSA_COCOA_UNAVAILABLE", detail=name)
        return value

    def selector(self, name: str) -> int:
        value = self.objc.sel_registerName(name.encode())
        if not value:
            raise OSATransportError("OSA_COCOA_UNAVAILABLE", detail=name)
        return value

    def message(
        self,
        receiver: Any,
        name: str,
        args: tuple[Any, ...] = (),
        types: tuple[Any, ...] = (),
        result: Any = C.c_void_p,
    ) -> Any:
        if not receiver:
            raise OSATransportError("OSA_COCOA_NIL_RECEIVER", detail=name)
        address = C.cast(self.objc.objc_msgSend, C.c_void_p).value
        call = C.CFUNCTYPE(result, C.c_void_p, C.c_void_p, *types)(address)
        return call(receiver, self.selector(name), *args)

    def ns_string(self, value: str) -> Any:
        result = self.message(
            self.objc_class("NSString"),
            "stringWithUTF8String:",
            (value.encode(),),
            (C.c_char_p,),
        )
        if not result:
            raise OSATransportError("OSA_STRING_CREATION_FAILED")
        return result

    def object_text(self, value: Any) -> str:
        pointer = self.message(value, "UTF8String", result=C.c_char_p)
        if pointer is None:
            raise OSATransportError("OSA_RESULT_DECODING_FAILED")
        return pointer.decode("utf-8")

    def check(self, status: int, code: str) -> None:
        if status:
            raise OSATransportError(code, detail=str(status))


class _ProcBSDInfo(C.Structure):
    _fields_ = [
        ("flags", C.c_uint32),
        ("status", C.c_uint32),
        ("xstatus", C.c_uint32),
        ("pid", C.c_uint32),
        ("ppid", C.c_uint32),
        ("uid", C.c_uint32),
        ("gid", C.c_uint32),
        ("ruid", C.c_uint32),
        ("rgid", C.c_uint32),
        ("svuid", C.c_uint32),
        ("svgid", C.c_uint32),
        ("reserved", C.c_uint32),
        ("command", C.c_char * 16),
        ("name", C.c_char * 32),
        ("nfiles", C.c_uint32),
        ("pgid", C.c_uint32),
        ("pjobc", C.c_uint32),
        ("tdev", C.c_uint32),
        ("tpgid", C.c_uint32),
        ("nice", C.c_int32),
        ("start_seconds", C.c_uint64),
        ("start_microseconds", C.c_uint64),
    ]


class _DarwinProcessLookup:
    def __init__(self, runtime: _DarwinRuntime) -> None:
        self.runtime = runtime

    def snapshot(
        self, pid: int, *, application: Any = None
    ) -> Optional[ExcelProcessIdentity]:
        info = _ProcBSDInfo()
        size = self.runtime.proc_pidinfo(
            pid, 3, 0, C.byref(info), C.sizeof(info)
        )
        if size != C.sizeof(info):
            return None
        if info.pid != pid:
            raise OSATransportError("OSA_PROCESS_IDENTITY_CHANGED", pid=pid)
        buffer = C.create_string_buffer(4096)
        length = self.runtime.proc_pidpath(pid, buffer, len(buffer))
        if length <= 0:
            return None
        executable = posixpath.realpath(buffer.value.decode("utf-8"))
        app = application
        if app is None:
            app = self.runtime.message(
                self.runtime.objc_class("NSRunningApplication"),
                "runningApplicationWithProcessIdentifier:",
                (pid,),
                (C.c_int32,),
            )
        elif self.runtime.message(
            app, "processIdentifier", result=C.c_int32
        ) != pid:
            raise OSATransportError("OSA_PROCESS_IDENTITY_CHANGED", pid=pid)
        if not app:
            return None
        bundle_obj = self.runtime.message(app, "bundleIdentifier")
        executable_url = self.runtime.message(app, "executableURL")
        if not bundle_obj or not executable_url:
            return None
        cocoa_path = self.runtime.object_text(
            self.runtime.message(executable_url, "path")
        )
        if posixpath.realpath(cocoa_path) != executable:
            raise OSATransportError("OSA_PROCESS_EXECUTABLE_CHANGED", pid=pid)
        final_info = _ProcBSDInfo()
        final_size = self.runtime.proc_pidinfo(
            pid, 3, 0, C.byref(final_info), C.sizeof(final_info)
        )
        if final_size != C.sizeof(final_info):
            return None
        if (
            final_info.pid != info.pid
            or final_info.start_seconds != info.start_seconds
            or final_info.start_microseconds != info.start_microseconds
        ):
            raise OSATransportError("OSA_PROCESS_IDENTITY_CHANGED", pid=pid)
        return ExcelProcessIdentity(
            pid=pid,
            start_seconds=info.start_seconds,
            start_microseconds=info.start_microseconds,
            executable=executable,
            bundle_id=self.runtime.object_text(bundle_obj),
        )


class _NativeEventAdapter:
    def __init__(self, runtime: _DarwinRuntime, original: Any, refcon: int) -> None:
        self.runtime = runtime
        self.original = original
        self.refcon = refcon

    def get_address(self, event: Any) -> _AEDesc:
        address = _AEDesc()
        self.runtime.check(
            self.runtime.ae_get_attribute(
                event, _four_char("addr"), _four_char("****"), C.byref(address)
            ),
            "OSA_ADDRESS_READ_FAILED",
        )
        return address

    def _data(self, descriptor: _AEDesc) -> bytes:
        size = self.runtime.ae_get_size(C.byref(descriptor))
        if size < 0:
            raise OSATransportError("OSA_ADDRESS_READ_FAILED")
        buffer = C.create_string_buffer(size or 1)
        self.runtime.check(
            self.runtime.ae_get_data(C.byref(descriptor), buffer, size),
            "OSA_ADDRESS_READ_FAILED",
        )
        return buffer.raw[:size]

    def decode_target(self, descriptor: _AEDesc) -> AppleEventTarget:
        kind = descriptor.descriptor_type
        data = self._data(descriptor)
        if kind == _four_char("psn ") and len(data) == C.sizeof(_ProcessSerialNumber):
            psn = _ProcessSerialNumber.from_buffer_copy(data)
            if psn.high == 0 and psn.low == 2:
                return AppleEventTarget.current_process()
            pid = C.c_int32()
            self.runtime.check(
                self.runtime.get_process_pid(C.byref(psn), C.byref(pid)),
                "OSA_TARGET_LOOKUP_FAILED",
            )
            return AppleEventTarget.kernel_pid(pid.value)
        if kind == _four_char("kpid") and len(data) == C.sizeof(C.c_int32):
            return AppleEventTarget.kernel_pid(C.c_int32.from_buffer_copy(data).value)
        if kind == _four_char("bund"):
            return AppleEventTarget.bundle(data.decode("utf-8"))
        if kind in (_four_char("aprl"), _four_char("furl")):
            value = data.decode("utf-8")
            if value.startswith("file://"):
                from urllib.parse import unquote, urlparse

                value = unquote(urlparse(value).path)
            return AppleEventTarget.application_path(value)
        return AppleEventTarget.unknown(
            kind.to_bytes(4, "big").decode("latin-1"), data
        )

    def duplicate_event(self, event: Any) -> _AEDesc:
        copy = _AEDesc()
        self.runtime.check(
            self.runtime.ae_duplicate(event, C.byref(copy)), "OSA_EVENT_COPY_FAILED"
        )
        return copy

    def replace_address_with_pid(self, event: _AEDesc, pid: int) -> None:
        native_pid = C.c_int32(pid)
        self.runtime.check(
            self.runtime.ae_put_attribute(
                C.byref(event),
                _four_char("addr"),
                _four_char("kpid"),
                C.byref(native_pid),
                C.sizeof(native_pid),
            ),
            "OSA_ADDRESS_WRITE_FAILED",
        )

    def invoke_original(
        self,
        event: Any,
        reply: Any,
        mode: int,
        priority: int,
        timeout: int,
        idle: Any,
        filter_: Any,
    ) -> int:
        return self.original(
            event, reply, mode, priority, timeout, idle, filter_, self.refcon
        )

    def dispose(self, descriptor: _AEDesc) -> None:
        self.runtime.check(
            self.runtime.ae_dispose(C.byref(descriptor)), "OSA_DESCRIPTOR_DISPOSE_FAILED"
        )


def _callback_status(error: BaseException) -> int:
    if isinstance(error, OSATransportError):
        if error.code == "OSA_TIMEOUT":
            return _ERR_AE_TIMEOUT
        if error.code in (
            "OSA_PROCESS_NOT_FOUND",
            "OSA_PROCESS_IDENTITY_CHANGED",
            "OSA_PROCESS_EXECUTABLE_CHANGED",
            "OSA_PROCESS_BUNDLE_CHANGED",
        ):
            return _PROC_NOT_FOUND
        if error.code == "OSA_TARGET_REJECTED":
            return _ERR_AE_EVENT_NOT_PERMITTED
    return _PARAM_ERR


def _error_description(runtime: _DarwinRuntime, error: Any) -> str:
    if not error:
        return "Unknown OSAKit error"
    try:
        return runtime.object_text(runtime.message(error, "description"))
    except BaseException:
        return "Unable to decode OSAKit error"


def _helper_router(
    identity: ExcelProcessIdentity, process_lookup: Any, *, helper_pid: int
) -> BoundSendRouter:
    return BoundSendRouter(
        identity, process_lookup, helper_pid=helper_pid, clock=_shared_monotonic
    )


def _execute_helper(request: HelperRequest, script_path: Path) -> str:
    runtime = _DarwinRuntime()
    deadline = request.shared_deadline
    require_excel_process(request.identity)
    source = script_path.read_text(encoding="utf-8")
    pool = instance = script = None
    component = None
    prior_pointer = C.c_void_p()
    prior_refcon = C.c_ssize_t()
    callback = None
    try:
        pool = runtime.message(runtime.objc_class("NSAutoreleasePool"), "new")
        process_lookup = _DarwinProcessLookup(runtime)
        require_same_process(
            request.identity, process_lookup.snapshot(request.identity.pid)
        )
        language = runtime.message(
            runtime.objc_class("OSALanguage"),
            "languageForName:",
            (runtime.ns_string("AppleScript"),),
            (C.c_void_p,),
        )
        if not language:
            raise OSATransportError("OSA_LANGUAGE_UNAVAILABLE")
        instance = runtime.message(runtime.objc_class("OSALanguageInstance"), "alloc")
        instance = runtime.message(
            instance,
            "initWithLanguage:",
            (language,),
            (C.c_void_p,),
        )
        script = runtime.message(runtime.objc_class("OSAScript"), "alloc")
        script = runtime.message(
            script,
            "initWithSource:fromURL:languageInstance:usingStorageOptions:",
            (runtime.ns_string(source), None, instance, 0),
            (C.c_void_p, C.c_void_p, C.c_void_p, C.c_ulong),
        )
        error = C.c_void_p()
        compiled = runtime.message(
            script,
            "compileAndReturnError:",
            (C.byref(error),),
            (C.POINTER(C.c_void_p),),
            C.c_bool,
        )
        if not compiled:
            raise OSATransportError(
                "OSA_COMPILE_FAILED", detail=_error_description(runtime, error.value)
            )
        component = runtime.message(instance, "componentInstance")
        if not component:
            raise OSATransportError("OSA_COMPONENT_UNAVAILABLE")
        runtime.check(
            runtime.osa_get_send(
                component, C.byref(prior_pointer), C.byref(prior_refcon)
            ),
            "OSA_SEND_CALLBACK_READ_FAILED",
        )
        if not prior_pointer.value:
            raise OSATransportError("OSA_SEND_CALLBACK_READ_FAILED")
        original = runtime.callback_type(prior_pointer.value)
        adapter = _NativeEventAdapter(runtime, original, prior_refcon.value)
        router = _helper_router(
            request.identity, process_lookup, helper_pid=os.getpid()
        )

        bridge = SendCallbackBridge(router, adapter=adapter, deadline=deadline)
        callback = runtime.callback_type(bridge)
        def set_send(send_callback: Any, refcon: int) -> None:
            runtime.check(
                runtime.osa_set_send(component, send_callback, refcon),
                "OSA_SEND_CALLBACK_SET_FAILED",
            )
        lease = SendCallbackLease(
            set_send, original=original, original_refcon=prior_refcon.value
        )
        with lease.installed(callback):
            result = runtime.message(
                script,
                "executeAndReturnError:",
                (C.byref(error),),
                (C.POINTER(C.c_void_p),),
            )
            if bridge.error is not None:
                raise bridge.error
            if not result:
                raise OSATransportError(
                    "OSA_EXECUTION_FAILED",
                    pid=request.identity.pid,
                    outcome_uncertain=True,
                    detail=_error_description(runtime, error.value),
                )
            string_value = runtime.message(result, "stringValue")
            if not string_value:
                raise OSATransportError("OSA_RESULT_DECODING_FAILED")
            return runtime.object_text(string_value)
    finally:
        for value in (script, instance):
            if value:
                runtime.message(value, "release")
        if pool:
            runtime.message(pool, "drain")
        _ = callback  # Keep the CFUNCTYPE closure alive through restoration.


def _helper_main(request_json: str, script_source: str) -> int:
    try:
        request = HelperRequest.from_json(request_json)
        script = Path(script_source).expanduser().resolve()
        if not script.is_file():
            raise FileNotFoundError(script)
        result = _execute_helper(request, script)
    except BaseException as exc:
        error = exc if isinstance(exc, OSATransportError) else OSATransportError(
            "OSA_HELPER_FAILED", detail=f"{type(exc).__name__}: {exc}"
        )
        print(json.dumps(error.as_dict(), separators=(",", ":")), file=sys.stderr)
        return 1
    print(result)
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if len(arguments) != 3 or arguments[1] != "--helper":
        print("usage: macos_osa_transport.py REQUEST_JSON --helper SCRIPT", file=sys.stderr)
        return 2
    return _helper_main(arguments[0], arguments[2])


if __name__ == "__main__":
    raise SystemExit(main())
