"""Bounded IPC for one AppKit launcher; Cocoa never executes in the client.

The helper owns the returned NSRunningApplication for its entire lifetime.
Stopping this owned helper never closes, quits, or signals Microsoft Excel.
"""
from __future__ import annotations

from dataclasses import asdict
import json
import math
import os
from pathlib import Path
import select
import socket
import struct
import subprocess
import sys
import threading
import time
from typing import Any, Callable, Optional
from uuid import uuid4

# Direct script invocation must use the adjacent installed package, not cwd.
if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
    from skills.WPSComposer.scripts.msoffice.macos_excel_process import LaunchedExcel
    from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity, require_excel_process, require_same_process
else:
    from .macos_excel_process import LaunchedExcel
    from .macos_osa_transport import ExcelProcessIdentity, require_excel_process, require_same_process

_MAX_FRAME = 131072


class LauncherError(RuntimeError):
    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


def shared_clock() -> float:
    return time.clock_gettime(time.CLOCK_MONOTONIC)


def _remaining(deadline: float, clock: Callable[[], float]) -> float:
    remaining = deadline - clock()
    if not math.isfinite(remaining) or remaining <= 0:
        raise LauncherError('EXCEL_LAUNCHER_TIMEOUT')
    return remaining


def _transfer(sock: socket.socket, value: Any, deadline: float,
              clock: Callable[[], float], *, receive: bool = False) -> Any:
    def io(data: Any, *, read: bool = False) -> bytes:
        out = bytearray()
        while (len(out) < data) if read else bool(data):
            readable, writable, _ = select.select([sock] if read else [], [] if read else [sock], [], _remaining(deadline, clock))
            if not readable and not writable:
                raise LauncherError('EXCEL_LAUNCHER_TIMEOUT')
            try:
                if read:
                    part = sock.recv(data-len(out))
                    if not part:
                        raise EOFError('Launcher peer closed')
                    out.extend(part)
                else:
                    count = sock.send(data)
                    if count == 0:
                        raise EOFError('Launcher peer closed')
                    data = data[count:]
            except BlockingIOError:
                continue
        return bytes(out)
    if receive:
        size = struct.unpack('!I', io(4, read=True))[0]
        if not 0 < size <= _MAX_FRAME:
            raise LauncherError('EXCEL_LAUNCHER_BAD_FRAME')
        try:
            result = json.loads(io(size, read=True))
        except (ValueError, UnicodeError):
            raise LauncherError('EXCEL_LAUNCHER_BAD_FRAME') from None
        if not isinstance(result, dict):
            raise LauncherError('EXCEL_LAUNCHER_BAD_FRAME')
        return result
    raw = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(',', ':')).encode()
    if not 0 < len(raw) <= _MAX_FRAME:
        raise LauncherError('EXCEL_LAUNCHER_BAD_FRAME')
    io(struct.pack('!I', len(raw)) + raw)


def _durable(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_name(path.name + '.' + uuid4().hex + '.tmp')
    try:
        with temporary.open('x', encoding='utf-8') as stream:
            json.dump(value, stream, ensure_ascii=False, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
        fd = os.open(str(path.parent), os.O_RDONLY)
        try:
            os.fsync(fd)
        finally:
            os.close(fd)
    finally:
        temporary.unlink(missing_ok=True)


class BoundedExcelLauncher:
    """One helper, one launch attempt; every IPC call uses the current deadline."""
    def __init__(self, *, evidence_dir: Path, deadline: float,
                 clock: Callable[[], float] = time.monotonic,
                 os_clock: Callable[[], float] = shared_clock,
                 helper_command: Optional[list[str]] = None) -> None:
        self._clock, self._os_clock = clock, os_clock
        _remaining(deadline, clock)
        self._session_deadline = self._deadline = deadline
        self.evidence_dir = Path(evidence_dir).resolve()
        self.evidence_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.nonce = uuid4().hex
        self.failed = False
        self.process: Any = None
        self.socket: Any = None
        self.recovery_candidate: Optional[dict[str, Any]] = None
        self.cleanup_errors: list[str] = []
        self._record: Optional[LaunchedExcel] = None
        self._launch_requested = False
        self._request = 0
        self._command = helper_command or [sys.executable, str(Path(__file__).resolve())]
        _durable(self.evidence_dir/'intent.json', {'nonce':self.nonce, 'state':'prepared', 'helper_pid':None})

    def set_deadline(self, deadline: float) -> None:
        _remaining(deadline, self._clock)
        self._deadline = min(self._session_deadline, deadline)

    def _spawn(self) -> None:
        parent, child = socket.socketpair()
        liveness_read, liveness_write = socket.socketpair()
        self._liveness = liveness_write
        parent.setblocking(False)
        self.socket = parent
        try:
            with (self.evidence_dir/'helper.stdout').open('xb') as stdout, (self.evidence_dir/'helper.stderr').open('xb') as stderr:
                self.process = subprocess.Popen(self._command + ['--helper-fd', str(child.fileno()), self.nonce, str(self.evidence_dir), str(liveness_read.fileno())],
                    stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
                    pass_fds=(child.fileno(), liveness_read.fileno()), close_fds=True)
            _durable(self.evidence_dir/'helper.json', {'nonce':self.nonce, 'helper_pid':self.process.pid})
        finally:
            child.close()
            liveness_read.close()

    def abort(self) -> None:
        """Stop only the child held by Popen; never signal an Office PID/group."""
        if getattr(self, "_liveness", None) is not None:
            self._liveness.close()
        if self.socket is not None:
            self.socket.close()
        if self.process is not None and self.process.poll() is None:
            self.process.kill()
            try:
                self.process.wait(timeout=.25)
            except subprocess.TimeoutExpired:
                if self.process.poll() is not None:
                    return
                error = LauncherError('EXCEL_LAUNCHER_HELPER_STILL_RUNNING')
                self.cleanup_errors.append(error.code)
                raise error
            if self.process.poll() is None:
                error = LauncherError('EXCEL_LAUNCHER_HELPER_STILL_RUNNING')
                self.cleanup_errors.append(error.code)
                raise error

    def _failure(self) -> None:
        self.failed = True
        try:
            self.abort()
        except BaseException as error:
            self.cleanup_errors.append(str(error))
        for name in ('identity.json', 'candidate.json'):
            try:
                path = self.evidence_dir/name
                if path.is_file():
                    value = json.loads(path.read_text())
                    if value.get('nonce') == self.nonce:
                        self.recovery_candidate = value['candidate']
                        break
            except BaseException as error:
                self.cleanup_errors.append(str(error))

    def _rpc(self, operation: str, data: dict[str, Any]) -> dict[str, Any]:
        if self.failed:
            raise LauncherError('EXCEL_LAUNCHER_UNAVAILABLE')
        try:
            _remaining(self._deadline, self._clock)
            if self.process is None:
                self._spawn()
            self._request += 1
            request_id = self._request
            remaining = _remaining(self._deadline, self._clock)
            shared_deadline = self._os_clock() + remaining
            request = {'version':1, 'nonce':self.nonce, 'id':request_id, 'operation':operation,
                       'deadline':shared_deadline, 'data':data}
            _durable(self.evidence_dir/('request-'+str(request_id)+'.json'), request)
            _transfer(self.socket, request, self._deadline, self._clock)
            while True:
                reply = _transfer(self.socket, None, self._deadline, self._clock, receive=True)
                if set(reply) != {'version','nonce','id','helper_pid','status','data'} or type(reply['version']) is not int or reply['version'] != 1 or reply['nonce'] != self.nonce or type(reply['id']) is not int or reply['id'] != request_id or reply['helper_pid'] != self.process.pid:
                    raise LauncherError('EXCEL_LAUNCHER_BAD_REPLY')
                if not isinstance(reply['data'], dict):
                    raise LauncherError('EXCEL_LAUNCHER_BAD_REPLY')
                _durable(self.evidence_dir/('reply-'+str(request_id)+'.json'), reply)
                if reply['status'] == 'progress':
                    self.recovery_candidate = reply['data']
                    continue
                if reply['status'] == 'error':
                    raise LauncherError(reply['data'].get('code','EXCEL_LAUNCHER_FAILED'))
                if reply['status'] != 'ok':
                    raise LauncherError('EXCEL_LAUNCHER_BAD_REPLY')
                return reply['data']
        except BaseException:
            self._failure()
            raise

    def existing_pids(self) -> set[int]:
        value = self._rpc('HELLO', {})
        pids = value.get('pids')
        if set(value) != {'pids'} or not isinstance(pids,list) or any(type(pid) is not int or pid<=0 for pid in pids):
            self._failure()
            raise LauncherError('EXCEL_LAUNCHER_BAD_REPLY')
        return set(pids)

    def launch(self, app_path: Path, deadline: float) -> LaunchedExcel:
        if self.failed:
            raise LauncherError('EXCEL_LAUNCHER_UNAVAILABLE')
        if self._launch_requested:
            raise LauncherError('EXCEL_LAUNCH_ALREADY_REQUESTED')
        self.set_deadline(deadline)
        self._launch_requested = True
        value = self._rpc('LAUNCH', {'app_path':str(app_path)})
        try:
            if set(value) != {'identity','launch_date','token'} or not isinstance(value['token'],str) or not value['token']:
                raise ValueError('Invalid launch reply')
            identity = ExcelProcessIdentity(**value['identity'])
            require_excel_process(identity)
            if type(value['launch_date']) not in (float,int) or not math.isfinite(value['launch_date']):
                raise ValueError('Invalid launch date')
            self._record = LaunchedExcel(identity, value['launch_date'], value['token'])
            self.recovery_candidate = {'identity':asdict(identity), 'launch_date':value['launch_date'], 'pid':identity.pid}
            return self._record
        except BaseException:
            self._failure()
            raise

    def _probe(self) -> dict[str, Any]:
        if self._record is None:
            raise LauncherError('EXCEL_LAUNCHER_UNAVAILABLE')
        result = self._rpc('PROBE', {'token':self._record.handle,'identity':asdict(self._record.identity)})
        if set(result) != {'identity','finished','terminated'} or type(result['finished']) is not bool or type(result['terminated']) is not bool:
            self._failure()
            raise LauncherError('EXCEL_LAUNCHER_BAD_REPLY')
        if result['identity'] is not None:
            try:
                result['identity'] = ExcelProcessIdentity(**result['identity'])
            except BaseException:
                self._failure()
                raise
        return result

    def snapshot(self, pid: int) -> Optional[ExcelProcessIdentity]:
        if self._record is None or pid != self._record.identity.pid:
            raise LauncherError('EXCEL_LAUNCHER_WRONG_PROCESS')
        return self._probe()['identity']

    def finished_launching(self, record: LaunchedExcel) -> bool:
        if record != self._record:
            raise LauncherError('EXCEL_LAUNCHER_WRONG_PROCESS')
        return self._probe()['finished']

    def has_terminated(self, record: LaunchedExcel) -> bool:
        if record != self._record:
            raise LauncherError('EXCEL_LAUNCHER_WRONG_PROCESS')
        return self._probe()['terminated']

    def wait_until(self, deadline: float) -> None:
        # Local adaptive backoff only; Cocoa runloop pumping is inside PROBE.
        timeout = min(_remaining(self._deadline,self._clock), max(0,deadline-self._clock()))
        select.select([], [], [], timeout)

    def release(self) -> None:
        if self._record is None:
            raise LauncherError('EXCEL_LAUNCHER_UNAVAILABLE')
        self._rpc('RELEASE', {'token':self._record.handle,'identity':asdict(self._record.identity)})
        try:
            if self.process is None:
                raise LauncherError('EXCEL_LAUNCHER_UNAVAILABLE')
            if self.process.poll() is None:
                try:
                    self.process.wait(timeout=_remaining(self._deadline, self._clock))
                except subprocess.TimeoutExpired:
                    if self.process.poll() is None:
                        raise LauncherError('EXCEL_LAUNCHER_HELPER_EXIT_TIMEOUT') from None
            if self.process.returncode != 0:
                raise LauncherError('EXCEL_LAUNCHER_HELPER_EXIT_FAILED')
            self.abort()
        except BaseException:
            self.failed = True
            try:
                self.abort()
            except BaseException as cleanup_error:
                self.cleanup_errors.append(str(cleanup_error))
            raise


def helper_main(argv: list[str], *, launcher_factory: Any = None) -> int:
    if len(argv) != 5 or argv[0] != '--helper-fd':
        return 2
    descriptor, nonce, folder = int(argv[1]), argv[2], Path(argv[3])
    liveness = socket.socket(fileno=int(argv[4]))
    liveness.set_inheritable(False)
    def parent_lifetime() -> None:
        try:
            while liveness.recv(1):
                pass
        finally:
            # Dedicated IPC monitor only: never call AppKit or Office off-main.
            os._exit(0)
    threading.Thread(target=parent_lifetime, daemon=True).start()
    sock = socket.socket(fileno=descriptor)
    sock.set_inheritable(False)
    sock.setblocking(False)
    native = record = None
    token = None
    launched = False
    previous: set[int] = set()
    last_id = 0
    # Parent owns the outer watchdog. Idle helpers exit on peer EOF immediately.
    lifetime = shared_clock()+86400
    try:
        while True:
            request = _transfer(sock,None,lifetime,shared_clock,receive=True)
            if set(request) != {'version','nonce','id','operation','deadline','data'} or type(request['version']) is not int or request['version'] != 1 or request['nonce'] != nonce or type(request['id']) is not int or request['id'] <= last_id or not isinstance(request['data'],dict):
                raise LauncherError('EXCEL_LAUNCHER_BAD_REQUEST')
            last_id = request['id']
            deadline = request['deadline']
            if type(deadline) not in (int,float):
                raise LauncherError('EXCEL_LAUNCHER_BAD_REQUEST')
            _remaining(deadline,shared_clock)
            if native is None:
                lifetime = deadline
            operation, data = request['operation'], request['data']
            def send(status: str, payload: dict[str,Any]) -> None:
                _transfer(sock,{'version':1,'nonce':nonce,'id':last_id,'helper_pid':os.getpid(),'status':status,'data':payload},deadline,shared_clock)
            def observe(stage: str, candidate: dict[str,Any]) -> None:
                _durable(folder/('identity.json' if stage=='identity' else 'candidate.json'), {'nonce':nonce,'candidate':candidate})
                send('progress',candidate)
            try:
                if operation == 'HELLO' and native is None and not data:
                    if launcher_factory is None:
                        from skills.WPSComposer.scripts.msoffice.macos_excel_process import AppKitExcelLauncher
                        factory = AppKitExcelLauncher
                    else:
                        factory = launcher_factory
                    native = factory(clock=shared_clock, observer=observe)
                    _remaining(deadline,shared_clock)
                    previous = native.existing_pids()
                    response = {'pids':sorted(previous)}
                elif operation == 'LAUNCH' and native is not None and not launched and set(data)=={'app_path'}:
                    launched = True
                    path = Path(data['app_path'])
                    if not path.is_absolute():
                        raise LauncherError('EXCEL_LAUNCHER_BAD_REQUEST')
                    _durable(folder/'launching.json', {'nonce':nonce,'operation':'LAUNCH','id':last_id,'app_path':str(path),'previous_pids':sorted(previous)})
                    _remaining(deadline,shared_clock)
                    record = native.launch(path,deadline)
                    if record.identity.pid in previous:
                        raise LauncherError('EXCEL_LAUNCH_REUSED_PROCESS')
                    require_excel_process(record.identity)
                    token = uuid4().hex
                    candidate = {'pid':record.identity.pid,'identity':asdict(record.identity),'launch_date':record.launch_date}
                    observe('identity',candidate)
                    response = {'identity':asdict(record.identity),'launch_date':record.launch_date,'token':token}
                elif operation in ('PROBE','RELEASE') and record is not None and set(data)=={'token','identity'} and data['token']==token and data['identity']==asdict(record.identity):
                    _remaining(deadline,shared_clock)
                    native.wait_until(shared_clock())
                    _remaining(deadline,shared_clock)
                    current = native.snapshot(record.identity.pid)
                    _remaining(deadline,shared_clock)
                    finished = bool(native.finished_launching(record))
                    _remaining(deadline,shared_clock)
                    terminated = bool(native.has_terminated(record))
                    _remaining(deadline,shared_clock)
                    same_birth = current is not None and (current.start_seconds,current.start_microseconds)==(record.identity.start_seconds,record.identity.start_microseconds)
                    if current is not None and not terminated:
                        require_same_process(record.identity, current)
                    if operation=='RELEASE':
                        if same_birth or not terminated:
                            raise LauncherError('EXCEL_LAUNCHER_NOT_RELEASED')
                        send('ok',{})
                        return 0
                    response = {'identity':asdict(current) if current else None,'finished':finished,'terminated':terminated}
                else:
                    raise LauncherError('EXCEL_LAUNCHER_BAD_REQUEST')
                _remaining(deadline,shared_clock)
                send('ok',response)
            except BaseException as error:
                try:
                    send('error',{'code':getattr(error,'code',type(error).__name__)})
                finally:
                    return 1
    except EOFError:
        return 0
    except BaseException:
        return 1
    finally:
        sock.close()
        # No Excel close/quit/terminate operation belongs to helper teardown.


if __name__ == '__main__':
    raise SystemExit(helper_main(sys.argv[1:]))
