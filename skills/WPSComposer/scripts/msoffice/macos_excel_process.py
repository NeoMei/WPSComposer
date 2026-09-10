"""Ownership of one explicitly launched Excel instance; no attached-app fallback.

The caller holds its component lock until close succeeds or recovery is persisted.
Importing this module does not load Cocoa or launch any native application.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
import json
import math
import os
from pathlib import Path
import tempfile
import time
from typing import Any, Callable, Optional
from uuid import uuid4

from .macos_osa_transport import (
    BoundOSAKitTransport, ExcelProcessIdentity, OSATransportError,
    require_excel_process, require_same_process,
)

_APP = Path('/Applications/Microsoft Excel.app')


class ExcelProcessError(RuntimeError):
    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


@dataclass(frozen=True)
class LaunchedExcel:
    identity: ExcelProcessIdentity
    launch_date: Optional[float]
    handle: Any


class PrivateExcelProcessOwner:
    """Launch, validate, and retire a single process within an absolute deadline.

    ``launcher`` supplies existing_pids, launch, snapshot, finished_launching,
    has_terminated and wait_until. ``transport_factory`` accepts the immutable
    process identity and returns a BoundOSAKitTransport-compatible object.
    ``run`` is the session transport boundary; nonzero or uncertain results
    quarantine this owner. Recovery is deliberately in memory for the caller to
    persist with its staging evidence BEFORE releasing the component lock.
    """

    def __init__(self, *, launcher: Any = None,
                 transport_factory: Callable[..., Any] = BoundOSAKitTransport,
                 clock: Callable[[], float] = time.monotonic,
                 app_path: Path = _APP,
                 evidence_dir: Optional[Path] = None) -> None:
        self._launcher = launcher
        self._transport_factory = transport_factory
        self._clock = clock
        self._app_path = Path(app_path).resolve()
        self._launch: Optional[LaunchedExcel] = None
        self._transport: Any = None
        self.identity: Optional[ExcelProcessIdentity] = None
        self.launch_date: Optional[float] = None
        self.owned_path: Optional[Path] = None
        self.state = 'new'
        self.recovery: Optional[dict[str, Any]] = None
        self._inventory: list[dict[str, Any]] = []
        self.evidence_dir = Path(evidence_dir).absolute() if evidence_dir is not None else Path(tempfile.mkdtemp(prefix='wpscomposer-excel-owner-'))
        self.evidence_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
        self.last_command: Optional[dict[str, Any]] = None

    @property
    def is_private_owned(self) -> bool:
        return self.state in ('ready', 'owned')

    def _deadline(self, deadline: float) -> None:
        if isinstance(deadline, bool) or not isinstance(deadline, (float, int)) or not math.isfinite(deadline):
            raise ValueError('A finite absolute deadline is required')
        if self._clock() >= deadline:
            raise ExcelProcessError('EXCEL_DEADLINE')
        setter = getattr(self._launcher, 'set_deadline', None)
        if setter is not None:
            setter(deadline)

    def _available(self) -> None:
        if not self.is_private_owned:
            raise ExcelProcessError('EXCEL_OWNER_UNAVAILABLE')

    def quarantine(self, error: BaseException) -> None:
        """Retain identity and path; never send cleanup after uncertain failure."""
        if self.state == 'quarantined':
            return
        cleanup_error = None
        abort = getattr(self._launcher, 'abort', None)
        if abort is not None:
            try:
                abort()
            except BaseException as failed_abort:
                cleanup_error = str(failed_abort)
        self.recovery = {
            'helper_cleanup_error': cleanup_error,
            'state': self.state,
            'code': getattr(error, 'code', type(error).__name__),
            'identity': asdict(self.identity) if self.identity else None,
            'launch_date': self.launch_date,
            'launch_candidate': getattr(self._launcher, 'recovery_candidate', None),
            'owned_path': str(self.owned_path) if self.owned_path else None,
            'inventory': self._inventory,
            'outcome_uncertain': True,
            'evidence_dir': str(self.evidence_dir),
            'last_command': self.last_command,
        }
        self.state = 'quarantined'

    def _same_process(self) -> None:
        require_same_process(self.identity, self._launcher.snapshot(self.identity.pid))

    def start(self, *, deadline: float) -> 'PrivateExcelProcessOwner':
        if self.state != 'new':
            raise ExcelProcessError('EXCEL_OWNER_UNAVAILABLE')
        self._deadline(deadline)
        self.state = 'starting'
        try:
            if self._launcher is None:
                from .macos_excel_launcher import BoundedExcelLauncher
                self._launcher = BoundedExcelLauncher(clock=self._clock, deadline=deadline,
                    evidence_dir=self.evidence_dir / 'launcher')
            previous = self._launcher.existing_pids()
            self._deadline(deadline)
            self._launch = self._launcher.launch(self._app_path, deadline)
            self.identity = self._launch.identity
            self.launch_date = self._launch.launch_date
            if self.identity.pid in previous:
                raise ExcelProcessError('EXCEL_LAUNCH_REUSED_PROCESS')
            require_excel_process(self.identity)
            if (type(self.launch_date) not in (float, int) or
                    not math.isfinite(self.launch_date)):
                raise ExcelProcessError('EXCEL_LAUNCH_DATE_MISMATCH')
            expected_executable = self._app_path / 'Contents/MacOS/Microsoft Excel'
            if os.path.realpath(self.identity.executable) != os.path.realpath(expected_executable):
                raise ExcelProcessError('EXCEL_LAUNCH_EXECUTABLE_MISMATCH')
            self._same_process()
            interval = 0.01
            while True:
                self._deadline(deadline)
                self._same_process()
                if self._launcher.has_terminated(self._launch):
                    raise ExcelProcessError('EXCEL_LAUNCH_EXITED')
                if self._launcher.finished_launching(self._launch):
                    break
                self._launcher.wait_until(min(deadline, self._clock() + interval))
                interval = min(interval * 2, 0.25)
            self._transport = self._transport_factory(self.identity)
            books = self._read_inventory(deadline)
            if books:
                if len(books) != 1 or not _pristine_record(books[0]):
                    raise ExcelProcessError('EXCEL_FOREIGN_WORKBOOKS')
                self._command('bootstrap', deadline, book=books[0], acknowledgment='empty')
                if self._read_inventory(deadline):
                    raise ExcelProcessError('EXCEL_FOREIGN_WORKBOOKS')
            self.state = 'ready'
            return self
        except BaseException as error:
            self.quarantine(error)
            raise

    def reserve_workbook(self, path: Path) -> None:
        """Retain the planned staging path before the first open/create event."""
        self._available()
        if self.state != 'ready' or self.owned_path is not None or not Path(path).is_absolute():
            raise ValueError('An unbound owner and an absolute workbook path are required')
        self.owned_path = Path(path)

    def claim_workbook(self, path: Path, *, deadline: float) -> None:
        """Bind the unique staged workbook after the caller creates/opens it.

        The caller reserves the planned path before its open/create command. A second claim cannot replace a
        binding. No path is inferred from the foreground/active workbook.
        """
        self._available()
        if self.state != 'ready' or not Path(path).is_absolute() or (self.owned_path is None or self.owned_path != Path(path)):
            raise ValueError('An unbound owner and an absolute workbook path are required')
        self.owned_path = Path(path)
        try:
            books = self._read_inventory(deadline)
            if len(books) != 1 or books[0]['full_name'] != str(self.owned_path) or books[0]['name'] != self.owned_path.name:
                raise ExcelProcessError('EXCEL_FOREIGN_WORKBOOKS')
            if books[0]['saved'] is not True:
                raise ExcelProcessError('EXCEL_WORKBOOK_NOT_SAVED')
            self.state = 'owned'
        except BaseException as error:
            self.quarantine(error)
            raise

    def run(self, script_path: Path, deadline: float) -> Any:
        self._available()
        if self.state == 'ready' and self.owned_path is None:
            raise ExcelProcessError('EXCEL_WORKBOOK_NOT_RESERVED')
        try:
            result = self._run(script_path, deadline)
            self._same_process()
            return result
        except BaseException as error:
            self.quarantine(error)
            raise

    def _run(self, script_path: Path, deadline: float) -> Any:
        self._deadline(deadline)
        self._same_process()
        result = self._transport.run(script_path, deadline)
        if result.returncode:
            raise OSATransportError('EXCEL_COMMAND_FAILED', pid=self.identity.pid,
                stdout=result.stdout, stderr=result.stderr, outcome_uncertain=True)
        self._deadline(deadline)
        return result

    def _command(self, operation: str, deadline: float, *, book: Any = None,
                 acknowledgment: Optional[str] = None) -> dict[str, Any]:
        nonce = uuid4().hex
        source = _script(operation, nonce, book).replace(_quote(str(_APP)), _quote(str(self._app_path)))
        folder = self.evidence_dir / (operation + '-' + nonce)
        folder.mkdir(mode=0o700)
        script_path = folder / 'owner.applescript'
        script_path.write_text(source, encoding='utf-8')
        diagnostic = {
            'operation': operation, 'script_path': str(script_path),
            'stdout_path': str(folder / 'stdout.txt'),
            'stderr_path': str(folder / 'stderr.txt'),
            'diagnostic_path': str(folder / 'diagnostic.json'),
            'stdout': '', 'stderr': '', 'returncode': None, 'error': None,
            'diagnostic_errors': [],
        }
        self.last_command = diagnostic
        try:
            result = self._run(script_path, deadline)
            diagnostic.update(stdout=result.stdout, stderr=result.stderr,
                              returncode=result.returncode)
            if operation != 'quit':
                self._same_process()
            try:
                value = json.loads(result.stdout)
            except (TypeError, ValueError):
                raise ExcelProcessError('EXCEL_BAD_ACK') from None
            if not isinstance(value, dict) or value.get('nonce') != nonce:
                raise ExcelProcessError('EXCEL_BAD_ACK')
            if acknowledgment and (set(value) != {'nonce', acknowledgment} or value.get(acknowledgment) is not True):
                raise ExcelProcessError('EXCEL_BAD_ACK')
            return value
        except BaseException as error:
            for stream in ('stdout', 'stderr'):
                raw = getattr(error, stream, None)
                if raw is not None:
                    text = raw.decode('utf-8', errors='replace') if isinstance(raw, bytes) else str(raw)
                    if diagnostic['returncode'] is None:
                        diagnostic[stream] = text
                    elif text:
                        diagnostic.setdefault('error_output', {})[stream] = text
            diagnostic['error'] = {
                'code': getattr(error, 'code', type(error).__name__),
                'detail': str(error),
            }
            raise
        finally:
            # Evidence is best-effort when preserving a primary native failure.
            # Every destination gets an attempt even if an earlier write failed.
            cancellation: Optional[BaseException] = None
            for stream in ('stdout', 'stderr', 'diagnostic'):
                destination = diagnostic[stream + '_path']
                try:
                    content = (json.dumps(diagnostic, ensure_ascii=False, indent=2)
                               if stream == 'diagnostic' else diagnostic[stream] or '')
                    Path(destination).write_text(content, encoding='utf-8')
                except BaseException as write_error:
                    if not isinstance(write_error, Exception) and cancellation is None:
                        cancellation = write_error
                    diagnostic['diagnostic_errors'].append({
                        'path': destination,
                        'code': type(write_error).__name__,
                        'detail': str(write_error),
                    })
            if diagnostic['error'] is None:
                if cancellation is not None:
                    raise cancellation
                if diagnostic['diagnostic_errors']:
                    raise ExcelProcessError('EXCEL_DIAGNOSTIC_WRITE_FAILED')

    def _read_inventory(self, deadline: float) -> list[dict[str, Any]]:
        value = self._command('inventory', deadline)
        books = value.get('workbooks')
        if set(value) != {'nonce', 'version', 'workbooks'} or not isinstance(value['version'], str) or not value['version'].strip() or not isinstance(books, list):
            raise ExcelProcessError('EXCEL_BAD_ACK')
        for book in books:
            if not isinstance(book, dict) or set(book) != {'name', 'path', 'full_name', 'saved', 'pristine'} or any(not isinstance(book[k], str) for k in ('name', 'path', 'full_name')) or any(type(book[k]) is not bool for k in ('saved', 'pristine')) or not book['name'] or not book['full_name']:
                raise ExcelProcessError('EXCEL_BAD_ACK')
        self._inventory = books
        return books

    def close(self, *, deadline: float) -> None:
        self._available()
        try:
            if self.owned_path is not None:
                self._command('close', deadline, book={'name': self.owned_path.name,
                    'full_name': str(self.owned_path)}, acknowledgment='empty')
            if self._read_inventory(deadline):
                raise ExcelProcessError('EXCEL_FOREIGN_WORKBOOKS')
            self._command('quit', deadline, acknowledgment='quit_sent')
            interval = 0.01
            while True:
                self._deadline(deadline)
                current = self._launcher.snapshot(self.identity.pid)
                same_birth = current is not None and (current.start_seconds, current.start_microseconds) == (self.identity.start_seconds, self.identity.start_microseconds)
                if self._launcher.has_terminated(self._launch) and not same_birth:
                    release = getattr(self._launcher, 'release', None)
                    if release is not None:
                        release()
                    self.state = 'closed'
                    return
                self._launcher.wait_until(min(deadline, self._clock() + interval))
                interval = min(interval * 2, 0.25)
        except BaseException as error:
            self.quarantine(error)
            raise


def _pristine_record(book: dict[str, Any]) -> bool:
    return (book['pristine'] is True and book['saved'] is True and
            book['path'] == '' and book['full_name'] == book['name'])


def _quote(value: str) -> str:
    return '"' + value.replace('\\', '\\\\').replace('"', '\\"').replace('\r', '\\r').replace('\n', '\\n') + '"'


_JSON = '''use framework "Foundation"
use scripting additions
on j(v)
 set box to current application's NSArray's arrayWithObject:v
 set d to current application's NSJSONSerialization's dataWithJSONObject:box options:0 |error|:(missing value)
 set s to (current application's NSString's alloc()'s initWithData:d encoding:4) as text
 return text 2 thru -2 of s
end j
on pristine(b)
 tell application "/Applications/Microsoft Excel.app"
  try
   if saved of b is not true or path of b is not "" or full name of b is not name of b then return false
   if has vb project of b or is add in of b then return false
   if (count of sheets of b) is not 1 or (count of worksheets of b) is not 1 then return false
   if (count of named items of b) is not 0 then return false
   set s to worksheet 1 of b
   if (count of shapes of s) is not 0 or (count of chart objects of s) is not 0 then return false
   if (count of named items of s) is not 0 or (count of Excel comments of s) is not 0 then return false
   if (count of hyperlinks of s) is not 0 or (count of query tables of s) is not 0 then return false
   if (count of cells of used range of s) is not 1 then return false
   if (get address of used range of s) is not "$A$1" then return false
   if value of range "A1" of s is not "" or formula of range "A1" of s is not "" then return false
   return true
  on error
   return false
  end try
 end tell
end pristine
'''


def _script(operation: str, nonce: str, book: Any) -> str:
    prefix = _JSON + '\n-- owner operation: ' + operation + '\nset ownerNonce to ' + _quote(nonce) + '\ntell application "/Applications/Microsoft Excel.app"\n'
    if operation == 'inventory':
        body = '''set rowsJSON to ""
repeat with ownerIndex from 1 to (count of workbooks)
 set b to workbook ownerIndex
 if rowsJSON is not "" then set rowsJSON to rowsJSON & ","
 set rowsJSON to rowsJSON & "{\\"name\\":" & my j(name of b) & ",\\"path\\":" & my j(path of b) & ",\\"full_name\\":" & my j(full name of b) & ",\\"saved\\":" & my j(saved of b) & ",\\"pristine\\":" & my j(my pristine(b)) & "}"
end repeat
return "{\\"nonce\\":" & my j(ownerNonce) & ",\\"version\\":" & my j(version as text) & ",\\"workbooks\\":[" & rowsJSON & "]}"
'''
    elif operation in ('bootstrap', 'close'):
        body = 'if (count of workbooks) is not 1 then error "Foreign workbooks"\nset b to workbook ' + _quote(book['name']) + '\nif full name of b is not ' + _quote(book['full_name']) + ' then error "Workbook identity mismatch"\n'
        if operation == 'bootstrap':
            body += 'if not my pristine(b) then error "Bootstrap workbook changed"\n'
        body += 'close b saving no\nreturn "{\\"nonce\\":" & my j(ownerNonce) & ",\\"empty\\":" & my j((count of workbooks) is 0) & "}"\n'
    elif operation == 'quit':
        body = 'if (count of workbooks) is not 0 then error "Foreign workbooks"\nquit saving no\n'
        # No further AppleEvents may address Excel after quit.
        return prefix + body + 'end tell\nreturn "{\\"nonce\\":" & my j(ownerNonce) & ",\\"quit_sent\\":true}"\n'
    else:
        raise ValueError('Unknown owner operation')
    return prefix + body + 'end tell\n'


class AppKitExcelLauncher:
    """Retain NSWorkspace's returned instance, never identify it by PID difference."""

    def __init__(self, *, clock: Callable[[], float] = time.monotonic,
                 observer: Optional[Callable[[str, dict[str, Any]], None]] = None) -> None:
        from .macos_osa_transport import _DarwinRuntime, _DarwinProcessLookup
        self.runtime = _DarwinRuntime()
        self._lookup = _DarwinProcessLookup(self.runtime)
        self._clock = clock
        self._retained: list[Any] = []
        self._applications: dict[int, Any] = {}
        self._observer = observer
        self.recovery_candidate: Optional[dict[str, Any]] = None

    def existing_pids(self) -> set[int]:
        import ctypes as C
        r = self.runtime
        workspace = r.message(r.objc_class('NSWorkspace'), 'sharedWorkspace')
        apps = r.message(workspace, 'runningApplications')
        return {r.message(r.message(apps, 'objectAtIndex:', (i,), (C.c_ulong,)),
            'processIdentifier', result=C.c_int32)
            for i in range(r.message(apps, 'count', result=C.c_ulong))}

    def launch(self, app_path: Path, deadline: float) -> LaunchedExcel:
        import ctypes as C
        if self._clock() >= deadline:
            raise ExcelProcessError('EXCEL_DEADLINE')
        r = self.runtime
        pool = r.message(r.objc_class('NSAutoreleasePool'), 'new')
        app = None
        try:
            url = r.message(r.objc_class('NSURL'), 'fileURLWithPath:',
                (r.ns_string(str(app_path)),), (C.c_void_p,))
            workspace = r.message(r.objc_class('NSWorkspace'), 'sharedWorkspace')
            config = r.message(r.objc_class('NSDictionary'), 'dictionary')
            error = C.c_void_p()
            # NSWorkspaceLaunchNewInstance | NSWorkspaceLaunchAsync. Activation
            # is allowed: background launch has been observed to silently no-op.
            app = r.message(workspace, 'launchApplicationAtURL:options:configuration:error:',
                (url, 0x80000 | 0x10000, config, C.byref(error)),
                (C.c_void_p, C.c_ulong, C.c_void_p, C.POINTER(C.c_void_p)))
            if not app:
                raise ExcelProcessError('EXCEL_LAUNCH_FAILED')
            r.message(app, 'retain')
            self._retained.append(app)
            pid = r.message(app, 'processIdentifier', result=C.c_int32)
            self.recovery_candidate = {'pid': pid, 'verified': False}
            if self._observer is not None:
                self._observer('candidate', dict(self.recovery_candidate))
            if pid <= 0:
                raise ExcelProcessError('EXCEL_LAUNCH_IDENTITY_UNAVAILABLE')
            self._applications[pid] = app
            bundle = r.message(app, 'bundleIdentifier')
            executable = r.message(app, 'executableURL')
            date = r.message(app, 'launchDate')
            launch_date = r.message(date, 'timeIntervalSince1970', result=C.c_double) if date else None
            self.recovery_candidate.update(
                bundle_id=r.object_text(bundle) if bundle else None,
                executable=os.path.realpath(r.object_text(r.message(executable, 'path'))) if executable else None,
                launch_date=launch_date,
            )
            if self._observer is not None:
                self._observer('candidate', dict(self.recovery_candidate))
            identity = self.snapshot(pid)
            if identity is None or not bundle or not executable:
                raise ExcelProcessError('EXCEL_LAUNCH_IDENTITY_UNAVAILABLE')
            if r.object_text(bundle) != identity.bundle_id or os.path.realpath(r.object_text(r.message(executable, 'path'))) != identity.executable:
                raise ExcelProcessError('EXCEL_LAUNCH_IDENTITY_MISMATCH')
            self.recovery_candidate.update(identity=asdict(identity))
            return LaunchedExcel(identity, launch_date, app)
        finally:
            if pool:
                r.message(pool, 'drain')

    def snapshot(self, pid: int) -> Optional[ExcelProcessIdentity]:
        return self._lookup.snapshot(pid, application=self._applications.get(pid))

    def finished_launching(self, record: LaunchedExcel) -> bool:
        import ctypes as C
        return self.runtime.message(record.handle, 'isFinishedLaunching', result=C.c_bool)

    def has_terminated(self, record: LaunchedExcel) -> bool:
        import ctypes as C
        return self.runtime.message(record.handle, 'isTerminated', result=C.c_bool)

    def wait_until(self, deadline: float) -> None:
        import ctypes as C
        r = self.runtime
        remaining = max(0.0, deadline - self._clock())
        date = r.message(r.objc_class('NSDate'), 'dateWithTimeIntervalSinceNow:',
            (remaining,), (C.c_double,))
        loop = r.message(r.objc_class('NSRunLoop'), 'currentRunLoop')
        r.message(loop, 'runUntilDate:', (date,), (C.c_void_p,), None)

    def __del__(self) -> None:
        # Release only our Objective-C references; never terminate applications.
        for app in getattr(self, '_retained', ()):
            self.runtime.message(app, 'release')
