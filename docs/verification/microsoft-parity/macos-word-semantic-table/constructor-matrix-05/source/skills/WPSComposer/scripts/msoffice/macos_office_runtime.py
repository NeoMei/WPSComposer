"""Private-container native Excel/PowerPoint jobs with atomic publication.

An Apple event timeout does not cancel an Office command. Uncertain jobs keep
their private files and quarantine the component until explicitly recovered.
"""
from __future__ import annotations

import errno
import importlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
from uuid import uuid4

from ..artifact_transport import (
    ValidatorSpec, copy_file_before_deadline, publish_artifact,
    validate_before_deadline, validate_office_package, validate_pdf,
)
from ..office_engines import engine_executable, validate_timeout
from .input_validation import validate_native_input

_APPS = {'spreadsheet': ('Excel', 'com.microsoft.Excel', 'xlsx'),
         'presentation': ('PowerPoint', 'com.microsoft.Powerpoint', 'pptx')}
from .office_errors import NativeOfficeError, NATIVE_OFFICE_ERROR_CODES


def _remaining(deadline):
    budget = deadline - time.monotonic()
    if budget <= 0:
        raise NativeOfficeError('NATIVE_OFFICE_TIMEOUT')
    return budget


def _container_root(component):
    # PowerPoint 16.112.3 cannot reopen its native PPTX from Data/tmp (-9074).
    # The same bytes and quarantine attributes reopen from Data/Documents.
    folder = 'Documents' if component == 'presentation' else 'tmp'
    return Path.home() / 'Library' / 'Containers' / _APPS[component][1] / 'Data' / folder / 'wpscomposer'


def _compiler(component):
    name = 'macos_excel_script' if component == 'spreadsheet' else 'macos_powerpoint_script'
    return importlib.import_module('.' + name, __package__)


class OfficeJobLock:
    def __init__(self, root):
        self.path = root / 'native-office.lock'
        self.quarantine_path = root / 'native-office.quarantine.json'
        self.stream = None

    def acquire(self, deadline, *, recovery=False):
        fd = os.open(self.path, os.O_CREAT | os.O_RDWR | getattr(os, 'O_NOFOLLOW', 0), 0o600)
        stream = os.fdopen(fd, 'r+b')
        try:
            while True:
                _remaining(deadline)
                try:
                    if os.name == 'nt':
                        import msvcrt
                        # CRT locks can extend beyond EOF. An initialization
                        # write would race another process holding this byte.
                        stream.seek(0)
                        msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                    else:
                        import fcntl
                        fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    break
                except OSError as exc:
                    contention = (errno.EACCES, errno.EAGAIN)
                    if os.name == 'nt':
                        contention += (errno.EDEADLK,)
                    if exc.errno not in contention:
                        raise
                    time.sleep(min(.05, _remaining(deadline)))
            if self.quarantine_path.exists() and not recovery:
                raise NativeOfficeError('NATIVE_OFFICE_QUARANTINED', quarantine_path=self.quarantine_path)
            self.stream = stream
        except BaseException:
            stream.close()
            raise

    def quarantine(self, detail):
        fd = os.open(self.quarantine_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY |
                     getattr(os, 'O_NOFOLLOW', 0), 0o600)
        with os.fdopen(fd, 'w') as stream:
            json.dump(detail, stream, ensure_ascii=False)
            stream.flush()
            os.fsync(stream.fileno())

    def close(self):
        if self.stream is not None:
            try:
                if os.name == 'nt':
                    import msvcrt
                    self.stream.seek(0)
                    msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self.stream.fileno(), fcntl.LOCK_UN)
            finally:
                self.stream.close()
                self.stream = None


def _publish(staged, output, *, overwrite, deadline):
    fmt = Path(output).suffix.lower().lstrip('.')
    validator = (ValidatorSpec.from_callable(validate_pdf) if fmt == 'pdf' else
                 ValidatorSpec.from_callable(validate_office_package, fmt))
    # Publication also validates its staged copy and final destination within
    # the same deadline; a bad output never replaces a previous deliverable.
    def bounded(path):
        validate_before_deadline(validator, path, deadline)
    return publish_artifact(staged, output, overwrite=overwrite, validator=bounded, deadline=deadline)


def _execute(component, output, *, overwrite, deadline, prepare):
    if sys.platform != 'darwin' or not engine_executable('msoffice', component):
        raise NativeOfficeError('NATIVE_OFFICE_UNAVAILABLE')
    output = Path(output).expanduser().resolve()
    if output.exists() and not overwrite:
        raise FileExistsError(output)
    root = _container_root(component)
    root.mkdir(mode=0o700, parents=True, exist_ok=True)
    lock = OfficeJobLock(root)
    job = None
    launched = False
    closed = False
    published = False
    stdout = stderr = ''
    logs_attempted = False
    diagnostic_io_failures = []
    primary_error = None

    def persist_diagnostic(label, write):
        try:
            write()
        except BaseException as error:
            # Diagnostics must not replace the primary native failure or
            # cancellation. Retain only safe failure categories on the error.
            diagnostic_io_failures.append((label, type(error).__name__))
            return error
        return None

    def write_logs():
        nonlocal logs_attempted
        logs_attempted = True
        first_error = None
        for name, value in (('stdout', stdout), ('stderr', stderr)):
            path = job / (name + '.log')
            def write(path=path, value=value):
                if isinstance(value, bytes):
                    path.write_bytes(value)
                else:
                    path.write_text(value or '', encoding='utf-8')
            error = persist_diagnostic(name, write)
            if first_error is None:
                first_error = error
        return first_error

    try:
        lock.acquire(deadline)
        job = Path(tempfile.mkdtemp(prefix='job-', dir=root))
        source, staged = prepare(job, deadline)
        script = job / 'job.applescript'
        script.write_text(source, encoding='utf-8')
        os.chmod(script, 0o600)
        budget = _remaining(deadline)
        launched = True
        try:
            result = subprocess.run(['/usr/bin/osascript', str(script)],
                                    capture_output=True, text=True, timeout=budget)
        except subprocess.TimeoutExpired as exc:
            stdout, stderr = exc.stdout, exc.stderr
            raise NativeOfficeError('NATIVE_OFFICE_TIMEOUT') from None
        stdout, stderr = result.stdout, result.stderr
        marker = 'WPSCOMPOSER_MS_OFFICE_OK:' + component
        closed = result.returncode == 0 and result.stdout.strip() == marker
        if not closed or not staged.is_file():
            raise NativeOfficeError('NATIVE_OFFICE_EXECUTION_FAILED')
        log_error = write_logs()
        if log_error is not None:
            raise log_error
        _remaining(deadline)
        _publish(staged, output, overwrite=overwrite, deadline=deadline)
        published = True
        return output
    except BaseException as exc:
        primary_error = exc
        if job is not None and launched:
            detail = {'component': component, 'staging_path': str(job), 'closed': closed,
                      'code': getattr(exc, 'code', 'NATIVE_OFFICE_EXECUTION_FAILED')}
            quarantined = False
            if not closed:
                quarantined = persist_diagnostic('quarantine', lambda: lock.quarantine(detail)) is None
            recovery = job / 'recovery.json'
            recovery_written = persist_diagnostic('recovery', lambda: recovery.write_text(
                json.dumps(detail), encoding='utf-8')) is None
            if not logs_attempted:
                write_logs()
            if isinstance(exc, NativeOfficeError):
                exc.staging_path = str(job)
                exc.diagnostic_path = str(recovery) if recovery_written else None
                exc.quarantine_path = str(lock.quarantine_path) if quarantined else None
            if diagnostic_io_failures:
                exc.diagnostic_io_failures = tuple(diagnostic_io_failures)
        raise
    finally:
        cleanups = [('lock.close', lock.close)]
        if job is not None and (published or not launched):
            cleanups.append(('staging.remove', lambda: shutil.rmtree(job)))
        cleanup_failures = []
        for label, cleanup in cleanups:
            try:
                cleanup()
            except BaseException as error:
                if primary_error is None:
                    raise
                cleanup_failures.append((label, type(error).__name__))
        if cleanup_failures:
            primary_error.cleanup_io_failures = tuple(cleanup_failures)


def generate_recorded(recorded, output, *, timeout=600, overwrite=False):
    deadline = time.monotonic() + validate_timeout(timeout)
    component = recorded.plan.component
    if component not in _APPS:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')
    output = Path(output).expanduser().resolve()
    fmt = _APPS[component][2]
    if output.suffix.lower() != '.' + fmt:
        raise ValueError('Output extension must match native component')
    compiler = _compiler(component)
    # Reject unsupported plans without launching Office or staging a job.
    compiler.compile_plan(recorded.plan, {r.id: Path('/preflight') / ('image-' + str(i) + r.source_path.suffix)
                          for i, r in enumerate(recorded.resources)}, Path('/preflight') / ('owned.' + fmt),
                          timeout=_remaining(deadline))
    def prepare(job, deadline):
        resources = {}
        for index, resource in enumerate(recorded.resources):
            path = job / ('image-' + str(index) + resource.source_path.suffix.lower())
            copy_file_before_deadline(resource.source_path, path, deadline=deadline)
            resources[resource.id] = path
        native = job / ('owned-' + uuid4().hex + '.' + fmt)
        script = compiler.compile_plan(recorded.plan, resources, native, timeout=_remaining(deadline))
        return script, native
    return _execute(component, output, overwrite=overwrite, deadline=deadline, prepare=prepare)


def generate(doc, format_name, output, preset=None, *, timeout=600, overwrite=False):
    from ..recording_composers import RecordingSheetComposer, RecordingSlideComposer
    from ..renderers import sheet_renderer, slide_renderer
    renderer, composer = {'xlsx': (sheet_renderer.render, RecordingSheetComposer),
                          'pptx': (slide_renderer.render, RecordingSlideComposer)}[format_name]
    deadline = time.monotonic() + validate_timeout(timeout)
    recorded = renderer(doc, str(output), preset=preset, composer_factory=composer)
    return generate_recorded(recorded, output, timeout=_remaining(deadline), overwrite=overwrite)


def _validate_input(path, component, deadline):
    try:
        validate_native_input(path, component, deadline=deadline)
    except TimeoutError:
        raise NativeOfficeError('NATIVE_OFFICE_TIMEOUT') from None
    except ValueError:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED') from None


def convert(request, timeout=600):
    deadline = time.monotonic() + validate_timeout(timeout)
    component = request.component
    if component not in _APPS or request.source.suffix.lower() != '.' + _APPS[component][2]:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')
    _validate_input(request.source, component, deadline)
    compiler = _compiler(component)
    compiler.compile_conversion(Path('/preflight/source' + request.source.suffix),
                                Path('/preflight/result.pdf'), timeout=_remaining(deadline))
    def prepare(job, deadline):
        source = job / ('source-' + uuid4().hex + request.source.suffix.lower())
        copy_file_before_deadline(request.source, source, deadline=deadline)
        _validate_input(source, component, deadline)
        pdf = job / ('export-' + uuid4().hex + '.pdf')
        return compiler.compile_conversion(source, pdf, timeout=_remaining(deadline)), pdf
    return _execute(component, request.output, overwrite=request.overwrite, deadline=deadline, prepare=prepare)


def recover_quarantine(component, *, timeout=30):
    """Clear a component quarantine after read-only cleanup verification.

    Retained artifacts are never deleted. A running script or an open document
    in the quarantined workspace requires explicit cleanup before recovery.
    """
    from .macos_script import apple_string
    if component not in _APPS:
        raise ValueError('Unknown native Office component')
    deadline = time.monotonic() + validate_timeout(timeout)
    root = _container_root(component).resolve()
    if not root.is_dir():
        return False
    lock = OfficeJobLock(root)
    lock.acquire(deadline, recovery=True)
    try:
        if not lock.quarantine_path.exists():
            return False
        if lock.quarantine_path.is_symlink():
            raise ValueError('Quarantine metadata must be a regular file')
        data = json.loads(lock.quarantine_path.read_text(encoding='utf-8'))
        if not isinstance(data, dict) or data.get('component') != component:
            raise ValueError('Quarantine component is invalid')
        stage = Path(data.get('staging_path', '')).resolve()
        if stage.parent != root or not stage.name.startswith(('job-', 'session-')) or not stage.is_dir():
            raise ValueError('Quarantine staging identity is invalid')
        bound = data.get('bound_path', data.get('owned_path'))
        if bound is not None and Path(bound).resolve().parent != stage:
            raise RuntimeError('Attached document completion requires manual verification; recovery refused')
        process = subprocess.run(['/bin/ps', '-axo', 'command'], capture_output=True,
                                 text=True, check=True, timeout=_remaining(deadline))
        if any('osascript' in line and str(stage) in line for line in process.stdout.splitlines()):
            raise RuntimeError('Previous native script is still running; recovery refused')
        app = 'Microsoft ' + _APPS[component][0]
        if '/' + app + '.app/Contents/MacOS/' in process.stdout:
            collection = 'workbooks' if component == 'spreadsheet' else 'presentations'
            item = 'workbook' if component == 'spreadsheet' else 'presentation'
            source = f'''with timeout of {max(1, int(_remaining(deadline)))} seconds
 tell application {apple_string(app)}
  repeat with documentIndex from 1 to (count of {collection})
   set documentPath to full name of {item} documentIndex
   if documentPath is missing value then error "Document identity unavailable"
   if documentPath does not start with "/" then error "Unidentified unsaved document prevents recovery"
   if documentPath starts with {apple_string(str(stage) + '/')} then error "Task document remains open"
  end repeat
 end tell
end timeout
return "WPSCOMPOSER_RECOVERY_OK"'''
            result = subprocess.run(['/usr/bin/osascript', '-e', source], capture_output=True,
                                    text=True, timeout=_remaining(deadline))
            if result.returncode or result.stdout.strip() != 'WPSCOMPOSER_RECOVERY_OK':
                raise RuntimeError('Native document cleanup is not verified; recovery refused')
        _remaining(deadline)
        lock.quarantine_path.unlink()
        return True
    finally:
        lock.close()
