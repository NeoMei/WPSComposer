"""Keep uncertain Office jobs isolated when evidence or cleanup I/O fails."""
from pathlib import Path
from types import SimpleNamespace
import sys
import errno
import subprocess
import time

import pytest

from skills.WPSComposer.scripts.msoffice import windows_office_runtime as runtime


@pytest.fixture
def isolated(monkeypatch, tmp_path):
    root = tmp_path / 'private'
    root.mkdir()
    events = []

    class Lock:
        def __init__(self, directory):
            self.quarantine_path = directory / 'quarantine.json'

        def acquire(self, deadline):
            if self.quarantine_path.exists():
                raise runtime.NativeOfficeError('NATIVE_OFFICE_QUARANTINED')

        def quarantine(self, detail):
            events.append('quarantine')
            self.quarantine_path.write_text('uncertain')

        def close(self):
            events.append('close')

    monkeypatch.setattr(runtime, '_component_root', lambda component: root)
    monkeypatch.setattr(runtime, 'OfficeJobLock', Lock)
    monkeypatch.setattr(runtime, 'engine_executable', lambda *a: 'native.exe')
    monkeypatch.setattr(runtime, 'sys', SimpleNamespace(platform='win32', executable=sys.executable))
    return root, events, Lock


def execute(tmp_path):
    return runtime._execute('spreadsheet', tmp_path / 'result.xlsx', overwrite=False,
                            deadline=time.monotonic() + 30, prepare=lambda *a: {})


@pytest.mark.parametrize('primary', [runtime.NativeOfficeError('NATIVE_OFFICE_TIMEOUT'), KeyboardInterrupt()])
def test_recovery_write_failure_preserves_primary_and_quarantines(monkeypatch, tmp_path, isolated, primary):
    root, events, _ = isolated
    def worker(*args):
        raise primary
    monkeypatch.setattr(runtime, '_run_worker', worker)
    def fail_write(path, detail):
        events.append('recovery')
        raise OSError('disk full')
    monkeypatch.setattr(runtime, '_write_json', fail_write)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert events == ['quarantine', 'recovery', 'close']
    assert (root / 'quarantine.json').exists()
    with pytest.raises(runtime.NativeOfficeError) as blocked:
        execute(tmp_path)
    assert blocked.value.code == 'NATIVE_OFFICE_QUARANTINED'
    assert not (tmp_path / 'result.xlsx').exists()
    if isinstance(primary, runtime.NativeOfficeError):
        assert primary.diagnostic_path is None
        assert primary.quarantine_path == str(root / 'quarantine.json')


def test_lock_cleanup_failure_keeps_native_error(monkeypatch, tmp_path, isolated):
    _, _, lock = isolated
    primary = runtime.NativeOfficeError('NATIVE_OFFICE_TIMEOUT')
    def worker(*args):
        raise primary
    def close(self):
        raise OSError('close failed')
    monkeypatch.setattr(runtime, '_run_worker', worker)
    monkeypatch.setattr(lock, 'close', close)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert primary.cleanup_io_failures == (('lock.close', 'OSError'),)


def test_failed_quarantine_write_keeps_primary_without_claiming_persistence(monkeypatch, tmp_path, isolated):
    root, _, lock = isolated
    primary = runtime.NativeOfficeError('NATIVE_OFFICE_TIMEOUT')
    def worker(*args):
        raise primary
    def quarantine(*args):
        raise OSError('disk full')
    monkeypatch.setattr(runtime, '_run_worker', worker)
    monkeypatch.setattr(lock, 'quarantine', quarantine)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert primary.quarantine_path is None
    assert not (root / 'quarantine.json').exists()
    assert primary.diagnostic_io_failures == (('quarantine', 'OSError'),)


def test_closed_native_failure_does_not_quarantine(monkeypatch, tmp_path, isolated):
    root, events, _ = isolated
    primary = runtime.NativeOfficeError('NATIVE_OFFICE_EXECUTION_FAILED')
    primary.cleanup_verified = True
    def worker(*args):
        raise primary
    monkeypatch.setattr(runtime, '_run_worker', worker)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert not (root / 'quarantine.json').exists()
    assert 'quarantine' not in events
    assert primary.quarantine_path is None


def test_cleanup_failure_after_publication_is_not_silenced(monkeypatch, tmp_path, isolated):
    _, _, lock = isolated
    failure = OSError('unlock failed')
    def worker(job, *args):
        target = job / 'result.xlsx'
        target.write_bytes(b'controlled native output')
        return {'path': str(target), 'cleanup_verified': True}
    def publish(source, output, **kwargs):
        output.write_bytes(source.read_bytes())
    def close(self):
        raise failure
    monkeypatch.setattr(runtime, '_run_worker', worker)
    monkeypatch.setattr(runtime, '_publish', publish)
    monkeypatch.setattr(lock, 'close', close)
    with pytest.raises(OSError) as caught:
        execute(tmp_path)
    assert caught.value is failure
    assert (tmp_path / 'result.xlsx').read_bytes() == b'controlled native output'


class TimedOutChild:
    pid = 812
    returncode = None
    def wait(self, timeout):
        if self.returncode is None:
            raise subprocess.TimeoutExpired('synthetic-python-worker', timeout)
    def kill(self):
        self.returncode = -9


def test_worker_timeout_diagnostic_failure_keeps_typed_timeout(monkeypatch, tmp_path):
    child = TimedOutChild()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: child)
    write = runtime._write_json
    def diagnostic_failure(path, value):
        if path.name == 'diagnostics.json':
            raise OSError(errno.ENOSPC, 'injected diagnostic write failure')
        return write(path, value)
    monkeypatch.setattr(runtime, '_write_json', diagnostic_failure)
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    assert child.returncode == -9
    assert getattr(caught.value, 'code', None) == 'NATIVE_OFFICE_TIMEOUT'


def test_failed_python_launch_with_diagnostic_failure_does_not_quarantine(monkeypatch, tmp_path, isolated):
    root, _, _ = isolated
    def launch(*args, **kwargs):
        raise OSError(errno.EACCES, 'synthetic child launch refused')
    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    write = runtime._write_json
    def diagnostic_failure(path, value):
        if path.name == 'diagnostics.json':
            raise OSError(errno.ENOSPC, 'injected diagnostic write failure')
        return write(path, value)
    monkeypatch.setattr(runtime, '_write_json', diagnostic_failure)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert not (root / 'quarantine.json').exists()
    assert getattr(caught.value, 'cleanup_verified', False) is True


def test_abrupt_worker_exit_reports_existing_recovery_path(monkeypatch, tmp_path, isolated):
    class FailedChild:
        pid = 812
        returncode = 1
        def wait(self, timeout):
            return 1
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: FailedChild())
    with pytest.raises(runtime.NativeOfficeError) as caught:
        execute(tmp_path)
    path = caught.value.diagnostic_path
    assert path is None or Path(path).is_file(), path
    assert Path(caught.value.staging_path, 'recovery.json').is_file()


@pytest.mark.parametrize('signal_type', [KeyboardInterrupt, SystemExit])
def test_launch_cancellation_keeps_identity_when_diagnostic_write_fails(monkeypatch, tmp_path, isolated, signal_type):
    root, _, _ = isolated
    primary = signal_type()
    def launch(*args, **kwargs):
        raise primary
    write = runtime._write_json
    def diagnostic_failure(path, value):
        if path.name == 'diagnostics.json':
            raise OSError('diagnostic failure')
        return write(path, value)
    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    monkeypatch.setattr(runtime, '_write_json', diagnostic_failure)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert primary.cleanup_verified is True
    assert not (root / 'quarantine.json').exists()


@pytest.mark.parametrize('value', ['[]', 'null', 'false'])
def test_non_object_worker_diagnostic_is_classified_as_native_failure(tmp_path, value):
    (tmp_path / 'diagnostics.json').write_text(value)
    error = runtime._worker_failure(tmp_path)
    assert error.code == 'NATIVE_OFFICE_EXECUTION_FAILED'
    assert error.cleanup_verified is False
    assert Path(error.diagnostic_path).is_file()


@pytest.mark.parametrize('signal_type', [KeyboardInterrupt, SystemExit])
def test_wait_cancellation_stops_only_owned_python_and_keeps_identity(monkeypatch, tmp_path, signal_type):
    primary = signal_type()
    class Child(TimedOutChild):
        def wait(self, timeout):
            if self.returncode is None:
                raise primary
    child = Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: child)
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    assert caught.value is primary
    assert child.returncode == -9


def test_timeout_child_cleanup_failure_does_not_replace_timeout(monkeypatch, tmp_path):
    class Child(TimedOutChild):
        def kill(self):
            raise OSError('cannot signal owned Python')
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: Child())
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    assert getattr(caught.value, 'code', None) == 'NATIVE_OFFICE_TIMEOUT'
    assert caught.value.cleanup_io_failures == (('worker.kill', 'OSError'), ('worker.wait', 'TimeoutExpired'))


@pytest.mark.parametrize('failed_name', ['diagnostics.json', 'worker.log'])
def test_diagnostic_stat_failure_keeps_timeout_and_stops_python(monkeypatch, tmp_path, failed_name):
    class Child:
        pid = 812
        returncode = None
        def wait(self, timeout):
            if self.returncode is None:
                raise subprocess.TimeoutExpired('synthetic-child', timeout)
        def kill(self):
            self.returncode = -9
    child = Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: child)
    is_file = Path.is_file
    def denied(path):
        if path == tmp_path / failed_name:
            raise OSError(errno.EACCES, 'injected evidence metadata failure')
        return is_file(path)
    monkeypatch.setattr(Path, 'is_file', denied)
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    observed = (getattr(caught.value, 'code', None), child.returncode)
    assert observed == ('NATIVE_OFFICE_TIMEOUT', -9), observed


@pytest.mark.parametrize('outcome', ['timeout', 'cancel'])
@pytest.mark.parametrize('secondary_type', [KeyboardInterrupt, SystemExit])
def test_secondary_diagnostic_cancel_does_not_replace_primary(monkeypatch, tmp_path, outcome, secondary_type):
    primary = KeyboardInterrupt('initial cancellation')
    class Child(TimedOutChild):
        def wait(self, timeout):
            if self.returncode is None:
                if outcome == 'cancel':
                    raise primary
                raise subprocess.TimeoutExpired('synthetic-child', timeout)
    child = Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: child)
    is_file = Path.is_file
    def interrupted(path):
        if path == tmp_path / 'diagnostics.json':
            raise secondary_type('secondary diagnostic interruption')
        return is_file(path)
    monkeypatch.setattr(Path, 'is_file', interrupted)
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    if outcome == 'timeout':
        assert getattr(caught.value, 'code', None) == 'NATIVE_OFFICE_TIMEOUT'
    else:
        assert caught.value is primary
    assert child.returncode == -9
    assert ('worker.diagnostic.stat', secondary_type.__name__) in caught.value.diagnostic_io_failures


@pytest.mark.parametrize('outcome', ['timeout', 'cancel', 'launch'])
def test_log_close_preserves_primary(monkeypatch, tmp_path, outcome):
    secondary = OSError(errno.EIO, 'secondary log close failure')
    primary = KeyboardInterrupt('initial cancellation')
    real_open = Path.open

    class Log:
        def __init__(self, stream): self.stream = stream
        def __enter__(self): return self
        def __exit__(self, *args): self.close()
        def close(self):
            self.stream.close()
            raise secondary

    class Child:
        pid = 812
        returncode = None
        def wait(self, timeout):
            if self.returncode is None:
                if outcome == 'cancel': raise primary
                raise subprocess.TimeoutExpired('exact owned Python', timeout)
        def kill(self): self.returncode = -9

    child = Child()
    def open_log(path, *args, **kwargs):
        stream = real_open(path, *args, **kwargs)
        return Log(stream) if path.name == 'worker.log' else stream
    def launch(*args, **kwargs):
        if outcome == 'launch': raise PermissionError('Python unavailable')
        return child
    monkeypatch.setattr(Path, 'open', open_log)
    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    with pytest.raises(BaseException) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic()+3)
    if outcome == 'cancel':
        assert caught.value is primary
    else:
        assert getattr(caught.value, 'code', None) == ('NATIVE_OFFICE_TIMEOUT' if outcome == 'timeout' else 'NATIVE_OFFICE_EXECUTION_FAILED')
    if outcome == 'launch': assert caught.value.cleanup_verified is True
    else: assert child.returncode == -9


@pytest.mark.parametrize('phase', ['request', 'log.open', 'log.chmod'])
def test_prelaunch_preparation_failure_does_not_quarantine(monkeypatch, tmp_path, isolated, phase):
    root, _, _ = isolated
    error = OSError('prelaunch failure')
    launches = []
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: launches.append(a))
    write = runtime._write_json
    open_file = Path.open
    chmod = runtime.os.chmod
    if phase == 'request':
        def fail(path, value):
            if path.name == 'request.json':
                raise error
            return write(path, value)
        monkeypatch.setattr(runtime, '_write_json', fail)
    elif phase == 'log.open':
        def fail(path, *a, **k):
            if path.name == 'worker.log':
                raise error
            return open_file(path, *a, **k)
        monkeypatch.setattr(Path, 'open', fail)
    else:
        def fail(path, *a, **k):
            if Path(path).name == 'worker.log':
                raise error
            return chmod(path, *a, **k)
        monkeypatch.setattr(runtime.os, 'chmod', fail)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is error
    assert error.cleanup_verified is True
    assert launches == []
    assert not (root / 'quarantine.json').exists()


def test_worker_log_close_failure_after_verified_cleanup_still_propagates(monkeypatch, tmp_path):
    failure = OSError('close failed after completed worker')
    open_file = Path.open
    class Log:
        def __init__(self, stream):
            self.stream = stream
        def close(self):
            self.stream.close()
            raise failure
    def open_log(path, *args, **kwargs):
        stream = open_file(path, *args, **kwargs)
        return Log(stream) if path.name == 'worker.log' else stream
    class Child:
        pid = 812
        returncode = 0
        def wait(self, timeout):
            runtime._write_json(tmp_path / 'response.json', {'status': 'ok', 'value': {'cleanup_verified': True}})
    monkeypatch.setattr(Path, 'open', open_log)
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *a, **k: Child())
    with pytest.raises(OSError) as caught:
        runtime._run_worker(tmp_path, {'component': 'spreadsheet'}, time.monotonic() + 3)
    assert caught.value is failure
    assert failure.cleanup_verified is True
