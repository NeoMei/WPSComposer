"""Word one-shot worker evidence faults must not change primary outcomes."""
from __future__ import annotations

import errno
import json
from pathlib import Path
import subprocess
import sys
import time

import pytest

from skills.WPSComposer.scripts.msoffice import windows_runtime as runtime


@pytest.fixture
def isolated(monkeypatch, tmp_path):
    shared = tmp_path / 'writer'
    locks = []

    class Lock:
        def __init__(self, root):
            self.quarantine_path = root / 'native-office.quarantine.json'
            self.closed = False
            locks.append(self)

        def acquire(self, deadline):
            if self.quarantine_path.exists():
                raise runtime.NativeOfficeError('NATIVE_OFFICE_QUARANTINED')

        def quarantine(self, detail):
            self.quarantine_path.write_text(json.dumps(detail))

        def close(self):
            self.closed = True

    monkeypatch.setattr(runtime, '_component_root', lambda component: shared)
    monkeypatch.setattr(runtime, 'OfficeJobLock', Lock)
    return shared, locks, Lock


def execute(tmp_path):
    return runtime._run_worker(tmp_path, {'action': 'export'}, time.monotonic() + 3)


@pytest.mark.parametrize('primary', [runtime.NativeWordTimeoutError(), KeyboardInterrupt()])
@pytest.mark.parametrize('fault', ['quarantine', 'close'])
def test_shared_lock_fault_preserves_primary_identity(monkeypatch, tmp_path, isolated, primary, fault):
    shared, locks, lock_type = isolated

    def worker(*args):
        raise primary

    def fail(*args):
        raise OSError(errno.EIO, 'controlled lock failure')

    monkeypatch.setattr(runtime, '_run_worker_unlocked', worker)
    monkeypatch.setattr(lock_type, fault, fail)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    if fault == 'quarantine':
        assert primary.diagnostic_io_failures == (('quarantine', 'OSError'),)
        assert primary.quarantine_path is None
        assert locks[0].closed
    else:
        assert primary.cleanup_io_failures == (('lock.close', 'OSError'),)
        assert (shared / 'native-office.quarantine.json').is_file()


class TimedOutChild:
    pid = 723
    returncode = None

    def wait(self, timeout):
        if self.returncode is None:
            raise subprocess.TimeoutExpired('owned Python', timeout)

    def kill(self):
        self.returncode = -9


def fail_diagnostic_writes(monkeypatch):
    write_json = runtime._write_json

    def write(path, value):
        if path.name == 'diagnostics.json':
            raise OSError(errno.ENOSPC, 'controlled diagnostic storage failure')
        return write_json(path, value)

    monkeypatch.setattr(runtime, '_write_json', write)


def test_timeout_diagnostic_failure_keeps_timeout_and_shared_quarantine(monkeypatch, tmp_path, isolated):
    child = TimedOutChild()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: child)
    fail_diagnostic_writes(monkeypatch)
    with pytest.raises(runtime.NativeWordTimeoutError) as caught:
        execute(tmp_path)
    assert child.returncode == -9
    assert caught.value.diagnostic_io_failures == (('diagnostics', 'OSError'),)
    assert Path(caught.value.diagnostic_path).name == 'worker.log'
    assert Path(caught.value.diagnostic_path).is_file()
    assert Path(caught.value.quarantine_path).is_file()
    with pytest.raises(runtime.NativeWordError) as blocked:
        execute(tmp_path)
    assert blocked.value.code == 'NATIVE_WORD_QUARANTINED'


@pytest.mark.parametrize('primary', [OSError(errno.EACCES, 'Python unavailable'), KeyboardInterrupt(), SystemExit(7)])
def test_launch_failure_is_known_clean_even_when_diagnostic_write_fails(monkeypatch, tmp_path, isolated, primary):
    def launch(*args, **kwargs):
        raise primary

    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    fail_diagnostic_writes(monkeypatch)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    if isinstance(primary, Exception):
        assert caught.value.code == 'NATIVE_WORD_EXECUTION_FAILED'
    else:
        assert caught.value is primary
    assert caught.value.cleanup_verified is True
    assert caught.value.diagnostic_io_failures == (('diagnostics', 'OSError'),)
    assert not (isolated[0] / 'native-office.quarantine.json').exists()
    assert isolated[1][0].closed


def test_request_write_failure_before_launch_is_known_clean(monkeypatch, tmp_path, isolated):
    primary = OSError(errno.ENOSPC, 'request storage failure')

    def fail(*args):
        raise primary

    monkeypatch.setattr(runtime, '_write_json', fail)
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: pytest.fail('launched'))
    with pytest.raises(OSError) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert primary.cleanup_verified is True
    assert not (isolated[0] / 'native-office.quarantine.json').exists()


@pytest.mark.parametrize('primary', [KeyboardInterrupt(), SystemExit(8)])
def test_wait_cancellation_stops_owned_python_and_preserves_identity(monkeypatch, tmp_path, isolated, primary):
    class Child(TimedOutChild):
        def wait(self, timeout):
            if self.returncode is None:
                raise primary

    child = Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: child)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    assert caught.value is primary
    assert child.returncode == -9
    assert (isolated[0] / 'native-office.quarantine.json').is_file()


def test_timeout_cleanup_errors_do_not_replace_timeout(monkeypatch, tmp_path, isolated):
    class Child(TimedOutChild):
        def kill(self):
            raise OSError(errno.EACCES, 'owned Python signal refused')

    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: Child())
    with pytest.raises(runtime.NativeWordTimeoutError) as caught:
        execute(tmp_path)
    assert caught.value.cleanup_io_failures == (('worker.kill', 'PermissionError'),
                                               ('worker.wait', 'TimeoutExpired'))
    assert (isolated[0] / 'native-office.quarantine.json').is_file()


def test_diagnostic_metadata_failure_cannot_skip_child_cleanup(monkeypatch, tmp_path, isolated):
    child = TimedOutChild()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: child)
    stat = Path.stat

    def fail_stat(path, *args, **kwargs):
        if path.name == 'diagnostics.json':
            raise OSError(errno.EACCES, 'diagnostic stat refused')
        return stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, 'stat', fail_stat)
    with pytest.raises(runtime.NativeWordTimeoutError) as caught:
        execute(tmp_path)
    assert child.returncode == -9
    assert caught.value.diagnostic_path is None or Path(caught.value.diagnostic_path).name == 'worker.log'
    assert (isolated[0] / 'native-office.quarantine.json').is_file()


@pytest.mark.parametrize('data', ['[]', 'null', 'false'])
def test_non_object_diagnostics_fall_back_to_typed_failure(tmp_path, data):
    (tmp_path / 'diagnostics.json').write_text(data)
    error = runtime._worker_failure(tmp_path)
    assert error.code == 'NATIVE_WORD_EXECUTION_FAILED'
    assert error.cleanup_verified is False
    assert Path(error.diagnostic_path).is_file()


def test_worker_failure_never_advertises_absent_evidence(tmp_path):
    error = runtime._worker_failure(tmp_path)
    assert error.diagnostic_path is None


def test_success_does_not_swallow_lock_close_failure(monkeypatch, tmp_path, isolated):
    primary = OSError(errno.EIO, 'unlock failed')
    monkeypatch.setattr(runtime, '_run_worker_unlocked', lambda *args: {'success': True})

    def fail(lock):
        raise primary

    monkeypatch.setattr(isolated[2], 'close', fail)
    with pytest.raises(OSError) as caught:
        execute(tmp_path)
    assert caught.value is primary


def test_real_python_wait_cancellation_does_not_kill_unrelated_child(monkeypatch, tmp_path, isolated):
    real_popen = subprocess.Popen
    unrelated = real_popen([sys.executable, '-c', 'import time; time.sleep(30)'])
    children = []
    primary = KeyboardInterrupt('cancel owned worker')

    def launch(command, **kwargs):
        child = real_popen([sys.executable, '-c', 'import time; time.sleep(30)'], **kwargs)
        children.append(child)
        wait = child.wait

        def cancelled_wait(timeout):
            child.wait = wait
            raise primary

        child.wait = cancelled_wait
        return child

    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    try:
        with pytest.raises(KeyboardInterrupt) as caught:
            execute(tmp_path)
        assert caught.value is primary
        assert children[0].poll() is not None
        assert unrelated.poll() is None
    finally:
        for child in [unrelated] + children:
            if child.poll() is None:
                child.kill()
            child.wait(timeout=2)


@pytest.mark.parametrize('outcome', ['timeout', 'success'])
def test_worker_log_close_preserves_primary_but_propagates_after_success(monkeypatch, tmp_path, isolated, outcome):
    close_error = OSError(errno.EIO, 'controlled log close failure')
    open_path = Path.open

    class Log:
        def __init__(self, stream):
            self.stream = stream

        def __enter__(self):
            return self

        def __exit__(self, *args):
            self.close()

        def close(self):
            self.stream.close()
            raise close_error

    def open_log(path, *args, **kwargs):
        stream = open_path(path, *args, **kwargs)
        return Log(stream) if path.name == 'worker.log' else stream

    def launch(command, **kwargs):
        child = TimedOutChild()
        if outcome == 'success':
            child.returncode = 0
            Path(command[-1]).write_text(json.dumps({'status': 'ok', 'value': {'native': 'success'}}))
        return child

    monkeypatch.setattr(Path, 'open', open_log)
    monkeypatch.setattr(runtime.subprocess, 'Popen', launch)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    if outcome == 'timeout':
        assert isinstance(caught.value, runtime.NativeWordTimeoutError)
        assert caught.value.cleanup_io_failures == (('worker.log.close', 'OSError'),)
    else:
        assert caught.value is close_error


@pytest.mark.parametrize('primary_kind', ['timeout', 'cancel'])
@pytest.mark.parametrize('secondary_type', [KeyboardInterrupt, SystemExit])
@pytest.mark.parametrize('target', ['diagnostics.json', 'native-office.quarantine.json'])
def test_secondary_evidence_cancellation_preserves_primary(monkeypatch, tmp_path, isolated, primary_kind, secondary_type, target):
    primary_cancel = KeyboardInterrupt('original wait cancellation')
    secondary = secondary_type('secondary metadata cancellation')

    class Child(TimedOutChild):
        def wait(self, timeout):
            if self.returncode is None and primary_kind == 'cancel':
                raise primary_cancel
            return super().wait(timeout)

    child = Child()
    monkeypatch.setattr(runtime.subprocess, 'Popen', lambda *args, **kwargs: child)
    is_file = Path.is_file

    def evidence(path):
        if path.name == target:
            raise secondary
        return is_file(path)

    monkeypatch.setattr(Path, 'is_file', evidence)
    with pytest.raises(BaseException) as caught:
        execute(tmp_path)
    error = caught.value
    if primary_kind == 'cancel':
        assert error is primary_cancel
    else:
        assert isinstance(error, runtime.NativeWordTimeoutError)
    assert error is not secondary
    assert error.diagnostic_io_failures == ((target + '.stat', secondary_type.__name__),)
    assert child.returncode == -9
    assert isolated[1][0].closed
    quarantine = isolated[0] / 'native-office.quarantine.json'
    data = json.loads(quarantine.read_text())
    assert data['code'] == ('NATIVE_WORD_TIMEOUT' if primary_kind == 'timeout' else 'NATIVE_WORD_EXECUTION_FAILED')
    if target == 'native-office.quarantine.json':
        assert error.quarantine_path is None
    else:
        assert error.diagnostic_path == str(next(tmp_path.glob('operation-*/worker.log')))
    with pytest.raises(runtime.NativeWordError) as blocked:
        execute(tmp_path)
    assert blocked.value.code == 'NATIVE_WORD_QUARANTINED'
