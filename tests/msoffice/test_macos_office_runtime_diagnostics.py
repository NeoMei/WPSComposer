"""Diagnostic disk failures cannot change native completion evidence."""
from __future__ import annotations

import errno
import json
import os
from pathlib import Path
import subprocess
import time
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice import macos_office_runtime as runtime


@pytest.fixture
def job_host(tmp_path, monkeypatch):
    root = tmp_path / 'container'
    output = tmp_path / 'output.xlsx'
    output.write_bytes(b'previous output')
    calls = []
    cancelled = KeyboardInterrupt('original cancellation')
    state = SimpleNamespace(outcome='native', faults=set(), writes=[], calls=calls,
                            root=root, output=output, cancelled=cancelled)
    monkeypatch.setattr(runtime.sys, 'platform', 'darwin')
    monkeypatch.setattr(runtime, 'engine_executable', lambda *args: '/fake/Excel')
    monkeypatch.setattr(runtime, '_container_root', lambda component: root)

    def run(*args, **kwargs):
        calls.append('launch')
        if state.outcome == 'timeout':
            raise subprocess.TimeoutExpired('osascript', 1, output=b'partial', stderr=b'failure')
        if state.outcome == 'cancel':
            raise cancelled
        return SimpleNamespace(returncode=0 if state.outcome == 'closed' else 1,
                               stdout='WPSCOMPOSER_MS_OFFICE_OK:spreadsheet' if state.outcome == 'closed' else 'partial',
                               stderr='')

    monkeypatch.setattr(runtime.subprocess, 'run', run)
    write_text, write_bytes, open_fd = Path.write_text, Path.write_bytes, os.open

    def check(path):
        state.writes.append(path.name)
        if path.name in state.faults:
            raise OSError(errno.ENOSPC, 'diagnostic storage exhausted')

    def checked_text(path, *args, **kwargs):
        check(path)
        return write_text(path, *args, **kwargs)

    def checked_bytes(path, *args, **kwargs):
        check(path)
        return write_bytes(path, *args, **kwargs)

    def checked_open(path, *args, **kwargs):
        if Path(path).name == 'native-office.quarantine.json':
            check(Path(path))
        return open_fd(path, *args, **kwargs)

    monkeypatch.setattr(Path, 'write_text', checked_text)
    monkeypatch.setattr(Path, 'write_bytes', checked_bytes)
    monkeypatch.setattr(runtime, 'os', SimpleNamespace(**dict(vars(os), open=checked_open)))

    def prepare(job, deadline):
        staged = job / 'owned.xlsx'
        staged.write_bytes(b'native output')
        return 'synthetic script', staged

    def publish(staged, target, **kwargs):
        calls.append('publish')
        target.write_bytes(staged.read_bytes())

    monkeypatch.setattr(runtime, '_publish', publish)
    state.execute = lambda **kwargs: runtime._execute(
        'spreadsheet', output, overwrite=True, deadline=time.monotonic() + 3,
        prepare=kwargs.get('prepare', prepare))
    return state


@pytest.mark.parametrize('outcome', ['timeout', 'native', 'cancel'])
@pytest.mark.parametrize('fault', ['stdout.log', 'stderr.log', 'recovery.json', 'native-office.quarantine.json'])
def test_failed_native_job_preserves_primary_error_despite_diagnostic_io(job_host, outcome, fault):
    state = job_host
    state.outcome, state.faults = outcome, {fault}
    with pytest.raises(BaseException) as caught:
        state.execute()
    exc = caught.value
    if outcome == 'cancel':
        assert exc is state.cancelled
    else:
        assert isinstance(exc, runtime.NativeOfficeError)
        assert exc.code == ('NATIVE_OFFICE_TIMEOUT' if outcome == 'timeout' else 'NATIVE_OFFICE_EXECUTION_FAILED')
        assert Path(exc.staging_path).is_dir()
        if fault == 'recovery.json':
            assert exc.diagnostic_path is None
        else:
            assert json.loads(Path(exc.diagnostic_path).read_text())['code'] == exc.code
        if fault == 'native-office.quarantine.json':
            assert exc.quarantine_path is None
        else:
            assert Path(exc.quarantine_path).is_file()
    assert exc.diagnostic_io_failures
    assert state.calls == ['launch']
    assert state.output.read_bytes() == b'previous output'
    assert state.writes.index('native-office.quarantine.json') < state.writes.index('recovery.json')
    for name in ['stdout.log', 'stderr.log', 'recovery.json', 'native-office.quarantine.json']:
        assert name in state.writes
    if fault != 'native-office.quarantine.json':
        with pytest.raises(runtime.NativeOfficeError) as blocked:
            state.execute()
        assert blocked.value.code == 'NATIVE_OFFICE_QUARANTINED'
        assert state.calls == ['launch']


def test_all_diagnostic_writes_can_fail_without_masking_timeout(job_host):
    state = job_host
    state.outcome = 'timeout'
    state.faults = {'stdout.log', 'stderr.log', 'recovery.json', 'native-office.quarantine.json'}
    with pytest.raises(runtime.NativeOfficeError) as caught:
        state.execute()
    assert caught.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert caught.value.quarantine_path is None and caught.value.diagnostic_path is None
    assert len(caught.value.diagnostic_io_failures) == 4
    assert state.output.read_bytes() == b'previous output'


@pytest.mark.parametrize('fault', ['stdout.log', 'stderr.log'])
def test_closed_ack_with_failed_logs_preserves_closed_state_and_refuses_publication(job_host, fault):
    state = job_host
    state.outcome, state.faults = 'closed', {fault}
    with pytest.raises(OSError) as caught:
        state.execute()
    assert caught.value.errno == errno.ENOSPC
    assert state.calls == ['launch']
    assert state.output.read_bytes() == b'previous output'
    assert not (state.root / 'native-office.quarantine.json').exists()
    recovery = next(state.root.glob('job-*/recovery.json'))
    assert json.loads(recovery.read_text())['closed'] is True
    assert {'stdout.log', 'stderr.log'} <= set(state.writes)


def test_closed_native_job_still_publishes_and_removes_private_staging(job_host):
    state = job_host
    state.outcome = 'closed'
    assert state.execute() == state.output
    assert state.output.read_bytes() == b'native output'
    assert state.calls == ['launch', 'publish']
    assert not list(state.root.glob('job-*'))


def test_expired_deadline_before_transport_does_not_quarantine_or_report_deleted_paths(job_host, monkeypatch):
    state = job_host

    def expire_after_script(path, mode):
        monkeypatch.setattr(runtime, '_remaining', lambda deadline: (_ for _ in ()).throw(
            runtime.NativeOfficeError('NATIVE_OFFICE_TIMEOUT')))

    monkeypatch.setattr(runtime.os, 'chmod', expire_after_script)
    with pytest.raises(runtime.NativeOfficeError) as caught:
        state.execute()
    assert caught.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert state.calls == []
    assert not (state.root / 'native-office.quarantine.json').exists()
    assert not list(state.root.glob('job-*'))
    assert caught.value.staging_path is None and caught.value.diagnostic_path is None


@pytest.mark.parametrize('outcome', ['native', 'timeout', 'cancel'])
def test_lock_cleanup_failure_does_not_replace_primary_error(job_host, monkeypatch, outcome):
    state = job_host
    state.outcome = outcome
    close = runtime.OfficeJobLock.close

    def failing_close(lock):
        close(lock)
        raise OSError(errno.EIO, 'injected unlock failure')

    monkeypatch.setattr(runtime.OfficeJobLock, 'close', failing_close)
    with pytest.raises(BaseException) as caught:
        state.execute()
    if outcome == 'cancel':
        assert caught.value is state.cancelled
    else:
        assert isinstance(caught.value, runtime.NativeOfficeError)
        assert caught.value.code == ('NATIVE_OFFICE_TIMEOUT' if outcome == 'timeout'
                                     else 'NATIVE_OFFICE_EXECUTION_FAILED')
    assert caught.value.cleanup_io_failures == (('lock.close', 'OSError'),)
    assert (state.root / 'native-office.quarantine.json').is_file()
    assert state.calls == ['launch']
    assert state.output.read_bytes() == b'previous output'


def test_staging_cleanup_failure_preserves_prepare_error_identity(job_host, monkeypatch):
    primary = ValueError('original prepare failure')

    def failing_prepare(job, deadline):
        raise primary

    def failing_remove(path):
        raise OSError(errno.EACCES, 'injected staging cleanup failure')

    monkeypatch.setattr(runtime.shutil, 'rmtree', failing_remove)
    with pytest.raises(ValueError) as caught:
        job_host.execute(prepare=failing_prepare)
    assert caught.value is primary
    assert caught.value.cleanup_io_failures == (('staging.remove', 'PermissionError'),)
    assert job_host.calls == []
    assert not (job_host.root / 'native-office.quarantine.json').exists()
    assert len(list(job_host.root.glob('job-*'))) == 1


@pytest.mark.parametrize('phase', ['lock.close', 'staging.remove'])
def test_successful_job_still_propagates_cleanup_failure(job_host, monkeypatch, phase):
    job_host.outcome = 'closed'
    failure = OSError(errno.EIO, 'injected cleanup failure')
    close = runtime.OfficeJobLock.close

    def failing_close(lock):
        close(lock)
        raise failure

    def failing_remove(path):
        raise failure

    if phase == 'lock.close':
        monkeypatch.setattr(runtime.OfficeJobLock, 'close', failing_close)
    else:
        monkeypatch.setattr(runtime.shutil, 'rmtree', failing_remove)
    with pytest.raises(OSError) as caught:
        job_host.execute()
    assert caught.value is failure
    assert job_host.calls == ['launch', 'publish']
    assert job_host.output.read_bytes() == b'native output'
