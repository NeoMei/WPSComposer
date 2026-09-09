"""Abort persistence failures with real Python transports, without COM."""
from __future__ import annotations

import errno
from pathlib import Path
import queue
import subprocess
import sys

import pytest

from test_windows_session_proxy import transport


COMPONENTS = [('ProxyWordSession', '/source.docx', 'NATIVE_WORD_TIMEOUT'),
              ('ProxyExcelSession', '/source.xlsx', 'NATIVE_OFFICE_TIMEOUT'),
              ('ProxyPowerPointSession', '/source.pptx', 'NATIVE_OFFICE_TIMEOUT')]


@pytest.mark.parametrize('name,source,timeout_code', COMPONENTS)
@pytest.mark.parametrize('fault', ['recovery.json', 'native-office.quarantine.json'])
@pytest.mark.parametrize('outcome', ['timeout', 'cancel'])
def test_abort_preserves_primary_and_truthful_paths_across_components(
    transport, monkeypatch, name, source, timeout_code, fault, outcome,
):
    module, _, children = transport
    session = getattr(module, name).open_document(source)
    write_text = Path.write_text
    attempted = []
    targets = {session.staging_root / 'recovery.json', session._lock.quarantine_path}

    def write(path, *args, **kwargs):
        if path in targets:
            attempted.append(path.name)
            if path.name == fault:
                raise OSError(errno.ENOSPC, 'controlled private storage failure')
        return write_text(path, *args, **kwargs)

    monkeypatch.setattr(Path, 'write_text', write)
    primary = KeyboardInterrupt('original cancellation') if outcome == 'cancel' else queue.Empty()

    def receive(*args, **kwargs):
        raise primary

    monkeypatch.setattr(session._responses, 'get', receive)
    with pytest.raises(BaseException) as caught:
        session.inspect_document()
    error = caught.value
    if outcome == 'cancel':
        assert error is primary
    else:
        assert error.code == timeout_code
    assert error.diagnostic_io_failures == ((
        'recovery' if fault == 'recovery.json' else 'quarantine', 'OSError'),)
    assert attempted == ['native-office.quarantine.json', 'recovery.json']
    assert session._lock.closed and children[0].poll() is not None
    later_error = session._error('QUARANTINED')
    assert later_error.diagnostic_path == (None if fault == 'recovery.json' else str(session.staging_root / 'recovery.json'))
    assert later_error.quarantine_path == (None if fault == 'native-office.quarantine.json' else str(session._lock.quarantine_path))
    with pytest.raises((module.NativeWordError, module.NativeOfficeError)):
        session.inspect_document()
    if fault == 'recovery.json':
        with pytest.raises(RuntimeError, match='quarantined'):
            getattr(module, name).open_document(source)
        assert len(children) == 1


def test_initial_binding_error_retains_identity_when_abort_diagnostics_fail(transport, monkeypatch):
    module, _, children = transport
    primary = ValueError('original binding failure')
    captured = []

    def fail_binding(self, *args, **kwargs):
        captured.append(self)
        raise primary

    monkeypatch.setattr(module.ProxyExcelSession, '_call', fail_binding)
    write_text = Path.write_text

    def write(path, *args, **kwargs):
        if path.name == 'recovery.json':
            raise OSError(errno.ENOSPC, 'controlled recovery failure')
        return write_text(path, *args, **kwargs)

    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(ValueError) as caught:
        module.ProxyExcelSession.open_document('/source.xlsx')
    assert caught.value is primary
    assert primary.diagnostic_io_failures == (('recovery', 'OSError'),)
    assert captured[0]._lock.closed and children[0].poll() is not None
    assert captured[0]._lock.quarantine_path.is_file()


@pytest.mark.parametrize('phase', ['child.kill', 'io.stop', 'lock.close'])
def test_abort_cleanup_faults_preserve_cancellation_and_attempt_remaining_cleanup(transport, monkeypatch, phase):
    module, _, children = transport
    session = module.ProxyExcelSession.open_document('/source.xlsx')
    primary = KeyboardInterrupt('original cancellation')

    def receive(*args, **kwargs):
        raise primary

    monkeypatch.setattr(session._responses, 'get', receive)
    target, method = {'child.kill': (children[0], 'kill'), 'io.stop': (session, '_stop_io'),
                      'lock.close': (session._lock, 'close')}[phase]
    original = getattr(target, method)

    def cleanup():
        original()
        raise OSError(errno.EIO, 'controlled cleanup failure')

    monkeypatch.setattr(target, method, cleanup)
    with pytest.raises(BaseException) as caught:
        session.inspect_document()
    assert caught.value is primary
    assert primary.cleanup_io_failures == ((phase, 'OSError'),)
    assert session._lock.closed and children[0].poll() is not None
    assert session._stop.is_set()


def test_abort_only_terminates_its_exact_python_child(transport):
    module, _, children = transport
    session = module.ProxyPowerPointSession.open_document('/source.pptx')
    unrelated = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(30)'])
    children.append(unrelated)
    session._abort('controlled uncertain response')
    assert children[0].poll() is not None
    assert unrelated.poll() is None
    assert session._lock.closed


def test_error_does_not_advertise_recovery_before_abort(transport):
    module, _, _ = transport
    with module.ProxyWordSession.open_document('/source.docx') as session:
        error = session._error('EXECUTION_FAILED')
        assert error.diagnostic_path is None and error.quarantine_path is None


def test_cancel_during_close_is_not_replaced_with_quarantine_error(transport, monkeypatch):
    module, _, children = transport
    session = module.ProxyExcelSession.open_document('/source.xlsx')
    primary = SystemExit('original cancellation')

    def receive(*args, **kwargs):
        raise primary

    monkeypatch.setattr(session._responses, 'get', receive)
    with pytest.raises(SystemExit) as caught:
        session.close()
    assert caught.value is primary
    assert session._lock.quarantine_path.is_file()
    assert session._lock.closed and children[0].poll() is not None


@pytest.mark.parametrize('name,source,timeout_code', COMPONENTS)
@pytest.mark.parametrize('secondary_type', [KeyboardInterrupt, SystemExit])
@pytest.mark.parametrize('target', ['recovery', 'quarantine'])
def test_timeout_survives_secondary_cancellation_in_evidence_lookup(
    transport, monkeypatch, name, source, timeout_code, secondary_type, target,
):
    module, _, children = transport
    session = getattr(module, name).open_document(source)
    recovery = session.staging_root / 'recovery.json'
    quarantine = session._lock.quarantine_path
    failing_path = recovery if target == 'recovery' else quarantine
    original_is_file = Path.is_file
    secondary = secondary_type('secondary metadata cancellation')

    def interrupted(path):
        if path == failing_path:
            raise secondary
        return original_is_file(path)

    def receive(*args, **kwargs):
        raise queue.Empty()

    monkeypatch.setattr(Path, 'is_file', interrupted)
    monkeypatch.setattr(session._responses, 'get', receive)
    with pytest.raises(BaseException) as caught:
        session.inspect_document()
    error = caught.value
    assert isinstance(error, (module.NativeWordError, module.NativeOfficeError))
    assert error.code == timeout_code
    if name == 'ProxyWordSession':
        assert isinstance(error, module.NativeWordTimeoutError)
    assert error.diagnostic_path == (None if target == 'recovery' else str(recovery))
    assert error.quarantine_path == (None if target == 'quarantine' else str(quarantine))
    assert session._uncertain and session._lock.closed
    assert children[0].poll() is not None
    assert original_is_file(recovery) and original_is_file(quarantine)
    with pytest.raises((module.NativeWordError, module.NativeOfficeError)) as later:
        session.inspect_document()
    assert later.value.code == timeout_code.replace('TIMEOUT', 'QUARANTINED')
    with pytest.raises(RuntimeError, match='quarantined'):
        getattr(module, name).open_document(source)
    assert len(children) == 1
