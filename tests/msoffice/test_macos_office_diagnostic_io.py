"""Exercise real session transports with process and diagnostic I/O faults."""
from pathlib import Path
import subprocess
import time

import pytest

from skills.WPSComposer.scripts.msoffice import macos_excel_session as excel
from skills.WPSComposer.scripts.msoffice import macos_powerpoint_session as powerpoint


class Lock:
    def __init__(self, root, fail=False):
        self.quarantine_path = root / 'quarantine.json'
        self.fail = fail
        self.closed = False

    def quarantine(self, detail):
        if self.fail:
            raise OSError('injected quarantine file write failure')
        self.quarantine_path.write_text('quarantined')

    def close(self):
        self.closed = True


def session(app, tmp_path, monkeypatch, execute, fault):
    lock = Lock(tmp_path, fail=fault in ('lock', 'all'))
    if app == 'excel':
        obj = excel.MacExcelSession.__new__(excel.MacExcelSession)
        obj._closed = obj._failed = obj._attached = obj._read_only = False
        obj._job = tmp_path
        obj._native = tmp_path / 'owned.xlsx'
        obj._counter = 0
        obj._deadline = time.monotonic() + 60
        obj._lock = lock
        obj._execute = execute
    else:
        obj = powerpoint.MacPowerPointSession()
        obj._job = tmp_path
        obj._path = str(tmp_path / 'owned.pptx')
        obj._name = 'owned.pptx'
        obj._entered = True
        obj._deadline = time.monotonic() + 60
        obj._lock = lock
        monkeypatch.setattr(powerpoint.subprocess, 'run', execute)
    real_write = Path.write_text
    real_bytes = Path.write_bytes

    def reject(path):
        is_log = path.suffix in ('.log', '.stdout', '.stderr')
        if (fault in ('logs', 'all') and is_log
                or fault in ('recovery', 'all') and path.name == 'recovery.json'
                or fault == 'script' and path.suffix == '.applescript'):
            raise OSError('injected diagnostic I/O failure')

    def write(path, data, *args, **kwargs):
        reject(path)
        return real_write(path, data, *args, **kwargs)

    def write_bytes(path, data):
        reject(path)
        return real_bytes(path, data)

    monkeypatch.setattr(Path, 'write_text', write)
    monkeypatch.setattr(Path, 'write_bytes', write_bytes)
    return obj


def quarantined(app, obj):
    return obj._failed if app == 'excel' else obj._uncertain


@pytest.mark.parametrize('app', ['excel', 'powerpoint'])
def test_success_log_failure_blocks_further_native_and_close(app, tmp_path, monkeypatch):
    calls = []

    def execute(command, **kwargs):
        calls.append(command)
        payload = '{}' if app == 'excel' else '["WPSCOMPOSER_PPT_SESSION_OK", "OK"]'
        return subprocess.CompletedProcess(command, 0, payload, '')

    obj = session(app, tmp_path, monkeypatch, execute, 'logs')
    with pytest.raises(Exception):
        obj._run('return "{}"')
    assert quarantined(app, obj)
    with pytest.raises(RuntimeError):
        obj._run('return "{}"')
    obj.close()
    assert len(calls) == 1


@pytest.mark.parametrize('app', ['excel', 'powerpoint'])
@pytest.mark.parametrize('primary', ['native', 'timeout', 'cancel'])
@pytest.mark.parametrize('fault', ['logs', 'recovery', 'lock', 'all'])
def test_diagnostic_failure_preserves_primary_and_quarantine(app, primary, fault, tmp_path, monkeypatch):
    calls = []
    cancellation = KeyboardInterrupt('injected primary cancellation')

    def execute(command, **kwargs):
        calls.append(command)
        if primary == 'cancel':
            raise cancellation
        if primary == 'timeout':
            raise subprocess.TimeoutExpired(command, 1, output=b'partial', stderr=b'timeout')
        return subprocess.CompletedProcess(command, 1, '', 'execution error: native (-2700)')

    obj = session(app, tmp_path, monkeypatch, execute, fault)
    expected = KeyboardInterrupt if primary == 'cancel' else RuntimeError
    with pytest.raises(expected) as caught:
        obj._run('return "{}"', **({'mutation': True} if app == 'powerpoint' else {}))
    if primary == 'cancel':
        assert caught.value is cancellation
    elif app == 'powerpoint':
        assert caught.value.code == ('NATIVE_OFFICE_TIMEOUT' if primary == 'timeout' else 'NATIVE_OFFICE_EXECUTION_FAILED')
    if fault == 'all' and primary != 'cancel':
        assert getattr(caught.value, 'diagnostic_path', None) is None
    assert quarantined(app, obj)
    with pytest.raises(RuntimeError):
        obj._run('return "{}"')
    obj.close()
    assert len(calls) == 1


@pytest.mark.parametrize('app', ['excel', 'powerpoint'])
def test_script_write_failure_stays_before_submission(app, tmp_path, monkeypatch):
    calls = []
    obj = session(app, tmp_path, monkeypatch, lambda *a, **kw: calls.append(a), 'script')
    with pytest.raises(OSError, match='diagnostic I/O'):
        obj._run('return "{}"')
    assert calls == []
    assert not quarantined(app, obj)
