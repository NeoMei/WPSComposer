"""Public failures retain recovery data without exposing native document text."""
import os
import subprocess
from pathlib import Path

import pytest

from skills.WPSComposer import generate, convert_to_pdf
from skills.WPSComposer.scripts import conversion
from skills.WPSComposer.scripts.msoffice.errors import NativeWordError, NativeWordTimeoutError
from skills.WPSComposer.scripts.longform.lifecycle import LongformLifecycleError
from skills.WPSComposer.scripts.longform import platform_runtime
from skills.WPSComposer.scripts.msoffice import macos_runtime, windows_runtime


PRIVATE_TEXT = 'PRIVATE DOCUMENT CONTENT MUST NOT ESCAPE'


def test_public_mac_unsupported_plan_names_capability_without_starting_word(monkeypatch, tmp_path):
    monkeypatch.setattr(platform_runtime.sys, 'platform', 'darwin')
    monkeypatch.setattr(macos_runtime.MacWordAdapter, '_ensure_started',
                        lambda *args: pytest.fail('unsupported plan started Word'))
    with pytest.raises(LongformLifecycleError) as caught:
        generate('# Title\n\n$$\n\\unsupported{PRIVATE DOCUMENT CONTENT MUST NOT ESCAPE}\n$$', source_is_text=True,
                 output=str(tmp_path/'out.docx'), engine='msoffice')
    assert caught.value.code == 'NATIVE_WORD_UNSUPPORTED'
    assert 'equation' in str(caught.value).lower()
    assert 'WPS' not in str(caught.value)
    assert PRIVATE_TEXT not in str(caught.value)
    assert not (tmp_path/'out.docx').exists()


@pytest.mark.parametrize('failure', ['timeout', 'cleanup'])
@pytest.mark.skipif(os.name != 'posix', reason='This macOS recovery test acquires a real POSIX job lock')
def test_public_mac_failure_retains_safe_recovery_locations(monkeypatch, tmp_path, failure):
    monkeypatch.setattr(platform_runtime.sys, 'platform', 'darwin')
    stage = tmp_path/'private-stage'
    stage.mkdir()
    def start(adapter, deadline):
        if adapter.staging_root is None:
            adapter.staging_root = stage
            adapter.lock = macos_runtime.WordJobLock(tmp_path/'native.lock')
            adapter.lock.acquire(deadline)
    monkeypatch.setattr(macos_runtime.MacWordAdapter, '_ensure_started', start)
    def native(*args, **kwargs):
        if failure == 'timeout':
            raise subprocess.TimeoutExpired(args[0], 1, output=PRIVATE_TEXT)
        return subprocess.CompletedProcess(args[0], 1, '', PRIVATE_TEXT)
    monkeypatch.setattr(macos_runtime.subprocess, 'run', native)
    with pytest.raises(LongformLifecycleError) as caught:
        generate('# Title\n\n## Chapter\n\nBody', source_is_text=True,
                 output=str(tmp_path/'out.docx'), engine='msoffice')
    assert caught.value.code == ('NATIVE_WORD_TIMEOUT' if failure == 'timeout' else 'NATIVE_WORD_QUARANTINED')
    assert caught.value.staging_path == str(stage)
    assert Path(caught.value.diagnostic_path).is_file()
    assert Path(caught.value.quarantine_path).is_file()
    assert str(stage) in str(caught.value)
    assert PRIVATE_TEXT not in str(caught.value)
    assert not (tmp_path/'out.docx').exists()


def test_public_windows_worker_failure_retains_evidence_without_raw_traceback(monkeypatch, tmp_path, word_lock):
    monkeypatch.setattr(platform_runtime.sys, 'platform', 'win32')
    stage = tmp_path/'windows-stage'
    stage.mkdir()
    monkeypatch.setattr(windows_runtime, '_private_root', lambda: stage)
    class FailedWorker:
        returncode = 1
        def wait(self, **kwargs):
            return 1
    monkeypatch.setattr(windows_runtime.subprocess, 'Popen', lambda *args, **kwargs: FailedWorker())
    with pytest.raises(LongformLifecycleError) as caught:
        generate('# Title\n\n## Chapter\n\nBody', source_is_text=True,
                 output=str(tmp_path/'out.docx'), engine='msoffice')
    assert caught.value.code == 'NATIVE_WORD_EXECUTION_FAILED'
    assert Path(caught.value.staging_path).parent == stage
    assert caught.value.staging_path in str(caught.value)


def test_public_wps_unstructured_errors_keep_existing_redaction(monkeypatch, tmp_path):
    monkeypatch.setattr(platform_runtime.sys, 'platform', 'darwin')
    def fail(*args, **kwargs):
        raise RuntimeError(PRIVATE_TEXT)
    monkeypatch.setattr(platform_runtime.MacLongformAdapter, 'execute', fail)
    with pytest.raises(LongformLifecycleError) as caught:
        generate('# Title\n\nBody', source_is_text=True,
                 output=str(tmp_path/'out.docx'), engine='wps')
    assert caught.value.code == 'NATIVE_GENERATION_FAILED'
    assert PRIVATE_TEXT not in str(caught.value)


@pytest.mark.parametrize('code', ['NATIVE_WORD_TIMEOUT', 'NATIVE_WORD_QUARANTINED'])
def test_conversion_keeps_exception_class_code_and_structured_recovery(monkeypatch, tmp_path, code):
    source = tmp_path/'source.docx'
    source.write_bytes(b'backend replaced; no native input is opened')
    stage = tmp_path/'private-stage'
    def fail(request):
        raise NativeWordError(code, staging_path=stage, diagnostic_path=stage/'native.log')
    monkeypatch.setattr(conversion, '_select_backend', lambda request: ('test-word', fail))
    with pytest.raises(conversion.ConversionError) as caught:
        convert_to_pdf(str(source), engine='msoffice')
    assert caught.value.code == code
    assert code in conversion.STABLE_CONVERSION_ERROR_CODES
    assert caught.value.staging_path == str(stage)
    assert caught.value.to_dict()['staging_path'] == str(stage)


def test_clean_native_timeout_keeps_diagnostics_without_quarantine(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    adapter = macos_runtime.MacWordAdapter(build_longform_generation(''))
    adapter.staging_root = tmp_path
    adapter.lock = macos_runtime.WordJobLock(tmp_path/'native.lock')
    monkeypatch.setattr(adapter, '_ensure_started', lambda deadline: None)
    monkeypatch.setattr(macos_runtime.subprocess, 'run', lambda *args, **kwargs:
                        subprocess.CompletedProcess(args[0], 1, '',
                            'WPSC_ERROR\t-1712\t' + PRIVATE_TEXT + '\nWPSC_CLEAN\n'))
    with pytest.raises(NativeWordTimeoutError) as caught:
        adapter._run('native source', macos_runtime.time.monotonic()+30)
    assert caught.value.staging_path == str(tmp_path)
    assert caught.value.quarantine_path is None
    assert PRIVATE_TEXT not in str(caught.value)


def test_deadline_crossed_after_native_completion_keeps_stage(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    adapter = macos_runtime.MacWordAdapter(build_longform_generation(''))
    adapter.staging_root = tmp_path
    monkeypatch.setattr(adapter, '_ensure_started', lambda deadline: None)
    checks = iter([10, None])
    def remaining(deadline):
        value = next(checks)
        if value is None:
            raise NativeWordTimeoutError()
        return value
    monkeypatch.setattr(macos_runtime, 'remaining', remaining)
    monkeypatch.setattr(macos_runtime.subprocess, 'run', lambda *args, **kwargs:
                        subprocess.CompletedProcess(args[0], 0, 'WPSC_OK\t0\n', ''))
    with pytest.raises(NativeWordTimeoutError) as caught:
        adapter._run('native source', 99)
    assert caught.value.staging_path == str(tmp_path)
    assert Path(caught.value.diagnostic_path).is_file()


def test_windows_completion_deadline_keeps_operation_evidence(monkeypatch, tmp_path, word_lock):
    checks = iter([True, True, True, False])
    def deadline(value):
        if not next(checks):
            raise NativeWordTimeoutError()
    monkeypatch.setattr(windows_runtime, '_require_deadline', deadline)
    class FinishedWorker:
        returncode = 0
        def wait(self, **kwargs):
            return 0
    def start(command, **kwargs):
        Path(command[-1]).write_text('{"status":"ok","value":{}}')
        return FinishedWorker()
    monkeypatch.setattr(windows_runtime.subprocess, 'Popen', start)
    with pytest.raises(NativeWordTimeoutError) as caught:
        windows_runtime._run_worker(tmp_path, {'action':'export'}, 99)
    assert Path(caught.value.staging_path).parent == tmp_path
    assert Path(caught.value.diagnostic_path).is_file()


@pytest.fixture
def word_lock(monkeypatch, tmp_path):
    # These tests simulate Windows workers on all hosts. Keep the shared lock
    # protocol without importing the host-specific msvcrt module or user roots.
    class Lock:
        def __init__(self, root):
            self.quarantine_path = root / 'native-office.quarantine.json'
        def acquire(self, deadline):
            pass
        def quarantine(self, detail):
            import json
            self.quarantine_path.write_text(json.dumps(detail))
        def close(self):
            pass
    monkeypatch.setattr(windows_runtime, 'OfficeJobLock', Lock)
    monkeypatch.setattr(windows_runtime, '_component_root', lambda component: tmp_path/'shared-word')
