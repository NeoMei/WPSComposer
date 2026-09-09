from pathlib import Path
from types import SimpleNamespace
import subprocess

import pytest

from skills.WPSComposer.scripts.msoffice import macos_office_runtime as runtime
from skills.WPSComposer.scripts.generation_plan import GenerationPlan, RecordedGeneration


@pytest.fixture
def native(monkeypatch, tmp_path):
    monkeypatch.setattr(runtime.sys, 'platform', 'darwin')
    monkeypatch.setattr(runtime, 'engine_executable', lambda *a: '/Applications/Microsoft Excel.app')
    monkeypatch.setattr(runtime, '_container_root', lambda component: tmp_path / component)
    compiler = SimpleNamespace(compile_plan=lambda *a, **kw: 'synthetic script',
                               compile_conversion=lambda *a, **kw: 'synthetic conversion')
    monkeypatch.setattr(runtime, '_compiler', lambda component: compiler)
    monkeypatch.setattr(runtime, '_publish', lambda staged, target, **kw: Path(target).write_bytes(staged.read_bytes()) or target)
    return compiler


def recording(component='spreadsheet'):
    return RecordedGeneration(GenerationPlan(component, ()), ())


def test_powerpoint_uses_reopenable_documents_container():
    assert runtime._container_root('presentation').parts[-3:] == ('Data', 'Documents', 'wpscomposer')
    assert runtime._container_root('spreadsheet').parts[-3:] == ('Data', 'tmp', 'wpscomposer')


def test_compile_failure_does_not_start_app_or_create_job(native, monkeypatch, tmp_path):
    def reject(*a, **kw):
        raise ValueError('unsupported plan')
    native.compile_plan = reject
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **k: pytest.fail('app launched'))
    with pytest.raises(ValueError, match='unsupported'):
        runtime.generate_recorded(recording(), tmp_path / 'out.xlsx')
    assert not (tmp_path / 'spreadsheet').exists()


def test_timeout_preserves_diagnostics_and_blocks_next_job(native, monkeypatch, tmp_path):
    calls = []
    def timeout(argv, **kw):
        calls.append(argv)
        raise subprocess.TimeoutExpired(argv, 1, output=b'private partial output')
    monkeypatch.setattr(runtime.subprocess, 'run', timeout)
    target = tmp_path / 'out.xlsx'
    target.write_bytes(b'original')
    with pytest.raises(runtime.NativeOfficeError) as caught:
        runtime.generate_recorded(recording(), target, overwrite=True, timeout=3)
    exc = caught.value
    assert exc.code == 'NATIVE_OFFICE_TIMEOUT'
    assert Path(exc.staging_path).is_dir()
    assert Path(exc.quarantine_path).is_file()
    assert 'private partial' not in str(exc)
    assert target.read_bytes() == b'original'
    with pytest.raises(runtime.NativeOfficeError, match='recovery'):
        runtime.generate_recorded(recording(), target, overwrite=True)
    assert len(calls) == 1


def test_success_requires_closed_marker_and_actual_output(native, monkeypatch, tmp_path):
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **kw: SimpleNamespace(
        returncode=0, stdout='WPSCOMPOSER_MS_OFFICE_OK:spreadsheet\n', stderr=''))
    with pytest.raises(runtime.NativeOfficeError) as caught:
        runtime.generate_recorded(recording(), tmp_path / 'out.xlsx')
    assert caught.value.code == 'NATIVE_OFFICE_EXECUTION_FAILED'
    assert not (tmp_path / 'out.xlsx').exists()


def test_zero_exit_without_marker_is_quarantined(native, monkeypatch, tmp_path):
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **kw: SimpleNamespace(
        returncode=0, stdout='OK', stderr=''))
    with pytest.raises(runtime.NativeOfficeError) as caught:
        runtime.generate_recorded(recording(), tmp_path / 'out.xlsx')
    assert Path(caught.value.quarantine_path).is_file()


def test_native_basename_is_unique_and_staging_is_component_private(native, monkeypatch, tmp_path):
    targets = []
    def compile_plan(plan, resources, native_path, **kw):
        targets.append(Path(native_path))
        return str(native_path)
    native.compile_plan = compile_plan
    def run(argv, **kw):
        target = Path(Path(argv[-1]).read_text())
        target.write_bytes(b'native output')
        return SimpleNamespace(returncode=0, stdout='WPSCOMPOSER_MS_OFFICE_OK:spreadsheet', stderr='')
    monkeypatch.setattr(runtime.subprocess, 'run', run)
    for index in range(2):
        runtime.generate_recorded(recording(), tmp_path / f'out{index}.xlsx')
    actual = [p for p in targets if str(p).startswith(str(tmp_path))]
    assert len(actual) == 2 and actual[0].name != actual[1].name
    assert all(p.parent.parent == tmp_path / 'spreadsheet' for p in actual)
    assert all(not p.parent.exists() for p in actual)


@pytest.mark.parametrize('timeout', [True, 0, -1, float('nan'), float('inf')])
def test_invalid_timeout_before_side_effects(native, monkeypatch, tmp_path, timeout):
    with pytest.raises((TypeError, ValueError)):
        runtime.generate_recorded(recording(), tmp_path / 'out.xlsx', timeout=timeout)
    assert not (tmp_path / 'spreadsheet').exists()


def test_recovery_preserves_quarantine_when_script_still_running(native, monkeypatch, tmp_path):
    import json
    root = tmp_path / 'spreadsheet'; root.mkdir()
    job = root / 'job-test'; job.mkdir()
    quarantine = root / 'native-office.quarantine.json'
    quarantine.write_text(json.dumps({'component': 'spreadsheet', 'staging_path': str(job)}))
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **kw: SimpleNamespace(returncode=0, stdout='/usr/bin/osascript '+str(job/'job.applescript'), stderr=''))
    with pytest.raises(RuntimeError, match='still running'):
        runtime.recover_quarantine('spreadsheet')
    assert quarantine.exists()


def test_recovery_checks_owned_documents_before_clearing(native, monkeypatch, tmp_path):
    import json
    root = tmp_path / 'presentation'; root.mkdir()
    job = root / 'session-test'; job.mkdir()
    quarantine = root / 'native-office.quarantine.json'
    quarantine.write_text(json.dumps({'component': 'presentation', 'staging_path': str(job)}))
    calls = []
    def run(argv, **kwargs):
        calls.append(argv)
        if argv[0] == '/bin/ps':
            return SimpleNamespace(returncode=0, stdout='/Applications/Microsoft PowerPoint.app/Contents/MacOS/Microsoft PowerPoint', stderr='')
        return SimpleNamespace(returncode=0, stdout='WPSCOMPOSER_RECOVERY_OK', stderr='')
    monkeypatch.setattr(runtime.subprocess, 'run', run)
    assert runtime.recover_quarantine('presentation')
    assert not quarantine.exists() and job.exists()
    script = calls[-1][-1]
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert apple_string(str(job) + '/') in script and 'full name' in script
    assert 'close ' not in script and 'quit' not in script.lower()


def test_recovery_rejects_foreign_staging_without_native_probe(native, monkeypatch, tmp_path):
    import json
    root = tmp_path / 'spreadsheet'; root.mkdir()
    quarantine = root / 'native-office.quarantine.json'
    quarantine.write_text(json.dumps({'component': 'spreadsheet', 'staging_path': str(tmp_path/'unrelated')}))
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **kw: pytest.fail('native probe'))
    with pytest.raises(ValueError):
        runtime.recover_quarantine('spreadsheet')
    assert quarantine.exists()


def test_recovery_does_not_clear_uncertain_attached_document(native, monkeypatch, tmp_path):
    import json
    root = tmp_path/'presentation'; root.mkdir()
    job = root/'session-test'; job.mkdir()
    quarantine = root/'native-office.quarantine.json'
    quarantine.write_text(json.dumps({'component':'presentation','staging_path':str(job),'bound_path':str(tmp_path/'user.pptx')}))
    monkeypatch.setattr(runtime.subprocess, 'run', lambda *a, **kw: pytest.fail('native probe'))
    with pytest.raises(RuntimeError, match='manual verification'):
        runtime.recover_quarantine('presentation')
    assert quarantine.exists()
