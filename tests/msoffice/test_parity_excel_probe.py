"""Contract checks for the bounded native Excel feasibility runner."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import subprocess

import pytest

SOURCE = Path(__file__).resolve().parents[2] / 'fixtures/microsoft_parity/macos_excel.py'


def probe_module():
    assert SOURCE.exists(), 'Excel probe runner has not been implemented'
    spec = importlib.util.spec_from_file_location('excel_parity_probe', SOURCE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_existing_output_is_rejected_before_execution(tmp_path):
    module = probe_module()
    marker = tmp_path / 'original'
    marker.write_text('keep')
    with pytest.raises(FileExistsError):
        module.run_probe(tmp_path, execute=lambda *a, **k: pytest.fail('must not launch'))
    assert marker.read_text() == 'keep'


def test_timeout_keeps_script_and_partial_diagnostics(tmp_path):
    module = probe_module()
    out = tmp_path / 'new'
    def timed_out(*args, **kwargs):
        assert 1 <= kwargs['timeout'] <= 90
        raise subprocess.TimeoutExpired(args[0], kwargs['timeout'], output=b'created owned workbook', stderr=b'waiting')
    report = module.run_probe(out, timeout=20, execute=timed_out)
    assert report['status'] == 'timeout'
    assert report['recovery_required'] is True
    assert (out / 'probe.applescript').is_file()
    assert (out / 'stdout.log').read_text() == 'created owned workbook'
    assert (out / 'stderr.log').read_text() == 'waiting'
    assert json.loads((out / 'report.json').read_text())['status'] == 'timeout'
    assert report['source_sha256']


def test_native_failure_is_not_reclassified_as_success(tmp_path):
    module = probe_module()
    def failed(*args, **kwargs):
        return subprocess.CompletedProcess(args[0], 1, 'stage=chart', 'native error -1708')
    report = module.run_probe(tmp_path / 'failed', execute=failed)
    assert report['status'] == 'failed'
    assert report['returncode'] == 1
    assert report['verified'] is False


def test_script_quotes_paths_and_closes_only_owned_documents(tmp_path):
    module = probe_module()
    script = module.compile_probe(tmp_path / 'quote"slash\\dir', 'abcdef012345', 30)
    assert 'quote\\"slash\\\\dir' in script
    assert 'quit' not in script.lower()
    assert 'display alerts' not in script.lower()
    assert 'do visual basic' not in script.lower()
    assert 'close ownedBook saving no' in script
    assert 'sentinelToken' in script
    assert 'full name of ownedBook' in script


@pytest.mark.parametrize('timeout', [0, -1, 901, True, 1.5])
def test_invalid_deadline_is_rejected_before_directory_creation(tmp_path, timeout):
    module = probe_module()
    out = tmp_path / 'new'
    with pytest.raises(ValueError):
        module.run_probe(out, timeout=timeout)
    assert not out.exists()


def test_success_token_without_native_artifacts_is_not_verified(tmp_path):
    module = probe_module()
    def incomplete(*args, **kwargs):
        return subprocess.CompletedProcess(args[0], 0, 'native_probe=pass', 'chart=pass')
    report = module.run_probe(tmp_path / 'incomplete', execute=incomplete)
    assert report['verified'] is False
    assert report['status'] == 'failed'
    assert 'FileNotFoundError' in report['error']


def test_records_optional_primitive_failure_and_exact_runner_source(tmp_path):
    module = probe_module()
    def failed(*args, **kwargs):
        return subprocess.CompletedProcess(args[0], 1, '', 'chart=pass\nsheet_create_rename_clone_delete=failed:-1728:missing object\n')
    out = tmp_path / 'partial'
    report = module.run_probe(out, execute=failed)
    assert report['primitives']['chart'] == 'pass'
    assert report['primitives']['sheet_create_rename_clone_delete'].startswith('failed:')
    assert (out / 'runner-source.py').read_bytes() == SOURCE.read_bytes()


def test_native_staging_copies_exact_output_and_keeps_failed_recovery(tmp_path):
    module = probe_module()
    staging_root = tmp_path / 'excel-container'
    staging_root.mkdir()
    native_bytes = b'invalid synthetic native output must never pass validation'
    def completed(*args, **kwargs):
        stages = list(staging_root.iterdir())
        assert len(stages) == 1
        (stages[0] / 'native.xlsx').write_bytes(native_bytes)
        return subprocess.CompletedProcess(args[0], 0, 'native_probe=pass', '')
    out = tmp_path / 'evidence'
    report = module.run_probe(out, execute=completed, native_staging_root=staging_root)
    assert report['verified'] is False
    assert (out / 'native.xlsx').read_bytes() == native_bytes
    native_dir = Path(report['native_directory'])
    assert native_dir.parent == staging_root
    assert (native_dir / 'native.xlsx').read_bytes() == native_bytes
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert apple_string(str(native_dir / 'native.xlsx')) in (out / 'probe.applescript').read_text()
