"""Safety/report contract tests; no Office launch in the portable suite."""
import importlib.util
import json
import subprocess
from pathlib import Path

import pytest

PATH = Path(__file__).resolve().parents[2] / 'fixtures/microsoft_parity/macos_powerpoint.py'


def probe():
    assert PATH.exists(), 'PowerPoint feasibility runner is missing'
    spec = importlib.util.spec_from_file_location('powerpoint_probe', PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_existing_output_rejected_without_altering_evidence(tmp_path):
    marker = tmp_path / 'report.json'
    marker.write_text('old evidence')
    with pytest.raises(FileExistsError):
        probe().run_probe(tmp_path)
    assert marker.read_text() == 'old evidence'


@pytest.mark.parametrize('timeout', [0, -1, True, float('inf'), float('nan')])
def test_invalid_budget_rejected_before_output_creation(tmp_path, timeout):
    output = tmp_path / 'new'
    with pytest.raises(ValueError):
        probe().run_probe(output, timeout=timeout)
    assert not output.exists()


def test_native_failure_retains_script_raw_logs_and_hashes(monkeypatch, tmp_path):
    module = probe()
    monkeypatch.setattr(module.sys, 'platform', 'darwin')
    monkeypatch.setattr(module, 'STAGING_PARENT', tmp_path / 'sandbox')
    def failed(*args, **kwargs):
        assert kwargs['timeout'] <= 12
        return subprocess.CompletedProcess(args[0], 1, 'partial output', 'native failure -1708')
    monkeypatch.setattr(module.subprocess, 'run', failed)
    output = tmp_path / 'failure'
    report = module.run_probe(output, timeout=12)
    assert report['status'] == 'FAIL'
    assert (output / 'native.stderr.txt').read_text() == 'native failure -1708'
    assert (output / 'native.applescript').exists()
    assert report['source_sha256'] == module.sha256(PATH)
    assert json.loads((output / 'report.json').read_text())['status'] == 'FAIL'
    assert report['checksums']['native.applescript'] == module.sha256(output / 'native.applescript')
    assert report['native_ui_acceptance'] == 'NOT_RUN'


def test_timeout_is_uncertain_and_retains_recovery_without_second_native_call(monkeypatch, tmp_path):
    module = probe()
    monkeypatch.setattr(module.sys, 'platform', 'darwin')
    monkeypatch.setattr(module, 'STAGING_PARENT', tmp_path / 'sandbox')
    calls = []
    def timeout(*args, **kwargs):
        calls.append(args)
        raise subprocess.TimeoutExpired(args[0], kwargs['timeout'], output=b'partial', stderr=b'pending')
    monkeypatch.setattr(module.subprocess, 'run', timeout)
    report = module.run_probe(tmp_path / 'timeout', timeout=1)
    assert len(calls) == 1
    assert report['status'] == 'UNCERTAIN_TIMEOUT'
    assert report['recovery']['may_have_open_owned_documents'] is True
    assert (tmp_path / 'timeout/native.stderr.txt').read_text() == 'pending'


def test_native_inputs_are_staged_inside_app_container(monkeypatch, tmp_path):
    module = probe()
    monkeypatch.setattr(module.sys, 'platform', 'darwin')
    monkeypatch.setattr(module, 'STAGING_PARENT', tmp_path / 'sandbox')
    sandbox = tmp_path / 'container'
    monkeypatch.setattr(module, 'STAGING_PARENT', sandbox, raising=False)
    def failed(args, **kwargs):
        source = Path(args[1]).read_text()
        assert str(sandbox) in source
        assert str(tmp_path / 'evidence/source.png') not in source
        return subprocess.CompletedProcess(args, 1, '', 'synthetic native failure')
    monkeypatch.setattr(module.subprocess, 'run', failed)
    report = module.run_probe(tmp_path / 'evidence')
    staging = Path(report['staging_dir'])
    assert staging.parent == sandbox
    assert (staging / 'source.png').read_bytes() == (tmp_path / 'evidence/source.png').read_bytes()
    assert staging.exists(), 'uncertain native artifacts must remain for recovery'


def test_missing_artifacts_cannot_pass_inspection(tmp_path):
    checks = probe().inspect_artifacts(tmp_path)['checks']
    assert checks['native_pptx_exists'] is False
    assert checks['native_pdf_exists'] is False


def test_saved_geometry_and_font_are_verified_not_only_text_presence(tmp_path):
    import zipfile
    p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    with zipfile.ZipFile(tmp_path/'native.pptx', 'w') as z:
        z.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{p}"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>')
        z.writestr('ppt/slides/slide1.xml', f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:sp><p:nvSpPr><p:cNvPr name="parity-title"/></p:nvSpPr><p:txBody><a:p><a:r><a:rPr sz="2100"/><a:t>Native editable PowerPoint</a:t></a:r></a:p></p:txBody></p:sp></p:sld>')
        z.writestr('ppt/slides/slide2.xml', f'<p:sld xmlns:p="{p}"/>')
    checks = probe().inspect_artifacts(tmp_path)['checks']
    assert checks['editable_text'] is True
    assert checks['title_font_28pt_bold'] is False
    assert checks['shape_geometry_and_fill'] is False
