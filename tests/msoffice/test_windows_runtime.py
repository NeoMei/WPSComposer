import base64
import json
import subprocess
import sys
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.msoffice import windows_runtime as runtime


def test_worker_deadline_terminates_only_child_and_retains_private_evidence(monkeypatch, tmp_path):
    real_popen = subprocess.Popen
    children = []

    def stalled_worker(command, **kwargs):
        child = real_popen([sys.executable, '-c', 'import time; time.sleep(30)'], **kwargs)
        children.append(child)
        return child

    monkeypatch.setattr(runtime.subprocess, 'Popen', stalled_worker)
    started = time.monotonic()
    with pytest.raises(TimeoutError, match='retained'):
        runtime._run_worker(tmp_path, {'action': 'export'}, started + 0.2)
    assert time.monotonic() - started < 3
    assert children[0].poll() is not None
    assert list(tmp_path.glob('*/request.json'))
    assert list(tmp_path.glob('*/diagnostics.json'))


@pytest.mark.parametrize('deadline', [float('nan'), float('inf'), -1, 0])
def test_invalid_deadline_does_not_start_child(monkeypatch, tmp_path, deadline):
    def forbidden(*args, **kwargs):
        pytest.fail('Child launched for invalid deadline')
    monkeypatch.setattr(runtime.subprocess, 'Popen', forbidden)
    with pytest.raises((ValueError, TimeoutError)):
        runtime._run_worker(tmp_path, {'action': 'export'}, deadline)
    assert not list(tmp_path.iterdir())


def test_unknown_worker_action_returns_error_and_diagnostics(tmp_path):
    with pytest.raises(RuntimeError, match='retained'):
        runtime._run_worker(tmp_path, {'action': 'not-an-operation'}, time.monotonic() + 10)
    diagnostics = json.loads(next(tmp_path.glob('*/diagnostics.json')).read_text())
    assert diagnostics['status'] == 'failed'
    assert diagnostics['error_type'] == 'ValueError'


def test_generation_transports_validated_plan_and_normalized_bytes(monkeypatch):
    build = build_longform_generation('# Native report\n\n## Chapter\n\nBody.')
    captured = []
    adapter = runtime.WindowsWordAdapter(build)

    def execute(root, payload, deadline):
        captured.append(payload)
        target = root / 'native.docx'
        target.write_bytes(b'owned')
        return {'outcome': {'stagedArtifact': str(target), 'paginationMap': {'version': 'M5-v1', 'nodes': []}, 'appliedOperations': len(payload['plan']['operations'])}}

    monkeypatch.setattr(runtime, '_run_worker', execute)
    try:
        outcome = adapter.execute(build, (), time.monotonic() + 10)
        assert outcome.pagination_map.version == 'M5-v1'
        assert captured[0]['action'] == 'execute'
        assert captured[0]['plan']['protocolVersion'] == 2
        assert captured[0]['resources'] == []
    finally:
        adapter.close()


def test_adapter_rejects_stub_pagination_and_retains_stage(monkeypatch):
    build = build_longform_generation('# Report\n\nBody.')
    adapter = runtime.WindowsWordAdapter(build)
    marker = adapter.staging_root / 'evidence.docx'
    marker.write_bytes(b'evidence')
    monkeypatch.setattr(runtime, '_run_worker', lambda *args: {'outcome': {'stagedArtifact': str(marker), 'paginationMap': {'version': 'M2-stub'}}})
    with pytest.raises(RuntimeError, match='M5'):
        adapter.execute(build, (), time.monotonic() + 10)
    adapter.cleanup(marker)
    adapter.close()
    assert marker.read_bytes() == b'evidence'
    import shutil
    shutil.rmtree(adapter.staging_root)


def test_convert_stages_source_before_native_open_and_publishes_atomically(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.conversion import ConversionRequest
    source = tmp_path / 'original.docx'
    source.write_bytes(b'original document')
    output = tmp_path / 'public.pdf'
    seen = []

    def export(root, payload, deadline):
        staged = Path(payload['source'])
        assert staged != source
        assert staged.read_bytes() == b'original document'
        staged.write_bytes(b'Word modified only the copy')
        result = root / 'export.pdf'
        result.write_bytes(b'pdf')
        seen.append(root)
        return {'path': str(result)}

    def publish(staged, target, **kwargs):
        assert not target.exists()
        target.write_bytes(staged.read_bytes())
        return target

    monkeypatch.setattr(runtime, '_run_worker', export)
    monkeypatch.setattr(runtime, 'publish_artifact', publish)
    result = runtime.convert(ConversionRequest(source, output, 'writer', False))
    assert result == output
    assert source.read_bytes() == b'original document'
    assert output.read_bytes() == b'pdf'
    assert not seen[0].exists()


def test_convert_rejects_existing_output_before_worker(monkeypatch, tmp_path):
    source = tmp_path / 'source.docx'
    source.write_bytes(b'docx')
    output = tmp_path / 'result.pdf'
    output.write_bytes(b'existing')
    monkeypatch.setattr(runtime, '_run_worker', lambda *args: pytest.fail('native worker started'))
    with pytest.raises(FileExistsError):
        runtime.convert(SimpleNamespace(source=source, output=output, component='writer', overwrite=False))
    assert output.read_bytes() == b'existing'


def test_normalized_image_resource_survives_json_transport_without_source_path():
    from skills.WPSComposer.scripts.longform.resources import ImageProfile, PreparedLongformResource
    resource = PreparedLongformResource(
        id='res-synthetic', media_type='image/png', source_sha256='a' * 64,
        payload_sha256='b' * 64, normalizer_id='raster-v1',
        payload_bytes=b'\x00\xff\xfe image payload',
        image_profile=ImageProfile(80, 60, 96.0, 96.0, 1, 'PNG', False),
    )
    wire = json.loads(json.dumps(runtime._resource_json(resource)))
    assert wire['payload_base64'] == 'AP/+IGltYWdlIHBheWxvYWQ='
    assert 'source_path' not in wire
    recovered = runtime._resource_from_json(wire)
    assert recovered.payload_bytes == b'\x00\xff\xfe image payload'
    assert recovered.image_profile.pixel_width == 80


def test_worker_returned_escape_path_is_rejected_and_evidence_retained(monkeypatch, tmp_path):
    adapter = runtime.WindowsWordAdapter(build_longform_generation('# Report\n\nText.'))
    source = adapter.staging_root / 'source.docx'
    source.write_bytes(b'owned')
    outside = tmp_path / 'unrelated.pdf'
    outside.write_bytes(b'user pdf')
    monkeypatch.setattr(runtime, '_run_worker', lambda *args: {'path': str(outside)})
    with pytest.raises(RuntimeError, match='outside'):
        adapter.export_pdf(source, time.monotonic() + 10)
    adapter.cleanup(source)
    adapter.close()
    assert source.exists()
    assert outside.read_bytes() == b'user pdf'
    import shutil
    shutil.rmtree(adapter.staging_root)


def test_unpublished_lifecycle_retains_evidence_even_after_native_success():
    adapter = runtime.WindowsWordAdapter(build_longform_generation('# Report\n\nText.'))
    marker = adapter.staging_root / 'native-result.docx'
    marker.write_bytes(b'diagnostic evidence')
    adapter.cleanup(marker)
    adapter.close()
    assert marker.read_bytes() == b'diagnostic evidence'
    import shutil
    shutil.rmtree(adapter.staging_root)


def test_successful_publication_allows_private_stage_cleanup(monkeypatch, tmp_path):
    adapter = runtime.WindowsWordAdapter(build_longform_generation('# Report\n\nText.'))
    staged = adapter.staging_root / 'native-result.docx'
    staged.write_bytes(b'owned document')
    output = tmp_path / 'result.docx'
    monkeypatch.setattr(runtime._BaseAdapter, 'publish', lambda *args: output)
    assert adapter.publish(staged, output, False, time.monotonic() + 10) == output
    adapter.cleanup(staged)
    adapter.close()
    assert not adapter.staging_root.exists()


def test_child_launch_failure_preserves_diagnostic(monkeypatch, tmp_path):
    def fail_launch(*args, **kwargs):
        raise OSError('Cannot launch Python')
    monkeypatch.setattr(runtime.subprocess, 'Popen', fail_launch)
    with pytest.raises(RuntimeError) as caught:
        runtime._run_worker(tmp_path, {'action': 'export'}, time.monotonic() + 10)
    assert caught.value.code == 'NATIVE_WORD_EXECUTION_FAILED'
    diagnostic = json.loads(next(tmp_path.glob('*/diagnostics.json')).read_text())
    assert diagnostic['error_type'] == 'OSError'


@pytest.mark.parametrize('signal', [KeyboardInterrupt(), SystemExit(7)])
def test_child_launch_cancellation_preserves_signal_and_diagnostic(monkeypatch, tmp_path, signal):
    def cancel_launch(*args, **kwargs):
        raise signal
    monkeypatch.setattr(runtime.subprocess, 'Popen', cancel_launch)
    with pytest.raises(type(signal)) as caught:
        runtime._run_worker(tmp_path, {'action': 'export'}, time.monotonic() + 10)
    assert caught.value is signal
    diagnostic = json.loads(next(tmp_path.glob('*/diagnostics.json')).read_text())
    assert diagnostic['error_type'] == type(signal).__name__
    assert diagnostic['word_termination_attempted'] is False
