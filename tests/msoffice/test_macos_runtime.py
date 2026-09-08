from pathlib import Path
import time
import pytest


def test_invalid_timeout_does_not_start_word(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_runtime import remaining
    for deadline in (time.monotonic()-1, float('nan'), float('inf')):
        with pytest.raises((ValueError, TimeoutError)):
            remaining(deadline)


def test_cross_process_lock_respects_total_deadline(tmp_path):
    pytest.importorskip("fcntl", reason="Native macOS/POSIX job locking is unavailable on Windows")
    from skills.WPSComposer.scripts.msoffice.macos_runtime import WordJobLock
    path = tmp_path/'word.lock'
    one, two = WordJobLock(path), WordJobLock(path)
    one.acquire(time.monotonic()+1)
    try:
        with pytest.raises(TimeoutError):
            two.acquire(time.monotonic()+0.05)
    finally:
        one.close()
        two.close()
    two.acquire(time.monotonic()+1)
    two.close()


def test_capability_rejection_precedes_staging(tmp_path, monkeypatch):
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
    from skills.WPSComposer.scripts.msoffice.macos_script import MacWordCapabilityError
    build=build_longform_generation('# Title\n\n$$\nx^2\n$$')
    adapter=MacWordAdapter(build)
    def forbidden(*args, **kwargs):
        pytest.fail('unsupported plan must not allocate native staging')
    monkeypatch.setattr(adapter, '_ensure_started', forbidden)
    with pytest.raises(MacWordCapabilityError):
        adapter.execute(build, (), time.monotonic()+30)
    adapter.close()


def test_cleanup_keeps_quarantined_staging(tmp_path):
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
    adapter=MacWordAdapter(build_longform_generation(''))
    adapter.staging_root=tmp_path
    source=tmp_path/'owned.docx'; source.write_bytes(b'uncertain')
    adapter.quarantined=True
    adapter.cleanup(source)
    adapter.close()
    assert source.read_bytes() == b'uncertain'


def test_failed_unpublished_job_preserves_staging(tmp_path):
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
    adapter=MacWordAdapter(build_longform_generation(''))
    adapter.staging_root=tmp_path
    p=tmp_path/'operation.log';p.write_text('diagnostic')
    adapter.close()
    assert p.is_file()


def test_persistent_quarantine_blocks_next_job(tmp_path):
    pytest.importorskip("fcntl", reason="Native macOS/POSIX job locking is unavailable on Windows")
    from skills.WPSComposer.scripts.msoffice.macos_runtime import WordJobLock
    lock=WordJobLock(tmp_path/'native.lock')
    lock.quarantine('uncertain AppleEvent; diagnostic.log')
    with pytest.raises(RuntimeError,match='quarantin'):
        lock.acquire(time.monotonic()+1)
    assert (tmp_path/'native.lock.quarantine').is_file()


def test_failed_native_cleanup_persists_job_quarantine(tmp_path, monkeypatch):
    import subprocess
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter, WordJobLock
    adapter=MacWordAdapter(build_longform_generation(''))
    adapter.staging_root=tmp_path
    adapter.lock=WordJobLock(tmp_path/'word.lock')
    monkeypatch.setattr(adapter,'_ensure_started',lambda deadline: None)
    monkeypatch.setattr(subprocess,'run',lambda *a,**kw: subprocess.CompletedProcess(a[0],1,'','cleanup failed'))
    with pytest.raises(RuntimeError):
        adapter._run('native source',time.monotonic()+30)
    assert adapter.lock.quarantine_path.is_file()
    adapter.close()


def test_timeout_preserves_partial_diagnostics_and_blocks_next_job(tmp_path, monkeypatch):
    pytest.importorskip("fcntl", reason="Native macOS/POSIX job locking is unavailable on Windows")
    import subprocess
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter, WordJobLock
    adapter=MacWordAdapter(build_longform_generation(''))
    adapter.staging_root=tmp_path;adapter.lock=WordJobLock(tmp_path/'word.lock')
    monkeypatch.setattr(adapter,'_ensure_started',lambda deadline: None)
    def timeout(*args,**kwargs):
        raise subprocess.TimeoutExpired(args[0],2,output=b'partial output',stderr=b'partial diagnostic')
    monkeypatch.setattr(subprocess,'run',timeout)
    with pytest.raises(TimeoutError) as caught:
        adapter._run('native source',time.monotonic()+30)
    assert caught.value.code == 'NATIVE_WORD_TIMEOUT'
    assert next(tmp_path.glob('*.log')).read_text() == 'partial diagnostic\npartial output'
    assert adapter.lock.quarantine_path.is_file()
    with pytest.raises(RuntimeError,match='quarantin'):
        WordJobLock(tmp_path/'word.lock').acquire(time.monotonic()+1)


def test_recovery_without_marker_returns_without_launch(tmp_path, monkeypatch):
    pytest.importorskip("fcntl", reason="Native macOS/POSIX job locking is unavailable on Windows")
    import subprocess
    from skills.WPSComposer.scripts.msoffice.macos_runtime import recover_quarantine
    monkeypatch.setattr(subprocess,'run',lambda *a,**kw: pytest.fail('unnecessary native launch'))
    assert recover_quarantine(lock_path=tmp_path/'word.lock') is False
