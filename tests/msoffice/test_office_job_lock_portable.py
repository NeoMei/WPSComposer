"""Exercise native exclusion and the Windows CRT boundary on every host."""
from __future__ import annotations

import errno
import json
import os
import sys
import time
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice import macos_office_runtime as runtime


@pytest.fixture
def windows_backend(monkeypatch):
    calls, failures, streams = [], [], []

    def locking(fd, mode, count):
        calls.append((mode, count, os.lseek(fd, 0, os.SEEK_CUR), os.fstat(fd).st_size))
        if failures:
            raise failures.pop(0)

    def fdopen(*args, **kwargs):
        stream = os.fdopen(*args, **kwargs)
        streams.append(stream)
        return stream

    monkeypatch.setattr(runtime, 'os', SimpleNamespace(**dict(vars(os), name='nt', fdopen=fdopen)))
    monkeypatch.setitem(sys.modules, 'msvcrt', SimpleNamespace(locking=locking, LK_NBLCK=2, LK_UNLCK=0))
    monkeypatch.setitem(sys.modules, 'fcntl', None)
    return calls, failures, streams


def test_native_lock_excludes_contender_until_close(tmp_path):
    owner, contender = runtime.OfficeJobLock(tmp_path), runtime.OfficeJobLock(tmp_path)
    owner.acquire(time.monotonic() + 3)
    try:
        with pytest.raises(runtime.NativeOfficeError) as caught:
            contender.acquire(time.monotonic() + .1)
        assert caught.value.code == 'NATIVE_OFFICE_TIMEOUT'
        assert contender.stream is None
    finally:
        owner.close()
    contender.acquire(time.monotonic() + 3)
    contender.close()


def test_windows_lock_releases_same_byte_without_initialization_write(tmp_path, windows_backend):
    calls, _, streams = windows_backend
    lock = runtime.OfficeJobLock(tmp_path)
    lock.acquire(time.monotonic() + 3)
    lock.close()
    lock.close()
    assert calls == [(2, 1, 0, 0), (0, 1, 0, 0)]
    assert streams[0].closed and lock.stream is None


@pytest.mark.parametrize('code', [errno.EACCES, errno.EAGAIN, errno.EDEADLK])
def test_windows_contention_retries_within_original_deadline(tmp_path, windows_backend, code):
    calls, failures, _ = windows_backend
    failures.append(OSError(code, 'locked'))
    lock = runtime.OfficeJobLock(tmp_path)
    lock.acquire(time.monotonic() + 3)
    lock.close()
    assert [call[0] for call in calls] == [2, 2, 0]


def test_windows_permanent_contention_times_out_and_closes(tmp_path, windows_backend, monkeypatch):
    calls, failures, streams = windows_backend
    failures.extend([OSError(errno.EACCES, 'locked')] * 3)
    clock = iter([0, 0, 1])
    monkeypatch.setattr(runtime, 'time', SimpleNamespace(monotonic=lambda: next(clock), sleep=lambda _: None))
    lock = runtime.OfficeJobLock(tmp_path)
    with pytest.raises(runtime.NativeOfficeError) as caught:
        lock.acquire(.5)
    assert caught.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert len(calls) == 1 and streams[0].closed and lock.stream is None


@pytest.mark.parametrize('failure', [OSError(errno.EBADF, 'bad descriptor'),
                                   BlockingIOError(errno.EIO, 'storage failure')])
def test_windows_unrelated_io_failure_is_not_retried_or_masked(tmp_path, windows_backend, failure):
    calls, failures, streams = windows_backend
    failures.append(failure)
    lock = runtime.OfficeJobLock(tmp_path)
    with pytest.raises(OSError) as caught:
        lock.acquire(time.monotonic() + 3)
    assert caught.value is failure
    assert len(calls) == 1 and streams[0].closed and lock.stream is None


def test_windows_unlock_error_still_closes_descriptor(tmp_path, windows_backend):
    _, failures, streams = windows_backend
    lock = runtime.OfficeJobLock(tmp_path)
    lock.acquire(time.monotonic() + 3)
    failure = OSError(errno.EBADF, 'unlock failed')
    failures.append(failure)
    with pytest.raises(OSError) as caught:
        lock.close()
    assert caught.value is failure
    assert streams[0].closed and lock.stream is None
    lock.close()


@pytest.mark.parametrize('backend', ['native', 'windows'])
def test_quarantine_blocks_normal_jobs_but_allows_recovery(tmp_path, backend, request):
    if backend == 'windows':
        request.getfixturevalue('windows_backend')
    lock = runtime.OfficeJobLock(tmp_path)
    detail = {'component': 'presentation', 'reason': 'completion uncertain'}
    lock.quarantine(detail)
    quarantine = tmp_path / 'native-office.quarantine.json'
    with pytest.raises(runtime.NativeOfficeError) as caught:
        lock.acquire(time.monotonic() + 3)
    assert caught.value.code == 'NATIVE_OFFICE_QUARANTINED'
    assert caught.value.quarantine_path == str(quarantine)
    assert lock.stream is None
    lock.acquire(time.monotonic() + 3, recovery=True)
    assert lock.path == tmp_path / 'native-office.lock'
    lock.close()
    assert json.loads(quarantine.read_text()) == detail
