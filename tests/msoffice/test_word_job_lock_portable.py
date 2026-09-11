"""OS-lock contracts: real local exclusion plus bounded Windows backend doubles."""
import errno
import os
import sys
import time
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice import macos_runtime as runtime


@pytest.fixture
def windows_backend(monkeypatch):
    calls = []
    failures = []

    def locking(fd, mode, count):
        calls.append((mode, count, os.lseek(fd, 0, os.SEEK_CUR), os.fstat(fd).st_size))
        if failures:
            raise failures.pop(0)

    monkeypatch.setattr(runtime, 'os', SimpleNamespace(**dict(vars(os), name='nt')))
    monkeypatch.setitem(sys.modules, 'msvcrt', SimpleNamespace(
        locking=locking, LK_NBLCK=2, LK_UNLCK=0))
    monkeypatch.setitem(sys.modules, 'fcntl', None)
    return calls, failures


def test_windows_lock_and_unlock_use_the_same_byte_without_initialization(tmp_path, windows_backend):
    calls, _ = windows_backend
    lock = runtime.WordJobLock(tmp_path / 'job.lock')
    lock.acquire(time.monotonic() + 3)
    stream = lock.file
    lock.acquire(time.monotonic() + 3)
    lock.close()
    lock.close()
    assert calls == [(2, 1, 0, 0), (0, 1, 0, 0)]
    assert stream.closed and lock.file is None


def test_windows_lock_retries_contention_with_the_original_deadline(tmp_path, windows_backend):
    calls, failures = windows_backend
    failures.append(OSError(errno.EACCES, 'locked'))
    lock = runtime.WordJobLock(tmp_path / 'job.lock')
    lock.acquire(time.monotonic() + 3)
    lock.close()
    assert [row[0] for row in calls] == [2, 2, 0]


def test_windows_lock_does_not_retry_unrelated_io_failure(tmp_path, windows_backend):
    calls, failures = windows_backend
    failure = OSError(errno.EBADF, 'bad descriptor')
    failures.append(failure)
    lock = runtime.WordJobLock(tmp_path / 'job.lock')
    with pytest.raises(OSError) as caught:
        lock.acquire(time.monotonic() + 3)
    assert caught.value is failure
    assert len(calls) == 1 and lock.file is None


def test_windows_quarantine_still_blocks_normal_acquisition(tmp_path, windows_backend):
    lock = runtime.WordJobLock(tmp_path / 'job.lock')
    lock.quarantine({'reason': 'unverified native completion'})
    with pytest.raises(runtime.NativeWordError):
        lock.acquire(time.monotonic() + 3)
    assert lock.file is None
    lock.acquire(time.monotonic() + 3, recovery=True)
    lock.close()
    assert lock.quarantine_path.is_file()
