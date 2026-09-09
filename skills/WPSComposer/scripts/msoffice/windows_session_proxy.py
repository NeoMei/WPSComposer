"""Deadline-bound Windows Microsoft sessions over one private Python worker.

Windows pipes are read/written by daemon threads, not select(). A timed-out
request terminates only the exact Python child and quarantines the component;
Office and all uncertain document/recovery files are left intact.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
import os
from pathlib import Path, PurePath
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time

from .errors import NativeWordError, NativeWordTimeoutError, NATIVE_WORD_ERROR_CODES
from .office_errors import NativeOfficeError, NATIVE_OFFICE_ERROR_CODES
from .windows_office_runtime import OfficeJobLock, _component_root
from .input_validation import validate_native_input
from .windows_session_worker import (PROTOCOL, MAX_FRAME_BYTES, COMMON_METHODS,
    BUSINESS_METHODS, GETTERS, UNSUPPORTED_BUSINESS, HANDLE_TYPES, encode_frame, decode_frame)
from .windows_business_protocol import DTO_TAG, decode_value, encode_value

_PACKAGE_ROOT = Path(__file__).resolve().parents[4]
_MODULE = 'skills.WPSComposer.scripts.msoffice.windows_session_worker'
_COMPONENT = {'writer': 'writer', 'sheet': 'spreadsheet', 'slide': 'presentation'}


@dataclass(frozen=True)
class NativeSessionHandle:
    """Opaque, session-bound result; never exposes a COM attribute/method chain."""
    _owner: object
    id: str
    type: str


def _arguments(value, owner):
    if isinstance(value, NativeSessionHandle):
        if value._owner is not owner:
            raise ValueError('Native handle belongs to another session')
        return {'__wpscomposer_handle__': {'id': value.id, 'type': value.type}}
    if isinstance(value, PurePath):
        return str(value)
    if isinstance(value, tuple):
        return {DTO_TAG: {
            'type': 'tuple',
            'value': [_arguments(item, owner) for item in value],
        }}
    if isinstance(value, list):
        return [_arguments(item, owner) for item in value]
    if isinstance(value, dict):
        return {key: _arguments(item, owner) for key, item in value.items()}
    return encode_value(value)


def _result(value, owner):
    if isinstance(value, list): return [_result(item, owner) for item in value]
    if isinstance(value, dict):
        if set(value) == {'__wpscomposer_handle__'}:
            ref = value['__wpscomposer_handle__']
            if (not isinstance(ref, dict) or set(ref) != {'id','type'} or
                    not isinstance(ref['id'], str) or not re.fullmatch('[0-9a-f]{32}', ref['id']) or
                    ref['type'] not in HANDLE_TYPES):
                raise ValueError('Invalid native handle response')
            return NativeSessionHandle(owner, ref['id'], ref['type'])
        if set(value) == {'__wpscomposer_tuple__'}:
            if not isinstance(value['__wpscomposer_tuple__'], list): raise ValueError('Invalid native tuple response')
            return tuple(_result(item, owner) for item in value['__wpscomposer_tuple__'])
        if set(value) == {DTO_TAG}:
            return decode_value(value)
        return {key: _result(item, owner) for key,item in value.items()}
    return decode_value(value)


def _spawn_worker(job):
    env = os.environ.copy()
    env['PYTHONPATH'] = str(_PACKAGE_ROOT) + os.pathsep + env.get('PYTHONPATH', '')
    for key in ('TMP', 'TEMP', 'TMPDIR'):
        env[key] = str(job)
    with (job / 'worker-stderr.log').open('wb') as log:
        os.chmod(job / 'worker-stderr.log', 0o600)
        return subprocess.Popen([sys.executable, '-u', '-m', _MODULE, str(job)],
            cwd=str(_PACKAGE_ROOT), env=env, stdin=subprocess.PIPE,
            stdout=subprocess.PIPE, stderr=log)


class _SessionProxy:
    engine = 'msoffice'
    kind = None

    def __init__(self):
        raise TypeError('Use open_document or attach_active')

    @classmethod
    def open_document(cls, path, *, read_only=False, visible=False):
        return cls._start('open_document', [str(Path(path).expanduser().resolve())],
                          {'read_only': read_only, 'visible': visible})

    @classmethod
    def attach_active(cls):
        return cls._start('attach_active', [], {})

    @classmethod
    def new_document(cls, *, visible=False):
        return cls._start('new_document', [], {'visible': visible})

    @classmethod
    def _start(cls, method, args, kwargs):
        self = cls.__new__(cls)
        self._deadline = time.monotonic() + 600
        self._closed = self._uncertain = False
        self._sequence = 0
        self._identity = {}
        self._last_request = None
        self._child = None
        self._serial = threading.Lock()
        self._responses = queue.Queue(maxsize=2)
        self._stop = threading.Event()
        self._reader_error = None
        self._reader = self._writer = None
        if method == 'open_document':
            validate_native_input(
                Path(args[0]), _COMPONENT[cls.kind], deadline=self._deadline
            )
        root = _component_root(_COMPONENT[cls.kind])
        if root.is_symlink():
            raise ValueError('Office staging root must not be a symlink')
        root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self._lock = OfficeJobLock(root)
        self._lock.acquire(self._deadline)
        try:
            self.staging_root = Path(tempfile.mkdtemp(prefix='session-', dir=root))
            os.chmod(self.staging_root, 0o700)
            self._child = _spawn_worker(self.staging_root)
            self._reader = threading.Thread(target=self._read_responses, daemon=True)
            self._reader.start()
            if self._call(method, *args, **kwargs) != {'kind': cls.kind}:
                raise ValueError('Worker did not acknowledge the requested document binding')
            return self
        except BaseException:
            if self._child is not None and not self._uncertain:
                self._abort('Initial document binding failed')
            self._lock.close()
            raise

    @property
    def publication_deadline(self):
        return self._deadline

    def _error(self, suffix):
        locations = {'staging_path': self.staging_root,
                     'diagnostic_path': self.staging_root / 'recovery.json',
                     'quarantine_path': self._lock.quarantine_path if self._uncertain else None}
        if self.kind == 'writer':
            return NativeWordTimeoutError(**locations) if suffix == 'TIMEOUT' else NativeWordError('NATIVE_WORD_' + suffix, **locations)
        return NativeOfficeError('NATIVE_OFFICE_' + suffix, **locations)

    def _remaining(self):
        value = self._deadline - time.monotonic()
        if value <= 0:
            raise TimeoutError('Session deadline expired')
        return value

    def _read_responses(self):
        try:
            with (self.staging_root / 'worker-stdout.log').open('wb') as log:
                os.chmod(self.staging_root / 'worker-stdout.log', 0o600)
                while not self._stop.is_set():
                    raw = self._child.stdout.readline(MAX_FRAME_BYTES + 1)
                    log.write(raw)
                    log.flush()
                    if not raw:
                        raise EOFError('Session worker exited without a response')
                    if len(raw) > MAX_FRAME_BYTES or not raw.endswith(b'\n'):
                        raise ValueError('Invalid or oversized worker response')
                    self._responses.put_nowait(raw)
        except BaseException as exc:
            try:
                self._responses.put_nowait(exc)
            except queue.Full:
                self._reader_error = ValueError('Session worker sent unsolicited responses')
        finally:
            self._child.stdout.close()

    def _stop_io(self):
        self._stop.set()
        for thread in (self._writer, self._reader):
            if thread is not None:
                thread.join(timeout=.2)
        # The reader owns stdout and closes it only after its blocking read ends.
        # Never block the deadline thread acquiring BufferedReader's read lock.
        if self._child is not None and (self._writer is None or not self._writer.is_alive()):
            try:
                self._child.stdin.close()
            except OSError:
                pass  # A broken pipe was already recorded as unverified cleanup.

    def _write_request(self, raw, done):
        try:
            self._child.stdin.write(raw)
            self._child.stdin.flush()
            done.put(None)
        except BaseException as exc:
            done.put(exc)

    def _abort(self, reason):
        if self._uncertain:
            return
        self._uncertain = True
        detail = {'protocol': PROTOCOL, 'component': _COMPONENT[self.kind],
                  'reason': reason, 'staging_path': str(self.staging_root),
                  'python_pid': getattr(self._child, 'pid', None), 'identity': self._identity,
                  'last_request': self._last_request, 'cleanup_verified': False,
                  'office_termination_attempted': False}
        try:
            (self.staging_root / 'recovery.json').write_text(json.dumps(detail, ensure_ascii=False), encoding='utf-8')
            if not self._lock.quarantine_path.exists():
                self._lock.quarantine(detail)
        finally:
            if self._child is not None and self._child.poll() is None:
                self._child.kill()
                try:
                    self._child.wait(timeout=1)
                except subprocess.TimeoutExpired:
                    pass
            self._stop_io()
            self._lock.close()

    def _remote_error(self, detail):
        if not isinstance(detail, dict):
            raise ValueError('Malformed remote error')
        code = detail.get('code')
        locations = dict(staging_path=self.staging_root, diagnostic_path=self.staging_root / 'worker-error.log')
        if code in NATIVE_WORD_ERROR_CODES:
            error = NativeWordTimeoutError(**locations) if code == 'NATIVE_WORD_TIMEOUT' else NativeWordError(code, **locations)
        elif code in NATIVE_OFFICE_ERROR_CODES:
            error = NativeOfficeError(code, **locations)
        else:
            error = None
        types = {'ValueError': ValueError, 'PermissionError': PermissionError,
                 'FileExistsError': FileExistsError, 'NotImplementedError': NotImplementedError,
                 'TypeError': TypeError, 'OSError': OSError, 'RuntimeError': RuntimeError,
                 'TimeoutError': TimeoutError}
        if error is None:
            if detail.get('type') not in types or not isinstance(detail.get('message'), str):
                raise ValueError('Malformed remote error type')
            error = types[detail['type']](detail['message'])
        if 'clipboard_changed' in detail:
            if type(detail['clipboard_changed']) is not bool:
                raise ValueError('Malformed clipboard side-effect flag')
            error.clipboard_changed = detail['clipboard_changed']
        return error

    def _call(self, method, *args, **kwargs):
        if self._uncertain:
            raise self._error('QUARANTINED')
        if self._closed:
            raise ValueError('Microsoft session is closed')
        if method not in COMMON_METHODS | BUSINESS_METHODS[self.kind] | {'open_document', 'attach_active', 'new_document', 'get_property'}:
            raise NotImplementedError('Native method has no value-only session contract')
        # Validate arguments before a request number or native operation exists.
        args = [_arguments(item, self) for item in args]
        kwargs = _arguments(kwargs, self)
        encode_frame({'args': args, 'kwargs': kwargs})
        acquired = False
        try:
            acquired = self._serial.acquire(timeout=self._remaining())
            if not acquired:
                raise TimeoutError('Session request queue deadline expired')
            self._remaining()
            self._sequence += 1
            request = {'protocol': PROTOCOL, 'id': self._sequence, 'kind': self.kind,
                       'method': method, 'args': list(args), 'kwargs': kwargs,
                       'deadline': self._deadline, 'remaining_seconds': self._remaining()}
            raw = encode_frame(request)
            self._last_request = request
            (self.staging_root / 'last-request.json').write_bytes(raw)
            done = queue.Queue(maxsize=1)
            self._writer = threading.Thread(target=self._write_request, args=(raw, done), daemon=True)
            self._writer.start()
            written = done.get(timeout=self._remaining())
            if isinstance(written, BaseException): raise written
            response = self._responses.get(timeout=self._remaining())
            if self._reader_error is not None: raise self._reader_error
            if isinstance(response, BaseException): raise response
            result = decode_frame(response)
            if (not isinstance(result, dict) or type(result.get('protocol')) is not int or result.get('protocol') != PROTOCOL or
                    type(result.get('id')) is not int or result['id'] != self._sequence or
                    result.get('status') not in {'ok', 'error'}):
                raise ValueError('Invalid session response envelope or sequence')
            if 'identity' in result:
                if not isinstance(result['identity'], dict): raise ValueError('Invalid Office identity report')
                self._identity = result['identity']
            self._remaining()
            if result['status'] == 'error':
                if method == 'close' and kwargs == {'save_changes': True}:
                    if result.get('session_state') != 'open':
                        raise ValueError(
                            'Saving close error did not acknowledge an open session'
                        )
                elif 'session_state' in result:
                    raise ValueError('Unexpected session state acknowledgement')
                error = self._remote_error(result.get('error'))
            else:
                if 'session_state' in result:
                    raise ValueError('Unexpected session state acknowledgement')
                if 'value' not in result: raise ValueError('Missing session response value')
                return _result(result['value'], self)
        except (TimeoutError, queue.Empty):
            self._abort('Native session request exceeded the total deadline')
            raise self._error('TIMEOUT') from None
        except BaseException:
            self._abort('Worker response or request completion could not be verified')
            raise self._error('QUARANTINED') from None
        finally:
            if acquired:
                self._serial.release()
        # A valid, completed native error remains catchable as its original
        # public category; close still gets a bounded cleanup opportunity.
        if isinstance(error, TimeoutError) or str(getattr(error, 'code', '')).endswith('_TIMEOUT'):
            self._abort('Worker reported an expired native operation')
            raise self._error('TIMEOUT')
        if str(getattr(error, 'code', '')).endswith('_QUARANTINED'):
            self._abort('Worker reported uncertain native cleanup')
            raise self._error('QUARANTINED')
        raise error

    def __enter__(self):
        if self._closed or self._uncertain:
            raise ValueError('Microsoft session is closed or quarantined')
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            self.close()
        except BaseException:
            if exc is None:
                raise

    def __getattr__(self, name):
        if name in GETTERS[self.kind]:
            return self._call('get_property', name)
        if name in COMMON_METHODS | BUSINESS_METHODS[self.kind]:
            return lambda *args, **kwargs: self._call(name, *args, **kwargs)
        if name in UNSUPPORTED_BUSINESS:
            def unsupported(*args, **kwargs):
                raise NotImplementedError(
                    'This business method is unavailable for this document component'
                )
            return unsupported
        raise AttributeError(name)

    def close(self, save_changes=False):
        if self._closed:
            return
        if self._uncertain:
            raise self._error('QUARANTINED')
        try:
            result = self._call('close', save_changes=save_changes)
        except BaseException:
            # _call leaves a verified remote error non-uncertain. In the saving
            # close path the worker remains alive with the session lock held,
            # so the caller can retry, save elsewhere, or explicitly discard.
            if save_changes and not self._uncertain:
                raise
            if not self._uncertain:
                self._abort('Native session close was not verified')
            raise self._error('QUARANTINED') from None
        try:
            if result != {'closed': True}:
                raise ValueError('Worker did not acknowledge native close')
            self._child.stdin.close()
            self._child.wait(timeout=self._remaining())
            if self._child.returncode != 0:
                raise ValueError('Session worker exited unsuccessfully')
            self._closed = True
            self._stop_io()
        except BaseException:
            self._abort('Native session close was not verified')
            raise self._error('QUARANTINED') from None
        self._lock.close()


class ProxyWordSession(_SessionProxy):
    kind = 'writer'


class ProxyExcelSession(_SessionProxy):
    kind = 'sheet'


class ProxyPowerPointSession(_SessionProxy):
    kind = 'slide'
