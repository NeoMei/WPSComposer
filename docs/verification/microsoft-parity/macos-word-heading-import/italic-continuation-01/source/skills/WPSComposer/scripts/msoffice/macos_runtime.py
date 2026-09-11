"""Deadline-bound, serialized native Microsoft Word adapter for macOS."""
from __future__ import annotations

import hashlib
import errno
import json
import math
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
from uuid import uuid4

from ..artifact_transport import validate_office_package, validate_pdf, publish_artifact, copy_file_before_deadline
from ..longform.pipeline import _build_executor_resources
from ..longform.platform_runtime import _BaseAdapter, _apply_relayout
from .errors import NativeWordError, NativeWordTimeoutError
from .macos_script import (
    MacWordCapabilityError, apple_string, compile_plan, parse_result,
    pagination_source, refresh_source, wrap_owned,
)


def remaining(deadline: float) -> float:
    if not isinstance(deadline, (int, float)) or not math.isfinite(deadline):
        raise ValueError('Word deadline must be finite')
    value = deadline - time.monotonic()
    if value <= 0:
        raise NativeWordTimeoutError()
    return value


class WordJobLock:
    """Advisory OS lock shared by native jobs, portable for host-independent tests."""
    def __init__(self, path: Path):
        self.path = Path(path)
        self.file = None
        self.quarantine_path = self.path.with_name(self.path.name + '.quarantine')

    def quarantine(self, detail):
        fd = os.open(self.quarantine_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, 'O_NOFOLLOW', 0), 0o600)
        with os.fdopen(fd, 'w') as stream:
            stream.write(json.dumps(detail, ensure_ascii=False))
            stream.flush()
            os.fsync(stream.fileno())

    def acquire(self, deadline: float, *, recovery=False):
        if self.file is not None:
            return
        fd = os.open(self.path, os.O_CREAT | os.O_RDWR | getattr(os, 'O_NOFOLLOW', 0), 0o600)
        stream = os.fdopen(fd, 'r+b')
        try:
            while True:
                remaining(deadline)
                try:
                    if os.name == 'nt':
                        import msvcrt
                        # CRT locking permits a region beyond EOF. Avoid an
                        # initialization write racing an already-held lock.
                        stream.seek(0)
                        try:
                            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                        except OSError as exc:
                            if exc.errno in (errno.EACCES, errno.EAGAIN, errno.EDEADLK):
                                raise BlockingIOError(*exc.args) from exc
                            raise
                    else:
                        import fcntl
                        fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    if self.quarantine_path.exists() and not recovery:
                        raise NativeWordError('NATIVE_WORD_QUARANTINED', quarantine_path=self.quarantine_path)
                    self.file = stream
                    return
                except BlockingIOError:
                    time.sleep(min(0.05, remaining(deadline)))
        except BaseException:
            stream.close()
            raise

    def close(self):
        if self.file is not None:
            try:
                if os.name == 'nt':
                    import msvcrt
                    self.file.seek(0)
                    msvcrt.locking(self.file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self.file.fileno(), fcntl.LOCK_UN)
            finally:
                self.file.close()
                self.file = None


class MacWordAdapter(_BaseAdapter):
    def __init__(self, build):
        super().__init__(build)
        self.staging_root = None
        self.lock = None
        self.quarantined = False
        self.published = False
        self.nodes = {}
        self.operations = 0
        self.issues = ()
        self.deadline = None

    def _ensure_started(self, deadline):
        remaining(deadline)
        if self.deadline is None:
            self.deadline = deadline
        elif deadline > self.deadline + 0.01:
            raise ValueError('Native Word job cannot extend its total deadline')
        if self.staging_root is not None:
            return
        if sys.platform != 'darwin' or not Path('/Applications/Microsoft Word.app').is_dir():
            raise NativeWordError('NATIVE_WORD_UNAVAILABLE')
        parent = Path.home() / 'Library/Containers/com.microsoft.Word/Data/tmp'
        if not parent.is_dir():
            raise NativeWordError('NATIVE_WORD_UNAVAILABLE')
        self.lock = WordJobLock(parent/'wpscomposer-native-word.lock')
        self.lock.acquire(deadline)
        self.staging_root = Path(tempfile.mkdtemp(prefix='wpscomposer-native-', dir=parent))
        os.chmod(self.staging_root, 0o700)

    def _run(self, source: str, deadline: float):
        self._ensure_started(deadline)
        path = self.staging_root / ('operation-' + uuid4().hex + '.applescript')
        path.write_text(source, encoding='utf-8')
        os.chmod(path, 0o600)
        try:
            result = subprocess.run(['/usr/bin/osascript', str(path)], capture_output=True, text=True, timeout=remaining(deadline))
        except (subprocess.TimeoutExpired, OSError) as error:
            self.quarantined = True
            stdout = getattr(error, 'stdout', '') or ''
            stderr = getattr(error, 'stderr', '') or ''
            stdout = stdout.decode('utf-8', 'replace') if isinstance(stdout, bytes) else stdout
            stderr = stderr.decode('utf-8', 'replace') if isinstance(stderr, bytes) else stderr
            diagnostic = path.with_suffix('.log')
            diagnostic.write_text(stderr + '\n' + stdout, encoding='utf-8')
            self.lock.quarantine({'schema':1, 'stagingRoot':str(self.staging_root), 'diagnostic':str(diagnostic), 'reason':'Native AppleEvent completion uncertain'})
            locations = dict(staging_path=self.staging_root, diagnostic_path=diagnostic,
                             quarantine_path=self.lock.quarantine_path)
            if isinstance(error, subprocess.TimeoutExpired):
                raise NativeWordTimeoutError(**locations) from None
            raise NativeWordError('NATIVE_WORD_QUARANTINED', **locations) from None
        diagnostic = result.stderr + '\n' + result.stdout
        # Diagnostics contain only the task's script/results, never sentinel text.
        path.with_suffix('.log').write_text(diagnostic, encoding='utf-8')
        if result.returncode:
            self.quarantined = True
            if 'WPSC_CLEAN' not in diagnostic.splitlines():
                self.lock.quarantine({'schema':1,'stagingRoot':str(self.staging_root),'diagnostic':str(path.with_suffix('.log')),'reason':'Native cleanup or sentinel preservation unverified'})
            locations = dict(staging_path=self.staging_root, diagnostic_path=path.with_suffix('.log'),
                             quarantine_path=self.lock.quarantine_path if self.lock.quarantine_path.exists() else None)
            if any(line.startswith('WPSC_ERROR\t-1712\t') for line in diagnostic.splitlines()):
                raise NativeWordTimeoutError(**locations)
            code = 'NATIVE_WORD_QUARANTINED' if locations['quarantine_path'] else 'NATIVE_WORD_EXECUTION_FAILED'
            raise NativeWordError(code, **locations)
        try:
            remaining(deadline)
        except NativeWordTimeoutError:
            raise NativeWordTimeoutError(staging_path=self.staging_root,
                                         diagnostic_path=path.with_suffix('.log')) from None
        return diagnostic

    def execute(self, build, directives, deadline):
        derived = _apply_relayout(build.plan, directives)
        # Compile before allocating staging or sending any AppleEvents.
        compile_plan(derived, {r.resource_id: Path('/preflight/resource') for r in build.preflight.resources}, Path('/preflight/owned.docx'), timeout=remaining(deadline))
        self._ensure_started(deadline)
        resources = {}
        for resource in _build_executor_resources(build.base_dir, build.preflight):
            if hashlib.sha256(resource.payload_bytes).hexdigest() != resource.payload_sha256:
                raise ValueError('Native Word resource payload digest mismatch')
            suffix = {'image/png':'.png','image/jpeg':'.jpg','image/svg+xml':'.svg'}.get(resource.media_type)
            if suffix is None:
                raise MacWordCapabilityError('Mac Word unsupported normalized resource type')
            path = self.staging_root / ('resource-' + uuid4().hex + suffix)
            path.write_bytes(resource.payload_bytes)
            resources[resource.id] = path
        target = self.staging_root / ('document-' + uuid4().hex + '.docx')
        compiled = compile_plan(derived, resources, target, timeout=max(1, remaining(deadline)-2))
        raw = self._run(compiled.source, deadline)
        outcome = parse_result(raw, compiled.nodes, target, compiled.operations, compiled.issues)
        validate_office_package(target, 'docx')
        self.nodes, self.operations, self.issues = compiled.nodes, compiled.operations, outcome.issues
        return outcome

    def _owned_path(self, path):
        path = Path(path).resolve()
        if self.staging_root is None or path.parent != self.staging_root.resolve() or not path.is_file():
            raise ValueError('Native Word operation requires an existing task-owned staged artifact')
        return path

    def export_pdf(self, docx, deadline):
        self._ensure_started(deadline)
        source = self._owned_path(docx)
        target = self.staging_root / ('quality-' + uuid4().hex + '.pdf')
        body = f'''repaginate ownedDoc
save as ownedDoc file name {apple_string(str(target))} file format format PDF add to recent files false
set ownedDoc to missing value
repeat with documentIndex from 1 to (count of documents)
 set d to document documentIndex
 if (posix full name of d is {apple_string(str(source))}) or (posix full name of d is {apple_string(str(target))}) then set ownedDoc to document (name of d)
end repeat'''
        raw = self._run(wrap_owned(body, target, max(1,remaining(deadline)-2), source=source), deadline)
        parse_result(raw, {}, target, 0)
        validate_pdf(target)
        return target

    def patch_quality_notices(self, docx, notices, deadline):
        source = self._owned_path(docx)
        target = self.staging_root / ('notices-' + uuid4().hex + '.docx')
        copy_file_before_deadline(source, target, deadline=deadline)
        body = []
        for notice in notices:
            text = '[' + notice.code + '] ' + notice.message
            bookmark = self.nodes.get(notice.node_id, self.nodes.get('doc:quality', (None,'')))[0]
            if not bookmark:
                raise ValueError('Native Word quality notice anchor is unavailable')
            body += [f'set r to text object of bookmark {apple_string(bookmark)} of ownedDoc', 'set p to end of content of r', 'set r to create range ownedDoc start p end p', f'set content of r to {apple_string(text)} & return']
        body += [refresh_source(), pagination_source(self.nodes), f'save as ownedDoc file name {apple_string(str(target))} file format format document default add to recent files false']
        raw = self._run(wrap_owned('\n'.join(body), target, max(1,remaining(deadline)-2), source=target), deadline)
        result = parse_result(raw, self.nodes, target, self.operations, self.issues)
        validate_office_package(target, 'docx')
        return result

    def cleanup(self, path):
        if self.published and not self.quarantined:
            self._owned_path(path).unlink(missing_ok=True)

    def publish(self, staged, output, overwrite, deadline):
        result = super().publish(staged, output, overwrite, deadline)
        self.published = True
        return result

    def close(self):
        # Every successful AppleScript operation closes its retained document.
        # Never delete uncertain files after a timeout or ownership failure.
        if self.staging_root is not None and self.published and not self.quarantined:
            shutil.rmtree(self.staging_root, ignore_errors=True)
        if self.lock is not None:
            self.lock.close()


def recover_quarantine(*, timeout=30, lock_path=None):
    """Explicitly clear quarantine only after read-only process/document checks."""
    if not isinstance(timeout, (int,float)) or isinstance(timeout,bool) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('Recovery timeout must be finite and positive')
    deadline = time.monotonic() + timeout
    parent = Path.home() / 'Library/Containers/com.microsoft.Word/Data/tmp'
    lock = WordJobLock(lock_path or parent/'wpscomposer-native-word.lock')
    lock.acquire(deadline, recovery=True)
    try:
        if not lock.quarantine_path.exists():
            return False
        data = json.loads(lock.quarantine_path.read_text(encoding='utf-8'))
        if not isinstance(data,dict) or data.get('schema') != 1:
            raise RuntimeError('Quarantine metadata is invalid; refusing automatic recovery')
        staging = Path(data['stagingRoot']).resolve()
        if staging.parent != parent.resolve() or not staging.name.startswith(('wpscomposer-native-', 'wpscomposer-session-')):
            raise ValueError('Quarantine staging identity is invalid')
        bound = data.get('documentPath')
        if bound is not None and Path(bound).resolve().parent != staging:
            raise RuntimeError('Attached Word recovery requires manual verification')
        process = subprocess.run(['/bin/ps','-axo','command'],capture_output=True,text=True,check=True,timeout=remaining(deadline))
        if any('osascript' in line and str(staging) in line for line in process.stdout.splitlines()):
            raise RuntimeError('Previous native Word script is still running; recovery refused')
        if '/Microsoft Word.app/Contents/MacOS/Microsoft Word' in process.stdout:
            script = f'''with timeout of {max(1,int(remaining(deadline)))} seconds
 tell application "Microsoft Word"
  repeat with documentIndex from 1 to (count of documents)
   set d to document documentIndex
   set docPath to posix full name of d
   if docPath is missing value then error "Word document identity unavailable"
   if docPath starts with {apple_string(str(staging) + '/')} then error "Previous task-owned document remains open"
  end repeat
 end tell
end timeout
return "WPSC_RECOVERY_OK"'''
            result = subprocess.run(['/usr/bin/osascript','-e',script],capture_output=True,text=True,timeout=remaining(deadline))
            if result.returncode or result.stdout.strip() != 'WPSC_RECOVERY_OK':
                raise RuntimeError('Native Word owned-document cleanup is not verified; recovery refused')
        remaining(deadline)
        lock.quarantine_path.unlink()
        return True
    finally:
        lock.close()


def convert(request, timeout=600) -> Path:
    if isinstance(timeout, bool) or not isinstance(timeout,(int,float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('Native Word timeout must be finite and positive')
    if request.component != 'writer' or request.source.suffix.lower() not in ('.doc', '.docx'):
        raise MacWordCapabilityError('Mac native Word conversion requires DOC or DOCX')
    if request.output.exists() and not request.overwrite:
        raise FileExistsError(request.output)
    if request.source.suffix.lower() == '.docx':
        validate_office_package(request.source, 'docx')
    from ..longform.pipeline import build_longform_generation
    adapter = MacWordAdapter(build_longform_generation(''))
    deadline = time.monotonic() + timeout
    try:
        adapter._ensure_started(deadline)
        staged = adapter.staging_root / ('source-' + uuid4().hex + request.source.suffix.lower())
        copy_file_before_deadline(request.source, staged, deadline=deadline)
        pdf = adapter.export_pdf(staged, deadline)
        return adapter.publish(pdf, request.output, request.overwrite, deadline)
    finally:
        adapter.close()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Explicit native Word quarantine recovery')
    parser.add_argument('--recover', action='store_true', required=True)
    parser.add_argument('--timeout', type=float, default=30)
    args = parser.parse_args()
    print('Quarantine cleared after verification' if recover_quarantine(timeout=args.timeout) else 'No quarantine marker exists')
