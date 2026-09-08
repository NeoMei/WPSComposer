"""Deadline-bound Word workers behind the existing M5 lifecycle.

The private JSON protocol carries a validated GenerationPlan, normalized
resource bytes as base64 plus ImageProfile metadata, QualityFinding dictionaries
and ExecutionOutcome dictionaries. It never deserializes executable objects or
re-reads original image paths. Each worker owns a private operation directory;
only Python is terminated on timeout, with uncertain Word files retained.
"""
from __future__ import annotations

import base64
from dataclasses import asdict
import json
import math
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from uuid import uuid4

from ..artifact_transport import (
    copy_file_before_deadline, publish_artifact, validate_pdf,
    validate_office_package,
)
from ..generation_plan import validate_generation_plan
from ..longform.executor import ExecutionOutcome
from ..longform.pipeline import _build_executor_resources
from ..longform.platform_runtime import _BaseAdapter, _apply_relayout
from ..longform.quality import QualityFinding
from ..longform.resources import ImageProfile, PreparedLongformResource


_MODULE = 'skills.WPSComposer.scripts.msoffice.windows_runtime'
_PACKAGE_ROOT = Path(__file__).resolve().parents[4]


def _require_deadline(deadline):
    if isinstance(deadline, bool) or not isinstance(deadline, (int, float)) or not math.isfinite(deadline):
        raise ValueError('Word deadline must be a finite monotonic timestamp')
    if time.monotonic() >= deadline:
        raise TimeoutError('Native Word deadline expired')


def _private_root():
    root = Path(tempfile.mkdtemp(prefix='wpscomposer-native-word-'))
    os.chmod(root, 0o700)
    return root


def _write_json(path, value):
    with Path(path).open('x', encoding='utf-8') as handle:
        os.chmod(path, 0o600)
        json.dump(value, handle, ensure_ascii=False, allow_nan=False)


def _owned_path(root, value):
    path = Path(value).expanduser().resolve()
    if Path(root).resolve() not in path.parents:
        raise RuntimeError('Native Word returned a path outside its private staging directory')
    if not path.is_file():
        raise RuntimeError('Native Word did not produce its staged artifact')
    return path


def _run_worker(root, payload, deadline):
    """Run a single COM phase within the caller's remaining total budget."""
    _require_deadline(deadline)
    root = Path(root).resolve()
    operation = root / ('operation-' + uuid4().hex)
    operation.mkdir(mode=0o700)
    request = operation / 'request.json'
    response = operation / 'response.json'
    diagnostic = operation / 'diagnostics.json'
    _write_json(request, dict(payload, protocol=1, deadline=deadline))
    _require_deadline(deadline)
    env = os.environ.copy()
    env['PYTHONPATH'] = str(_PACKAGE_ROOT) + os.pathsep + env.get('PYTHONPATH', '')
    command = [sys.executable, '-m', _MODULE, '--worker', str(request), str(response)]
    with (operation / 'worker.log').open('wb') as log:
        os.chmod(operation / 'worker.log', 0o600)
        try:
            child = subprocess.Popen(
                command, cwd=str(_PACKAGE_ROOT), env=env,
                stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
            )
        except BaseException as exc:
            _write_json(diagnostic, {
                'status': 'failed', 'error_type': type(exc).__name__,
                'message': str(exc), 'word_termination_attempted': False,
            })
            raise
        try:
            child.wait(timeout=max(0.001, deadline - time.monotonic()))
        except subprocess.TimeoutExpired:
            # This is the child Python handle, never the Word PID and never a
            # process-tree kill. Word and uncertain open files are left intact.
            child.kill()
            child.wait(timeout=5)
            if not diagnostic.exists():
                _write_json(diagnostic, {'status': 'timeout', 'python_pid': child.pid, 'word_termination_attempted': False})
            raise TimeoutError(f'Native Word timed out; private evidence retained at {operation}') from None
    if child.returncode != 0 or not response.is_file():
        raise RuntimeError(f'Native Word worker failed; private evidence retained at {operation}')
    result = json.loads(response.read_text(encoding='utf-8'))
    if not isinstance(result, dict) or result.get('status') != 'ok':
        raise RuntimeError(f'Native Word worker failed; private evidence retained at {operation}')
    _require_deadline(deadline)
    return result['value']


def _resource_json(resource):
    return {
        'id': resource.id, 'media_type': resource.media_type,
        'source_sha256': resource.source_sha256, 'payload_sha256': resource.payload_sha256,
        'normalizer_id': resource.normalizer_id,
        'payload_base64': base64.b64encode(resource.payload_bytes).decode('ascii'),
        'image_profile': asdict(resource.image_profile),
    }


def _resource_from_json(raw):
    return PreparedLongformResource(
        id=raw['id'], media_type=raw['media_type'],
        source_sha256=raw['source_sha256'], payload_sha256=raw['payload_sha256'],
        normalizer_id=raw['normalizer_id'],
        payload_bytes=base64.b64decode(raw['payload_base64'], validate=True),
        image_profile=ImageProfile(**raw['image_profile']),
    )


class WindowsWordAdapter(_BaseAdapter):
    def __init__(self, build):
        super().__init__(build)
        self.staging_root = _private_root()
        self._failed = False
        self._published = False

    def _call(self, payload, deadline):
        try:
            return _run_worker(self.staging_root, payload, deadline)
        except BaseException:
            self._failed = True
            raise

    def _outcome(self, result):
        try:
            outcome = ExecutionOutcome.from_dict(result['outcome'])
            _owned_path(self.staging_root, outcome.staged_artifact)
            if outcome.pagination_map.version != 'M5-v1':
                raise RuntimeError('Native Word did not return real M5 pagination')
            return outcome
        except BaseException:
            self._failed = True
            raise

    def execute(self, build, directives, deadline):
        _require_deadline(deadline)
        plan = _apply_relayout(build.plan, directives)
        resources = _build_executor_resources(build.base_dir, build.preflight)
        return self._outcome(self._call({
            'action': 'execute', 'plan': plan.to_dict(),
            'resources': [_resource_json(resource) for resource in resources],
        }, deadline))

    def export_pdf(self, docx, deadline):
        source = _owned_path(self.staging_root, docx)
        result = self._call({'action': 'export', 'source': str(source)}, deadline)
        try:
            return _owned_path(self.staging_root, result['path'])
        except BaseException:
            self._failed = True
            raise

    def patch_quality_notices(self, docx, notices, deadline):
        source = _owned_path(self.staging_root, docx)
        return self._outcome(self._call({
            'action': 'patch', 'source': str(source),
            'notices': [notice.to_dict() for notice in notices],
            'bookmarks': self.bookmarks,
        }, deadline))

    def publish(self, staged, output, overwrite, deadline):
        result = super().publish(staged, output, overwrite, deadline)
        self._published = True
        return result

    def cleanup(self, path):
        if self._published and not self._failed:
            _owned_path(self.staging_root, path).unlink(missing_ok=True)

    def close(self):
        if self._published and not self._failed and self.staging_root.exists():
            shutil.rmtree(self.staging_root)


def convert(request, timeout=600) -> Path:
    """Convert a private source copy through verified Word and publish atomically."""
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('Native Word timeout must be a finite positive number')
    deadline = time.monotonic() + timeout
    source = Path(request.source).expanduser().resolve()
    output = Path(request.output).expanduser().resolve()
    if request.component != 'writer' or source.suffix.lower() not in {'.doc', '.docx'}:
        raise ValueError('Native Word PDF conversion supports DOC and DOCX only')
    if output.suffix.lower() != '.pdf':
        raise ValueError('Native Word conversion output must be PDF')
    if not source.is_file():
        raise FileNotFoundError(source)
    if output.exists() and not request.overwrite:
        raise FileExistsError(output)
    root = _private_root()
    succeeded = False
    try:
        staged_source = root / ('source' + source.suffix.lower())
        copy_file_before_deadline(source, staged_source, deadline=deadline)
        result = _run_worker(root, {'action': 'export', 'source': str(staged_source)}, deadline)
        pdf = _owned_path(root, result['path'])
        published = publish_artifact(
            pdf, output, overwrite=request.overwrite,
            validator=validate_pdf, deadline=deadline,
        )
        succeeded = True
        return published
    finally:
        if succeeded:
            shutil.rmtree(root)


def _execute_request(payload, operation):
    action = payload.get('action')
    if payload.get('protocol') != 1 or action not in {'execute', 'export', 'patch'}:
        raise ValueError('Unsupported native Word worker protocol or action')
    deadline = payload['deadline']
    _require_deadline(deadline)
    from .windows_host import create_dedicated_composer
    from ..longform.windows_executor import WindowsLongformExecutor

    created = []

    def factory(staging_dir=None):
        composer = create_dedicated_composer(staging_dir)
        created.append(composer)
        return composer

    factory.allow_host_retry = False
    factory.strict_cleanup = True
    executor = WindowsLongformExecutor(staging_dir=str(operation), composer_factory=factory)
    result = None
    try:
        if action == 'execute':
            plan = validate_generation_plan(payload['plan'], component='writer')
            resources = tuple(_resource_from_json(raw) for raw in payload['resources'])
            outcome = executor.execute(plan, resources, deadline=deadline)
            if outcome.pagination_map.version != 'M5-v1':
                raise RuntimeError('Native Word did not return real M5 pagination')
            validate_office_package(Path(outcome.staged_artifact), 'docx')
            result = {'outcome': outcome.to_dict()}
        else:
            # Always work on a copy, including quality exports and notice patch.
            source = _owned_path(operation.parent, payload['source'])
            owned_source = operation / ('input' + source.suffix.lower())
            copy_file_before_deadline(source, owned_source, deadline=deadline)
            if action == 'patch':
                notices = tuple(QualityFinding.from_dict(raw) for raw in payload['notices'])
                outcome = executor.patch_quality_notices(
                    owned_source, notices, payload['bookmarks'], deadline=deadline,
                )
                validate_office_package(Path(outcome.staged_artifact), 'docx')
                result = {'outcome': outcome.to_dict()}
            else:
                composer = factory(str(operation))
                try:
                    composer.open_owned_document(owned_source, read_only=True)
                    target = operation / 'export.pdf'
                    composer.export_pdf(target)
                finally:
                    composer.close(save_changes=False)
                validate_pdf(target)
                result = {'path': str(target)}
        _require_deadline(deadline)
        return result
    finally:
        # Existing executor versions swallow cleanup errors; do not allow that
        # to report success in this adapter, even before their strict hook lands.
        for composer in created:
            if composer.cleanup_error is not None or not composer._closed:
                raise RuntimeError('Native Word cleanup was not verified; owned stage retained')


def _worker_main(request, response):
    request = Path(request).resolve()
    response = Path(response).resolve()
    operation = request.parent
    if request.name != 'request.json' or response != operation / 'response.json':
        raise ValueError('Worker files must be the private request/response pair')
    try:
        payload = json.loads(request.read_text(encoding='utf-8'))
        value = _execute_request(payload, operation)
        _write_json(response, {'status': 'ok', 'value': value})
        _write_json(operation / 'diagnostics.json', {'status': 'succeeded'})
        return 0
    except BaseException as exc:
        _write_json(operation / 'diagnostics.json', {
            'status': 'failed', 'error_type': type(exc).__name__,
            'message': str(exc), 'traceback': traceback.format_exc(),
            'word_termination_attempted': False,
        })
        return 1


if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[1] != '--worker':
        raise SystemExit('This module only accepts the private --worker protocol')
    raise SystemExit(_worker_main(sys.argv[2], sys.argv[3]))
