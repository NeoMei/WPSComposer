"""Closed-plan Excel/PowerPoint jobs in deadline-bound private Python workers.

This initial adapter implements the seven spreadsheet/nine presentation plan
operations. It does not claim complete Composer parity or native acceptance.
Only Python is terminated on timeout. Uncertain Office state is quarantined;
user processes and source documents are never killed, saved or replaced.
"""
from __future__ import annotations

import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from uuid import uuid4

from ..artifact_transport import (ValidatorSpec, copy_file_before_deadline, publish_artifact,
                                 validate_before_deadline, validate_office_package, validate_pdf)
from ..generation_plan import validate_generation_plan
from ..office_engines import engine_executable, validate_timeout
from .office_errors import NativeOfficeError

_MODULE = 'skills.WPSComposer.scripts.msoffice.windows_office_runtime'
_PACKAGE_ROOT = Path(__file__).resolve().parents[4]
_FORMATS = {'spreadsheet': 'xlsx', 'presentation': 'pptx'}


def _remaining(deadline):
    if isinstance(deadline, bool) or not isinstance(deadline, (int, float)) or not math.isfinite(deadline):
        raise ValueError('Native Office deadline must be finite')
    result = deadline - time.monotonic()
    if result <= 0:
        raise NativeOfficeError('NATIVE_OFFICE_TIMEOUT')
    return result


def _component_root(component):
    base = Path(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir())
    return base / 'WPSComposer' / 'native-office' / component


def _write_json(path, value):
    with Path(path).open('x', encoding='utf-8') as stream:
        os.chmod(path, 0o600)
        json.dump(value, stream, ensure_ascii=False, allow_nan=False)


class OfficeJobLock:
    def __init__(self, root):
        self.path = root / 'native-office.lock'
        self.quarantine_path = root / 'native-office.quarantine.json'
        self.stream = None
        self.locked = False

    def acquire(self, deadline):
        import msvcrt
        if self.path.is_symlink():
            raise ValueError('Office lock must not be a symlink')
        self.stream = self.path.open('a+b')
        os.chmod(self.path, 0o600)
        if self.stream.tell() == 0:
            self.stream.write(b'\0')
            self.stream.flush()
        try:
            while True:
                _remaining(deadline)
                self.stream.seek(0)
                try:
                    msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
                    self.locked = True
                    break
                except OSError:
                    time.sleep(min(.05, _remaining(deadline)))
            if self.quarantine_path.exists():
                raise NativeOfficeError('NATIVE_OFFICE_QUARANTINED', quarantine_path=self.quarantine_path)
        except BaseException:
            self.close()
            raise

    def quarantine(self, detail):
        _write_json(self.quarantine_path, detail)

    def close(self):
        if self.stream is not None:
            try:
                if self.locked:
                    import msvcrt
                    self.stream.seek(0)
                    msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
            finally:
                self.stream.close()
                self.stream = None
                self.locked = False


def _owned_path(root, value):
    raw = Path(value).expanduser()
    path = raw.resolve()
    if raw.is_symlink() or Path(root).resolve() not in path.parents or not path.is_file():
        raise ValueError('Native Office returned a file outside private staging')
    return path


def _worker_diagnostic_path(job, error):
    for path in (job / 'diagnostics.json', job / 'worker.log'):
        try:
            if path.is_file():
                return path
        except BaseException as failure:
            # Evidence discovery cannot substitute an inaccessible path for
            # the native failure being reported.
            error.diagnostic_io_failures = (getattr(error, 'diagnostic_io_failures', ())
                                           + (('worker.diagnostic.stat', type(failure).__name__),))
    return None


def _retain_worker_diagnostic(error, job, detail):
    try:
        _write_json(job / 'diagnostics.json', detail)
    except BaseException as failure:
        error.diagnostic_io_failures = (getattr(error, 'diagnostic_io_failures', ())
                                       + (('worker.diagnostics', type(failure).__name__),))
    if isinstance(error, NativeOfficeError):
        path = _worker_diagnostic_path(job, error)
        error.diagnostic_path = str(path) if path is not None else None


def _worker_failure(job):
    diagnostic = job / 'diagnostics.json'
    data = {}
    try:
        data = json.loads(diagnostic.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        pass
    if not isinstance(data, dict):
        data = {}
    code = data.get('public_code')
    if code not in {'NATIVE_OFFICE_TIMEOUT', 'NATIVE_OFFICE_QUARANTINED', 'NATIVE_OFFICE_UNSUPPORTED', 'NATIVE_OFFICE_UNAVAILABLE'}:
        code = 'NATIVE_OFFICE_EXECUTION_FAILED'
    error = NativeOfficeError(code, staging_path=job)
    path = _worker_diagnostic_path(job, error)
    error.diagnostic_path = str(path) if path is not None else None
    error.cleanup_verified = data.get('cleanup_verified') is True
    return error


def _run_worker(job, payload, deadline):
    primary_error = None
    log = None
    launched = closed = False
    try:
        _remaining(deadline)
        request, response = job / 'request.json', job / 'response.json'
        diagnostic = job / 'diagnostics.json'
        _write_json(request, dict(payload, protocol=1, deadline=deadline))
        env = os.environ.copy()
        env['PYTHONPATH'] = str(_PACKAGE_ROOT) + os.pathsep + env.get('PYTHONPATH', '')
        command = [sys.executable, '-m', _MODULE, '--worker', str(request), str(response)]
        log = (job / 'worker.log').open('wb')
        os.chmod(job / 'worker.log', 0o600)
        try:
            child = subprocess.Popen(command, cwd=str(_PACKAGE_ROOT), env=env,
                                     stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT)
            launched = True
        except BaseException as exc:
            error = (NativeOfficeError('NATIVE_OFFICE_EXECUTION_FAILED', staging_path=job)
                     if isinstance(exc, Exception) else exc)
            error.cleanup_verified = True
            _retain_worker_diagnostic(error, job, {'status': 'launch_failed', 'message': str(exc),
                                      'cleanup_verified': True, 'office_termination_attempted': False})
            raise error from None
        try:
            child.wait(timeout=_remaining(deadline))
        except BaseException as exc:
            # This exact handle is Python, not EXCEL/POWERPNT and not a tree.
            timed_out = isinstance(exc, (subprocess.TimeoutExpired, NativeOfficeError))
            error = (NativeOfficeError('NATIVE_OFFICE_TIMEOUT', staging_path=job)
                     if timed_out else exc)
            cleanup_failures = []
            for label, cleanup in (('worker.kill', child.kill),
                                   ('worker.wait', lambda: child.wait(timeout=1))):
                try:
                    cleanup()
                except BaseException as failure:
                    cleanup_failures.append((label, type(failure).__name__))
            if cleanup_failures:
                error.cleanup_io_failures = tuple(cleanup_failures)
            if _worker_diagnostic_path(job, error) != diagnostic:
                _retain_worker_diagnostic(error, job, {'status': 'timeout' if timed_out else 'interrupted', 'python_pid': child.pid,
                                        'cleanup_verified': False, 'office_termination_attempted': False})
            elif isinstance(error, NativeOfficeError):
                error.diagnostic_path = str(diagnostic)
            raise error from None
        if child.returncode != 0 or not response.is_file():
            raise _worker_failure(job)
        try:
            response_value = json.loads(response.read_text(encoding='utf-8'))
            if response_value.get('status') != 'ok' or response_value['value'].get('cleanup_verified') is not True:
                raise ValueError('Native Office cleanup was not verified')
            closed = True
            _remaining(deadline)
            return response_value['value']
        except (ValueError, KeyError, AttributeError):
            raise _worker_failure(job) from None
    except BaseException as error:
        primary_error = error
        if not launched or closed:
            error.cleanup_verified = True
        raise
    finally:
        if log is not None:
            try:
                log.close()
            except BaseException as failure:
                if primary_error is None:
                    failure.cleanup_verified = not launched or closed
                    raise
                primary_error.cleanup_io_failures = (getattr(primary_error, 'cleanup_io_failures', ())
                                                     + (('worker.log.close', type(failure).__name__),))


def validate_plan(plan, resources):
    raw = plan.to_dict() if hasattr(plan, 'to_dict') else plan
    plan = validate_generation_plan(raw, component=raw.get('component'))
    if plan.component not in _FORMATS:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')
    prefix = 'sheet' if plan.component == 'spreadsheet' else 'slide'
    if not plan.operations or plan.operations[0].op != prefix + '.reset':
        raise ValueError('Native Office plans must start with one reset')
    count = 1 if prefix == 'sheet' else 0
    for index, operation in enumerate(plan.operations):
        name, args = operation.op, operation.args
        if index and name == prefix + '.reset':
            raise ValueError('Repeated reset is not allowed')
        if name in {'sheet.add', 'slide.add_title', 'slide.add_section', 'slide.add_bullets', 'slide.add_blank'}:
            count += 1
        if name in {'sheet.select', 'sheet.rename'} and args['index'] > count:
            raise ValueError('Worksheet index is outside the plan')
        if name in {'slide.add_image', 'slide.add_table'} and args['slide'] > count:
            raise ValueError('Slide index is outside the plan')
        if name == 'slide.add_image' and args['imageId'] not in resources:
            raise ValueError('Slide image resource is missing')
        if name in {'sheet.rename', 'sheet.add'}:
            sheet_name = args['name']
            if not sheet_name or len(sheet_name) > 31 or re.search(r'[\[\]:*?/\\]', sheet_name):
                raise ValueError('Invalid Excel sheet name')
        if name == 'sheet.write_table':
            for row in args['values']:
                for value in row:
                    if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')):
                        # No DDE, external workbooks or callable automation in
                        # renderer input. Formula expansion is a separate gate.
                        functions = re.findall(r'([A-Za-z_][\w.]*)\s*\(', value)
                        if (any(char in value for char in '|[]') or any(fn.upper() not in
                            {'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF', 'ROUND', 'ABS'} for fn in functions)):
                            raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')
    return plan


def _publish(source, output, *, overwrite, deadline):
    fmt = output.suffix.lower().lstrip('.')
    spec = ValidatorSpec.from_callable(validate_pdf) if fmt == 'pdf' else ValidatorSpec.from_callable(validate_office_package, fmt)
    return publish_artifact(source, output, overwrite=overwrite,
                            validator=lambda path: validate_before_deadline(spec, path, deadline), deadline=deadline)


def _execute(component, output, *, overwrite, deadline, prepare):
    if sys.platform != 'win32' or not engine_executable('msoffice', component):
        raise NativeOfficeError('NATIVE_OFFICE_UNAVAILABLE')
    output = Path(output).expanduser().resolve()
    if output.exists() and not overwrite:
        raise FileExistsError(output)
    root = _component_root(component)
    if root.is_symlink():
        raise ValueError('Office staging root must not be a symlink')
    root.mkdir(parents=True, exist_ok=True, mode=0o700)
    lock = OfficeJobLock(root)
    job = None
    launched = closed = published = False
    primary_error = None
    diagnostic_failures = []

    def persist_diagnostic(label, write):
        try:
            write()
        except BaseException as error:
            diagnostic_failures.append((label, type(error).__name__))
            return False
        return True

    try:
        lock.acquire(deadline)
        job = Path(tempfile.mkdtemp(prefix='job-', dir=root))
        payload = prepare(job, deadline)
        launched = True
        result = _run_worker(job, dict(payload, component=component), deadline)
        closed = result.get('cleanup_verified') is True
        if not closed:
            raise NativeOfficeError('NATIVE_OFFICE_QUARANTINED')
        artifact = _owned_path(job, result['path'])
        _publish(artifact, output, overwrite=overwrite, deadline=deadline)
        published = True
        return output
    except BaseException as exc:
        primary_error = exc
        if job is not None:
            closed = closed or getattr(exc, 'cleanup_verified', False)
            detail = {'component': component, 'staging_path': str(job), 'cleanup_verified': closed,
                      'office_termination_attempted': False, 'code': getattr(exc, 'code', 'NATIVE_OFFICE_EXECUTION_FAILED')}
            quarantined = False
            if launched and not closed:
                quarantined = persist_diagnostic('quarantine', lambda: lock.quarantine(detail))
            recovery = job / 'recovery.json'
            recovery_written = persist_diagnostic('recovery', lambda: _write_json(recovery, detail))
            if isinstance(exc, NativeOfficeError):
                exc.staging_path = str(job)
                exc.diagnostic_path = exc.diagnostic_path or (str(recovery) if recovery_written else None)
                exc.quarantine_path = str(lock.quarantine_path) if quarantined else None
            if diagnostic_failures:
                exc.diagnostic_io_failures = (getattr(exc, 'diagnostic_io_failures', ())
                                              + tuple(diagnostic_failures))
        raise
    finally:
        cleanups = [('lock.close', lock.close)]
        if published and job is not None:
            cleanups.append(('staging.remove', lambda: shutil.rmtree(job)))
        cleanup_failures = []
        for label, cleanup in cleanups:
            try:
                cleanup()
            except BaseException as error:
                if primary_error is None:
                    raise
                cleanup_failures.append((label, type(error).__name__))
        if cleanup_failures:
            primary_error.cleanup_io_failures = (getattr(primary_error, 'cleanup_io_failures', ())
                                                 + tuple(cleanup_failures))


def generate_recorded(recorded, output, *, timeout=600, overwrite=False):
    deadline = time.monotonic() + validate_timeout(timeout)
    resources = {resource.id: resource for resource in recorded.resources}
    plan = validate_plan(recorded.plan, resources)
    output = Path(output).expanduser().resolve()
    if output.suffix.lower() != '.' + _FORMATS[plan.component]:
        raise ValueError('Native Office output format must match its component')
    def prepare(job, deadline):
        paths = {}
        for index, resource in enumerate(recorded.resources):
            source = Path(resource.source_path)
            target = job / ('image-' + str(index) + source.suffix.lower())
            copy_file_before_deadline(source, target, deadline=deadline)
            paths[resource.id] = str(target)
        return {'action': 'generate', 'plan': plan.to_dict(), 'resources': paths}
    return _execute(plan.component, output, overwrite=overwrite, deadline=deadline, prepare=prepare)


def generate(doc, format_name, output, preset=None, *, timeout=600, overwrite=False):
    from ..recording_composers import RecordingSheetComposer, RecordingSlideComposer
    from ..renderers import sheet_renderer, slide_renderer
    deadline = time.monotonic() + validate_timeout(timeout)
    try:
        renderer, factory = {'xlsx': (sheet_renderer.render, RecordingSheetComposer),
                             'pptx': (slide_renderer.render, RecordingSlideComposer)}[format_name]
    except KeyError:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED') from None
    recorded = renderer(doc, str(output), preset=preset, composer_factory=factory)
    return generate_recorded(recorded, output, timeout=_remaining(deadline), overwrite=overwrite)


def convert(request, timeout=600):
    from .input_validation import validate_native_input
    deadline = time.monotonic() + validate_timeout(timeout)
    source, output = Path(request.source).expanduser().resolve(), Path(request.output).expanduser().resolve()
    component = request.component
    if component not in _FORMATS or source.suffix.lower() != '.' + _FORMATS[component]:
        raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')
    if output.suffix.lower() != '.pdf' or source == output:
        raise ValueError('Native conversion requires a distinct PDF output')
    validate_native_input(source, component, deadline=deadline)
    def prepare(job, deadline):
        target = job / ('source-' + uuid4().hex + source.suffix.lower())
        copy_file_before_deadline(source, target, deadline=deadline)
        validate_native_input(target, component, deadline=deadline)
        return {'action': 'export', 'source': str(target)}
    return _execute(component, output, overwrite=request.overwrite, deadline=deadline, prepare=prepare)


def execute_plan(composer, plan, resources):
    from .._colors import hex_to_rgb_long
    from ..design_presets import DesignPreset
    plan = validate_plan(plan, resources)
    for operation in plan.operations:
        composer._verify_document()
        name, a = operation.op, operation.args
        if name == 'sheet.reset':
            if composer._doc.Worksheets.Count != 1:
                raise ValueError('Excel initial workbook must contain exactly one worksheet')
            composer.select_sheet(1)
        elif name == 'sheet.rename':
            composer.rename_sheet(a['index'], a['name'])
        elif name == 'sheet.add':
            composer.add_sheet(a['name'])
        elif name == 'sheet.select':
            composer.select_sheet(a['index'])
        elif name == 'sheet.write_table':
            composer.write_table(a['startRow'], a['startCol'], a['values'],
                                 header_bold=a.get('headerBold', True), header_shade=a.get('headerShade', '#4472C4'),
                                 header_font_color=a.get('headerFontColor', '#FFFFFF'), font_size=a.get('fontSize', 11))
            for row_index, row in enumerate(a['values']):
                for col_index, value in enumerate(row):
                    if isinstance(value, str) and value.startswith('='):
                        composer.set_formula(a['startRow'] + row_index, a['startCol'] + col_index, value)
        elif name == 'sheet.set_column_width':
            composer.set_column_width(a['column'], a['width'])
        elif name == 'sheet.autofit':
            composer.ws.UsedRange.Columns.AutoFit()  # shared helper swallows errors
        elif name == 'slide.reset':
            if composer.slide_count != 0:
                raise ValueError('PowerPoint initial presentation must have no slides')
        elif name == 'slide.set_size':
            composer._doc.PageSetup.SlideWidth = a['width']
            composer._doc.PageSetup.SlideHeight = a['height']
        elif name == 'slide.apply_preset':
            raw = a['preset']
            colors = dict(raw['colors'])
            colors['bg'] = colors.pop('background', '#FFFFFF')
            fonts = {role: (v['family'], v['size'], v['color']) for role, v in raw['fonts'].items()}
            composer._design_preset = DesignPreset(raw['name'], colors, fonts, dict(raw['spacing']), {})
            fill = composer._doc.SlideMaster.Background.Fill
            fill.ForeColor.RGB = hex_to_rgb_long(colors['bg'])
            fill.Solid()
        elif name in {'slide.add_title', 'slide.add_section', 'slide.add_bullets'}:
            layout = {'slide.add_title': composer.LAYOUT_TITLE, 'slide.add_section': composer.LAYOUT_SECTION,
                      'slide.add_bullets': composer.LAYOUT_TITLE_CONTENT}[name]
            slide, _ = composer._new_slide(layout)
            title = slide.Shapes.Title.TextFrame.TextRange
            title.Text = a['title']
            default = 40 if name == 'slide.add_title' else 32
            composer._apply_preset_font(title.Font, 'title', a.get('titleSize', default), a.get('titleColor'), default_size=default)
            body = a.get('subtitle') if name == 'slide.add_title' else a.get('items')
            if body:
                text_range = slide.Shapes.Placeholders(2).TextFrame.TextRange
                text_range.Text = body if isinstance(body, str) else '\r'.join(body)
                role, key, default = ('subtitle', 'subtitleSize', 20) if name == 'slide.add_title' else ('body', 'bodySize', 18)
                composer._apply_preset_font(text_range.Font, role, a.get(key, default), None, default_size=default)
        elif name == 'slide.add_blank':
            composer.add_blank_slide()
        elif name == 'slide.add_image':
            composer.add_image(a['slide'], resources[a['imageId']], a['left'], a['top'], a.get('width'), a.get('height'))
        elif name == 'slide.add_table':
            shape = composer.add_table(a['slide'], a['rows'], a['cols'], a['left'], a['top'], a['width'], a['height'],
                                       a['data'], header_shade=a.get('headerShade', '#4472C4'),
                                       header_font=a.get('headerFont', '#FFFFFF'), font_size=a.get('fontSize', 11))
            for column in range(1, a['cols'] + 1):
                shape.Table.Cell(1, column).Shape.Fill.ForeColor.RGB = hex_to_rgb_long(a.get('headerShade', '#4472C4'))
        else:
            raise NativeOfficeError('NATIVE_OFFICE_UNSUPPORTED')


def _execute_request(payload, job):
    from .windows_office_host import create_composer
    from .input_validation import validate_native_input
    component = payload.get('component')
    action = payload.get('action')
    if payload.get('protocol') != 1 or component not in _FORMATS or action not in {'generate', 'export'}:
        raise ValueError('Invalid native Office worker protocol')
    deadline = payload['deadline']
    _remaining(deadline)
    resources = {key: _owned_path(job, path) for key, path in payload.get('resources', {}).items()}
    if action == 'generate':
        plan = validate_plan(payload['plan'], resources)
        if plan.component != component:
            raise ValueError('Worker plan/component mismatch')
        target = job / ('native.' + _FORMATS[component])
    else:
        source = _owned_path(job, payload['source'])
        validate_native_input(source, component, deadline=deadline)
        target = job / 'export.pdf'
    composer = None
    error = None
    try:
        composer = create_composer(component, job)
        if action == 'generate':
            execute_plan(composer, plan, resources)
            _remaining(deadline)
            (composer.save_xlsx if component == 'spreadsheet' else composer.save_pptx)(target)
        else:
            composer.open_owned_document(source, read_only=True)
            composer.export_pdf(target)
    except BaseException as exc:
        error = exc
        raise
    finally:
        try:
            if composer is not None:
                composer.close(save_changes=False)
            if error is not None:
                error.cleanup_verified = composer is not None and composer._closed
        except BaseException as cleanup:
            if error is not None:
                raise cleanup from error
            raise
    _remaining(deadline)
    if action == 'generate':
        validate_office_package(target, _FORMATS[component])
    else:
        validate_pdf(target)
    return {'path': str(target), 'cleanup_verified': True, 'engine': 'msoffice',
            'component': component, 'executable': composer.executable, 'pid': composer.pid,
            'application_owned': composer._owns_app}


def _worker_main(request, response):
    request, response = Path(request).resolve(), Path(response).resolve()
    job = request.parent
    if request.name != 'request.json' or response != job / 'response.json':
        raise ValueError('Worker paths must be the private request/response pair')
    try:
        result = _execute_request(json.loads(request.read_text(encoding='utf-8')), job)
        _write_json(response, {'status': 'ok', 'value': result})
        _write_json(job / 'diagnostics.json', {'status': 'succeeded', 'cleanup_verified': True})
        return 0
    except BaseException as exc:
        _write_json(job / 'diagnostics.json', {'status': 'failed', 'error_type': type(exc).__name__,
                    'message': str(exc), 'traceback': traceback.format_exc(),
                    'cleanup_verified': getattr(exc, 'cleanup_verified', False),
                    'office_termination_attempted': False, 'public_code': getattr(exc, 'code', None)})
        return 1


if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[1] != '--worker':
        raise SystemExit('Only the private --worker protocol is supported')
    raise SystemExit(_worker_main(sys.argv[2], sys.argv[3]))
