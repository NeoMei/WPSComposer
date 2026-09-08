"""Closed JSON protocol; one Windows COM apartment/session per Python child.

This module imports no COM dependencies until the first validated open request.
Only JSON values cross the boundary, never native objects or attribute paths.
"""
from __future__ import annotations

import contextlib
import json
import math
import os
from pathlib import Path
import sys
import tempfile
import time
import traceback
from uuid import uuid4

PROTOCOL = 1
MAX_FRAME_BYTES = 32 * 1024 * 1024
COMMON_METHODS = frozenset({
    'inspect_document', 'inspect_selection', 'apply_format_patch',
    'apply_structural_op', 'preflight_save', 'save', 'save_copy', 'save_current',
    'export_pdf', 'is_bound_to', 'supports_attached_save_copy', 'close',
})
# These existing business methods have value-only inputs and return values.
BUSINESS_METHODS = {
    'writer': frozenset({'ensure_styles', 'ensure_heading_styles', 'apply_heading_text_color',
        'add_heading', 'add_heading2', 'add_heading_level', 'add_paragraph', 'add_centered',
        'add_bullet_list', 'add_numbered_list', 'add_styled_paragraph', 'add_rich_paragraph',
        'add_code_lines', 'set_columns', 'set_orientation', 'set_margins', 'set_page_size',
        'add_section', 'set_header', 'set_footer', 'set_page_number_in_footer',
        'add_table', 'add_floating_textbox', 'add_wordart', 'add_image', 'add_image_block',
        'set_page_role', 'set_document_metadata', 'set_page_numbering', 'set_header_footer',
        'configure_section', 'add_page_break', 'add_horizontal_line', 'add_paragraph_horizontal_line',
        'compact_terminal_paragraph', 'update_fields', 'refresh_fields', 'refresh_indexes',
        'repaginate_and_update_numbering', 'refresh_bookmarks_and_references',
        'repaginate_and_update_page_fields', 'snapshot_fields', 'save_docx'}),
    'sheet': frozenset({'select_sheet', 'rename_sheet', 'add_sheet', 'add_chart', 'write_cell', 'set_formula',
        'write_table', 'set_cell_style', 'set_range_style', 'set_borders', 'merge_cells',
        'freeze_panes', 'set_column_width', 'set_row_height', 'autofit', 'add_title_row',
        'conditional_format', 'set_header_footer', 'save_xlsx'}),
    'slide': frozenset({'set_slide_size', 'add_title_slide', 'add_section_slide',
        'add_text_slide', 'add_bullets_slide', 'set_background_color', 'set_notes', 'save_pptx',
        'add_blank_slide', 'add_textbox', 'add_shape', 'add_image', 'add_table'}),
}
GETTERS = {'writer': frozenset(), 'sheet': frozenset(), 'slide': frozenset({'slide_count'})}
UNSUPPORTED_BUSINESS = frozenset({'apply_design_preset', 'apply_layout_template'})
HANDLE_TYPES = frozenset({'slide', 'sheet', 'shape', 'table', 'inline_shape', 'chart', 'range'})
_HANDLE_TAG = '__wpscomposer_handle__'
_TUPLE_TAG = '__wpscomposer_tuple__'


class NativeHandleRegistry:
    """An opaque reference is usable only while its object remains in this doc."""
    def __init__(self, session):
        self.session = session
        self.records = {}

    @staticmethod
    def _items(collection):
        for index in range(1, int(collection.Count) + 1):
            yield index, collection.Item(index)

    def _location(self, native, kind):
        self.session._verify()
        c = self.session._composer
        identity = c._deps.identity
        token = identity(native)
        document = c._doc
        family = self.session.kind
        if family == 'writer' and kind in {'table', 'shape', 'inline_shape'}:
            collection = getattr(document, {'table':'Tables', 'shape':'Shapes', 'inline_shape':'InlineShapes'}[kind])
            for index, item in self._items(collection):
                if identity(item) == token: return f'{kind}:{index}', index
        elif family == 'slide':
            for index, slide in self._items(document.Slides):
                if kind == 'slide' and identity(slide) == token: return f'slide:{index}', index
                if kind == 'shape':
                    for shape_index, shape in self._items(slide.Shapes):
                        if identity(shape) == token: return f'slide:{index}/shape:{shape_index}', shape_index
        elif family == 'sheet':
            for index, sheet in self._items(document.Worksheets):
                if kind == 'sheet' and identity(sheet) == token: return f'sheet:{index}', index
                if kind in {'shape', 'chart'}:
                    collection = sheet.Shapes if kind == 'shape' else sheet.ChartObjects()
                    for object_index, item in self._items(collection):
                        if identity(item) == token: return f'sheet:{index}/{kind}:{object_index}', object_index
                if kind == 'range' and identity(native.Parent) == identity(sheet):
                    return f'sheet:{index}/range:{native.Address}', str(native.Address)
        if family == 'writer' and kind == 'range' and identity(native.Document) == identity(document):
            start, end = int(native.Start), int(native.End)
            if 0 <= start <= end <= int(document.Content.End): return f'range:{start}-{end}', (start, end)
        raise ValueError('Native handle is stale or no longer belongs to this session')

    def register(self, native, kind):
        if kind not in HANDLE_TYPES: raise ValueError('Unsupported native handle type')
        if len(self.records) >= 1024: raise ValueError('Native session handle limit exceeded')
        self._location(native, kind)
        token = uuid4().hex
        self.records[token] = (native, kind)
        return {_HANDLE_TAG: {'id': token, 'type': kind}}

    def _reference(self, value):
        if not isinstance(value, dict) or set(value) != {_HANDLE_TAG}: return None
        reference = value[_HANDLE_TAG]
        if not isinstance(reference, dict) or set(reference) != {'id', 'type'} or not isinstance(reference['id'], str):
            raise ValueError('Malformed native handle')
        record = self.records.get(reference['id'])
        if record is None: raise ValueError('Native handle belongs to another session or has expired')
        native, kind = record
        if kind != reference['type']: raise ValueError('Native handle type mismatch')
        path, index = self._location(native, kind)
        return native, kind, path, index

    def _value(self, value, role=None):
        reference = self._reference(value)
        if reference:
            native, kind, path, index = reference
            if role in {'target', 'parent'}: return path
            if role == 'anchor':
                family = self.session.kind
                if family == 'sheet' and kind == 'sheet': return index
                if family == 'slide' and kind == 'slide': return path
                if family == 'writer' and kind == 'range':
                    for paragraph_index, paragraph in self._items(self.session._composer._doc.Paragraphs):
                        candidate = paragraph.Range
                        if (int(candidate.Start), int(candidate.End)) == index:
                            return f'paragraph:{paragraph_index}'
                    raise ValueError('Word anchor handle must cover exactly one complete paragraph')
                raise ValueError('Native handle has the wrong anchor type')
            if role in {'slide_index', 'sheet_index'}:
                if kind != role.split('_')[0]: raise ValueError('Native handle has the wrong argument type')
                return index
            if role == 'source_range':
                if kind != 'range': raise ValueError('Native handle has the wrong argument type')
                return native
            if role == 'range_str':
                if kind != 'range' or self.session.kind != 'sheet': raise ValueError('Native handle has the wrong argument type')
                return index
            raise ValueError('Native handle is not supported for this argument type')
        if isinstance(value, list): return [self._value(item) for item in value]
        if isinstance(value, dict):
            nested_roles = {'target':'target', 'parent':'parent', 'before':'anchor',
                            'after':'anchor', 'slide':'slide_index'}
            return {key: self._value(item, nested_roles.get(key)) for key,item in value.items()}
        return value

    def prepare_call(self, method, args, kwargs):
        if len(self.records) >= 1024 and method in {'add_table','add_floating_textbox','add_wordart',
                'add_image','add_image_block','select_sheet','rename_sheet','add_sheet','add_chart',
                'add_blank_slide','add_textbox','add_shape'}:
            raise ValueError('Native session handle limit exceeded before mutation')
        roles = {}
        if method == 'apply_format_patch': roles = {0:'target', 'target':'target'}
        elif self.session.kind == 'slide' and method in {'add_textbox','add_shape','add_image','add_table','set_background_color','set_notes'}:
            roles = {0:'slide_index', 'slide_index':'slide_index'}
        elif self.session.kind == 'sheet' and method in {'select_sheet','rename_sheet'}:
            roles = {0:'sheet_index', 'index':'sheet_index'}
        elif self.session.kind == 'sheet' and method in {'set_range_style','set_borders','merge_cells','add_title_row','conditional_format'}:
            roles = {0:'range_str', 'range_str':'range_str'}
        elif self.session.kind == 'sheet' and method == 'add_chart':
            roles = {5:'source_range', 'source_range':'source_range'}
        return ([self._value(value, roles.get(index)) for index,value in enumerate(args)],
                {key:self._value(value, roles.get(key)) for key,value in kwargs.items()})

    def encode_result(self, method, value, args, kwargs):
        family = self.session.kind
        kind = None
        if family == 'writer':
            kind = {'add_table':'table', 'add_floating_textbox':'shape', 'add_wordart':'shape'}.get(method)
            if method in {'add_image','add_image_block'}: kind = 'inline_shape' if kwargs.get('inline', True) else 'shape'
        elif family == 'sheet':
            if method in {'select_sheet','rename_sheet','add_sheet'}: kind = 'sheet'
            elif method == 'add_chart': kind = 'chart' if kwargs.get('source_range', args[5] if len(args)>5 else None) else 'shape'
        elif family == 'slide':
            if method == 'add_blank_slide':
                if not isinstance(value, tuple) or len(value) != 2 or type(value[1]) is not int:
                    raise ValueError('Native blank-slide result contract changed')
                return {_TUPLE_TAG: [self.register(value[0], 'slide'), value[1]]}
            if method in {'add_textbox','add_shape','add_image','add_table'}: kind = 'shape'
        return self.register(value, kind) if kind else value


def encode_frame(value):
    raw = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(',', ':')).encode('utf-8') + b'\n'
    if len(raw) > MAX_FRAME_BYTES:
        raise ValueError('Session protocol frame exceeds size limit')
    return raw


def decode_frame(raw):
    if not raw or len(raw) > MAX_FRAME_BYTES or not raw.endswith(b'\n'):
        raise ValueError('Session protocol frame is incomplete or oversized')
    return json.loads(raw, parse_constant=lambda value: (_ for _ in ()).throw(ValueError('Non-finite JSON value')))


def default_factory(kind, method, args, kwargs, *, deadline=None):
    from .windows_document_api import WindowsWordSession, WindowsExcelSession, WindowsPowerPointSession
    cls = {'writer': WindowsWordSession, 'sheet': WindowsExcelSession, 'slide': WindowsPowerPointSession}[kind]
    return getattr(cls, method)(*args, **dict(kwargs, _deadline=deadline))


def _identity(session):
    # Reading only already-captured Python fields cannot introduce a new COM
    # call while trying to report a hung/failed native command.
    state = vars(session)
    composer = state.get('_composer')
    native = vars(composer) if composer is not None and hasattr(composer, '__dict__') else {}
    cached_identity = native.get('identity')
    cached = vars(cached_identity) if cached_identity is not None and hasattr(cached_identity, '__dict__') else {}
    return {'staging_path': str(state.get('staging_root', '')),
            'office_pid': native.get('pid', cached.get('pid')),
            'office_executable': native.get('executable', cached.get('executable')),
            'attached': state.get('_attached', False)}


def _error(exc):
    kind = type(exc).__name__
    if isinstance(exc, PermissionError): kind = 'PermissionError'
    elif isinstance(exc, TimeoutError): kind = 'TimeoutError'
    elif isinstance(exc, FileExistsError): kind = 'FileExistsError'
    elif isinstance(exc, NotImplementedError): kind = 'NotImplementedError'
    elif isinstance(exc, ValueError): kind = 'ValueError'
    elif isinstance(exc, TypeError): kind = 'TypeError'
    elif isinstance(exc, OSError): kind = 'OSError'
    elif not hasattr(exc, 'code'): kind = 'RuntimeError'
    result = {'type': kind, 'message': str(exc), 'code': getattr(exc, 'code', None)}
    changed = getattr(exc, 'clipboard_changed', None)
    if type(changed) is bool:
        result['clipboard_changed'] = changed
    return result


def serve(incoming, outgoing, job, *, factory=default_factory):
    """Run the actual protocol; injectable factory is for COM-free tests only."""
    session = None
    registry = None
    last_id = 0
    deadline = None
    client_deadline = None
    kind = None
    closed = False
    identity = {}
    job = Path(job)
    try:
        while True:
            raw = incoming.readline(MAX_FRAME_BYTES + 1)
            if not raw:
                break
            response = {'protocol': PROTOCOL, 'id': None, 'status': 'error'}
            try:
                request = decode_frame(raw)
                if not isinstance(request, dict) or set(request) != {'protocol', 'id', 'kind', 'method', 'args', 'kwargs', 'deadline', 'remaining_seconds'}:
                    raise ValueError('Invalid session request envelope')
                response['id'] = request['id']
                if type(request['protocol']) is not int or request['protocol'] != PROTOCOL or type(request['id']) is not int or request['id'] != last_id + 1:
                    raise ValueError('Invalid session request sequence')
                last_id = request['id']
                requested_deadline = request['deadline']
                if type(requested_deadline) not in (int, float) or not math.isfinite(requested_deadline):
                    raise ValueError('Invalid session deadline')
                budget = request['remaining_seconds']
                if type(budget) not in (int, float) or not math.isfinite(budget) or not 0 < budget <= 600:
                    raise ValueError('Invalid remaining session budget')
                if deadline is None:
                    # Older Python/macOS monotonic clocks have process-local
                    # origins. Translate the remaining budget once, then only
                    # shorten this child's absolute deadline on later calls.
                    deadline = time.monotonic() + budget
                    client_deadline = requested_deadline
                    kind = request['kind']
                elif requested_deadline > client_deadline:
                    raise ValueError('A session cannot extend its deadline')
                client_deadline = min(client_deadline, requested_deadline)
                deadline = min(deadline, time.monotonic() + budget)
                if session is not None:
                    session._deadline = deadline
                if time.monotonic() >= deadline:
                    raise TimeoutError('Session deadline expired')
                if kind not in BUSINESS_METHODS or request['kind'] != kind:
                    raise ValueError('Invalid or changed session component')
                method, args, kwargs = request['method'], request['args'], request['kwargs']
                if not isinstance(method, str) or not isinstance(args, list) or not isinstance(kwargs, dict):
                    raise ValueError('Invalid session call')
                if session is None:
                    if method not in {'open_document', 'attach_active', 'new_document'}:
                        raise ValueError('The first request must establish a document binding')
                    with contextlib.redirect_stdout(sys.stderr):
                        if factory is default_factory:
                            session = factory(kind, method, args, kwargs, deadline=deadline)
                        else:
                            session = factory(kind, method, args, kwargs)
                        session._deadline = deadline
                        session.kind = kind
                    registry = NativeHandleRegistry(session)
                    value = {'kind': kind}
                else:
                    allowed = COMMON_METHODS | BUSINESS_METHODS[kind]
                    if method == 'get_property':
                        if len(args) != 1 or args[0] not in GETTERS[kind] or kwargs:
                            raise ValueError('Unsupported native property')
                        value = getattr(session, args[0])
                    elif method not in allowed:
                        raise NotImplementedError('Native method has no value-only session contract')
                    else:
                        args, kwargs = registry.prepare_call(method, args, kwargs)
                        if method == 'close': registry.records.clear()
                        with contextlib.redirect_stdout(sys.stderr):
                            value = getattr(session, method)(*args, **kwargs)
                        value = registry.encode_result(method, value, args, kwargs)
                        if method == 'close':
                            value = {'closed': True}
                            closed = True
                identity.update({key: value for key, value in _identity(session).items() if value is not None})
                (job / 'office-identity.json').write_text(json.dumps(identity), encoding='utf-8')
                response.update(status='ok', value=value, identity=identity)
                # Encode before acknowledging; arbitrary COM objects never get a
                # repr/default serializer that could hide an invalid contract.
                frame = encode_frame(response)
            except BaseException as exc:
                if session is not None:
                    try:
                        retain = object.__getattribute__(session, '_retain_error')
                    except AttributeError:
                        pass
                    else:
                        retain(exc)
                (job / 'worker-error.log').write_text(traceback.format_exc(), encoding='utf-8')
                response.update(status='error', error=_error(exc))
                response.pop('value', None)
                frame = encode_frame(response)
            outgoing.write(frame)
            outgoing.flush()
            if closed:
                break
    finally:
        if session is not None and not closed:
            # Parent EOF is a request to release this one session. A stuck close
            # remains subject to the parent's exact-child kill, never Office kill.
            with contextlib.redirect_stdout(sys.stderr):
                if registry is not None: registry.records.clear()
                session.close(save_changes=False)


def main():
    if len(sys.argv) != 2:
        raise SystemExit('Usage: windows_session_worker PRIVATE_JOB')
    job = Path(sys.argv[1]).resolve(strict=True)
    if not job.is_dir():
        raise SystemExit('Private job directory is missing')
    tempfile.tempdir = str(job)
    for key in ('TMP', 'TEMP', 'TMPDIR'):
        os.environ[key] = str(job)
    serve(sys.stdin.buffer, sys.stdout.buffer, job)


if __name__ == '__main__':
    main()
