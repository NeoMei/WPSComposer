"""Identity-bound Windows Microsoft document sessions.

File sessions edit a screened private copy. Live sessions retain the exact COM
document, never close it, and use non-rebinding save-copy primitives only.
Native acceptance is separate from these adapters' platform-independent tests.
"""
from __future__ import annotations

import hashlib
import inspect
import json
import math
import ntpath
import os
import re
from pathlib import Path
import shutil
import tempfile
import time
import traceback
from uuid import uuid4

from ..artifact_transport import (copy_file_before_deadline, publish_artifact,
    validate_office_package, validate_pdf, ValidatorSpec, validate_before_deadline)
from .input_validation import validate_native_input
from . import windows_host as word_host
from . import windows_office_host as office_host


_COMPONENT = {'writer': 'writer', 'sheet': 'spreadsheet', 'slide': 'presentation'}
_FORMAT = {'writer': 'docx', 'sheet': 'xlsx', 'slide': 'pptx'}
_APP = {'writer': ('Word.Application', 'Microsoft Word', 'WINWORD.EXE', 'Documents', 'ActiveDocument'),
        'sheet': ('Excel.Application', 'Microsoft Excel', 'EXCEL.EXE', 'Workbooks', 'ActiveWorkbook'),
        'slide': ('PowerPoint.Application', 'Microsoft PowerPoint', 'POWERPNT.EXE', 'Presentations', 'ActivePresentation')}


class DocumentIdentityError(RuntimeError):
    """An Office session can no longer prove its original document binding."""


def _path(value):
    return ntpath.normcase(ntpath.normpath(str(value)))


def _deadline_value(deadline=None):
    if deadline is None:
        return time.monotonic() + 600
    if type(deadline) not in (float, int) or not math.isfinite(deadline):
        raise ValueError('Session deadline must be finite')
    result = min(deadline, time.monotonic() + 600)
    if time.monotonic() >= result:
        raise TimeoutError('Microsoft session deadline expired')
    return result


def _require_deadline(deadline):
    if deadline is not None and time.monotonic() >= deadline:
        raise TimeoutError('Microsoft session deadline expired')


def _digest(path, deadline=None):
    _require_deadline(deadline)
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            _require_deadline(deadline)
            result.update(chunk)
    _require_deadline(deadline)
    return result.hexdigest()


def _create_owned_composer(component, root):
    if component == 'writer':
        composer = word_host.create_dedicated_composer(str(root))
        # Only the identity-verified dedicated Word process is affected.
        try:
            composer._app.AutomationSecurity = 3
        except BaseException:
            composer.close(save_changes=False)
            raise
        return composer
    return office_host.create_composer(component, root)


def _attach_dependencies(component):
    if component == 'writer':
        return word_host._load_dependencies()
    deps = office_host._load_dependencies(component)
    import win32com.client
    deps.active = lambda: win32com.client.GetActiveObject(office_host.APPS[component][0])
    return deps


def _attached_composer(kind, deps, app, doc):
    token = deps.identity(app)
    if kind == 'writer':
        pid = deps.window_pid(int(doc.Windows.Item(1).Hwnd))
    else:
        pid = office_host._application_pid(app, _COMPONENT[kind], deps)
    image = deps.process_image(pid)
    if (not image or ntpath.basename(image).casefold() != _APP[kind][2].casefold()
            or str(app.Name) != _APP[kind][1]
            or _path(app.Path) != _path(ntpath.dirname(image))
            or deps.identity(doc.Application) != token):
        raise DocumentIdentityError('Active document is not bound to the requested Microsoft executable')
    native = {'writer': word_host.NativeWordComposer, 'sheet': office_host.NativeSheetComposer,
              'slide': office_host.NativeSlideComposer}[kind]
    composer = native.__new__(native)
    composer._deps, composer._app, composer._doc = deps, app, doc
    composer._application_token, composer._document_token = token, deps.identity(doc)
    composer.pid, composer.executable = pid, image
    composer.identity = word_host.WordIdentity(pid, image)
    composer.component = _COMPONENT[kind]
    composer._collection = getattr(app, _APP[kind][3])
    composer._staging_dir = None
    composer._path = str(doc.FullName)
    composer._read_only = bool(doc.ReadOnly)
    composer._owns_app = composer._owns_doc = False
    composer._visible = True
    composer._ws = None
    composer._closed = False
    composer._first_section_configured = False
    composer.cleanup_error = None
    composer._com_initialized = True
    return composer


class _WindowsSession:
    engine = 'msoffice'
    kind = None

    def __init__(self, *args, **kwargs):
        raise TypeError('Use open_document or attach_active to establish a verified native binding')

    @classmethod
    def _initialize(cls, root, *, attached, read_only, deadline=None):
        session = cls.__new__(cls)
        session.staging_root = root
        session._attached = attached
        session._owns_doc = not attached
        session._read_only = bool(read_only)
        session._closed = session._failed = False
        session._composer = None
        session._source = session._source_digest = session._logical_path = None
        session._logical_digest = None
        session._deadline = _deadline_value(deadline)
        return session

    @property
    def publication_deadline(self):
        return self._deadline

    def _retain_error(self, error):
        self._failed = True
        self.staging_root.mkdir(mode=0o700, parents=True, exist_ok=True)
        record = self.staging_root / ('failure-' + uuid4().hex + '.json')
        with record.open('x', encoding='utf-8') as stream:
            os.chmod(record, 0o600)
            json.dump({'type': type(error).__name__, 'message': str(error),
                       'traceback': traceback.format_exc(), 'attached': self._attached,
                       'native_cleanup_verified': False}, stream, ensure_ascii=False)

    @classmethod
    def new_document(cls, *, visible=False, _deadline=None):
        """Use the identity-verified host's one new document, saved privately."""
        deadline = _deadline_value(_deadline)
        root = Path(tempfile.mkdtemp(prefix='wpscomposer-ms-session-'))
        os.chmod(root, 0o700)
        session = cls._initialize(root, attached=False, read_only=False, deadline=deadline)
        try:
            session._composer = _create_owned_composer(_COMPONENT[cls.kind], root)
            session._capture_binding()
            session._verify()
            private = root / ('new-' + uuid4().hex + '.' + _FORMAT[cls.kind])
            getattr(session._composer, 'save_' + _FORMAT[cls.kind])(private)
            validate_native_input(private, _COMPONENT[cls.kind], deadline=deadline)
            session._verify()
            if visible and session._composer._owns_app:
                session._composer._app.Visible = True
            return session
        except BaseException as exc:
            session._retain_error(exc)
            if session._composer is not None:
                try:
                    session._composer.close(save_changes=False)
                except BaseException as cleanup:
                    session._retain_error(cleanup)
            raise

    @classmethod
    def open_document(cls, path, *, read_only=False, visible=False, _deadline=None):
        source = Path(path).expanduser().resolve()
        deadline = _deadline_value(_deadline)
        validate_native_input(source, _COMPONENT[cls.kind], deadline=deadline)
        before = _digest(source, deadline)
        root = Path(tempfile.mkdtemp(prefix='wpscomposer-ms-session-'))
        os.chmod(root, 0o700)
        session = cls._initialize(root, attached=False, read_only=read_only, deadline=deadline)
        session._source, session._logical_path = source, source
        session._source_digest = session._logical_digest = before
        try:
            owned = root / ('input.' + _FORMAT[cls.kind])
            copy_file_before_deadline(source, owned, deadline=deadline)
            if _digest(owned, deadline) != before or _digest(source, deadline) != before:
                raise RuntimeError('Source changed while creating its private copy')
            validate_native_input(owned, _COMPONENT[cls.kind], deadline=deadline)
            session._composer = _create_owned_composer(_COMPONENT[cls.kind], root)
            session._composer.open_owned_document(owned, read_only=read_only)
            session._capture_binding()
            # Never change app visibility in an attached/shared process.
            if visible and session._composer._owns_app:
                session._composer._app.Visible = True
            session._verify()
            return session
        except BaseException as exc:
            session._retain_error(exc)
            if session._composer is not None:
                try:
                    session._composer.close(save_changes=False)
                except BaseException as cleanup:
                    session._retain_error(cleanup)
            raise

    @classmethod
    def attach_active(cls, *, _deadline=None):
        deadline = _deadline_value(_deadline)
        deps = _attach_dependencies(_COMPONENT[cls.kind])
        deps.initialize()
        session = None
        try:
            app = deps.active()
            if app is None:
                raise DocumentIdentityError('No registered Microsoft application')
            doc = getattr(app, _APP[cls.kind][4])
            if doc is None:
                raise DocumentIdentityError('Microsoft application has no active document')
            composer = _attached_composer(cls.kind, deps, app, doc)
            root = Path(tempfile.mkdtemp(prefix='wpscomposer-ms-live-'))
            os.chmod(root, 0o700)
            session = cls._initialize(root, attached=True, read_only=composer._read_only, deadline=deadline)
            session._composer = composer
            session._capture_binding()
            session._verify(selection=True)
            return session
        except BaseException as exc:
            if session is not None:
                session._retain_error(exc)
            deps.uninitialize()
            raise

    def _capture_binding(self):
        c = self._composer
        self._application_token = c._deps.identity(c._app)
        self._document_token = c._deps.identity(c._doc)

    def _verify(self, *, mutation=False, selection=False):
        _require_deadline(self._deadline)
        if self._closed or self._composer is None:
            raise DocumentIdentityError('Microsoft document session is closed')
        c = self._composer
        if (c._deps.identity(c._app) != self._application_token or
                c._deps.identity(c._doc) != self._document_token):
            raise DocumentIdentityError('Microsoft document session identity changed')
        if self.kind == 'writer':
            c._verify_document(c._doc)
        else:
            c._verify_document()
        if mutation and self._read_only:
            raise PermissionError('Read-only Microsoft session cannot mutate the document')
        if mutation:
            self._verify_live_macro_free()
        if selection:
            active = getattr(c._app, _APP[self.kind][4])
            if c._deps.identity(active) != self._document_token:
                raise DocumentIdentityError('Selection belongs to a different Microsoft document')
            if self.kind == 'sheet':
                selected_doc = c._app.Selection.Parent.Parent
            elif self.kind == 'slide':
                selected_doc = c._app.ActiveWindow.Presentation
            else:
                selected_doc = c._app.ActiveWindow.Selection.Document
            if c._deps.identity(selected_doc) != self._document_token:
                raise DocumentIdentityError('Active selection is not bound to the original document')

    def _verify_live_macro_free(self):
        if not self._attached:
            return
        doc = self._composer._doc
        try:
            macros = bool(doc.HasVBProject)
            if self.kind == 'sheet':
                macros = macros or bool(doc.Excel4MacroSheets.Count) or bool(doc.Excel4IntlMacroSheets.Count)
        except Exception as exc:
            raise PermissionError('Cannot verify live document macro status') from exc
        if macros:
            raise PermissionError('Live macro-bearing documents cannot be edited, saved, or exported')

    def __enter__(self):
        self._verify()
        return self

    def __exit__(self, *exc):
        if exc and exc[1] is not None:
            self._retain_error(exc[1])
        self.close(save_changes=False)

    def inspect_document(self, **options):
        self._verify()
        result = self._composer.inspect_document(**options)
        if self.kind == 'sheet' and isinstance(result, dict):
            # Shared SheetComposer reads Application.ActiveWindow for this one
            # field. Bind it back to this workbook's verified native window.
            try:
                frozen = bool(self._composer._doc.Windows.Item(1).FreezePanes)
            except Exception:
                frozen = None
            for sheet in result.get('sheets', []):
                if sheet.get('freeze_panes') is not None:
                    sheet['freeze_panes'] = frozen
        if not self._attached and isinstance(result, dict):
            result['path'] = str(self._logical_path) if self._logical_path is not None else None
        return result

    def inspect_selection(self):
        self._verify(selection=True)
        return self._composer.inspect_selection()

    def apply_format_patch(self, target, **patch):
        try:
            from .edit_preflight import classify_windows_set_op
            operation = {'op': 'set', 'target': target, **patch}
            if classify_windows_set_op(self.kind, operation) == 'unsupported':
                raise ValueError('Windows Microsoft set operation is unsupported')
            self._verify(mutation=True, selection=target == 'selection')
            self._screen_values(patch)
            return self._composer.apply_format_patch(target, **patch)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def _screen_values(self, values):
        if self.kind != 'sheet':
            return
        # Reuse the generation gate before assigning native Formula or Value;
        # DDE/automation formulas never reach an already-open user workbook.
        from .windows_office_runtime import validate_plan
        def strings(value):
            if isinstance(value, dict):
                for nested in value.values():
                    yield from strings(nested)
            elif isinstance(value, (list, tuple)):
                for nested in value:
                    yield from strings(nested)
            elif isinstance(value, str):
                yield value
        cells = [[s] for s in strings(values)]
        if cells:
            validate_plan({'component': 'spreadsheet', 'operations': [
                {'op': 'sheet.reset', 'args': {}}, {'op': 'sheet.write_table',
                'args': {'startRow': 1, 'startCol': 1, 'values': cells}}]}, {})

    def apply_structural_op(self, op):
        try:
            from .edit_preflight import classify_windows_structural_op
            if classify_windows_structural_op(
                    self.kind, op, engine='msoffice') == 'unsupported':
                raise ValueError('Windows Microsoft structural operation is unsupported')
            # Word shape move/clone uses the app clipboard selection internally.
            selection = self.kind == 'writer' and op.get('op') in {'move', 'clone'} and str(op.get('target', '')).startswith('shape:')
            self._verify(mutation=True, selection=selection)
            self._screen_values(op)
            if op.get('op') == 'insert' and op.get('type') == 'image':
                op = dict(op)
                op['props'] = dict(op.get('props') or {})
                op['props']['path'] = self._stage_resource(op['props'].get('path', ''))
            if self.kind == 'slide' and op.get('op') in {'move','clone'}:
                return self._powerpoint_relocate(op)
            if self.kind == 'sheet' and op.get('op') in {'move','clone'} and re.fullmatch(r'sheet:[1-9][0-9]*',str(op.get('target',''))):
                return self._excel_relocate_sheet(op)
            if self.kind == 'sheet' and op.get('op') == 'remove' and re.fullmatch(r'sheet:[1-9][0-9]*',str(op.get('target',''))):
                return self._excel_delete_sheet(op['target'])
            return self._composer.apply_structural_op(op)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def _excel_relocate_sheet(self, op):
        c = self._composer
        sheets = c._doc.Worksheets
        count = int(sheets.Count)
        index = int(op['target'].split(':')[1])
        if not 1 <= index <= count:
            raise ValueError('Native worksheet target is outside the workbook')
        source = sheets.Item(index)
        to, verb = op.get('to'), op['op']
        if to is None:
            anchor_index, key = (index if verb=='clone' else count), 'After'
        elif to == 'end':
            anchor_index,key = count,'After'
        elif to == 'start':
            anchor_index,key = 1,'Before'
        elif isinstance(to,dict) and set(to) in ({'before'},{'after'}):
            name = next(iter(to))
            anchor_index,key = to[name],name.title()
        else:
            raise ValueError('Native worksheet destination requires before/after or start/end')
        if isinstance(anchor_index,bool) or not isinstance(anchor_index,int) or not 1 <= anchor_index <= count:
            raise ValueError('Native worksheet destination is outside the workbook')
        anchor = sheets.Item(anchor_index)
        before = [c._deps.identity(sheets.Item(i)) for i in range(1,count+1)]
        if verb == 'move':
            if index != anchor_index:
                # Microsoft supports named optional arguments; never pass the
                # WPS-specific positional None, which becomes COM VT_NULL.
                source.Move(**{key:anchor})
            if int(sheets.Count) != count:
                raise DocumentIdentityError('Native worksheet move changed the workbook collection')
            result_sheet = source
        else:
            source.Copy(**{key:anchor})
            added = [sheets.Item(i) for i in range(1,int(sheets.Count)+1)
                     if c._deps.identity(sheets.Item(i)) not in before]
            if int(sheets.Count) != count+1 or len(added) != 1:
                raise DocumentIdentityError('Native worksheet clone did not create exactly one bound sheet')
            result_sheet = added[0]
        self._verify()
        return {'type':'sheet', 'moved' if verb=='move' else 'cloned':True,
                'from':op['target'], 'path':f'sheet:{int(result_sheet.Index)}'}

    def _excel_delete_sheet(self, target):
        """Never open an unbounded shared-app confirmation dialog."""
        c = self._composer
        sheets = c._doc.Worksheets
        index = int(target.split(':')[1])
        count = int(sheets.Count)
        if not 1 <= index <= count or count <= 1:
            raise ValueError('Worksheet deletion requires an existing sheet and a remaining sheet')
        sheet = sheets.Item(index)
        alerts = bool(c._app.DisplayAlerts)
        if alerts and not c._owns_app:
            raise RuntimeError('Shared Excel worksheet deletion would require an interactive confirmation')
        try:
            if alerts:
                c._app.DisplayAlerts = False
            result = sheet.Delete()
            if result is False or int(sheets.Count) != count - 1:
                raise RuntimeError('Native worksheet deletion was cancelled or not completed')
        finally:
            if alerts:
                c._app.DisplayAlerts = alerts
        self._verify()
        return {'removed':target}

    def _powerpoint_relocate(self, op):
        c = self._composer
        slides = c._doc.Slides
        count = int(slides.Count)
        target = str(op.get('target',''))
        match = re.fullmatch(r'slide:([1-9][0-9]*)(?:/shape:(.+))?', target)
        if not match or not 1 <= int(match[1]) <= count:
            raise ValueError('Native slide target is outside the presentation')
        source_index = int(match[1])
        source_slide = slides.Item(source_index)
        verb = op['op']
        if match[2] is not None:
            shape = c._resolve_shape(source_index,match[2])
            if shape is None:
                raise ValueError('Native shape target does not exist')
            to = op.get('to')
            if to in (None,'start','end'):
                destination = source_index
            elif isinstance(to,dict) and set(to)=={'slide'}:
                destination = to['slide']
            else:
                raise ValueError('Native shape destination requires a slide index')
            if isinstance(destination,bool) or not isinstance(destination,int) or not 1 <= destination <= count:
                raise ValueError('Native shape destination is outside the presentation')
            destination_slide = slides.Item(destination)
            if destination == source_index:
                if verb == 'clone':
                    shape.Duplicate()
                return {'type':'shape', 'moved' if verb=='move' else 'cloned':True, 'from':target, 'to_slide':destination}
            before_count = int(destination_slide.Shapes.Count)
            pasted = None
            try:
                # Native cross-slide transfer uses the clipboard. Never read or
                # restore the user's clipboard: another app may change it while
                # these COM calls run. Report the side effect explicitly.
                shape.Copy()
                pasted = destination_slide.Shapes.Paste()
                if int(pasted.Count) != 1 or int(destination_slide.Shapes.Count) != before_count+1:
                    raise RuntimeError('Native shape paste did not create exactly one copy')
                self._verify()
                if verb == 'move':
                    shape.Delete()
            except BaseException as exc:
                exc.clipboard_changed = True
                # The source remains intact through copy/paste failures.
                # Roll back only exact newly returned destination objects.
                try:
                    if pasted is not None:
                        for index in range(int(pasted.Count),0,-1):
                            pasted.Item(index).Delete()
                except BaseException as cleanup:
                    cleanup.clipboard_changed = True
                    raise cleanup from exc
                raise
            return {'type':'shape', 'moved' if verb=='move' else 'cloned':True,
                    'from':target, 'to_slide':destination, 'clipboard_changed':True}
        to = op.get('to','end')
        anchor = None
        if to in (None,'end'):
            destination = count + (1 if verb=='clone' else 0)
        elif to == 'start':
            destination = 1
        elif isinstance(to,dict) and set(to)=={'index'}:
            destination = to['index']
            if isinstance(destination,bool) or not isinstance(destination,int) or not 1 <= destination <= count+(verb=='clone'):
                raise ValueError('Native slide destination index is outside the presentation')
        elif isinstance(to,dict) and set(to) in ({'before'},{'after'}):
            key = next(iter(to))
            ref = re.fullmatch(r'slide:([1-9][0-9]*)',str(to[key]))
            if not ref or not 1 <= int(ref[1]) <= count:
                raise ValueError('Native slide anchor is outside the presentation')
            anchor = slides.Item(int(ref[1]))
            destination = int(ref[1])+(key=='after')
            if verb == 'move' and source_index < destination:
                destination -= 1
        else:
            raise ValueError('Unsupported native slide destination')
        if verb == 'move':
            if destination != source_index:
                # Microsoft uses the FINAL collection index. The inherited
                # WPS path adds an offset for its different MoveTo semantics.
                source_slide.MoveTo(destination)
        else:
            duplicate = source_slide.Duplicate().Item(1)
            try:
                if anchor is not None:
                    destination = anchor.SlideIndex+(key=='after')
                    if duplicate.SlideIndex < destination:
                        destination -= 1
                duplicate.MoveTo(destination)
            except BaseException:
                duplicate.Delete()
                raise
        self._verify()
        return {'type':'slide', 'moved' if verb=='move' else 'cloned':True, 'from':target}

    def _stage_resource(self, path):
        source = Path(path).expanduser().resolve()
        if not source.is_file():
            raise ValueError('Native image resource must be a local file')
        before = _digest(source, self._deadline)
        target = self.staging_root / ('resource-' + uuid4().hex + source.suffix)
        copy_file_before_deadline(source, target, deadline=self._deadline)
        if _digest(target, self._deadline) != before or _digest(source, self._deadline) != before:
            raise RuntimeError('Image resource changed while creating its private copy')
        return str(target)

    def _output(self, path, pdf=False):
        target = Path(path).expanduser().resolve()
        expected = '.pdf' if pdf else '.' + _FORMAT[self.kind]
        if target.suffix.lower() != expected:
            raise ValueError('Session output must use its native format, or explicit PDF export')
        return target

    def _publish(self, staged, target, *, pdf=False, replace_current=False):
        deadline = self._deadline
        if not pdf:
            validate_native_input(staged, _COMPONENT[self.kind], deadline=deadline)
        approved_digest = _digest(staged, deadline)
        spec = ValidatorSpec.from_callable(validate_pdf) if pdf else ValidatorSpec.from_callable(validate_office_package, _FORMAT[self.kind])
        def native_validator(path):
            validate_before_deadline(spec, path, deadline)
            # Transport's sibling temporary file deliberately has .tmp suffix;
            # validate its bytes against the screened native-format artifact.
            if _digest(path, deadline) != approved_digest:
                raise ValueError('Native artifact changed during publication')
            if replace_current and Path(path) != target:
                self._check_current_unchanged()
        result = str(publish_artifact(staged, target, overwrite=replace_current,
            validator=native_validator, deadline=deadline))
        if replace_current or not pdf:
            self._last_published_digest = approved_digest
        return result

    def _check_current_unchanged(self):
        if self._logical_path is None:
            raise ValueError('New document requires an explicit first save destination')
        try:
            unchanged = _digest(self._logical_path, self._deadline) == self._logical_digest
        except OSError:
            unchanged = False
        if not unchanged:
            self._failed = True
            raise RuntimeError('Source or saved output changed since this Microsoft session saved it')

    def _owned_save(self, target, *, pdf=False, replace_current=False):
        try:
            self._verify(mutation=not pdf)
            staged = self.staging_root / ('output-' + uuid4().hex + ('.pdf' if pdf else '.' + _FORMAT[self.kind]))
            method = 'export_pdf' if pdf else {'writer': 'save_docx', 'sheet': 'save_xlsx', 'slide': 'save_pptx'}[self.kind]
            getattr(self._composer, method)(staged)
            self._verify()
            return self._publish(staged, target, pdf=pdf, replace_current=replace_current)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def save(self, path, fmt=None):
        target = self._output(path)
        if target.exists():
            raise FileExistsError('Output already exists')
        if fmt is not None:
            raise ValueError('Explicit native format numbers are not accepted by bound sessions')
        if self._attached:
            return self.save_copy(target)
        try:
            result = self._owned_save(target)
            self._logical_path = target
            self._logical_digest = self._last_published_digest
            return result
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def save_current(self):
        self._verify(mutation=True)
        if self._attached:
            if not str(self._composer._doc.Path):
                raise ValueError('Unsaved attached document requires an explicit copy destination')
            self._composer._doc.Save()
            self._verify()
            return str(Path(self._composer._doc.FullName).resolve())
        self._check_current_unchanged()
        result = self._owned_save(self._logical_path, replace_current=True)
        self._logical_digest = self._last_published_digest
        if self._logical_path == self._source:
            self._source_digest = self._logical_digest
        return result

    def supports_attached_save_copy(self):
        return self.kind in {'sheet', 'slide'}

    def preflight_save(self, output=None, *, overwrite=False):
        """Check a complete live edit's save contract before any mutation."""
        self._verify(mutation=True)
        doc = self._composer._doc
        if bool(doc.ReadOnly):
            raise PermissionError('Read-only Microsoft document cannot be edited or saved')
        if output is None:
            if self._attached:
                if not str(doc.Path) or not Path(str(doc.FullName)).is_absolute():
                    raise ValueError('Live document has no already-saved native path')
                target = self._output(doc.FullName)
                if not target.is_file():
                    raise ValueError('Live document has no existing saved native file')
            else:
                self._check_current_unchanged()
                target = self._output(self._logical_path)
        else:
            target = self._output(output)
            if self._attached:
                if not self.supports_attached_save_copy():
                    raise RuntimeError('Word has no verified non-rebinding save-copy primitive')
                if self.is_bound_to(target):
                    raise ValueError('Save-copy destination must differ from the live document')
                if self.kind == 'sheet' and str(doc.Path):
                    self._output(doc.FullName)
            if target.exists() and not overwrite:
                raise FileExistsError('Save-copy destination already exists')
        if target.exists():
            if not target.is_file() or not os.access(target, os.W_OK):
                raise PermissionError('Native save destination is not writable')
        parent = target.parent
        while not parent.exists() and parent != parent.parent:
            parent = parent.parent
        if not parent.is_dir() or not os.access(parent, os.W_OK):
            raise PermissionError('Native save destination directory is not writable')
        return str(target)

    def save_copy(self, path, fmt=None):
        target = self._output(path)
        if target.exists():
            raise FileExistsError('Output already exists')
        if fmt is not None:
            raise ValueError('Explicit native format numbers are not accepted by bound sessions')
        if not self._attached:
            return self._owned_save(target)
        self._verify()
        self._verify_live_macro_free()
        if not self.supports_attached_save_copy():
            raise RuntimeError('Word has no verified non-rebinding save-copy primitive')
        if self.is_bound_to(target):
            raise ValueError('Save-copy destination must differ from the live document')
        doc = self._composer._doc
        before = (str(doc.FullName), bool(doc.Saved))
        staged = self.staging_root / ('copy-' + uuid4().hex + '.' + _FORMAT[self.kind])
        try:
            if self.kind == 'sheet':
                # Excel SaveCopyAs cannot convert an XLS/XLSM workbook to XLSX.
                if str(doc.Path) and Path(str(doc.FullName)).suffix.lower() != '.xlsx':
                    raise ValueError('Excel live copies require an XLSX source or unsaved workbook')
                doc.SaveCopyAs(str(staged))
            else:
                doc.SaveCopyAs(str(staged), 24)
            self._verify()
            if before != (str(doc.FullName), bool(doc.Saved)):
                raise DocumentIdentityError('Save-copy altered the live document binding or saved state')
            return self._publish(staged, target)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def export_pdf(self, path):
        target = self._output(path, pdf=True)
        if target.exists():
            raise FileExistsError('Output already exists')
        if not self._attached:
            return self._owned_save(target, pdf=True)
        self._verify()
        self._verify_live_macro_free()
        doc = self._composer._doc
        before = (str(doc.FullName), bool(doc.Saved))
        staged = self.staging_root / ('export-' + uuid4().hex + '.pdf')
        try:
            if self.kind == 'writer':
                doc.ExportAsFixedFormat(OutputFileName=str(staged), ExportFormat=17, OpenAfterExport=False)
            elif self.kind == 'sheet':
                doc.ExportAsFixedFormat(Type=0, Filename=str(staged), OpenAfterPublish=False)
            else:
                doc.ExportAsFixedFormat(Path=str(staged), FixedFormatType=2)
            self._verify()
            if before != (str(doc.FullName), bool(doc.Saved)):
                raise DocumentIdentityError('PDF export changed the live document binding or saved state')
            return self._publish(staged, target, pdf=True)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def is_bound_to(self, path):
        self._verify()
        current = self._composer._doc.FullName if self._attached else self._logical_path
        return _path(Path(current).expanduser().resolve()) == _path(Path(path).expanduser().resolve())

    def close(self, save_changes=False):
        if self._closed:
            return
        try:
            if self._attached:
                # Release references only. Save, close and Quit are separate,
                # explicit operations; even save_changes=True cannot save live.
                self._composer._deps.uninitialize()
                self._composer._com_initialized = False
            else:
                if save_changes:
                    self.save_current()
                self._composer.close(save_changes=False)
            self._closed = True
            self._composer = None
            if not self._failed:
                shutil.rmtree(self.staging_root)
        except BaseException as exc:
            self._retain_error(exc)
            raise

    def __getattr__(self, name):
        # Expose existing Composer business methods while preventing alternate
        # lifecycle entrypoints from re-dispatching outside this binding.
        if name == 'save_' + _FORMAT[self.kind]:
            return self.save
        if name.startswith('_') or name.startswith(('open', 'attach', 'save', 'close')):
            raise AttributeError(name)
        composer = object.__getattribute__(self, '_composer')
        self._verify()
        value = getattr(composer, name)
        if not callable(value):
            return value
        def bound(*args, **kwargs):
            try:
                signature = inspect.signature(value)
                call = signature.bind(*args, **kwargs)
                read = (
                    name.startswith(('inspect', 'snapshot', 'get_'))
                    or name in {
                        'degradation_checkpoint',
                        'pagination_fragment_for_bookmark',
                    }
                )
                self._verify(mutation=not read, selection=not read)
                self._screen_values((call.args, call.kwargs))
                if name in {'add_image', 'add_image_block'}:
                    image_args, image_kwargs = list(args), dict(kwargs)
                    if 'path' in image_kwargs:
                        image_kwargs['path'] = self._stage_resource(
                            image_kwargs['path']
                        )
                    else:
                        index = 1 if self.kind == 'slide' else 0
                        image_args[index] = self._stage_resource(image_args[index])
                    call = signature.bind(*image_args, **image_kwargs)
                if name in {
                    'add_captioned_figure_native',
                    'add_captioned_figure_fallback',
                }:
                    locators = call.arguments.get('resource_locators')
                    if locators is not None:
                        if (not isinstance(locators, dict) or
                                any(type(key) is not str for key in locators) or
                                any(not isinstance(path, (str, os.PathLike)) for path in locators.values())):
                            raise ValueError('Invalid resource locator contract')
                        resolved = []
                        for key, path in locators.items():
                            source = Path(path).expanduser().resolve()
                            if not source.is_file():
                                raise ValueError('Native image resource must be a local file')
                            resolved.append((key, source))
                        call.arguments['resource_locators'] = {
                            key: self._stage_resource(path)
                            for key, path in resolved
                        }
                if name in {
                    'add_equation_native',
                    'add_equation_native_fallback',
                }:
                    locator = call.arguments.get('fallback_resource_locator')
                    if locator is not None:
                        if not isinstance(locator, (str, os.PathLike)):
                            raise ValueError('Invalid fallback resource locator contract')
                        call.arguments['fallback_resource_locator'] = (
                            self._stage_resource(locator)
                        )
                return value(*call.args, **call.kwargs)
            except BaseException as exc:
                self._retain_error(exc)
                raise
        return bound


class WindowsWordSession(_WindowsSession):
    kind = 'writer'


class WindowsExcelSession(_WindowsSession):
    kind = 'sheet'


class WindowsPowerPointSession(_WindowsSession):
    kind = 'slide'
