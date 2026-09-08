"""Bound Microsoft Word editing sessions using the macOS scripting dictionary.

File sessions operate on private copies and verify their unique path on every
AppleEvent. Attached sessions require a native window ID and never close their
document. Word versions returning no ID reject attachment before mutation.
"""
from __future__ import annotations

import json
import hashlib
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
from uuid import uuid4

from ..artifact_transport import copy_file_before_deadline, publish_artifact, validate_office_package, validate_pdf, validate_before_deadline, ValidatorSpec
from .errors import NativeWordCapabilityError, NativeWordError, NativeWordTimeoutError
from .input_validation import validate_native_input
from .macos_runtime import WordJobLock, remaining
from .macos_script import apple_string

WORD_APP = Path('/Applications/Microsoft Word.app')
OPERATION_TIMEOUT = 60.0
_COMPLETION_MARKER = 'WPSCOMPOSER_WORD_SESSION_OK'


def _temporary_root():
    return Path.home() / 'Library/Containers/com.microsoft.Word/Data/tmp'


_JSON = '''use framework "Foundation"
use scripting additions
on jsonRows(rows)
  set dataValue to current application's NSJSONSerialization's dataWithJSONObject:rows options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:dataValue encoding:4) as text
end jsonRows
on enumIndex(v, choices)
  repeat with i from 1 to count choices
    if v is item i of choices then return i - 1
  end repeat
  return -1
end enumIndex
'''

_ALIGN = ['align paragraph left', 'align paragraph center', 'align paragraph right', 'align paragraph justify', 'align paragraph distribute']
_SPACING = ['line space single', 'line space1 pt5', 'line space double', 'line space at least', 'line space exactly', 'line space multiple']
_FONT = {'name': 'name', 'size': 'font size', 'bold': 'bold', 'italic': 'italic', 'underline': 'underline', 'strikethrough': 'strike through', 'color': 'color'}
_PARA = {'alignment': 'alignment', 'left_indent': 'paragraph format left indent', 'right_indent': 'paragraph format right indent', 'first_line_indent': 'first line indent', 'space_before': 'space before', 'space_after': 'space after', 'line_spacing': 'line spacing', 'line_spacing_rule': 'line spacing rule', 'keep_together': 'keep together', 'keep_with_next': 'keep with next', 'page_break_before': 'page break before', 'widow_control': 'widow control'}
_PAGE = {k: k.replace('_', ' ') for k in ('orientation', 'page_width', 'page_height', 'top_margin', 'bottom_margin', 'left_margin', 'right_margin', 'header_distance', 'footer_distance', 'gutter')}
_GEOMETRY = {'left': 'left position', 'top': 'top', 'width': 'width', 'height': 'height', 'rotation': 'rotation'}
_WRAP = ['wrap square', 'wrap tight', 'wrap through', 'wrap none', 'wrap top bottom', 'wrap behind', 'wrap front', 'wrap inline']
_SHAPE_TYPES = ['unset', 'auto', 'callout', 'chart', 'comment', 'free form', 'group', 'embedded OLE control', 'form control', 'line', 'linked OLE object', 'linked picture', 'OLE control', 'picture', 'place holder', 'word art', 'media', 'text box', 'script anchor', 'table', 'canvas', 'diagram', 'ink', 'ink comment', 'smartart graphic', 'slicer', 'web video', 'content application', 'graphic', 'linked graphic']


def _number(value):
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value):
        raise ValueError('Expected finite number')
    return str(value)


def _value(key, value):
    if key in ('color', 'back_color'):
        if isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= 0xffffff:
            rgb = [value & 255, (value >> 8) & 255, (value >> 16) & 255]
        elif isinstance(value, str) and re.fullmatch(r'#[0-9a-fA-F]{6}', value):
            rgb = [int(value[i:i+2], 16) for i in (1, 3, 5)]
        else:
            raise ValueError('Invalid color')
        return '{' + ', '.join(str(c * 257) for c in rgb) + '}'
    if key in ('name', 'text', 'style'):
        if not isinstance(value, str):
            raise ValueError('Expected text')
        return apple_string(value)
    if key in ('bold', 'italic', 'strikethrough', 'keep_together', 'keep_with_next', 'page_break_before', 'widow_control', 'visible'):
        if value not in (True, False, -1, 0, 1):
            raise ValueError('Expected boolean')
        return 'true' if value else 'false'
    enum = {'alignment': _ALIGN, 'line_spacing_rule': _SPACING, 'orientation': ['orient portrait', 'orient landscape'], 'vertical_alignment': ['cell align vertical top', 'cell align vertical center', 'cell align vertical bottom']}.get(key)
    if enum:
        if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value < len(enum):
            raise ValueError('Unsupported enumeration')
        return enum[value]
    if key == 'underline':
        if value not in (0, 1):
            raise ValueError('Unsupported underline style')
        return 'underline single' if value else 'underline none'
    rendered = _number(value)
    if key in ('size', 'width', 'height', 'page_width', 'page_height') and value <= 0:
        raise ValueError('Expected positive dimension')
    if key in ('weight', 'space_before', 'space_after', 'line_spacing', 'top_margin', 'bottom_margin', 'left_margin', 'right_margin', 'header_distance', 'footer_distance', 'gutter') and value < 0:
        raise ValueError('Expected nonnegative dimension')
    if key == 'transparency' and not 0 <= value <= 1:
        raise ValueError('Expected transparency between zero and one')
    return rendered


def _row(category, index, key, expression, *, optional=False):
    line = f'set end of nativeRows to {{{apple_string(category)}, {index}, {apple_string(key)}, {expression}}}'
    return ['try', line, 'end try'] if optional else [line]


def _decode_snapshot(rows):
    result = {'kind': 'writer', 'counts': {}, 'paragraphs': [], 'tables': [], 'shapes': [], 'sections': []}
    objects = {}
    for category, index, key, value in rows:
        if category == 'doc':
            obj = result
        elif category == 'cells':
            table, row, column = [int(part) for part in index.split(',')]
            parent = objects[('tables', table)]
            token = (category, index)
            if token not in objects:
                objects[token] = {'id': f'table:{table}/cell:{row},{column}', 'row': row, 'column': column}
                parent.setdefault('cells', []).append(objects[token])
            obj = objects[token]
        else:
            token = (category, index)
            if token not in objects:
                singular = {'paragraphs': 'paragraph', 'tables': 'table', 'shapes': 'shape', 'sections': 'section', 'selection': 'selection'}[category]
                objects[token] = {'id': f'{singular}:{index}' if singular != 'selection' else 'selection', 'index': index}
                result.setdefault(category, []).append(objects[token])
            obj = objects[token]
        parts = key.split('.')
        for part in parts[:-1]:
            obj = obj.setdefault(part, {})
        if parts[-1] == 'text' and isinstance(value, str):
            value = value.rstrip('\r\x07')
        if parts[-1] == 'color' and isinstance(value, list) and len(value) == 3:
            value = '#' + ''.join(f'{min(255, max(0, round(c / 257))):02X}' for c in value)
        obj[parts[-1]] = value
    return result


class MacWordSession:
    kind = 'writer'
    engine = 'msoffice'

    def __init__(self):
        self._deadline = time.monotonic() + 600
        self.staging_root = None
        self.lock = None
        self._window_id = None
        self._bound_path = None
        self._bound_name = None
        self._source_path = None
        self._source_digest = None
        self._private_path = None
        self._owns_doc = False
        self._read_only = False
        self._closed = False
        self._quarantined = False
        self._structural_changed = False
        self._retain_evidence = False

    def _prepare(self):
        if sys.platform != 'darwin' or not WORD_APP.is_dir() or not _temporary_root().is_dir():
            raise NativeWordError('NATIVE_WORD_UNAVAILABLE')
        self.lock = WordJobLock(_temporary_root() / 'wpscomposer-native-word.lock')
        self.lock.acquire(self._deadline)
        try:
            self.staging_root = Path(tempfile.mkdtemp(prefix='wpscomposer-session-', dir=_temporary_root()))
            os.chmod(self.staging_root, 0o700)
        except BaseException:
            self.lock.close()
            raise

    @classmethod
    def open_document(cls, path, *, read_only=False, visible=False):
        source = Path(path).expanduser().resolve(strict=True)
        self = cls()
        validate_native_input(source, 'writer', deadline=self._deadline)
        self._source_path, self._owns_doc, self._read_only = source, True, bool(read_only)
        self._prepare()
        target = self.staging_root / ('document-' + uuid4().hex + '.docx')
        self._private_path = target
        activation_attempted = False
        try:
            self._source_digest = self._digest(source)
            copy_file_before_deadline(source, target, deadline=self._deadline)
            validate_native_input(target, 'writer', deadline=self._deadline)
            if self._digest(target) != self._source_digest:
                raise ValueError('Source changed while preparing private document')
            self._check_source_unchanged()
            activation_attempted = True
            rows = self._execute([
                '-- WPSC_BIND_OPEN',
                f'open file name {apple_string(str(target))} read only {str(bool(read_only)).lower()} add to recent files false',
                'set matches to {}',
                'repeat with di from 1 to (count of documents)',
                'set d to document di',
                f'if (posix full name of d as text) is {apple_string(str(target))} then set end of matches to d',
                'end repeat',
                'if (count matches) is not 1 then error "WPSC_BINDING_FAILED"',
                'set boundDoc to item 1 of matches',
                'set boundWindow to active window of boundDoc',
                'set wid to id of boundWindow',
                *(['set visible of boundWindow to true'] if visible else []),
                'set nativeRows to {{"binding", wid, posix full name of boundDoc as text}}',
            ], bind=False)
            self._bind(rows)
            return self
        except BaseException:
            if activation_attempted:
                self._retain('Open completion or document binding unverified')
            else:
                shutil.rmtree(self.staging_root)
            self.lock.close()
            raise

    @classmethod
    def new_document(cls, *, visible=False):
        self = cls()
        self._owns_doc = True
        self._prepare()
        self._private_path = self.staging_root / ('document-' + uuid4().hex + '.docx')
        try:
            rows = self._execute([
                '-- WPSC_BIND_NEW',
                'set previousNames to name of every document',
                'set boundDoc to make new document',
                'set newName to name of boundDoc',
                'if previousNames contains newName then error "WPSC_NEW_IDENTITY_COLLISION"',
                'set boundDoc to document newName',
                f'save as boundDoc file name {apple_string(str(self._private_path))} file format format document default add to recent files false',
                f'set boundDoc to document {apple_string(self._private_path.name)}',
                f'if posix full name of boundDoc is not {apple_string(str(self._private_path))} then error "WPSC_BINDING_FAILED"',
                'set boundWindow to active window of boundDoc',
                *(['set visible of boundWindow to true'] if visible else []),
                'set nativeRows to {{"binding", id of boundWindow, posix full name of boundDoc as text}}',
            ], bind=False)
            self._bind(rows)
            return self
        except BaseException:
            self._retain('New document creation or binding unverified')
            self.lock.close()
            raise

    @classmethod
    def attach_active(cls):
        self = cls()
        self._prepare()
        try:
            rows = self._execute([
                '-- WPSC_BIND_ATTACH',
                'if not running then error "WPSC_NO_ACTIVE_DOCUMENT"',
                'if (count documents) is 0 then error "WPSC_NO_ACTIVE_DOCUMENT"',
                'set boundWindow to active window',
                'set boundDoc to document of boundWindow',
                'set wid to id of boundWindow',
                'set nativeRows to {{"binding", wid, posix full name of boundDoc as text, name of boundDoc as text, read only of boundDoc}}',
            ], bind=False, launch=False)
            self._bind(rows)
            return self
        except BaseException:
            self.lock.close()
            if not self._quarantined:
                shutil.rmtree(self.staging_root)
            raise

    def _bind(self, rows):
        if len(rows) != 1 or len(rows[0]) < 3 or rows[0][0] != 'binding':
            raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED')
        if not self._owns_doc and (type(rows[0][1]) is not int or rows[0][1] < 1):
            raise NativeWordCapabilityError('Word does not expose a trustworthy active-document window identity')
        if self._owns_doc and Path(rows[0][2]).resolve() != self._private_path.resolve():
            raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED')
        self._window_id, self._bound_path = rows[0][1:3]
        self._bound_name = rows[0][3] if len(rows[0]) > 3 else Path(self._bound_path).name
        if not self._owns_doc:
            if len(rows[0]) < 5 or type(rows[0][4]) is not bool:
                raise NativeWordCapabilityError('Attached Word native read-only state is unavailable')
            self._read_only = rows[0][4]
            if self._bound_path and Path(self._bound_path).is_absolute() and Path(self._bound_path).is_file():
                self._source_path = Path(self._bound_path)
                self._source_digest = self._digest(self._source_path)

    def _binding(self):
        if self._owns_doc and self._bound_path:
            lines = [f'set boundDoc to document {apple_string(Path(self._bound_path).name)}', 'set boundWindow to active window of boundDoc']
        elif self._window_id is None:
            raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED')
        else:
            lines = [f'set boundWindow to window id {self._window_id}', 'set boundDoc to document of boundWindow']
        if self._bound_path:
            lines += [f'if (posix full name of boundDoc as text) is not {apple_string(self._bound_path)} then error "WPSC_STALE_DOCUMENT"']
        else:
            lines += [f'if (name of boundDoc as text) is not {apple_string(self._bound_name)} then error "WPSC_STALE_DOCUMENT"', 'if (path of boundDoc as text) is not "" then error "WPSC_STALE_DOCUMENT"']
        return lines

    def _retain(self, reason):
        self._quarantined = True
        if self.lock and not self.lock.quarantine_path.exists():
            self.lock.quarantine({'schema': 1, 'stagingRoot': str(self.staging_root), 'reason': reason, 'windowId': self._window_id, 'documentPath': self._bound_path})

    def _execute(self, lines, *, bind=True, launch=True):
        if self._closed or self._quarantined:
            raise NativeWordError('NATIVE_WORD_QUARANTINED', staging_path=self.staging_root)
        budget = self._remaining()
        source = _JSON
        if not launch:
            source += 'if not application "Microsoft Word" is running then error "WPSC_NO_ACTIVE_DOCUMENT"\n'
        source += f'with timeout of {max(1, math.ceil(budget))} seconds\ntell application {apple_string(str(WORD_APP))}\nset nativeRows to {{}}\n'
        source += '\n'.join((self._binding() if bind else []) + list(lines))
        source += '\nend tell\nend timeout\nreturn my jsonRows({' + apple_string(_COMPLETION_MARKER) + ', nativeRows})\n'
        script = self.staging_root / (uuid4().hex + '.applescript')
        script.write_text(source, encoding='utf-8')
        os.chmod(script, 0o600)
        log = script.with_suffix('.log')
        try:
            result = subprocess.run(['/usr/bin/osascript', str(script)], capture_output=True, text=True, timeout=self._remaining())
        except (subprocess.TimeoutExpired, OSError) as exc:
            self._retain('Native AppleEvent completion uncertain')
            def partial(value):
                return value.decode('utf-8', errors='replace') if isinstance(value, bytes) else str(value or '')
            log.write_text(type(exc).__name__ + '\n' + partial(getattr(exc, 'stdout', None)) + '\n' + partial(getattr(exc, 'stderr', None)), encoding='utf-8')
            cls = NativeWordTimeoutError if isinstance(exc, subprocess.TimeoutExpired) else NativeWordError
            kw = dict(staging_path=self.staging_root, diagnostic_path=log, quarantine_path=self.lock.quarantine_path)
            raise (cls(**kw) if cls is NativeWordTimeoutError else cls('NATIVE_WORD_QUARANTINED', **kw)) from None
        log.write_text(result.stderr + '\n' + result.stdout, encoding='utf-8')
        if result.returncode:
            self._retain_evidence = True
            if '-1712' in result.stderr:
                self._retain('Native AppleEvent completion uncertain')
                raise NativeWordTimeoutError(staging_path=self.staging_root, diagnostic_path=log, quarantine_path=self.lock.quarantine_path)
            raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED', staging_path=self.staging_root, diagnostic_path=log)
        try:
            envelope = json.loads(result.stdout)
            if not isinstance(envelope, list) or len(envelope) != 2 or envelope[0] != _COMPLETION_MARKER or not isinstance(envelope[1], list):
                raise ValueError()
            self._remaining()
            return envelope[1]
        except (TypeError, ValueError):
            self._retain('Native completion acknowledgement invalid')
            raise NativeWordError('NATIVE_WORD_QUARANTINED', staging_path=self.staging_root, diagnostic_path=log, quarantine_path=self.lock.quarantine_path) from None

    def __enter__(self):
        if self._closed:
            raise ValueError('Session is closed')
        return self

    @property
    def publication_deadline(self):
        return self._deadline

    def _remaining(self):
        try:
            return min(OPERATION_TIMEOUT, remaining(self._deadline))
        except NativeWordTimeoutError:
            if self._bound_path is not None:
                self._retain('Session deadline expired before verified document close')
            raise NativeWordTimeoutError(staging_path=self.staging_root, quarantine_path=self.lock.quarantine_path if self._quarantined else None) from None

    def _digest(self, path):
        digest = hashlib.sha256()
        self._remaining()
        with Path(path).open('rb') as stream:
            while True:
                self._remaining()
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
        self._remaining()
        return digest.hexdigest()

    def _check_source_unchanged(self):
        if self._source_path is None:
            raise ValueError('New Word document requires an explicit save destination')
        try:
            unchanged = self._digest(self._source_path) == self._source_digest
        except OSError:
            unchanged = False
        if not unchanged:
            self._retain_evidence = True
            raise ValueError('Source changed since this Word session opened')

    def __exit__(self, exc_type, exc, tb):
        try:
            self.close()
        except BaseException:
            if exc is None:
                raise

    def _writable(self):
        if self._read_only:
            raise ValueError('Word session is read-only')
        if self._closed:
            raise ValueError('Session is closed')

    def _range(self, target):
        if target == 'selection':
            return ['set targetRange to text object of selection of boundWindow']
        match = re.fullmatch(r'paragraph:([1-9]\d*)', target or '')
        if match:
            return [f'set targetRange to text object of paragraph {int(match[1])} of boundDoc']
        match = re.fullmatch(r'paragraph:@paraId=([0-9a-fA-F]+)', target or '')
        if match:
            if self._structural_changed or not self._owns_doc:
                raise ValueError('Stable paragraph target is stale; re-inspect the document')
            from ..writer import read_paraids_from_docx
            ids = read_paraids_from_docx(self._bound_path)
            indices = [i + 1 for i, value in enumerate(ids) if value and value.lower() == match[1].lower()]
            if len(indices) != 1:
                raise ValueError('Unknown stable paragraph target')
            return [f'if (count paragraphs of boundDoc) is not {len(ids)} then error "WPSC_STALE_PARAGRAPH"', f'set targetRange to text object of paragraph {indices[0]} of boundDoc']
        match = re.fullmatch(r'range:(\d+)-(\d+)', target or '')
        if match and int(match[1]) <= int(match[2]):
            return [f'if {int(match[2])} > (end of content of text object of boundDoc) then error "WPSC_INVALID_RANGE"', f'set targetRange to create range boundDoc start {int(match[1])} end {int(match[2])}']
        match = re.fullmatch(r'table:([1-9]\d*)/cell:([1-9]\d*),([1-9]\d*)', target or '')
        if match:
            return [f'set targetCell to get cell from table (table {int(match[1])} of boundDoc) row {int(match[2])} column {int(match[3])}', 'set targetRange to text object of targetCell']
        raise ValueError('Unsupported Writer target')

    def apply_format_patch(self, target, *, text=None, font=None, paragraph=None, geometry=None, fill=None, line=None, style=None, wrap=None, vertical_alignment=None, page_setup=None, columns=None):
        self._writable()
        patch = {k: v for k, v in locals().copy().items() if k not in ('self', 'target') and v is not None}
        accepted, rejected, mutations = [], [], []
        shape = re.fullmatch(r'shape:([1-9]\d*)', target or '')
        section = re.fullmatch(r'section:([1-9]\d*)', target or '')
        cell = target.startswith('table:') if isinstance(target, str) else False
        if shape:
            resolve = [f'set targetShape to shape {int(shape[1])} of boundDoc', 'set targetRange to text range of text frame of targetShape']
        elif section:
            resolve = [f'set targetSetup to page setup of section {int(section[1])} of boundDoc']
        else:
            resolve = self._range(target)
        mappings = {'font': (_FONT, 'font object of targetRange'), 'paragraph': (_PARA, 'paragraph format of targetRange')}
        if shape:
            mappings.update({'geometry': (_GEOMETRY, 'targetShape'), 'fill': ({'color': 'fore color', 'back_color': 'back color', 'visible': 'visible', 'transparency': 'transparency'}, 'fill format of targetShape'), 'line': ({'color': 'fore color', 'weight': 'weight', 'visible': 'visible', 'transparency': 'transparency'}, 'line format of targetShape')})
        if cell:
            mappings['fill'] = ({'color': 'background pattern color'}, 'shading of targetCell')
        if section:
            mappings = {'page_setup': (_PAGE, 'targetSetup')}
        for group, values in patch.items():
            if group in mappings and isinstance(values, dict):
                names, obj = mappings[group]
                for key, value in values.items():
                    label = group + '.' + key
                    try:
                        if key not in names:
                            raise ValueError()
                        rendered = _value(key, value)
                        mutations.append(f'set {names[key]} of {obj} to {rendered}')
                        accepted.append(label)
                    except (ValueError, TypeError):
                        rejected.append(label)
            elif group in ('text', 'style') and not section:
                try:
                    rendered = _value(group, values)
                    if group == 'text' and (target.startswith('paragraph:') or cell):
                        mutations += ['set replacementStart to start of content of targetRange', 'set replacementEnd to end of content of targetRange', 'set replacementRange to create range boundDoc start replacementStart end (replacementEnd - 1)', f'set content of replacementRange to {rendered}']
                    else:
                        mutations.append(f'set {"content" if group == "text" else "style"} of targetRange to {rendered}')
                    accepted.append(group)
                except (TypeError, ValueError):
                    rejected.append(group)
            elif group == 'vertical_alignment' and cell:
                try:
                    mutations.append('set vertical alignment of targetCell to ' + _value(group, values))
                    accepted.append(group)
                except (TypeError, ValueError):
                    rejected.append(group)
            elif group == 'columns' and section and type(values) is int and 1 <= values <= 45:
                mutations.append(f'set number of text columns targetSetup number of columns {values}')
                accepted.append(group)
            elif group == 'wrap' and shape and type(values) is int and 0 <= values < len(_WRAP):
                mutations.append(f'set wrap type of wrap format of targetShape to {_WRAP[values]}')
                accepted.append(group)
            else:
                rejected.extend(group + '.' + k for k in values) if isinstance(values, dict) else rejected.append(group)
        if rejected:
            return {'accepted': [], 'rejected': rejected}
        if mutations:
            # Range/selection/cell replacement may remove old paragraph marks even
            # when the replacement contains none. Do not reuse on-disk identities.
            if text is not None and (any(mark in text for mark in ('\r', '\n', '\u2029')) or target == 'selection' or target.startswith(('range:', 'table:'))):
                self._structural_changed = True
            self._execute(resolve + mutations + ['set nativeRows to {{"ok"}}'])
        return {'accepted': accepted, 'rejected': []}

    def _position(self, position):
        if position in (None, 'end'):
            return ['set insertionPoint to (end of content of text object of boundDoc) - 1']
        if position == 'start':
            return ['set insertionPoint to 0']
        if isinstance(position, dict) and len(position) == 1:
            key, value = next(iter(position.items()))
            target = f'paragraph:{value}' if key == 'index' else value
            lines = self._range(target)
            if key in ('before', 'index'):
                return lines + ['set insertionPoint to start of content of targetRange']
            if key == 'after':
                return lines + ['set insertionPoint to end of content of targetRange', 'if insertionPoint >= (end of content of text object of boundDoc) then set insertionPoint to (end of content of text object of boundDoc) - 1']
        raise ValueError('Unsupported insertion position')

    @staticmethod
    def _paragraph_boundary():
        # Word's terminal insertion point is before its mandatory final CR.
        # Appending a block there needs a separator if the last paragraph is
        # nonempty, otherwise the new text/style becomes part of that paragraph.
        return ['if insertionPoint > 0 then',
                'set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint',
                'if (content of precedingRange as text) is not return then',
                'set boundaryRange to create range boundDoc start insertionPoint end insertionPoint',
                'set content of boundaryRange to return',
                'set insertionPoint to insertionPoint + 1',
                'end if', 'end if']

    def apply_structural_op(self, op):
        self._writable()
        verb = op.get('op')
        if verb == 'insert':
            if op.get('parent', 'body') != 'body':
                raise ValueError('Unsupported Writer insert parent')
            props = op.get('props') or {}
            kind = op.get('type')
            lines = self._position(op.get('position', 'end')) + ['set insertionRange to create range boundDoc start insertionPoint end insertionPoint']
            if kind in ('paragraph', 'heading'):
                if set(props) - {'text', 'style', 'level'}:
                    raise NativeWordCapabilityError('Mac Word paragraph insertion attributes are unsupported')
                text = props.get('text', '')
                if not isinstance(text, str):
                    raise ValueError('Expected paragraph text')
                lines += self._paragraph_boundary() + ['set insertionRange to create range boundDoc start insertionPoint end insertionPoint', f'set content of insertionRange to {apple_string(text)} & return', f'set targetRange to create range boundDoc start insertionPoint end (insertionPoint + {len(text.encode("utf-16-le")) // 2 + 1})']
                if kind == 'heading':
                    level = props.get('level', 1)
                    if type(level) is not int or not 1 <= level <= 9:
                        raise ValueError('Invalid heading level')
                    lines += [f'set style of targetRange to style heading{level}']
                elif props.get('style'):
                    lines += [f'set style of targetRange to {apple_string(props["style"])}']
            elif kind == 'page_break' and not props:
                lines += ['insert break at insertionRange break type page break']
            elif kind == 'table':
                if set(props) - {'rows', 'cols', 'data'}:
                    raise NativeWordCapabilityError('Mac Word table insertion attributes are unsupported')
                row_count, col_count = props.get('rows', 2), props.get('cols', 2)
                if any(type(n) is not int or not 1 <= n <= 1000 for n in (row_count, col_count)) or row_count * col_count > 10000:
                    raise ValueError('Invalid table dimensions')
                lines += [f'set insertedTable to make new table at boundDoc with properties {{text object:insertionRange, number of rows:{row_count}, number of columns:{col_count}}}']
                for ri, row in enumerate(props.get('data') or []):
                    if ri >= row_count or not isinstance(row, (tuple, list)) or len(row) > col_count:
                        raise ValueError('Table data exceeds dimensions')
                    for ci, value in enumerate(row):
                        lines += [f'set insertedCell to get cell from table insertedTable row {ri + 1} column {ci + 1}', f'set content of text object of insertedCell to {apple_string(str(value))}']
            elif kind == 'textbox':
                if set(props) - {'text', 'left', 'top', 'width', 'height'}:
                    raise NativeWordCapabilityError('Mac Word textbox insertion attributes are unsupported')
                geometry = {k: _number(props.get(k, default)) for k, default in [('left', 100), ('top', 100), ('width', 200), ('height', 50)]}
                lines += [f'set insertedShape to make new text box at boundDoc with properties {{anchor:insertionRange, left position:{geometry["left"]}, top:{geometry["top"]}, width:{geometry["width"]}, height:{geometry["height"]}}}', f'set content of text range of text frame of insertedShape to {apple_string(str(props.get("text", "")))}']
            elif kind == 'image':
                if set(props) != {'path'}:
                    raise NativeWordCapabilityError('Mac Word image insertion attributes are unsupported')
                from PIL import Image
                image_path = Path(props['path']).expanduser().resolve(strict=True)
                with Image.open(image_path) as picture:
                    if picture.format not in ('PNG', 'JPEG'):
                        raise NativeWordCapabilityError('Mac Word requires PNG or JPEG images')
                    suffix = '.png' if picture.format == 'PNG' else '.jpg'
                    picture.verify()
                private_image = self.staging_root / ('image-' + uuid4().hex + suffix)
                copy_file_before_deadline(image_path, private_image, deadline=self._deadline)
                with Image.open(private_image) as picture:
                    picture.verify()
                lines += [f'set insertedImage to make new inline picture at insertionRange with properties {{file name:{apple_string(str(private_image))}, link to file:false, save with document:true}}']
            else:
                raise NativeWordCapabilityError(f'Mac Word {kind} insertion is not implemented')
        elif verb in ('remove', 'move', 'clone'):
            target = op.get('target', '')
            object_match = re.fullmatch(r'(table|shape|inline_shape):([1-9]\d*)', target)
            if object_match:
                native = object_match[1].replace('_', ' ')
                lines = [f'set targetObject to {native} {int(object_match[2])} of boundDoc']
                if verb != 'remove':
                    raise NativeWordCapabilityError('Mac Word object move/clone is not implemented')
                self._execute(lines + ['delete targetObject', 'set nativeRows to {{"ok"}}'])
                self._structural_changed = True
                return {'path': None, 'reinspect_required': True}
            if not (target.startswith('paragraph:') or target.startswith('range:')):
                raise NativeWordCapabilityError('Mac Word structural target is not implemented')
            lines = self._range(target) + ['set sourceRange to targetRange']
            if verb == 'remove':
                lines += ['delete sourceRange']
            else:
                lines += ['set sourceStart to start of content of sourceRange', 'set sourceEnd to end of content of sourceRange'] + self._position(op.get('to', 'end'))
                lines += ['if insertionPoint >= sourceStart and insertionPoint <= sourceEnd then error "WPSC_OVERLAPPING_RANGE"']
                if target.startswith('paragraph:'):
                    lines += self._paragraph_boundary() + ['set sourceStart to start of content of sourceRange', 'set sourceEnd to end of content of sourceRange']
                lines += ['set insertionRange to create range boundDoc start insertionPoint end insertionPoint', 'set formatted text of insertionRange to formatted text of sourceRange']
                if verb == 'move':
                    lines += ['if insertionPoint < sourceStart then', 'set sourceRange to create range boundDoc start (sourceStart + sourceEnd - sourceStart) end (sourceEnd + sourceEnd - sourceStart)', 'end if', 'delete sourceRange']
        else:
            raise ValueError('Unsupported structural operation')
        self._execute(lines + ['set nativeRows to {{"ok"}}'])
        self._structural_changed = True
        return {'path': None, 'reinspect_required': True}

    def inspect_document(self, include_text=True, max_elements=None):
        if max_elements is not None and (type(max_elements) is not int or max_elements < 0):
            raise ValueError('max_elements must be a nonnegative integer')
        lines = []
        for key in ('name', 'saved'):
            lines += _row('doc', 0, key, f'{key} of boundDoc')
        lines += _row('doc', 0, 'path', 'posix full name of boundDoc as text')
        for plural, native in [('paragraphs', 'paragraphs'), ('tables', 'tables'), ('shapes', 'shapes'), ('sections', 'sections'), ('inline_shapes', 'inline shapes'), ('styles', 'Word styles')]:
            lines += _row('doc', 0, 'counts.' + plural, f'count {native} of boundDoc')
        for plural, singular in [('paragraphs', 'paragraph'), ('tables', 'table'), ('shapes', 'shape'), ('sections', 'section')]:
            lines += [f'set objectCount to count {plural} of boundDoc']
            if max_elements is not None:
                lines += [f'if objectCount > {max_elements} then set objectCount to {max_elements}']
            lines += ['repeat with i from 1 to objectCount', f'set targetObject to {singular} i of boundDoc']
            if plural == 'paragraphs':
                lines += ['set targetRange to text object of targetObject'] + self._range_rows(plural, 'i', include_text)
            elif plural == 'tables':
                lines += _row(plural, 'i', 'rows', 'number of rows of targetObject') + _row(plural, 'i', 'columns', 'number of columns of targetObject') + _row(plural, 'i', 'allow_autofit', 'allow auto fit of targetObject')
                lines += ['repeat with cellIndex from 1 to (count cells of text object of targetObject)', 'set targetCell to cell cellIndex of text object of targetObject', 'set cellKey to (i as text) & "," & (row index of targetCell as text) & "," & (column index of targetCell as text)', 'set targetRange to text object of targetCell']
                lines += self._range_rows('cells', 'cellKey', include_text)
                lines += _row('cells', 'cellKey', 'width', 'width of targetCell', optional=True)
                lines += _row('cells', 'cellKey', 'fill.color', 'background pattern color of shading of targetCell', optional=True)
                lines += _row('cells', 'cellKey', 'vertical_alignment', 'my enumIndex(vertical alignment of targetCell, {cell align vertical top, cell align vertical center, cell align vertical bottom})', optional=True)
                lines += ['end repeat']
            elif plural == 'sections':
                for key, prop in _PAGE.items():
                    expression = f'{prop} of page setup of targetObject'
                    if key == 'orientation':
                        expression = 'my enumIndex(' + expression + ', {orient portrait, orient landscape})'
                    lines += _row(plural, 'i', 'page_setup.' + key, expression, optional=True)
                lines += _row(plural, 'i', 'columns', 'count text columns of page setup of targetObject', optional=True)
            else:
                lines += _row(plural, 'i', 'name', 'name of targetObject')
                lines += _row(plural, 'i', 'type', 'my enumIndex(shape type of targetObject, {' + ', '.join('shape type ' + name for name in _SHAPE_TYPES) + '})', optional=True)
                for key, prop in _GEOMETRY.items():
                    lines += _row(plural, 'i', 'geometry.' + key, prop + ' of targetObject', optional=True)
                lines += _row(plural, 'i', 'wrap', 'my enumIndex(wrap type of wrap format of targetObject, {' + ', '.join(_WRAP) + '})', optional=True)
                for group, attributes in [('fill', {'color': 'fore color', 'back_color': 'back color', 'visible': 'visible', 'transparency': 'transparency'}), ('line', {'color': 'fore color', 'weight': 'weight', 'visible': 'visible', 'transparency': 'transparency'})]:
                    for key, prop in attributes.items():
                        lines += _row(plural, 'i', group + '.' + key, prop + ' of ' + group + ' format of targetObject', optional=True)
                if include_text:
                    lines += _row(plural, 'i', 'text', 'content of text range of text frame of targetObject', optional=True)
            lines += ['end repeat']
        result = _decode_snapshot(self._execute(lines))
        if self._owns_doc and not self._structural_changed:
            from ..writer import read_paraids_from_docx
            ids = read_paraids_from_docx(self._bound_path)
            if len(ids) == result['counts'].get('paragraphs'):
                for paragraph in result['paragraphs']:
                    identity = ids[paragraph['index'] - 1]
                    if identity:
                        paragraph['para_id'] = identity
                        paragraph['id'] = 'paragraph:@paraId=' + identity
        return result

    def _range_rows(self, category, index, include_text=True):
        lines = []
        for key, expr in [('start', 'start of content'), ('end', 'end of content')]:
            lines += _row(category, index, key, expr + ' of targetRange')
        if include_text:
            lines += _row(category, index, 'text', 'content of targetRange')
        for group, mapping, obj in [('font', _FONT, 'font object'), ('paragraph', _PARA, 'paragraph format')]:
            for key, prop in mapping.items():
                expression = f'{prop} of {obj} of targetRange'
                if key == 'alignment':
                    expression = f'my enumIndex({expression}, {{{", ".join(_ALIGN)}}})'
                elif key == 'line_spacing_rule':
                    expression = f'my enumIndex({expression}, {{{", ".join(_SPACING)}}})'
                elif key == 'underline':
                    expression = f'my enumIndex({expression}, {{underline none, underline single}})'
                lines += _row(category, index, group + '.' + key, expression, optional=True)
        lines += _row(category, index, 'style', 'name local of style of targetRange', optional=True)
        return lines

    def inspect_selection(self):
        result = _decode_snapshot(self._execute(self._range('selection') + self._range_rows('selection', 1)))
        return result['selection'][0]

    def preflight_save(self, output=None, *, overwrite=False):
        """Reject unsupported attached saves before the orchestration mutates."""
        self._writable()
        if output is not None:
            output = Path(output).expanduser().resolve()
        if output is not None and Path(output).suffix.lower() != '.docx':
            raise NativeWordCapabilityError('Mac Word session output requires DOCX')
        if output is not None and Path(output).exists() and not overwrite:
            raise FileExistsError('Output already exists')
        if self._owns_doc and output is None:
            self._check_source_unchanged()
        if not self._owns_doc:
            if output is not None:
                raise NativeWordCapabilityError('Attached Word document has no verified non-rebinding copy primitive')
            if not self._bound_path or not Path(self._bound_path).is_absolute() or not Path(self._bound_path).is_file():
                raise NativeWordCapabilityError('Unsaved attached document requires an explicit native save')
            if not os.access(self._bound_path, os.W_OK):
                raise NativeWordCapabilityError('Attached Word source is not writable')
            self._check_source_unchanged()
            validate_native_input(Path(self._bound_path), 'writer', deadline=self._deadline)
            self._execute(['if read only of boundDoc then error "WPSC_READ_ONLY"', 'set nativeRows to {{"ok"}}'])

    def save(self, path, fmt=None):
        self._writable()
        self.preflight_save(path)
        destination = Path(path).expanduser().resolve()
        if destination.suffix.lower() != '.docx' or fmt not in (None, 12, 16):
            raise NativeWordCapabilityError('Mac Word session output requires DOCX')
        if not self._owns_doc:
            raise NativeWordCapabilityError('Attached Word document cannot be rebound; use save_current')
        return self._save_to(destination, replace_source=False)

    def _save_to(self, destination, *, replace_source):
        self._execute(['save boundDoc', 'set nativeRows to {{"ok"}}'])
        validator = ValidatorSpec.from_callable(validate_office_package, 'docx')
        def validate(path):
            validate_before_deadline(validator, path, self._deadline)
            # publish_artifact validates the staged copy, then the local temporary
            # file immediately before replacement, and finally the published file.
            if replace_source and Path(path) != destination:
                self._check_source_unchanged()
        try:
            new_digest = self._digest(self._bound_path) if replace_source else None
            publish_artifact(Path(self._bound_path), destination, overwrite=replace_source, validator=validate, deadline=self._deadline)
        except BaseException:
            self._retain_evidence = True
            raise
        if replace_source:
            self._source_digest = new_digest
        return str(destination)

    def save_current(self):
        self._writable()
        if self._owns_doc:
            self._check_source_unchanged()
            return self._save_to(self._source_path, replace_source=True)
        self.preflight_save()
        self._execute(['if read only of boundDoc then error "WPSC_READ_ONLY"', 'save boundDoc', 'set nativeRows to {{"ok"}}'])
        self._source_digest = self._digest(self._source_path)
        return self._bound_path

    def save_copy(self, path, fmt=None):
        if not self._owns_doc:
            raise NativeWordCapabilityError('Attached Word document has no verified non-rebinding copy primitive')
        return self.save(path, fmt)

    def supports_attached_save_copy(self):
        return False

    def is_bound_to(self, path):
        rows = self._execute([_row('doc', 0, 'path', 'posix full name of boundDoc as text')[0]])
        return Path(rows[0][3]).resolve() == Path(path).expanduser().resolve()

    def export_pdf(self, path):
        if not self._owns_doc:
            raise NativeWordCapabilityError('Attached Word PDF export has no verified non-rebinding primitive')
        destination = Path(path).expanduser().resolve()
        if destination.suffix.lower() != '.pdf':
            raise NativeWordCapabilityError('Mac Word PDF output requires a PDF destination')
        if destination.exists():
            raise FileExistsError('Output already exists')
        target = self.staging_root / (uuid4().hex + '.pdf')
        # Word's PDF save-as does not rebind a DOCX document; verify this before publication.
        self._execute([f'save as boundDoc file name {apple_string(str(target))} file format format PDF add to recent files false'] + self._binding() + ['set nativeRows to {{"ok"}}'])
        validator = ValidatorSpec.from_callable(validate_pdf)
        try:
            publish_artifact(target, destination, overwrite=False, validator=lambda p: validate_before_deadline(validator, p, self._deadline), deadline=self._deadline)
        except BaseException:
            self._retain_evidence = True
            raise
        return str(destination)

    def close(self, save_changes=False):
        if self._closed:
            return
        if save_changes and self._owns_doc and self._bound_path is not None and not self._quarantined:
            self.save_current()
        try:
            if self._quarantined:
                raise NativeWordError('NATIVE_WORD_QUARANTINED', staging_path=self.staging_root, quarantine_path=self.lock.quarantine_path)
            if self._owns_doc and self._bound_path is not None:
                try:
                    self._execute(['close boundDoc saving no', 'set nativeRows to {{"ok"}}'])
                except BaseException:
                    self._retain('Bound document close was not verified')
                    raise
            if self.staging_root and not self._retain_evidence:
                shutil.rmtree(self.staging_root)
            self._closed = True
        finally:
            if self.lock:
                self.lock.close()

    @staticmethod
    def _business_format(target, props):
        """Direct Writer units differ from generation-plan line-spacing units."""
        lines = []
        font_names = {'size': 'font size', 'bold': 'bold', 'italic': 'italic',
                      'color': 'color', 'strikethrough': 'strike through', 'underline': 'underline'}
        para_names = {'align': 'alignment', 'indent_first': 'first line indent',
                      'left_indent': 'paragraph format left indent', 'right_indent': 'paragraph format right indent',
                      'space_before': 'space before', 'space_after': 'space after',
                      'keep_together': 'keep together', 'keep_with_next': 'keep with next'}
        for key, value in props.items():
            if value is None:
                continue
            if key in font_names:
                lines.append(f'set {font_names[key]} of font object of {target} to {_value(key, value)}')
            elif key in para_names:
                label = 'alignment' if key == 'align' else key
                lines.append(f'set {para_names[key]} of paragraph format of {target} to {_value(label, value)}')
        family = props.get('font_name')
        latin = props.get('font_name_ascii')
        if family is not None:
            _value('name', family)
            if not family:
                raise ValueError('Font name must not be empty')
            latin = latin or family
            _value('name', latin)
            for prop, value in [('name', family), ('east asian name', family), ('ascii name', latin), ('other name', latin), ('complex script name', latin)]:
                lines.append(f'set {prop} of font object of {target} to {apple_string(value)}')
        elif latin is not None:
            # The baseline ignores ascii_name without font_name; still validate it.
            _value('name', latin)
        value, rule = props.get('line_spacing'), props.get('line_spacing_rule')
        if value is not None:
            _value('line_spacing', value)
        if rule is None:
            if value is not None:
                lines.append(f'set line spacing of paragraph format of {target} to {_number(value)}')
        else:
            aliases = {'single': 0, 'one_and_half': 1, 'double': 2, 'at_least': 3, 'exact': 4, 'multiple': 5}
            code = aliases.get(rule, rule)
            rendered = _value('line_spacing_rule', code)
            lines.append(f'set line spacing rule of paragraph format of {target} to {rendered}')
            if value is not None and code in (3, 4, 5):
                points = float(value) * 12 if code == 5 else value
                lines.append(f'set line spacing of paragraph format of {target} to {_number(points)}')
        return lines

    @staticmethod
    def _business_style(style_name):
        _value('style', style_name)
        if not style_name:
            raise ValueError('Style name must not be empty')
        names = {'Normal': 'style normal', 'Body Text': 'style body text',
                 'Title': 'style title', 'List Paragraph': 'style list paragraph'}
        if re.fullmatch(r'Heading [1-6]', style_name):
            expression = 'style heading' + style_name[-1]
        else:
            expression = names.get(style_name, apple_string(style_name))
        if expression.startswith('style '):
            expression = '(' + expression + ')'
        return f'set semanticStyle to Word style {expression} of boundDoc'

    def _business_paragraph(self, text, style_name='Normal', props=None, *, style_props=None, suffix=(), create_reference_style=False):
        _value('text', text)
        formatting = self._business_format('semanticRange', props or {})
        lines = [self._business_style(style_name)]
        if create_reference_style:
            if style_name != 'First Paragraph':
                raise NativeWordCapabilityError('Only the First Paragraph reference style can be created here')
            lines = ['try'] + lines + ['on error',
                     'set semanticStyle to make new Word style at boundDoc with properties {name local:"First Paragraph"}',
                     'set base style of semanticStyle to style body text', 'end try']
        if style_props:
            lines += self._business_format('semanticStyle', style_props)
        lines += self._position('end') + self._paragraph_boundary()
        lines += ['set semanticRange to create range boundDoc start insertionPoint end insertionPoint',
                  f'set content of semanticRange to {apple_string(text)} & return',
                  f'set semanticRange to create range boundDoc start insertionPoint end (insertionPoint + {len(text.encode("utf-16-le")) // 2 + 1})',
                  'set style of semanticRange to semanticStyle',
                  'reset font object of semanticRange',
                  'reset paragraph format of semanticRange']
        lines += formatting + list(suffix)
        lines += ['set trailingPoint to (end of content of text object of boundDoc) - 1',
                  'set trailingRange to create range boundDoc start trailingPoint end trailingPoint',
                  'set style of trailingRange to style normal',
                  'reset font object of trailingRange', 'reset paragraph format of trailingRange']
        return lines

    def _business_commit(self, lines, *, structural=False):
        self._writable()
        if not lines:
            return
        if structural:
            self._structural_changed = True
        self._execute(list(lines) + ['set nativeRows to {{"ok"}}'])

    def add_paragraph(self, text, size=None, bold=None, italic=None,
                      color=None, align=None, indent_first=None,
                      line_spacing=None, space_after=None, font_name=None,
                      line_spacing_rule=None, font_name_ascii=None,
                      space_before=None):
        props = {key: value for key, value in locals().copy().items() if key not in ('self', 'text')}
        self._business_commit(self._business_paragraph(text, props=props), structural=True)

    def add_heading_level(self, text, level=1, size=None, color=None,
                          line_spacing=None, space_after=None,
                          line_spacing_rule=None, bold=None):
        from ..reference_styles import get_heading_style
        level = min(max(int(level), 1), 6)
        defaults = dict(get_heading_style(level))
        defaults['size'] = defaults.pop('font_size')
        overrides = dict(size=size, color=color, line_spacing=line_spacing,
                         space_after=space_after, line_spacing_rule=line_spacing_rule, bold=bold)
        self._business_commit(self._business_paragraph(text, f'Heading {level}', overrides,
                              style_props=defaults,
                              suffix=[f'set outline level of paragraph format of semanticRange to outline level{level}']), structural=True)

    def add_heading(self, text, size=None, bold=True, color=None):
        return self.add_heading_level(text, level=1, size=size, bold=bold, color=color)

    def add_heading2(self, text, size=None, color=None):
        return self.add_heading_level(text, level=2, size=size, color=color)

    def add_centered(self, text, size=24, bold=True, color=None):
        self.add_paragraph(text, size=size, bold=bold, color=color, align=1)

    def add_styled_paragraph(self, text, style_name):
        self._business_commit(self._business_paragraph(text, style_name), structural=True)

    def add_rich_paragraph(self, spans, style_name):
        from ..document_model import Span
        from .. import reference_styles as RS
        spans = list(spans)
        if any(not isinstance(span, Span) for span in spans):
            raise ValueError('Rich paragraphs require Span objects')
        for span in spans:
            _value('text', span.text)
            if span.link:
                raise NativeWordCapabilityError('Native Word direct rich hyperlink spans have not been verified')
            for key in ('bold', 'italic', 'code', 'strikethrough'):
                if type(getattr(span, key)) is not bool:
                    raise ValueError('Span format flags must be boolean')
        props = dict(left_indent=0, right_indent=0)
        props.update(RS.STYLES['BodyText'])
        if style_name == 'First Paragraph':
            props.update(RS.STYLES['FirstParagraph'])
        props['size'] = 12
        props.update(bold=False, italic=False, color='#000000')
        suffix = []; offset = 0
        for span in spans:
            count = len(span.text.encode('utf-16-le')) // 2
            suffix.append(f'set spanRange to create range boundDoc start (insertionPoint + {offset}) end (insertionPoint + {offset + count})')
            family_lines = self._business_format('spanRange', dict(
                font_name=RS.MONO_FONT if span.code else RS.BODY_FONT,
                font_name_ascii=RS.MONO_FONT if span.code else RS.LATIN_FONT, size=9 if span.code else 12))
            if span.code:
                suffix += ['if semanticCodeStyle is missing value then'] + family_lines + [
                    'else', 'set style of spanRange to semanticCodeStyle', 'end if']
            else:
                suffix += family_lines
            suffix += self._business_format('spanRange', dict(bold=span.bold, italic=span.italic,
                strikethrough=span.strikethrough, underline=0, color='#000000'))
            offset += count
        prefix = []
        if any(span.code for span in spans):
            prefix = ['set semanticCodeStyle to missing value', 'try',
                      'set semanticCodeStyle to Word style "Verbatim Char" of boundDoc', 'end try']
        self._business_commit(prefix + self._business_paragraph(''.join(span.text for span in spans), style_name,
                              props, suffix=suffix, create_reference_style=style_name == 'First Paragraph'), structural=True)

    def _business_list(self, items, glyph, indent, *, ordered):
        from ..reference_styles import BODY_FONT, LATIN_FONT
        items = list(items)
        if any(not isinstance(item, str) for item in items):
            raise ValueError('List items must be text')
        _value('text', glyph); _number(indent)
        if indent < 0:
            raise ValueError('List indent must be nonnegative')
        props = dict(left_indent=indent, indent_first=-indent, line_spacing_rule='one_and_half',
                     space_before=0, space_after=3, font_name=BODY_FONT, font_name_ascii=LATIN_FONT,
                     size=12, bold=False, color='#000000')
        lines = []
        for index, item in enumerate(items, 1):
            prefix = f'{index}.' if ordered else glyph
            tab = f'make new tab stop at paragraph 1 of semanticRange with properties {{tab stop position:{indent}}}'
            lines += self._business_paragraph(prefix + '\t' + item, 'List Paragraph', props, suffix=[tab])
        self._business_commit(lines, structural=True)

    def add_bullet_list(self, items, glyph="\u2022", indent=24):
        self._business_list(items, glyph, indent, ordered=False)

    def add_numbered_list(self, items, indent=24):
        self._business_list(items, '', indent, ordered=True)

    def add_page_break(self):
        self._business_commit(self._position('end') + [
            'set insertionRange to create range boundDoc start insertionPoint end insertionPoint',
            'insert break at insertionRange break type page break'], structural=True)

    def set_margins(self, top, bottom, left, right):
        props = dict(top_margin=top, bottom_margin=bottom, left_margin=left, right_margin=right)
        lines = [f'set {_PAGE[key]} of page setup of boundDoc to {_value(key, value)}' for key, value in props.items()]
        self._business_commit(lines)

    def set_page_size(self, width, height):
        lines = [f'set {name} of page setup of boundDoc to {_value(key, value)}'
                 for name, key, value in [('page width', 'page_width', width), ('page height', 'page_height', height)]]
        self._business_commit(lines)

    def set_orientation(self, landscape):
        value = _value('bold', landscape)
        self._business_commit(['set orientation of page setup of boundDoc to orient ' + ('landscape' if value == 'true' else 'portrait')])

    def set_header(self, text):
        rendered = _value('text', text)
        self._business_commit(['set pagePart to get header (section 1 of boundDoc) index header footer primary',
                               f'set content of text object of pagePart to {rendered}'])

    def set_footer(self, text):
        rendered = _value('text', text)
        self._business_commit(['set pagePart to get footer (section 1 of boundDoc) index header footer primary',
                               f'set content of text object of pagePart to {rendered}'])

    def set_header_footer(self, header=None, footer=None, link_to_previous_header=None,
                          link_to_previous_footer=None):
        lines = ['set semanticSection to section (count sections of boundDoc) of boundDoc']
        for part, link in [('header', link_to_previous_header), ('footer', link_to_previous_footer)]:
            if link is not None:
                value = _value('bold', link)
                lines += ['if (count sections of boundDoc) > 1 then',
                          f'set pagePart to get {part} semanticSection index header footer primary',
                          f'set link to previous of pagePart to {value}',
                          f'if link to previous of pagePart is not {value} then error "WPSC_HEADER_LINK_FAILED"', 'end if']
        for part, text in [('header', header), ('footer', footer)]:
            if text is not None:
                lines += [f'set pagePart to get {part} semanticSection index header footer primary',
                          f'set content of text object of pagePart to {apple_string(str(text))}']
                if part == 'header':
                    lines += ['set alignment of paragraph format of text object of pagePart to align paragraph center',
                              'set pageBorder to get border (paragraph format of text object of pagePart) which border border bottom',
                              'set line style of pageBorder to line style single',
                              'set line width of pageBorder to line width75 point',
                              'set color of pageBorder to {0, 0, 0}']
        self._business_commit(lines)

    def save_docx(self, path):
        return self.save(path, 12)

    @classmethod
    def _style_definition(cls, key, props):
        from ..writer import WriterComposer
        if not isinstance(key, str) or not isinstance(props, dict):
            raise ValueError('Style definitions require named property mappings')
        normalized = {}
        for name, value in props.items():
            canonical = WriterComposer._STYLE_CAMEL_KEYS.get(name, name)
            if canonical in normalized and normalized[canonical] != value:
                raise ValueError('Conflicting style property aliases')
            normalized[canonical] = value
        allowed = {'name', 'type', 'based_on', 'font_name', 'font_name_ascii', 'font_size',
                   'bold', 'italic', 'underline', 'strikethrough', 'color', 'align',
                   'indent_first', 'left_indent', 'right_indent', 'line_spacing',
                   'line_spacing_rule', 'space_before', 'space_after', 'keep_together',
                   'keep_with_next', 'outline_level', 'shading', 'left_border', 'border_color'}
        if set(normalized) - allowed or normalized.get('type', 'paragraph') != 'paragraph':
            raise NativeWordCapabilityError('Only the supported native paragraph style properties are available')
        name = normalized.get('name', key)
        cls._business_style(name)
        base = normalized.get('based_on')
        if base is not None:
            cls._business_style(base)
        formatting = dict(normalized)
        if 'font_size' in formatting:
            formatting['size'] = formatting.pop('font_size')
        lines = cls._business_format('semanticStyle', formatting)
        if normalized.get('outline_level') is not None:
            level = normalized['outline_level']
            if type(level) is not int or not 1 <= level <= 10:
                raise ValueError('Outline level must be between one and ten')
            enum = 'outline level body text' if level == 10 else f'outline level{level}'
            lines.append(f'set outline level of paragraph format of semanticStyle to {enum}')
        if normalized.get('shading') is not None:
            lines.append('set background pattern color of shading of semanticStyle to ' + _value('color', normalized['shading']))
        if normalized.get('left_border') is not None:
            _value('visible', normalized['left_border'])
        if normalized.get('border_color') is not None:
            _value('color', normalized['border_color'])
        if normalized.get('left_border'):
            lines += ['set semanticBorder to get border (paragraph format of semanticStyle) which border border left',
                      'set line style of semanticBorder to line style single',
                      'set line width of semanticBorder to line width225 point',
                      'set color of semanticBorder to ' + _value('color', normalized.get('border_color', '#CCCCCC'))]
        return name, base, lines

    @staticmethod
    def _style_inheritance_order(definitions):
        canonical_names = {name.casefold(): name for name in definitions}
        if len(canonical_names) != len(definitions):
            raise ValueError('Duplicate style name')
        ordered = []; visiting = set(); visited = set()
        def visit(name):
            if name in visiting:
                raise ValueError('Cyclic style inheritance')
            if name in visited:
                return
            visiting.add(name)
            base = definitions[name][0]
            base_name = canonical_names.get(base.casefold()) if base else None
            if base_name is not None:
                visit(base_name)
            visiting.remove(name); visited.add(name); ordered.append(name)
        for name in definitions:
            visit(name)
        return ordered

    @staticmethod
    def _paragraph_style_guard(ref):
        return (f'if (style type of {ref}) is not in '
                '{style type paragraph, style type paragraph only, style type linked} '
                'then error "WPSC_STYLE_TYPE_MISMATCH"')

    @classmethod
    def _style_batch(cls, styles_dict):
        if not isinstance(styles_dict, dict):
            raise ValueError('Styles must be a mapping')
        definitions = {}
        for key, props in styles_dict.items():
            name, base, lines = cls._style_definition(key, props)
            if name in definitions:
                raise ValueError('Duplicate style name')
            definitions[name] = (base, lines)
        ordered = cls._style_inheritance_order(definitions)
        declared_refs = {name.casefold(): f'requestedStyle{index}' for index, name in enumerate(ordered)}
        preflight = []; mutations = []
        for index, name in enumerate(ordered):
            base, formatting = definitions[name]
            ref = f'requestedStyle{index}'
            preflight += [f'set {ref} to missing value', 'try',
                          cls._business_style(name).replace('set semanticStyle to ', f'set {ref} to ', 1), 'end try',
                          f'if {ref} is not missing value then',
                          cls._paragraph_style_guard(ref), 'end if']
            declared_base = declared_refs.get(base.casefold()) if base else None
            if base and declared_base is None:
                preflight += [cls._business_style(base).replace('set semanticStyle to ', f'set baseStyle{index} to ', 1),
                              cls._paragraph_style_guard(f'baseStyle{index}')]
            mutations += [f'set semanticStyle to {ref}', 'if semanticStyle is missing value then',
                          f'set semanticStyle to make new Word style at boundDoc with properties {{name local:{apple_string(name)}}}', 'end if']
            if base:
                if declared_base is not None:
                    base_ref = declared_base
                else:
                    base_ref = f'baseStyle{index}'
                mutations.append(f'set base style of semanticStyle to {base_ref}')
            mutations += formatting + [f'set {ref} to semanticStyle']
        return preflight + mutations

    def ensure_styles(self, styles_dict):
        self._business_commit(self._style_batch(styles_dict))

    def ensure_heading_styles(self, styles_by_level):
        if not isinstance(styles_by_level, dict):
            raise ValueError('Heading styles must be a mapping')
        definitions = {}
        for level, props in styles_by_level.items():
            if isinstance(level, str) and re.fullmatch(r'[1-6]', level):
                level = int(level)
            if type(level) is not int or not 1 <= level <= 6:
                raise ValueError('Heading level must be between one and six')
            _name, base, formatting = self._style_definition(f'Heading {level}', props)
            name = f'Heading {level}'
            if name in definitions:
                raise ValueError('Duplicate heading level')
            definitions[name] = (base, formatting)
        ordered = self._style_inheritance_order(definitions)
        preflight = []; mutations = []
        for name in ordered:
            level = int(name[-1])
            base, formatting = definitions[name]
            preflight += [f'set requestedHeading{level} to Word style (style heading{level}) of boundDoc',
                          self._paragraph_style_guard(f'requestedHeading{level}')]
            if base:
                preflight += [self._business_style(base).replace('set semanticStyle to ', f'set headingBase{level} to ', 1),
                              self._paragraph_style_guard(f'headingBase{level}')]
            mutations += [f'set semanticStyle to requestedHeading{level}']
            if base:
                mutations.append(f'set base style of semanticStyle to headingBase{level}')
            mutations += formatting
        self._business_commit(preflight + mutations)

    def apply_heading_text_color(self, color):
        rendered = _value('color', color)
        preflight = [f'set requestedHeading{level} to Word style (style heading{level}) of boundDoc' for level in range(1, 7)]
        self._business_commit(preflight + [f'set color of font object of requestedHeading{level} to {rendered}' for level in range(1, 7)])

    def add_code_lines(self, lines):
        values = list(lines)
        if any(not isinstance(value, str) for value in values):
            raise ValueError('Code lines must be text')
        commands = []
        for value in values:
            commands += self._business_paragraph(value if value else ' ', 'Source Code')
        commands += self._business_paragraph('', props={'size': 4})
        self._business_commit(commands, structural=True)

    def set_columns(self, count):
        if type(count) is not int or not 1 <= count <= 45:
            raise ValueError('Column count must be between one and 45')
        self._business_commit([f'set number of text columns (page setup of boundDoc) number of columns {count}'])

    def add_section(self, landscape=None, *, continuous=False):
        if type(continuous) is not bool:
            raise ValueError('continuous must be boolean')
        orientation = None if landscape is None else _value('bold', landscape)
        commands = self._position('end') + [
            'set insertionRange to create range boundDoc start insertionPoint end insertionPoint',
            'insert break at insertionRange break type section break ' + ('continuous' if continuous else 'next page')]
        if orientation is not None:
            commands += ['set semanticSection to section (count sections of boundDoc) of boundDoc',
                         'set orientation of page setup of semanticSection to orient ' + ('landscape' if orientation == 'true' else 'portrait')]
        self._business_commit(commands, structural=True)

    def set_page_number_in_footer(self):
        self._business_commit(['set pagePart to get footer (section 1 of boundDoc) index header footer primary',
            'set content of text object of pagePart to "Page "',
            'set footerRange to collapse range (character 5 of text object of pagePart) direction collapse end',
            'create new field text range footerRange field type field page preserve formatting true'])

    def compact_terminal_paragraph(self):
        # Match Python str.strip used by the frozen implementation, plus the
        # Word cell terminator. The character list is local script data.
        whitespace_codes = [7, *range(9, 14), *range(28, 33), 133, 160, 5760,
                            *range(8192, 8203), 8232, 8233, 8239, 8287, 12288]
        whitespace = ', '.join(f'character id {code}' for code in whitespace_codes)
        self._business_commit(['set terminalRange to text object of last paragraph of boundDoc',
            'set compactable to true', 'repeat with terminalCharacter in characters of (content of terminalRange as text)',
            f'if (terminalCharacter as text) is not in {{{whitespace}}} then set compactable to false',
            'end repeat', 'if compactable then',
            'set font size of font object of terminalRange to 1',
            'set space before of paragraph format of terminalRange to 0',
            'set space after of paragraph format of terminalRange to 0',
            'set line spacing rule of paragraph format of terminalRange to line space exactly',
            'set line spacing of paragraph format of terminalRange to 1',
            'set keep together of paragraph format of terminalRange to false',
            'set keep with next of paragraph format of terminalRange to false', 'end if'])

    def reset(self):
        """Frozen Writer reset is a no-op; retain the bound document."""
        return None

    @staticmethod
    def _refresh_commands(*, figures=False, fields=False):
        lines = ['repeat with tocIndex from 1 to (count of tables of contents of boundDoc)',
                 'update (table of contents tocIndex of boundDoc)', 'end repeat']
        if figures:
            lines += ['repeat with figureIndex from 1 to (count of tables of figures of boundDoc)',
                      'update (table of figures figureIndex of boundDoc)', 'end repeat']
        if fields:
            lines += ['repeat with fieldIndex from 1 to (count fields of boundDoc)',
                      'if (update field (field fieldIndex of boundDoc)) is false then error "WPSC_FIELD_UPDATE_FAILED"', 'end repeat']
        return lines

    def refresh_indexes(self):
        self._business_commit(self._refresh_commands(figures=True))

    def update_fields(self):
        self._business_commit(self._refresh_commands(fields=True))

    def finalize_fields(self, *, max_rounds=3):
        if type(max_rounds) is not int or not 1 <= max_rounds <= 100:
            raise ValueError('max_rounds must be a bounded positive integer')
        # The frozen direct method performs one update; convergence belongs to
        # the generation executor and is not silently added to this alias.
        self.update_fields()

    def refresh_fields(self, round_index):
        from ..longform.executor import FieldSnapshot
        if type(round_index) is not int or round_index < 0:
            raise ValueError('round_index must be nonnegative')
        self._writable()
        rows = self._execute(self._refresh_commands(fields=True) + [
            'set nativeRows to {{"stats", compute statistics boundDoc statistic statistic pages, count tables of contents of boundDoc}}'])
        if len(rows) != 1 or len(rows[0]) != 3 or rows[0][0] != 'stats' or any(type(value) is not int or value < 0 for value in rows[0][1:]):
            raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED', staging_path=self.staging_root)
        pages, toc_count = rows[0][1:]
        return (FieldSnapshot(stable_key=('doc:finalize', 'PAGE', 0), field_category='page',
                result_hash=f'{pages}-{toc_count}', toc_page_count=toc_count,
                figure_index_page_count=0, table_index_page_count=0, total_pages=pages),)
