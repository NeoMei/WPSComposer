"""Native SEQ/STYLEREF shells at the exact bound Word selection.

The cursor is deliberately range-local. Append checkpoint recovery retains its
existing boundary; this module does not pretend it can undo arbitrary edits.
"""
from __future__ import annotations

from dataclasses import dataclass
from uuid import uuid4

from ..writer import _caption_field_codes, _NATIVE_BOOKMARK_RE
from .errors import NativeWordError
from .macos_script import apple_string
from .macos_word_fields import NativeIndexHandle


def _exact(expression, value):
    return f"((current application's NSString's stringWithString:({expression}))'s isEqualToString:{apple_string(value)})"


def _units(value):
    return len(value.encode('utf-16-le')) // 2


def _single_row(rows, label, size):
    return (isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list)
            and len(rows[0])==size and rows[0][0]==label)


def _exact_ack(actual, expected):
    """Native booleans and integer coordinates are different protocol types."""
    if type(actual) is not type(expected): return False
    if isinstance(expected,list):
        return len(actual)==len(expected) and all(_exact_ack(a,b) for a,b in zip(actual,expected))
    return actual==expected


@dataclass(frozen=True)
class NativeCaptionRange:
    """Acknowledged caption range before its terminating paragraph mark."""

    Start: int
    End: int


class BoundNumberingCursor:
    """Validated native cursor; each mutation acknowledges its resulting bounds."""
    def __init__(self, session):
        self.session = session
        session._mutation_preflight()
        rows = session._execute(self._binding_guard() + [
            'set numberSelection to selection of boundWindow',
            'set numberRange to text object of numberSelection',
            'set nativeRows to {{"selection",start of content of numberRange,end of content of numberRange}}',
        ])
        if (not _single_row(rows,'selection',3) or any(type(v) is not int for v in rows[0][1:])
                or not 0<=rows[0][1]<=rows[0][2]):
            self._invalid()
        self.start, self.position = rows[0][1:]

    def _invalid(self):
        self.session._retain_evidence = True
        self.session._retain('Native numbering acknowledgement invalid')
        raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED', staging_path=self.session.staging_root)

    def _binding_guard(self):
        lines = ['set numberSelection to selection of boundWindow',
                 'if story type of numberSelection is not main text story then error "WPSC_NUMBERING_WRONG_STORY"']
        if self.session._bound_path:
            for obj in ('boundDoc','document of boundWindow','document of numberSelection'):
                lines += [f'if not {_exact("posix full name of " + obj + " as text",self.session._bound_path)} then error "WPSC_STALE_DOCUMENT"']
        else:
            for obj in ('boundDoc','document of numberSelection'):
                lines += [f'if not {_exact("name of " + obj + " as text",self.session._bound_name or "")} then error "WPSC_STALE_DOCUMENT"',
                          f'if (path of {obj} as text) is not "" then error "WPSC_STALE_DOCUMENT"']
        return lines

    def _guard(self):
        return self._binding_guard() + [
            'set numberSelectionRange to text object of numberSelection',
            f'if (start of content of numberSelectionRange) is not {self.start} or (end of content of numberSelectionRange) is not {self.position} then error "WPSC_NUMBERING_SELECTION_CHANGED"',
        ]

    @staticmethod
    def _move(expression):
        return [f'set selection start of numberSelection to {expression}',
                f'set selection end of numberSelection to {expression}']

    def _mutate(self, lines):
        session = self.session
        session._mutation_preflight()
        commands = self._guard()+lines
        # _execute clears this only immediately before native submission, after
        # script I/O and its final deadline gate (same contract as recovery).
        session._field_topology_mutation_pending = True
        try:
            rows = session._execute(commands)
        except BaseException:
            if not session._field_topology_mutation_pending:
                session._structural_changed = True
                session._pending_heading = None
                session._retain_evidence = True
                session._retain('Native numbering completion uncertain')
            raise
        else:
            session._invalidate_field_topology()
            session._structural_changed = True
            session._pending_heading = None
            return rows
        finally:
            session._field_topology_mutation_pending = False

    def text(self, value):
        if not isinstance(value,str):
            raise TypeError('native text must be text')
        start, end = self.start, self.start+_units(value)
        lines = [f'set numberRange to create range boundDoc start {self.start} end {self.position}',
                 f'set content of numberRange to {apple_string(value)}',
                 f'set numberRange to create range boundDoc start {start} end {end}',
                 'set numberText to content of numberRange',
                 'if numberText is missing value then set numberText to ""',
                 *self._move(str(end)),
                 'set nativeRows to {{"text",start of content of numberRange,end of content of numberRange,numberText as text}}']
        rows = self._mutate(lines)
        if not _exact_ack(rows,[['text',start,end,value]]): self._invalid()
        self.start = self.position = end

    def localized_style(self):
        rows = self.session._execute(self._guard()+[
            'set numberStyleName to "Heading 1"', 'try',
            'set numberStyleName to name local of Word style (style heading1) of boundDoc as text',
            'on error numberMessage number numberError',
            'if numberError is -1712 or numberError is -609 or numberError is -128 then error numberMessage number numberError',
            'end try', 'set nativeRows to {{"style",numberStyleName}}',
        ])
        if not _single_row(rows,'style',2) or not isinstance(rows[0][1],str): self._invalid()
        return rows[0][1]

    def field(self, code, owner, kind):
        identity = 'WPSC_N_'+uuid4().hex[:30]
        native_type = 'field style ref' if kind=='STYLEREF' else 'field sequence'
        field_text = code.split(' ',1)[1]
        p = self.position
        rows = self._mutate([
            'set numberBefore to count fields of boundDoc',
            f'set numberRange to create range boundDoc start {p} end {p}',
            f'create new field text range numberRange field type {native_type} field text {apple_string(field_text)} preserve formatting true',
            'set numberAfter to count fields of boundDoc',
            'if numberAfter is not numberBefore + 1 then error "WPSC_NUMBERING_FIELD_DELTA"',
            'set numberMatches to 0', 'set numberField to missing value',
            'repeat with numberIndex from 1 to numberAfter',
            'set numberCandidate to field numberIndex of boundDoc',
            f'if (start of content of field code of numberCandidate) is {p+1} then',
            'set numberMatches to numberMatches + 1', 'set numberField to numberCandidate',
            'end if', 'end repeat',
            'if numberMatches is not 1 then error "WPSC_NUMBERING_FIELD_IDENTITY"',
            f'if field type of numberField is not {native_type} then error "WPSC_NUMBERING_FIELD_TYPE"',
            'set numberCode to content of field code of numberField as text',
            'set numberCodeStart to start of content of field code of numberField',
            'set numberCodeEnd to end of content of field code of numberField',
            'set numberResultStart to start of content of result range of numberField',
            'set numberResultEnd to end of content of result range of numberField',
            'set numberNext to numberResultEnd + 1',
            'if numberNext > (end of content of text object of boundDoc) - 1 then error "WPSC_NUMBERING_FIELD_RANGE"',
            f'make new bookmark at boundDoc with properties {{name:{apple_string(identity)},text object:field code of numberField}}',
            f'set numberIdentity to text object of bookmark {apple_string(identity)} of boundDoc',
            *self._move('numberNext'),
            'set nativeRows to {{"field",numberBefore,numberAfter,numberCodeStart,numberCodeEnd,numberResultStart,numberResultEnd,numberCode,'
            f'{apple_string(identity)},start of content of numberIdentity,end of content of numberIdentity,content of numberIdentity as text}}}}',
        ])
        if not _single_row(rows,'field',12):
            self._invalid()
        row = rows[0]
        if (any(type(row[i]) is not int for i in (1,2,3,4,5,6,9,10))
                or row[1]<0 or row[2]!=row[1]+1 or row[3]!=p+1
                or not row[3]<row[4]<row[5]<=row[6]
                or not isinstance(row[7],str)
                or row[7].strip().removesuffix(' \\* MERGEFORMAT')!=code
                or row[8]!=identity or row[9:12]!=[row[3],row[4],row[7]]):
            self._invalid()
        session = self.session
        if not hasattr(session,'_field_session_id'): session._field_session_id = uuid4().hex
        if not hasattr(session,'_tracked_numbering'): session._tracked_numbering = []
        handle = NativeIndexHandle(session._field_session_id,identity,owner or 'doc:native',kind,'numbering')
        session._tracked_numbering.append((handle,row[7]))
        self.start = self.position = row[6]+1

    def bookmark(self, name, start, end):
        rows = self._mutate([
            f'set numberRange to create range boundDoc start {start} end {end}',
            f'make new bookmark at boundDoc with properties {{name:{apple_string(name)},text object:numberRange}}',
            f'set numberBookmark to bookmark {apple_string(name)} of boundDoc',
            'set nativeRows to {{"bookmark",name of numberBookmark,start of bookmark of numberBookmark,end of bookmark of numberBookmark}}',
        ])
        if not _exact_ack(rows,[['bookmark',name,start,end]]): self._invalid()

    def format_paragraph(self, start):
        rows = self._mutate([
            f'set numberRange to create range boundDoc start {start} end {self.position}',
            'set alignment of paragraph format of numberRange to align paragraph right',
            'set keep together of paragraph format of numberRange to true',
            'set nativeRows to {{"format",(alignment of paragraph format of numberRange is align paragraph right),keep together of paragraph format of numberRange}}',
        ])
        if not _exact_ack(rows,[['format',True,True]]): self._invalid()

    def format_caption(self, start, keep_with_next):
        end = self.position
        if type(start) is not int or not 0 <= start <= end:
            self._invalid()
        rows = self._mutate([
            f'set numberRange to create range boundDoc start {start} end {end}',
            'set alignment of paragraph format of numberRange to align paragraph center',
            'set keep together of paragraph format of numberRange to true',
            'set nativeRows to {{"caption-format-base",start of content of numberRange,end of content of numberRange,'
            '(alignment of paragraph format of numberRange is align paragraph center),'
            'keep together of paragraph format of numberRange}}',
        ])
        if not _exact_ack(rows,[['caption-format-base',start,end,True,True]]):
            self._invalid()
        keep = bool(keep_with_next)
        rows = self._mutate([
            f'set numberRange to create range boundDoc start {start} end {end}',
            f'set keep with next of paragraph format of numberRange to {str(keep).lower()}',
            'set nativeRows to {{"caption-format-keep",start of content of numberRange,end of content of numberRange,'
            'keep with next of paragraph format of numberRange}}',
        ])
        if not _exact_ack(rows,[['caption-format-keep',start,end,keep]]):
            self._invalid()
        return NativeCaptionRange(start, end)

    def paragraph(self):
        self.text('\r')


def add_number_shell(cursor, numbering, bookmark_name, owner_node_id):
    _, sequence_code = _caption_field_codes(numbering)
    prefix, suffix = numbering['prefix'], numbering['suffix']
    cursor.text(prefix)
    number_start = cursor.position
    if numbering.get('mode') == 'chapter':
        name = str(cursor.localized_style())
        if not name or '"' in name: name = 'Heading 1'
        cursor.field(f'STYLEREF "{name}" \\s',owner_node_id,'STYLEREF')
        cursor.text('-')
    kind = {'WPSC_FIG':'SEQ_FIG','WPSC_TAB':'SEQ_TAB','WPSC_EQ':'SEQ_EQ'}[numbering['sequenceId']]
    cursor.field(sequence_code,owner_node_id,kind)
    number_end = cursor.position
    if bookmark_name is not None:
        if not _NATIVE_BOOKMARK_RE.fullmatch(bookmark_name): raise ValueError('invalid native bookmark')
        cursor.bookmark(bookmark_name,number_start,number_end)
    if suffix: cursor.text(suffix)
    return number_start,number_end


def add_native_caption(session, caption, numbering, bookmark_name, owner_node_id,
                       *, keep_with_next=False):
    cursor = BoundNumberingCursor(session)
    start = cursor.start
    add_number_shell(cursor, numbering, bookmark_name, owner_node_id)
    if caption:
        cursor.text(' ' + str(caption))
    paragraph_range = cursor.format_caption(start, keep_with_next)
    cursor.paragraph()
    return paragraph_range


def add_equation_number_native(session, *, source, numbering, bookmarkName,
                               fallbackText, owner_node_id=None):
    session._mutation_preflight()
    cursor = BoundNumberingCursor(session)
    start = cursor.position
    cursor.text(str(source or fallbackText))
    cursor.text('\t')
    add_number_shell(cursor,numbering,bookmarkName,owner_node_id)
    cursor.format_paragraph(start)
    cursor.paragraph()
    return {'issues':[]}
