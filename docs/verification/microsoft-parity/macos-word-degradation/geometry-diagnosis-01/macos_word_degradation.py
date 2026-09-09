"""Bound, acknowledged inline/block notices; native validation is a separate gate.

Returned handles are immutable native range snapshots, not live COM proxies or
persistent edit targets. A failed table batch may fall back only after the real
checkpoint recovery has acknowledged restoration of the document preimage.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional
from uuid import uuid4

from ..writer import NativeWriterObjectError, WriterComposer
from . import macos_word_recovery
from .macos_script import apple_string


@dataclass(frozen=True)
class NativeDegradationRange:
    """Session-tagged range facts at acknowledgement time, in UTF-16 units."""
    session_id: str
    Start: int
    End: int
    Text: str


@dataclass(frozen=True)
class NativeDegradationBox:
    Range: NativeDegradationRange
    table_index: Optional[int] = None


def _units(text):
    return len(text.encode('utf-16-le')) // 2


def _display(code, fallback_text, *, inline):
    text = WriterComposer._degradation_display(code, fallback_text, inline=inline)
    try:
        valid = (_units(text) <= 1000000
                 and not any(ord(c) < 32 and c not in '\t\r\n' for c in text))
    except UnicodeError:
        valid = False
    if not valid:
        raise ValueError('Degradation display is not safe native text')
    return text


def _style_commands(*, block):
    lines = [
        'set italic of font object of noticeRange to true',
        'set color of font object of noticeRange to {40092, 0, 1542}',
        'set background pattern color of shading of noticeRange to {64764, 59624, 59110}',
    ]
    if block:
        lines += [
            'set space before of paragraph format of noticeRange to 0',
            'set space after of paragraph format of noticeRange to 3',
            'set keep together of paragraph format of noticeRange to true',
            'set outline level of paragraph format of noticeRange to outline level body text',
        ]
    return lines


def _style_ack(*, block):
    lines = [
        'set end of nativeRows to {"degradation-style",italic of font object of noticeRange,'
        '(color of font object of noticeRange is {40092, 0, 1542}),'
        '(background pattern color of shading of noticeRange is {64764, 59624, 59110})}',
    ]
    if block:
        lines += [
            'set end of nativeRows to {"degradation-paragraph",'
            'space before of paragraph format of noticeRange,'
            'space after of paragraph format of noticeRange,'
            'keep together of paragraph format of noticeRange,'
            '(outline level of paragraph format of noticeRange is outline level body text)}',
        ]
    return lines


def range_commands(session, display, *, block=False, position=None):
    """Compile one literal insertion; no implicit paragraph or UI selection."""
    lines = session._position('end') + [
        'set noticeBeforeEnd to end of content of text object of boundDoc',
        'set noticeBeforeTables to count tables of boundDoc',
    ]
    if position is not None:
        lines += [f'if insertionPoint is not {position} then error "WPSC_DEGRADATION_ANCHOR_CHANGED"']
    mode = 'block-range' if block else 'inline'
    lines += [
        'set noticeStart to insertionPoint',
        'set noticeRange to create range boundDoc start noticeStart end noticeStart',
        f'set content of noticeRange to {apple_string(display)}',
        f'set noticeEnd to noticeStart + {_units(display)}',
        'set noticeRange to create range boundDoc start noticeStart end noticeEnd',
    ] + _style_commands(block=block) + [
        f'set nativeRows to {{{{"degradation-range",{apple_string(mode)},'
        'start of content of noticeRange,end of content of noticeRange,'
        'content of noticeRange as text,noticeBeforeEnd,'
        'end of content of text object of boundDoc,noticeBeforeTables,count tables of boundDoc}}',
    ] + _style_ack(block=block)
    return lines


def table_commands(session, display, position):
    """Catch only acknowledged table failures, leaving rollback to Python."""
    return session._position('end') + [
        f'if insertionPoint is not {position} then error "WPSC_DEGRADATION_ANCHOR_CHANGED"',
        'set noticeBeforeEnd to end of content of text object of boundDoc',
        'set noticeBeforeTables to count tables of boundDoc',
        'try',
        'activate object boundWindow',
        'set noticeTarget to create range boundDoc start insertionPoint end insertionPoint',
        'set noticeTable to make new table at boundDoc with properties '
        '{text object:noticeTarget,number of rows:1,number of columns:1}',
        'set noticeCell to get cell from table noticeTable row 1 column 1',
        f'set content of text object of noticeCell to {apple_string(display)}',
        'set noticeRange to text object of noticeCell',
    ] + _style_commands(block=True) + [
        'set allow break across pages of row 1 of noticeTable to false',
        'set noticeStart to start of content of text object of noticeCell',
        f'set noticeEnd to noticeStart + {_units(display)}',
        'set noticeDisplayRange to create range boundDoc start noticeStart end noticeEnd',
        f'set nativeRows to {{{{"degradation-table",{position},'
        'start of content of text object of noticeTable,end of content of text object of noticeTable,'
        'content of text object of noticeTable as text,noticeBeforeEnd,'
        'end of content of text object of boundDoc,noticeBeforeTables,count tables of boundDoc,'
        'count rows of noticeTable,count columns of noticeTable,noticeStart,noticeEnd,'
        'content of noticeDisplayRange as text}}',
    ] + _style_ack(block=True) + [
        'set end of nativeRows to {"degradation-row",allow break across pages of row 1 of noticeTable}',
        'on error noticeError number noticeNumber',
        # Timeout/invalid connection cannot authorize further AppleEvents.
        'if noticeNumber is -1712 or noticeNumber is -609 or noticeNumber is -128 then error noticeError number noticeNumber',
        f'set nativeRows to {{{{"degradation-table-failed",{position}}}}}',
        'end try',
    ]


def _failed(session, *, quarantine=True):
    session._retain_evidence = True
    if quarantine:
        session._retain('Native degradation acknowledgement or completion uncertain')
    raise NativeWriterObjectError(
        'DEGRADATION_INSERT_FAILED', 'degradation insertion failed'
    ) from None


def _execute(session, lines):
    session._mutation_preflight()
    try:
        return session._execute(lines)
    except BaseException as exc:
        session._retain_evidence = True
        session._retain('Native degradation completion uncertain')
        if not isinstance(exc, Exception):
            raise
        _failed(session)


def _integers(row, positions):
    return all(type(row[i]) is int for i in positions)


def _style_valid(rows, *, block):
    if not (isinstance(rows, list) and len(rows) == (2 if block else 1)):
        return False
    row = rows[0]
    if not (isinstance(row, list) and len(row) == 4
            and row[0] == 'degradation-style' and all(v is True for v in row[1:])):
        return False
    if block:
        row = rows[1]
        if not (isinstance(row, list) and len(row) == 5
                and row[0] == 'degradation-paragraph'
                and all(type(v) in (int, float) for v in row[1:3])
                and row[1:3] == [0, 3] and all(v is True for v in row[3:])):
            return False
    return True


def _range_ack(rows, display, *, block, position=None):
    if not isinstance(rows, list) or len(rows) != (3 if block else 2):
        return None
    row = rows[0]
    mode = 'block-range' if block else 'inline'
    if not (isinstance(row, list) and len(row) == 9
            and row[:2] == ['degradation-range', mode]
            and _integers(row, (2, 3, 5, 6, 7, 8))
            and row[2] >= 0 and row[3] == row[2] + _units(display)
            and row[4] == display and row[5] == row[2] + 1
            and row[6] == row[3] + 1 and 0 <= row[7] == row[8]
            and (position is None or row[2] == position)
            and _style_valid(rows[1:], block=block)):
        return None
    return row[2], row[3], row[4]


def _table_ack(rows, display, position):
    if not isinstance(rows, list) or len(rows) != 4:
        return None
    row = rows[0]
    if not (isinstance(row, list) and len(row) == 14
            and row[0] == 'degradation-table'
            and _integers(row, (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12))
            and row[1] == position and position <= row[2] <= position+1
            and isinstance(row[4], str) and row[4] == display+'\r\x07'
            and row[3] == row[2]+_units(row[4])
            and row[5] == position+1 and row[6] == row[3]+1
            and row[7] >= 0 and row[8] == row[7]+1
            and row[9:11] == [1, 1] and row[11] == row[2]
            and row[12] == row[11]+_units(display) and row[13] == display
            and _style_valid(rows[1:3], block=True)
            and isinstance(rows[3], list) and len(rows[3]) == 2
            and rows[3][0] == 'degradation-row' and rows[3][1] is False):
        return None
    return row[2], row[3], row[4], row[8]


def _committed_range(session, facts):
    if not hasattr(session, '_degradation_session_id'):
        session._degradation_session_id = uuid4().hex
    session._pending_heading = None
    session._structural_changed = True
    session._invalidate_field_topology()
    return NativeDegradationRange(session._degradation_session_id, *facts)


def _insert_range(session, display, *, block=False, position=None):
    rows = _execute(session, range_commands(session, display, block=block, position=position))
    facts = _range_ack(rows, display, block=block, position=position)
    if facts is None:
        _failed(session)
    handle = _committed_range(session, facts)
    return NativeDegradationBox(handle) if block else handle


def add_inline_degradation(session, code, message, fallback_text):
    display = _display(code, fallback_text, inline=True)
    return _insert_range(session, display)


def add_degradation_notice(session, code, message, fallback_text, placement='block'):
    if placement == 'inline':
        return add_inline_degradation(session, code, message, fallback_text)
    display = _display(code, fallback_text, inline=False)
    session._mutation_preflight()
    position = macos_word_recovery.checkpoint(session)
    rows = _execute(session, table_commands(session, display, position))
    if (isinstance(rows, list) and len(rows) == 1
            and isinstance(rows[0], list) and len(rows[0]) == 2
            and rows[0][0] == 'degradation-table-failed'
            and type(rows[0][1]) is int and rows[0][1] == position):
        try:
            macos_word_recovery.rollback(session, position)
        except BaseException:
            session._retain_evidence = True
            session._retain('Native degradation rollback unverified')
            raise
        return _insert_range(session, display, block=True, position=position)
    facts = _table_ack(rows, display, position)
    if facts is None:
        _failed(session)
    return NativeDegradationBox(_committed_range(session, facts[:3]), facts[3])
