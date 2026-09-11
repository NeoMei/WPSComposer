"""Bound quality notices with serial Writer semantics and local middle recovery.

Native enablement is staged: insertion accepts a main-body paragraph start,
separated from existing tables, in a bounded document without drawing objects.
Interior/cell/terminal/adjacent-table positions remain explicit native gates.
Snapshots contain hashes of text and formatting, never captured body text.
"""
from __future__ import annotations

from dataclasses import dataclass
import re
from uuid import uuid4

from ..writer import NativeWriterObjectError, WriterComposer
from ..longform.privacy import redact_private_text
from .macos_script import apple_string
from .errors import NativeWordError
from .macos_word_degradation import (
    NativeDegradationBox, _committed_range, _style_commands, _style_ack,
    _style_valid, _units,
)
from .macos_word_recovery import hash_commands, _literal

ANCHOR = 'wpsc_document_quality_anchor'
_CODE = re.compile(r'^[A-Z][A-Z0-9_]{0,63}$')
_HASH = re.compile(r'^[0-9a-f]{64}$')
# Full character-format observations are intentionally bounded until native
# performance/large-document acceptance expands this gate.
_MAX_END = 10000


def _error(message):
    return NativeWriterObjectError('DEGRADATION_INSERT_FAILED', message)


def reserve_document_quality_anchor(session, title="生成质量提示", notices=()):
    if not hasattr(session, '_quality_notice_seen'):
        session._quality_notice_seen = set()
    if not hasattr(session, '_quality_notice_anchor_position'):
        safe_title = redact_private_text(str(title or '生成质量提示'))
        try:
            position = _reserve_native(session)
            session._quality_notice_anchor_position = position
            session._quality_notice_title = safe_title
        except Exception:
            raise _error('quality anchor insertion failed') from None
    for notice in notices or ():
        _upsert_quality_notice_mapping(session, notice)


def _upsert_quality_notice_mapping(session, notice):
    raw = notice.get('code') if isinstance(notice, dict) else getattr(notice, 'code', None)
    code = raw if isinstance(raw, str) and _CODE.fullmatch(raw) else 'QUALITY_NOTICE'
    placement = notice.get('placement', 'document') if isinstance(notice, dict) else getattr(notice, 'placement', 'document')
    node = notice.get('nodeId') if isinstance(notice, dict) else getattr(notice, 'node_id', None)
    identity = (code, placement, redact_private_text(str(node or '')))
    if identity in session._quality_notice_seen:
        return
    fallback_text = (notice.get('fallbackText') or notice.get('message', '')
                     if isinstance(notice, dict) else getattr(notice, 'message', ''))
    try:
        target = _target_at(session, int(session._quality_notice_anchor_position))
        display = WriterComposer._degradation_display(code, fallback_text)
        if not session._quality_notice_seen:
            display = f'{session._quality_notice_title}\r{display}'
        box = _insert_box_native(session, target, display)
        session._quality_notice_anchor_position = int(box.Range.End)
    except Exception:
        raise _error('quality notice upsert failed') from None
    session._quality_notice_seen.add(identity)


def upsert_document_quality_notice(session, issue):
    if not hasattr(session, '_quality_notice_anchor_position'):
        raise _error('quality anchor is unavailable')
    _upsert_quality_notice_mapping(session, issue)


def add_document_quality_notice(session, notices):
    if not hasattr(session, '_quality_notice_anchor_position'):
        reserve_document_quality_anchor(session, notices=notices)
        return
    for notice in notices or ():
        _upsert_quality_notice_mapping(session, notice)


def add_quality_notice_at_bookmark(session, *, code, message, fallback, node_id, page, bookmark_name=None):
    try:
        target = _target_bookmark(session, bookmark_name)
        display = WriterComposer._degradation_display(
            code, f'{redact_private_text(str(message))} '
            f'(page {int(page)}; {redact_private_text(str(fallback))})')
        return _insert_box_native(session, target, display)
    except Exception:
        raise _error('quality notice insertion failed') from None


def _freeze(value):
    return tuple(_freeze(v) for v in value) if isinstance(value, (list, tuple)) else value


def _identity(session):
    if not hasattr(session, '_quality_session_id'):
        session._quality_session_id = uuid4().hex
    return (session._quality_session_id, session._bound_path, session._bound_name, session._window_id)


@dataclass(frozen=True)
class QualityTarget:
    identity: tuple
    position: int
    document_end: int
    state: tuple
    bookmark: object = None


def _exact(expression, value):
    return f"((current application's NSString's stringWithString:({expression}))'s isEqualToString:{_literal(value)})"


def _guard(session):
    lines = ['set qualitySelection to selection of boundWindow']
    for obj in ('boundDoc', 'document of boundWindow', 'document of qualitySelection'):
        if session._bound_path:
            lines += [f'if not {_exact("posix full name of " + obj + " as text", session._bound_path)} then error "WPSC_STALE_DOCUMENT"']
        else:
            lines += [f'if not {_exact("name of " + obj + " as text", session._bound_name or "")} then error "WPSC_STALE_DOCUMENT"',
                      f'if not {_exact("path of " + obj + " as text", "")} then error "WPSC_STALE_DOCUMENT"']
    if session._window_id is not None:
        lines += [f'if id of boundWindow is not {_literal(session._window_id)} then error "WPSC_STALE_DOCUMENT"']
    return lines


def _native_path(session):
    return session._bound_path or ''


def _failed(session):
    session._retain_evidence = True
    session._retain('Native quality acknowledgement or restoration uncertain')
    raise _error('quality native completion unverified') from None


def _mutate(session, lines):
    session._mutation_preflight()
    session._field_topology_mutation_pending = True
    try:
        rows = session._execute(lines)
    except BaseException:
        if not session._field_topology_mutation_pending:
            session._retain_evidence = True
            session._retain('Native quality completion uncertain')
        raise
    else:
        session._invalidate_field_topology()
        return rows
    finally:
        session._field_topology_mutation_pending = False


def _reserve_native(session):
    session._mutation_preflight()
    lines = _guard(session) + [
        'if story type of qualitySelection is not main text story then error "WPSC_QUALITY_WRONG_STORY"',
        'set qualitySelectionRange to text object of qualitySelection',
        'set qualityPoint to end of content of qualitySelectionRange',
        'set qualitySelectionStart to start of content of qualitySelectionRange',
        'set qualityBeforeEnd to end of content of text object of boundDoc',
        'if qualitySelectionStart < 0 or qualityPoint < qualitySelectionStart or qualityPoint > qualityBeforeEnd then error "WPSC_QUALITY_RANGE_INVALID"',
        *hash_commands('content of text object of boundDoc as text', 'qualityBeforeHash'),
        'set qualityTarget to create range boundDoc start qualityPoint end qualityPoint',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(ANCHOR)},text object:qualityTarget}}',
        f'set qualityBookmark to bookmark {apple_string(ANCHOR)} of boundDoc',
        *hash_commands('content of text object of boundDoc as text', 'qualityAfterHash'),
        f'set nativeRows to {{{{"quality-reserved",{apple_string(_native_path(session))},qualitySelectionStart,qualityPoint,qualityBeforeEnd,'
        'start of bookmark of qualityBookmark,end of bookmark of qualityBookmark,'
        'end of content of text object of boundDoc,qualityBeforeHash,qualityAfterHash}}',
    ]
    # Bookmark-only reservation has no text/topology mutation. The transport
    # still owns uncertainty retention; no title/cursor is committed on failure.
    try:
        rows = session._execute(lines)
    except NativeWordError as exc:
        session._retain_evidence = True
        # A private transport diagnostic is created only after launch. Local
        # script/deadline failures cannot masquerade as a submitted bookmark.
        if exc.diagnostic_path is not None:
            session._retain("Native quality reservation completion uncertain")
        raise
    if not (isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list) and len(rows[0]) == 10):
        _failed(session)
    r = rows[0]
    if not (r[:2] == ['quality-reserved', _native_path(session)]
            and all(type(v) is int for v in r[2:8]) and 0 <= r[2] <= r[3] <= r[4]
            and r[5] == r[6] == r[3] and r[7] == r[4]
            and isinstance(r[8], str) and _HASH.fullmatch(r[8]) and r[8] == r[9]):
        _failed(session)
    return r[3]


def _position_gate():
    return [
        'set qualityDocEnd to end of content of text object of boundDoc',
        f'if qualityDocEnd > {_MAX_END} then error "WPSC_QUALITY_SIZE_UNVERIFIED"',
        'if qualityPoint <= 0 or qualityPoint >= qualityDocEnd - 1 then error "WPSC_QUALITY_POSITION_UNVERIFIED"',
        'if (count shapes of boundDoc) is not 0 or (count inline shapes of boundDoc) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"',
        'set qualityTarget to create range boundDoc start qualityPoint end qualityPoint',
        'if start of content of qualityTarget is not qualityPoint or end of content of qualityTarget is not qualityPoint then error "WPSC_QUALITY_RANGE_INVALID"',
        'set qualityParagraph to text object of paragraph 1 of qualityTarget',
        'if start of content of qualityParagraph is not qualityPoint then error "WPSC_QUALITY_POSITION_UNVERIFIED"',
        # Word's collection iterator supplies an item reference here. Resolve
        # by native ordinal before reading table bounds (target-read-probe-01).
        'repeat with qualityTableOrdinal from 1 to count tables of boundDoc',
        'set qt to table qualityTableOrdinal of boundDoc',
        'if (start of content of text object of qt) <= qualityPoint + 1 and (end of content of text object of qt) >= qualityPoint - 1 then error "WPSC_QUALITY_TABLE_POSITION_UNVERIFIED"',
        'end repeat',
        'repeat with qualityFieldOrdinal from 1 to count fields of boundDoc',
        'set qf to field qualityFieldOrdinal of boundDoc',
        'if (start of content of field code of qf) <= qualityPoint and (end of content of result range of qf) >= qualityPoint then error "WPSC_QUALITY_FIELD_POSITION_UNVERIFIED"',
        'end repeat',
        'repeat with qualityBookmarkOrdinal from 1 to count bookmarks of boundDoc',
        'set qb to bookmark qualityBookmarkOrdinal of boundDoc',
        'if start of bookmark of qb < qualityPoint and end of bookmark of qb > qualityPoint then error "WPSC_QUALITY_CROSSING_BOOKMARK_UNVERIFIED"',
        'end repeat',
    ]


def _layout_commands():
    """Hash section settings, all list definitions and non-main field identity."""
    return [
        'set qualityLayout to {}',
        'repeat with qi from 1 to count sections of boundDoc',
        'set qsection to section qi of boundDoc', 'set qsetup to page setup of qsection',
        'set end of qualityLayout to {"section",qi as integer,orientation of qsetup as text,page width of qsetup,page height of qsetup,'
        'top margin of qsetup,bottom margin of qsetup,left margin of qsetup,right margin of qsetup,header distance of qsetup,footer distance of qsetup,gutter of qsetup,count text columns of qsetup}',
        'repeat with qindex in {header footer primary,header footer first page,header footer even pages}',
        'set qheader to get header qsection index qindex',
        'set qfooter to get footer qsection index qindex',
        'repeat with qpart in {qheader,qfooter}',
        'if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"',
        'set qr to text object of qpart',
        'set end of qualityLayout to {"page-part",header footer index of qpart as text,is header of qpart,link to previous of qpart,content of qr as text}',
        'repeat with qfield in fields of qr',
        'set end of qualityLayout to {"page-field",field type of qfield as text,content of field code of qfield as text,content of result range of qfield as text,'
        'start of content of field code of qfield,end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,locked of qfield}',
        'end repeat', 'end repeat', 'end repeat', 'end repeat',
        'repeat with qi from 1 to count list templates of boundDoc',
        'set qtemplate to list template qi of boundDoc',
        'set end of qualityLayout to {"list-template",qi as integer,name of qtemplate as text,outline numbered of qtemplate}',
        'repeat with qlevel in list levels of qtemplate',
        'set end of qualityLayout to {"list-level",entry_index of qlevel,linked style of qlevel as text,number format of qlevel as text,number style of qlevel as text,'
        'start at of qlevel,reset on higher of qlevel,number position of qlevel,text position of qlevel,tab position of qlevel,trailing character of qlevel as text,list level alignment of qlevel as text}',
        'end repeat', 'end repeat',
        'repeat with qstyle in Word styles of boundDoc',
        'set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}',
        'end repeat',
        *hash_commands('my jsonRows(qualityLayout)', 'qualityLayoutHash'),
        'set end of qualityState to {"layout",qualityLayoutHash}',
    ]


def _snapshot_commands():
    """Project unchanged objects around [qualityPoint, qualityPoint+qualityDelta).

    In the zero-delta preimage every paragraph/object is recorded. A nonzero
    projection omits only fully enclosed new paragraphs and the one new table;
    any crossing paragraph, field or original object fails exact equality.
    Character/paragraph style values are hashed as typed JSON in native memory.
    """
    lines = [
        'set qe to end of content of text object of boundDoc',
        'set qp to create range boundDoc start 0 end qualityPoint',
        'set qs to create range boundDoc start (qualityPoint + qualityDelta) end qe',
        *hash_commands('content of qp as text', 'qualityPrefixHash'),
        *hash_commands('content of qs as text', 'qualitySuffixHash'),
        'set qualityState to {{"quality-state",qe - qualityDelta,qualityPrefixHash,qualitySuffixHash}}',
        *_layout_commands(),
        'repeat with qi from 1 to count paragraphs of boundDoc',
        'set qr to text object of paragraph qi of boundDoc',
        'set qa to start of content of qr', 'set qz to end of content of qr',
        'if qualityDelta is 0 or qa < qualityPoint or qz > qualityPoint + qualityDelta then',
        # Range fallback can share the final inserted paragraph with the old
        # suffix. Observe only that original suffix, including its formatting.
        'if qualityDelta > 0 and qa >= qualityPoint and qa < qualityPoint + qualityDelta and qz > qualityPoint + qualityDelta then',
        'set qa to qualityPoint + qualityDelta',
        'set qr to create range boundDoc start qa end qz', 'end if',
        *hash_commands('content of qr as text', 'qualityTextHash'),
        'set qpf to paragraph format of qr',
        'set qualityFormats to {{name local of style of qr as text,first line indent of qpf,paragraph format left indent of qpf,paragraph format right indent of qpf,'
        'space before of qpf,space after of qpf,line spacing of qpf,line spacing rule of qpf as text,alignment of qpf as text,'
        'keep with next of qpf,keep together of qpf,widow control of qpf,outline level of qpf as text,character unit first line indent of qpf,'
        'list type of list format of qr as text,list level number of list format of qr,list value of list format of qr,list string of list format of qr as text}}',
        'repeat with qc from qa to qz - 1',
        'set qcr to create range boundDoc start qc end (qc + 1)',
        # Resolve each property record once; subsequent projections are local.
        # Native evidence compares all seven values at every UTF-16 coordinate.
        'set qcf to (get properties of font object of qcr) as record',
        'set qcs to (get properties of shading of qcr) as record',
        'set end of qualityFormats to {name of qcf as text,font size of qcf,bold of qcf,italic of qcf,underline of qcf as text,'
        'color of qcf,background pattern color of qcs}',
        'end repeat',
        *hash_commands('my jsonRows(qualityFormats)', 'qualityFormatHash'),
        'if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta',
        'if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta',
        'set end of qualityState to {"paragraph",qa,qz,qualityTextHash,qualityFormatHash}',
        'end if', 'end repeat',
        'repeat with qi from 1 to count fields of boundDoc',
        'set qf to field qi of boundDoc',
        'set qa to start of content of field code of qf', 'set qz to end of content of field code of qf',
        'set qra to start of content of result range of qf', 'set qrz to end of content of result range of qf',
        *hash_commands('content of field code of qf as text', 'qualityCodeHash'),
        *hash_commands('content of result range of qf as text', 'qualityResultHash'),
        'if qa >= qualityPoint + qualityDelta then',
        'set qa to qa - qualityDelta', 'set qz to qz - qualityDelta',
        'set qra to qra - qualityDelta', 'set qrz to qrz - qualityDelta', 'end if',
        'set end of qualityState to {"field",qi as integer,field type of qf as text,qa,qz,qra,qrz,qualityCodeHash,qualityResultHash,locked of qf}',
        'end repeat',
        'set qualityOldOrdinal to 0',
        'repeat with qi from 1 to count tables of boundDoc', 'set qt to table qi of boundDoc',
        'set qa to start of content of text object of qt', 'set qz to end of content of text object of qt',
        'if qualityDelta is 0 or qa is not qualityPoint or qz is not qualityPoint + qualityDelta then',
        *hash_commands('content of text object of qt as text', 'qualityTextHash'),
        'if qa >= qualityPoint + qualityDelta then', 'set qa to qa - qualityDelta', 'set qz to qz - qualityDelta', 'end if',
        'set qualityOldOrdinal to qualityOldOrdinal + 1',
        'set qualityRowFlags to {}', 'repeat with qualityRowOrdinal from 1 to count rows of qt',
        'set qrow to row qualityRowOrdinal of qt',
        'set end of qualityRowFlags to allow break across pages of qrow', 'end repeat',
        'set end of qualityState to {"table",qualityOldOrdinal,qa,qz,count rows of qt,count columns of qt,qualityTextHash,qualityRowFlags}',
        'end if', 'end repeat',
        'repeat with qi from 1 to count bookmarks of boundDoc', 'set qb to bookmark qi of boundDoc',
        'set qa to start of bookmark of qb', 'set qz to end of bookmark of qb',
        'set qhashStart to qa',
        'if qualityDelta > 0 and qa is qualityPoint and qz > qualityPoint then set qhashStart to qa + qualityDelta',
        'set qr to create range boundDoc start qhashStart end qz',
        *hash_commands('content of qr as text', 'qualityTextHash'),
        'if qa >= qualityPoint + qualityDelta then set qa to qa - qualityDelta',
        'if qz > qualityPoint + qualityDelta then set qz to qz - qualityDelta',
        'set end of qualityState to {"bookmark",name of qb as text,qa,qz,qualityTextHash}',
        'end repeat', 'set end of qualityState to {"quality-state-end"}',
    ]
    return lines


def _valid_state(rows):
    if not (isinstance(rows, (list, tuple)) and len(rows) >= 2
            and isinstance(rows[-1], (list, tuple)) and tuple(rows[-1]) == ('quality-state-end',)):
        return False
    h = rows[0]
    digest = lambda v: isinstance(v, str) and _HASH.fullmatch(v) is not None
    if not (isinstance(h, (list, tuple)) and len(h) == 4 and h[0] == 'quality-state' and type(h[1]) is int and 1 <= h[1] <= _MAX_END
            and digest(h[2]) and digest(h[3])):
        return False
    names = set(); counts = {'field': 0, 'table': 0}; prior_paragraph_end = 0; layouts = 0
    for r in rows[1:-1]:
        if not isinstance(r, (list, tuple)) or not r:
            return False
        kind = r[0]
        if kind == 'layout':
            if not (len(r) == 2 and digest(r[1])): return False
            layouts += 1
            continue
        if kind == 'paragraph':
            if not (len(r) == 5 and digest(r[3]) and digest(r[4])): return False
            positions = r[1:3]
            if positions[0] != prior_paragraph_end: return False
            prior_paragraph_end = positions[1]
        elif kind == 'field':
            counts[kind] += 1
            if not (len(r) == 10 and type(r[1]) is int and r[1] == counts[kind] and isinstance(r[2], str)
                    and digest(r[7]) and digest(r[8]) and type(r[9]) is bool): return False
            positions = r[3:7]
        elif kind == 'table':
            counts[kind] += 1
            if not (len(r) == 8 and type(r[1]) is int and r[1] == counts[kind]
                    and all(type(v) is int and v > 0 for v in r[4:6]) and digest(r[6])
                    and isinstance(r[7], (list, tuple)) and len(r[7]) == r[4]
                    and all(type(v) is bool for v in r[7])): return False
            positions = r[2:4]
        elif kind == 'bookmark':
            if not (len(r) == 5 and isinstance(r[1], str) and r[1] and r[1] not in names and digest(r[4])): return False
            names.add(r[1]); positions = r[2:4]
        else:
            return False
        if any(type(v) is not int or not 0 <= v <= h[1] for v in positions): return False
        if any(a > b for a, b in zip(positions, positions[1:])): return False
    return prior_paragraph_end == h[1] and layouts == 1


def _target_rows(session, prefix, bookmark=None):
    session._mutation_preflight()
    rows = session._execute(_guard(session) + prefix + _position_gate() + [
        'set qualityDelta to 0', *_snapshot_commands(),
        f'set nativeRows to {{{{"quality-target",{apple_string(_native_path(session))},qualityPoint,qualityDocEnd}},{{"preimage",qualityState}}}}',
    ])
    if not (isinstance(rows, list) and len(rows) == 2 and isinstance(rows[0], list)
            and len(rows[0]) == 4 and rows[0][:2] == ['quality-target', _native_path(session)]
            and all(type(v) is int for v in rows[0][2:]) and 0 < rows[0][2] < rows[0][3]-1
            and isinstance(rows[1], list) and len(rows[1]) == 2 and rows[1][0] == 'preimage'
            and _valid_state(rows[1][1]) and rows[1][1][0][1] == rows[0][3]):
        _failed(session)
    return QualityTarget(_identity(session), rows[0][2], rows[0][3], _freeze(rows[1][1]), bookmark)


def _target_at(session, position):
    if type(position) is not int or position < 0:
        raise ValueError('invalid native quality coordinate')
    target = _target_rows(session, [f'set qualityPoint to {position}'])
    if target.position != position:
        _failed(session)
    return target


def _bookmark_lookup(bookmark_name):
    name = bookmark_name or ANCHOR
    # Word's native collection accepts a name or an integer ordinal. Do not
    # stringify an arbitrary object into a different bookmark identity.
    if not isinstance(name, str) and type(name) is not int:
        raise ValueError('unrepresentable native bookmark lookup')
    literal = _literal(name)
    return [f'if not (exists bookmark {literal} of boundDoc) then error "WPSC_QUALITY_BOOKMARK_MISSING"',
            f'set qualityBookmark to bookmark {literal} of boundDoc',
            ('set qualityPoint to end of content of text object of paragraph 1 of (text object of qualityBookmark)'
             if bookmark_name else 'set qualityPoint to start of bookmark of qualityBookmark')]


def _target_bookmark(session, bookmark_name):
    return _target_rows(session, _bookmark_lookup(bookmark_name), (bookmark_name,))


def _target_guard(session, target):
    lines = _guard(session) + [f'set qualityPoint to {target.position}']
    if target.bookmark is not None:
        lines += _bookmark_lookup(target.bookmark[0]) + [f'if qualityPoint is not {target.position} then error "WPSC_QUALITY_TARGET_STALE"']
    return lines + _position_gate() + [
        'set qualityDelta to 0', *_snapshot_commands(),
        f'set qualityExpectedState to {_literal(target.state)}',
        'if not ((current application\'s NSString\'s stringWithString:(my jsonRows(qualityState)))\'s isEqualToString:(my jsonRows(qualityExpectedState))) then error "WPSC_QUALITY_TARGET_STALE"',
        'set qualityBeforeState to qualityState',
    ]


def _same_native_state(message):
    # Compare exact native serialization rather than AppleScript's case-insensitive
    # list/text equality. Both operands are native values, preserving JSON types.
    return 'if not ((current application\'s NSString\'s stringWithString:(my jsonRows(qualityState)))\'s isEqualToString:(my jsonRows(qualityBeforeState))) then error "' + message + '"'


def _table_commands(session, target, display):
    return _target_guard(session, target) + [
        'activate object boundWindow', *_guard(session),
        'set qualityNewTable to missing value',
        'try',
        'set qualityTarget to create range boundDoc start qualityPoint end qualityPoint',
        'set qualityNewTable to make new table at boundDoc with properties {text object:qualityTarget,number of rows:1,number of columns:1}',
        'set qualityCell to get cell from table qualityNewTable row 1 column 1',
        f'set content of text object of qualityCell to {apple_string(display)}',
        'set noticeRange to text object of qualityCell', *_style_commands(block=True),
        'set allow break across pages of row 1 of qualityNewTable to false',
        'set qualityStart to start of content of text object of qualityNewTable',
        'set qualityEnd to end of content of text object of qualityNewTable',
        'set qualityDelta to qualityEnd - qualityStart',
        'if qualityStart is not qualityPoint then error "WPSC_QUALITY_TABLE_SHIFTED"',
        *_snapshot_commands(), _same_native_state('WPSC_QUALITY_OUTSIDE_CHANGED'),
        'set qualityOrdinal to 0',
        'repeat with qi from 1 to count tables of boundDoc',
        'set qt to table qi of boundDoc',
        'if start of content of text object of qt is qualityStart and end of content of text object of qt is qualityEnd then',
        'if qualityOrdinal is not 0 then error "WPSC_QUALITY_TABLE_AMBIGUOUS"',
        'set qualityOrdinal to qi as integer', 'end if', 'end repeat',
        f'set qualityDisplayRange to create range boundDoc start qualityStart end (qualityStart + {_units(display)})',
        f'set nativeRows to {{{{"quality-insert",{apple_string(_native_path(session))},qualityPoint,qualityStart,qualityEnd,'
        f'content of text object of qualityNewTable as text,{target.document_end},end of content of text object of boundDoc,qualityOrdinal,'
        'count rows of qualityNewTable,count columns of qualityNewTable,start of content of text object of qualityCell,'
        'end of content of qualityDisplayRange,content of qualityDisplayRange as text}}',
        *_style_ack(block=True),
        'set end of nativeRows to {"degradation-row",allow break across pages of row 1 of qualityNewTable}',
        'set end of nativeRows to {"preimage",qualityBeforeState}',
        'set end of nativeRows to {"preserved",qualityState}',
        'on error qualityError number qualityNumber',
        'if qualityNumber is -1712 or qualityNumber is -609 or qualityNumber is -128 then error qualityError number qualityNumber',
        *_guard(session),
        'if qualityNewTable is not missing value then',
        'set qualityStart to start of content of text object of qualityNewTable',
        'set qualityEnd to end of content of text object of qualityNewTable',
        'if qualityStart is not qualityPoint or qualityEnd <= qualityStart then error "WPSC_QUALITY_PARTIAL_UNSAFE"',
        'if count rows of qualityNewTable is not 1 or count columns of qualityNewTable is not 1 then error "WPSC_QUALITY_PARTIAL_UNSAFE"',
        'set qualityDelta to qualityEnd - qualityStart',
        *_snapshot_commands(), _same_native_state('WPSC_QUALITY_PARTIAL_UNSAFE'),
        'delete qualityNewTable', 'end if',
        'set qualityDelta to 0', *_snapshot_commands(), _same_native_state('WPSC_QUALITY_RESTORE_FAILED'),
        f'set nativeRows to {{{{"quality-restored",{apple_string(_native_path(session))},qualityPoint}},{{"preimage",qualityBeforeState}},{{"restored",qualityState}}}}',
        'end try',
    ]


def _range_commands(session, target, display):
    return _target_guard(session, target) + [
        # Without a trailing paragraph mark, Word applies paragraph formatting
        # to the suffix too. Until broader native proof, allow only an already
        # compatible suffix; never silently change its paragraph formatting.
        'set qpf to paragraph format of qualityParagraph',
        'if space before of qpf is not 0 or space after of qpf is not 3 or keep together of qpf is not true or outline level of qpf is not outline level body text then error "WPSC_QUALITY_FALLBACK_FORMAT_UNVERIFIED"',
        'set noticeRange to create range boundDoc start qualityPoint end qualityPoint',
        f'set content of noticeRange to {apple_string(display)}',
        f'set qualityDelta to {_units(display)}',
        'set noticeRange to create range boundDoc start qualityPoint end (qualityPoint + qualityDelta)',
        *_style_commands(block=True), *_snapshot_commands(),
        _same_native_state('WPSC_QUALITY_RANGE_OUTSIDE_CHANGED'),
        f'set nativeRows to {{{{"quality-range",{apple_string(_native_path(session))},start of content of noticeRange,end of content of noticeRange,'
        f'content of noticeRange as text,{target.document_end},end of content of text object of boundDoc}}}}',
        *_style_ack(block=True),
        'set end of nativeRows to {"preimage",qualityBeforeState}',
        'set end of nativeRows to {"preserved",qualityState}',
    ]


def _states_match(rows, target, labels=('preimage', 'preserved')):
    return (isinstance(rows, list) and len(rows) == 2
            and all(isinstance(r, list) and len(r) == 2 and r[0] == label
                    and _valid_state(r[1]) and _freeze(r[1]) == target.state
                    for r, label in zip(rows, labels)))


def _table_ack(rows, target, display):
    if not (isinstance(rows, list) and len(rows) == 6 and isinstance(rows[0], list) and len(rows[0]) == 14): return None
    r = rows[0]; p = target.position; n = _units(display)
    ordinal = 1 + sum(1 for row in target.state if row[0] == 'table' and row[2] < p)
    if not (r[:2] == ['quality-insert', target.identity[1] or '']
            and all(type(r[i]) is int for i in (2,3,4,6,7,8,9,10,11,12))
            and r[2] == r[3] == r[11] == p and r[4] == p+n+2
            and r[5] == display+'\r\x07\r\x07' and r[6] == target.document_end
            and r[7] == target.document_end+n+2 and r[8] == ordinal and r[9:11] == [1,1]
            and r[12] == p+n and r[13] == display and _style_valid(rows[1:3], block=True)
            and rows[3] == ['degradation-row',False] and rows[3][1] is False
            and _states_match(rows[4:],target)): return None
    return p,r[4],r[5],ordinal


def _range_ack(rows, target, display):
    if not (isinstance(rows,list) and len(rows)==5 and isinstance(rows[0],list) and len(rows[0])==7): return None
    r=rows[0];p=target.position;n=_units(display)
    if not (r[:2]==['quality-range',target.identity[1] or ''] and all(type(r[i]) is int for i in (2,3,5,6))
            and r[2]==p and r[3]==p+n and r[4]==display and r[5]==target.document_end and r[6]==target.document_end+n
            and _style_valid(rows[1:3],block=True) and _states_match(rows[3:],target)):return None
    return p,p+n,display


def _insert_box_native(session, target, display):
    if not isinstance(target,QualityTarget) or target.identity != _identity(session) or not _valid_state(target.state):
        raise ValueError('foreign or invalid native quality target')
    if not isinstance(display,str) or any(ord(c)<32 and c not in '\r\n\t' for c in display):
        raise ValueError('unrepresentable native quality display')
    if _units(display)>1000000:
        raise ValueError('native quality display exceeds bound')
    rows=_mutate(session,_table_commands(session,target,display))
    if (isinstance(rows,list) and len(rows)==3 and isinstance(rows[0],list)
            and len(rows[0])==3 and rows[0][:2]==['quality-restored',_native_path(session)]
            and type(rows[0][2]) is int and rows[0][2]==target.position
            and _states_match(rows[1:],target,('preimage','restored'))):
        rows=_mutate(session,_range_commands(session,target,display))
        facts=_range_ack(rows,target,display)
        if facts is None:_failed(session)
        return NativeDegradationBox(_committed_range(session,facts))
    facts=_table_ack(rows,target,display)
    if facts is None:_failed(session)
    return NativeDegradationBox(_committed_range(session,facts[:3]),facts[3])
