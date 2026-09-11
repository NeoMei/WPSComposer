"""Bound direct Word references; no generation-schema restrictions or file rewriting.

Builders are pure. Native batches acknowledge the inserted ranges and required
properties; only acknowledged code-bookmark identities survive the batch.
"""
from __future__ import annotations

from collections.abc import Mapping
from itertools import islice
import math
from uuid import uuid4

from ..longform.privacy import redact_private_text
from ..writer import _reference_field_code, NativeWriterObjectError
from .errors import NativeWordError
from .macos_script import apple_string
from .macos_word_fields import NativeIndexHandle

_MAX_ITEMS = 10000
_MAX_TEXT = 1000000
_MAX_POINTS = 1584


def _items(value):
    if isinstance(value, (str, bytes, Mapping)):
        raise ValueError('Expected a bounded sequence')
    try:
        result = list(islice(iter(value), _MAX_ITEMS + 1))
    except Exception:
        raise ValueError('Expected a bounded sequence') from None
    if len(result) > _MAX_ITEMS:
        raise ValueError('Too many entries or runs')
    return result


def _text(value):
    try:
        text = str(value)
        size = _units(text)
    except Exception:
        raise ValueError('Invalid display text') from None
    if size > _MAX_TEXT or any(ord(c) < 32 and c not in '\t\r\n' for c in text):
        raise ValueError('Invalid display text')
    return text


def _units(text):
    return len(text.encode('utf-16-le')) // 2


def _point(value):
    try:
        result = float(value)
    except (ValueError, TypeError, OverflowError):
        raise ValueError('Expected finite paragraph geometry') from None
    if not math.isfinite(result) or abs(result) > _MAX_POINTS:
        raise ValueError('Expected bounded paragraph geometry')
    return result


def normalize_runs(runs, list_formatting):
    normalized = []
    try:
        for run in _items(runs):
            if not isinstance(run, Mapping):
                raise ValueError('Expected a run mapping')
            kind = run['type']
            item = {'type': kind}
            if kind == 'text':
                item['text'] = _text(run['text'])
            elif kind in ('citation', 'degradation', 'reference'):
                item['fallbackText'] = _text(run['fallbackText'])
                item['prefix'] = _text(run.get('prefix', ''))
                item['suffix'] = _text(run.get('suffix', ''))
                if kind == 'reference':
                    item['bookmarkName'] = run['bookmarkName']
                    _reference_field_code(item['bookmarkName'])
                    item['prefix'] = _text(run['prefix'])
                    item['suffix'] = _text(run['suffix'])
                elif kind == 'degradation':
                    item['code'] = redact_private_text(_text(run['code']))
                    item['nodeId'] = redact_private_text(_text(run['nodeId']))
            else:
                raise ValueError('Unsupported reference run type')
            normalized.append(item)
        if list_formatting is not None and not isinstance(list_formatting, Mapping):
            raise ValueError('Expected list formatting mapping')
        indent = None if list_formatting is None else _point(list_formatting['indentPt'])
    except (KeyError, TypeError):
        raise ValueError('Missing consumed reference property') from None
    if sum(_units(value) for run in normalized for value in run.values() if isinstance(value,str)) > _MAX_TEXT:
        raise ValueError('Reference paragraph is too large')
    return normalized, indent


def _start(session):
    return session._position('end') + session._paragraph_boundary() + [
        'set operationStart to insertionPoint',
        'set beforeParagraphs to count paragraphs of boundDoc',
        'set nativeRows to {}',
    ]


def _finish():
    return ['set end of nativeRows to {"complete",beforeParagraphs,count paragraphs of boundDoc,operationStart,insertionPoint}']


def _literal(text, index, *, style=False):
    extent = _units(text)
    lines = ['set literalStart to insertionPoint',
             'set literalRange to create range boundDoc start insertionPoint end insertionPoint',
             f'set content of literalRange to {apple_string(text)}',
             f'set insertionPoint to insertionPoint + {extent}',
             'set literalRange to create range boundDoc start literalStart end insertionPoint',
             f'if (content of literalRange as text) is not {apple_string(text)} then error "WPSC_REFERENCE_TEXT_UNVERIFIED"',
             f'set end of nativeRows to {{"literal",{index},literalStart,insertionPoint,content of literalRange as text}}']
    if not text:
        # Word returns missing value for a collapsed empty range. An empty
        # literal is a no-op: acknowledge its native bounds without writing.
        lines = ['set literalStart to insertionPoint',
                 'set literalRange to create range boundDoc start insertionPoint end insertionPoint',
                 'set emptyValue to content of literalRange',
                 'if emptyValue is not missing value then',
                 'if (emptyValue as text) is not "" then error "WPSC_REFERENCE_TEXT_UNVERIFIED"',
                 'end if',
                 f'set end of nativeRows to {{"literal",{index},start of content of literalRange,end of content of literalRange,""}}']
    if style and text:
        # Apply only after all content (including affixes and CR) exists. Word
        # otherwise extends the styled insertion point into following text.
        lines += [f'set end of styleRanges to {{{index},literalStart,insertionPoint}}']
    return lines


def _apply_styles():
    return ['repeat with styleIndex from 1 to count styleRanges',
            'set styleSpec to item styleIndex of styleRanges',
            'set styledRange to create range boundDoc start (item 2 of styleSpec) end (item 3 of styleSpec)',
            'set italic of font object of styledRange to true',
            'set color of font object of styledRange to {40092, 0, 1542}',
            'set background pattern color of shading of styledRange to {64764, 59624, 59110}',
            'set end of nativeRows to {"styled",item 1 of styleSpec,start of content of styledRange,end of content of styledRange,italic of font object of styledRange,(color of font object of styledRange is {40092, 0, 1542}),(background pattern color of shading of styledRange is {64764, 59624, 59110})}',
            'end repeat']


def _normal_tail():
    return ['set trailingRange to create range boundDoc start insertionPoint end insertionPoint',
            'set style of trailingRange to style normal',
            'reset font object of trailingRange', 'reset paragraph format of trailingRange']


def _list_commands(indent):
    return ['set paragraphRange to create range boundDoc start operationStart end insertionPoint',
            'try', 'set style of paragraphRange to Word style "List Paragraph" of boundDoc', 'end try',
            f'set paragraph format left indent of paragraph format of paragraphRange to {indent}',
            f'set first line indent of paragraph format of paragraphRange to {-indent}',
            f'make new tab stop at paragraph 1 of paragraphRange with properties {{tab stop position:{indent}}}',
            'set line spacing rule of paragraph format of paragraphRange to line space1 pt5',
            'set space before of paragraph format of paragraphRange to 0',
            'set space after of paragraph format of paragraphRange to 3',
            'set tabVerified to false',
            'repeat with tabIndex from 1 to count tab stops of paragraph 1 of paragraphRange',
            'set ownTab to tab stop tabIndex of paragraph 1 of paragraphRange',
            f'if (tab stop position of ownTab) is {indent} then set tabVerified to true', 'end repeat',
            'set end of nativeRows to {"list",paragraph format left indent of paragraph format of paragraphRange,first line indent of paragraph format of paragraphRange,(line spacing rule of paragraph format of paragraphRange is line space1 pt5),space before of paragraph format of paragraphRange,space after of paragraph format of paragraphRange,tabVerified}'] + _normal_tail()


def _reference_commands(run, index, identity, controller_owned):
    field_text = run['bookmarkName'] + ' \\h'
    code = _reference_field_code(run['bookmarkName'])
    merged_code = ' REF ' + field_text + ' \\* MERGEFORMAT '
    lines = ['set referenceStart to insertionPoint',
             'set previousDocumentEnd to end of content of text object of boundDoc',
             'set previousFields to count fields of boundDoc', 'try',
             'set ownRange to create range boundDoc start referenceStart end referenceStart',
             f'create new field text range ownRange field type field ref field text {apple_string(field_text)} preserve formatting true',
             'if (count fields of boundDoc) is not previousFields + 1 then error "WPSC_REFERENCE_COUNT_UNVERIFIED"',
             'set ownField to missing value', 'set matchingFields to 0',
             'repeat with fi from 1 to count fields of boundDoc',
             'set candidateField to field fi of boundDoc',
             'if (start of content of field code of candidateField) is referenceStart + 1 then',
             'set ownField to candidateField', 'set matchingFields to matchingFields + 1', 'end if', 'end repeat',
             'if matchingFields is not 1 then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"',
             'if field type of ownField is not field ref then error "WPSC_REFERENCE_TYPE_UNVERIFIED"',
             'set referenceCode to content of field code of ownField as text',
             f'if referenceCode is not {apple_string(" REF " + field_text + " ")} and referenceCode is not {apple_string(merged_code)} and referenceCode is not {apple_string(code)} then error "WPSC_REFERENCE_CODE_UNVERIFIED"',
             'set referenceResultStart to start of content of result range of ownField',
             'set referenceResultEnd to end of content of result range of ownField',
             'if referenceResultStart < referenceStart + 2 or referenceResultEnd < referenceResultStart then error "WPSC_REFERENCE_RANGE_UNVERIFIED"',
             # A field-end character follows its result; write after it.
             'set insertionPoint to referenceResultEnd + 1',
             'if insertionPoint > (end of content of text object of boundDoc) - 1 then error "WPSC_REFERENCE_RANGE_UNVERIFIED"',
             'on error referenceError number referenceNumber',
             'if referenceError starts with "WPSC_REFERENCE_" then error referenceError number referenceNumber',]
    if controller_owned:
        lines += ['set referenceFailed to true',
                  f'set end of nativeRows to {{"reference_failure",{index},referenceStart,(end of content of text object of boundDoc) - 1}}']
    else:
        lines += ['set actualPartialEnd to (end of content of text object of boundDoc) - 1',
                  'if actualPartialEnd < referenceStart then error "WPSC_REFERENCE_ROLLBACK_UNVERIFIED"',
                  'set rollbackRange to create range boundDoc start referenceStart end actualPartialEnd',
                  'set content of rollbackRange to ""',
                  'if (end of content of text object of boundDoc) is not previousDocumentEnd or (count fields of boundDoc) is not previousFields then error "WPSC_REFERENCE_ROLLBACK_UNVERIFIED"',
                  'set insertionPoint to referenceStart', 'set referenceDegraded to true',
                  f'set end of nativeRows to {{"fallback",{index},referenceStart,actualPartialEnd,previousFields,count fields of boundDoc}}']
        lines += _literal(run['fallbackText'], index)
    lines += ['end try', 'if not referenceFailed and not referenceDegraded then',
              f'make new bookmark at boundDoc with properties {{name:{apple_string(identity)}, text object:field code of ownField}}',
              f'if (content of text object of bookmark {apple_string(identity)} of boundDoc as text) is not referenceCode then error "WPSC_REFERENCE_IDENTITY_UNVERIFIED"',
              f'set end of nativeRows to {{"reference",{index},previousFields,count fields of boundDoc,referenceStart + 1,referenceResultStart,referenceResultEnd,referenceCode,{apple_string(identity)}}}',
              'end if']
    return lines


def paragraph_commands(session, runs, indent, controller_owned, static):
    lines = _start(session) + ['set referenceFailed to false', 'set styleRanges to {}']
    handles = {}
    for index, run in enumerate(runs):
        lines += ['if not referenceFailed then']
        kind = run['type']
        if kind == 'text':
            lines += _literal(run['text'], index)
        elif static:
            lines += _literal(run['prefix'], index) + _literal(run['fallbackText'],index,style=True) + _literal(run['suffix'],index)
        elif kind in ('citation','degradation'):
            lines += _literal(run['fallbackText'],index,style=kind=='degradation')
        else:
            identity = 'WPSC_R_' + uuid4().hex[:30]
            handles[index] = identity
            lines += _literal(run['prefix'],index) + ['set referenceDegraded to false']
            lines += _reference_commands(run,index,identity,controller_owned)
            lines += ['if not referenceFailed then'] + _literal(run['suffix'],index) + ['end if']
        lines += ['end if']
    lines += ['if not referenceFailed then'] + _literal('\r',len(runs))
    lines += _apply_styles()
    if indent is not None:
        lines += _list_commands(indent)
    lines += _finish() + ['end if']
    return lines, handles


def bibliography_commands(session, texts, geometry):
    lines = _start(session)
    for index, text in enumerate(texts):
        text += '\r'
        lines += ['set paragraphStart to insertionPoint',
                  'set paragraphRange to create range boundDoc start insertionPoint end insertionPoint',
                  f'set content of paragraphRange to {apple_string(text)}',
                  f'set insertionPoint to insertionPoint + {_units(text)}',
                  'set paragraphRange to create range boundDoc start paragraphStart end insertionPoint']
        props = ''
        if geometry is not None:
            hanging,left,after = geometry
            lines += ['set alignment of paragraph format of paragraphRange to align paragraph left',
                      f'set paragraph format left indent of paragraph format of paragraphRange to {left}',
                      f'set first line indent of paragraph format of paragraphRange to {-hanging}',
                      'set space before of paragraph format of paragraphRange to 0',
                      f'set space after of paragraph format of paragraphRange to {after}',
                      'set keep together of paragraph format of paragraphRange to true']
            props = ',my enumIndex(alignment of paragraph format of paragraphRange,{align paragraph left}),paragraph format left indent of paragraph format of paragraphRange,first line indent of paragraph format of paragraphRange,space before of paragraph format of paragraphRange,space after of paragraph format of paragraphRange,keep together of paragraph format of paragraphRange'
        lines += [f'set end of nativeRows to {{"paragraph",{index},paragraphStart,insertionPoint,content of paragraphRange as text{props}}}']
    return lines + _finish()


def _invalid(session):
    session._retain('Native reference acknowledgement invalid')
    raise NativeWordError('NATIVE_WORD_EXECUTION_FAILED',staging_path=session.staging_root)


def _execute(session, lines):
    try:
        return session._execute_structural(lines)
    except NativeWordError:
        session._retain('Native reference completion uncertain')
        raise


def _complete(row, start, end, paragraph_delta):
    return (isinstance(row,list) and len(row)==5 and row[0]=='complete'
            and all(type(x) is int for x in row[1:]) and row[1]>=1
            and row[2]-row[1]==paragraph_delta and row[3:]==[start,end])


def bibliography(session, entries, *, geometry=None, schema=1, style='numeric'):
    session._writable()
    if geometry is not None and (schema!=1 or style!='numeric'):
        raise ValueError('Invalid structured bibliography contract')
    geometry = None if geometry is None else tuple(_point(x) for x in geometry)
    texts = []
    for entry in _items(entries):
        try:
            if geometry is not None:
                if not isinstance(entry,Mapping):raise ValueError('Expected bibliography mapping')
                number = int(entry['number'])
                if abs(number)>10**12:raise ValueError('Bibliography number is too large')
                texts.append(_text(f'[{number}] {_text(entry["text"])}'))
            else:
                texts.append(_text(entry))
        except Exception:
            raise ValueError('Invalid bibliography entry') from None
    if sum(map(_units,texts))>_MAX_TEXT:raise ValueError('Bibliography is too large')
    if not texts:return {'issues':[]}
    rows = _execute(session,bibliography_commands(session,texts,geometry))
    if not isinstance(rows,list) or len(rows)!=len(texts)+1:_invalid(session)
    start = rows[0][2] if isinstance(rows[0],list) and len(rows[0])>2 else None
    if type(start) is not int or start<0:_invalid(session)
    position = start
    for index,text in enumerate(texts):
        end=position+_units(text)+1
        expected=['paragraph',index,position,end,text+'\r']
        if geometry is not None:
            hanging,left,after=geometry;expected += [0,left,-hanging,0,after,True]
        if rows[index]!=expected:_invalid(session)
        position=end
    if not _complete(rows[-1],start,position,sum(t.count('\r')+1 for t in texts)):_invalid(session)
    return {'issues':[]}


def paragraph(session, runs, owner, list_formatting, controller_owned=False, static=False):
    session._writable()
    runs,indent=normalize_runs(runs,list_formatting)
    if not static:
        for run in runs:
            if run['type']=='degradation':
                run['fallbackText']=redact_private_text(run['fallbackText'])
    # Only native REF tracking consumes owner; other modes leave it inert.
    owner = _text(owner or 'doc:native') if any(r['type']=='reference' for r in runs) and not static else 'doc:native'
    controller_owned=bool(controller_owned)
    lines,handles=paragraph_commands(session,runs,indent,controller_owned,static)
    rows=_execute(session,lines)
    if not isinstance(rows,list):_invalid(session)
    pending=list(rows)
    position=None;start=None;tracked=[];degraded=False;literal_paragraphs=0;styled_ranges=[]
    def take():
        if not pending:_invalid(session)
        row=pending.pop(0)
        if not isinstance(row,list):_invalid(session)
        return row
    def literal(text,index,styled=False):
        nonlocal position,start,literal_paragraphs
        row=take()
        if position is None:
            if len(row)!=5 or type(row[2]) is not int or row[2]<0:_invalid(session)
            position=start=row[2]
        end=position+_units(text)
        if row!=['literal',index,position,end,text]:_invalid(session)
        if styled and text:styled_ranges.append(['styled',index,position,end,True,True,True])
        position=end
        literal_paragraphs += text.count('\r')
    for index,run in enumerate(runs):
        kind=run['type']
        if kind=='text':literal(run['text'],index)
        elif static:
            literal(run['prefix'],index);literal(run['fallbackText'],index,True);literal(run['suffix'],index)
        elif kind in ('citation','degradation'):literal(run['fallbackText'],index,kind=='degradation')
        else:
            literal(run['prefix'],index)
            row=take()
            if row and row[0]=='reference_failure':
                if (not controller_owned or len(row)!=4 or row[:3]!=['reference_failure',index,position]
                        or type(row[3]) is not int or row[3]<position or pending):_invalid(session)
                session._retain_evidence=True
                raise NativeWriterObjectError('CROSS_REFERENCE_FAILED')
            if row and row[0]=='fallback':
                if controller_owned or len(row)!=6 or row[:3]!=['fallback',index,position] or any(type(x) is not int for x in row[2:]) or row[3]<position or row[4]<0 or row[4]!=row[5]:_invalid(session)
                literal(run['fallbackText'],index);degraded=True
            else:
                if (len(row)!=9 or row[:2]!=['reference',index] or any(type(x) is not int for x in row[2:7])
                        or row[2]<0 or row[3]!=row[2]+1 or row[4]!=position+1 or row[5]<=row[4] or row[6]<row[5]
                        or not isinstance(row[7],str) or row[7].strip().removesuffix(' \\* MERGEFORMAT')!=_reference_field_code(run['bookmarkName']) or row[8]!=handles[index]):_invalid(session)
                tracked.append((row[8],row[7]));position=row[6]+1
            literal(run['suffix'],index)
    literal('\r',len(runs))
    for expected_style in styled_ranges:
        if take()!=expected_style:_invalid(session)
    if indent is not None and take()!=['list',indent,-indent,True,0,3,True]:_invalid(session)
    paragraph_delta=literal_paragraphs
    if not _complete(take(),start,position,paragraph_delta) or pending:_invalid(session)
    if tracked:
        if not hasattr(session,'_field_session_id'):session._field_session_id=uuid4().hex
        if not hasattr(session,'_tracked_references'):session._tracked_references=[]
        for identity,code in tracked:
            session._tracked_references.append((NativeIndexHandle(session._field_session_id,identity,owner,'REF','reference'),code))
    if static:return None
    issues=[{'code':r['code'],'message':'Reference target is unresolved','placement':'inline','fallback':'inline','nodeId':r['nodeId']} for r in runs if r['type']=='degradation']
    if degraded:issues.append({'code':'CROSS_REFERENCE_FAILED','message':'Cross-reference used its inline fallback','placement':'inline'})
    return {'issues':issues}
