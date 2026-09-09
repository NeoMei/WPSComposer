"""Local append rollback with immutable preimages and acknowledged native topology.

Coordinates are plain native story offsets, scoped to this session, not origin
credentials. Hashing stays inside the bounded AppleEvent and uses no shell.
"""
from __future__ import annotations

from dataclasses import dataclass
import re
import math

from .macos_script import apple_string


@dataclass(frozen=True)
class CheckpointSnapshot:
    coordinate: int
    state: tuple
    tracked_indexes_prefix: tuple
    tracked_references_prefix: tuple
    observed_field_topology: object


def hash_commands(text_expression, output):
    """Fixed executable, private text through stdin; never shell/argv/disk text."""
    return [
        'set recoveryTask to current application\'s NSTask\'s alloc()\'s init()',
        'recoveryTask\'s setLaunchPath:"/usr/bin/shasum"',
        'recoveryTask\'s setArguments:{"-a", "256"}',
        'set recoveryInput to current application\'s NSPipe\'s pipe()',
        'set recoveryOutput to current application\'s NSPipe\'s pipe()',
        'recoveryTask\'s setStandardInput:recoveryInput',
        'recoveryTask\'s setStandardOutput:recoveryOutput',
        'recoveryTask\'s setStandardError:(current application\'s NSFileHandle\'s fileHandleWithNullDevice())',
        f'set recoveryData to (current application\'s NSString\'s stringWithString:({text_expression}))\'s dataUsingEncoding:4',
        'recoveryTask\'s |launch|()',
        '(recoveryInput\'s fileHandleForWriting())\'s writeData:recoveryData',
        '(recoveryInput\'s fileHandleForWriting())\'s closeFile()',
        'set recoveryDigestData to (recoveryOutput\'s fileHandleForReading())\'s readDataToEndOfFile()',
        'recoveryTask\'s waitUntilExit()',
        'if (recoveryTask\'s terminationStatus() as integer) is not 0 then error "WPSC_CHECKPOINT_HASH_FAILED"',
        'set recoveryDigestText to (current application\'s NSString\'s alloc()\'s initWithData:recoveryDigestData |encoding|:4) as text',
        'if (length of recoveryDigestText) is not 68 then error "WPSC_CHECKPOINT_HASH_FAILED"',
        'if text 65 thru 68 of recoveryDigestText is not "  -" & linefeed then error "WPSC_CHECKPOINT_HASH_FAILED"',
        f'set {output} to text 1 thru 64 of recoveryDigestText',
    ]


def _literal(value):
    if isinstance(value,str):
        # Native table and field text contains object terminators. Keep them
        # literal data via character id rather than escaping executable source.
        parts=re.split(r'([\x00-\x08\x0b\x0c\x0e-\x1f])',value)
        return '('+' & '.join(f'(character id {ord(p)})' if len(p)==1 and ord(p)<32 and p not in '\r\n\t' else apple_string(p) for p in parts)+')'
    if type(value) is bool:return str(value).lower()
    if type(value) is int:return str(value)
    if type(value) is float and math.isfinite(value):return repr(value)
    if isinstance(value,(list,tuple)):return '{'+','.join(_literal(v) for v in value)+'}'
    raise ValueError('invalid native fact')


def state_commands(tracked, coordinate=None):
    lines=['set recoveryEnd to end of content of text object of boundDoc',
           'set recoveryBound to recoveryEnd - 1',
           'set recoveryTarget to recoveryBound' if coordinate is None else f'set recoveryTarget to {coordinate}',
           'if recoveryTarget < 0 or recoveryTarget > recoveryBound then error "WPSC_CHECKPOINT_BOUND_FAILED"',
           'set recoveryPrefix to create range boundDoc start 0 end recoveryTarget',
           *hash_commands('content of recoveryPrefix as text','recoveryPrefixHash'),
           'set recoveryGuardStart to recoveryTarget - 512',
           'if recoveryGuardStart < 0 then set recoveryGuardStart to 0',
           'set recoveryGuard to create range boundDoc start recoveryGuardStart end recoveryTarget',
           'set recoveryParagraph to text object of paragraph (count paragraphs of boundDoc) of boundDoc',
           *hash_commands('content of recoveryParagraph as text','recoveryParagraphHash'),
           'set nativeRows to {{"checkpoint-state",1,recoveryBound,recoveryEnd,count paragraphs of boundDoc,start of content of recoveryParagraph,end of content of recoveryParagraph,recoveryPrefixHash,recoveryGuardStart,recoveryTarget,content of recoveryGuard as text,recoveryParagraphHash}}',
           'repeat with recoveryOrdinal from 1 to count tables of boundDoc',
           'set recoveryObject to table recoveryOrdinal of boundDoc',
           'set end of nativeRows to {"table",recoveryOrdinal as integer,start of content of text object of recoveryObject,end of content of text object of recoveryObject,count rows of recoveryObject,count columns of recoveryObject}',
           'end repeat', *_field_commands()]
    for handle,code in tracked:
        name=apple_string(handle.bookmark)
        lines += [f'if not (exists bookmark {name} of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"',
                  f'set recoveryIdentity to text object of bookmark {name} of boundDoc',
                  f'if not ((current application\'s NSString\'s stringWithString:(content of recoveryIdentity as text))\'s isEqualToString:{_literal(code)}) then error "WPSC_FIELD_IDENTITY_STALE"',
                  f'set end of nativeRows to {{"bookmark",{name},start of content of recoveryIdentity,end of content of recoveryIdentity,content of recoveryIdentity as text}}']
    lines += object_commands()
    names=[handle.bookmark for handle,code in tracked]
    lines += ['repeat with recoveryOrdinal from 1 to count bookmarks of boundDoc',
              'set recoveryBookmark to bookmark recoveryOrdinal of boundDoc',
              'set recoveryBookmarkName to name of recoveryBookmark as text',
              f'if {_literal(names)} does not contain recoveryBookmarkName then',
              'set recoveryIdentity to text object of recoveryBookmark',
              *hash_commands('content of recoveryIdentity as text','recoveryBookmarkHash'),
              'set end of nativeRows to {"bookmark-hash",recoveryBookmarkName,start of content of recoveryIdentity,end of content of recoveryIdentity,recoveryBookmarkHash}',
              'end if','end repeat']
    return lines+['set end of nativeRows to {"objects",count shapes of boundDoc,count inline shapes of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}', 'set end of nativeRows to {"checkpoint-end"}']


def object_commands():
    # These objects are not deletion candidates in the bounded REF/table slice.
    # A new or changed object blocks recovery before any destructive command.
    return [
        'repeat with recoveryOrdinal from 1 to count shapes of boundDoc',
        'set recoveryShape to shape recoveryOrdinal of boundDoc',
        'set recoveryShapeText to ""',
        'if has text of text frame of recoveryShape then set recoveryShapeText to content of text range of text frame of recoveryShape as text',
        *hash_commands('recoveryShapeText','recoveryShapeHash'),
        'set end of nativeRows to {"shape",recoveryOrdinal as integer,name of recoveryShape as text,shape type of recoveryShape as text,anchorID of recoveryShape,editID of recoveryShape,start of content of anchor of recoveryShape,end of content of anchor of recoveryShape,left position of recoveryShape,top of recoveryShape,width of recoveryShape,height of recoveryShape,rotation of recoveryShape,z order position of recoveryShape,visible of recoveryShape,recoveryShapeHash}',
        'end repeat',
        'repeat with recoveryOrdinal from 1 to count inline shapes of boundDoc',
        'set recoveryPicture to inline shape recoveryOrdinal of boundDoc',
        *hash_commands('alternative text of recoveryPicture as text','recoveryPictureHash'),
        'set end of nativeRows to {"inline-shape",recoveryOrdinal as integer,anchorID of recoveryPicture,editID of recoveryPicture,start of content of text object of recoveryPicture,end of content of text object of recoveryPicture,width of recoveryPicture,height of recoveryPicture,inline shape type of recoveryPicture as text,recoveryPictureHash}',
        'end repeat',
    ]


def _field_commands():
    return ['repeat with recoveryOrdinal from 1 to count fields of boundDoc',
            'set recoveryObject to field recoveryOrdinal of boundDoc',
            'set end of nativeRows to {"field","main",recoveryOrdinal as integer,field type of recoveryObject as text,content of field code of recoveryObject as text,start of content of field code of recoveryObject,end of content of field code of recoveryObject,start of content of result range of recoveryObject,end of content of result range of recoveryObject}',
            'end repeat']


def _malformed(session):
    session._retain('Native recovery acknowledgement invalid')
    raise ValueError('invalid recovery acknowledgement')


def _parse(session,rows,target=None):
    if not isinstance(rows,list) or len(rows)<2 or rows[-1]!=['checkpoint-end']:_malformed(session)
    h=rows[0]
    if not isinstance(h,list) or len(h)!=12 or h[:2]!=['checkpoint-state',1]:_malformed(session)
    if any(type(h[i]) is not int for i in (1,2,3,4,5,6,8,9)):_malformed(session)
    if not (0<=h[2]==h[3]-1 and h[4]>=1 and 0<=h[5]<h[6]<=h[3]):_malformed(session)
    if h[9]!=(h[2] if target is None else target) or h[8]!=max(0,h[9]-512):_malformed(session)
    if any(not isinstance(h[i],str) or not re.fullmatch('[0-9a-f]{64}',h[i]) for i in (7,11)):_malformed(session)
    if not isinstance(h[10],str) or len(h[10].encode('utf-16-le',errors='surrogatepass'))>1024:_malformed(session)
    seen=set();counts={'table':0,'field':0,'shape':0,'inline-shape':0}; phase=0;objects=None
    for row in rows[1:-1]:
        if not isinstance(row,list) or not row:_malformed(session)
        kind=row[0]
        if kind=='table':
            if phase>0 or len(row)!=6 or any(type(v) is not int for v in row[1:]):_malformed(session)
            counts[kind]+=1
            if row[1]!=counts[kind] or not 0<=row[2]<row[3]<=h[3] or min(row[4:])<1:_malformed(session)
        elif kind=='field':
            if phase>1 or len(row)!=9 or row[1]!='main' or any(type(row[i]) is not int for i in (2,5,6,7,8)) or any(not isinstance(row[i],str) for i in (3,4)):_malformed(session)
            phase=1;counts[kind]+=1
            if row[2]!=counts[kind] or not 1<=row[5]<=row[6]<=row[7]<=row[8]<h[3]:_malformed(session)
        elif kind=='shape':
            if len(row)!=16 or any(type(row[i]) is not int for i in (1,4,5,6,7,13)) or any(not isinstance(row[i],str) for i in (2,3,15)) or type(row[14]) is not bool:_malformed(session)
            if any(type(row[i]) not in (int,float) or not math.isfinite(row[i]) for i in (8,9,10,11,12)) or not re.fullmatch('[0-9a-f]{64}',row[15]):_malformed(session)
            counts[kind]+=1
            if row[1]!=counts[kind] or not 0<=row[6]<=row[7]<=h[3] or min(row[10:12])<0:_malformed(session)
        elif kind=='inline-shape':
            if len(row)!=10 or any(type(row[i]) is not int for i in (1,2,3,4,5)) or any(not isinstance(row[i],str) for i in (8,9)):_malformed(session)
            if any(type(row[i]) not in (int,float) or not math.isfinite(row[i]) or row[i]<0 for i in (6,7)) or not re.fullmatch('[0-9a-f]{64}',row[9]):_malformed(session)
            counts[kind]+=1
            if row[1]!=counts[kind] or not 0<=row[4]<=row[5]<=h[3]:_malformed(session)
        elif kind=='objects':
            if objects is not None or len(row)!=6 or any(type(v) is not int or v<0 for v in row[1:]):_malformed(session)
            objects=row
        elif kind in ('bookmark','bookmark-hash'):
            if len(row)!=5 or not isinstance(row[1],str) or not isinstance(row[4],str) or any(type(row[i]) is not int for i in (2,3)) or not 0<=row[2]<=row[3]<=h[3] or row[1] in seen:_malformed(session)
            if kind=='bookmark-hash' and not re.fullmatch('[0-9a-f]{64}',row[4]):_malformed(session)
            phase=2;seen.add(row[1])
        else:_malformed(session)
    if objects!=['objects',counts['shape'],counts['inline-shape'],counts['table'],counts['field'],len(seen)]:_malformed(session)
    return tuple(tuple(row) for row in rows)


def _tracking(session):
    return (tuple(tuple(item) for item in getattr(session,'_tracked_indexes',())),
            tuple(tuple(item) for item in getattr(session,'_tracked_references',())))


def _fail(session,code,message):
    from ..writer import NativeWriterObjectError
    session._retain_evidence=True
    raise NativeWriterObjectError(code,message) from None


def checkpoint(session):
    try:
        if session._closed or session._quarantined:raise ValueError('session unavailable')
        session._remaining()
        indexes,refs=_tracking(session)
        state=_parse(session,session._execute(state_commands(indexes+refs)))
        identities={row[1]:row[4] for row in state if row[0]=='bookmark'}
        if identities!={handle.bookmark:code for handle,code in indexes+refs}:_malformed(session)
        item=CheckpointSnapshot(state[0][2],state,indexes,refs,getattr(session,'_observed_field_topology',None))
        mapping=getattr(session,'_degradation_checkpoints',{})
        old=mapping.get(item.coordinate)
        if old is not None and (old.state,old.tracked_indexes_prefix,old.tracked_references_prefix)!=(item.state,item.tracked_indexes_prefix,item.tracked_references_prefix):raise ValueError('ambiguous checkpoint')
        mapping[item.coordinate]=old or item
        session._degradation_checkpoints=mapping
        return item.coordinate
    except Exception:
        _fail(session,'LOCAL_MUTATION_CHECKPOINT_FAILED','checkpoint failed')


def _preflight(saved,current):
    target=saved.coordinate
    if current[0][2]<target:raise ValueError('invalid bound')
    # Prefix digest and guard must agree before any deletion. Terminal paragraph
    # geometry can legitimately change while appending to its previous end.
    if current[0][7:11]!=saved.state[0][7:11]:raise ValueError('prefix changed')
    original={row[0]:[r for r in saved.state[1:-1] if r[0]==row[0]] for row in saved.state[1:-1]}
    for kind in ('shape','inline-shape'):
        if [r for r in saved.state if r[0]==kind]!=[r for r in current if r[0]==kind]:raise ValueError('unsupported object change')
    appended={'field':[],'table':[],'bookmark':[]}
    for kind in ('table','field','bookmark','bookmark-hash'):
        old=original.get(kind,[]);now=[r for r in current[1:-1] if r[0]==kind]
        if kind in ('bookmark','bookmark-hash'):
            by_name={r[1]:r for r in now}
            if any(by_name.get(r[1])!=r for r in old):raise ValueError('bookmark topology changed')
            old_names={r[1] for r in old}
            for fact in now:
                if fact[1] not in old_names:
                    if fact[2]<target:raise ValueError('crossing bookmark')
                    appended['bookmark'].append(fact)
        else:
            if now[:len(old)]!=old:raise ValueError('old topology changed')
            for fact in now[len(old):]:
                start,end=(fact[2],fact[3]) if kind=='table' else (fact[5]-1,fact[8]+1)
                if start<target or end>current[0][3]:raise ValueError('crossing object')
                appended[kind].append(fact)
    tables=appended['table']
    if any(a[2]<b[3] and b[2]<a[3] for i,a in enumerate(tables) for b in tables[i+1:]):raise ValueError('overlapping appended tables')
    return appended


def rollback_commands(saved,current,appended):
    tracked=saved.tracked_indexes_prefix+saved.tracked_references_prefix
    lines=state_commands(tracked,saved.coordinate)+[
        f'if not ((current application\'s NSArray\'s arrayWithArray:nativeRows)\'s isEqualToArray:{_literal(current)}) then error "WPSC_CHECKPOINT_PREFLIGHT_CHANGED"',
        'set recoveryDeletedFields to {}','set recoveryDeletedTables to {}',
        f'set recoveryLiveTables to {_literal([r for r in current if r[0]=="table"])}']
    for fact in sorted(appended['field'],key=lambda f:(f[5],f[2]),reverse=True):
        lines += [f'set recoveryField to field {fact[2]} of boundDoc',
                  f'if not ((current application\'s NSString\'s stringWithString:(content of field code of recoveryField as text))\'s isEqualToString:{_literal(fact[4])}) or (start of content of field code of recoveryField) is not {fact[5]} or (end of content of result range of recoveryField) is not {fact[8]} then error "WPSC_CHECKPOINT_FIELD_CHANGED"',
                  'set recoveryBeforeDelete to end of content of text object of boundDoc',
                  'delete recoveryField',
                  f'set end of recoveryDeletedFields to {_literal(fact)}',
                  'set recoveryDelta to recoveryBeforeDelete - (end of content of text object of boundDoc)',
                  'if recoveryDelta < 0 then error "WPSC_CHECKPOINT_FIELD_DELETE_FAILED"',
                  'repeat with recoveryTableFact in recoveryLiveTables',
                  f'if item 3 of recoveryTableFact >= {fact[8]+1} then set item 3 of recoveryTableFact to (item 3 of recoveryTableFact) - recoveryDelta',
                  f'if item 4 of recoveryTableFact >= {fact[8]+1} then set item 4 of recoveryTableFact to (item 4 of recoveryTableFact) - recoveryDelta',
                  'end repeat']
    for fact in appended['bookmark']:
        name=apple_string(fact[1])
        lines += [f'if exists bookmark {name} of boundDoc then delete bookmark {name} of boundDoc']
    original_fields=[r for r in saved.state if r[0]=='field']
    lines += ['set nativeRows to {}',*_field_commands(),f'if not ((current application\'s NSArray\'s arrayWithArray:nativeRows)\'s isEqualToArray:{_literal(original_fields)}) then error "WPSC_CHECKPOINT_UNPLANNED_FIELD_REMOVAL"',
              'set recoveryActualTables to {}','repeat with recoveryOrdinal from 1 to count tables of boundDoc',
              'set recoveryTable to table recoveryOrdinal of boundDoc',
              'set end of recoveryActualTables to {"table",recoveryOrdinal as integer,start of content of text object of recoveryTable,end of content of text object of recoveryTable,count rows of recoveryTable,count columns of recoveryTable}',
              'end repeat','if not ((current application\'s NSArray\'s arrayWithArray:recoveryActualTables)\'s isEqualToArray:recoveryLiveTables) then error "WPSC_CHECKPOINT_TABLE_CHANGED"']
    for fact in sorted(appended['table'],key=lambda t:(t[2],t[1]),reverse=True):
        lines += [f'set recoveryTable to table {fact[1]} of boundDoc',
                  'delete recoveryTable',f'set end of recoveryDeletedTables to {_literal(fact)}']
    lines += ['set recoveryBound to (end of content of text object of boundDoc) - 1',
              f'if recoveryBound < {saved.coordinate} then error "WPSC_CHECKPOINT_BOUND_FAILED"',
              f'set rollbackRange to create range boundDoc start {saved.coordinate} end recoveryBound',
              'set content of rollbackRange to ""',
              *state_commands(tracked,saved.coordinate),
              f'if not ((current application\'s NSArray\'s arrayWithArray:nativeRows)\'s isEqualToArray:{_literal(saved.state)}) then error "WPSC_CHECKPOINT_POSTCONDITION_FAILED"',
              'set nativeRows to {{"rollback-ack",recoveryDeletedFields,recoveryDeletedTables}} & nativeRows']
    return lines


def rollback(session,checkpoint):
    try:
        target=int(checkpoint)
        session._mutation_preflight()
        saved=getattr(session,'_degradation_checkpoints',{}).get(target)
        if saved is None or target<0:raise ValueError('missing preimage')
        indexes,refs=_tracking(session)
        if indexes[:len(saved.tracked_indexes_prefix)]!=saved.tracked_indexes_prefix or refs[:len(saved.tracked_references_prefix)]!=saved.tracked_references_prefix:raise ValueError('tracking prefix changed')
        current=_parse(session,session._execute(state_commands(saved.tracked_indexes_prefix+saved.tracked_references_prefix,target)),target)
        appended=_preflight(saved,current)
        commands=rollback_commands(saved,current,appended)
        # _execute consumes this existing marker only after local script I/O
        # and the final deadline gate, immediately before native submission.
        session._field_topology_mutation_pending=True
        try:
            result=session._execute(commands)
        except BaseException:
            if not session._field_topology_mutation_pending and not session._quarantined:
                session._retain('Destructive rollback completion unverified')
            raise
        else:
            session._invalidate_field_topology()
        finally:
            session._field_topology_mutation_pending=False
        expected=['rollback-ack',[list(f) for f in sorted(appended['field'],key=lambda f:(f[5],f[2]),reverse=True)],[list(t) for t in sorted(appended['table'],key=lambda t:(t[2],t[1]),reverse=True)]]
        if not isinstance(result,list) or not result or result[0]!=expected:_malformed(session)
        if _parse(session,result[1:],target)!=saved.state:_malformed(session)
        session._tracked_indexes=list(saved.tracked_indexes_prefix)
        session._tracked_references=list(saved.tracked_references_prefix)
        if saved.observed_field_topology is None:session.__dict__.pop('_observed_field_topology',None)
        else:session._observed_field_topology=saved.observed_field_topology
        session._pending_heading=None
        session._structural_changed=True
    except Exception:
        _fail(session,'LOCAL_MUTATION_ROLLBACK_FAILED','rollback failed')
