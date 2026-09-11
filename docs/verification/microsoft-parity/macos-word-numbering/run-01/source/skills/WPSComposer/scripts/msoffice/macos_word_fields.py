"""Bound Word field primitives, with code-bookmark identities for owned indexes.

Inspection never adds bookmarks. Fields discovered in an existing document have
conservative observational identities; a changed native topology is stale, not
permission to silently retarget an earlier ordinal.
"""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
from uuid import uuid4

from ..longform.field_contract import snapshot_visible_field
from .errors import NativeWordError
from .macos_script import apple_string


@dataclass(frozen=True)
class NativeIndexHandle:
    """Session-local semantic identity; never a COM/native object pointer."""
    session_id: str
    bookmark: str
    owner_node_id: str
    kind: str
    category: str = 'index'


def _error(session, code='NATIVE_WORD_FIELD_IDENTITY_STALE'):
    session._retain_evidence = True
    raise NativeWordError(code, staging_path=session.staging_root)


def _ack(session, rows, expected):
    if rows != expected:
        session._retain('Native field acknowledgement invalid')
        _error(session, 'NATIVE_WORD_EXECUTION_FAILED')


def _density_commands(density):
    if density is None:
        density = {}
    if not isinstance(density, dict):
        raise ValueError('density must be a mapping')
    allowed = {'minFontSizePt':'font size of font object', 'minSpaceBeforePt':'space before of paragraph format', 'minSpaceAfterPt':'space after of paragraph format'}
    # Other baseline density metadata is inert here, but its shapes are checked.
    metadata = {'tocTitle','levels','includeFigureIndex','includeTableIndex','figureIndexTitle','tableIndexTitle'}
    if set(density) - set(allowed) - metadata:
        raise ValueError('Unknown density property')
    lines=[]
    from .macos_word_session import _number
    for key, value in density.items():
        if key in allowed:
            if not isinstance(value, dict) or set(value)-{'toc1','toc2','toc3'}:
                raise ValueError('Density minima must map TOC levels')
            for level, amount in value.items():
                if amount is None:
                    continue
                rendered=_number(amount)
                if amount < 0 or (key=='minFontSizePt' and amount==0):
                    raise ValueError('Invalid density minimum')
                lines += [f'set ownStyle to Word style (style {level}) of boundDoc', f'set {allowed[key]} of ownStyle to {rendered}']
        elif key=='levels':
            if type(value) is not int or not 1<=value<=3:raise ValueError('Invalid density levels')
        elif key.startswith('include'):
            if type(value) is not bool:raise ValueError('Invalid density boolean')
        elif not isinstance(value,str):
            raise ValueError('Invalid density title')
    return lines


def insert_index(session, title, *, kind='TOC', sequence_id=None, title_style_id=None, owner_node_id=None, density=None):
    session._writable()
    formatting=_density_commands(density)
    if kind!='TOC' and sequence_id not in {'WPSC_FIG','WPSC_TAB'}:
        raise ValueError('invalid native caption index sequence')
    if owner_node_id is not None and (not isinstance(owner_node_id,str) or not owner_node_id):
        raise ValueError('owner_node_id must be nonempty text')
    if title and kind=='TOC' and not isinstance(title,str):
        raise ValueError('TOC title must be text')
    lines=[]
    if title:
        if kind=='TOC':
            lines=session._business_paragraph(title,'Body Text',{'size':18,'bold':False,'color':0,'align':1,'line_spacing':18,'space_after':5},suffix=['set outline level of paragraph format of semanticRange to outline level body text'])
        else:
            # Resolve requested native style (including built-in numeric IDs),
            # then only the documented safe fallback on missing style.
            if isinstance(title_style_id,str):requested=apple_string(title_style_id)
            elif type(title_style_id) is int:requested=str(title_style_id)
            else:raise ValueError('title_style_id must be a native style name or ID')
            lines=session._business_paragraph(str(title),'Body Text')
            index=lines.index('set style of semanticRange to semanticStyle')
            lines[index:index]=['try',f'set semanticStyle to Word style {requested} of boundDoc','on error','set semanticStyle to Word style (style body text) of boundDoc','end try']
    bookmark='WPSC_F_'+uuid4().hex[:30]
    code='\\o "1-3" \\h \\z' if kind=='TOC' else '\\c "'+sequence_id+'" \\h \\z'
    lines += formatting+session._position('end')+session._paragraph_boundary()+[
        'set indexPoint to insertionPoint',
        'set indexRange to create range boundDoc start indexPoint end indexPoint',
        f'create new field text range indexRange field type field toc field text {apple_string(code)} preserve formatting true',
        'set ownField to missing value',
        'repeat with fi from 1 to count fields of boundDoc',
        'set candidateField to field fi of boundDoc',
        'if (start of content of field code of candidateField) is indexPoint + 1 then set ownField to candidateField',
        'end repeat',
        'if ownField is missing value then error "WPSC_INDEX_CREATION_UNVERIFIED"',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(bookmark)}, text object:field code of ownField}}',
        'if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"',
        'set codeStart to start of content of field code of ownField',
        'set resultEnd to end of content of result range of ownField',
        'set insertionPoint to (end of content of text object of boundDoc) - 1',
        'set terminalRange to create range boundDoc start insertionPoint end insertionPoint',
        *(['insert break at terminalRange break type page break'] if kind=='TOC' else ['set content of terminalRange to return']),
        'set nativeRows to {{"index",codeStart,resultEnd,content of field code of ownField as text}}',
    ]
    rows=session._execute_structural(lines)
    if len(rows)!=1 or len(rows[0])!=4 or rows[0][0]!='index' or any(type(v) is not int for v in rows[0][1:3]) or rows[0][1]<0 or rows[0][2]<rows[0][1]:
        session._retain('Native index acknowledgement invalid');_error(session,'NATIVE_WORD_EXECUTION_FAILED')
    if not isinstance(rows[0][3],str) or rows[0][3].strip().removesuffix(' \\* MERGEFORMAT') != 'TOC '+code:
        session._retain('Native index field code invalid');_error(session,'NATIVE_WORD_EXECUTION_FAILED')
    if not hasattr(session,'_field_session_id'):session._field_session_id=uuid4().hex
    handle=NativeIndexHandle(session._field_session_id,bookmark,owner_node_id or ('doc:toc' if kind=='TOC' else 'doc:index'),kind)
    if not hasattr(session,'_tracked_indexes'):session._tracked_indexes=[]
    session._tracked_indexes.append((handle,rows[0][3]))
    return handle


def placeholder(session,title,word):
    session._writable()
    lines=[]
    if title:lines+=session._business_paragraph(str(title),'Normal',{'size':14,'bold':True})
    lines+=session._business_paragraph(f'[{word} index placeholder]','Body Text')
    _ack(session,session._execute_structural(lines+['set nativeRows to {{"ok"}}']),[['ok']])


def _story_loop(body):
    # Word advertises story ranges but its collection count throws -1708.
    # Query the closed native story enum; -5941 means this story is absent.
    stories = ['main text', 'footnotes', 'endnotes', 'comments', 'text frame',
               'even pages header', 'primary header', 'even pages footer',
               'primary footer', 'first page header', 'first page footer',
               'footnote separator', 'footnote continuation separator ',
               'footnote continuation notice', 'endnote separator',
               'endnote continuation separator ', 'endnote continuation notice']
    lines=[]
    for story in stories:
        lines += ['set storyAvailable to false', 'set ownStory to missing value', 'try',
                  f'set ownStory to get story range boundDoc story type ' + ({'footnote continuation separator ': '13', 'endnote continuation separator ': '16'}.get(story, story+' story')),
                  'on error errorMessage number errorNumber',
                  'if errorNumber is not -5941 then error errorMessage number errorNumber',
                  'end try', 'try',
                  'set storyAvailable to (ownStory is not missing value)',
                  'on error errorMessage number errorNumber',
                  'if errorNumber is not -2753 then error errorMessage number errorNumber',
                  'end try', 'set chainIndex to 0',
                  'repeat while storyAvailable',
                  'set chainIndex to chainIndex + 1',
                  'if chainIndex > 10000 then error "WPSC_STORY_CHAIN_INVALID"',
                  f'set storyOwner to "story:{story}/chain:" & chainIndex',
                  'repeat with fieldIndex from 1 to count fields of ownStory',
                  'set ownField to field fieldIndex of ownStory', *body, 'end repeat',
                  'set storyAvailable to false', 'try',
                  'set ownStory to next story range of ownStory',
                  'on error errorMessage number errorNumber',
                  'if errorNumber is not -5941 then error errorMessage number errorNumber',
                  'end try', 'try',
                  'set storyAvailable to (ownStory is not missing value)',
                  'on error errorMessage number errorNumber',
                  'if errorNumber is not -2753 then error errorMessage number errorNumber',
                  'end try', 'end repeat']
    return lines


def refresh(session,phase):
    session._writable()
    conditions={'numbering':'field type of ownField is field style ref or field type of ownField is field sequence', 'references':'field type of ownField is field ref', 'page':'field type of ownField is field page or field type of ownField is field num pages'}
    lines=['set bookmarkHealth to count bookmarks of boundDoc'] if phase=='references' else ['repaginate boundDoc']
    lines+=_story_loop([f'if {conditions[phase]} then','if (update field ownField) is false then error "WPSC_FIELD_UPDATE_FAILED"','end if'])
    _ack(session,session._execute_structural(lines+['set nativeRows to {{"ok"}}'], field_topology_change=False),[['ok']])


def snapshot_commands(session):
    lines=['set nativeRows to {{"stats",compute statistics boundDoc statistic statistic pages}}']
    for handle,code in list(getattr(session,'_tracked_indexes',[])) + list(getattr(session,'_tracked_references',[]))+list(getattr(session,'_tracked_numbering',[])):
        lines += [f'if not (exists bookmark {apple_string(handle.bookmark)} of boundDoc) then error "WPSC_FIELD_IDENTITY_STALE"',f'set identityRange to text object of bookmark {apple_string(handle.bookmark)} of boundDoc',f'if (content of identityRange as text) is not {apple_string(code)} then error "WPSC_FIELD_IDENTITY_STALE"',f'set end of nativeRows to {{"identity",{apple_string(handle.bookmark)},start of content of identityRange}}']
    body=[
        'set ownKind to ""', 'set ownType to field type of ownField',
        'if ownType is field page then set ownKind to "PAGE"',
        'if ownType is field num pages then set ownKind to "NUMPAGES"',
        'if ownType is field ref then set ownKind to "REF"',
        'if ownType is field style ref then set ownKind to "STYLEREF"',
        'set codeText to ""',
        'if ownKind is not "" or ownType is field sequence or ownType is field toc then set codeText to content of field code of ownField as text',
        'if ownType is field sequence then',
        'if (word 2 of codeText as text) is "WPSC_FIG" then set ownKind to "SEQ_FIG"',
        'if (word 2 of codeText as text) is "WPSC_TAB" then set ownKind to "SEQ_TAB"',
        'if (word 2 of codeText as text) is "WPSC_EQ" then set ownKind to "SEQ_EQ"','end if',
        'if ownType is field toc then set ownKind to "INDEX"',
        'if ownKind is not "" then',
        'set resultRange to result range of ownField',
        'set nativeStart to start of content of resultRange',
        'set nativeEnd to end of content of resultRange',
        'set firstPage to 0', 'set lastPage to 0',
        'if ownKind is "INDEX" then',
        'set firstRange to create range boundDoc start nativeStart end nativeStart',
        'set lastPoint to nativeStart',
        'if nativeEnd > nativeStart then set lastPoint to nativeEnd - 1',
        'set lastRange to create range boundDoc start lastPoint end lastPoint',
        'set firstPage to (get range information firstRange information type active end page number) as integer',
        'set lastPage to (get range information lastRange information type active end page number) as integer', 'end if',
        'set end of nativeRows to {"field",storyOwner,fieldIndex,ownKind,codeText,start of content of field code of ownField,content of resultRange as text,firstPage,lastPage,nativeEnd - nativeStart}',
        'end if',
    ]
    return lines+_story_loop(body)


def snapshot(session):
    rows=session._execute(snapshot_commands(session))
    if not rows or len(rows[0])!=2 or rows[0][0]!='stats' or type(rows[0][1]) is not int or rows[0][1]<1:_error(session,'NATIVE_WORD_EXECUTION_FAILED')
    total=rows[0][1];identities={};fields=[]
    for row in rows[1:]:
        if len(row)==3 and row[0]=='identity' and isinstance(row[1],str) and type(row[2]) is int:
            if row[1] in identities:_error(session)
            identities[row[1]]=row[2]
        elif len(row)==10 and row[0]=='field' and all(isinstance(row[i],str) for i in (1,3,4,6)) and all(type(row[i]) is int for i in (2,5,7,8,9)) and row[2]>0 and row[5]>=0 and row[9]>=0 and ((1<=row[7]<=row[8]<=total) if row[3]=='INDEX' else row[7]==row[8]==0):fields.append(row)
        else:_error(session,'NATIVE_WORD_EXECUTION_FAILED')
    tracked=list(getattr(session,'_tracked_indexes',[]))+list(getattr(session,'_tracked_references',[]))+list(getattr(session,'_tracked_numbering',[]));by_position={}
    # All tracked creators bind a main-story range. Offsets in another story
    # are not the same position, even when a header repeats the same field.
    tracked_ordinals={};owner_ordinals={}
    for handle,code in tracked:
        if handle.bookmark not in identities:_error(session)
        matches=[r for r in fields if r[1]=='story:main text/chain:1' and r[3]==(handle.kind if handle.kind=='REF' or handle.category=='numbering' else 'INDEX') and r[5]==identities[handle.bookmark] and r[4]==code]
        if len(matches)!=1:_error(session)
        by_position[(matches[0][1],matches[0][2])]=handle
        key=(handle.owner_node_id,handle.kind)
        ordinal=owner_ordinals.get(key,0);owner_ordinals[key]=ordinal+1
        tracked_ordinals[handle.bookmark]=ordinal
    counts={'TOC':0,'TOF_FIG':0,'TOF_TAB':0};resolved=[];observed=[];story_ordinals={}
    for row in fields:
        _,owner,index,kind,code,start,visible,first,last,extent=row
        ordinal=story_ordinals.get(owner,0);story_ordinals[owner]=ordinal+1
        observed.append((owner,ordinal,kind,hashlib.sha256(code.encode()).hexdigest(),start,extent))
        handle=by_position.get((owner,index))
        if handle:owner,kind,category=handle.owner_node_id,handle.kind,handle.category
        elif kind=='INDEX':continue  # shared baseline tracks only inserted indexes
        else:
            category='page' if kind in {'PAGE','NUMPAGES'} else 'field'
        if kind in counts:counts[kind]+=last-first+1
        resolved.append((owner,kind,category,visible,tracked_ordinals[handle.bookmark] if handle else None))
    previous=getattr(session,'_observed_field_topology',None)
    if previous is not None:
        if len(observed)!=len(previous):_error(session)
        shifts={}
        for before,after in zip(previous,observed):
            if before[:4]!=after[:4]:_error(session)
            story=after[0]
            if after[4]!=before[4]+shifts.get(story,0):_error(session)
            shifts[story]=shifts.get(story,0)+after[5]-before[5]
    session._observed_field_topology=tuple(observed)
    ordinals={};snapshots=[]
    for owner,kind,category,visible,tracked_ordinal in resolved:
        key=(owner,kind)
        ordinal=tracked_ordinal if tracked_ordinal is not None else ordinals.get(key,0)
        ordinals[key]=ordinal+1
        snapshots.append(snapshot_visible_field(owner_node_id=owner,field_kind=kind,ordinal_within_node=ordinal,visible_result=visible,field_category=category,toc_page_count=counts['TOC'],figure_index_page_count=counts['TOF_FIG'],table_index_page_count=counts['TOF_TAB'],total_pages=total))
    return tuple(snapshots)
