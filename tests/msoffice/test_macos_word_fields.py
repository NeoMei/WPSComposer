"""Direct native field API contracts; transport fakes do not certify Word."""
import inspect
from copy import deepcopy
import subprocess
import sys
import pytest
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, NativeWordError
from skills.WPSComposer.scripts.writer import WriterComposer

METHODS=['insert_toc','insert_toc_with_styles','insert_caption_index_native','insert_figure_index','insert_table_index','refresh_bookmarks_and_references','repaginate_and_update_numbering','repaginate_and_update_page_fields','snapshot_fields']
@pytest.fixture
def session(monkeypatch):
    s=MacWordSession();calls=[]
    def execute(lines,**kwargs):
        calls.append(lines)
        return [['index',10,40,' TOC '+ ('\\c "WPSC_FIG" \\h \\z' if any('WPSC_FIG' in x for x in lines) else '\\o "1-3" \\h \\z')+' \\* MERGEFORMAT ']] if any('"index",' in x for x in lines) else [['ok']]
    monkeypatch.setattr(s,'_execute',execute)
    return s,calls

@pytest.mark.parametrize('name',METHODS)
def test_exact_signatures(name):
    assert hasattr(MacWordSession,name)
    assert inspect.signature(getattr(MacWordSession,name))==inspect.signature(getattr(WriterComposer,name))

@pytest.mark.parametrize('density',[[],{'minFontSizePt':3},{'minFontSizePt':{'toc2':0}}, {'minSpaceAfterPt':{'toc3':float('nan')}},{'minFontSizePt':{'toc1':True}}, {'bad':3}])
def test_density_preflight_before_toc(session,density):
    s,calls=session
    with pytest.raises(ValueError):s.insert_toc_with_styles('Contents',density)
    assert not calls

@pytest.mark.parametrize('name,args,kwargs',[
 ('insert_toc',(),{}),('insert_toc_with_styles',('Contents',{}),{}),
 ('insert_caption_index_native',(),dict(title='Figures',sequence_id='WPSC_FIG',title_style_id='Absent')),
 ('insert_figure_index',(),{}),('insert_table_index',(),{}),
 ('refresh_bookmarks_and_references',(),{}),('repaginate_and_update_numbering',(),{}),('repaginate_and_update_page_fields',(),{}),
])
def test_read_only_no_transport(session,name,args,kwargs):
    s,calls=session;s._read_only=True
    with pytest.raises(ValueError,match='read.only'):getattr(s,name)(*args,**kwargs)
    assert not calls

def test_insert_toc_owned_handle_and_title_semantics(session):
    s,calls=session;handle=s.insert_toc('目录😀');code='\n'.join(calls[0])
    assert handle.kind=='TOC' and handle.owner_node_id=='doc:toc'
    assert 'style body text' in code and 'outline level' in code
    assert '\\\\o' in code and '1-3' in code and 'page break' in code
    assert s._structural_changed

def test_caption_native_validates_and_tracks_category(session):
    s,calls=session
    with pytest.raises(ValueError):s.insert_caption_index_native(title='F',sequence_id='Other',title_style_id='Missing')
    assert not calls
    h=s.insert_caption_index_native(title='F',sequence_id='WPSC_FIG',title_style_id='Missing',owner_node_id='figs')
    assert (h.kind,h.owner_node_id,h.category)==('TOF_FIG','figs','index')
    assert 'WPSC_FIG' in '\n'.join(calls[0]) and 'on error' in '\n'.join(calls[0])

@pytest.mark.parametrize('name,word',[('insert_figure_index','Figure'),('insert_table_index','Table')])
def test_direct_placeholder_is_not_native_index(session,name,word):
    s,calls=session;assert getattr(s,name)('Heading') is None
    code='\n'.join(sum(calls,[]));assert f'[{word} index placeholder]' in code
    assert 'create new field' not in code

def test_exact_native_page_field_types_no_plain_text_match(session):
    s,calls=session;s.repaginate_and_update_page_fields();code='\n'.join(calls[0])
    assert 'field type of ownField is field page' in code and 'field type of ownField is field num pages' in code
    assert 'contains "PAGE"' not in code and 'repaginate boundDoc' in code
    assert 'get story range boundDoc story type' in code

def test_bookmark_health_is_required(session):
    s,calls=session;s.refresh_bookmarks_and_references();code='\n'.join(calls[0])
    assert 'count bookmarks of boundDoc' in code and 'field ref' in code

def test_bad_index_ack_quarantines(session,monkeypatch):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda *_a,**_kw:[])
    with pytest.raises(NativeWordError):s.insert_toc()
    assert s._quarantined

@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary compilation')
def test_all_new_commands_compile(session,tmp_path):
    s,calls=session;s.insert_toc_with_styles('目录😀',{'minFontSizePt':{'toc1':11}})
    s.insert_caption_index_native(title='Figures',sequence_id='WPSC_FIG',title_style_id='Missing')
    s.repaginate_and_update_numbering();s.refresh_bookmarks_and_references();s.repaginate_and_update_page_fields()
    from skills.WPSComposer.scripts.msoffice.macos_word_fields import snapshot_commands
    calls.append(snapshot_commands(s))
    for i,lines in enumerate(calls):
        p=tmp_path/f'{i}.applescript';p.write_text('tell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell')
        r=subprocess.run(['osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True)
        assert r.returncode==0,r.stderr

def field_rows():
    return [['stats',7],['field','story:main/chain:1',1,'REF',' REF b ',12,'cached 😀',0,0,9],['field','story:footer/chain:1',1,'PAGE',' PAGE ',1,'7',0,0,1]]

def test_snapshot_is_shared_privacy_safe_readonly_semantic_tuple(session,monkeypatch):
    from skills.WPSComposer.scripts.longform.executor import FieldSnapshot
    s,calls=session;s._read_only=True
    monkeypatch.setattr(s,'_execute',lambda _lines:field_rows())
    a=s.snapshot_fields();b=s.snapshot_fields()
    assert a==b and isinstance(a,tuple) and all(isinstance(x,FieldSnapshot) for x in a)
    assert a[0].stable_key==('story:main/chain:1','REF',0)
    assert len(a[0].result_hash)==64 and 'cached' not in repr(a)
    assert a[1].field_category=='page' and a[1].total_pages==7

@pytest.mark.parametrize('change',['delete','code','position','reorder'])
def test_observable_untracked_topology_change_fails_stale(session,monkeypatch,change):
    s,calls=session;rows=field_rows();monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    s.snapshot_fields()
    if change=='delete':rows.pop()
    elif change=='code':rows[1][4]=' REF replacement '
    elif change=='position':rows[1][5]+=1
    else:rows[1:]=list(reversed(rows[1:]))
    with pytest.raises(NativeWordError):s.snapshot_fields()

def test_result_refresh_changes_hash_not_semantic_key(session,monkeypatch):
    s,calls=session;rows=field_rows();monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    first=s.snapshot_fields();rows[1][6]='updated😀'
    last=s.snapshot_fields();assert first[0].stable_key==last[0].stable_key and first[0].result_hash!=last[0].result_hash

def test_snapshot_tracks_real_index_page_span_and_rejects_missing_identity(session,monkeypatch):
    s,calls=session;h=s.insert_toc()
    code=s._tracked_indexes[0][1]
    rows=[['stats',9],['identity',h.bookmark,10],['field','main',1,'INDEX',code,10,'Entry\t3\rEntry\t5',2,4,15]]
    monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    a=s.snapshot_fields();assert a[0].toc_page_count==3 and a[0].total_pages==9
    rows.pop(1)
    with pytest.raises(NativeWordError):s.snapshot_fields()

@pytest.mark.parametrize('rows',[[],[['stats',0]],[['stats',2],['field']], [['stats',2],['field','main',1,'PAGE',' PAGE ',2,'1',2,1]]])
def test_snapshot_rejects_malformed_required_readback(session,monkeypatch,rows):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    with pytest.raises(NativeWordError):s.snapshot_fields()

def test_story_enumeration_uses_native_supported_stories_not_broken_collection(session):
    s,calls=session;s.repaginate_and_update_numbering();code='\n'.join(calls[0])
    assert 'count story ranges of boundDoc' not in code
    assert 'get story range boundDoc story type main text story' in code
    assert 'if errorNumber is not -5941 then error' in code
    assert 'if errorNumber is not -2753 then error' in code

def test_cached_result_length_change_keeps_following_semantic_field_identity(session,monkeypatch):
    s,calls=session;rows=[['stats',1],['field','main',1,'REF',' REF a ',1,'x',0,0,1],['field','main',2,'REF',' REF b ',15,'y',0,0,1]]
    monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    first=s.snapshot_fields();rows[1][6]='longer😀';rows[1][9]=len('longer😀'.encode('utf-16-le'))//2;rows[2][5]+=len('longer😀'.encode('utf-16-le'))//2-1
    last=s.snapshot_fields();assert [x.stable_key for x in first]==[x.stable_key for x in last]

def test_stale_field_error_is_closed_public_diagnostic():
    from skills.WPSComposer.scripts.msoffice.errors import NATIVE_WORD_ERROR_CODES
    assert 'NATIVE_WORD_FIELD_IDENTITY_STALE' in NATIVE_WORD_ERROR_CODES

def test_snapshot_page_ranges_do_not_duplicate_native_content(session):
    from skills.WPSComposer.scripts.msoffice.macos_word_fields import snapshot_commands
    s,calls=session;code='\n'.join(snapshot_commands(s))
    assert 'duplicate of' not in code
    assert 'create range boundDoc start nativeStart end nativeStart' in code

def test_index_span_uses_document_ranges_not_unaddressable_story_result_refs(session):
    from skills.WPSComposer.scripts.msoffice.macos_word_fields import snapshot_commands
    s,calls=session;code='\n'.join(snapshot_commands(s))
    assert 'if ownKind is "INDEX" then' in code
    assert 'create range boundDoc start nativeStart end nativeStart' in code

@pytest.mark.parametrize('method',['refresh_bookmarks_and_references','repaginate_and_update_numbering','repaginate_and_update_page_fields'])
def test_refresh_invalidates_cached_targets_before_native_submission(session,monkeypatch,method):
    s,calls=session
    def execute(_lines):
        assert s._structural_changed
        return [['ok']]
    monkeypatch.setattr(s,'_execute',execute)
    getattr(s,method)()

def test_snapshot_reads_native_field_type_once_per_field(session):
    from skills.WPSComposer.scripts.msoffice.macos_word_fields import snapshot_commands
    s,calls=session;code='\n'.join(snapshot_commands(s))
    assert 'set ownType to field type of ownField' in code
    assert 'if ownType is field toc' in code

def test_native_stale_marker_maps_to_closed_identity_error(monkeypatch,tmp_path):
    from skills.WPSComposer.scripts.msoffice import macos_word_session as module
    s=MacWordSession();s.staging_root=tmp_path
    monkeypatch.setattr(module.subprocess,'run',lambda *a,**kw:subprocess.CompletedProcess(a,1,'','execution error: WPSC_FIELD_IDENTITY_STALE (-2700)'))
    with pytest.raises(NativeWordError) as caught:s._execute([],bind=False)
    assert caught.value.code=='NATIVE_WORD_FIELD_IDENTITY_STALE'

def test_tracked_toc_growth_compensates_following_ref_with_nested_native_extent(session,monkeypatch):
    s,calls=session;h=s.insert_toc();code=s._tracked_indexes[0][1]
    rows=[['stats',3],['identity',h.bookmark,10],['field','main',1,'INDEX',code,10,'A',1,1,80],['field','main',4,'REF',' REF target ',120,'X',0,0,1]]
    monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    before=s.snapshot_fields();rows[2][6]='ABCDE';rows[2][9]+=40;rows[3][2]+=2;rows[3][5]+=40
    after=s.snapshot_fields();assert [x.stable_key for x in before]==[x.stable_key for x in after]

@pytest.mark.parametrize('first,last,total',[(1,99,2),(3,3,2),(0,1,2)])
def test_snapshot_rejects_impossible_index_page_bounds(session,monkeypatch,first,last,total):
    s,calls=session;h=s.insert_toc();code=s._tracked_indexes[0][1]
    rows=[['stats',total],['identity',h.bookmark,10],['field','main',1,'INDEX',code,10,'A',first,last,1]]
    monkeypatch.setattr(s,'_execute',lambda _lines:rows)
    with pytest.raises(NativeWordError) as error:s.snapshot_fields()
    assert error.value.code=='NATIVE_WORD_EXECUTION_FAILED'


def test_fixture_rejects_duplicate_semantic_snapshot_and_preserves_every_source(tmp_path):
    import importlib.util
    from pathlib import Path
    p=Path('fixtures/microsoft_parity/macos_word_fields.py')
    spec=importlib.util.spec_from_file_location('fields_fixture_review',p);fixture=importlib.util.module_from_spec(spec);spec.loader.exec_module(fixture)
    assert hasattr(fixture,'validate_snapshots') and hasattr(fixture,'retain_sources')
    from collections import Counter
    expected=Counter({'STYLEREF':1,'SEQ_FIG':1,'SEQ_TAB':1,'REF':1,'TOC':2,'TOF_FIG':1,'TOF_TAB':1,'PAGE':1,'NUMPAGES':1})
    rows=[]
    for kind,count in expected.items():
        owner='doc:toc' if kind=='TOC' else 'figures' if kind=='TOF_FIG' else 'tables' if kind=='TOF_TAB' else 'story:primary footer/chain:1' if kind in {'PAGE','NUMPAGES'} else 'story:main text/chain:1'
        for ordinal in range(count):rows.append({'stable_key':[owner,kind,ordinal],'field_category':'index' if kind in {'TOC','TOF_FIG','TOF_TAB'} else 'page' if kind in {'PAGE','NUMPAGES'} else 'field'})
    fixture.validate_snapshots(rows)
    with pytest.raises(AssertionError):fixture.validate_snapshots(rows+[rows[0]])
    sources=[fixture.ROOT/'fixtures/microsoft_parity/macos_word_fields.py',fixture.ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_fields.py']
    hashes=fixture.retain_sources(tmp_path,sources)
    import hashlib
    for relative,digest in hashes.items():assert hashlib.sha256((tmp_path/'source'/relative).read_bytes()).hexdigest()==digest

def test_malformed_snapshot_retains_native_evidence(session,monkeypatch):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda _lines:[['stats',2],['field','main',1,'PAGE',' PAGE ',1,'1',1,1,1]])
    with pytest.raises(NativeWordError) as error:s.snapshot_fields()
    assert error.value.code=='NATIVE_WORD_EXECUTION_FAILED' and s._retain_evidence


def test_owned_reference_insertion_rebases_field_topology(session, monkeypatch):
    s, calls = session
    bookmark = 'wpsc_fig_' + 'a' * 24
    rows = [['stats', 3], ['field', 'main', 1, 'REF', ' REF existing ', 1, '1', 0, 0, 1]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    original = s.snapshot_fields()

    def insert(lines):
        identity = next(line for line in lines if 'make new bookmark' in line).split('name:"', 1)[1].split('"', 1)[0]
        return [['literal', 0, 20, 24, '见😀('],
                ['reference', 0, 1, 2, 25, 66, 67, f' REF {bookmark} \\h ', identity],
                ['literal', 0, 68, 70, ')尾'], ['literal', 1, 70, 71, '\r'],
                ['complete', 1, 2, 20, 71]]

    monkeypatch.setattr(s, '_execute', insert)
    s.add_cross_reference_paragraph(runs=[{
        'type': 'reference', 'bookmarkName': bookmark, 'prefix': '见😀(',
        'suffix': ')尾', 'fallbackText': '静态',
    }], owner_node_id='owner:new')
    handle, code = s._tracked_references[0]
    rows += [['identity', handle.bookmark, 25],
             ['field', 'main', 2, 'REF', code, 25, '2', 0, 0, 1]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))

    current = s.snapshot_fields()
    assert [item.stable_key for item in current] == [original[0].stable_key, ('owner:new', 'REF', 0)]


def test_owned_toc_and_structural_insert_rebase_field_topology(session, monkeypatch):
    s, calls = session
    rows = [['stats', 3], ['field', 'main', 1, 'REF', ' REF existing ', 12, '1', 0, 0, 1]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    original = s.snapshot_fields()
    monkeypatch.setattr(s, '_execute', lambda _lines: [['index', 30, 60, ' TOC \\o "1-3" \\h \\z \\* MERGEFORMAT ']])
    handle = s.insert_toc()
    code = s._tracked_indexes[0][1]
    rows += [['identity', handle.bookmark, 30],
             ['field', 'main', 2, 'INDEX', code, 30, 'Heading\t1', 1, 1, 20]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    after_toc = s.snapshot_fields()
    assert [item.stable_key for item in after_toc] == [original[0].stable_key, ('doc:toc', 'TOC', 0)]

    monkeypatch.setattr(s, '_execute', lambda _lines: [['ok']])
    s.apply_structural_op({'op': 'insert', 'type': 'paragraph', 'position': 'start', 'props': {'text': 'prefix'}})
    rows[1][5] += 7
    rows[2][2] += 7
    rows[3][5] += 7
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    assert [item.stable_key for item in s.snapshot_fields()] == [item.stable_key for item in after_toc]


def test_same_paragraph_text_replacement_rebases_later_field_position(session, monkeypatch):
    s, calls = session
    rows = [['stats', 1], ['field', 'main', 1, 'REF', ' REF later ', 20, '1', 0, 0, 1]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    original = s.snapshot_fields()
    monkeypatch.setattr(s, '_execute', lambda _lines: [['ok']])
    assert s.apply_format_patch('paragraph:1', text='longer text')['rejected'] == []
    rows[1][5] = 27
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(rows))
    assert s.snapshot_fields()[0].stable_key == original[0].stable_key


def test_field_refresh_keeps_drift_guard_for_position_and_code(session, monkeypatch):
    s, calls = session
    baseline = [['stats', 1], ['field', 'main', 1, 'REF', ' REF target ', 20, '1', 0, 0, 1]]
    monkeypatch.setattr(s, '_execute', lambda _lines: deepcopy(baseline))
    s.snapshot_fields()
    monkeypatch.setattr(s, '_execute', lambda _lines: [['ok']])
    s.refresh_bookmarks_and_references()

    for changed in ('position', 'code'):
        rows = deepcopy(baseline)
        if changed == 'position':
            rows[1][5] += 1
        else:
            rows[1][4] = ' REF replacement '
        monkeypatch.setattr(s, '_execute', lambda _lines, rows=rows: rows)
        with pytest.raises(NativeWordError) as caught:
            s.snapshot_fields()
        assert caught.value.code == 'NATIVE_WORD_FIELD_IDENTITY_STALE'


def test_header_field_topology_change_preserves_pending_body_heading(session):
    s, calls = session
    s._observed_field_topology = (('main', 0, 'REF', 'digest', 1, 1),)
    s._pending_heading = (10, 20, 2)
    s.set_header('header')
    assert not hasattr(s, '_observed_field_topology')
    assert s._pending_heading == (10, 20, 2)
    assert not s._structural_changed


def test_header_link_change_invalidates_field_topology(session):
    s, calls = session
    s._observed_field_topology = (('header', 0, 'PAGE', 'digest', 1, 1),)
    s.set_header_footer(link_to_previous_header=True)
    assert not hasattr(s, '_observed_field_topology')


@pytest.mark.parametrize('operation', ['structural', 'business', 'format'])
def test_local_preflight_failure_keeps_field_topology(session, monkeypatch, operation):
    s, calls = session
    baseline = (('main', 0, 'REF', 'digest', 1, 1),)
    s._observed_field_topology = baseline
    monkeypatch.setattr(s, '_remaining', lambda: (_ for _ in ()).throw(RuntimeError('preflight')))
    with pytest.raises(RuntimeError, match='preflight'):
        if operation == 'structural':
            s._execute_structural(['set nativeRows to {{"ok"}}'])
        elif operation == 'business':
            s._business_commit(['set content of text object of boundDoc to "x"'], structural=True)
        else:
            s.apply_format_patch('paragraph:1', text='longer')
    assert s._observed_field_topology == baseline


def test_transport_pre_submission_failure_keeps_field_topology(session, monkeypatch, tmp_path):
    import types
    s, calls = session
    baseline = (('main', 0, 'REF', 'digest', 1, 1),)
    s._observed_field_topology = baseline
    s._retain_evidence = True
    s.staging_root = tmp_path
    s._owns_doc = True
    s._bound_path = str(tmp_path / 'bound.docx')
    s._execute = types.MethodType(MacWordSession._execute, s)
    remaining_calls = 0
    def remaining():
        nonlocal remaining_calls
        remaining_calls += 1
        if remaining_calls == 3:
            raise RuntimeError('before subprocess')
        return 30
    monkeypatch.setattr(s, '_remaining', remaining)
    with pytest.raises(RuntimeError, match='before subprocess'):
        s._execute_topology_mutation(['set nativeRows to {{"ok"}}'])
    assert s._observed_field_topology == baseline
