"""Frozen pagination semantics; only the native AppleEvent transport is replaced."""
from __future__ import annotations

import importlib
import json
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_degradation import NativeDegradationRange
from skills.WPSComposer.scripts.writer import NativeWriterObjectError

NAME = 'skills.WPSComposer.scripts.msoffice.macos_word_pagination'


def pagination():
    assert importlib.util.find_spec(NAME), 'Missing Mac Word pagination implementation'
    return importlib.import_module(NAME)


def session_with(rows):
    session = MacWordSession()
    session._read_only = True
    session.commands = []
    def execute(lines):
        session.commands.append(list(lines))
        return rows
    session._execute = execute
    return session


def setup(width=600, height=800, left=70, right=60, top=50, bottom=40):
    return ['setup', width, height, left, right, top, bottom]


def tracked(node='节点😀', start=2, end=25, **values):
    return dict(nodeId=node, range=SimpleNamespace(Start=start, End=end), **values)


def test_bookmark_whole_paragraph_geometry_and_node_redaction():
    p = pagination()
    session = session_with([['bookmark', 2, 25, 3, 72, 91.5]])
    result = p.pagination_fragment_for_bookmark(session, '/private/client/secret.docx', 'WPSC_Unicode')
    assert result == {'nodeId':'<redacted>', 'story':'main', 'sections':['body'],
                      'pageStart':3, 'pageEnd':3, 'range':'2:25',
                      'fragments':[{'page':3,'bounds':[72.0,91.5,73.0,103.5]}]}
    assert '/private/client' not in json.dumps(session.commands)


@pytest.mark.parametrize('x,y', [(-1, 10), (10, -1), (float('nan'), 2), (float('inf'), 2)])
def test_bookmark_unavailable_coordinates_omit_bounds(x, y):
    p = pagination()
    session = session_with([['bookmark', 0, 1, 1, x, y]])
    assert p.pagination_fragment_for_bookmark(session, None, 'B')['fragments'] == [{'page':1}]


def test_map_visual_three_pages_has_native_first_middle_last_geometry():
    p = pagination()
    session = session_with([setup(), ['range', 0, 1, 3, 80, 100, 120]])
    result = p.pagination_map_for_ranges(session, [tracked(op='writer.add_heading')])
    assert result == {'version':'M5-v1','nodes':[{
        'nodeId':'节点😀','story':'main','sections':['body'],'pageStart':1,'pageEnd':3,
        'range':'2:25','fragments':[{'page':1,'bounds':[80.0,100.0,540.0,760.0]},
                                 {'page':2,'bounds':[70.0,50.0,540.0,760.0]},
                                 {'page':3,'bounds':[70.0,50.0,540.0,132.0]}]}]}


def test_map_skips_falsy_and_duplicate_ids_before_accessing_range_and_coerces_values():
    p = pagination()
    session = session_with([setup(), ['range', 0, 2, 2, 80, 100, 100],
                           ['range', 1, 2, 2, 80, 100, 100]])
    values = [dict(nodeId=0), dict(nodeId=None), tracked(7, '2', 3.9, role=42),
              dict(nodeId='7'), tracked('next', True, '2', role='')]
    result = p.pagination_map_for_ranges(session, iter(values))
    assert [(n['nodeId'],n['range'],n['sections']) for n in result['nodes']] == [
        ('7','2:3',['42']),('next','1:2',['body'])]
    assert all(n['fragments']==[{'page':2}] for n in result['nodes'])


def test_map_deduplicates_original_node_ids_before_redaction():
    p = pagination()
    session = session_with([setup(), ['range', 0, 1, 1, 0, 0, 0],
                           ['range', 1, 1, 1, 0, 0, 0]])
    result = p.pagination_map_for_ranges(session, [tracked('/private/a.docx'), tracked('/private/b.docx')])
    assert [n['nodeId'] for n in result['nodes']] == ['<redacted>', '<redacted>']
    assert '/private/' not in json.dumps(result)


@pytest.mark.parametrize('page_setup,points', [
    (setup(width=99), (80,100,120)), (setup(right=700), (80,100,120)),
    (setup(), (-1,100,120)), (setup(), (80,-1,120)),
    (setup(), (80,100,801)), (setup(), (float('nan'),100,120)),
])
def test_map_unusable_native_geometry_keeps_page_span(page_setup, points):
    p = pagination()
    session = session_with([page_setup, ['range', 0, 1, 2, *points]])
    result = p.pagination_map_for_ranges(session, [tracked(op='writer.add_equation')])
    assert result['nodes'][0]['fragments'] == [{'page':1},{'page':2}]


def test_map_negative_margins_clamped_and_bounds_have_minimum_extent():
    p = pagination()
    session = session_with([setup(left=-20,right=-5,top=-1,bottom=-2),
                           ['range',0,1,1,600,800,780]])
    result = p.pagination_map_for_ranges(session, [tracked(op='writer.add_semantic_table')])
    assert result['nodes'][0]['fragments'] == [{'page':1,'bounds':[600.0,800.0,601.0,801.0]}]


def test_empty_map_still_gets_native_setup_and_repagination():
    p = pagination()
    session = session_with([setup()])
    assert p.pagination_map_for_ranges(session, []) == {'version':'M5-v1','nodes':[]}
    assert session.commands and session.commands[0][0] == 'repaginate boundDoc'


@pytest.mark.parametrize('rows', [[], [setup()], [setup(), ['range',0,0,1,0,0,0]],
    [setup(), ['range',0,3,2,0,0,0]], [setup(), ['range',9,1,1,0,0,0]],
    [setup(), ['range',0,1,1,'private text',0,0]]])
def test_bad_native_map_acknowledgement_is_typed_and_private(rows):
    p = pagination()
    session = session_with(rows)
    with pytest.raises(NativeWriterObjectError) as error:
        p.pagination_map_for_ranges(session, [tracked()])
    assert error.value.code == 'PAGINATION_SNAPSHOT_FAILED'
    assert str(error.value) == 'pagination snapshot failed'


@pytest.mark.parametrize('value', [tracked(start=-1),tracked(start=3,end=2),
                                   tracked(start='private failure'), {'nodeId':'x'}])
def test_invalid_ranges_fail_locally_without_native_submission(value):
    p = pagination()
    session = session_with([])
    with pytest.raises(NativeWriterObjectError) as error:
        p.pagination_map_for_ranges(session, [value])
    assert error.value.code == 'PAGINATION_SNAPSHOT_FAILED'
    assert session.commands == []


def test_existing_tagged_range_rejects_other_session_but_duck_id_is_not_a_schema():
    p = pagination()
    session = session_with([setup(), ['range',0,1,1,0,0,0]])
    session._degradation_session_id = 'owned'
    foreign = NativeDegradationRange('other',0,2,'😀')
    with pytest.raises(NativeWriterObjectError):
        p.pagination_map_for_ranges(session, [dict(nodeId='x',range=foreign)])
    assert session.commands == []
    duck = SimpleNamespace(Start=0,End=2,session_id='other')
    assert p.pagination_map_for_ranges(session,[dict(nodeId='x',range=duck)])['nodes'][0]['range']=='0:2'


@pytest.mark.parametrize('method,args', [('pagination_fragment_for_bookmark',('node','B')),
                                       ('pagination_map_for_ranges',([],))])
def test_real_session_lifecycle_rejects_closed_and_quarantined_without_transport(method,args):
    p = pagination()
    for state in ('_closed','_quarantined'):
        session = MacWordSession()
        setattr(session,state,True)
        with pytest.raises(NativeWriterObjectError) as error:
            getattr(p,method)(session,*args)
        assert error.value.code=='PAGINATION_SNAPSHOT_FAILED'


def test_bookmark_native_ordinal_is_accepted_like_frozen_collection_lookup():
    p = pagination()
    session = session_with([['bookmark',0,2,1,0,0]])
    assert p.pagination_fragment_for_bookmark(session,'😀',1)['range']=='0:2'
    assert session.commands[0][0]=='set paginationBookmark to bookmark 1 of boundDoc'


def test_expired_real_session_maps_deadline_failure_without_starting_office():
    p = pagination()
    session = MacWordSession()
    session._deadline = 0
    with pytest.raises(NativeWriterObjectError) as error:
        p.pagination_map_for_ranges(session, [])
    assert error.value.code=='PAGINATION_SNAPSHOT_FAILED'


@pytest.mark.parametrize('method,args', [('pagination_fragment_for_bookmark',('node','B')),
                                       ('pagination_map_for_ranges',([],))])
def test_native_errors_preserve_privacy_without_swallowing_interrupts(method,args):
    p = pagination()
    session = MacWordSession()
    def fail(lines):
        raise RuntimeError('private full document text /private/client.docx')
    session._execute = fail
    with pytest.raises(NativeWriterObjectError) as error:
        getattr(p,method)(session,*args)
    assert str(error.value)=='pagination snapshot failed'
    assert error.value.__suppress_context__ is True
    def interrupt(lines):
        raise KeyboardInterrupt()
    session._execute = interrupt
    with pytest.raises(KeyboardInterrupt):
        getattr(p,method)(session,*args)


def test_native_owned_range_uses_utf16_offsets_without_accessing_full_text():
    p = pagination()
    session = session_with([setup(),['range',0,1,1,0,0,0]])
    session._degradation_session_id = 'owned'
    handle = NativeDegradationRange('owned',4,6,'😀 private payload')
    result = p.pagination_map_for_ranges(session,[dict(nodeId='emoji',range=handle)])
    assert result['nodes'][0]['range']=='4:6'
    assert 'private payload' not in json.dumps(result)
    assert 'private payload' not in json.dumps(session.commands)


@pytest.mark.parametrize('failure', [None,'owned-remains','sentinel-changed','close-failed'])
def test_fixture_waits_for_independent_owned_and_reopen_close_before_sentinel(monkeypatch,tmp_path,failure):
    """Catches close-before-exit and trusting _closed despite native leftovers."""
    from fixtures.microsoft_parity import macos_word_pagination as fixture
    import pdfplumber
    events = []
    starting = [['User document','/user.docx',False,'user-hash']]
    sentinel_row = ['Synthetic sentinel','Synthetic sentinel',False,'sentinel-hash']
    native_documents = list(starting)

    class NativeSession:
        _read_only = True
        _quarantined = False
        _closed = False
        def __init__(self,kind):
            self.kind = kind
            self.staging_root = tmp_path/(kind+'-runtime')
            self.staging_root.mkdir()
            self.row = [kind,'/private/'+kind+'.docx',True,'target-hash']
            self._bound_path = self.row[1]
        def __enter__(self):
            events.append(self.kind+'-enter')
            native_documents.append(self.row)
            return self
        def __exit__(self,*args):
            events.append(self.kind+'-exit')
            self._closed = True  # deliberately insufficient native close proof
            if failure == 'close-failed' and self.kind == 'owned':
                raise RuntimeError('owned close failed')
            if failure != 'owned-remains' or self.kind != 'owned':
                native_documents.remove(self.row)
            if failure == 'sentinel-changed' and self.kind == 'owned':
                native_documents[native_documents.index(sentinel_row)] = sentinel_row[:3]+['changed-hash']
        def _execute(self,lines):
            commands = '\n'.join(lines)
            if 'make new document' in commands:
                native_documents.append(sentinel_row)
                return [[sentinel_row[0]]]
            if 'close sentinelDoc saving no' in commands:
                events.append('unsafe-bound-sentinel-close')
                native_documents.remove(sentinel_row)
                return [['ok']]
            if 'repeat with di' in commands:
                return sorted(r[:3]+[r[3]+'  -'] for r in native_documents if r != self.row)
            return [['ok']]
        def add_heading_level(self,*args): pass
        def add_paragraph(self,*args): pass
        def add_page_break(self): pass
        def pagination_fragment_for_bookmark(self,*args):
            raise NativeWriterObjectError('PAGINATION_SNAPSHOT_FAILED','pagination snapshot failed')
        def save_docx(self,path): path.write_bytes(b'docx')
        def export_pdf(self,path): path.write_bytes(b'pdf')

    def inventory(output,label):
        events.append('inventory:'+label)
        return sorted(native_documents)
    def close_sentinel(output,name,token):
        events.append('independent-sentinel-close')
        assert name == sentinel_row[0] and token.startswith('Pagination sentinel ')
        native_documents.remove(sentinel_row)
        return [['sentinel-closed',name]]
    monkeypatch.setattr(fixture,'inventory',inventory,raising=False)
    monkeypatch.setattr(fixture,'close_sentinel',close_sentinel,raising=False)
    monkeypatch.setattr(fixture,'create_document',lambda *args,**kwargs:NativeSession('owned'))
    monkeypatch.setattr(fixture,'open_document',lambda *args,**kwargs:NativeSession('reopen'))
    monkeypatch.setattr(fixture,'_state',lambda session:[['state','body-hash',100]])
    monkeypatch.setattr(fixture,'_snapshots',lambda *args:{'native':'snapshot'})
    monkeypatch.setattr(fixture,'_validate',lambda *args:None)
    class PDF:
        pages = [SimpleNamespace(extract_text=lambda t=t:t) for t in
                 ('First page body','Second page body','Third page body')]
        def __enter__(self): return self
        def __exit__(self,*args): pass
    monkeypatch.setattr(pdfplumber,'open',lambda path:PDF())
    result = fixture.run(tmp_path/'output')
    assert 'unsafe-bound-sentinel-close' not in events
    if failure:
        assert result['passed'] is False
        assert 'independent-sentinel-close' not in events
        assert result['retained_sentinel_name'] == sentinel_row[0]
    else:
        assert result['passed'] is True
        assert events.index('owned-exit') < events.index('inventory:after-owned-close')
        assert events.index('reopen-exit') < events.index('inventory:after-reopen-close')
        assert events.index('inventory:after-reopen-close') < events.index('independent-sentinel-close')
        assert events[-1] == 'inventory:after-sentinel-close'
        assert native_documents == starting
        assert 'fixtures/microsoft_parity/macos_word_recovery.py' in result['source_hashes']


@pytest.mark.parametrize('method,args,rows,expected', [
    ('pagination_fragment_for_bookmark',('public😀','B'),[['bookmark',0,2,1,72,90]],
     {'nodeId':'public😀','story':'main','sections':['body'],'pageStart':1,'pageEnd':1,
      'range':'0:2','fragments':[{'page':1,'bounds':[72.0,90.0,73.0,102.0]}]}),
    ('pagination_map_for_ranges',([],),[setup()],{'version':'M5-v1','nodes':[]}),
])
def test_public_session_pagination_forwards_preserve_read_only_result(method,args,rows,expected):
    session = session_with(rows)
    operation = getattr(session,method,None)
    assert callable(operation), 'Missing public pagination forward'
    assert operation(*args) == expected
