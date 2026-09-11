"""Frozen Writer differential contracts; native ports are doubles, not native proof."""
from __future__ import annotations

import ast
from copy import deepcopy
import importlib
import inspect
import re
import subprocess
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_fields import NativeIndexHandle
from skills.WPSComposer.scripts.msoffice.errors import NativeWordError, NativeWordTimeoutError

BOOKMARK = 'wpsc_eq_' + 'a' * 24
GLOBAL = dict(sequenceId='WPSC_EQ', mode='global', prefix='(', suffix=')')
CHAPTER = dict(GLOBAL, mode='chapter', chapterStyleLevel=1, resetLevel=1)


def implementation():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_numbering')
    except ModuleNotFoundError:
        pytest.fail('Bound native numbering implementation is missing')


@pytest.fixture(scope='module')
def frozen_writer():
    source = subprocess.check_output(['git', 'show', '6dd3a00:skills/WPSComposer/scripts/writer.py'], text=True)
    tree = ast.parse(source)
    methods = {'_native_fields', '_native_position', '_native_insert_field',
               '_localized_styleref_code', '_add_native_number_shell', 'add_equation_number_native', 'snapshot_fields'}
    cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'WriterComposer')
    cls.bases = []
    cls.body = [n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name in methods]
    nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_caption_field_codes']
    namespace = dict(__name__='skills.WPSComposer.scripts._frozen_writer', __package__='skills.WPSComposer.scripts', re=re, _NATIVE_SEQUENCE_IDS={'WPSC_FIG', 'WPSC_TAB', 'WPSC_EQ'},
                     _NATIVE_BOOKMARK_RE=re.compile(r'^wpsc_(?:fig|tab|eq)_[0-9a-f]{24}$'))
    exec(compile(ast.Module(body=nodes+[cls], type_ignores=[]), '<frozen-6dd3a00>', 'exec'), namespace)
    return namespace['WriterComposer']


class NativeModel:
    """COM/AppleEvent boundary model. Logical orchestration remains real code."""
    def __init__(self, start=4, end=4, style='标题 1', failure=None):
        self.start, self.position, self.style, self.failure = start, end, style, failure
        self.events = []
        self.fields = []

    def text(self, text):
        if not isinstance(text, str):
            raise TypeError('native text must be text')
        self.events.append(('text', self.start, self.position, text))
        self.position = self.start + len(text.encode('utf-16-le')) // 2
        self.start = self.position

    def field(self, code, owner, kind):
        self.events.append(('field', self.position, code))
        if self.failure == 'field':
            raise RuntimeError('native field failed')
        self.position += 3  # Simulated result + field end, no rendered-number assumption.
        self.start = self.position
        self.fields.append((owner or 'doc:native', kind, 'numbering'))

    def localized_style(self):
        self.events.append(('style',))
        if isinstance(self.style, Exception):
            return 'Heading 1'
        return self.style

    def bookmark(self, name, start, end):
        self.events.append(('bookmark', name, start, end))
        if self.failure == 'bookmark':
            raise RuntimeError('native bookmark failed')

    def format_paragraph(self, start):
        self.events.append(('format', start, self.position, 2, -1))
        if self.failure == 'format':
            raise RuntimeError('native paragraph failed')

    def paragraph(self):
        self.events.append(('paragraph', self.position))
        self.position += 1
        self.start = self.position


def oracle(cls, model):
    writer = cls()
    class Selection:
        @property
        def Range(self):
            return SimpleNamespace(Start=model.start, End=model.position, Collapse=lambda direction: None)
        def TypeText(self, text): model.text(text)
        def TypeParagraph(self): model.paragraph()
        def SetRange(self, start, end): model.start, model.position = start, end
        def MoveRight(self, unit, count): model.start += 1; model.position += 1
    class Fields:
        def Add(self, rng, field_type, code, preserve):
            # Tracking happens in real frozen _native_insert_field, not here.
            model.events.append(('field', rng.End, code))
            if model.failure == 'field': raise RuntimeError('native field failed')
            return SimpleNamespace(Result=SimpleNamespace(End=rng.End+2))
    class Format:
        def __init__(self, start, end): self.start, self.end = start, end
        def __setattr__(self, name, value):
            object.__setattr__(self, name, value)
            if name == 'KeepTogether':
                model.events.append(('format', self.start, self.end, self.Alignment, value))
                if model.failure == 'format': raise RuntimeError('native paragraph failed')
    def style(index):
        model.events.append(('style',))
        if isinstance(model.style, Exception): raise model.style
        return SimpleNamespace(NameLocal=model.style)
    writer.selection = Selection()
    writer._doc = SimpleNamespace(Fields=Fields(), Styles=style,
        Range=lambda start,end: SimpleNamespace(Start=start, End=end, ParagraphFormat=Format(start,end)),
        Bookmarks=SimpleNamespace(Add=lambda name,rng: model.bookmark(name,rng.Start,rng.End)))
    return writer


def outcome(call):
    try: return ('return', call())
    except Exception as error: return ('error', type(error), str(error))


@pytest.mark.parametrize('numbering,source,fallback,bookmark,style,start,end,failure', [
    (GLOBAL, 'x=1', 'unused', BOOKMARK, '标题 1', 4, 4, None),
    (CHAPTER, '中文😀', 'unused', BOOKMARK, '标题 1', 4, 4, None),
    (CHAPTER, 0, 7, None, 'Heading 1', 4, 9, None),
    (CHAPTER, False, None, BOOKMARK, '', 4, 4, None),
    (CHAPTER, '', 'fallback', BOOKMARK, 'bad"name', 4, 4, None),
    (CHAPTER, '', 'fallback', BOOKMARK, RuntimeError('style unavailable'), 4, 4, None),
    (dict(GLOBAL,prefix='😀(',suffix=')尾'), [], False, BOOKMARK, '标题 1', 4, 4, None),
    (dict(GLOBAL,sequenceId='WPSC_TAB'), 'x', '', BOOKMARK, '标题 1', 0, 0, None),
    (dict(CHAPTER,sequenceId='WPSC_FIG'), 'x', '', BOOKMARK, '标题 1', 0, 0, None),
    ({}, 'already inserted', '', 'bad', '标题 1', 4, 4, None),
    (dict(GLOBAL,resetLevel=1), 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    (dict(CHAPTER,chapterStyleLevel=2), 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    ({k:v for k,v in GLOBAL.items() if k!='suffix'}, 'x', '', BOOKMARK, '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', 'bad', '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', '', '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', 42, '标题 1', 4, 4, None),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'field'),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'bookmark'),
    (GLOBAL, 'x', '', BOOKMARK, '标题 1', 4, 4, 'format'),
])
def test_equation_preserves_frozen_order_and_partial_failure(monkeypatch,frozen_writer,
        numbering,source,fallback,bookmark,style,start,end,failure):
    expected = NativeModel(start,end,style,failure)
    baseline = oracle(frozen_writer,expected)
    kwargs = dict(source=source,numbering=numbering,bookmarkName=bookmark,
                  fallbackText=fallback,owner_node_id='node:equation')
    want = outcome(lambda:baseline.add_equation_number_native(**kwargs))
    actual = NativeModel(start,end,style,failure)
    m = implementation()
    monkeypatch.setattr(m,'BoundNumberingCursor',lambda session:actual)
    got = outcome(lambda:MacWordSession().add_equation_number_native(**kwargs))
    assert got == want
    assert actual.events == expected.events
    assert actual.fields == [(owner,kind,category) for owner,kind,field,category in baseline._native_fields()]


def test_public_signature_matches_frozen(frozen_writer):
    assert hasattr(MacWordSession,'add_equation_number_native')
    assert inspect.signature(MacWordSession.add_equation_number_native) == inspect.signature(frozen_writer.add_equation_number_native)


def test_numbering_fields_have_owner_category_and_per_kind_ordinal(monkeypatch):
    s = MacWordSession()
    codes = [' STYLEREF "标题 1" \\s ', ' SEQ WPSC_EQ \\* ARABIC \\s 1 ', ' SEQ WPSC_EQ \\* ARABIC ']
    kinds = ['STYLEREF','SEQ_EQ','SEQ_EQ']
    s._tracked_numbering = [(NativeIndexHandle('s',f'own{i}','eq:one',kind,'numbering'),code)
                           for i,(kind,code) in enumerate(zip(kinds,codes))]
    rows = [['stats',1]]
    for i,(kind,code) in enumerate(zip(kinds,codes)):
        rows += [['identity',f'own{i}',i*50+1], ['field','story:main text/chain:1',i+1,kind,code,i*50+1,str(i+1),0,0,1]]
    monkeypatch.setattr(s,'_execute',lambda lines:deepcopy(rows))
    got = s.snapshot_fields()
    assert [f.stable_key for f in got] == [('eq:one','STYLEREF',0),('eq:one','SEQ_EQ',0),('eq:one','SEQ_EQ',1)]
    assert all(f.field_category=='numbering' for f in got)
    rows[2][4] = ' SEQ WPSC_TAB \\* ARABIC '
    with pytest.raises(NativeWordError): s.snapshot_fields()


def test_checkpoint_restores_numbering_tracking_after_confirmed_rollback(monkeypatch):
    from tests.msoffice.test_macos_word_recovery import rows,bind
    a = rows()
    field = ['field','main',1,'SEQ_EQ',' SEQ WPSC_EQ \\* ARABIC ',8,33,34,35]
    s,t = bind(monkeypatch,[a,rows(end=40,fields=[field]),[['rollback-ack',[field],[]],*a]])
    token = s.degradation_checkpoint()
    s._tracked_numbering = [(NativeIndexHandle('s','new','n','SEQ_EQ','numbering'),field[4])]
    assert s.rollback_degradation_checkpoint(token) is None
    assert s._tracked_numbering == []


def test_numbering_tracking_prefix_cannot_be_rewritten_before_rollback(monkeypatch):
    from tests.msoffice.test_macos_word_recovery import rows,bind
    handle = NativeIndexHandle('s','old','n','SEQ_EQ','numbering')
    a = rows(bookmarks=[['bookmark','old',1,4,'abc']])
    s,t = bind(monkeypatch,[a])
    s._tracked_numbering = [(handle,'abc')]
    token = s.degradation_checkpoint()
    s._tracked_numbering[0] = (handle,'rewritten')
    with pytest.raises(Exception) as error: s.rollback_degradation_checkpoint(token)
    assert error.value.code == 'LOCAL_MUTATION_ROLLBACK_FAILED'
    assert len(t.calls) == 1


def bind_cursor(monkeypatch, replies):
    m = implementation()
    s = MacWordSession()
    s._bound_path = '/private/owned.docx'
    calls = []
    def execute(lines):
        calls.append(lines)
        row = replies.pop(0)
        if isinstance(row,BaseException):
            # This double represents submitted native errors. Local script-I/O
            # failures are tested separately through the real transport below.
            s._invalidate_field_topology()
            s._field_topology_mutation_pending=False
            raise row
        return deepcopy(row)
    monkeypatch.setattr(s,'_execute',execute)
    monkeypatch.setattr(m,'uuid4',lambda:SimpleNamespace(hex='a'*32))
    return s,calls,m


def field_row():
    return ['field',2,3,10,35,36,37,' SEQ WPSC_EQ \\* ARABIC ',
            'WPSC_N_'+'a'*30,10,35,' SEQ WPSC_EQ \\* ARABIC ']


def test_real_cursor_consumes_exact_noncollapsed_selection_and_native_field_end(monkeypatch):
    replies = [[['selection',4,9]],[['text',4,7,'x😀']],[['text',7,8,'\t']],
               [['text',8,9,'(']],[field_row()],
               [['bookmark',BOOKMARK,9,38]],[['text',38,39,')']],
               [['format',True,True]],[['text',39,40,'\r']]]
    s,calls,m = bind_cursor(monkeypatch,replies)
    s._observed_field_topology = (('before',),)
    assert s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=BOOKMARK,
                                        fallbackText='unused') == {'issues':[]}
    assert not replies
    assert s._structural_changed and not hasattr(s,'_observed_field_topology')
    handle,code = s._tracked_numbering[0]
    assert (handle.owner_node_id,handle.kind,handle.category)==('doc:native','SEQ_EQ','numbering')
    assert code == ' SEQ WPSC_EQ \\* ARABIC '


@pytest.mark.parametrize('change', ['count','position','bounds','code','identity','identity-bounds','identity-text'])
def test_malformed_native_field_ack_never_commits_tracking(monkeypatch,change):
    row = field_row()
    if change=='count': row[2]=4
    elif change=='position': row[3]=11
    elif change=='bounds': row[5]=34
    elif change=='code': row[7]=row[11]=' SEQ ATTACK \\* ARABIC '
    elif change=='identity': row[8]='foreign'
    elif change=='identity-bounds': row[9]=11
    else: row[11]='different'
    s,calls,m = bind_cursor(monkeypatch,[[['selection',9,9]],[row]])
    c = m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.field('SEQ WPSC_EQ \\* ARABIC','owner','SEQ_EQ')
    assert s._quarantined and s._retain_evidence and not hasattr(s,'_tracked_numbering')


@pytest.mark.parametrize('bad', [[],[['text',4,7,'wrong']],[['text',4,8,'x😀']]])
def test_literal_bad_ack_quarantines_without_continuing(monkeypatch,bad):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',4,9]],bad])
    with pytest.raises(NativeWordError):
        s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=None,fallbackText='')
    assert s._quarantined and len(calls)==2


@pytest.mark.parametrize('error',[NativeWordTimeoutError(),NativeWordError('NATIVE_WORD_EXECUTION_FAILED'),KeyboardInterrupt()])
def test_uncertain_mutation_stops_without_retry(monkeypatch,error):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',4,4]],error])
    with pytest.raises(type(error)):
        s.add_equation_number_native(source='x',numbering=GLOBAL,bookmarkName=None,fallbackText='')
    assert s._quarantined and s._retain_evidence and len(calls)==2


@pytest.mark.parametrize('state',['_closed','_read_only','_quarantined'])
def test_unavailable_session_cannot_submit_any_native_call(monkeypatch,state):
    s,calls,m = bind_cursor(monkeypatch,[])
    setattr(s,state,True)
    with pytest.raises((ValueError,NativeWordError)):
        s.add_equation_number_native(source='x',numbering={},bookmarkName=None,fallbackText='')
    assert not calls


def test_late_bookmark_validation_preserves_confirmed_native_field_tracking(monkeypatch):
    s,calls,m = bind_cursor(monkeypatch,[[['selection',6,6]],[['text',6,7,'x']],
        [['text',7,8,'\t']],[['text',8,9,'(']],[field_row()]])
    with pytest.raises(ValueError,match='invalid native bookmark'):
        s.add_equation_number_native(source='x',numbering=GLOBAL,bookmarkName='bad',fallbackText='')
    assert len(s._tracked_numbering)==1 and not s._quarantined
    assert len(calls)==5


def test_numbering_native_scripts_compile_without_execution(monkeypatch,tmp_path):
    import sys
    if sys.platform!='darwin': pytest.skip('Word dictionary compile only')
    replies = [[['selection',4,9]],[['text',4,7,'x😀']],[['text',7,8,'\t']],
               [['text',8,9,'(']],[field_row()],
               [['bookmark',BOOKMARK,9,38]],[['text',38,39,')']],
               [['format',True,True]],[['text',39,40,'\r']]]
    s,calls,m = bind_cursor(monkeypatch,replies)
    s.add_equation_number_native(source='x😀',numbering=GLOBAL,bookmarkName=BOOKMARK,fallbackText='')
    c = m.BoundNumberingCursor.__new__(m.BoundNumberingCursor)
    c.session=s; c.start=c.position=40
    replies.append([['style','标题 1']]); c.localized_style()
    for index,lines in enumerate(calls):
        path = tmp_path/f'{index}.applescript'
        path.write_text('use framework "Foundation"\nuse scripting additions\ntell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell\n')
        result = subprocess.run(['/usr/bin/osacompile','-o',str(path.with_suffix('.scpt')),str(path)],capture_output=True,text=True,timeout=20)
        assert result.returncode==0,result.stderr


def test_native_fixture_requires_explicit_execute_before_creating_directory(tmp_path):
    import sys
    result=subprocess.run([sys.executable,'fixtures/microsoft_parity/macos_word_numbering.py',
        '--output',str(tmp_path/'untouched')],capture_output=True,text=True)
    assert result.returncode==2 and '--execute' in result.stderr
    assert not (tmp_path/'untouched').exists()


def test_native_fixture_validates_number_only_bookmarks_and_real_fields():
    try: m=importlib.import_module('fixtures.microsoft_parity.macos_word_numbering')
    except ModuleNotFoundError: pytest.fail('Guarded native numbering fixture is missing')
    rows=[['body','Chapter\rleft x=1\t😀(1)尾\rfallback\t[1-1]\r7\t(2)\r right\rTAIL\r'],
          ['field','SEQ WPSC_EQ \\* ARABIC','1'],['field','STYLEREF "标题 1" \\s','1'],
          ['field','SEQ WPSC_FIG \\* ARABIC \\s 1','1'],['field','SEQ WPSC_EQ \\* ARABIC','2'],
          ['bookmark',m.BOOKMARKS[0],20,50,'1'],['bookmark',m.BOOKMARKS[1],60,100,'1-1'],
          ['bookmark',m.BOOKMARKS[2],110,140,'2'],
          ['counts',0,0,0],['paragraph','left x=1\t😀(1)尾\r',True,True]]
    assert m.validate_readback(rows)
    broken=deepcopy(rows);broken[5][-1]='😀(1)尾'
    assert not m.validate_readback(broken)
    broken=deepcopy(rows);broken[2][1]='REF fake'
    assert not m.validate_readback(broken)


from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_fields import NativeIndexHandle
from skills.WPSComposer.scripts.longform.field_contract import snapshot_visible_field


def test_main_numbering_does_not_collide_with_header(monkeypatch):
    s = MacWordSession()
    code = ' STYLEREF "Heading 1" \\s '
    s._tracked_numbering = [(NativeIndexHandle('s','own','eq:one','STYLEREF','numbering'),code)]
    monkeypatch.setattr(s, '_execute', lambda lines: [
        ['stats',1], ['identity','own',10],
        ['field','story:main text/chain:1',1,'STYLEREF',code,10,'1',0,0,1],
        ['field','story:primary header/chain:1',1,'STYLEREF',code,10,'1',0,0,1],
    ])
    result = s.snapshot_fields()
    assert [r.stable_key for r in result] == [
        ('eq:one','STYLEREF',0), ('story:primary header/chain:1','STYLEREF',0),
    ]


def test_numbering_snapshot_preserves_frozen_creation_order_ordinals(monkeypatch):
    s = MacWordSession()
    code = ' SEQ WPSC_EQ \\* ARABIC '
    s._tracked_numbering = [
        (NativeIndexHandle('s','first','eq:one','SEQ_EQ','numbering'),code),
        (NativeIndexHandle('s','second','eq:one','SEQ_EQ','numbering'),code),
    ]
    monkeypatch.setattr(s, '_execute', lambda lines: [
        ['stats',1], ['identity','first',100], ['identity','second',10],
        ['field','story:main text/chain:1',1,'SEQ_EQ',code,10,'SECOND',0,0,6],
        ['field','story:main text/chain:1',2,'SEQ_EQ',code,100,'FIRST',0,0,5],
    ])
    by_key = {r.stable_key:r.result_hash for r in s.snapshot_fields()}
    for ordinal,visible in enumerate(('FIRST','SECOND')):
        expected = snapshot_visible_field(owner_node_id='eq:one',field_kind='SEQ_EQ',
            ordinal_within_node=ordinal,visible_result=visible,field_category='numbering',
            toc_page_count=0,figure_index_page_count=0,table_index_page_count=0,total_pages=1)
        assert by_key[expected.stable_key] == expected.result_hash
    # Topology must retain native document order independently of semantic ordinals.
    assert [r[4] for r in s._observed_field_topology] == [10,100]


from skills.WPSComposer.scripts.msoffice.errors import NativeWordError
import pytest

@pytest.mark.parametrize("rows", [[None],[3],[True]])
def test_malformed_selection_ack_quarantines(monkeypatch,rows):
    s,calls,m=bind_cursor(monkeypatch,[rows])
    with pytest.raises(NativeWordError): m.BoundNumberingCursor(s)
    assert s._quarantined

def test_bool_coordinate_is_not_text_ack(monkeypatch):
    s,calls,m=bind_cursor(monkeypatch,[[["selection",0,0]],[["text",0,True,"x"]]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.text("x")
    assert s._quarantined

@pytest.mark.parametrize("kind", ["field","style"])
def test_malformed_postmutation_ack_quarantines(monkeypatch,kind):
    s,calls,m=bind_cursor(monkeypatch,[[["selection",0,0]],[None]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError):
        if kind=="field": c.field("SEQ WPSC_EQ \\* ARABIC","n","SEQ_EQ")
        else: c.localized_style()
    assert s._quarantined


@pytest.mark.parametrize('owner',['doc:native','eq:one'])
@pytest.mark.parametrize('kind,code',[('SEQ_EQ',' SEQ WPSC_EQ \\* ARABIC '),('STYLEREF',' STYLEREF "Heading 1" \\s ')])
def test_reverse_position_snapshot_matches_frozen_writer(monkeypatch,frozen_writer,owner,kind,code):
    baseline=frozen_writer()
    baseline._doc=SimpleNamespace(ComputeStatistics=lambda kind:1)
    baseline._iter_section_page_fields=lambda:[]
    baseline._wpsc_native_fields=[(owner,kind,SimpleNamespace(Result=SimpleNamespace(Text=value)),'numbering') for value in ('FIRST','SECOND')]
    s=MacWordSession()
    s._tracked_numbering=[(NativeIndexHandle('s',name,owner,kind,'numbering'),code) for name in ('first','second')]
    monkeypatch.setattr(s,'_execute',lambda lines:[['stats',1],['identity','first',100],['identity','second',10],
        ['field','story:main text/chain:1',1,kind,code,10,'SECOND',0,0,6],
        ['field','story:main text/chain:1',2,kind,code,100,'FIRST',0,0,5]])
    assert {f.stable_key:f for f in s.snapshot_fields()}=={f.stable_key:f for f in baseline.snapshot_fields()}
    assert [r[4] for r in s._observed_field_topology]==[10,100]


def test_main_ref_identity_ignores_same_offset_header_ref(monkeypatch):
    s=MacWordSession();code=' REF target \\h '
    s._tracked_references=[(NativeIndexHandle('s','own','ref:one','REF','reference'),code)]
    monkeypatch.setattr(s,'_execute',lambda lines:[['stats',1],['identity','own',10],
        ['field','story:main text/chain:1',1,'REF',code,10,'main',0,0,4],
        ['field','story:primary header/chain:1',1,'REF',code,10,'header',0,0,6]])
    assert [f.stable_key for f in s.snapshot_fields()]==[('ref:one','REF',0),('story:primary header/chain:1','REF',0)]


@pytest.mark.parametrize('bad',[None,3,True,{},'field'])
def test_field_malformed_nested_rows_quarantine_after_native_mutation(monkeypatch,bad):
    s,calls,m=bind_cursor(monkeypatch,[[['selection',0,0]],[bad]])
    c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError): c.field('SEQ WPSC_EQ \\* ARABIC','n','SEQ_EQ')
    assert s._quarantined and s._structural_changed and not hasattr(s,'_tracked_numbering')


@pytest.mark.parametrize('action', ['bookmark','format'])
def test_numbering_ack_booleans_and_integers_are_not_interchangeable(monkeypatch,action):
    ack=[['bookmark','b',False,True]] if action=='bookmark' else [['format',1,1]]
    s,calls,m=bind_cursor(monkeypatch,[[['selection',0,0]],ack]);c=m.BoundNumberingCursor(s)
    with pytest.raises(NativeWordError):
        if action=='bookmark': c.bookmark('b',0,1)
        else: c.format_paragraph(0)
    assert s._quarantined


@pytest.mark.parametrize('failure',['script-write','deadline','read-only','submitted-error'])
def test_real_transport_distinguishes_presubmission_failure(monkeypatch,tmp_path,failure):
    from pathlib import Path
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=MacWordSession();s.staging_root=tmp_path;s._bound_path='/private/owned.docx';s._owns_doc=True
    reasons=[];calls=[]
    s.lock=SimpleNamespace(quarantine_path=tmp_path/'quarantine.json',quarantine=lambda data:reasons.append(data['reason']))
    s._observed_field_topology=(('before',1),);s._pending_heading=(0,4,1)
    c=implementation().BoundNumberingCursor.__new__(implementation().BoundNumberingCursor)
    c.session=s;c.start=c.position=0
    original_write=Path.write_text
    def write(path,data,**kwargs):
        if 'set content of numberRange' in data:
            if failure=='script-write': raise OSError('local disk full')
            if failure=='deadline': s._deadline=-1
        return original_write(path,data,**kwargs)
    monkeypatch.setattr(Path,'write_text',write)
    def run(command,**kwargs):
        calls.append(command)
        return subprocess.CompletedProcess(command,1,'','execution error: NUMBERING_PARTIAL_FAILURE (-2700)')
    monkeypatch.setattr(transport.subprocess,'run',run)
    if failure=='read-only': s._read_only=True
    with pytest.raises((OSError,NativeWordError,ValueError)): c.text('x')
    if failure=='submitted-error':
        assert len(calls)==1 and s._quarantined and reasons==['Native numbering completion uncertain']
        assert not hasattr(s,'_observed_field_topology') and s._pending_heading is None
    else:
        assert not calls and s._observed_field_topology==( ('before',1), )
        assert s._pending_heading==(0,4,1)
        if failure=='deadline':
            assert s._quarantined and reasons==['Session deadline expired before verified document close']
        else: assert not s._quarantined and not reasons
    assert not getattr(s,'_field_topology_mutation_pending',False)
