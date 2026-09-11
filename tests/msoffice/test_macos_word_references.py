"""Direct reference contracts. Transport doubles and compilation are not native proof."""
import inspect
import subprocess
import sys

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, NativeWordError
from skills.WPSComposer.scripts.writer import WriterComposer, NativeWriterObjectError

METHODS = ('add_bibliography_native', 'add_bibliography_legacy', 'add_cross_reference_paragraph', 'add_citation_paragraph', 'add_cross_reference_fallback')
BOOKMARK = 'wpsc_fig_' + 'a' * 24
REF = dict(type='reference', bookmarkName=BOOKMARK, prefix='见😀(', suffix=')尾', fallbackText='静态')

@pytest.fixture
def session(monkeypatch):
    s = MacWordSession()
    calls = []
    def execute(lines):
        calls.append(lines)
        assert s._structural_changed and s._pending_heading is None
        return []
    monkeypatch.setattr(s, '_execute', execute)
    return s, calls

@pytest.mark.parametrize('name', METHODS)
def test_direct_signatures_are_compatible(name):
    assert hasattr(MacWordSession, name)
    assert inspect.signature(getattr(MacWordSession, name)) == inspect.signature(getattr(WriterComposer, name))

@pytest.mark.parametrize('method,kwargs', [
    ('add_bibliography_native', dict(entries=[{'number':1}], style='numeric', hangingIndentPt=18,leftIndentPt=18,spaceAfterPt=6)),
    ('add_bibliography_native', dict(entries=[], style='numeric', hangingIndentPt=float('nan'),leftIndentPt=18,spaceAfterPt=6)),
    ('add_bibliography_native', dict(entries=[], style='other', hangingIndentPt=18,leftIndentPt=18,spaceAfterPt=6)),
    ('add_bibliography_legacy', dict(entries='not a sequence of entries')),
    ('add_cross_reference_paragraph', dict(runs=[dict(REF, bookmarkName='bad')])),
    ('add_citation_paragraph', dict(runs=[{'type':'citation'}])),
    ('add_citation_paragraph', dict(runs=[{'type':'degradation','fallbackText':'x','code':'X'}])),
    ('add_cross_reference_fallback', dict(runs=[REF], listFormatting={'indentPt':float('inf')})),
    ('add_cross_reference_paragraph', dict(runs=[REF] * 10001)),
])
def test_whole_call_validation_precedes_mutation(session, method, kwargs):
    s,calls=session
    before=s.__dict__.copy()
    with pytest.raises(ValueError):getattr(s,method)(**kwargs)
    assert not calls and s.__dict__ == before

@pytest.mark.parametrize('name', METHODS)
@pytest.mark.parametrize('state', ['_read_only','_closed'])
def test_closed_readonly_never_submit(session,name,state):
    s,calls=session;setattr(s,state,True)
    kwargs = dict(entries=[]) if 'bibliography' in name else dict(runs=[])
    if name.endswith('_native'):kwargs.update(style='numeric',hangingIndentPt=18,leftIndentPt=18,spaceAfterPt=6)
    with pytest.raises(ValueError):getattr(s,name)(**kwargs)
    assert not calls


def test_native_bibliography_direct_order_geometry_unicode(session,monkeypatch):
    s,calls=session
    def execute(lines):
        calls.append(lines)
        return [['paragraph',0,0,9,'[9] 中文😀\r',0,21.5,-12.25,0,4.5,True],
                ['paragraph',1,9,15,'[2] 7\r',0,21.5,-12.25,0,4.5,True],['complete',1,3,0,15]]
    monkeypatch.setattr(s,'_execute',execute)
    result=s.add_bibliography_native(entries=[{'number':'9','text':'中文😀','extra':object()},{'number':2,'text':7}],style='numeric',hangingIndentPt=12.25,leftIndentPt=21.5,spaceAfterPt=4.5,owner_node_id=object(),controller_owned=object())
    assert result=={'issues':[]}
    code='\n'.join(calls[0]);assert 'style body text' not in code
    assert 'paragraph format left indent' in code and 'keep together' in code


def test_legacy_literals_and_precomputed_scalar_coercion(session,monkeypatch):
    s,calls=session
    rows=[['paragraph',0,0,9,'[7] 中文😀\r'],['paragraph',1,9,11,'7\r'],['paragraph',2,11,17,'False\r'],['paragraph',3,17,22,'None\r'],['complete',1,5,0,22]]
    monkeypatch.setattr(s,'_execute',lambda lines: calls.append(lines) or rows)
    assert s.add_bibliography_legacy(entries=['[7] 中文😀',7,False,None],style=object())=={'issues':[]}
    assert '[1] ' not in '\n'.join(calls[0])


def test_late_conversion_failure_is_pure(session):
    class Bad:
        def __str__(self):raise RuntimeError('private conversion')
    s,calls=session
    with pytest.raises(ValueError):s.add_bibliography_legacy(entries=['first',Bad()])
    assert not calls and not s._structural_changed


def test_empty_bibliography_is_noop(session):
    s,calls=session
    assert s.add_bibliography_native(entries=[],style='numeric',hangingIndentPt=9,leftIndentPt=22,spaceAfterPt=2)=={'issues':[]}
    assert s.add_bibliography_legacy(entries=[])=={'issues':[]}
    assert not calls


def test_mixed_direct_normalization_ignores_plan_metadata():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import normalize_runs
    runs,indent=normalize_runs([{'type':'text','text':'中文😀'},dict(REF,targetKind='wrong',targetNodeId=object()),{'type':'citation','fallbackText':'[9; 3]','extra':object()},{'type':'degradation','fallbackText':'[CUSTOM 待核]','code':'CUSTOM','nodeId':'node:1'}],{'indentPt':'31.5','kind':'inert'})
    assert runs[0]['text']=='中文😀' and runs[1]['prefix']=='见😀('
    assert 'targetKind' not in runs[1] and runs[2]['fallbackText']=='[9; 3]'
    assert indent==31.5


def test_bad_ack_quarantines_and_invalidates(session):
    s,calls=session
    with pytest.raises(NativeWordError):s.add_citation_paragraph(runs=[{'type':'citation','fallbackText':'[1]'}])
    assert calls and s._quarantined and s._structural_changed


def test_controller_failure_is_controlled_without_local_fallback(session,monkeypatch):
    s,calls=session
    monkeypatch.setattr(s,'_execute',lambda lines:calls.append(lines) or [['literal',0,0,4,'见😀('],['reference_failure',0,4,4]])
    with pytest.raises(NativeWriterObjectError) as e:s.add_cross_reference_paragraph(runs=[REF],controller_owned=True)
    assert e.value.code=='CROSS_REFERENCE_FAILED' and not s._quarantined
    assert 'set content of rollbackRange' not in '\n'.join(calls[0])


def test_pure_commands_cover_native_and_static_modes():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import normalize_runs, paragraph_commands
    runs,indent=normalize_runs([{'type':'text','text':'•\t中文😀'},REF,{'type':'citation','fallbackText':'[8]'},{'type':'degradation','fallbackText':'[MISS 缺失]','code':'MISS','nodeId':'one'}],{'indentPt':24})
    lines,handles=paragraph_commands(MacWordSession(),runs,indent,False,False)
    code='\n'.join(lines)
    assert 'field type field ref' in code and 'preserve formatting true' in code
    assert 'start referenceStart end referenceStart' in code
    assert 'referenceResultEnd + 1' in code
    assert 'set content of rollbackRange to ""' in code
    assert 'List Paragraph' in code and 'line space1 pt5' in code and 'tab stop position:24.0' in code
    assert 'style of trailingRange to style normal' in code
    assert 'set insertionPoint to insertionPoint + 6' in code  # bullet/tab/CJK/emoji
    assert len(handles)==1
    static,_=paragraph_commands(MacWordSession(),runs,indent,False,True)
    assert 'create new field' not in '\n'.join(static)
    assert not any(x in code.lower() for x in ('quit','clipboard','activate','do shell script'))

@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax check only')
def test_all_reference_applescript_branches_compile(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_word_references import normalize_runs, paragraph_commands, bibliography_commands
    s=MacWordSession();runs,indent=normalize_runs([{'type':'text','text':'中文😀'},REF,{'type':'degradation','fallbackText':'缺失','code':'MISS','nodeId':'n'}],{'indentPt':24})
    batches=[paragraph_commands(s,runs,indent,c,f)[0] for c,f in [(False,False),(True,False),(False,True)]]
    batches += [bibliography_commands(s,['[2] 中文😀'],(18.,18.,6.)),bibliography_commands(s,['literal'],None)]
    for i,lines in enumerate(batches):
        source=tmp_path/f'{i}.applescript';source.write_text('tell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell')
        result=subprocess.run(['/usr/bin/osacompile','-o',str(source.with_suffix('.scpt')),str(source)],capture_output=True,text=True)
        assert result.returncode==0,result.stderr


def test_citation_degradation_occurrences_and_utf16_ack(session,monkeypatch):
    s,calls=session
    runs=[{'type':'text','text':'1.\t中文😀'}, {'type':'citation','fallbackText':'[8]'},
          {'type':'degradation','fallbackText':'[MISS 缺失]','code':'MISS','nodeId':'one'},
          {'type':'degradation','fallbackText':'[MISS 缺失]','code':'MISS','nodeId':'two'}]
    rows=[['literal',0,0,7,'1.\t中文😀'],['literal',1,7,10,'[8]'],
          ['literal',2,10,19,'[MISS 缺失]'],['literal',3,19,28,'[MISS 缺失]'],
          ['literal',4,28,29,'\r'],['styled',2,10,19,True,True,True],['styled',3,19,28,True,True,True],['list',24.,-24.,True,0,3,True],['complete',1,2,0,29]]
    monkeypatch.setattr(s,'_execute',lambda lines:calls.append(lines) or rows)
    result=s.add_citation_paragraph(runs=runs,listFormatting={'indentPt':24})
    assert [i['nodeId'] for i in result['issues']]==['one','two']
    assert [i['code'] for i in result['issues']]==['MISS','MISS']


def test_local_reference_fallback_one_issue_and_exact_range(session,monkeypatch):
    s,calls=session
    rows=[['literal',0,0,4,'见😀('],['fallback',0,4,20,3,3],
          ['literal',0,4,6,'静态'],['literal',0,6,8,')尾'],['literal',1,8,9,'\r'],['complete',1,2,0,9]]
    monkeypatch.setattr(s,'_execute',lambda lines:calls.append(lines) or rows)
    result=s.add_cross_reference_paragraph(runs=[REF])
    assert result=={'issues':[{'code':'CROSS_REFERENCE_FAILED','message':'Cross-reference used its inline fallback','placement':'inline'}]}
    assert not hasattr(s,'_tracked_references')


def test_static_fallback_styles_only_fallback_subrange(session,monkeypatch):
    s,calls=session
    rows=[['literal',0,0,4,'见😀('],['literal',0,4,6,'静态'],['literal',0,6,8,')尾'],['literal',1,8,9,'\r'],['styled',0,4,6,True,True,True],['complete',1,2,0,9]]
    monkeypatch.setattr(s,'_execute',lambda lines:calls.append(lines) or rows)
    assert s.add_cross_reference_fallback(runs=[REF],failure_code=object(),owner_node_id=object()) is None


def test_native_field_owner_identity_survives_snapshot(session,monkeypatch):
    s,calls=session
    def execute(lines):
        calls.append(lines)
        identity=next(x for x in lines if 'make new bookmark' in x).split('name:"')[1].split('"')[0]
        return [['literal',0,0,4,'见😀('],['reference',0,2,3,5,46,47,' REF '+BOOKMARK+' \\h ',identity],
                ['literal',0,48,50,')尾'],['literal',1,50,51,'\r'],['complete',1,2,0,51]]
    monkeypatch.setattr(s,'_execute',execute)
    assert s.add_cross_reference_paragraph(runs=[REF],owner_node_id='owner:two')=={'issues':[]}
    handle,code=s._tracked_references[0]
    rows=[['stats',1],['identity',handle.bookmark,5],['field','story:main text/chain:1',3,'REF',code,5,'2',0,0,1]]
    monkeypatch.setattr(s,'_execute',lambda lines:rows)
    snapshot=s.snapshot_fields()
    assert snapshot[0].stable_key==('owner:two','REF',0) and snapshot[0].field_category=='reference'
    rows[1][2]=6
    with pytest.raises(NativeWordError):s.snapshot_fields()


def test_native_readback_failure_is_not_silently_recovered():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import normalize_runs,paragraph_commands
    runs,indent=normalize_runs([REF],None)
    code='\n'.join(paragraph_commands(MacWordSession(),runs,indent,False,False)[0])
    assert 'if referenceError starts with "WPSC_REFERENCE_" then error referenceError number referenceNumber' in code


def test_fallback_cr_counts_as_native_paragraph_delta(session,monkeypatch):
    s,calls=session
    rows=[['literal',0,0,0,''],['fallback',0,0,0,0,0],['literal',0,0,3,'a\rb'],['literal',0,3,3,''],['literal',1,3,4,'\r'],['complete',1,3,0,4]]
    monkeypatch.setattr(s,'_execute',lambda lines:rows)
    assert len(s.add_cross_reference_paragraph(runs=[dict(REF,prefix='',suffix='',fallbackText='a\rb')])['issues'])==1


def test_malformed_controller_prefix_cannot_report_recoverable_failure(session,monkeypatch):
    s,calls=session
    monkeypatch.setattr(s,'_execute',lambda lines:[['literal',0,0,2,'wrong'],['reference_failure',0,4,4]])
    with pytest.raises(NativeWordError):s.add_cross_reference_paragraph(runs=[REF],controller_owned=True)
    assert s._quarantined


def test_transport_uncertainty_retains_no_fallback(session,monkeypatch):
    from skills.WPSComposer.scripts.msoffice.errors import NativeWordTimeoutError
    s,calls=session
    def execute(lines):
        calls.append(lines)
        raise NativeWordTimeoutError()
    monkeypatch.setattr(s,'_execute',execute)
    with pytest.raises(NativeWordTimeoutError):s.add_cross_reference_paragraph(runs=[REF])
    assert len(calls)==1 and s._quarantined and s._retain_evidence


def test_empty_run_paragraph_has_exact_one_cr(session,monkeypatch):
    s,calls=session
    monkeypatch.setattr(s,'_execute',lambda lines:[['literal',0,7,8,'\r'],['complete',3,4,7,8]])
    assert s.add_citation_paragraph(runs=[])=={'issues':[]}


def test_guarded_fixture_no_execute_cannot_touch_word(tmp_path):
    result=subprocess.run([sys.executable,'fixtures/microsoft_parity/macos_word_references.py','--output',str(tmp_path/'never-created')],capture_output=True,text=True)
    assert result.returncode==2 and '--execute required' in result.stderr
    assert not (tmp_path/'never-created').exists()


def test_fixture_native_style_ranges_use_utf16_and_both_list_kinds(monkeypatch):
    from fixtures.microsoft_parity.macos_word_references import inspect_native_styles
    class Session:
        def _execute(self,lines):
            self.lines=lines
            return [['list',24,-24,True,0,3,True],['list',31.5,-31.5,True,0,3,True],
                    ['degradation',26,35,'[MISS 缺失]',True,True,True],
                    ['degradation',35,44,'[MISS 缺失]',True,True,True]]
    s=Session()
    rows=[['paragraph',1,'•\t见8\r',24,-24,0,3,False,0,6],
          ['paragraph',2,'1.\t😀x[MISS 缺失][MISS 缺失]\r',31.5,-31.5,0,3,False,20,45]]
    observed=inspect_native_styles(s,rows)
    assert observed[2][1:3]==[26,35] and observed[3][1:3]==[35,44]
    assert 'start 26 end 35' in '\n'.join(s.lines)


def test_empty_literal_ack_accepts_native_missing_value_without_writing():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import _literal
    code='\n'.join(_literal('',3))
    assert 'set content of literalRange' not in code
    assert 'emptyValue is not missing value' in code
    assert 'start of content of literalRange,end of content of literalRange,""' in code


def test_tab_readback_uses_native_indexed_collection_not_every_iteration():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import _list_commands
    code='\n'.join(_list_commands(24.0))
    assert 'repeat with tabIndex from 1 to count tab stops of paragraph 1 of paragraphRange' in code
    assert 'set ownTab to tab stop tabIndex of paragraph 1 of paragraphRange' in code
    assert 'repeat with ownTab in' not in code


def test_degradation_styling_is_deferred_until_suffix_and_final_cr_exist():
    from skills.WPSComposer.scripts.msoffice.macos_word_references import normalize_runs,paragraph_commands
    runs,indent=normalize_runs([REF],None)
    code='\n'.join(paragraph_commands(MacWordSession(),runs,indent,False,True)[0])
    assert code.index('set italic of font object of styledRange') > code.index('set content of literalRange to ")尾"')
    assert 'set end of styleRanges to {0,literalStart,insertionPoint}' in code
