"""Recovery contracts exercised through real recovery functions with native ACKs."""
from copy import deepcopy
import importlib
import hashlib
import subprocess
import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_fields import NativeIndexHandle
from skills.WPSComposer.scripts.writer import NativeWriterObjectError

MODULE = 'skills.WPSComposer.scripts.msoffice.macos_word_recovery'


def recovery():
    try:
        return importlib.import_module(MODULE)
    except ModuleNotFoundError:
        pytest.fail('Mac Word recovery module is missing')


def rows(end=6, *, prefix='target', tables=(), fields=(), bookmarks=()):
    return [['checkpoint-state', 1, end, end + 1, 1, 0, end + 1,
             hashlib.sha256(prefix.encode()).hexdigest(), 0, 6, prefix,
             hashlib.sha256((prefix+'\r').encode()).hexdigest()],
            *map(list,tables), *map(list,fields), *map(list,bookmarks), ['objects',0,0,len(tables),len(fields),len(bookmarks)], ['checkpoint-end']]


class Transport:
    def __init__(self, session, replies):
        self.session, self.replies, self.calls = session, list(replies), []

    def __call__(self, lines, **kwargs):
        if self.session._quarantined:
            raise RuntimeError('quarantined')
        self.calls.append(lines)
        result = self.replies.pop(0)
        if isinstance(result, BaseException):
            self.session._retain('uncertain native completion')
            raise result
        return deepcopy(result)


def bind(monkeypatch, replies):
    s=MacWordSession(); transport=Transport(s,replies)
    monkeypatch.setattr(s,'_execute',transport)
    return s,transport


def test_plain_integer_checkpoint_and_repeat(monkeypatch):
    r=recovery(); s,t=bind(monkeypatch,[rows(),rows()])
    token=r.checkpoint(s)
    assert type(token) is int and token==6
    assert r.checkpoint(s)==token
    assert len(s._degradation_checkpoints)==1


def test_noop_retry_and_int_coercion(monkeypatch):
    r=recovery(); a=rows(); s,t=bind(monkeypatch,[a,a,[['rollback-ack',[],[]],*a],a,[['rollback-ack',[],[]],*a]])
    r.checkpoint(s)
    assert r.rollback(s,'6') is None
    assert r.rollback(s,6.9) is None
    assert s._structural_changed and s._pending_heading is None


def test_same_coordinate_different_preimage_is_not_overwritten(monkeypatch):
    r=recovery(); s,t=bind(monkeypatch,[rows(),rows(prefix='change')]); r.checkpoint(s)
    snapshot=s._degradation_checkpoints[6]
    with pytest.raises(NativeWriterObjectError,match='checkpoint failed'):r.checkpoint(s)
    assert s._degradation_checkpoints[6] is snapshot


@pytest.mark.parametrize('token',[-1,9,5,'bad'])
def test_invalid_target_no_delete(monkeypatch,token):
    r=recovery(); s,t=bind(monkeypatch,[rows(),rows()]); r.checkpoint(s)
    with pytest.raises(NativeWriterObjectError) as caught:r.rollback(s,token)
    assert caught.value.code=='LOCAL_MUTATION_ROLLBACK_FAILED'
    assert not any('set content of rollbackRange to ""' in c for c in t.calls)


def test_prefix_edit_far_before_boundary_rejects_before_delete(monkeypatch):
    r=recovery(); a=rows(); b=rows(end=20); b[0][7]='f'*64
    s,t=bind(monkeypatch,[a,b]);r.checkpoint(s)
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==2


@pytest.mark.parametrize('bad',[[],[['checkpoint-state']],rows()[:-1],rows()+[['extra']]])
def test_bad_checkpoint_ack_quarantines(monkeypatch,bad):
    r=recovery();s,t=bind(monkeypatch,[bad])
    with pytest.raises(NativeWriterObjectError):r.checkpoint(s)
    assert s._quarantined
    with pytest.raises(NativeWriterObjectError):r.checkpoint(s)
    assert len(t.calls)==1


def test_append_field_table_and_tracking_clean_only_after_ack(monkeypatch):
    r=recovery();a=rows(); field=['field','main',1,'REF',' REF b ',8,15,16,19]
    table=['table',1,6,23,1,1]; b=rows(end=23,fields=[field],tables=[table])
    s,t=bind(monkeypatch,[a,b,[['rollback-ack',[field],[table]],*a]])
    r.checkpoint(s);h=NativeIndexHandle('s','new','n','REF');s._tracked_references=[(h,' REF b ')]
    assert r.rollback(s,6) is None
    assert s._tracked_references==[]
    script='\n'.join(t.calls[-1])
    assert script.index('delete recoveryField')<script.index('delete recoveryTable')<script.index('set content of rollbackRange to ""')


def test_failed_rollback_ack_preserves_tracking_and_discards_observation(monkeypatch):
    r=recovery();s,t=bind(monkeypatch,[rows(),rows(end=10),[['rollback-ack',[],[]],*rows(prefix='change')]])
    s._observed_field_topology=(('old',1),);r.checkpoint(s)
    s._observed_field_topology=(('new',2),);h=NativeIndexHandle('s','new','n','REF');s._tracked_references=[(h,' REF b ')]
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert s._tracked_references==[(h,' REF b ')] and not hasattr(s,'_observed_field_topology')
    assert s._quarantined


def test_tracking_prefix_rewrite_rejected_before_native_mutation(monkeypatch):
    r=recovery();h=NativeIndexHandle('s','old','n','REF');a=rows(bookmarks=[['bookmark','old',1,4,'abc']])
    s,t=bind(monkeypatch,[a]);s._tracked_references=[(h,'abc')];r.checkpoint(s)
    s._tracked_references[0]=(h,'rewritten')
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==1


@pytest.mark.parametrize('error',[TimeoutError(),KeyboardInterrupt()])
def test_uncertain_rollback_quarantines_without_tracking_commit(monkeypatch,error):
    r=recovery();s,t=bind(monkeypatch,[rows(),rows(end=9),error]);r.checkpoint(s)
    h=NativeIndexHandle('s','new','n','REF');s._tracked_references=[(h,'new')]
    with pytest.raises((NativeWriterObjectError,KeyboardInterrupt)):r.rollback(s,6)
    assert s._quarantined and s._tracked_references==[(h,'new')]
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==3


def test_full_prefix_hash_is_streamed_without_shell_or_argv():
    r=recovery();script='\n'.join(r.hash_commands('privateText','hashValue'))
    assert 'do shell script' not in script and 'quoted form' not in script
    assert 'setArguments:{"-a", "256"}' in script
    assert script.index("recoveryTask's |launch|()")<script.index('writeData:')<script.index("closeFile()")


def test_hash_helper_native_synthetic_values(tmp_path):
    import subprocess, json, sys
    if sys.platform!='darwin':pytest.skip('Foundation native synthetic helper')
    r=recovery()
    values=['','中文😀','quote\'";$(touch /tmp/unwanted)\nline\r','文😀'*100000]
    for index,value in enumerate(values):
        # Synthetic text may be saved in this test, never a private Word prefix.
        input_path=tmp_path/f'input-{index}.txt';input_path.write_bytes(value.encode())
        script=tmp_path/f'hash-{index}.applescript'
        script.write_text('use framework "Foundation"\nuse scripting additions\nset privateText to (current application\'s NSString\'s stringWithContentsOfFile:'+json.dumps(str(input_path))+' encoding:4 |error|:(missing value)) as text\n'+'\n'.join(r.hash_commands('privateText','hashValue'))+'\nreturn hashValue\n')
        result=subprocess.run(['/usr/bin/osascript',str(script)],capture_output=True,text=True,timeout=20)
        assert result.returncode==0,result.stderr
        assert result.stdout.strip()==hashlib.sha256(value.encode()).hexdigest()


def test_native_scripts_compile_without_word_launch(monkeypatch,tmp_path):
    import subprocess,sys
    if sys.platform!='darwin':pytest.skip('Word dictionary compile only')
    from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
    r=recovery();s,t=bind(monkeypatch,[rows()]);r.checkpoint(s)
    old=s._degradation_checkpoints[6]
    inside=('field','main',1,'REF',' REF "quoted" ',8,15,16,19)
    outside=('field','main',2,'REF',' REF outside ',25,32,33,35)
    table=('table',1,6,23,1,1)
    current=r._parse(s,rows(end=40,fields=[inside,outside],tables=[table]),6)
    scripts=[r.state_commands(()),r.rollback_commands(old,current,r._preflight(old,current))]
    for index,lines in enumerate(scripts):
        path=tmp_path/f'{index}.applescript';path.write_text(_JSON+'\ntell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell\n')
        result=subprocess.run(['/usr/bin/osacompile','-o',str(path.with_suffix('.scpt')),str(path)],capture_output=True,text=True,timeout=20)
        assert result.returncode==0,result.stderr


def test_untracked_bookmark_deleted_or_retargeted_rejects(monkeypatch):
    r=recovery();a=rows();a.insert(-1,['bookmark-hash','target_name',0,4,'a'*64]);a[-3][-1]=1
    s,t=bind(monkeypatch,[a,rows(end=10)])
    assert r.checkpoint(s)==6
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==2


@pytest.mark.parametrize('fail_rollback',[False,True])
def test_real_executor_recovery_controller(monkeypatch,fail_rollback):
    from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
    from skills.WPSComposer.scripts.generation_plan import GenerationOperation
    r=recovery();a=rows();b=rows(end=20)
    s,t=bind(monkeypatch,[a,[['partial-ack']],b,[] if fail_rollback else [['rollback-ack',[],[]],*a],[['fallback-ack']]])
    # Only transport and the deliberately failing primitive are doubles. The
    # production controller, checkpoint, transaction and tracking commit run.
    monkeypatch.setattr(s,'degradation_checkpoint',lambda:r.checkpoint(s),raising=False)
    monkeypatch.setattr(s,'rollback_degradation_checkpoint',lambda token:r.rollback(s,token),raising=False)
    calls=[]
    def native(**kwargs):
        calls.append('native');s._execute(['set nativeRows to {{"partial-ack"}}'])
        raise NativeWriterObjectError('CROSS_REFERENCE_FAILED')
    def fallback(**kwargs):
        calls.append('fallback');s._execute(['set nativeRows to {{"fallback-ack"}}'])
    monkeypatch.setattr(s,'add_cross_reference_paragraph',native)
    monkeypatch.setattr(s,'add_cross_reference_fallback',fallback)
    executor=WindowsLongformExecutor()
    op=GenerationOperation(op='writer.add_cross_reference',args={'runs':[{'text':'fallback'}]},node_id='ref:partial',failure_policy={'mode':'degrade','recoverableCodes':['CROSS_REFERENCE_FAILED'],'fallback':'inline-fallback'})
    if fail_rollback:
        with pytest.raises(Exception):executor._run_op(s,op)
        assert calls==['native'] and executor._issues==[] and s._quarantined
    else:
        executor._run_op(s,op)
        assert calls==['native','fallback'] and len(executor._issues)==1
        assert executor._issues[0].code=='CROSS_REFERENCE_FAILED'


def test_readonly_field_observation_does_not_make_coordinate_ambiguous(monkeypatch):
    r=recovery();s,t=bind(monkeypatch,[rows(),rows()]);r.checkpoint(s)
    s._observed_field_topology=(('observed',1),)
    assert r.checkpoint(s)==6


def test_new_bookmark_sorting_before_old_is_safe_when_wholly_appended(monkeypatch):
    r=recovery();old=['bookmark-hash','z_old',0,4,'a'*64];new=['bookmark-hash','a_new',6,8,'b'*64]
    a=rows();a.insert(-1,old);a[-3][-1]=1;b=rows(end=10);b[-2][-1]=2;b[-1:-1]=[new,old]
    s,t=bind(monkeypatch,[a,b,[['rollback-ack',[],[]],*a]]);r.checkpoint(s)
    assert r.rollback(s,6) is None
    assert any('delete bookmark "a_new"' in line for line in t.calls[-1])


def test_public_session_forwards_match_writer_contract(monkeypatch):
    import inspect
    from skills.WPSComposer.scripts.writer import WriterComposer
    for name in ('degradation_checkpoint','rollback_degradation_checkpoint'):
        assert hasattr(MacWordSession,name)
        assert inspect.signature(getattr(MacWordSession,name))==inspect.signature(getattr(WriterComposer,name))
    a=rows();s,t=bind(monkeypatch,[a,a,[['rollback-ack',[],[]],*a]])
    assert s.degradation_checkpoint()==6
    assert s.rollback_degradation_checkpoint('6') is None


def test_real_execute_stale_binding_quarantines_and_blocks_later_events(monkeypatch,tmp_path):
    import subprocess
    from skills.WPSComposer.scripts.msoffice import macos_word_session as m
    s=MacWordSession();s.staging_root=tmp_path;s._bound_path=str(tmp_path/'owned.docx');s._owns_doc=True
    calls=[]
    def run(command,**kwargs):
        calls.append(command)
        return subprocess.CompletedProcess(command,1,'','execution error: WPSC_STALE_DOCUMENT (-2700)')
    monkeypatch.setattr(m.subprocess,'run',run)
    with pytest.raises(NativeWriterObjectError):s.degradation_checkpoint()
    assert s._quarantined
    with pytest.raises(NativeWriterObjectError):s.degradation_checkpoint()
    assert len(calls)==1


def test_readonly_checkpoint_reads_but_rollback_cannot_mutate(monkeypatch):
    r=recovery();s,t=bind(monkeypatch,[rows()]);s._read_only=True
    assert s.degradation_checkpoint()==6
    with pytest.raises(NativeWriterObjectError):s.rollback_degradation_checkpoint(6)
    assert len(t.calls)==1


def test_native_fixture_session_call_signatures_are_available():
    import ast,inspect
    from pathlib import Path
    path=Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_recovery.py'
    tree=ast.parse(path.read_text())
    for node in ast.walk(tree):
        if not isinstance(node,ast.Call) or not isinstance(node.func,ast.Attribute):continue
        if not isinstance(node.func.value,ast.Name) or node.func.value.id!='session':continue
        method=getattr(MacWordSession,node.func.attr,None)
        assert callable(method),node.func.attr
        if all(keyword.arg is not None for keyword in node.keywords):
            inspect.signature(method).bind(None,*[None for arg in node.args],**{keyword.arg:None for keyword in node.keywords})


def test_native_preflight_compares_case_sensitive_identity(monkeypatch):
    r=recovery();s,t=bind(monkeypatch,[rows()]);r.checkpoint(s)
    snap=s._degradation_checkpoints[6]
    code=r.rollback_commands(snap,snap.state,r._preflight(snap,snap.state))
    assert any('isEqualToArray:' in line and 'WPSC_CHECKPOINT_PREFLIGHT_CHANGED' in line for line in code)
    assert not any('if nativeRows is not' in line for line in code)


@pytest.mark.parametrize('object_fact',[
    ['table',1,3,10,1,1],
    ['field','main',1,'REF',' REF old ',4,7,8,10],
    ['bookmark-hash','crossing',3,9,'a'*64],
])
def test_crossing_original_objects_fail_before_deletion(monkeypatch,object_fact):
    r=recovery();a=rows();b=rows(end=12);b.insert(-1,object_fact);b[-3][{'table':3,'field':4,'bookmark-hash':5}[object_fact[0]]]=1
    s,t=bind(monkeypatch,[a,b]);r.checkpoint(s)
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==2 and not s._quarantined


def test_old_field_code_or_bounds_changes_fail_before_deletion(monkeypatch):
    r=recovery();field=['field','main',1,'REF',' REF old ',1,2,3,4]
    a=rows(fields=[field]);b=rows(end=12,fields=[field.copy()]);b[1][4]=' REF OLD '
    s,t=bind(monkeypatch,[a,b]);r.checkpoint(s)
    with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
    assert len(t.calls)==2


def test_old_handles_instances_and_immutable_tracking_snapshot_survive(monkeypatch):
    r=recovery();h=NativeIndexHandle('s','old','n','REF');a=rows(bookmarks=[['bookmark','old',1,4,'abc']])
    s,t=bind(monkeypatch,[a,rows(end=10,bookmarks=[['bookmark','old',1,4,'abc']]),[['rollback-ack',[],[]],*a]])
    s._tracked_references=[(h,'abc')];s._observed_field_topology=(('old',1),);r.checkpoint(s)
    snap=s._degradation_checkpoints[6];s._tracked_references.append((NativeIndexHandle('s','new','n2','REF'),'def'))
    assert snap.tracked_references_prefix==((h,'abc'),)
    r.rollback(s,6)
    assert s._tracked_references==[(h,'abc')] and s._tracked_references[0][0] is h
    assert s._observed_field_topology==( ('old',1), )


def test_foundation_native_fact_equality_is_exact(monkeypatch,tmp_path):
    import subprocess,sys
    from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
    if sys.platform!='darwin':pytest.skip('Foundation synthetic comparison')
    r=recovery();s,t=bind(monkeypatch,[rows()]);r.checkpoint(s)
    snap=s._degradation_checkpoints[6]
    code=r.rollback_commands(snap,snap.state,r._preflight(snap,snap.state))
    check=next(line for line in code if 'WPSC_CHECKPOINT_PREFLIGHT_CHANGED' in line)
    for index,change in enumerate((None,'case','diacritic','whitespace')):
        facts=[list(row) for row in snap.state]
        if change=='case':facts[0][10]='TARGET'
        elif change=='diacritic':facts[0][10]='tárget'
        elif change=='whitespace':facts[0][10]='target '
        path=tmp_path/f'{index}.applescript';path.write_text(_JSON+'set nativeRows to '+r._literal(facts)+'\n'+check+'\nreturn "accepted"\n')
        result=subprocess.run(['/usr/bin/osascript',str(path)],capture_output=True,text=True,timeout=10)
        assert (result.returncode==0)==(change is None)
        if change is not None:assert 'WPSC_CHECKPOINT_PREFLIGHT_CHANGED' in result.stderr


def test_real_execute_disconnect_quarantines_without_cleanup_event(monkeypatch,tmp_path):
    import subprocess
    from skills.WPSComposer.scripts.msoffice import macos_word_session as m
    class Lock:
        quarantine_path=tmp_path/'quarantine.json'
        def quarantine(self,data):self.quarantine_path.write_text('quarantined')
        def close(self):pass
    s=MacWordSession();s.staging_root=tmp_path;s._bound_path=str(tmp_path/'owned.docx');s._owns_doc=True;s.lock=Lock()
    calls=[]
    def run(command,**kwargs):
        calls.append(command);return subprocess.CompletedProcess(command,1,'','execution error: Connection is invalid. (-609)')
    monkeypatch.setattr(m.subprocess,'run',run)
    with pytest.raises(m.NativeWordError):s._execute(['set nativeRows to {{"attempt"}}'])
    assert s._quarantined
    with pytest.raises(m.NativeWordError):s.close()
    assert len(calls)==1


@pytest.mark.parametrize('changed',[True,False])
def test_floating_shape_identity_is_preserved_or_blocks_rollback(monkeypatch,changed):
    r=recovery()
    shape=['shape',1,'Box','text box',100,200,0,1,0,0,50,20,0,1,True,'a'*64]
    a=rows();a.insert(-1,shape);a[-3]=['objects',1,0,0,0,0]
    b=deepcopy(a)
    if changed:b[-2][2]='Changed box'
    s,t=bind(monkeypatch,[a,b,[['rollback-ack',[],[]],*a]])
    assert r.checkpoint(s)==6
    if changed:
        with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
        assert len(t.calls)==2
    else:assert r.rollback(s,6) is None


def test_appended_floating_shape_prevents_controller_fallback(monkeypatch):
    from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
    from skills.WPSComposer.scripts.generation_plan import GenerationOperation
    r=recovery();a=rows();b=rows();b.insert(-1,['shape',1,'New','text box',100,200,0,1,0,0,50,20,0,1,True,'a'*64]);b[-3]=['objects',1,0,0,0,0]
    s,t=bind(monkeypatch,[a,[['native-shape-created']],b]);fallback=[]
    def partial(**kwargs):
        s._execute(['set nativeRows to {{"native-shape-created"}}']);raise NativeWriterObjectError('CROSS_REFERENCE_FAILED')
    monkeypatch.setattr(s,'add_cross_reference_paragraph',partial)
    monkeypatch.setattr(s,'add_cross_reference_fallback',lambda **kwargs:fallback.append(1))
    executor=WindowsLongformExecutor();op=GenerationOperation(op='writer.add_cross_reference',args={'runs':[{'type':'text','text':'fallback'}]},node_id='shape-partial',failure_policy={'mode':'degrade','recoverableCodes':['CROSS_REFERENCE_FAILED'],'fallback':'inline-fallback'})
    with pytest.raises(Exception):executor._run_op(s,op)
    assert not fallback and not executor._issues and len(t.calls)==3


def test_missing_tracked_identity_cannot_form_checkpoint(monkeypatch):
    r=recovery();s,t=bind(monkeypatch,[rows()]);s._tracked_references=[(NativeIndexHandle('s','old','n','REF'),'abc')]
    with pytest.raises(NativeWriterObjectError):r.checkpoint(s)
    assert s._quarantined


def test_ack_missing_native_table_row_is_invalid(monkeypatch):
    r=recovery();a=rows();a[-2]=['objects',0,0,1,0,0]
    s,t=bind(monkeypatch,[a])
    with pytest.raises(NativeWriterObjectError):r.checkpoint(s)
    assert s._quarantined


@pytest.mark.parametrize('changed',[True,False])
def test_generic_inline_object_preservation_or_typed_rejection(monkeypatch,changed):
    r=recovery();item=['inline-shape',1,100,200,0,1,80,20,'inline horizontal line','a'*64]
    a=rows();a[-2][2]=1;a.insert(-1,item);b=deepcopy(a)
    if changed:b[-2][8]='inline OLE object'
    s,t=bind(monkeypatch,[a,b,[['rollback-ack',[],[]],*a]])
    assert r.checkpoint(s)==6
    if changed:
        with pytest.raises(NativeWriterObjectError):r.rollback(s,6)
        assert len(t.calls)==2
    else:assert r.rollback(s,6) is None


def test_new_generic_inline_object_prevents_controller_fallback(monkeypatch):
    from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
    from skills.WPSComposer.scripts.generation_plan import GenerationOperation
    a=rows();b=rows();b[-2][2]=1;b.insert(-1,['inline-shape',1,100,200,0,1,80,20,'inline horizontal line','a'*64])
    s,t=bind(monkeypatch,[a,[['native-inline-created']],b]);fallback=[]
    def partial(**kwargs):
        s._execute(['set nativeRows to {{"native-inline-created"}}']);raise NativeWriterObjectError('CROSS_REFERENCE_FAILED')
    monkeypatch.setattr(s,'add_cross_reference_paragraph',partial)
    monkeypatch.setattr(s,'add_cross_reference_fallback',lambda **kwargs:fallback.append(1))
    executor=WindowsLongformExecutor();op=GenerationOperation(op='writer.add_cross_reference',args={'runs':[{'type':'text','text':'fallback'}]},node_id='inline-partial',failure_policy={'mode':'degrade','recoverableCodes':['CROSS_REFERENCE_FAILED'],'fallback':'inline-fallback'})
    with pytest.raises(Exception):executor._run_op(s,op)
    assert not fallback and not executor._issues and len(t.calls)==3


def test_generic_inline_count_cannot_ack_without_identity_row(monkeypatch):
    r=recovery();a=rows();a[-2][2]=1;s,t=bind(monkeypatch,[a])
    with pytest.raises(NativeWriterObjectError):r.checkpoint(s)
    assert s._quarantined


def test_fixture_sentinel_cleanup_is_independent_and_follows_owned_context():
    import ast
    from pathlib import Path
    source=(Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_recovery.py').read_text()
    tree=ast.parse(source);main=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='main')
    primary=next(n for n in ast.walk(main) if isinstance(n,ast.With))
    assert 'close d saving no' not in ast.get_source_segment(source,primary)
    cleanup=[n for n in ast.walk(main) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='close_sentinel']
    assert len(cleanup)==1 and cleanup[0].lineno>primary.end_lineno


def test_fixture_independent_sentinel_close_uses_exact_guard(monkeypatch,tmp_path):
    import subprocess,json
    from fixtures.microsoft_parity import macos_word_recovery as fixture
    assert hasattr(fixture,'close_sentinel')
    commands=[]
    def run(command,**kwargs):
        from pathlib import Path
        commands.append(Path(command[-1]).read_text())
        return subprocess.CompletedProcess(command,0,json.dumps([['sentinel-closed','Test sentinel']]),'')
    monkeypatch.setattr(fixture.subprocess,'run',run)
    assert fixture.close_sentinel(tmp_path,'Test sentinel','token')==[['sentinel-closed','Test sentinel']]
    script=commands[0]
    assert 'document "Test sentinel"' in script and 'SENTINEL_CHANGED' in script and 'SENTINEL_SAVED' in script
    assert 'boundDoc' not in script and 'close sentinelDocument saving no' in script


def real_transport(monkeypatch, tmp_path, outcome):
    """Keep the actual session envelope; replace only the external process."""
    import json
    from pathlib import Path
    from types import SimpleNamespace
    from skills.WPSComposer.scripts.msoffice import macos_word_session as m
    s = MacWordSession(); s.staging_root = tmp_path
    s._owns_doc = True; s._bound_path = str(tmp_path/'owned.docx')
    reasons = []
    s.lock = SimpleNamespace(quarantine_path=tmp_path/'quarantine.json',
                             quarantine=lambda data: reasons.append(data['reason']), close=lambda: None)
    calls = []
    observed = []
    def run(command, **kwargs):
        source = Path(command[1]).read_text()
        calls.append(source)
        observed.append(getattr(s, '_observed_field_topology', None))
        if len(calls) == 1: result = rows()
        elif len(calls) == 2: result = rows(end=10)
        elif len(calls) > 3: result = [['ok']]
        else:
            assert 'set content of rollbackRange to ""' in source
            if isinstance(outcome, BaseException): raise outcome
            if outcome == 'ordinary-error':
                return m.subprocess.CompletedProcess(command, 1, '', 'execution error: WPSC_CHECKPOINT_POSTCONDITION_FAILED (-2700)')
            result = outcome
        return m.subprocess.CompletedProcess(command, 0, json.dumps([m._COMPLETION_MARKER, result]), '')
    monkeypatch.setattr(m.subprocess, 'run', run)
    s._observed_field_topology = (('baseline', 1),)
    checkpoint = s.degradation_checkpoint()
    s._observed_field_topology = (('appended', 2),)
    s._pending_heading = (6, 10, 1)
    handle = NativeIndexHandle('s', 'new', 'n', 'REF')
    s._tracked_references = [(handle, ' REF new ')]
    return s, checkpoint, calls, observed, reasons


@pytest.mark.parametrize('outcome', ['ordinary-error', [], [['rollback-ack', [], []], *rows(prefix='change')],
                                     KeyboardInterrupt(), OSError('connection unavailable'),
                                     subprocess.TimeoutExpired('osascript', 1, stderr='private timeout evidence')])
def test_real_destructive_rollback_failure_isolates_session(monkeypatch, tmp_path, outcome):
    from skills.WPSComposer.scripts.msoffice.macos_word_session import NativeWordError
    s, token, calls, observed, reasons = real_transport(monkeypatch, tmp_path, outcome)
    tracking = list(s._tracked_references)
    expected_error = KeyboardInterrupt if isinstance(outcome, KeyboardInterrupt) else NativeWriterObjectError
    with pytest.raises(expected_error) as caught:
        with s:
            s.rollback_degradation_checkpoint(token)
    if expected_error is NativeWriterObjectError:
        assert caught.value.code == 'LOCAL_MUTATION_ROLLBACK_FAILED'
    assert s._quarantined and reasons
    if expected_error is NativeWriterObjectError:
        assert s._retain_evidence
    assert list(tmp_path.glob('*.log'))
    assert observed[-1] is None
    assert not hasattr(s, '_observed_field_topology')
    assert s._tracked_references == tracking and s._pending_heading == (6, 10, 1)
    with pytest.raises(NativeWordError) as later:
        s._execute(['set nativeRows to {{"unexpected-write"}}'])
    assert later.value.code == 'NATIVE_WORD_QUARANTINED'
    assert len(calls) == 3
    if outcome == 'ordinary-error':
        assert any('execution error: WPSC_CHECKPOINT_POSTCONDITION_FAILED (-2700)' in p.read_text() for p in tmp_path.glob('*.log'))


@pytest.mark.parametrize('failure', ['argument', 'preimage', 'script-write', 'deadline', 'raw-quarantine'])
def test_real_rollback_presubmit_failure_preserves_tracking_and_observation(monkeypatch, tmp_path, failure):
    from pathlib import Path
    from skills.WPSComposer.scripts.msoffice import macos_word_session as m
    s, token, calls, observed, reasons = real_transport(monkeypatch, tmp_path, None)
    tracking = list(s._tracked_references)
    if failure == 'argument': token = 'invalid'
    if failure == 'preimage': token = 999
    if failure == 'raw-quarantine':
        original_compile = recovery().rollback_commands
        def compile_commands(*args):
            commands = original_compile(*args)
            s._quarantined = True
            return commands
        monkeypatch.setattr(recovery(), 'rollback_commands', compile_commands)
    original_write = Path.write_text
    def write(path, data, **kwargs):
        if 'set content of rollbackRange to ""' in data:
            if failure == 'script-write': raise OSError('local disk full')
            if failure == 'deadline': s._deadline = -1
        return original_write(path, data, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(NativeWriterObjectError) as caught:
        s.rollback_degradation_checkpoint(token)
    assert caught.value.code == 'LOCAL_MUTATION_ROLLBACK_FAILED'
    assert len(calls) == (1 if failure in ('argument', 'preimage') else 2)
    assert s._tracked_references == tracking
    assert s._observed_field_topology == (('appended', 2),)
    assert s._pending_heading == (6, 10, 1)
    assert not getattr(s, '_field_topology_mutation_pending', False)
    if failure == 'deadline':
        # Existing session deadline quarantine is authoritative, not partial mutation.
        assert s._quarantined
        assert reasons == ['Session deadline expired before verified document close']
    elif failure == 'raw-quarantine':
        assert s._quarantined and not reasons
    else:
        assert not s._quarantined and not reasons


def test_real_rollback_restores_checkpoint_cache_only_after_valid_ack(monkeypatch, tmp_path):
    s, token, calls, observed, reasons = real_transport(monkeypatch, tmp_path, [['rollback-ack', [], []], *rows()])
    assert s.rollback_degradation_checkpoint(token) is None
    assert observed == [(('baseline', 1),), (('appended', 2),), None]
    assert s._observed_field_topology == (('baseline', 1),)
    assert s._tracked_references == [] and s._pending_heading is None
    assert s._structural_changed and not s._quarantined
