"""Native boundary contracts; subprocess doubles cannot certify Word layout."""
from copy import deepcopy
import importlib
import json
from pathlib import Path
import subprocess
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, _COMPLETION_MARKER


def module():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_quality')
    except ModuleNotFoundError:
        pytest.fail('quality native boundary is missing')


def state():
    return [['quality-state', 30, 'a'*64, 'b'*64],
            ['paragraph', 0, 10, 'c'*64, 'd'*64],
            ['paragraph', 10, 30, 'e'*64, 'f'*64],
            ['bookmark', 'wpsc_document_quality_anchor', 10, 10, '0'*64],
            ['layout','9'*64],['quality-state-end']]


def session(tmp_path):
    s=MacWordSession();s.staging_root=tmp_path;s._bound_path='/private/Owned.docx';s._owns_doc=True
    s.lock=SimpleNamespace(quarantine_path=tmp_path/'quarantine.json',quarantine=lambda data:None)
    s._observed_field_topology=(('before',1),);s._pending_heading=(0,4,1)
    return s


def target(s):
    return module().QualityTarget(module()._identity(s),10,30,tuple(module()._freeze(state())),None)


def ack(s, *, mode='table'):
    return [['quality-insert',s._bound_path,10,10,15,'abc\r\x07\r\x07',30,35,1,1,1,10,13,'abc'],
            ['degradation-style',True,True,True],['degradation-paragraph',0,3,True,True],
            ['degradation-row',False],['preimage',state()],['preserved',state()]]


def test_middle_ack_uses_real_ordinal_and_delta(tmp_path):
    s=session(tmp_path);m=module();t=target(s)
    assert m._table_ack(ack(s),t,'abc') == (10,15,'abc\r\x07\r\x07',1)


@pytest.mark.parametrize('change', ['path-case','bool-position','wrong-delta','wrong-ordinal','extra-row','wrong-text','style','suffix','boundary-format','bookmark-case'])
def test_forged_middle_ack_rejected(tmp_path,change):
    s=session(tmp_path);m=module();t=target(s);r=ack(s)
    if change=='path-case':r[0][1]='/private/owned.docx'
    elif change=='bool-position':r[0][2]=True
    elif change=='wrong-delta':r[0][7]=34
    elif change=='wrong-ordinal':r[0][8]=2
    elif change=='extra-row':r[0][9]=2
    elif change=='wrong-text':r[0][13]='Abc'
    elif change=='style':r[1][1]=1
    elif change=='suffix':r[-1][1][0][3]='B'*64
    elif change=='boundary-format':r[-1][1][2][4]='1'*64
    else:r[-1][1][3][1]='WPSC_DOCUMENT_QUALITY_ANCHOR'
    assert m._table_ack(r,t,'abc') is None


@pytest.mark.parametrize('bad', [True,-1,10.5,'10'])
def test_numeric_target_invalid_before_native(tmp_path,monkeypatch,bad):
    s=session(tmp_path);calls=[]
    monkeypatch.setattr(s,'_execute',lambda lines:calls.append(lines))
    with pytest.raises(Exception):module()._target_at(s,bad)
    assert not calls


def test_foreign_descriptor_fails_without_native(tmp_path,monkeypatch):
    s=session(tmp_path);t=target(s);other=session(tmp_path);calls=[]
    monkeypatch.setattr(other,'_execute',lambda lines:calls.append(lines))
    with pytest.raises(Exception):module()._insert_box_native(other,t,'abc')
    assert not calls and not other._quarantined


@pytest.mark.parametrize('failure',['script-write','deadline','read-only','timeout','ordinary-error','json','malformed-ack'])
def test_actual_submission_controls_invalidation_and_quarantine(tmp_path,monkeypatch,failure):
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=session(tmp_path);m=module();t=target(s);calls=[];original=Path.write_text
    def write(p,data,**kwargs):
        if failure=='script-write':raise OSError('disk full')
        if failure=='deadline':s._deadline=-1
        return original(p,data,**kwargs)
    monkeypatch.setattr(Path,'write_text',write)
    def run(command,**kwargs):
        calls.append(command)
        if failure=='timeout':raise subprocess.TimeoutExpired(command,1)
        if failure=='ordinary-error':return subprocess.CompletedProcess(command,1,'','execution error: ordinary (-2700)')
        payload='bad' if failure=='json' else json.dumps([_COMPLETION_MARKER,[['malformed']]])
        return subprocess.CompletedProcess(command,0,payload,'')
    monkeypatch.setattr(transport.subprocess,'run',run)
    if failure=='read-only':s._read_only=True
    with pytest.raises(Exception):m._insert_box_native(s,t,'abc')
    if failure in ('script-write','deadline','read-only'):
        assert not calls and s._observed_field_topology==( ('before',1), )
        assert s._pending_heading==(0,4,1)
    else:
        assert len(calls)==1 and s._quarantined
        assert not hasattr(s,'_observed_field_topology')
        assert s._pending_heading==(0,4,1) and not s._structural_changed
    assert not getattr(s,'_field_topology_mutation_pending',False)


@pytest.mark.parametrize('changed',[False,True])
def test_fallback_requires_exact_restoration_ack(tmp_path,monkeypatch,changed):
    s=session(tmp_path);m=module();t=target(s);restored=state()
    if changed:restored[2][4]='1'*64
    replies=[[['quality-restored',s._bound_path,10],['preimage',state()],['restored',restored]],
             [['quality-range',s._bound_path,10,13,'abc',30,33],['degradation-style',True,True,True],
              ['degradation-paragraph',0,3,True,True],['preimage',state()],['preserved',state()]]]
    calls=[]
    def execute(lines):calls.append(lines);return replies.pop(0)
    monkeypatch.setattr(s,'_execute',execute)
    if changed:
        with pytest.raises(Exception):m._insert_box_native(s,t,'abc')
        assert len(calls)==1 and s._quarantined
    else:
        box=m._insert_box_native(s,t,'abc')
        assert len(calls)==2 and box.Range.Start==10 and box.Range.End==13 and box.table_index is None
        assert not s._quarantined


def test_compiler_guards_preimage_before_delete_and_rechecks_after(tmp_path):
    s=session(tmp_path);m=module();lines=m._table_commands(s,target(s),'abc');source='\n'.join(lines)
    assert source.index('WPSC_QUALITY_TARGET_STALE')<source.index('make new table')
    assert source.index('WPSC_QUALITY_PARTIAL_UNSAFE')<source.index('delete qualityNewTable')
    assert source.index('delete qualityNewTable')<source.index('WPSC_QUALITY_RESTORE_FAILED')
    assert 'selection of boundWindow' in source and 'isEqualToString:' in source
    assert 'undo' not in source.lower() and 'rollback_degradation' not in source


def test_empty_reservation_keeps_text_topology_and_pending_heading(tmp_path,monkeypatch):
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=session(tmp_path)
    payload=json.dumps([_COMPLETION_MARKER,[['quality-reserved',s._bound_path,3,9,30,9,9,30,'a'*64,'a'*64]]])
    monkeypatch.setattr(transport.subprocess,'run',lambda cmd,**kw:subprocess.CompletedProcess(cmd,0,payload,''))
    assert module()._reserve_native(s)==9
    assert s._observed_field_topology==( ('before',1), ) and s._pending_heading==(0,4,1)
    assert not s._structural_changed


def test_stale_target_readback_does_not_substitute_document_end(tmp_path,monkeypatch):
    s=session(tmp_path)
    monkeypatch.setattr(s,'_execute',lambda lines:[['quality-target',s._bound_path,29,30],['preimage',state()]])
    with pytest.raises(Exception):module()._target_at(s,31)
    assert s._quarantined


@pytest.mark.parametrize('bad',[None,True,[],['x'],[None],[['quality-state',True,'a'*64,'b'*64],['quality-state-end']]])
def test_state_parser_malformed_input_fails_closed(bad):
    assert module()._valid_state(bad) is False


def test_target_guard_compares_native_serializations_not_json_spelling(tmp_path):
    s=session(tmp_path);m=module();src='\n'.join(m._target_guard(s,target(s)))
    assert 'my jsonRows(qualityExpectedState)' in src


@pytest.mark.skipif(__import__('sys').platform!='darwin',reason='Mac dictionary compiler only')
@pytest.mark.parametrize('mode',['table','range','target'])
def test_quality_scripts_compile_without_running_word(tmp_path,mode):
    from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
    s=session(tmp_path);m=module();t=target(s)
    lines=m._table_commands(s,t,'中文😀') if mode=='table' else m._range_commands(s,t,'中文😀') if mode=='range' else m._target_guard(s,t)
    src=tmp_path/'quality.applescript'
    src.write_text(_JSON+'\ntell application "Microsoft Word"\n'+ '\n'.join(lines)+'\nend tell\n')
    result=subprocess.run(['/usr/bin/osacompile','-o',str(tmp_path/'quality.scpt'),str(src)],capture_output=True,text=True,timeout=20)
    assert result.returncode==0,result.stderr


def test_snapshot_rejects_missing_section_and_list_identity():
    m=module();rows=[r for r in state() if r[0]!='layout']
    assert not m._valid_state(rows)


@pytest.mark.parametrize('submitted',[False,True])
def test_reservation_failure_quarantines_only_submitted_uncertainty(tmp_path,monkeypatch,submitted):
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=session(tmp_path);calls=[]
    if not submitted:
        monkeypatch.setattr(Path,'write_text',lambda *a,**kw:(_ for _ in ()).throw(OSError('disk')))
    def run(cmd,**kw):
        calls.append(cmd);return subprocess.CompletedProcess(cmd,1,'','execution error: bookmark partial (-2700)')
    monkeypatch.setattr(transport.subprocess,'run',run)
    with pytest.raises(Exception):module()._reserve_native(s)
    assert s._quarantined is submitted
    assert s._observed_field_topology==( ('before',1), ) and s._pending_heading==(0,4,1)
    assert not s._structural_changed


def test_middle_table_ordinal_is_not_total_count(tmp_path):
    s=session(tmp_path);m=module();before=state()
    before[3:3]=[['field',1,'field sequence',2,3,3,4,'1'*64,'2'*64,False],
                 ['field',2,'field reference',20,21,21,22,'3'*64,'4'*64,True],
                 ['table',1,5,8,1,1,'5'*64,[False]],
                 ['table',2,24,27,1,1,'6'*64,[True]]]
    t=m.QualityTarget(m._identity(s),10,30,m._freeze(before),None)
    r=ack(s);r[0][8]=2;r[-2][1]=deepcopy(before);r[-1][1]=deepcopy(before)
    assert m._table_ack(r,t,'abc')==(10,15,'abc\r\x07\r\x07',2)
    r[0][8]=3
    assert m._table_ack(r,t,'abc') is None
    r[0][8]=2;r[-1][1][3][7]='7'*64
    assert m._table_ack(r,t,'abc') is None


@pytest.mark.parametrize('bad',[[[],None],[True,['quality-state-end']], [['quality-state',30,'a'*64,'b'*64],None]])
def test_nested_invalid_state_cannot_escape_parser(bad):
    assert module()._valid_state(bad) is False


def test_reservation_log_failure_preserves_topology_and_never_commits_or_retries(tmp_path,monkeypatch):
    """Permanent regression for the independent review's real-transport repro."""
    from skills.WPSComposer.scripts.msoffice import macos_word_session as transport
    s=session(tmp_path);calls=[];original=Path.write_text
    payload=json.dumps([_COMPLETION_MARKER,[['quality-reserved',s._bound_path,3,9,30,9,9,30,'a'*64,'a'*64]]])
    def run(cmd,**kw):calls.append(cmd);return subprocess.CompletedProcess(cmd,0,payload,'')
    def write(path,data,**kw):
        if path.suffix=='.log':raise OSError('post-submission log failure')
        return original(path,data,**kw)
    monkeypatch.setattr(transport.subprocess,'run',run);monkeypatch.setattr(Path,'write_text',write)
    with pytest.raises(Exception,match='quality anchor insertion failed'):
        module().reserve_document_quality_anchor(s)
    assert s._quarantined and len(calls)==1
    assert not hasattr(s,'_quality_notice_anchor_position') and not hasattr(s,'_quality_notice_title')
    assert s._observed_field_topology==( ('before',1), ) and s._pending_heading==(0,4,1) and not s._structural_changed
    with pytest.raises(Exception,match='quality anchor insertion failed'):
        module().reserve_document_quality_anchor(s)
    assert len(calls)==1


def test_table_gate_uses_the_native_proven_ordinal_lookup():
    """Guard the resolver distinction reproduced in target-read-probe-01."""
    lines=module()._position_gate()
    index=lines.index('repeat with qualityTableOrdinal from 1 to count tables of boundDoc')
    assert lines[index+1]=='set qt to table qualityTableOrdinal of boundDoc'
    assert 'repeat with qt in tables of boundDoc' not in lines


@pytest.mark.parametrize(('start','end','reject'),[
    (59,77,False),(173,191,False),
    (100,109,False),(100,110,True),(100,111,True),
    (111,130,True),(112,130,True),(113,130,False),(100,130,True),
])
def test_table_gate_keeps_its_inclusive_one_unit_neighbor_rejection(start,end,reject):
    """Evaluate the emitted scalar condition; this does not simulate Word getters."""
    line=next(s for s in module()._position_gate() if 'then error "WPSC_QUALITY_TABLE_POSITION_UNVERIFIED"' in s)
    condition=line.removeprefix('if ').split(' then error ',1)[0]
    condition=condition.replace('(start of content of text object of qt)','table_start').replace('(end of content of text object of qt)','table_end')
    result=eval(condition,{'__builtins__':{}},{'table_start':start,'table_end':end,'qualityPoint':111})
    assert result is reject


def test_retained_native_probe_discriminates_resolver_without_changing_comparison():
    root=Path(__file__).resolve().parents[2]
    evidence=json.loads((root/'docs/verification/microsoft-parity/macos-word-quality/target-read-probe-01/report.json').read_text(encoding='utf-8'))
    rows={step['mode']:step['rows'] for step in evidence['steps']}
    assert rows['original']==[['bounds',111,110,112],['expected-read-error','compound-comparison',-2763]]
    assert rows['indexed-original']==[['bounds',111,110,112],['compound',1,False],['compound',2,False],['done']]
    # Preserve the raw FAIL: the probe stopped on an unexpectedly successful
    # diagnostic branch; parent cleanup is independent evidence, not a rerun.
    assert evidence['status']=='FAIL'
    cleanup=json.loads((root/'docs/verification/microsoft-parity/macos-word-quality/target-read-probe-01-parent-cleanup/report.json').read_text(encoding='utf-8'))
    assert cleanup['after']==[] and cleanup['quarantine_recovered'] is True


@pytest.mark.parametrize(('plural','singular','variable','ordinal'),[
    ('fields','field','qf','qualityFieldOrdinal'),
    ('bookmarks','bookmark','qb','qualityBookmarkOrdinal'),
])
def test_other_position_guards_resolve_native_objects_by_ordinal(plural,singular,variable,ordinal):
    """Same collection-reference boundary; these two native paths remain unprobed."""
    lines=module()._position_gate()
    index=lines.index(f'repeat with {ordinal} from 1 to count {plural} of boundDoc')
    assert lines[index+1]==f'set {variable} to {singular} {ordinal} of boundDoc'
    assert f'repeat with {variable} in {plural} of boundDoc' not in lines


def test_ordinal_resolution_patch_preserves_every_position_rejection_condition():
    """The authorized resolver-only fix must not weaken its admission gates."""
    root=Path(__file__).resolve().parents[2]
    raw=(root/'docs/verification/microsoft-parity/macos-word-quality/run-01/native-runtime/0efcad4e246b4b5d9c9dcf3ff66d37fc.applescript').read_text(encoding='utf-8')
    old=[line for line in raw.split('set qualityDelta to 0',1)[0].splitlines() if line.startswith('if ') and 'then error "WPSC_QUALITY_' in line]
    current=[line for line in module()._position_gate() if line.startswith('if ')]
    assert current==old
