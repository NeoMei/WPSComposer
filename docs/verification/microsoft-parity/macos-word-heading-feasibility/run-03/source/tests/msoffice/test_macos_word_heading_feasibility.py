"""Pure guards and syntax only; never run Office feasibility during pytest."""
from __future__ import annotations
import importlib.util
from pathlib import Path
import subprocess
import sys
import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT/'fixtures/microsoft_parity/macos_word_heading_feasibility.py'

def module():
    assert FIXTURE.exists(), 'heading feasibility fixture absent'
    spec = importlib.util.spec_from_file_location('heading_feasibility', FIXTURE)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m


def test_execute_guard_has_no_filesystem_or_native_side_effect(tmp_path):
    m = module()
    assert m.main(['--mode','font-properties','--output',str(tmp_path/'absent')]) == 2
    assert not (tmp_path/'absent').exists()


@pytest.mark.parametrize('mode',['font-properties','paragraph-object','detached-properties','list-identity'])
def test_exact_native_rows_fail_closed(mode):
    m=module(); labels=m.expected_labels(mode)
    rows=[[label,True] for label in labels]
    assert m.native_valid(mode,rows)
    for bad in ([],rows+rows,rows[:-1],list(reversed(rows))):
        assert not m.native_valid(mode,bad)
    for i in range(len(rows)):
        bad=[r[:] for r in rows];bad[i][1]=1
        assert not m.native_valid(mode,bad)
        bad[i][1]=False
        assert not m.native_valid(mode,bad)


def test_whole_property_probe_does_not_assign_readonly_font_or_mask_errors():
    m=module(); source='\n'.join(m.clone_commands('detached-properties'))
    assert 'set properties of destinationFont to properties of sourceFont' in source
    assert 'set paragraph format of detachedStyle to paragraph format of sourceStyle' in source
    assert 'set font object of detachedStyle' not in source
    assert 'on error' not in source
    labels=m.expected_labels('detached-properties')
    for key in ('font:scaling','font:kerning','font-shading:texture','font-borders:outside line width','paragraph:widow control','paragraph-shading:background pattern color','paragraph-borders:outside color','paragraph:tabs'):
        assert key in labels


@pytest.mark.parametrize('value',['','bad','WPSCHeading_'+'A'*32,'WPSCHeading_'+'0'*31])
def test_template_names_are_owned_uuid_only(value):
    with pytest.raises(ValueError):module().list_commands(value)


def test_shared_list_has_four_levels_geometry_and_name_resolution():
    m=module(); source='\n'.join(m.list_commands('WPSCHeading_'+'a'*32))
    assert source.count('set linked style of ownLevel')==4
    assert 'set name of ownList' not in source  # creation supplies UUID name
    assert 'number position of ownLevel to 54' in source
    assert 'text position of ownLevel to 72' in source
    source='\n'.join(m.resolve_list('WPSCHeading_'+'a'*32))
    assert 'count list templates of boundDoc' in source and 'matches is not 1' in source
    assert 'isEqualToString:' in source


@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax compiler')
@pytest.mark.parametrize('mode',['font-properties','paragraph-object','detached-properties','list-identity'])
def test_all_phases_compile_without_launch(mode,tmp_path):
    m=module()
    for i,commands in enumerate(m.phases(mode,'WPSCHeading_'+'a'*32)):
        source="use framework \"Foundation\"\nuse scripting additions\ntell application \"/Applications/Microsoft Word.app\"\n"+'\n'.join(commands)+'\nend tell\n'
        path=tmp_path/('%s-%s.applescript'%(mode,i));path.write_text(source)
        r=subprocess.run(['/usr/bin/osacompile','-o',str(path.with_suffix('.scpt')),str(path)],capture_output=True,text=True,timeout=20)
        assert r.returncode==0,r.stderr


def test_xml_clone_rejects_lost_nested_appearance():
    m=module(); w=m.W
    r='<w:rPr><w:b/><w:shd w:fill="FCE8E6"/><w:bdr w:val="single" w:sz="6"/></w:rPr>'
    p='<w:pPr><w:spacing w:before="140"/><w:tabs><w:tab w:val="right" w:pos="800"/></w:tabs></w:pPr>'
    xml=('<w:styles xmlns:w="'+w[1:-1]+'"><w:style w:styleId="Heading1"><w:name w:val="heading 1"/>'+r+p+'</w:style><w:style w:styleId="Detached"><w:name w:val="WPSC Heading Clone"/><w:basedOn w:val="Normal"/>'+r+p+'</w:style></w:styles>').encode()
    assert m.clone_xml_valid(xml,'detached-properties')
    for a,b in [('<w:shd w:fill="FCE8E6"/>',''),('<w:bdr w:val="single" w:sz="6"/>',''),('<w:tab w:val="right" w:pos="800"/>','')]:
        first,second=xml.decode().split('<w:style w:styleId="Detached">')
        assert not m.clone_xml_valid((first+'<w:style w:styleId="Detached">'+second.replace(a,b)).encode(),'detached-properties')


def test_invalid_mode_never_emits_native():
    m=module()
    with pytest.raises(ValueError):m.phases('all','WPSCHeading_'+'a'*32)


def test_invalid_run_mode_rejected_before_source_copy_or_native(tmp_path,monkeypatch):
    m=module()
    monkeypatch.setattr(m,'retain_sources',lambda *a:pytest.fail('source copied before mode validation'))
    with pytest.raises(ValueError):m.run(tmp_path/'absent','all')
    assert not (tmp_path/'absent').exists()


@pytest.mark.skipif(sys.platform!='darwin',reason='local Word dictionary')
def test_all_dictionary_writable_scalars_are_compared():
    import xml.etree.ElementTree as ET
    m=module();root=ET.parse('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef').getroot()
    for name,properties in [('font',m.FONT),('paragraph format',m.PARAGRAPH),('shading',m.SHADING),('border options',m.BORDERS)]:
        cls=next(c for c in root.iter('class') if c.get('name')==name)
        assert tuple(p.get('name') for p in cls.findall('property') if p.get('access')!='r')==properties
    cls=next(c for c in root.iter('class') if c.get('name')=='Word style')
    assert next(p for p in cls.findall('property') if p.get('name')=='font object').get('access')=='r'


def test_list_xml_rejects_independent_lists_and_wrong_level_formats():
    m=module();ns=m.W[1:-1]
    styles='<w:styles xmlns:w="'+ns+'">'+''.join('<w:style w:styleId="Heading%s"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>'%i for i in range(1,5))+'</w:styles>'
    levels=''.join('<w:lvl w:ilvl="%s"><w:pStyle w:val="Heading%s"/><w:lvlText w:val="%s"/><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>'%(i-1,i,'.'.join('%%%s'%j for j in range(1,i+1))) for i in range(1,5))
    numbering='<w:numbering xmlns:w="'+ns+'"><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num><w:abstractNum w:abstractNumId="3">'+levels+'</w:abstractNum></w:numbering>'
    assert m.list_xml_valid(styles.encode(),numbering.encode())
    assert not m.list_xml_valid(styles.replace('w:val="7"','w:val="8"',1).encode(),numbering.encode())
    assert not m.list_xml_valid(styles.encode(),numbering.replace('%1.%2.%3.%4','%4').encode())
    assert not m.list_xml_valid(styles.encode(),numbering.replace('Heading4','Heading3').encode())


def test_reopen_commands_are_read_only_and_numbering_is_observed():
    m=module();name='WPSCHeading_'+'a'*32
    source='\n'.join(m.list_readback(name))
    assert source.count('set observedListString to list string')==4
    assert source.count('set actualNumberedLevel to')==4
    assert 'make new' not in source and 'set linked style of' not in source
    for mode in ('font-properties','paragraph-object','detached-properties'):
        source='\n'.join(m.clone_readback(mode))
        assert 'set properties of' not in source and 'set paragraph format of' not in source
        assert 'isEqualToString:' in source
    assert 'isEqualToString:' in '\n'.join(m.detached_readback())


def test_font_seed_omits_only_host_rejected_width_without_narrowing_readback():
    m=module()
    for mode in ('font-properties','detached-properties'):
        commands=m.seed_commands(mode)
        assert 'set outside line width of border options of sourceFont to line width75 point' not in commands
        assert 'set outside line style of border options of sourceFont to line style single' in commands
        assert 'set outside color of border options of sourceFont to {49344,49344,49344}' in commands
        assert 'font-borders:outside line width' in m.expected_labels(mode)
    assert 'set outside line width of border options of sourceParagraph to line width75 point' in m.seed_commands('paragraph-object')


def diagnostic_rows(m):
    return [[label,['number',17],['number',17],True] for label in m.expected_labels('font-properties')]


def test_order_diagnostic_schema_keeps_actual_mismatches_without_accepting_copy():
    m=module();rows=diagnostic_rows(m)
    assert m.diagnostic_rows_valid(rows)
    rows[0]=[rows[0][0],['number',17],['number',12],False]
    assert m.diagnostic_rows_valid(rows)
    report={'diagnostic':{}}
    m.record_observation(report,'immediate_after_copy',rows)
    assert report['diagnostic']['immediate_after_copy']['mismatch_labels']==[rows[0][0]]
    assert report['diagnostic']['immediate_after_copy']['appearance_equal'] is False
    for invalid in ([],rows+rows,rows[:-1],list(reversed(rows))):assert not m.diagnostic_rows_valid(invalid)
    invalid=[r[:] for r in rows];invalid[0]=[rows[0][0],['number',17],['number',12],True]
    assert not m.diagnostic_rows_valid(invalid)
    invalid[0]=[rows[0][0],['number',True],['number',12],False]
    assert not m.diagnostic_rows_valid(invalid)
    assert not m.native_valid('font-properties',[[r[0],r[3]] for r in rows])


def test_order_diagnostic_copy_precedes_readback_and_isolated_base_reset():
    m=module();phases=m.diagnostic_phases()
    assert len(phases)==5
    copy=phases[1]
    assert copy.index('set base style of detachedStyle to style normal') < copy.index('set properties of destinationFont to properties of sourceFont')
    assert copy.count('set base style of detachedStyle to style normal')==1
    reset=phases[3]
    assert [x for x in reset if x.startswith('set base style of')]==['set base style of detachedStyle to style normal']
    assert not any('set properties of' in x for x in reset)
    assert phases[2]==phases[4]
    for i in (2,4):
        assert not any('set properties of' in x or 'set base style of' in x for x in phases[i])
        assert any('leftKind' in x and 'rightKind' in x for x in phases[i])


@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax compiler')
def test_order_diagnostic_compiles_without_launch(tmp_path):
    m=module()
    for i,commands in enumerate(m.diagnostic_phases()):
        source='use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n'+'\n'.join(commands)+'\nend tell\n'
        p=tmp_path/('diagnostic-%s.applescript'%i);p.write_text(source)
        r=subprocess.run(['/usr/bin/osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True,timeout=20)
        assert r.returncode==0,r.stderr


def test_diagnostic_mode_guard_and_malformed_observation_stop(tmp_path,monkeypatch):
    m=module()
    monkeypatch.setattr(m,'run_order_diagnostic',lambda *a:pytest.fail('native diagnostic started without execute'))
    assert m.main(['--mode','font-order-diagnostic','--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
    with pytest.raises(AssertionError):m.record_observation({'diagnostic':{}},'immediate_after_copy',[])


def _diagnostic_harness(m,tmp_path,monkeypatch,malformed=False):
    events=[]
    phases=m.diagnostic_phases(); rows=diagnostic_rows(m)
    rows[0]=[rows[0][0],['number',17],['number',12],False]
    class Owner:
        def __init__(self,readonly=False):
            self._closed=False;self._quarantined=False;self.staging_root=tmp_path/'fake-stage'
            self.staging_root.mkdir(exist_ok=True);self.readonly=readonly
        def __enter__(self):return self
        def __exit__(self,*args):self._closed=True;events.append('closed-readonly' if self.readonly else 'closed-owned')
        def _execute(self,commands):
            if commands==phases[2]:
                events.append('readback-readonly' if self.readonly else 'readback')
                return [] if malformed else rows
            if self.readonly:pytest.fail('read-only reopen attempted a mutation')
            if commands==phases[0]:events.append('seed');return [['stage',True]]
            if commands==phases[1]:events.append('copy');return [['stage',True]]
            if commands==phases[3]:events.append('reset');return [['stage',True]]
            if 'set sentinelDoc to make new document' in commands:return [['Sentinel']]
            if commands==['set nativeRows to {{version as text}}']:return [['test']]
            pytest.fail('unexpected native commands')
        def save_docx(self,path):events.append('save-'+path.stem);path.write_bytes(path.name.encode())
        def export_pdf(self,path):path.write_bytes(b'pdf')
    class Sessions:
        @staticmethod
        def new_document(**kwargs):return Owner()
        @staticmethod
        def open_document(path,**kwargs):
            assert kwargs['read_only'] is True
            return Owner(readonly=True)
    monkeypatch.setattr(m,'MacWordSession',Sessions)
    monkeypatch.setattr(m,'retain_sources',lambda *args:{})
    monkeypatch.setattr(m,'inventory',lambda *args:[])
    monkeypatch.setattr(m,'sentinel_preimage',lambda *args:['Sentinel','Sentinel',False,'hash'])
    def cleanup(owner,output,report,name,token):
        assert owner._closed and not owner._quarantined
        report['sentinel_cleanup_attempted']=True;report['checks']['owned_and_sentinel_closed']=True
        events.append('sentinel-closed')
    monkeypatch.setattr(m,'close_after_owned',cleanup)
    monkeypatch.setattr(m,'diagnostic_artifacts',lambda out,report:report.update({'artifact_appearance_observations':{'before-base-reset':{'complete_font_xml_equal':False},'after-base-reset':{'complete_font_xml_equal':False}}}))
    old_sha=m.sha
    monkeypatch.setattr(m,'sha',lambda p:'dictionary-hash' if str(p).endswith('Word.sdef') else old_sha(p))
    return events


def test_diagnostic_mismatch_runs_control_but_never_reports_clone_acceptance(tmp_path,monkeypatch):
    m=module();events=_diagnostic_harness(m,tmp_path,monkeypatch)
    result=m.run_order_diagnostic(tmp_path)
    assert result['status']=='DIAGNOSTIC_COMPLETE'
    assert result['diagnostic_observation_complete'] is True
    assert result['diagnostic']['immediate_after_copy']['appearance_equal'] is False
    assert result['diagnostic']['after_base_reset']['appearance_equal'] is False
    assert events.count('reset')==1
    assert events.index('readback')<events.index('reset')
    assert events.index('closed-owned')<events.index('sentinel-closed')
    assert events.count('readback-readonly')==2
    assert result['checks']['snapshot_bytes_preserved'] is True


def test_malformed_immediate_readback_stops_before_reset_or_save(tmp_path,monkeypatch):
    m=module();events=_diagnostic_harness(m,tmp_path,monkeypatch,malformed=True)
    result=m.run_order_diagnostic(tmp_path)
    assert result['status']=='FAIL'
    assert 'reset' not in events and not any(e.startswith('save-') for e in events)
    assert result['raw_observation_rows']['immediate_after_copy']==[]
    assert events[-2:]==['closed-owned','sentinel-closed']
    assert result['error']['type']=='AssertionError'
