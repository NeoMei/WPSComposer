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
    xml=('<w:styles xmlns:w="'+w[1:-1]+'"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'+r+p+'</w:style><w:style w:type="paragraph" w:styleId="Detached"><w:name w:val="WPSC Heading Clone"/><w:basedOn w:val="Normal"/>'+r+p+'</w:style></w:styles>').encode()
    assert m.clone_xml_valid(xml,'detached-properties')
    for a,b in [('<w:shd w:fill="FCE8E6"/>',''),('<w:bdr w:val="single" w:sz="6"/>',''),('<w:tab w:val="right" w:pos="800"/>','')]:
        first,second=xml.decode().split('<w:style w:type="paragraph" w:styleId="Detached">')
        assert not m.clone_xml_valid((first+'<w:style w:type="paragraph" w:styleId="Detached">'+second.replace(a,b)).encode(),'detached-properties')


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
    styles='<w:styles xmlns:w="'+ns+'">'+''.join('<w:style w:type="paragraph" w:styleId="Heading%s"><w:name w:val="heading %s"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>'%(i,i) for i in range(1,5))+'</w:styles>'
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


def test_scalar_copy_covers_every_font_dimension_and_only_writes_differences():
    m=module();commands=m.scalar_copy_commands();source='\n'.join(commands)
    for label,left,right in m.dimensions('font-properties'):
        assert 'set scalarSourceValue to '+left in commands
        assert 'set scalarDestinationValue to '+right in commands
        assert 'set '+right+' to scalarSourceValue' in commands
    assert source.count('if not scalarEqual then')==len(m.dimensions('font-properties'))
    assert 'set properties of' not in source and 'on error' not in source
    assert source.index('set base style of detachedStyle to style normal')<source.index('set scalarSourceValue')
    assert 'isEqualToString:' in source


def test_scalar_copy_checks_entire_source_preimage_and_complete_target():
    m=module();before=diagnostic_rows(m);after=diagnostic_rows(m)
    assert m.scalar_copy_verdict(before,after)=={'source_properties_preserved':True,'complete_font_appearance_equal':True}
    changed=[r[:] for r in after]
    changed[0]=[after[0][0],['number',18],['number',18],True]
    assert m.scalar_copy_verdict(before,changed)=={'source_properties_preserved':False,'complete_font_appearance_equal':True}
    changed[0]=[after[0][0],['number',17],['number',12],False]
    assert m.scalar_copy_verdict(before,changed)=={'source_properties_preserved':True,'complete_font_appearance_equal':False}
    with pytest.raises(AssertionError):m.scalar_copy_verdict(before,[])


@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax compiler')
def test_scalar_copy_compiles_without_launch(tmp_path):
    m=module()
    source='use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n'+'\n'.join(m.scalar_copy_commands())+'\nend tell\n'
    p=tmp_path/'scalar.applescript';p.write_text(source)
    r=subprocess.run(['/usr/bin/osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True,timeout=20)
    assert r.returncode==0,r.stderr


def test_scalar_mode_has_no_execute_side_effect(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('scalar mode started without execute'))
    assert m.main(['--mode','font-scalar-copy','--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()


def test_scalar_copy_cannot_hide_source_changes_in_any_nested_property():
    m=module();before=diagnostic_rows(m)
    for index in range(len(before)):
        after=[r[:] for r in before]
        after[index]=[before[index][0],['text','changed source'],['text','changed source'],True]
        assert m.scalar_copy_verdict(before,after)['source_properties_preserved'] is False


@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax compiler')
def test_all_scalar_candidate_phases_compile_without_launch(tmp_path):
    m=module()
    for i,commands in enumerate(m.phases('font-scalar-copy','WPSCHeading_'+'a'*32)):
        source='use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n'+'\n'.join(commands)+'\nend tell\n'
        p=tmp_path/('scalar-phase-%s.applescript'%i);p.write_text(source)
        r=subprocess.run(['/usr/bin/osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True,timeout=20)
        assert r.returncode==0,r.stderr


def test_scalar_acceptance_rejects_same_appearance_with_source_mutation(tmp_path,monkeypatch):
    m=module();before=diagnostic_rows(m);after=diagnostic_rows(m)
    after[-2]=[after[-2][0],['number',99],['number',99],True]
    phases=m.phases('font-scalar-copy','WPSCHeading_'+'a'*32);events=[]
    class Owner:
        def __init__(self):
            self._closed=False;self._quarantined=False;self.staging_root=tmp_path/'fake-stage'
            self.staging_root.mkdir();self.reads=0
        def __enter__(self):return self
        def __exit__(self,*args):self._closed=True;events.append('owned-closed')
        def _execute(self,commands):
            if commands==phases[1]:self.reads+=1;return before if self.reads==1 else after
            if commands in (phases[0],phases[2]):return [['stage',True]]
            if 'set sentinelDoc to make new document' in commands:return [['Sentinel']]
            if commands==['set nativeRows to {{version as text}}']:return [['test']]
            pytest.fail('Unexpected mutation')
        def save_docx(self,path):pytest.fail('Saved after source mutation')
    class Sessions:
        @staticmethod
        def new_document(**kwargs):return Owner()
    monkeypatch.setattr(m,'MacWordSession',Sessions)
    monkeypatch.setattr(m,'retain_sources',lambda *a:{})
    monkeypatch.setattr(m,'inventory',lambda *a:[])
    monkeypatch.setattr(m,'sentinel_preimage',lambda *a:['Sentinel','Sentinel',False,'hash'])
    def cleanup(owner,output,report,name,token):
        assert owner._closed
        report['sentinel_cleanup_attempted']=True;events.append('sentinel-closed')
    monkeypatch.setattr(m,'close_after_owned',cleanup)
    old_sha=m.sha;monkeypatch.setattr(m,'sha',lambda p:'dictionary-hash' if str(p).endswith('Word.sdef') else old_sha(p))
    result=m.run(tmp_path,'font-scalar-copy')
    assert result['status']=='FAIL'
    assert result['checks']['source_properties_preserved'] is False
    assert result['checks']['complete_font_appearance_equal'] is True
    assert events==['owned-closed','sentinel-closed']


def numeric_style_xml(clone_szcs=True):
    m=module();ns=m.W[1:-1]
    return ('<w:styles xmlns:w="'+ns+'"><w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:rPr><w:sz w:val="34"/><w:szCs w:val="48"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="本地克隆"><w:name w:val="WPSC Heading Clone"/><w:basedOn w:val="a"/><w:rPr><w:sz w:val="34"/>'
        +('<w:szCs w:val="48"/>' if clone_szcs else '')+'</w:rPr></w:style></w:styles>').encode()


def test_numeric_style_ids_resolve_by_unique_paragraph_name_preserving_szcs():
    m=module()
    assert m.clone_xml_valid(numeric_style_xml(),'font-scalar-copy')
    assert not m.clone_xml_valid(numeric_style_xml(False),'font-scalar-copy')
    xml=numeric_style_xml().decode()
    for old,new in [('<w:basedOn w:val="a"/><w:rPr><w:sz w:val="34"/><w:szCs w:val="48"/></w:rPr></w:style></w:styles>','<w:basedOn w:val="1"/><w:rPr><w:sz w:val="34"/><w:szCs w:val="48"/></w:rPr></w:style></w:styles>'),('w:name w:val="heading 1"','w:name w:val="Other"'),('w:type="paragraph" w:styleId="1"','w:type="character" w:styleId="1"')]:
        assert not m.clone_xml_valid(xml.replace(old,new).encode(),'font-scalar-copy')
    duplicate='<w:style w:type="paragraph" w:styleId="duplicate"><w:name w:val="heading 1"/></w:style>'
    assert not m.clone_xml_valid(xml.replace('</w:styles>',duplicate+'</w:styles>').encode(),'font-scalar-copy')


def test_numeric_heading_ids_must_link_to_actual_numbering_styles():
    m=module();ns=m.W[1:-1]
    styles='<w:styles xmlns:w="'+ns+'">'+''.join('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="heading %s"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>'%(i,i) for i in range(1,5))+'</w:styles>'
    levels=''.join('<w:lvl w:ilvl="%s"><w:pStyle w:val="%s"/><w:lvlText w:val="%s"/><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>'%(i-1,i,'.'.join('%%%s'%j for j in range(1,i+1))) for i in range(1,5))
    numbering='<w:numbering xmlns:w="'+ns+'"><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num><w:abstractNum w:abstractNumId="3">'+levels+'</w:abstractNum></w:numbering>'
    assert m.list_xml_valid(styles.encode(),numbering.encode())
    assert not m.list_xml_valid(styles.encode(),numbering.replace('w:pStyle w:val="4"','w:pStyle w:val="Heading4"').encode())


def test_formatted_transfer_guards_detached_identity_and_redefines_private_style_only():
    m=module();source='\n'.join(m.formatted_transfer_commands())
    assert 'set formatted text of targetTextRange to formatted text of sourceTextRange' in source
    assert source.count('end of content of text object of paragraph')>=2
    assert source.count('WPSC_FORMATTED_STYLE_IDENTITY')>=2
    assert 'set automatically update of detachedStyle to true' in source
    assert 'set automatically update of detachedStyle to false' in source
    assert 'set style of selection of boundWindow to detachedStyle' in source
    for forbidden in ('copy format','paste format','update styles','do Visual Basic','set properties of','clipboard'):
        assert forbidden not in source
    assert 'set paragraph format of detachedStyle to paragraph format of sourceStyle' in source


@pytest.mark.skipif(sys.platform!='darwin',reason='Word dictionary syntax compiler')
def test_formatted_diagnostic_phases_compile_without_launch(tmp_path):
    m=module()
    for i,commands in enumerate(m.formatted_phases()):
        source='use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n'+'\n'.join(commands)+'\nend tell\n'
        p=tmp_path/('formatted-%s.applescript'%i);p.write_text(source)
        r=subprocess.run(['/usr/bin/osacompile','-o',str(p.with_suffix('.scpt')),str(p)],capture_output=True,text=True,timeout=20)
        assert r.returncode==0,r.stderr


def test_formatted_mode_has_no_execute_side_effect(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run_formatted_candidate',lambda *a:pytest.fail('native without execute'))
    assert m.main(['--mode','formatted-style-copy','--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()


def test_full_style_snapshot_rejects_changed_complex_size_and_wrong_identity():
    m=module();xml=numeric_style_xml()
    source=m.style_definition(xml,'heading 1')
    assert source is not None
    assert source!=m.style_definition(xml.replace(b'w:val="48"',b'w:val="24"'),'heading 1')
    assert m.style_definition(xml.replace(b'w:val="heading 1"',b'w:val="wrong"'),'heading 1') is None


def test_formatted_fresh_paragraph_has_no_copy_and_no_carried_direct_formatting():
    m=module();commands=m.formatted_fresh_paragraph();source='\n'.join(commands)
    assert 'set style of newRange to detachedStyle' in commands
    assert 'reset font object of newRange' in commands and 'reset paragraph format of newRange' in commands
    assert 'formatted text' not in source
    assert 'paragraph 3 of boundDoc' in source
    source='\n'.join(m.formatted_fresh_readback())
    assert 'set sourceFont to font object of text object of paragraph 3 of boundDoc' in source


def test_style_definition_ignores_only_revision_marker_not_inherited_font_size():
    m=module();xml=numeric_style_xml()
    with_revision=xml.replace(b'<w:name w:val="heading 1"/>',b'<w:name w:val="heading 1"/><w:rsid w:val="ABC123"/>')
    assert m.style_definition(xml,'heading 1')==m.style_definition(with_revision,'heading 1')
    with_link=xml.replace(b'<w:name w:val="heading 1"/>',b'<w:name w:val="heading 1"/><w:link w:val="OtherStyle"/>')
    assert m.style_definition(xml,'heading 1')!=m.style_definition(with_link,'heading 1')
    with_base=xml.replace(b'<w:basedOn w:val="a"/>',b'<w:basedOn w:val="OtherBase"/>',1)
    assert m.style_definition(xml,'heading 1')!=m.style_definition(with_base,'heading 1')


def test_source_challenge_is_exact_two_scalar_properties_without_global_updates():
    m=module();commands=m.formatted_source_change(17,14)
    assert commands==['set sourceStyle to Word style (style heading1) of boundDoc',
        'set font size of font object of sourceStyle to 17','set space before of paragraph format of sourceStyle to 14',
        'set nativeRows to {{"stage",true}}']
    for bad in (float('nan'),float('inf'),'17; unsafe'):
        with pytest.raises((ValueError,TypeError)):m.formatted_source_change(bad,14)
