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
