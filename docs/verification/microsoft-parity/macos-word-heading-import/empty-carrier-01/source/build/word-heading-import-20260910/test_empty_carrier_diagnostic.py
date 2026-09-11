from pathlib import Path
import importlib.util
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import pytest
PATH=Path(__file__).with_name('empty-carrier-diagnostic.py')
def module():
    assert PATH.exists(),'empty carrier diagnostic absent'
    spec=importlib.util.spec_from_file_location('empty_carrier',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_existing_empty_carrier_style_precedes_insert_without_resets():
    m=module();body='SOURCE\rIMPORTED\rSUFFIX\r\r';p=len(body.encode('utf-16-le'))//2-1
    code='\n'.join(m.creation_commands('B',body,p))
    assert code.index('set style of carrierRange')<code.index('set content of freshInsert')
    assert 'reset' not in code and 'set content of carrierRange' not in code
    assert 'WPSC_EMPTY_CARRIER_TEXT' in code and 'WPSC_EMPTY_CARRIER_BOUNDS' in code
    assert 'set content of freshInsert to '+m.apple_string(m.FRESH) in code
    assert next(line for line in code.splitlines() if line.startswith('set content of freshInsert'))=='set content of freshInsert to '+m.apple_string(m.FRESH)

def test_pinned_carrier_is_empty_without_properties_or_runs():
    m=module()
    with ZipFile(m.INPUT) as z:xml=z.read('word/document.xml')
    assert m.empty_carrier_xml_valid(xml)
    root=ET.fromstring(xml);p=root.find(m.W+'body').findall(m.W+'p')[3]
    ET.SubElement(p,m.W+'pPr')
    assert not m.empty_carrier_xml_valid(ET.tostring(root))

@pytest.mark.parametrize('change',['run','text','extra_paragraph'])
def test_nonempty_or_ambiguous_carrier_rejected(change):
    m=module()
    with ZipFile(m.INPUT) as z:root=ET.fromstring(z.read('word/document.xml'))
    body=root.find(m.W+'body');p=body.findall(m.W+'p')[3]
    if change=='run':ET.SubElement(p,m.W+'r')
    elif change=='text':ET.SubElement(ET.SubElement(p,m.W+'r'),m.W+'t').text='existing'
    else:ET.SubElement(body,m.W+'p')
    assert not m.empty_carrier_xml_valid(ET.tostring(root))

def test_bound_preimage_rejects_nonterminal_or_wrong_mode():
    m=module()
    with pytest.raises(ValueError):m.creation_commands('A','x\r',1)
    with pytest.raises(ValueError):m.creation_commands('B','x\r',0)

def test_saved_evidence_still_precedes_false_ack():
    m=module();events=[]
    class Owner:
        def save_docx(self,path):events.append('save')
    rows=m.save_after_creation(Owner(),Path('/tmp/owned.docx'),lambda:events.append('mutate') or [['stage',False]])
    assert rows==[['stage',False]] and events==['mutate','save']

def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
