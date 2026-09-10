from pathlib import Path
import importlib.util
import xml.etree.ElementTree as ET
import pytest
PATH=Path(__file__).with_name('fresh-order-diagnostic.py')
def module():
    assert PATH.exists(),'fresh-order diagnostic absent'
    spec=importlib.util.spec_from_file_location('fresh_order',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_orders_use_same_text_style_resets_with_only_insertion_order_changed():
    m=module();pre='SOURCE\rIMPORTED\rSUFFIX\r\r';p=len(pre.encode('utf-16-le'))//2-1
    a='\n'.join(m.creation_commands('A',pre,p));b='\n'.join(m.creation_commands('B',pre,p))
    assert a.index('set content of freshRange to')<a.index('set style of freshRange')<a.index('reset font object')
    assert b.index('set style of freshRange')<b.index('reset font object')<b.index('set content of freshInsert to')
    for code in (a,b):
        assert 'WPSC_FRESH_PREIMAGE' in code and 'WPSC_FRESH_TERMINAL' in code
        assert 'reset paragraph format of freshRange' in code and m.apple_string(m.FRESH) in code

def test_fresh_anchors_use_common_latin_e_and_whole_emoji():
    m=module();body=m.FRESH+'\r';anchors=m.range_anchors(body,0,len(body.encode('utf-16-le'))//2,m.FRESH)
    assert anchors[0]['scalar']=='E' and anchors[-1]['scalar']=='😀'
    assert anchors[-1]['end']-anchors[-1]['start']==2

@pytest.mark.parametrize('change',['wrong_text','surrogate','bool'])
def test_fresh_anchor_bad_preimage_rejected(change):
    m=module();body=m.FRESH+'\r';start=0;end=len(body.encode('utf-16-le'))//2
    if change=='wrong_text':body='wrong\r'
    elif change=='surrogate':start=end-2
    else:start=True
    with pytest.raises(ValueError):m.range_anchors(body,start,end,m.FRESH)

def test_xml_observation_records_direct_font_overrides_as_data():
    m=module();w=m.W
    doc=ET.Element(w+'document');p=ET.SubElement(ET.SubElement(doc,w+'body'),w+'p')
    ET.SubElement(ET.SubElement(p,w+'pPr'),w+'pStyle',{w+'val':'clone'})
    r=ET.SubElement(p,w+'r');ET.SubElement(ET.SubElement(r,w+'rPr'),w+'rFonts',{w+'cs':'Times New Roman'})
    ET.SubElement(r,w+'t').text=m.FRESH
    result=m.paragraph_xml_observation(ET.tostring(doc),m.FRESH)
    assert result['unique'] and 'Times New Roman' in result['runs'][0]['rPr_xml']

def test_save_evidence_precedes_diagnostic_readback_and_assertions():
    m=module();events=[]
    class Owner:
        def save_docx(self,path):events.append(('save',path))
    m.save_after_creation(Owner(),Path('/tmp/owned-evidence.docx'),lambda:events.append(('mutation',)) or [['stage',False]])
    assert [e[0] for e in events]==['mutation','save']

def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
