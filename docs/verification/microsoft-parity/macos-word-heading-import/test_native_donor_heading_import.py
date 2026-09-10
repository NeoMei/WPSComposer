from pathlib import Path
import importlib.util
from copy import deepcopy
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import pytest
PATH=Path(__file__).with_name('native-donor-heading-import.py')

def module():
    assert PATH.exists(), 'native donor heading probe absent'
    spec=importlib.util.spec_from_file_location('native_donor_import',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def styles(m):
    before=m.SOURCE_STYLES.read_bytes();donor=m.package_styles(m.DONOR)
    root=ET.fromstring(before)
    clone=m.paragraph_style_by_name(ET.fromstring(donor).findall(m.W+'style'),m.CLONE)
    root.append(deepcopy(clone))
    return before,donor,root

def test_exact_recipient_source_and_only_complete_clone_addition():
    m=module();before,donor,after=styles(m)
    assert m.recipient_matches_donor(before,donor)
    assert m.only_clone_added(before,ET.tostring(after),donor)

@pytest.mark.parametrize('change',['source_size','clone_szCs','unexpected_style','old_style','clone_link','clone_missing'])
def test_reject_style_change_before_target_deletion(change):
    m=module();before,donor,after=styles(m)
    clone=m.paragraph_style_by_name(after.findall(m.W+'style'),m.CLONE)
    if change=='source_size':
        source=m.paragraph_style_by_name(after.findall(m.W+'style'),'heading 1');source.find(m.W+'rPr').find(m.W+'sz').set(m.W+'val','11')
    elif change=='clone_szCs':clone.find(m.W+'rPr').remove(clone.find(m.W+'rPr').find(m.W+'szCs'))
    elif change=='unexpected_style':ET.SubElement(after,m.W+'style',{m.W+'styleId':'Extra'})
    elif change=='old_style':after.findall(m.W+'style')[0].set('altered','yes')
    elif change=='clone_link':ET.SubElement(clone,m.W+'link',{m.W+'val':'Heading1'})
    else:after.remove(clone)
    assert not m.only_clone_added(before,ET.tostring(after),donor)

def test_recipient_source_mismatch_or_existing_clone_rejected():
    m=module();before,donor,after=styles(m)
    assert not m.recipient_matches_donor(ET.tostring(after),donor)
    root=ET.fromstring(before)
    m.paragraph_style_by_name(root.findall(m.W+'style'),'heading 1').find(m.W+'rPr').find(m.W+'szCs').set(m.W+'val','24')
    assert not m.recipient_matches_donor(ET.tostring(root),donor)

def test_astral_preimage_replacement_and_wrong_ack_fail():
    m=module();before='😀\rREPLACE\r尾\r\r'
    inserted,final,start,end=m.collapsed.expected_stages(before,3,11,m.IMPORTED+'\r')
    assert final=='😀\r'+m.IMPORTED+'\r尾\r\r'
    assert m.collapsed.exact_ack([['insert',3,3,inserted]],'insert',3,3,inserted)
    assert not m.collapsed.exact_ack([['insert',3,11,inserted]],'insert',3,3,inserted)
    assert 'WPSC_COLLAPSED_FULL_PREIMAGE' in '\n'.join(m.collapsed.clear_commands(inserted,start,end))

@pytest.mark.parametrize('change',['wrong_style','direct_szCs','direct_indent','duplicate'])
def test_imported_paragraph_must_retain_unique_style_only_body(tmp_path,change):
    m=module()
    with ZipFile(m.DONOR) as z:parts={name:z.read(name) for name in z.namelist()}
    assert m.imported_paragraph_valid(parts['word/document.xml'],parts['word/styles.xml'])
    doc=ET.fromstring(parts['word/document.xml']);p=doc.find('.//'+m.W+'p')
    if change=='wrong_style':p.find('./'+m.W+'pPr/'+m.W+'pStyle').set(m.W+'val','Heading1')
    elif change=='direct_szCs':ET.SubElement(ET.SubElement(p.find(m.W+'r'),m.W+'rPr'),m.W+'szCs',{m.W+'val':'24'})
    elif change=='direct_indent':ET.SubElement(p.find(m.W+'pPr'),m.W+'ind',{m.W+'left':'40'})
    else:doc.find(m.W+'body').append(deepcopy(p))
    assert not m.imported_paragraph_valid(ET.tostring(doc),parts['word/styles.xml'])

def test_predelete_guard_requires_text_inventory_style_and_body():
    m=module();before,donor,after=styles(m);after_xml=ET.tostring(after)
    with ZipFile(m.DONOR) as z:doc=z.read('word/document.xml')
    inventory=[['own','/private/own.docx','body',True],['sentinel','', 'token',False]]
    args=[[['insert',3,3,'EXPECTED']],3,'EXPECTED',inventory,inventory,'/private/own.docx',before,after_xml,donor,doc]
    assert m.deletion_gate(*args)
    altered=list(args);altered[0]=[['insert',3,3,'WRONG']];assert not m.deletion_gate(*altered)
    altered=list(args);altered[4]=inventory[:1];assert not m.deletion_gate(*altered)
    altered=list(args);altered[7]=before;assert not m.deletion_gate(*altered)
    changed=ET.fromstring(doc);changed.find('.//'+m.W+'t').text='CHANGED'
    altered=list(args);altered[9]=ET.tostring(changed);assert not m.deletion_gate(*altered)


def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()

@pytest.mark.parametrize('change',['paragraph','defaults','normal'])
def test_recipient_inheritance_and_paragraph_dependencies_must_match(change):
    m=module();before,donor,_=styles(m);root=ET.fromstring(before)
    if change=='paragraph':
        source=m.paragraph_style_by_name(root.findall(m.W+'style'),'heading 1')
        ET.SubElement(source.find(m.W+'pPr'),m.W+'ind',{m.W+'left':'1'})
    elif change=='defaults':root.find('.//'+m.W+'docDefaults//'+m.W+'szCs').set(m.W+'val','25')
    else:m.paragraph_style_by_name(root.findall(m.W+'style'),'Normal').set('changed','yes')
    assert not m.recipient_matches_donor(ET.tostring(root),donor)
