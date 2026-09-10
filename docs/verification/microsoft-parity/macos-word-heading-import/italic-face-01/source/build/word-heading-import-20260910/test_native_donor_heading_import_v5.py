from pathlib import Path
import importlib.util
from copy import deepcopy
import xml.etree.ElementTree as ET
from zipfile import ZipFile
import pytest
PATH=Path(__file__).with_name('native-donor-heading-import-v5.py')
def module():
    assert PATH.exists(),'v5 proofing-aware scalar signatures absent'
    spec=importlib.util.spec_from_file_location('donor_v5',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def doc(m):
    root=ET.Element(m.W+'document');body=ET.SubElement(root,m.W+'body')
    for i in range(1,5):
        p=ET.SubElement(body,m.W+'p')
        if i in (1,4):
            ET.SubElement(ET.SubElement(p,m.W+'pPr'),m.W+'pStyle',{m.W+'val':'H1' if i==1 else 'CLONE'})
            r=ET.SubElement(p,m.W+'r');ET.SubElement(ET.SubElement(r,m.W+'rPr'),m.W+'rFonts',{m.W+'ascii':'Exact Font'})
            ET.SubElement(r,m.W+'t').text=m.FRESH
    return root

def split_and_mark(m,p):
    run=p.find(m.W+'r');p.remove(run);offset=m.FRESH.index('العربية');pieces=[m.FRESH[:offset],'العربية',m.FRESH[offset+7:]]
    for index,text in enumerate(pieces):
        if index==1:ET.SubElement(p,m.W+'proofErr',{m.W+'type':'spellStart'})
        r=deepcopy(run);r.find(m.W+'t').text=text;p.append(r)
        if index==1:ET.SubElement(p,m.W+'proofErr',{m.W+'type':'spellEnd'})

def test_identical_format_split_and_valid_spell_markers_do_not_change_scalar_signature():
    m=module();root=doc(m);before=m.control_signatures(ET.tostring(root),'H1','CLONE')
    split_and_mark(m,root.find(m.W+'body')[3])
    after=m.control_signatures(ET.tostring(root),'H1','CLONE')
    assert before==after and after[0]==after[1]
    assert len(after[0][1])==len(m.FRESH)

@pytest.mark.parametrize('change',['child','text','tail','extra_attribute','unknown_type','unbalanced','nested','unknown_node'])
def test_invalid_proof_metadata_or_unknown_content_rejected(change):
    m=module();root=doc(m);p=root.find(m.W+'body')[3];split_and_mark(m,p)
    mark=p.find(m.W+'proofErr')
    if change=='child':ET.SubElement(mark,m.W+'t').text='hidden'
    elif change=='text':mark.text='hidden'
    elif change=='tail':mark.tail='hidden'
    elif change=='extra_attribute':mark.set('hidden','x')
    elif change=='unknown_type':mark.set(m.W+'type','gramStart')
    elif change=='unbalanced':p.remove(p.findall(m.W+'proofErr')[1])
    elif change=='nested':p.insert(list(p).index(mark),deepcopy(mark))
    else:p.append(ET.Element(m.W+'fldSimple'))
    with pytest.raises(ValueError):m.control_signatures(ET.tostring(root),'H1','CLONE')

def test_changed_glyph_font_is_not_hidden_by_run_split():
    m=module();root=doc(m);p=root.find(m.W+'body')[3];split_and_mark(m,p)
    p.findall(m.W+'r')[1].find('./'+m.W+'rPr/'+m.W+'rFonts').set(m.W+'ascii','Different Font')
    left,right=m.control_signatures(ET.tostring(root),'H1','CLONE');assert left!=right

def test_unknown_run_children_and_duplicate_rpr_rejected():
    m=module()
    for child in (m.W+'fldChar',m.W+'rPr'):
        root=doc(m);ET.SubElement(root.find(m.W+'body')[3].find(m.W+'r'),child)
        with pytest.raises(ValueError):m.control_signatures(ET.tostring(root),'H1','CLONE')

def test_actual_native_proofing_split_preserves_before_challenge_signatures():
    m=module();folder=PATH.parent/'native-donor-import-04'
    assert folder.exists(), 'Required retained native v4 evidence absent'
    def read(name):
        with ZipFile(folder/name) as z:
            styles=ET.fromstring(z.read('word/styles.xml'));source=m.paragraph_style_by_name(styles.findall(m.W+'style'),'heading 1');clone=m.paragraph_style_by_name(styles.findall(m.W+'style'),m.CLONE)
            return m.control_signatures(z.read('word/document.xml'),source.get(m.W+'styleId'),clone.get(m.W+'styleId'))
    assert read('before-challenge.docx')==read('source-challenged.docx')

def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()

@pytest.mark.parametrize('change',['ppr_attribute','ppr_text','pstyle_child','pstyle_text','pstyle_tail'])
def test_complete_paragraph_properties_cannot_hide_content_or_attributes(change):
    m=module();root=doc(m);props=root.find(m.W+'body')[3].find(m.W+'pPr');style=props.find(m.W+'pStyle')
    if change=='ppr_attribute':props.set('unexpected','yes')
    elif change=='ppr_text':props.text='hidden'
    elif change=='pstyle_child':ET.SubElement(style,m.W+'rFonts')
    elif change=='pstyle_text':style.text='hidden'
    else:style.tail='hidden'
    try:left,right=m.control_signatures(ET.tostring(root),'H1','CLONE')
    except ValueError:return
    assert left!=right

def test_unknown_run_attributes_rejected_but_explicit_revision_markers_ignored():
    m=module();root=doc(m);run=root.find(m.W+'body')[3].find(m.W+'r')
    run.set(m.W+'rsidR','00112233');run.set(m.W+'rsidRPr','33445566')
    left,right=m.control_signatures(ET.tostring(root),'H1','CLONE');assert left==right
    run.set('unknown','hidden')
    with pytest.raises(ValueError):m.control_signatures(ET.tostring(root),'H1','CLONE')
