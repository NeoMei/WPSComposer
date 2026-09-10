from pathlib import Path
import importlib.util
from copy import deepcopy
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import pytest
PATH=Path(__file__).with_name('native-donor-heading-import-v4.py')
def module():
    assert PATH.exists(),'v4 reciprocal linked-style oracle absent'
    spec=importlib.util.spec_from_file_location('donor_v4',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def before_xml():
    with ZipFile(PATH.parent/'native-donor-import-03/before-challenge.docx') as z:return z.read('word/styles.xml')

def changed(m,root):
    after=deepcopy(root)
    for style in after.findall(m.W+'style'):
        if style.get(m.W+'styleId') in ('1','10'):
            size=style.find('./'+m.W+'rPr/'+m.W+'sz');size.set(m.W+'val',str(int(size.get(m.W+'val'))+22))
        if style.get(m.W+'styleId')=='1':
            spacing=style.find('./'+m.W+'pPr/'+m.W+'spacing');spacing.set(m.W+'before',str(int(spacing.get(m.W+'before'))+140))
    return after

def test_actual_native_reciprocal_pair_three_attribute_change():
    m=module()
    with ZipFile(PATH.parent/'native-donor-import-03/source-challenged.docx') as z:after=z.read('word/styles.xml')
    assert m.source_only_challenge_xml(before_xml(),after)

@pytest.mark.parametrize('change',['missing_link','oneway','duplicate_id','wrong_type','wrong_name','unequal_size','duplicate_size'])
def test_pair_precondition_rejects_unverified_relationship(change):
    m=module();root=ET.fromstring(before_xml());styles=root.findall(m.W+'style')
    source=next(s for s in styles if s.get(m.W+'styleId')=='1');pair=next(s for s in styles if s.get(m.W+'styleId')=='10')
    if change=='missing_link':source.remove(source.find(m.W+'link'))
    elif change=='oneway':pair.find(m.W+'link').set(m.W+'val','other')
    elif change=='duplicate_id':root.append(deepcopy(pair))
    elif change=='wrong_type':pair.set(m.W+'type','paragraph')
    elif change=='wrong_name':pair.find(m.W+'name').set(m.W+'val','Unrelated Character')
    elif change=='unequal_size':pair.find('./'+m.W+'rPr/'+m.W+'sz').set(m.W+'val','33')
    else:pair.find(m.W+'rPr').append(deepcopy(pair.find('./'+m.W+'rPr/'+m.W+'sz')))
    assert not m.source_only_challenge_xml(ET.tostring(root),ET.tostring(changed(m,root)))

@pytest.mark.parametrize('change',['docdefaults','root','source_szCs','pair_szCs','clone','spacing_after'])
def test_any_fourth_change_rejected(change):
    m=module();root=changed(m,ET.fromstring(before_xml()));styles=root.findall(m.W+'style')
    source=next(s for s in styles if s.get(m.W+'styleId')=='1');pair=next(s for s in styles if s.get(m.W+'styleId')=='10')
    if change=='docdefaults':root.find('.//'+m.W+'docDefaults//'+m.W+'szCs').set(m.W+'val','99')
    elif change=='root':root.set('unexpected','yes')
    elif change=='source_szCs':source.find('./'+m.W+'rPr/'+m.W+'szCs').set(m.W+'val','99')
    elif change=='pair_szCs':pair.find('./'+m.W+'rPr/'+m.W+'szCs').set(m.W+'val','99')
    elif change=='clone':m.paragraph_style_by_name(styles,m.CLONE).find('./'+m.W+'rPr/'+m.W+'sz').set(m.W+'val','99')
    else:source.find('./'+m.W+'pPr/'+m.W+'spacing').set(m.W+'after','999')
    assert not m.source_only_challenge_xml(before_xml(),ET.tostring(root))

def test_restoration_keeps_full_baseline_tree_requirement():
    m=module();before=before_xml()
    assert m.complete_styles_equal(before,before)
    assert not m.complete_styles_equal(before,ET.tostring(changed(m,ET.fromstring(before))))

def test_no_execute(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
