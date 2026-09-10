from pathlib import Path
import importlib.util
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import pytest

PATH = Path(__file__).with_name('docx-package-materialization.py')

def module():
    assert PATH.exists(), 'DOCX package materialization diagnostic absent'
    spec = importlib.util.spec_from_file_location('docx_materialization', PATH)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m

def test_four_parts_preserved_and_content_types_relationships(tmp_path):
    m=module(); target=tmp_path/'input.docx'
    m.package_input(m.FRAGMENT.read_bytes(), target)
    assert m.input_valid(target,m.FRAGMENT.read_bytes())
    assert m.package_valid(target,m.SOURCE_STYLES.read_bytes())
    with ZipFile(target) as z:
        assert set(z.namelist())==set(m.PARTS)|{'[Content_Types].xml'}
        types=ET.fromstring(z.read('[Content_Types].xml'))
        assert {x.get('PartName'):x.get('ContentType') for x in types}=={'/'+k:v for k,v in m.PARTS.items()}

@pytest.mark.parametrize('change',['missing','extra','duplicate','external'])
def test_input_rejects_altered_part_graph(tmp_path,change):
    m=module(); root=ET.fromstring(m.FRAGMENT.read_bytes())
    if change=='missing':root.remove(root[-1])
    elif change=='extra':
        p=ET.SubElement(root,m.PKG+'part',{m.PKG+'name':'/word/vbaProject.bin',m.PKG+'contentType':'active'})
    elif change=='duplicate':root.append(root[0])
    else:root[0].find(m.PKG+'xmlData')[0][0].set('TargetMode','External')
    with pytest.raises(ValueError):m.package_input(ET.tostring(root),tmp_path/'bad.docx')
    assert not (tmp_path/'bad.docx').exists()

@pytest.mark.parametrize('change',['szCs','indent','style','direct','body'])
def test_saved_oracle_rejects_property_style_and_body_drift(tmp_path,change):
    m=module(); source=m.SOURCE_STYLES.read_bytes(); input_path=tmp_path/'input.docx'
    m.package_input(m.FRAGMENT.read_bytes(),input_path)
    with ZipFile(input_path) as z:parts={k:z.read(k) for k in z.namelist()}
    styles=ET.fromstring(parts['word/styles.xml']); doc=ET.fromstring(parts['word/document.xml'])
    clone=m.paragraph_style_by_name(styles.findall(m.W+'style'),m.CLONE)
    if change=='szCs':clone.find(m.W+'rPr').remove(clone.find(m.W+'rPr').find(m.W+'szCs'))
    elif change=='indent':ET.SubElement(clone.find(m.W+'pPr'),m.W+'ind',{m.W+'left':'123'})
    elif change=='style':doc.find('.//'+m.W+'pStyle').set(m.W+'val','a')
    elif change=='direct':ET.SubElement(ET.SubElement(doc.find('.//'+m.W+'r'),m.W+'rPr'),m.W+'szCs',{m.W+'val':'24'})
    else:doc.find('.//'+m.W+'t').text='changed'
    parts['word/styles.xml']=ET.tostring(styles);parts['word/document.xml']=ET.tostring(doc)
    altered=tmp_path/'altered.docx'
    with ZipFile(altered,'w') as z:
        for name,data in parts.items():z.writestr(name,data)
    assert not m.package_valid(altered,source)

def test_input_comparison_rejects_changed_complete_part(tmp_path):
    m=module(); target=tmp_path/'input.docx'; m.package_input(m.FRAGMENT.read_bytes(),target)
    root=ET.fromstring(m.FRAGMENT.read_bytes());root[1].find(m.PKG+'xmlData')[0].find('.//'+m.W+'szCs').set(m.W+'val','99')
    assert not m.input_valid(target,ET.tostring(root))

def test_no_execute_never_enters_native(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *args:pytest.fail('native entered'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()

def test_unchanged_copy_cannot_count_as_materialization(tmp_path):
    m=module(); source=tmp_path/'input.docx'; output=tmp_path/'saved.docx'
    m.package_input(m.FRAGMENT.read_bytes(),source);output.write_bytes(source.read_bytes())
    assert not m.native_serialized_valid(source,output)
