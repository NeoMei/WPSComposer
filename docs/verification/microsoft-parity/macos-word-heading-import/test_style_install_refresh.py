from pathlib import Path
import importlib.util
from copy import deepcopy
import xml.etree.ElementTree as E
import pytest

def m():
 p=Path(__file__).with_name('style-install-refresh.py');s=importlib.util.spec_from_file_location('install_probe',p);v=importlib.util.module_from_spec(s);s.loader.exec_module(v);return v

def seeded(v):
 styles,_=v.base.package(v.INPUT);root=E.fromstring(styles);source=v.v5.paragraph_style_by_name(root.findall(v.W+'style'),'heading 1')
 parent=E.SubElement(root,v.W+'style',{v.W+'type':'paragraph',v.W+'styleId':'Parent'})
 E.SubElement(parent,v.W+'name',{v.W+'val':v.PARENT});E.SubElement(parent,v.W+'basedOn',{v.W+'val':'a'})
 E.SubElement(E.SubElement(parent,v.W+'pPr'),v.W+'ind',{v.W+'left':'340'})
 source.find(v.W+'basedOn').set(v.W+'val','Parent');return root

def test_clone_materializes_exact_one_inherited_indent():
 v=m();root=seeded(v);clone=v.expected_clone(E.tostring(root))
 assert clone.find('./'+v.W+'pPr/'+v.W+'ind').attrib=={v.W+'left':'340'}
 assert clone.find(v.W+'basedOn').attrib=={v.W+'val':'a'}
 assert clone.find(v.W+'link') is None

@pytest.mark.parametrize('kind',['source_indent','bad_parent','inherited_font'])
def test_unsupported_inheritance_rejected(kind):
 v=m();root=seeded(v);source=v.v5.paragraph_style_by_name(root.findall(v.W+'style'),'heading 1');parent=root[-1]
 if kind=='source_indent':E.SubElement(source.find(v.W+'pPr'),v.W+'ind',{v.W+'left':'2'})
 elif kind=='bad_parent':parent.find(v.W+'basedOn').set(v.W+'val','unknown')
 else:E.SubElement(E.SubElement(parent,v.W+'rPr'),v.W+'b')
 with pytest.raises(ValueError):v.expected_clone(E.tostring(root))

def test_only_target_style_add_or_refresh_allowed():
 v=m();root=seeded(v);before=E.tostring(root);clone=v.expected_clone(before);root.append(clone)
 assert v.other_styles_preserved(before,E.tostring(root))
 root.set('unexpected','x');assert not v.other_styles_preserved(before,E.tostring(root))

def test_same_name_stale_clone_does_not_count_as_refresh():
 v=m();root=seeded(v);a=v.expected_clone(E.tostring(root));root.append(deepcopy(a));source=v.v5.paragraph_style_by_name(root.findall(v.W+'style'),'heading 1');source.find('./'+v.W+'rPr/'+v.W+'sz').set(v.W+'val','46')
 assert not v.clone_matches(E.tostring(root),v.expected_clone(E.tostring(root)))

def test_clear_requires_exact_utf16_carrier():
 v=m();body='A😀\rB\r';start=len('A😀\r'.encode('utf-16-le'))//2;inserted=body[:3]+v.CARRIER+'\r'+body[3:]
 assert 'WPSC_CARRIER_TEXT' in '\n'.join(v.clear_commands(inserted,start,start+v.units(v.CARRIER+'\r')))
 with pytest.raises(ValueError):v.clear_commands(inserted,start+1,start+v.units(v.CARRIER+'\r'))

def test_input_package_is_valid_office_input_and_keeps_theme(tmp_path):
 from zipfile import ZipFile
 v=m();root=seeded(v)
 with ZipFile(v.INPUT) as z:theme=z.read('word/theme/theme1.xml')
 path=tmp_path/'input.docx';v.input_fragment(E.tostring(root),theme,path)
 from skills.WPSComposer.scripts.msoffice.input_validation import validate_native_input
 validate_native_input(path,'writer')
 with ZipFile(path) as z:
  assert z.read('word/theme/theme1.xml')==theme
  assert v.clone_matches(z.read('word/styles.xml'),v.expected_clone(E.tostring(root)))
  assert len(E.fromstring(z.read('word/document.xml')).find(v.W+'body'))==1

def test_renamed_clone_and_duplicate_identity_failclosed():
 v=m();root=seeded(v);before=E.tostring(root);clone=v.expected_clone(before);clone.set(v.W+'styleId','Renamed');root.append(clone)
 assert not v.other_styles_preserved(before,E.tostring(root))
 root[-1].set(v.W+'styleId',v.STYLE_ID);root.append(deepcopy(root[-1]));assert not v.other_styles_preserved(before,E.tostring(root))

def test_state_requires_exact_selection_bookmark_and_boolean():
 v=m();row=['state',True,'body',0,5,0,v.units(v.v5.FRESH),v.v5.FRESH,0,0,17,17]
 assert v.state_valid([row],'body',17)
 for index,value in [(1,1),(3,1),(6,1),(7,'wrong'),(10,True)]:
  bad=list(row);bad[index]=value;assert not v.state_valid([bad],'body',17)

def test_setup_and_readback_use_verified_paragraph_indent_property():
 v=m()
 assert 'set paragraph format left indent of paragraph format of parentStyle to 17' in v.setup_commands()
 assert 'paragraph format left indent of paragraph format of sourceStyle' in '\n'.join(v.state_commands())
 assert not any(line.startswith('set left indent of paragraph format') for line in v.setup_commands())
