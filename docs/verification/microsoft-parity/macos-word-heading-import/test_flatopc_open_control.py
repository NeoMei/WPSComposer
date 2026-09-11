from pathlib import Path
import importlib.util
import xml.etree.ElementTree as ET
import pytest

PATH = Path(__file__).with_name('flatopc-open-control.py')

def module():
    assert PATH.exists(), 'explicit converter control absent'
    spec = importlib.util.spec_from_file_location('flatopc_control', PATH)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m

def test_exact_original_native_properties_in_distinct_imported_style():
    m = module()
    source = m.SOURCE_STYLES.read_bytes()
    package = ET.fromstring(m.FRAGMENT.read_bytes())
    imported = next(p for p in package if p.get(m.PKG+'name') == '/word/styles.xml').find(m.PKG+'xmlData')[0]
    xml = ET.tostring(imported)
    assert m.materialized_style_valid(source, xml)
    clone = m.paragraph_style_by_name(imported.findall(m.W+'style'), m.CLONE)
    clone.find(m.W+'rPr').remove(clone.find(m.W+'rPr').find(m.W+'szCs'))
    assert not m.materialized_style_valid(source, ET.tostring(imported))

def test_style_oracle_rejects_changed_base_or_source_identity():
    m = module(); source = m.SOURCE_STYLES.read_bytes()
    package = ET.fromstring(m.FRAGMENT.read_bytes())
    styles = next(p for p in package if p.get(m.PKG+'name') == '/word/styles.xml').find(m.PKG+'xmlData')[0]
    clone = m.paragraph_style_by_name(styles.findall(m.W+'style'), m.CLONE)
    clone.find(m.W+'basedOn').set(m.W+'val', 'Heading1')
    assert not m.materialized_style_valid(source, ET.tostring(styles))

def test_binding_requires_single_exact_private_path_and_expected_name(tmp_path):
    m = module(); p = tmp_path/'own.xml'
    assert m.binding_valid([['binding',0,str(p),p.name]],p)
    assert not m.binding_valid([['binding',0,str(p),'other']],p)
    assert not m.binding_valid([['binding',0,str(p),p.name]]*2,p)
    assert not m.binding_valid([['binding',0,str(p)+'x',p.name]],p)

def test_explicit_converter_not_default_global_change(tmp_path):
    m = module(); p = tmp_path/'own.xml'
    code = '\n'.join(m.open_commands(p))
    assert 'file converter open format xmldocument serialized' in code
    assert 'if (count documents) is not 0' in code
    assert 'default open format' not in code
    assert 'active document' not in code

def test_no_execute_guard(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m,'run',lambda *args: pytest.fail('native entered'))
    assert m.main(['--output',str(tmp_path/'absent')]) == 2
    assert not (tmp_path/'absent').exists()
