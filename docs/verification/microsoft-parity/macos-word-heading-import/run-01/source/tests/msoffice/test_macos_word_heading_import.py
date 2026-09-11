"""Source-only contracts for an opt-in native Flat OPC style import probe."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / 'fixtures/microsoft_parity/macos_word_heading_import.py'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PKG = '{http://schemas.microsoft.com/office/2006/xmlPackage}'


def module():
    assert FIXTURE.exists(), 'native import hypothesis fixture absent'
    spec = importlib.util.spec_from_file_location('heading_import_probe', FIXTURE)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def native_styles():
    return ('<w:styles xmlns:w="' + W[1:-1] + '">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/>'
            '<w:basedOn w:val="a"/><w:link w:val="Heading1Char"/>'
            '<w:pPr><w:keepNext/><w:spacing w:before="280"/><w:outlineLvl w:val="0"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="34"/><w:szCs w:val="48"/><w:shd w:fill="FCE8E6"/></w:rPr>'
            '</w:style></w:styles>').encode()


def test_fragment_preserves_native_appearance_without_heading_identity():
    m = module()
    original = native_styles()
    fragment = m.flat_opc_fragment(original)
    assert native_styles() == original
    package = ET.fromstring(fragment)
    parts = {p.get(PKG+'name'): p for p in package}
    assert set(parts) == {'/_rels/.rels', '/word/document.xml', '/word/styles.xml', '/word/_rels/document.xml.rels'}
    styles = parts['/word/styles.xml'].find(PKG+'xmlData')[0]
    clone = next(s for s in styles.findall(W+'style') if s.find(W+'name').get(W+'val') == m.CLONE)
    assert clone.find(W+'basedOn').get(W+'val') == 'a'
    assert clone.find(W+'link') is None
    assert clone.find(W+'rPr').find(W+'szCs').get(W+'val') == '48'
    assert clone.find(W+'rPr').find(W+'shd').get(W+'fill') == 'FCE8E6'
    assert clone.find(W+'pPr').find(W+'spacing').get(W+'before') == '280'
    assert m.clone_xml_valid(ET.tostring(styles), 'detached-properties') is False  # source heading is deliberately absent
    document = parts['/word/document.xml'].find(PKG+'xmlData')[0]
    paragraph = document.find(W+'body').find(W+'p')
    assert paragraph.find(W+'pPr').find(W+'pStyle').get(W+'val') == clone.get(W+'styleId')
    assert ''.join(document.itertext()) == m.IMPORTED
    assert b'TargetMode="External"' not in fragment
    assert b'macroEnabled' not in fragment


@pytest.mark.parametrize('change', ['duplicate', 'missing_complex_size', 'style_collision'])
def test_untrusted_or_non_discriminating_seed_is_rejected(change):
    m = module()
    xml = native_styles().decode()
    if change == 'duplicate':
        xml = xml.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style></w:styles>')
    elif change == 'missing_complex_size':
        xml = xml.replace('<w:szCs w:val="48"/>', '')
    else:
        xml = xml.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="x"><w:name w:val="'+m.CLONE+'"/></w:style></w:styles>')
    with pytest.raises(ValueError):
        m.flat_opc_fragment(xml.encode())


def test_import_is_one_native_file_command_bound_to_explicit_range(tmp_path):
    m = module()
    lines = m.import_commands(tmp_path / 'fragment.xml')
    commands = [line for line in lines if line.startswith('insert file ')]
    assert len(commands) == 1
    assert 'at importRange' in commands[0]
    assert 'confirm conversions false link false' in commands[0]
    assert any('paragraph 2 of boundDoc' in line for line in lines)
    assert not any('active document' in line or 'open file' in line for line in lines)


def test_no_execute_guard_leaves_no_files_and_never_enters_runner(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m, 'run', lambda *a: pytest.fail('native runner entered'))
    assert m.main(['--output', str(tmp_path/'absent')]) == 2
    assert not (tmp_path/'absent').exists()


def test_style_preservation_rejects_changed_existing_style():
    m = module()
    original = native_styles()
    assert m.existing_styles_preserved(original, original)
    assert not m.existing_styles_preserved(original, original.replace(b'w:val="48"', b'w:val="24"'))
    assert not m.existing_styles_preserved(original, original.replace(b'w:styleId="1"', b'w:styleId="changed"'))


@pytest.mark.parametrize('mutation', ['', 'wrong_style', 'complex_size', 'paragraph_indent', 'paragraph_mark_font', 'duplicate'])
def test_fresh_paragraph_uses_only_complete_imported_style(mutation):
    m = module()
    package = ET.fromstring(m.flat_opc_fragment(native_styles()))
    imported = next(p for p in package if p.get(PKG+'name') == '/word/styles.xml').find(PKG+'xmlData')[0]
    styles = ET.fromstring(native_styles())
    styles.append(next(s for s in imported.findall(W+'style') if s.get(W+'styleId') == m.STYLE_ID))
    paragraph = '<w:p><w:pPr><w:pStyle w:val="'+m.STYLE_ID+'"/></w:pPr><w:r><w:t>'+m.FRESH+'</w:t></w:r></w:p>'
    if mutation == 'wrong_style':
        paragraph = paragraph.replace(m.STYLE_ID, '1')
    elif mutation == 'complex_size':
        paragraph = paragraph.replace('<w:r>', '<w:r><w:rPr><w:szCs w:val="24"/></w:rPr>')
    elif mutation == 'paragraph_indent':
        paragraph = paragraph.replace('</w:pPr>', '<w:ind w:left="120"/></w:pPr>')
    elif mutation == 'paragraph_mark_font':
        paragraph = paragraph.replace('</w:pPr>', '<w:rPr><w:szCs w:val="24"/></w:rPr></w:pPr>')
    elif mutation == 'duplicate':
        paragraph *= 2
    document = ('<w:document xmlns:w="'+W[1:-1]+'"><w:body>'+paragraph+'</w:body></w:document>').encode()
    assert m.fresh_paragraph_xml_valid(document, ET.tostring(styles)) is (mutation == '')


def test_invalid_preimage_stops_before_file_import(tmp_path, monkeypatch):
    """Malformed readback cannot authorize the later native mutation."""
    from zipfile import ZipFile
    m = module()
    class Session:
        staging_root = tmp_path/'runtime'
        _closed = False
        _quarantined = False
        def __enter__(self):
            self.staging_root.mkdir()
            return self
        def __exit__(self, *args):
            self._closed = True
        def _execute(self, lines):
            if any(line.startswith('insert file ') for line in lines):
                pytest.fail('Invalid preimage reached native import')
            if 'set sentinelDoc to make new document' in lines:
                return [['sentinel']]
            if lines == ['set nativeRows to {{version as text}}']:
                return [['test-host']]
            if lines == m.seed():
                return [['stage', True]]
            return []
        def save_docx(self, path):
            with ZipFile(path, 'w') as z:
                z.writestr('word/styles.xml', native_styles())
    session = Session()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: session)
    monkeypatch.setattr(m, 'retain_sources', lambda *args: {})
    monkeypatch.setattr(m, 'sha', lambda *args: 'hash')
    monkeypatch.setattr(m, 'inventory', lambda *args: [])
    monkeypatch.setattr(m, 'sentinel_preimage', lambda *args: [])
    def closed(owner, output, report, *args):
        assert owner._closed
        report['sentinel_cleanup_attempted'] = True
    monkeypatch.setattr(m, 'close_after_owned', closed)
    report = m.run(tmp_path)
    assert report['status'] == 'FAIL'
    assert report['checks']['source_font_preimage_valid'] is False
    assert report['current_phase'] == 'source_font_preimage'
    assert 'native_import' not in report['raw_rows']
    assert report['remaining_sentinel'] is None
