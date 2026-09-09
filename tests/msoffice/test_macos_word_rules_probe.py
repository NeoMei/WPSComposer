"""Fail-closed evidence validation for the isolated native rules probe."""
from __future__ import annotations
import importlib.util
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile
import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / 'fixtures/microsoft_parity/macos_word_rules_probe.py'


def module():
    assert FIXTURE.exists(), 'isolated rules probe has not been built'
    spec = importlib.util.spec_from_file_location('rules_probe', FIXTURE)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


@pytest.mark.parametrize('case,rows', [
    ('inline', [['inline', 1, True, 432, 2.25]]),
    ('paragraph', [['paragraph', ' \r', True, True, True, True, 'FOLLOWING\r', True]]),
    ('wordart', [['wordart', 1, True, 'WORDART 中文😀', 'Arial', 36, False, False, 200, 400, True, 140, 40]]),
])
def test_native_readback_requires_exact_types_and_all_semantic_dimensions(case, rows):
    m = module()
    assert m.verify_native_rows(case, rows)
    assert not m.verify_native_rows(case, [])
    assert not m.verify_native_rows(case, rows + rows)
    assert not m.verify_native_rows(case, [rows[0][:-1]])
    for index, value in enumerate(rows[0]):
        broken = [list(rows[0])]
        broken[0][index] = None
        assert not m.verify_native_rows(case, broken), (case, index)
    for index, value in enumerate(rows[0]):
        if type(value) is bool:
            broken = [list(rows[0])]; broken[0][index] = int(value)
            assert not m.verify_native_rows(case, broken)


@pytest.mark.skipif(sys.platform != 'darwin', reason='macOS dictionary compiler')
@pytest.mark.parametrize('case', ['inline', 'paragraph', 'wordart'])
def test_dictionary_scripts_compile_without_launch(case, tmp_path):
    m = module()
    source = 'tell application "/Applications/Microsoft Word.app"\n' + '\n'.join(m.build_probe_commands(case)) + '\nend tell\n'
    script = tmp_path / (case + '.applescript'); script.write_text(source)
    result = subprocess.run(['/usr/bin/osacompile', '-o', str(tmp_path / (case + '.scpt')), str(script)], capture_output=True, text=True, timeout=20)
    assert result.returncode == 0, result.stderr


def test_cli_refuses_native_without_explicit_execute(tmp_path):
    m = module()
    assert m.main(['--output', str(tmp_path / 'absent')]) == 2
    assert not (tmp_path / 'absent').exists()


def test_xml_verifier_does_not_accept_text_or_textbox_as_native_wordart():
    m = module()
    xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>WORDART 中文😀</w:t></w:r></w:p></w:body></w:document>'
    assert not m.verify_xml('wordart', xml.encode())
    assert not m.verify_xml('inline', xml.encode())
    assert not m.verify_xml('paragraph', xml.encode())


def test_paragraph_xml_requires_exact_border_and_clean_following():
    m = module()
    xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="C0C0C0"/></w:pBdr></w:pPr><w:r><w:t> </w:t></w:r></w:p><w:p><w:r><w:t>FOLLOWING</w:t></w:r></w:p></w:body></w:document>'
    assert m.verify_xml('paragraph', xml.encode())
    for old, new in [('w:sz="6"','w:sz="8"'), ('C0C0C0','000000'), ('center','left'), ('FOLLOWING','WRONG')]:
        assert not m.verify_xml('paragraph', xml.replace(old,new).encode())
    assert not m.verify_xml('paragraph', xml.replace('<w:p><w:r><w:t>FOLLOWING', '<w:p><w:pPr><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr><w:r><w:t>FOLLOWING').encode())
