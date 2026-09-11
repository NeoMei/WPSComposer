from pathlib import Path
import importlib.util
import pytest

PATH = Path(__file__).with_name('docx-collapsed-control.py')

def module():
    assert PATH.exists(), 'collapsed control absent'
    spec = importlib.util.spec_from_file_location('collapsed', PATH)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value

def test_native_splice_handles_astral_and_terminal_paragraph():
    m = module()
    before = '😀\rREPLACE\r尾\r\r'
    inserted, deleted, start, end = m.expected_stages(before, 3, 11, 'DOC😀\r')
    assert inserted == '😀\rDOC😀\rREPLACE\r尾\r\r'
    assert deleted == '😀\rDOC😀\r尾\r\r'
    assert (start, end) == (9, 17)

def test_clear_compiles_full_preimage_and_shifted_target_guard():
    m = module()
    code = '\n'.join(m.clear_commands('DOC\rREPLACE\r尾\r\r', 4, 12))
    assert code.index('WPSC_COLLAPSED_FULL_PREIMAGE') < code.index('set content of clearRange to ""')
    assert code.index('WPSC_COLLAPSED_TARGET_CHANGED') < code.index('set content of clearRange to ""')

def test_stage_ack_rejects_append_or_wrong_native_bounds():
    m = module()
    assert m.exact_ack([['insert',3,3,'A']], 'insert', 3, 3, 'A')
    assert not m.exact_ack([['insert',3,11,'A']], 'insert', 3, 3, 'A')
    assert not m.exact_ack([['insert',3,3,'AREPLACE']], 'insert', 3, 3, 'A')

def test_missing_execute_never_touches_output(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m, 'run', lambda *args: pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')]) == 2
    assert not (tmp_path/'absent').exists()
