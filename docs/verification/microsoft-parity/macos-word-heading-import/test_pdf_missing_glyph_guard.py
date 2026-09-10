import importlib.util
from pathlib import Path


def load():
    path = Path(__file__).with_name('pdf-missing-glyph-diagnostic.py')
    spec = importlib.util.spec_from_file_location('glyph_guard', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_healthy_mapped_glyphs_only_satisfy_negative_gate():
    assert load().missing_glyph_gate('Eا😀', [(69, 1), (1575, 2), (128512, 3)])


def test_missing_unicode_markers_fail_even_with_positive_glyph_id():
    for char in ('\0', '\ufffd'):
        assert not load().missing_glyph_gate(char, [(ord(char), 12)])


def test_glyph_zero_fails_even_with_plausible_unicode_text():
    assert not load().missing_glyph_gate('ا', [(1575, 0)])


def test_empty_or_malformed_trace_is_not_success():
    for rows in ([], [(69, True)], [(True, 5)], [(69, -1)], [('69', 5)]):
        assert not load().missing_glyph_gate('E', rows)


def test_existing_native_failure_is_rejected_without_native_execution():
    import json
    path = Path(__file__).with_name('native-donor-import-05-parent-review') / 'pdf-font-observations.json'
    text = ''.join(row['text'] for row in json.loads(path.read_text())['spans'])
    assert text.count('\0') == 21
    assert not load().missing_glyph_gate(text, [(65533, 0)] * 21)
