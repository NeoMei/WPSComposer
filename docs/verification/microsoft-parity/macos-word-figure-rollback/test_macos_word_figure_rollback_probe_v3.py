"""Build-only inline-picture rollback probe; no production support declaration."""
from __future__ import annotations

from copy import deepcopy
import importlib.util
from pathlib import Path
import subprocess
import sys

import pytest


PROBE = Path(__file__).with_name('macos_word_figure_rollback_probe_v3.py')


def fixture():
    spec = importlib.util.spec_from_file_location(
        'build_only_macos_word_figure_rollback_probe', PROBE)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_probe_requires_explicit_execute_before_creating_directory(tmp_path):
    result = subprocess.run([
        sys.executable,
        str(PROBE),
        '--output', str(tmp_path / 'untouched'),
    ], capture_output=True, text=True)
    assert result.returncode == 2 and '--execute' in result.stderr
    assert not (tmp_path / 'untouched').exists()


def test_probe_deterministic_png_is_valid_and_stable(tmp_path):
    module = fixture()
    first, second = tmp_path / 'first.png', tmp_path / 'second.png'
    module.write_probe_png(first)
    module.write_probe_png(second)
    assert first.read_bytes().startswith(b'\x89PNG\r\n\x1a\n')
    assert module.sha(first) == module.sha(second)


def test_probe_state_validator_is_closed_and_strictly_typed():
    module = fixture()
    rows = module.expected_state()
    assert module.validate_state(rows)
    for row_index, column, value in (
            (0, 1, rows[0][1] + 'extra'),
            (1, 1, True),
            (2, 4, 'changed'),
            (3, 2, module.END + 1),
            (4, 7, 1),
            (4, 9, False),
            (4, 10, True),
            (5, 1, -1)):
        broken = deepcopy(rows)
        broken[row_index][column] = value
        assert not module.validate_state(broken)
    assert not module.validate_state(rows + [['extra']])


def scenario_rows(module, label, failed):
    return [
        ['inserted', label, module.START, module.START + 1, 1, 0,
         f'figure:probe/{label}', 80, 50, True, True, True,
         module.START + 2, module.START + 2, failed],
        *module.expected_state(),
    ]


@pytest.mark.parametrize(('label', 'failed'), [
    ('one-image', False), ('second-failure', True),
])
def test_probe_scenario_validator_requires_exact_insert_and_restoration(label, failed):
    module = fixture()
    rows = scenario_rows(module, label, failed)
    assert module.validate_scenario(rows, label, failed=failed)
    for column, value in (
            (2, module.START - 1), (3, module.START + 2),
            (4, True), (5, 1), (6, 'other'),
            (7, 78), (9, 1), (12, module.START + 3),
            (13, module.START + 3), (14, not failed)):
        broken = deepcopy(rows)
        broken[0][column] = value
        assert not module.validate_scenario(broken, label, failed=failed)
    broken = deepcopy(rows)
    broken[-1][-1] = 3
    assert not module.validate_scenario(broken, label, failed=failed)


def test_probe_second_failure_is_after_first_ack_and_rollback_uses_suffix_bookmark():
    module = fixture()
    lines = module.scenario_commands(
        'second-failure', '/private/staged.png',
        missing='/private/definitely-missing.png')
    source = '\n'.join(lines)
    first = source.index('set probePicture to make new inline picture')
    acknowledged = source.index('set probePictureEnd to end of content')
    second = source.index('make new inline picture at probeSecondInsertion')
    rollback_bound = source.index('set probeRollbackEnd to probeMutationEnd')
    rollback = source.index('set content of probeRollback to probeOriginalText')
    assert first < acknowledged < second < rollback_bound < rollback
    assert 'WPSC_FIGURE_PROBE_SECOND_FAILURE_SIDE_EFFECT' in source
    assert 'end of content of text object of boundDoc' not in source[rollback_bound:rollback]


def test_probe_remains_build_only_and_preserves_owned_source():
    source = PROBE.read_text()
    assert "source_document = output / 'figure-rollback-source.docx'" in source
    assert 'session.save_copy(output_document)' in source
    assert "report['checks']['owned_document_source_preserved'] = True" in source
    assert 'def add_captioned_figure' not in source


def test_probe_generated_applescripts_compile_without_execution(tmp_path):
    if sys.platform != 'darwin':
        pytest.skip('Word dictionary compile only')
    module = fixture()
    groups = (
        module.seed_commands(),
        module.state_commands(),
        module.scenario_commands('one-image', '/private/staged.png'),
        module.scenario_commands(
            'second-failure', '/private/staged.png',
            missing='/private/definitely-missing.png'),
    )
    for index, lines in enumerate(groups):
        path = tmp_path / f'figure-rollback-{index}.applescript'
        path.write_text(
            'use framework "Foundation"\nuse scripting additions\n'
            'tell application "Microsoft Word"\n'
            'set boundDoc to active document\nset boundWindow to active window\n'
            + '\n'.join(lines) + '\nend tell\n')
        result = subprocess.run([
            '/usr/bin/osacompile', '-o', str(path.with_suffix('.scpt')), str(path),
        ], capture_output=True, text=True, timeout=20)
        assert result.returncode == 0, result.stderr


@pytest.mark.parametrize('column', [7, 8])
@pytest.mark.parametrize('value', [float('nan'), float('inf'), -float('inf')])
def test_probe_rejects_nonfinite_picture_dimensions(column, value):
    module = fixture()
    rows = scenario_rows(module, 'one-image', False)
    rows[0][column] = value
    assert not module.validate_scenario(rows, 'one-image', failed=False)


def test_full_document_oracle_rejects_actual_native_v2_proofing_drift():
    module = fixture()
    evidence = PROBE.parent / 'native-v2-01'
    source = evidence / 'figure-rollback-source.docx'
    restored = evidence / 'figure-rollback.docx'
    assert module.document_signature(source) == module.document_signature(source)
    assert module.document_signature(source) != module.document_signature(restored)


def test_full_document_oracle_ignores_only_known_edit_metadata(tmp_path):
    import zipfile
    from xml.etree import ElementTree as ET
    module = fixture()
    source = PROBE.parent / 'native-v2-01/figure-rollback-source.docx'
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
    with zipfile.ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    def altered(name, mutate):
        doc = ET.fromstring(parts['word/document.xml'])
        mutate(doc)
        path = tmp_path / name
        with zipfile.ZipFile(path, 'w') as archive:
            for key, value in parts.items():
                archive.writestr(key, ET.tostring(doc) if key == 'word/document.xml' else value)
        return module.document_signature(path)
    expected = module.document_signature(source)
    assert altered('metadata.docx', lambda d: d.find('.//' + W + 'p').set(W14+'textId','NEW')) == expected
    assert altered('identity.docx', lambda d: d.find('.//' + W + 'p').set(W14+'paraId','NEW')) != expected
    assert altered('size.docx', lambda d: d.find('.//' + W + 'sz').set(W+'val','99')) != expected
    assert altered('bookmark.docx', lambda d: d.find('.//' + W + 'bookmarkStart').set(W+'name','changed')) != expected
