"""Pure feasibility guard/observation tests; these never launch Word."""
from __future__ import annotations

import copy
import importlib.util
from pathlib import Path
import subprocess
import sys

import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / 'fixtures/microsoft_parity/macos_word_quality_feasibility.py'


def module():
    assert FIXTURE.exists(), 'quality feasibility fixture is not implemented'
    spec = importlib.util.spec_from_file_location('quality_feasibility', FIXTURE)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


@pytest.mark.parametrize('mode', ['bound-selection', 'middle-table-recovery'])
def test_missing_execute_leaves_no_output_and_does_not_launch(mode, tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: pytest.fail('native launch'))
    target = tmp_path / 'absent'
    assert m.main(['--mode', mode, '--output', str(target)]) == 2
    assert not target.exists()
    with pytest.raises(ValueError, match='execute'):
        m.run(target, mode)
    assert not target.exists()


def test_unknown_mode_and_existing_output_fail_before_native(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: pytest.fail('native launch'))
    with pytest.raises(ValueError, match='mode'):
        m.run(tmp_path / 'absent', 'all', execute=True)
    with pytest.raises(FileExistsError):
        m.run(tmp_path, 'bound-selection', execute=True)


def state():
    return [['body', 100, 'a' * 64, 1, 1, 1, 1],
            ['paragraph', 1, 0, 15, 'b' * 64, [13], ['normal', 6, 9, 12, 18]],
            ['field', 1, 'field ref', 15, 19, 20, 25, 'c' * 64, 'd' * 64, False],
            ['table', 1, 28, 37, 1, 1, 'e' * 64],
            ['bookmark', 'wpsc_quality_splice', 50, 50, 'f' * 64],
            ['end']]


def recovery_rows():
    before = state()
    partial = copy.deepcopy(before)
    partial[0][1] = 122
    partial[0][2] = '9' * 64
    partial[0][4] = 2
    partial.insert(-2, ['table', 2, 52, 71, 1, 1, '8' * 64])
    return [['middle', 50, 100, 52, 71, 'native-table-marker', -2700, 0],
            ['before', before], ['partial', partial], ['after-delete', copy.deepcopy(before)]]


def test_middle_recovery_requires_actual_nonterminal_insert_and_all_state_equality():
    m = module()
    rows = recovery_rows()
    assert m.assess_middle(rows)['restored'] is True
    for group, index, value in [(0, 2, 'A' * 64), (1, 5, ['normal', 6, 9, 24, 18]),
                                (2, 7, 'C' * 64), (3, 3, 38), (4, 1, 'WPSC_quality_splice')]:
        bad = copy.deepcopy(rows)
        bad[3][1][group][index] = value
        assert m.assess_middle(bad)['restored'] is False
    bad = copy.deepcopy(rows)
    bad[0][1] = 99
    assert m.assess_middle(bad)['observed_middle'] is False
    bad = copy.deepcopy(rows)
    bad[0][7] = -1728
    assert m.assess_middle(bad)['restored'] is False


@pytest.mark.parametrize('bad', [[], [['ok']], recovery_rows() + [['extra']], [['middle', True]]])
def test_malformed_middle_observations_never_certify_recovery(bad):
    assert not module().assess_middle(bad)['restored']


def test_selection_requires_different_active_sentinel_and_exact_body_preservation():
    m = module()
    rows = [['selection', '/owned/Case.docx', 7, 9, 9, 2, 2, 60, 9, 9, True, True, True, 'a' * 64, 'a' * 64]]
    assert m.assess_selection(rows, '/owned/Case.docx', 7, 9)['reserved']
    for column, value in [(1, '/owned/case.docx'), (2, 9), (3, 7), (5, 9), (8, 10), (10, 1), (14, 'b' * 64)]:
        bad = copy.deepcopy(rows)
        bad[0][column] = value
        assert not m.assess_selection(bad, '/owned/Case.docx', 7, 9)['reserved']


def test_xml_observation_reports_exact_changes_and_retains_bookmark_collapse():
    m = module()
    xml = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:bookmarkStart w:id="1" w:name="wpsc_document_quality_anchor"/><w:bookmarkEnd w:id="1"/><w:r><w:t>prefix</w:t></w:r></w:p></w:body></w:document>'
    assert m.xml_observation(xml, xml)['document_xml_exact']
    assert m.xml_observation(xml, xml)['collapsed_quality_bookmark']
    changed = xml.replace(b'prefix', b'Prefix')
    assert not m.xml_observation(xml, changed)['document_xml_exact']
    changed = xml.replace(b'<w:bookmarkEnd w:id="1"/>', b'').replace(b'</w:r>', b'</w:r><w:bookmarkEnd w:id="1"/>')
    assert not m.xml_observation(xml, changed)['collapsed_quality_bookmark']


def test_recovery_phase_has_one_owned_delete_and_no_automatic_repair():
    m = module()
    source = '\n'.join(m.middle_commands())
    assert source.count('delete qualityNewTable') == 1
    assert source.index('set partialRows to nativeRows') < source.index('delete qualityNewTable')
    assert source.index('delete qualityNewTable') < source.index('set afterRows to nativeRows')
    for forbidden in ('undo', 'clipboard', 'rollback(', 'rollbackRange', 'set content of rollback', 'delete table 1'):
        assert forbidden not in source.lower()


def test_selection_probe_reads_inactive_bound_window_and_exact_identity():
    m = module()
    source = '\n'.join(m.selection_commands('/owned/Case.docx', 7, 9))
    assert 'selection of boundWindow' in source
    assert 'document of qualitySelection' in source
    assert 'isEqualToString:' in source
    assert 'activate object boundWindow' not in source


def test_source_digest_check_detects_live_and_retained_drift(tmp_path):
    m = module()
    source = tmp_path / 'live.py'
    retained = tmp_path / 'retained.py'
    source.write_bytes(b'original')
    retained.write_bytes(b'original')
    digest = m.sha(source)
    assert m.source_pair_matches(source, retained, digest)
    source.write_bytes(b'Original')
    assert not m.source_pair_matches(source, retained, digest)
    source.write_bytes(b'original')
    retained.write_bytes(b'other')
    assert not m.source_pair_matches(source, retained, digest)


def test_observation_rejects_inconsistent_topology_and_unobserved_marker():
    m = module()
    rows = recovery_rows()
    rows[0][5] = 'not the inserted marker'
    assert not m.assess_middle(rows)['restored']
    rows = recovery_rows()
    rows[1][1][0][4] = 100
    rows[2][1][0][4] = 101
    rows[3][1][0][4] = 100
    assert not m.assess_middle(rows)['restored']


@pytest.mark.skipif(sys.platform != 'darwin', reason='local dictionary syntax compiler only')
@pytest.mark.parametrize('mode', ['bound-selection', 'middle-table-recovery'])
def test_compiles_local_dictionary_phases_without_running_word(mode, tmp_path):
    m = module()
    phases = [m.seed_commands(mode), m.snapshot_commands()]
    phases += [m.selection_commands('/owned/Case.docx', 7, 9)] if mode == 'bound-selection' else [m.middle_commands()]
    for index, commands in enumerate(phases):
        source = tmp_path / ('phase-%s.applescript' % index)
        source.write_text(m._JSON+'\ntell application "/Applications/Microsoft Word.app"\n'+'\n'.join(commands)+'\nend tell\n')
        result = subprocess.run(['/usr/bin/osacompile', '-o', str(source.with_suffix('.scpt')), str(source)], capture_output=True, text=True, timeout=20)
        assert result.returncode == 0, result.stderr


def test_sentinel_cleanup_refuses_native_calls_before_owned_close(tmp_path, monkeypatch):
    from types import SimpleNamespace
    m = module()
    monkeypatch.setattr(m, 'inventory', lambda *a: pytest.fail('inventory before owned close'))
    for closed, quarantined in ((False, False), (True, True)):
        with pytest.raises(RuntimeError, match='Owned close unverified'):
            m.close_sentinel_after_owned(SimpleNamespace(_closed=closed, _quarantined=quarantined), tmp_path, {}, 'sentinel', 'token')


def test_changed_sentinel_never_reaches_delete_script(tmp_path, monkeypatch):
    from types import SimpleNamespace
    m = module()
    monkeypatch.setattr(m, 'inventory', lambda *a: [['sentinel', '', False, 'changed']])
    monkeypatch.setattr(m.subprocess, 'run', lambda *a, **kw: pytest.fail('sentinel close after drift'))
    report = {'sentinel_before': ['sentinel', '', False, 'original'], 'inventory_before': []}
    with pytest.raises(RuntimeError, match='Inventory or sentinel changed'):
        m.close_sentinel_after_owned(SimpleNamespace(_closed=True, _quarantined=False), tmp_path, report, 'sentinel', 'token')
    assert not (tmp_path / 'sentinel-close.applescript').exists()
