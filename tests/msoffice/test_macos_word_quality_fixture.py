"""Pure stage-one runner gates, independent validators and cleanup faults."""
from __future__ import annotations

import importlib.util
import copy
import json
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT/'fixtures/microsoft_parity/macos_word_quality.py'


def module():
    assert FIXTURE.exists(), 'stage-one quality fixture missing'
    spec = importlib.util.spec_from_file_location('quality_fixture', FIXTURE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def test_explicit_execute_and_new_output_precede_native_calls(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: pytest.fail('Word launch'))
    monkeypatch.setattr(m.safe, 'inventory', lambda *a: pytest.fail('Word inventory'))
    target = tmp_path/'absent'
    assert m.main(['--output', str(target)]) == 2
    for value in (False, None, 1, 'yes'):
        with pytest.raises(ValueError, match='execute'):
            m.run(target, execute=value)
    assert not target.exists()
    with pytest.raises(FileExistsError):
        m.run(tmp_path, execute=True)


def test_dictionary_failure_retains_primary_error_without_native_calls(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    monkeypatch.setattr(m, 'DICTIONARY', tmp_path/'missing')
    monkeypatch.setattr(m.safe, 'inventory', lambda *a: pytest.fail('native inventory'))
    report = m.run(tmp_path/'run', execute=True)
    assert report['status'] == 'FAIL'
    assert report['error']['type'] == 'FileNotFoundError'
    assert json.loads((tmp_path/'run/report.json').read_text()) == report


def reservation():
    return [['reservation', '/owned.docx', 'sentinel', 97, 100, 2, 2, 100, 100, True]]


def test_reservation_requires_actual_noncollapsed_inactive_bound_selection():
    m = module()
    good = reservation()
    assert m.reservation_valid(good, '/owned.docx', 'sentinel', 100)
    for index, value in ((1, '/Owned.docx'), (2, 'other'), (3, 100), (4, 99),
                         (5, 3), (7, 101), (8, 101), (9, 1)):
        bad = [good[0].copy()]
        bad[0][index] = value
        assert not m.reservation_valid(bad, '/owned.docx', 'sentinel', 100)
    assert not m.reservation_valid(good + good, '/owned.docx', 'sentinel', 100)


def notice():
    # Display has 54 UTF-16 units; the actual returned table adds two units.
    return [['notice', 100, 156, 256, 3, 2,
             'QUALITY TITLE 中文😀\r[QUALITY_TEST] Notice 中文😀 retained',
             100, 100, True],
            ['style', True, True, True, 0, 3, True, True, False]]


def test_notice_requires_exact_geometry_title_style_bookmark_and_cursor():
    m = module()
    good = notice()
    # Independent literal coordinates are intentionally checked before mutations.
    assert m.notice_valid(good, 100, 200, 156)
    for index, value in ((1, 101), (2, 154), (3, 254), (4, 4), (5, 3),
                         (6, 'wrong text'), (7, 155), (8, 155), (9, 1)):
        bad = [row.copy() for row in good]
        bad[0][index] = value
        assert not m.notice_valid(bad, 100, 200, 156)
    for index, value in ((1, 1), (2, False), (3, False), (4, True), (5, 4),
                         (6, False), (7, False), (8, True)):
        bad = [row.copy() for row in good]
        bad[1][index] = value
        assert not m.notice_valid(bad, 100, 200, 156)
    assert not m.notice_valid(good, 100, 200, 155)


@pytest.mark.parametrize('fault', ['timeout', 'malformed', 'wrong-ack', 'after-close-inventory'])
def test_cleanup_uncertainty_stops_all_native_followup_and_keeps_identity(tmp_path, monkeypatch, fault):
    m = module()
    events = []
    owner = SimpleNamespace(_closed=True, _quarantined=False, _bound_path='/owned.docx', staging_root=tmp_path/'stage')
    report = {'status': 'FAIL', 'checks': {}, 'error': {'message': 'original failure'},
              'confirmed': ['reserved', 'first-notice'], 'inventory_before': [],
              'sentinel_name': 'sentinel', 'sentinel_before': ['sentinel', '', False, 'a'*64]}
    def inventory(output, label):
        events.append(label)
        if label == 'final':
            raise m.safe.subprocess.TimeoutExpired('inventory', 30)
        return [report['sentinel_before']]
    def close(*args, **kwargs):
        events.append('close')
        if fault == 'timeout':
            raise m.safe.subprocess.TimeoutExpired('close', 30, output=b'partial', stderr=b'err')
        return SimpleNamespace(returncode=0, stdout={'malformed': '{', 'wrong-ack': '[]'}.get(fault, '[["sentinel-closed"]]'), stderr='')
    monkeypatch.setattr(m.safe, 'inventory', inventory)
    monkeypatch.setattr(m.safe.subprocess, 'run', close)
    m.cleanup(owner, tmp_path, report)
    frozen = list(events)
    m.cleanup(owner, tmp_path, report)
    assert events == frozen
    assert report['quarantined'] is True
    assert report['error'] == {'message': 'original failure'}
    assert report['confirmed'] == ['reserved', 'first-notice']
    marker = json.loads((tmp_path/'native-uncertainty.json').read_text())
    assert marker['owned_document']['path'] == '/owned.docx'
    assert marker['owned_document']['closed'] is True
    assert marker['sentinel']['preimage'] == report['sentinel_before']
    assert marker['sentinel']['state'] == ('closed-acknowledged' if fault == 'after-close-inventory' else 'unknown')
    assert events == (['after-owned-close', 'close', 'final'] if fault == 'after-close-inventory' else ['after-owned-close', 'close'])


def test_quarantined_owner_never_closes_sentinel_or_inventories(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.safe, 'inventory', lambda *a: pytest.fail('native followup'))
    report = {'checks': {}, 'sentinel_name': 'sentinel', 'error': {'message': 'failed mutation'}}
    m.cleanup(SimpleNamespace(_closed=False, _quarantined=True), tmp_path, report)
    assert report['quarantined'] is True
    assert report['remaining_sentinel'] == 'sentinel'


def test_cleanup_never_retries_a_close_already_attempted_by_main_flow(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.safe, 'close_sentinel_after_owned', lambda *a: pytest.fail('second close attempt'))
    monkeypatch.setattr(m.safe, 'inventory', lambda *a: [['sentinel', '', False, 'a'*64]])
    report = {'checks': {}, 'sentinel_name': 'sentinel', 'sentinel_cleanup_attempted': True,
              'inventory_before': [], 'error': {'message': 'guard refused'}}
    owner = SimpleNamespace(_closed=True, _quarantined=False)
    m.cleanup(owner, tmp_path, report)
    assert report['remaining_sentinel'] == 'sentinel'
    assert report['error']['message'] == 'guard refused'
    assert 'cleanup_error' not in report


def test_final_inventory_after_prior_close_keeps_sentinel_identity(tmp_path, monkeypatch):
    m = module()
    def fail(*args):
        raise m.safe.subprocess.TimeoutExpired('final', 30)
    monkeypatch.setattr(m.safe, 'inventory', fail)
    monkeypatch.setattr(m.safe, 'close_sentinel_after_owned', lambda *a: pytest.fail('sentinel already closed'))
    report = {'checks': {}, 'sentinel_name': 'sentinel', 'sentinel_closed': True,
              'sentinel_before': ['sentinel', '', False, 'a'*64], 'inventory_before': []}
    owner = SimpleNamespace(_closed=True, _quarantined=False, _bound_path='/owned.docx', staging_root=tmp_path/'stage')
    m.cleanup(owner, tmp_path, report)
    assert report['remaining_sentinel'] is None
    assert report['native_uncertainty']['sentinel']['name'] == 'sentinel'
    assert report['native_uncertainty']['sentinel']['state'] == 'closed-acknowledged'
    assert 'cleanup_error' not in report


def test_confirmation_record_survives_local_failure_and_cleanup(tmp_path):
    m = module()
    report = {'error': {'message': 'first failure'}, 'confirmed': ['first-notice']}
    m.record_failure(tmp_path, report, RuntimeError('cleanup error'), 'cleanup')
    assert report['error']['message'] == 'first failure'
    assert report['confirmed'] == ['first-notice']
    assert report['cleanup_error']['message'] == 'cleanup error'


def test_run_keeps_primary_native_error_when_quarantine_retention_itself_fails(tmp_path, monkeypatch):
    m = module()
    dictionary = tmp_path/'Word.sdef'
    dictionary.write_text('controlled dictionary')
    monkeypatch.setattr(m, 'DICTIONARY', dictionary)
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    events = []
    def inventory(output, label):
        events.append(label)
        return []
    monkeypatch.setattr(m.safe, 'inventory', inventory)
    class Owner:
        _closed = False
        _quarantined = False
        _bound_path = '/owned.docx'
        staging_root = tmp_path/'stage'
        def _execute_structural(self, lines):
            events.append('seed')
            raise RuntimeError('original native failure')
        def _retain(self, reason):
            self._quarantined = True
            raise OSError('quarantine marker disk failed')
    owner = Owner()
    owner.staging_root.mkdir()
    (owner.staging_root/'original.log').write_text('original native log')
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: owner)
    output = tmp_path/'run'
    report = m.run(output, execute=True)
    assert report['error']['message'] == 'original native failure'
    assert report['quarantine_retention_error']['message'] == 'quarantine marker disk failed'
    assert report['quarantined'] is True and report['status'] == 'FAIL'
    assert events == ['before', 'seed']
    assert (output/'native-runtime/original.log').read_text() == 'original native log'
    assert 'original native failure' in (output/'primary-failure.txt').read_text()


def test_creation_failure_before_session_return_blocks_final_inventory(tmp_path, monkeypatch):
    m = module()
    dictionary = tmp_path/'Word.sdef'
    dictionary.write_text('controlled dictionary')
    monkeypatch.setattr(m, 'DICTIONARY', dictionary)
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    events = []
    def inventory(output, label):
        events.append(label)
        return []
    def failed_creation(**kwargs):
        raise RuntimeError('new document completion uncertain')
    monkeypatch.setattr(m.safe, 'inventory', inventory)
    monkeypatch.setattr(m.MacWordSession, 'new_document', failed_creation)
    report = m.run(tmp_path/'run', execute=True)
    assert events == ['before']
    assert report['quarantined'] is True
    assert report['native_uncertainty']['stage'] == 'open-owned'
    assert report['error']['message'] == 'new document completion uncertain'


def test_xml_preservation_rejects_field_or_surrounding_style_changes():
    m = module()
    def xml(body):
        return ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '</w:body></w:document>').encode()
    prefix = '<w:p><w:pPr><w:spacing w:after="140"/></w:pPr><w:r><w:t>prefix</w:t></w:r><w:r><w:instrText> SEQ Quality2 </w:instrText></w:r></w:p>'
    suffix = '<w:p><w:r><w:t>suffix</w:t></w:r></w:p>'
    table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>QUALITY TITLE 中文😀</w:t></w:r></w:p><w:p><w:r><w:t>[QUALITY_TEST] Notice 中文😀 retained</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    before = xml(prefix+suffix)
    after = xml(prefix+table+suffix)
    assert m.xml_preserved(before, after)
    assert not m.xml_preserved(before, after.replace(b'Quality2', b'Quality3'))
    assert not m.xml_preserved(before, after.replace(b'140', b'160'))
    assert not m.xml_preserved(before, xml(prefix+table+table+suffix))


def test_source_manifest_includes_cleanup_transitive_dependencies_and_test():
    m = module()
    paths = {p.relative_to(ROOT).as_posix() for p in m.SOURCES}
    assert {'fixtures/microsoft_parity/macos_word_inline_rule_feasibility.py',
            'fixtures/microsoft_parity/macos_word_quality.py',
            'fixtures/microsoft_parity/macos_word_quality_feasibility.py',
            'skills/WPSComposer/scripts/msoffice/macos_word_quality.py',
            'tests/msoffice/test_macos_word_quality_fixture.py', 'pyproject.toml'} <= paths


@pytest.mark.parametrize('fault', ['close-timeout', 'close-malformed', 'after-owned-timeout', 'final-malformed', 'reopen-timeout'])
def test_run_preserves_confirmed_first_notice_and_never_reopens_after_cleanup_uncertainty(tmp_path, monkeypatch, fault):
    m = module()
    dictionary = tmp_path/'Word.sdef'
    dictionary.write_text('controlled dictionary')
    monkeypatch.setattr(m, 'DICTIONARY', dictionary)
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    monkeypatch.setattr(m, 'pdf_checks', lambda output: True)
    events = []
    h = 'a'*64
    base = [['body', 200, h, 2, 2, 2, 1],
            ['paragraph', 1, 0, 100, h, [13], []], ['paragraph', 2, 100, 200, h, [13], []],
            ['field', 1, 'sequence', 30, 34, 35, 36, h, h, False],
            ['field', 2, 'sequence', 140, 145, 146, 147, h, h, False],
            ['table', 1, 10, 20, 1, 1, h], ['table', 2, 160, 170, 1, 1, h],
            ['bookmark', 'wpsc_quality_splice', 100, 100, h], ['end']]
    empty = copy.deepcopy(base)
    empty[0][6] = 2
    empty.insert(-1, ['bookmark', m.ANCHOR, 100, 100, h])
    after = copy.deepcopy(empty)
    after[0][1] = 256
    after[0][4] = 3
    after[6][1] = 3
    after.insert(6, ['table', 2, 100, 156, 1, 1, h])
    assert m.state_valid(base) and m.state_valid(empty) and m.state_valid(after)
    class Owner:
        _closed = False
        _quarantined = False
        _bound_path = '/owned.docx'
        reserved = False
        inserted = False
        staging_root = tmp_path/'owned-stage'
        def _execute(self, lines):
            script = '\n'.join(lines)
            if 'set qualitySentinel to make new document' in script:
                # Recover the actual synthetic token for the independent inventory.
                import re
                self.token = re.search(r'to "(QUALITY STAGE ONE SENTINEL[^"\n]+)"', script).group(1)
                events.append('create-sentinel')
                return [['sentinel', '16.test']]
            if 'set nativeRows to {{"selected",true}}' in script:
                return [['selected', True]]
            if 'set nativeRows to {{"seed",true}}' in script:
                return [['seed', True]]
            if 'set nativeRows to {{"reservation"' in script:
                return reservation()
            if 'set qt to table 2 of boundDoc' in script:
                return notice()
            return copy.deepcopy(after if self.inserted else empty if self.reserved else base)
        _execute_structural = _execute
        def reserve_document_quality_anchor(self, title):
            self.reserved = True
            self._quality_notice_anchor_position = 100
        def upsert_document_quality_notice(self, issue):
            self.inserted = True
            self._quality_notice_anchor_position = 156
        def save_docx(self, path):
            prefix = '<w:p><w:r><w:t>prefix</w:t></w:r></w:p>'
            box = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>QUALITY TITLE 中文😀</w:t></w:r></w:p><w:p><w:r><w:t>[QUALITY_TEST] Notice 中文😀 retained</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+prefix+(box if self.inserted else '')+'</w:body></w:document>'
            with ZipFile(path, 'w') as package:
                package.writestr('word/document.xml', xml)
                package.writestr('word/styles.xml', '<styles/>')
        def export_pdf(self, path):
            path.write_bytes(b'controlled PDF boundary')
        def close(self):
            events.append('owned-close' if self is owner else 'reopen-close')
            self._closed = True
        def _retain(self, reason):
            self._quarantined = True
    owner = Owner()
    owner.staging_root.mkdir()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: owner)
    def reopen(*args, **kwargs):
        events.append('reopen')
        assert kwargs['read_only'] is True
        if fault == 'reopen-timeout':
            raise RuntimeError('readonly open completion uncertain')
        other = Owner()
        other.inserted = other.reserved = True
        other._quality_notice_anchor_position = 156
        other.staging_root = tmp_path/'reopen-stage'
        other.staging_root.mkdir()
        return other
    monkeypatch.setattr(m.MacWordSession, 'open_document', reopen)
    closed = False
    def inventory(output, label):
        events.append(label)
        if label == 'before':
            return []
        if fault == 'after-owned-timeout' and label == 'after-owned-close':
            raise m.safe.subprocess.TimeoutExpired('inventory', 30)
        if label == 'final':
            return [['malformed']]
        import hashlib
        return [] if closed else [['sentinel', '', False, hashlib.sha256((owner.token+'\r').encode()).hexdigest()]]
    def close(*args, **kwargs):
        nonlocal closed
        events.append('sentinel-close')
        if fault == 'close-timeout':
            raise m.safe.subprocess.TimeoutExpired('sentinel', 30)
        closed = True
        return SimpleNamespace(returncode=0, stdout='{}' if fault == 'close-malformed' else '[["sentinel-closed"]]', stderr='')
    monkeypatch.setattr(m.safe, 'inventory', inventory)
    monkeypatch.setattr(m.safe.subprocess, 'run', close)
    output = tmp_path/'run'
    report = m.run(output, execute=True)
    assert report['confirmed'] == ['reserved', 'first-notice', 'duplicate']
    assert report['status'] == 'FAIL' and report['quarantined'] is True
    assert (output/'partial.docx').exists() and (output/'partial.pdf').exists()
    assert (output/'after.docx').exists()
    assert report['native-runtime-identity']['closed'] is True
    assert report['native_uncertainty']['sentinel']['name'] == 'sentinel'
    assert report['native_uncertainty']['sentinel']['preimage'] == report['sentinel_before']
    if fault == 'reopen-timeout':
        assert events[-1] == 'reopen' and 'final' not in events
        assert report['remaining_sentinel'] is None
        assert report['native_uncertainty']['stage'] == 'open-readonly'
        assert report['error']['message'] == 'readonly open completion uncertain'
    elif fault == 'final-malformed':
        assert events[-1] == 'final'
        assert report['remaining_sentinel'] is None
        assert report['native_uncertainty']['sentinel']['state'] == 'closed-acknowledged'
        assert 'inventory_error' in report
    else:
        assert 'reopen' not in events and 'final' not in events
        assert events.count('sentinel-close') == (0 if fault == 'after-owned-timeout' else 1)
        assert 'error' in report
    assert json.loads((output/'report.json').read_text()) == report
