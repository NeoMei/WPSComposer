"""Pure guards and retained evidence; never execute native Word."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT/'fixtures/microsoft_parity/macos_word_inline_rule_feasibility.py'


def module():
    assert FIXTURE.exists(), 'dedicated inline runner missing'
    spec = importlib.util.spec_from_file_location('inline_feasibility', FIXTURE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def test_execute_guard_precedes_any_output_or_native_work(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: pytest.fail('native launch'))
    target = tmp_path/'absent'
    assert m.main(['--output', str(target)]) == 2
    for value in (False, None, 1, 'true'):
        with pytest.raises(ValueError, match='execute'):
            m.run(target, execute=value)
    assert not target.exists()
    with pytest.raises(FileExistsError):
        m.run(tmp_path, execute=True)


@pytest.mark.parametrize('hash_audit_fails', [False, True])
def test_dictionary_preflight_failure_keeps_report_without_native_inventory(tmp_path, monkeypatch, hash_audit_fails):
    m = module()
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    monkeypatch.setattr(m, 'DICTIONARY', tmp_path/'missing.sdef')
    monkeypatch.setattr(m, 'inventory', lambda *a: pytest.fail('native inventory after preflight failed'))
    if hash_audit_fails:
        def missing_source(*args):
            raise FileNotFoundError('source disappeared during final audit')
        monkeypatch.setattr(m, 'source_pair_matches', missing_source)
    output = tmp_path/'run'
    report = m.run(output, execute=True)
    assert report['status'] == 'FAIL'
    assert report['error']['type'] == 'FileNotFoundError'
    assert 'inventory_failure' not in report
    assert ('evidence_hash_failure' in report) is hash_audit_fails
    assert json.loads((output/'report.json').read_text()) == report


def test_emitted_probe_has_only_exact_document_container_constructor():
    m = module()
    lines = m.build_commands()
    constructors = [line for line in lines if 'make new' in line]
    assert constructors == ['set creationResult to make new standard inline horizontal line at boundDoc with properties {text object:insertionRange}']
    assert 'if (count inline shapes of boundDoc) is not beforeCount + 1 then error "WPSC_RULE_COUNT_DELTA_FAILED"' in lines
    assert '(inline shape type of ownLine is inline shape horizontal line)' in '\n'.join(lines)


@pytest.mark.parametrize('row', [[], ['inline', 0, True, 400, 1], ['inline', True, True, 400, 1],
    ['inline', 2, True, 400, 1], ['inline', 1, 1, 400, 1], ['inline', 1, False, 400, 1],
    ['inline', 1, True, 0, 1], ['inline', 1, True, 400, float('nan')],
    ['inline', 1, True, float('inf'), 1], ['inline', 1, True, 400, True]])
def test_native_rejects_count_type_and_geometry_substitutes(row):
    m = module()
    assert not m.native_valid([row])
    assert m.native_valid([['inline', 1, True, 400, 1]])


def xml(body):
    return ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'
            '<w:body>'+body+'</w:body></w:document>').encode()


def test_xml_requires_native_inline_rule_not_border_image_or_plain_text():
    m = module()
    native = '<w:p><w:r><w:pict><v:rect o:hr="t"/></w:pict></w:r></w:p>'
    assert m.inline_xml_valid(xml(native))
    for body in ('<w:p><w:pPr><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr></w:p>',
                 '<w:p><w:r><w:t>horizontal line</w:t></w:r></w:p>',
                 native+native, native.replace('o:hr="t"', 'o:hr="false"'),
                 native.replace('<v:rect o:hr="t"/>', '<v:rect o:hr="t"><v:imagedata/></v:rect>'),
                 '<v:rect o:hr="t"/>'):
        assert not m.inline_xml_valid(xml(body))


def test_pdf_requires_opaque_horizontal_vector_matching_native_geometry():
    import fitz
    m = module()
    page = fitz.Rect(0, 0, 600, 800)
    rect = fitz.Rect(80, 100, 480, 101)
    good = {'type': 'f', 'rect': rect, 'fill_opacity': 1, 'fill': (0.5, 0.5, 0.5),
            'items': [('re', rect, 1)]}
    assert m.pdf_rule_matches(good, page, [400, 1])
    for change in ({'fill_opacity': 0.1}, {'items': [('l', fitz.Point(80, 100), fitz.Point(480, 101))]},
                   {'items': []}, {'fill': None}, {'rect': fitz.Rect(80, 100, 480, 130)},
                   {'rect': fitz.Rect(80, 100, 100, 101)}):
        assert not m.pdf_rule_matches(dict(good, **change), page, [400, 1])
    stroke = {'type': 's', 'rect': fitz.Rect(80, 100, 480, 100), 'width': 1,
              'stroke_opacity': 1, 'color': (0.5, 0.5, 0.5), 'dashes': '[] 0',
              'items': [('l', fitz.Point(80, 100), fitz.Point(480, 100))]}
    assert m.pdf_rule_matches(stroke, page, [400, 1])
    assert not m.pdf_rule_matches(dict(stroke, dashes='[2 2] 0'), page, [400, 1])
    assert not m.pdf_rule_matches(good, page, [float('nan'), 1])


def test_cleanup_refuses_unclosed_quarantined_and_changed_sentinel(tmp_path, monkeypatch):
    m = module()
    monkeypatch.setattr(m, 'inventory', lambda *args: [['sentinel', '', False, 'b'*64]])
    monkeypatch.setattr(m.subprocess, 'run', lambda *args, **kw: pytest.fail('unsafe close'))
    report = {'checks': {}, 'inventory_before': [], 'sentinel_before': ['sentinel', '', False, 'a'*64]}
    for closed, quarantined in ((False, False), (True, True), (True, False)):
        with pytest.raises(RuntimeError):
            m.close_sentinel_after_owned(SimpleNamespace(_closed=closed, _quarantined=quarantined), tmp_path, report, 'sentinel', 'token')
    assert not (tmp_path/'sentinel-close.applescript').exists()
    assert report['sentinel_close_state'] == 'identity-refused-before-close'
    assert not report.get('native_uncertainty')


def test_ordinary_failure_keeps_partial_native_artifacts_and_primary_error(tmp_path):
    m = module()
    events = []
    class Owner:
        _quarantined = False
        _closed = False
        _bound_path = '/private/owned.docx'
        def _execute(self, lines):
            events.append('read')
            return [['partial-counts', 0, 0]]
        def save_docx(self, path):
            events.append('save')
            path.write_bytes(b'partial native package')
        def export_pdf(self, path):
            events.append('pdf')
            raise RuntimeError('PDF failure')
    report = {'error': {'message': 'original constructor failure'}, 'steps': []}
    m.retain_partial(Owner(), tmp_path, report)
    assert events == ['read', 'save', 'pdf']
    assert (tmp_path/'partial.docx').read_bytes() == b'partial native package'
    assert report['error']['message'] == 'original constructor failure'
    assert 'PDF failure' in report['partial_retention_error']


def test_quarantine_stops_partial_native_reads_and_saves(tmp_path):
    m = module()
    owner = SimpleNamespace(_quarantined=True, _closed=False)
    report = {}
    m.retain_partial(owner, tmp_path, report)
    assert report['partial_retention_skipped'] == 'session closed or quarantined'
    assert list(tmp_path.iterdir()) == []


def test_live_and_retained_sources_both_bound_to_original_hash(tmp_path):
    m = module()
    live = tmp_path/'live'
    retained = tmp_path/'retained'
    live.write_bytes(b'original')
    retained.write_bytes(b'original')
    digest = m.sha(live)
    assert m.source_pair_matches(live, retained, digest)
    retained.write_bytes(b'drift')
    assert not m.source_pair_matches(live, retained, digest)
    retained.write_bytes(b'original')
    live.write_bytes(b'drift')
    assert not m.source_pair_matches(live, retained, digest)


@pytest.mark.parametrize('constructor_fails', [False, True])
@pytest.mark.parametrize('cleanup_fault', [None, 'timeout', 'connection', 'appleevent-timeout', 'malformed-json',
    'wrong-ack', 'presubmission', 'identity-refusal', 'inventory-timeout', 'inventory-malformed', 'final-inventory-timeout'])
def test_runner_closes_owned_before_exact_sentinel_and_preserves_failure_evidence(tmp_path, monkeypatch, constructor_fails, cleanup_fault):
    m = module()
    events = []
    state = {'sentinel': None}
    dictionary = tmp_path/'Word.sdef'
    dictionary.write_bytes(b'controlled dictionary')
    monkeypatch.setattr(m, 'DICTIONARY', dictionary)
    monkeypatch.setattr(m, 'SOURCES', [FIXTURE])
    monkeypatch.setattr(m, 'pdf_checks', lambda *a: {'pdf_anchor_visible': True, 'pdf_native_geometry_visible': True})

    class Owner:
        _closed = False
        _quarantined = False
        _bound_path = '/private/exact-owned.docx'
        def __init__(self, label):
            self.label = label
            self.staging_root = tmp_path/label
            self.staging_root.mkdir()
            self.line = label == 'reopen'
        def __enter__(self):
            return self
        def __exit__(self, *args):
            self._closed = True
            events.append(self.label+'-close')
            (self.staging_root/'close.log').write_text('exact close ACK')
        def _execute_structural(self, lines):
            return self._execute(lines)
        def _execute(self, lines):
            commands = '\n'.join(lines)
            if 'set qualitySentinel to make new document' in commands:
                import re
                match = re.search(r'set content of text object of qualitySentinel to (".*")', commands)
                token = json.loads(match.group(1))
                import hashlib
                state['sentinel'] = ['sentinel', 'sentinel', False, hashlib.sha256((token+'\r').encode()).hexdigest()]
                return [['sentinel', 'test-version']]
            if '"seed"' in commands:
                return [['seed', 0, 0]]
            if '"partial-counts"' in commands:
                return [['partial-counts', 0, 0]]
            if 'set creationResult to make new standard inline horizontal line' in commands:
                events.append('constructor')
                (self.staging_root/'constructor.log').write_text('retained native command output')
                if constructor_fails:
                    raise RuntimeError('WPSC_RULE_COUNT_DELTA_FAILED (-2700)')
                self.line = True
            return [['inline', 1, True, 400, 1]]
        def save_docx(self, path):
            events.append(path.name)
            body = '<w:p><w:r><w:t>'+m.PREFIX+'</w:t></w:r></w:p>'
            if self.line:
                body += '<w:p><w:r><w:pict><v:rect o:hr="t"/></w:pict></w:r></w:p>'
            with ZipFile(path, 'w') as package:
                package.writestr('word/document.xml', xml(body))
                package.writestr('word/styles.xml', b'<styles/>')
        def export_pdf(self, path):
            events.append(path.name)
            path.write_bytes(b'%PDF controlled native export')

    owner = Owner('owned')
    monkeypatch.setattr(m.MacWordSession, 'new_document', lambda **kw: owner)
    monkeypatch.setattr(m.MacWordSession, 'open_document', lambda *a, **kw: Owner('reopen'))

    def inventory(_output, label):
        if label == 'after-owned-close':
            assert owner._closed and events[-1] == 'owned-close'
            events.append('post-owned-inventory')
            if cleanup_fault == 'inventory-timeout':
                raise m.subprocess.TimeoutExpired('synthetic inventory', 30, output=b'inventory partial', stderr=b'inventory error')
            if cleanup_fault == 'inventory-malformed':
                return [['unverified inventory ACK']]
        if label == 'final':
            events.append('final-inventory')
            if cleanup_fault == 'final-inventory-timeout':
                raise m.subprocess.TimeoutExpired('synthetic final inventory', 30)
        return [state['sentinel']] if state['sentinel'] else []
    monkeypatch.setattr(m, 'inventory', inventory)

    def native_close(args, **kwargs):
        assert args[0] == '/usr/bin/osascript'
        assert owner._closed
        assert events[-1] == 'post-owned-inventory'
        source = Path(args[1]).read_text()
        assert 'isEqualToString:' in source and 'path of qualitySentinel' in source
        assert 'if saved of qualitySentinel then' in source
        assert source.index('SENTINEL_TEXT') < source.index('close qualitySentinel saving no')
        events.append('sentinel-close-attempt')
        if cleanup_fault == 'timeout':
            raise m.subprocess.TimeoutExpired('synthetic sentinel close', 30, output=b'partial stdout', stderr=b'partial stderr')
        if cleanup_fault == 'presubmission':
            raise FileNotFoundError('synthetic executable absent before spawn')
        responses = {
            'connection': (1, '', 'execution error: connection invalid (-609)'),
            'appleevent-timeout': (1, '', 'execution error: timeout (-1712)'),
            'malformed-json': (0, 'not JSON', ''),
            'wrong-ack': (0, '[["sentinel-closed", true]]', ''),
            'identity-refusal': (1, '', 'execution error: QUALITY_SENTINEL_TEXT (-2700)'),
        }
        if cleanup_fault in responses:
            code, stdout, stderr = responses[cleanup_fault]
            return SimpleNamespace(returncode=code, stdout=stdout, stderr=stderr)
        state['sentinel'] = None
        events.append('sentinel-close')
        return SimpleNamespace(returncode=0, stdout='[["sentinel-closed"]]', stderr='')
    monkeypatch.setattr(m.subprocess, 'run', native_close)

    output = tmp_path/'run'
    report = m.run(output, execute=True)
    assert events.count('constructor') == 1
    assert events.index('owned-close') < events.index('post-owned-inventory')
    assert (output/'native-runtime/close.log').read_text() == 'exact close ACK'
    assert json.loads((output/'report.json').read_text()) == report
    if cleanup_fault == 'final-inventory-timeout':
        assert report['status'] == 'FAIL' and report['quarantined'] is True
        assert report['remaining_sentinel'] is None and report['sentinel_closed'] is True
        marker = json.loads((output/'native-uncertainty.json').read_text())
        assert marker['owned_document']['closed'] is True
        assert marker['sentinel']['state'] == 'closed-acknowledged'
        assert report['sentinel_state'] == 'closed-acknowledged'
        assert events[-1] == 'final-inventory'
        if constructor_fails:
            assert 'WPSC_RULE_COUNT_DELTA_FAILED' in report['error']['message']
        return
    if cleanup_fault:
        uncertain = cleanup_fault not in ('presubmission', 'identity-refusal')
        assert report['status'] == 'FAIL'
        assert report['remaining_sentinel'] == 'sentinel'
        assert report['native-runtime-identity']['closed'] is True
        assert owner._closed is True and owner._quarantined is False
        assert report['quarantined'] is uncertain
        assert ('final-inventory' in events) is not uncertain
        assert 'reopen-close' not in events and 'sentinel-close' not in events
        assert events.count('sentinel-close-attempt') == (0 if cleanup_fault.startswith('inventory-') else 1)
        assert (output/'native-uncertainty.json').exists() is uncertain
        if uncertain:
            marker = json.loads((output/'native-uncertainty.json').read_text())
            assert marker['owned_document']['closed'] is True
            assert marker['sentinel']['name'] == 'sentinel'
            assert marker['sentinel']['preimage'] == report['sentinel_before']
            assert marker['sentinel']['state'] == 'unknown'
            assert report['sentinel_state'] == 'unknown'
            if cleanup_fault == 'timeout':
                assert 'partial stdout' in (output/'sentinel-close.log').read_text()
                assert 'partial stderr' in (output/'sentinel-close.log').read_text()
            if cleanup_fault == 'inventory-timeout':
                assert 'inventory partial' in (output/'after-owned-close-uncertainty.log').read_text()
        else:
            assert report['sentinel_close_state'] == ('not-submitted' if cleanup_fault == 'presubmission' else 'identity-refused-before-close')
        if constructor_fails:
            assert 'WPSC_RULE_COUNT_DELTA_FAILED' in report['error']['message']
            assert 'cleanup_failure' in report
            assert (output/'partial.docx').is_file() and (output/'partial.pdf').is_file()
            assert 'WPSC_RULE_COUNT_DELTA_FAILED' in (output/'constructor-failure.txt').read_text()
        return
    assert events.index('post-owned-inventory') < events.index('sentinel-close') < events.index('final-inventory')
    assert report['remaining_sentinel'] is None
    assert report['checks']['inventory_preserved'] is True
    if constructor_fails:
        assert report['status'] == 'FAIL'
        assert 'WPSC_RULE_COUNT_DELTA_FAILED' in report['error']['message']
        assert 'WPSC_RULE_COUNT_DELTA_FAILED' in (output/'constructor-failure.txt').read_text()
        assert (output/'partial.docx').is_file() and (output/'partial.pdf').is_file()
        assert (output/'partial-xml/word/document.xml').is_file()
        assert (output/'native-runtime/constructor.log').is_file()
        assert not (output/'after.docx').exists()
        assert 'reopen-close' not in events
    else:
        assert report['status'] == 'PASS'
        assert events.index('sentinel-close') < events.index('reopen-close') < events.index('final-inventory')
        assert (output/'after-xml/word/styles.xml').read_bytes() == b'<styles/>'


def test_live_inventory_uncertainty_quarantines_owner_and_blocks_later_native_calls(tmp_path, monkeypatch):
    m = module()
    events = []
    class Owner:
        _closed = False
        _quarantined = False
        _bound_path = '/private/exact-owned.docx'
        staging_root = tmp_path/'owned'
        def _retain(self, reason):
            self._quarantined = True
            events.append('owner-quarantined')
    owner = Owner()
    def uncertain_inventory(*args):
        events.append('inventory')
        raise m.subprocess.TimeoutExpired('synthetic inventory', 30)
    monkeypatch.setattr(m, 'inventory', uncertain_inventory)
    report = {'sentinel_name': 'sentinel', 'sentinel_before': ['sentinel', '', False, 'a'*64]}
    with pytest.raises(m.subprocess.TimeoutExpired):
        m.independent_inventory(tmp_path, 'sentinel-before', report, owner)
    assert owner._quarantined and not owner._closed
    with pytest.raises(RuntimeError, match='uncertain'):
        m.independent_inventory(tmp_path, 'final', report, owner)
    with pytest.raises(RuntimeError, match='uncertain'):
        m.close_sentinel_after_owned(owner, tmp_path, report, 'sentinel', 'token')
    assert events == ['inventory', 'owner-quarantined']
