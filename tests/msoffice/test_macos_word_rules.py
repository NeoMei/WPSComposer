"""Paragraph-rule transport contracts; native execution is a separate gate."""
from copy import deepcopy
import importlib
import json
from pathlib import Path
import subprocess
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.errors import NativeWordError, NativeWordTimeoutError
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession


def api():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_rules')
    except ModuleNotFoundError:
        pytest.fail('paragraph rule implementation is missing')


# Nonempty terminal paragraph starts at 4 and ends at 9; space+CR adds two units.
ACK = [['paragraph-rule', 9, 4, 11, 10, 12, 2, 3, True, ' \r', '\r',
        True, True, True, True, True, True]]


def bind(monkeypatch, tmp_path, rows=ACK):
    session = MacWordSession()
    session.staging_root = tmp_path
    session._window_id = 77
    session._bound_path = '/private/task/document.docx'
    session._pending_heading = (1, 8, 1)
    session._observed_field_topology = ('old-field',)
    calls = []
    def run(args, **kwargs):
        source = Path(args[1]).read_text()
        calls.append((source, kwargs))
        assert not hasattr(session, '_observed_field_topology')
        return SimpleNamespace(returncode=0, stderr='', stdout=json.dumps(['WPSCOMPOSER_WORD_SESSION_OK', rows]))
    monkeypatch.setattr('skills.WPSComposer.scripts.msoffice.macos_word_session.subprocess.run', run)
    return session, calls


def test_success_commits_only_acknowledged_state(monkeypatch, tmp_path):
    m = api()
    session, calls = bind(monkeypatch, tmp_path)
    assert m.add_paragraph_horizontal_line(session) is None
    assert session._structural_changed and session._pending_heading is None
    assert not session._quarantined
    source, options = calls[0]
    assert 'window id 77' in source and '/private/task/document.docx' in source
    assert 0 < options['timeout'] <= 120


@pytest.mark.parametrize('state', ['_closed', '_read_only', '_quarantined', 'expired'])
def test_preflight_rejection_does_not_invalidate_state(monkeypatch, tmp_path, state):
    m = api()
    session, calls = bind(monkeypatch, tmp_path)
    if state == 'expired':
        session._bound_path = None
        session._deadline = 0
    else:
        setattr(session, state, True)
    with pytest.raises((ValueError, NativeWordError)):
        m.add_paragraph_horizontal_line(session)
    assert not calls and session._pending_heading == (1, 8, 1)
    assert session._observed_field_topology == ('old-field',)
    assert not session._structural_changed


def test_script_write_failure_before_submission_preserves_targets(monkeypatch, tmp_path):
    m = api()
    session, calls = bind(monkeypatch, tmp_path)
    session.staging_root = tmp_path / 'absent'
    with pytest.raises(OSError):
        m.add_paragraph_horizontal_line(session)
    assert not calls and session._pending_heading == (1, 8, 1)
    assert session._observed_field_topology == ('old-field',)
    assert not session._structural_changed and not session._quarantined


@pytest.mark.parametrize('bad', [[], ACK+ACK, [ACK[0][:-1]]] + [
    [ACK[0][:i] + [None] + ACK[0][i+1:]] for i in range(len(ACK[0]))
] + [[ACK[0][:i] + [1] + ACK[0][i+1:]] for i in (8,11,12,13,14,15,16)])
def test_malformed_or_partial_ack_never_reports_completion(monkeypatch, tmp_path, bad):
    m = api()
    session, calls = bind(monkeypatch, tmp_path, bad)
    with pytest.raises(NativeWordError) as error:
        m.add_paragraph_horizontal_line(session)
    assert error.value.code == 'NATIVE_WORD_QUARANTINED'
    assert error.value.diagnostic_path and Path(error.value.diagnostic_path).exists()
    assert calls and session._quarantined and session._retain_evidence
    assert not hasattr(session, '_observed_field_topology')


def test_ack_rejects_changed_prefix_geometry_and_missing_boundary(monkeypatch, tmp_path):
    m = api()
    for index, value in [(1, True), (2, 10), (3, 99), (4, 90), (5, 14), (6, 0), (7, 2), (8, False), (9, 'x\r'), (10, 'FOLLOWING\r'), (16, False)]:
        rows = deepcopy(ACK); rows[0][index] = value
        session, _ = bind(monkeypatch, tmp_path, rows)
        with pytest.raises(NativeWordError):
            m.add_paragraph_horizontal_line(session)


def test_empty_document_ack_has_no_extra_separator(monkeypatch, tmp_path):
    m = api()
    rows = [['paragraph-rule', 0, 0, 2, 1, 3, 1, 2, True, ' \r', '\r', True, True, True, True, True, True]]
    session, _ = bind(monkeypatch, tmp_path, rows)
    assert m.add_paragraph_horizontal_line(session) is None


def test_native_execution_failure_retains_primary_diagnostic(monkeypatch, tmp_path):
    m = api()
    session, _ = bind(monkeypatch, tmp_path)
    def fail(*args, **kwargs):
        return SimpleNamespace(returncode=1, stderr='Native rejected border (-2700)', stdout='')
    monkeypatch.setattr('skills.WPSComposer.scripts.msoffice.macos_word_session.subprocess.run', fail)
    with pytest.raises(NativeWordError) as error:
        m.add_paragraph_horizontal_line(session)
    assert error.value.code == 'NATIVE_WORD_EXECUTION_FAILED'
    assert error.value.diagnostic_path and session._quarantined


def test_guarded_fixture_requires_execute_and_rejects_existing_output(tmp_path):
    try:
        m = importlib.import_module('fixtures.microsoft_parity.macos_word_paragraph_rule')
    except ModuleNotFoundError:
        pytest.fail('guarded paragraph rule fixture is missing')
    assert m.main(['--output', str(tmp_path / 'absent')]) == 2
    assert not (tmp_path / 'absent').exists()
    with pytest.raises(FileExistsError):
        m.main(['--execute', '--output', str(tmp_path)])


def test_generated_rule_script_compiles_without_launch(tmp_path):
    m = api()
    source = 'use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n' + '\n'.join(m.paragraph_rule_commands(MacWordSession())) + '\nend tell\n'
    script = tmp_path / 'rule.applescript'; script.write_text(source)
    result = subprocess.run(['/usr/bin/osacompile', '-o', str(tmp_path / 'rule.scpt'), str(script)], capture_output=True, text=True, timeout=20)
    assert result.returncode == 0, result.stderr


def test_ack_rejects_previous_extra_separator_even_with_all_true_flags(monkeypatch, tmp_path):
    m = api()
    old = [['paragraph-rule', 9, 10, 12, 10, 13, 2, 4, True, ' \r', '\r', True, True, True, True, True, True]]
    session, _ = bind(monkeypatch, tmp_path, old)
    with pytest.raises(NativeWordError):
        m.add_paragraph_horizontal_line(session)


def test_compiler_emits_current_paragraph_format_before_insertion():
    code = '\n'.join(api().paragraph_rule_commands(MacWordSession()))
    # Compiler payload is the public native boundary: no paragraph separator;
    # baseline ordering determines which paragraph gets the border/inheritance.
    assert 'set content of boundaryRange' not in code
    assert code.index('set line style of ownBorder to line style single') < code.index('set content of insertionRange to " " & return')
    assert code.index('set alignment of paragraph format of ruleRange to align paragraph center') < code.index('set content of insertionRange to " " & return')
    assert "isEqualToString:rulePrefix" in code
    assert '(alignment of followingFormat is align paragraph center)' in code


def test_fixture_readback_requires_current_ruled_paragraph_and_centered_following():
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    valid = [['public-paragraph-rule', fixture.PREFIX+' \r', 'FOLLOWING\r', True, True, True, True, True, True, True, 6, 9, 12, 18]]
    assert fixture.native_valid(valid)
    for index in range(len(valid[0])):
        broken = deepcopy(valid); broken[0][index] = None
        assert not fixture.native_valid(broken)


@pytest.mark.parametrize('closed,changed,close_fails', [(False, False, False), (True, True, False), (True, False, True), (True, False, False)])
def test_fixture_cleanup_proves_owned_close_and_exact_sentinel_before_independent_close(monkeypatch, tmp_path, closed, changed, close_fails):
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    owned = SimpleNamespace(_closed=closed, _quarantined=False)
    before = ['sentinel', 'sentinel', False, 'hash-exact']
    observed = ['sentinel', 'sentinel', False, 'hash-CHANGED' if changed else 'hash-exact']
    events = []
    def inventory(output, label):
        events.append(label)
        return [observed] if label == 'after-owned-close' else []
    def close(output, name, token):
        events.append('independent-close')
        if close_fails:
            raise RuntimeError('guard refused close')
        return [['sentinel-closed', name]]
    monkeypatch.setattr(fixture, 'inventory', inventory, raising=False)
    monkeypatch.setattr(fixture, 'close_sentinel', close, raising=False)
    assert hasattr(fixture, 'close_after_owned'), 'owned-exit sentinel proof missing'
    report = {'checks':{}, 'inventory_before':[], 'sentinel_before':before}
    if not closed or changed or close_fails:
        with pytest.raises((RuntimeError, AssertionError)):
            fixture.close_after_owned(owned, tmp_path, report, 'sentinel', 'TOKEN')
    else:
        assert fixture.close_after_owned(owned, tmp_path, report, 'sentinel', 'TOKEN') is None
        assert report['checks']['owned_closed_with_sentinel_preserved']
        assert report['checks']['inventory_after_sentinel_close_preserved']
    if not closed or changed:
        assert 'independent-close' not in events
    if not closed:
        assert events == []
    if closed and not changed and not close_fails:
        assert events == ['after-owned-close', 'independent-close', 'after-sentinel-close']


def test_fixture_xml_checks_current_paragraph_border_and_inherited_center():
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    xml = ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           '<w:p><w:pPr><w:jc w:val="center"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="C0C0C0"/></w:pBdr></w:pPr>'
           '<w:r><w:t>'+fixture.PREFIX+' </w:t></w:r></w:p>'
           '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:jc w:val="center"/>'
           '<w:spacing w:before="120" w:after="180" w:line="360" w:lineRule="exact"/><w:ind w:firstLine="240"/></w:pPr>'
           '<w:r><w:t>FOLLOWING</w:t></w:r></w:p></w:body></w:document>')
    styles = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/></w:style></w:styles>'
    assert all(fixture.xml_checks(xml, styles).values())
    for old, new in [('w:sz="6"', 'w:sz="8"'), ('center', 'left'), ('BodyText', 'Normal'),
                     (fixture.PREFIX, fixture.PREFIX.lower()), ('w:after="180"','w:after="0"'),
                     ('</w:body>', '<w:p/></w:body>')]:
        assert not all(fixture.xml_checks(xml.replace(old, new), styles).values())


def test_fixture_records_real_unsaved_basename_and_rejects_invalid_preimage():
    import hashlib
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    assert hasattr(fixture, 'sentinel_preimage'), 'actual sentinel inventory row selection missing'
    token = 'TOKEN 中文😀'
    row = ['Document7', 'Document7', False, hashlib.sha256((token+'\r').encode()).hexdigest()]
    assert fixture.sentinel_preimage([['other', '/tmp/other.docx', True, 'irrelevant'], row], 'Document7', token) == row
    for rows in ([], [row, row], [row[:2]+[True,row[3]]], [row[:3]+['wronghash']], [['other']+row[1:]], [row[:2]+[0,row[3]]]):
        with pytest.raises(ValueError):
            fixture.sentinel_preimage(rows, 'Document7', token)


@pytest.mark.parametrize('rows', [ACK, []])
def test_public_session_forward_uses_bound_rule_ack_contract(monkeypatch, tmp_path, rows):
    import inspect
    from skills.WPSComposer.scripts.writer import WriterComposer
    assert hasattr(MacWordSession, 'add_paragraph_horizontal_line'), 'public paragraph rule forward missing'
    assert inspect.signature(MacWordSession.add_paragraph_horizontal_line) == inspect.signature(WriterComposer.add_paragraph_horizontal_line)
    session, calls = bind(monkeypatch, tmp_path, rows)
    if rows:
        assert session.add_paragraph_horizontal_line() is None
        assert session._structural_changed and session._pending_heading is None
    else:
        with pytest.raises(NativeWordError):
            session.add_paragraph_horizontal_line()
        assert session._quarantined
    assert len(calls) == 1 and not hasattr(session, '_observed_field_topology')


def test_fixture_resolves_actual_native_localized_style_id_and_rejects_wrong_style():
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    evidence = Path(__file__).resolve().parents[2] / 'docs/verification/microsoft-parity/macos-word-paragraph-rule/diagnosis-01'
    xml = (evidence/'after-following.document.xml').read_bytes()
    styles = (evidence/'after-following.styles.xml').read_bytes()
    assert all(fixture.xml_checks(xml, styles).values())
    # Actual native style ID ae must resolve to a paragraph style named Body Text.
    for wrong in (
        styles.replace(b'w:val="Body Text"', b'w:val="Normal"'),
        styles.replace(b'w:val="Body Text"', b'w:val="Body text"'),
        styles.replace(b'w:type="paragraph" w:styleId="ae"', b'w:type="character" w:styleId="ae"'),
        styles.replace(b'w:styleId="ae"', b'w:styleId="ae-missing"'),
        styles.replace(b'</w:styles>', b'<w:style w:type="paragraph" w:styleId="ae"><w:name w:val="Body Text"/></w:style></w:styles>'),
    ):
        assert wrong != styles
        assert not fixture.xml_checks(xml, wrong)['following_style_spacing_xml']


def test_fixture_native_style_readback_uses_exact_names_and_compiles(tmp_path):
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    code = '\n'.join(fixture.readback())
    assert "NSString" in code and 'isEqualToString:' in code
    assert 'name local of style of followingRange' in code
    assert 'name local of Word style (style body text) of boundDoc' in code
    source = 'use framework "Foundation"\nuse scripting additions\ntell application "/Applications/Microsoft Word.app"\n'+code+'\nend tell\n'
    script = tmp_path/'fixture-style.applescript';script.write_text(source)
    result = subprocess.run(['/usr/bin/osacompile','-o',str(tmp_path/'fixture-style.scpt'),str(script)],capture_output=True,text=True,timeout=20)
    assert result.returncode == 0, result.stderr


def test_fixture_accepts_retained_run02_pdf_rule_offline(tmp_path):
    import shutil
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    evidence = Path(__file__).resolve().parents[2]/'docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02'
    for name in ('paragraph-rule.docx', 'paragraph-rule.pdf'):
        shutil.copyfile(evidence/name, tmp_path/name)
    assert all(fixture.artifacts(tmp_path).values())


@pytest.mark.parametrize('change', [
    'wrong-color', 'transparent-fill', 'thick-box', 'short-box', 'vertical',
    'wrong-x', 'wrong-y', 'wrong-page', 'wrong-page-size', 'triangle',
    'extra-path', 'transparent-stroke', 'thick-stroke', 'diagonal-stroke', 'dashed-stroke',
])
def test_fixture_rule_detector_rejects_wrong_geometry_and_style(change):
    import fitz
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    rect = fitz.Rect(88.56, 96.96, 506.64, 97.68)
    drawing = {'type':'f', 'fill':(192/255,)*3, 'fill_opacity':1.0,
               'rect':rect, 'items':[('re', rect, -1)]}
    page, page_rect = 1, fitz.Rect(0, 0, 595.2, 841.92)
    if change == 'wrong-color': drawing['fill'] = (0.5,)*3
    if change == 'transparent-fill': drawing['fill_opacity'] = 0.2
    if change == 'thick-box': rect.y1 = 110
    if change == 'short-box': rect.x1 = 300
    if change == 'vertical': rect.x1, rect.y1 = 89.28, 500
    if change == 'wrong-x': rect.x0 += 25; rect.x1 += 25
    if change == 'wrong-y': rect.y0 += 25; rect.y1 += 25
    if change == 'wrong-page': page = 2
    if change == 'wrong-page-size': page_rect = fitz.Rect(0, 0, 200, 200)
    if change == 'triangle': drawing['items'] = [('l', rect.tl, rect.br)]
    if change == 'extra-path': drawing['items'].append(('re', rect, -1))
    if change.endswith('stroke'):
        drawing = {'type':'s', 'color':(192/255,)*3, 'stroke_opacity':1.0,
                   'width':0.75, 'dashes':'[] 0',
                   'items':[('l', fitz.Point(88.56, 97.32), fitz.Point(506.64, 97.32))]}
        if change == 'transparent-stroke': drawing['stroke_opacity'] = 0.0
        if change == 'thick-stroke': drawing['width'] = 5
        if change == 'diagonal-stroke': drawing['items'][0][2].y = 98
        if change == 'dashed-stroke': drawing['dashes'] = '[3 2] 0'
    assert not fixture.pdf_rule_matches(drawing, page, page_rect)


def test_fixture_rule_detector_accepts_horizontal_opaque_stroke():
    import fitz
    from fixtures.microsoft_parity import macos_word_paragraph_rule as fixture
    drawing = {'type':'s', 'color':(192/255,)*3, 'stroke_opacity':1.0, 'width':0.75,
               'dashes':'[] 0', 'items':[('l', fitz.Point(88.56, 97.32), fitz.Point(506.64, 97.32))]}
    assert fixture.pdf_rule_matches(drawing, 1, fitz.Rect(0, 0, 595.2, 841.92))
