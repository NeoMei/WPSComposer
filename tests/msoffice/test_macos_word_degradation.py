"""Direct degradation contracts at the native transport/acknowledgement boundary."""
from copy import deepcopy
from dataclasses import FrozenInstanceError
import hashlib
import importlib

import pytest

from skills.WPSComposer.scripts.msoffice.errors import NativeWordTimeoutError
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.writer import NativeWriterObjectError


def api():
    try:
        return importlib.import_module(
            'skills.WPSComposer.scripts.msoffice.macos_word_degradation'
        )
    except ModuleNotFoundError:
        pytest.fail('Direct Word degradation implementation is missing')


def range_ack(text, mode='inline', start=3):
    end = start + len(text.encode('utf-16-le')) // 2
    rows = [['degradation-range', mode, start, end, text, start+1, end+1, 0, 0],
            ['degradation-style', True, True, True]]
    if mode == 'block-range':
        rows.append(['degradation-paragraph', 0, 3, True, True])
    return rows


def table_ack(text, start=3):
    end = start + len(text.encode('utf-16-le')) // 2
    return [['degradation-table', start, start, end+2, text+'\r\x07\r\x07',
             start+1, end+3, 0, 1, 1, 1, start, end, text],
            ['degradation-style', True, True, True],
            ['degradation-paragraph', 0, 3, True, True],
            ['degradation-row', False]]


def checkpoint_rows(*, appended=False):
    prefix = '旧😀'
    end = 8 if appended else 3
    table = [['table', 1, 3, 8, 1, 1]] if appended else []
    return [['checkpoint-state', 1, end, end+1, 1, 0, end+1,
             hashlib.sha256(prefix.encode()).hexdigest(), 0, 3, prefix,
             hashlib.sha256((prefix+'\r').encode()).hexdigest()],
            *table, ['objects', 0, 0, len(table), 0, 0], ['checkpoint-end']]


class Transport:
    def __init__(self, session, responses):
        self.session = session
        self.responses = list(responses)
        self.calls = []
        self.states = []

    def __call__(self, lines, **kwargs):
        assert not self.session._quarantined
        self.calls.append(list(lines))
        self.states.append((self.session._structural_changed,
                            self.session._pending_heading,
                            getattr(self.session, '_observed_field_topology', None)))
        response = self.responses.pop(0)
        if isinstance(response, BaseException):
            raise response
        return deepcopy(response)


def bind(monkeypatch, responses):
    session = MacWordSession()
    session._pending_heading = (0, 3, 1)
    session._observed_field_topology = ('original',)
    transport = Transport(session, responses)
    monkeypatch.setattr(session, '_execute', transport)
    return session, transport


@pytest.mark.parametrize(('code', 'fallback', 'want'), [
    ('REF_FAILED', '缺少😀', '[REF_FAILED: 缺少😀]'),
    ('invalid code', 'fallback', '[DEGRADATION: fallback]'),
    ('REF_FAILED', '[REF_FAILED: ready]', '[REF_FAILED: ready]'),
    ('REF_FAILED', 0, '[REF_FAILED]'),
    ('REF_FAILED', 17, '[REF_FAILED: 17]'),
])
def test_inline_shared_display_utf16_and_semantic_range(monkeypatch, code, fallback, want):
    module = api()
    session, transport = bind(monkeypatch, [range_ack(want)])
    handle = module.add_inline_degradation(session, code, object(), fallback)
    assert (handle.Start, handle.End, handle.Text) == (
        3, 3 + len(want.encode('utf-16-le')) // 2, want)
    assert handle.session_id
    with pytest.raises(FrozenInstanceError):
        handle.End = 999
    assert transport.states == [(False, (0, 3, 1), ('original',))]
    assert session._structural_changed and session._pending_heading is None
    assert not hasattr(session, '_observed_field_topology')
    assert not session._quarantined


def test_inline_reuses_redaction_and_does_not_consume_message(monkeypatch):
    module = api()
    from skills.WPSComposer.scripts.writer import WriterComposer
    class Inert:
        def __str__(self):
            raise AssertionError('message must be ignored')
    text = 'file /Users/alice/private/key.txt token=sk-secretvalue'
    want = WriterComposer._degradation_display('REF_FAILED', text, inline=True)
    session, transport = bind(monkeypatch, [range_ack(want)])
    handle = module.add_inline_degradation(session, 'REF_FAILED', Inert(), text)
    assert handle.Text == want and '/Users/alice' not in handle.Text
    assert '/Users/alice' not in '\n'.join(transport.calls[0])


def test_block_single_cell_ack_returns_box_range(monkeypatch):
    module = api()
    text = '[TABLE_FAILED] 回退😀'
    session, transport = bind(monkeypatch, [checkpoint_rows(), table_ack(text)])
    box = module.add_degradation_notice(session, 'TABLE_FAILED', None, '回退😀')
    assert box.table_index == 1
    assert box.Range.Start == 3 and box.Range.Text == text+'\r\x07\r\x07'
    assert box.Range.End == 3+len((text+'\r\x07').encode('utf-16-le'))//2
    assert transport.states == [(False, (0, 3, 1), ('original',))]*2
    assert session._structural_changed and session._pending_heading is None


@pytest.mark.parametrize('placement', [None, '', 'INLINE', 0, [], {'x': 'inline'}])
def test_every_non_inline_placement_selects_block(monkeypatch, placement):
    module = api()
    session, transport = bind(monkeypatch, [checkpoint_rows(), table_ack('[NOTICE] x')])
    box = module.add_degradation_notice(session, 'NOTICE', None, 'x', placement)
    assert box.table_index == 1 and len(transport.calls) == 2


def test_exact_inline_placement_delegates_without_checkpoint(monkeypatch):
    module = api()
    session, transport = bind(monkeypatch, [range_ack('[NOTICE: x]')])
    result = module.add_degradation_notice(session, 'NOTICE', None, 'x', 'inline')
    assert result.Text == '[NOTICE: x]' and len(transport.calls) == 1


@pytest.mark.parametrize('bad', [[], [['ok']], range_ack('[NOTICE: x]')+[['extra']]])
def test_malformed_inline_ack_quarantines_without_committing_state(monkeypatch, bad):
    module = api()
    session, transport = bind(monkeypatch, [bad])
    with pytest.raises(NativeWriterObjectError) as caught:
        module.add_inline_degradation(session, 'NOTICE', None, 'x')
    assert caught.value.code == 'DEGRADATION_INSERT_FAILED'
    assert session._quarantined and session._retain_evidence
    assert not session._structural_changed and session._pending_heading == (0, 3, 1)
    assert session._observed_field_topology == ('original',)
    assert not hasattr(session, '_degradation_session_id')
    with pytest.raises(Exception):
        module.add_inline_degradation(session, 'NOTICE', None, 'x')
    assert len(transport.calls) == 1


@pytest.mark.parametrize(('row', 'column', 'value'), [
    (0, 2, True), (0, 3, 999), (0, 4, 'wrong'), (0, 5, 100),
    (0, 6, 1000), (0, 7, -1), (0, 8, 1), (1, 1, 1), (1, 2, False),
])
def test_range_ack_checks_native_bounds_text_types_and_style(monkeypatch, row, column, value):
    module = api()
    ack = range_ack('[NOTICE: x]')
    ack[row][column] = value
    session, _ = bind(monkeypatch, [ack])
    with pytest.raises(NativeWriterObjectError):
        module.add_inline_degradation(session, 'NOTICE', None, 'x')
    assert session._quarantined


def test_table_failure_uses_real_acknowledged_rollback_before_range_fallback(monkeypatch):
    module = api()
    before = checkpoint_rows()
    session, transport = bind(monkeypatch, [
        before, [['degradation-table-failed', 3]], checkpoint_rows(appended=True),
        [['rollback-ack', [], [['table', 1, 3, 8, 1, 1]]], *before],
        range_ack('[NOTICE] x', 'block-range'),
    ])
    box = module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert box.table_index is None and box.Range.Text == '[NOTICE] x'
    assert len(transport.calls) == 5
    assert 'delete recoveryTable' in transport.calls[3]
    assert transport.states[:4] == [(False, (0, 3, 1), ('original',))]*4
    assert transport.states[4][0] is True  # Only acknowledged rollback has committed.


def test_failed_rollback_stops_fallback_and_retains_evidence(monkeypatch):
    module = api()
    session, transport = bind(monkeypatch, [
        checkpoint_rows(), [['degradation-table-failed', 3]],
        checkpoint_rows(appended=True), [['rollback-ack', [], []]],
    ])
    with pytest.raises(NativeWriterObjectError) as caught:
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert caught.value.code == 'LOCAL_MUTATION_ROLLBACK_FAILED'
    assert len(transport.calls) == 4 and session._quarantined


@pytest.mark.parametrize('bad', [[['degradation-table-failed', True]],
                                [['degradation-table-failed', 4]],
                                [['degradation-table-failed', 3], ['extra']], []])
def test_malformed_table_failure_cannot_trigger_rollback_or_fallback(monkeypatch, bad):
    module = api()
    session, transport = bind(monkeypatch, [checkpoint_rows(), bad])
    with pytest.raises(NativeWriterObjectError):
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert len(transport.calls) == 2 and session._quarantined
    assert not session._structural_changed


@pytest.mark.parametrize(('row', 'column', 'value'), [
    (0, 1, 4), (0, 2, -1), (0, 3, 0), (0, 4, 'wrong'), (0, 7, True),
    (0, 8, 2), (0, 9, 2), (0, 10, 2), (0, 11, -1), (0, 13, 'wrong'),
    (1, 3, False), (2, 1, 2), (2, 4, False), (3, 1, True),
])
def test_bad_table_success_never_falls_back(monkeypatch, row, column, value):
    module = api()
    ack = table_ack('[NOTICE] x')
    ack[row][column] = value
    session, transport = bind(monkeypatch, [checkpoint_rows(), ack])
    with pytest.raises(NativeWriterObjectError):
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert len(transport.calls) == 2 and session._quarantined


@pytest.mark.parametrize('block', [False, True])
@pytest.mark.parametrize('error', [NativeWordTimeoutError(), RuntimeError('private detail'), KeyboardInterrupt()])
def test_uncertain_transport_or_cancellation_never_retries(monkeypatch, block, error):
    module = api()
    replies = [checkpoint_rows(), error] if block else [error]
    session, transport = bind(monkeypatch, replies)
    call = module.add_degradation_notice if block else module.add_inline_degradation
    with pytest.raises(KeyboardInterrupt if isinstance(error, KeyboardInterrupt) else NativeWriterObjectError):
        call(session, 'NOTICE', None, 'x')
    assert len(transport.calls) == len(replies)
    assert session._quarantined and not session._structural_changed
    assert session._retain_evidence


@pytest.mark.parametrize('text', ['bad\x00value', '\ud800', 'x'*1000001],
                         ids=['control', 'surrogate', 'oversized'])
@pytest.mark.parametrize('block', [False, True])
def test_native_unsafe_display_rejected_before_checkpoint_or_mutation(monkeypatch, text, block):
    module = api()
    session, transport = bind(monkeypatch, [])
    call = module.add_degradation_notice if block else module.add_inline_degradation
    with pytest.raises(ValueError):
        call(session, 'NOTICE', None, text)
    assert transport.calls == [] and not session._structural_changed


@pytest.mark.parametrize('condition', ['_read_only', '_closed', '_quarantined'])
def test_local_preflight_rejects_before_native_calls(monkeypatch, condition):
    module = api()
    session, transport = bind(monkeypatch, [])
    setattr(session, condition, True)
    with pytest.raises(Exception):
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert transport.calls == [] and not session._structural_changed


def test_failed_fallback_is_fatal_without_another_attempt(monkeypatch):
    module = api()
    before = checkpoint_rows()
    session, transport = bind(monkeypatch, [
        before, [['degradation-table-failed', 3]], before,
        [['rollback-ack', [], []], *before], [],
    ])
    with pytest.raises(NativeWriterObjectError) as caught:
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert caught.value.code == 'DEGRADATION_INSERT_FAILED'
    assert len(transport.calls) == 5 and session._quarantined
    assert not hasattr(session, '_degradation_session_id')


def test_inline_preserves_existing_reference_tracking(monkeypatch):
    module = api()
    session, _ = bind(monkeypatch, [range_ack('[NOTICE: x]')])
    references = [('old-ref', ' REF target ')]
    indexes = [('old-toc', ' TOC ')]
    session._tracked_references = references
    session._tracked_indexes = indexes
    module.add_inline_degradation(session, 'NOTICE', None, 'x')
    assert session._tracked_references is references
    assert session._tracked_indexes is indexes


def test_failed_checkpoint_cannot_submit_notice(monkeypatch):
    module = api()
    session, transport = bind(monkeypatch, [[]])
    with pytest.raises(NativeWriterObjectError) as caught:
        module.add_degradation_notice(session, 'NOTICE', None, 'x')
    assert caught.value.code == 'LOCAL_MUTATION_CHECKPOINT_FAILED'
    assert len(transport.calls) == 1 and session._quarantined
    assert not session._structural_changed


def test_public_session_inline_forward_runs_real_helper(monkeypatch):
    session, transport = bind(monkeypatch, [range_ack('[NOTICE: 中文😀]')])
    handle = session.add_inline_degradation('NOTICE', object(), '中文😀')
    assert handle.Text == '[NOTICE: 中文😀]' and handle.Start == 3
    assert session._structural_changed and len(transport.calls) == 1


def test_public_session_block_forward_preserves_placement_and_return(monkeypatch):
    session, transport = bind(monkeypatch, [checkpoint_rows(), table_ack('[NOTICE] x')])
    handle = session.add_degradation_notice('NOTICE', object(), 'x', None)
    assert handle.Range.Text == '[NOTICE] x\r\x07\r\x07' and handle.table_index == 1
    assert session._structural_changed and len(transport.calls) == 2


def test_native_fixture_requires_execute_without_creating_output(tmp_path):
    import subprocess
    import sys
    from pathlib import Path
    fixture = Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_degradation.py'
    result = subprocess.run([sys.executable, str(fixture), '--output', str(tmp_path/'out')],
                            capture_output=True, text=True)
    assert result.returncode == 2 and '--execute required' in result.stderr
    assert not (tmp_path/'out').exists()


@pytest.mark.parametrize('fault', [None, 'notice-color', 'following-italic', 'row-split'])
def test_saved_ooxml_gate_checks_notice_style_and_normal_tail(tmp_path, fault):
    import zipfile
    try:
        from fixtures.microsoft_parity.macos_word_degradation import inspect_ooxml
    except ModuleNotFoundError:
        pytest.fail('Native degradation fixture is missing')
    notice_color = '000000' if fault == 'notice-color' else '9C0006'
    following_italic = '<w:i/>' if fault == 'following-italic' else ''
    cant_split = '' if fault == 'row-split' else '<w:cantSplit/>'
    xml = f'''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:tbl><w:tr><w:trPr>{cant_split}</w:trPr><w:tc><w:p>
    <w:pPr><w:spacing w:before="0" w:after="60"/><w:keepLines/><w:outlineLvl w:val="9"/></w:pPr>
    <w:r><w:rPr><w:i/><w:color w:val="{notice_color}"/><w:shd w:fill="FCE8E6"/></w:rPr><w:t>[NOTICE] 中文😀</w:t></w:r>
    </w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr>{following_italic}</w:rPr><w:t>NORMAL-TAIL</w:t></w:r></w:p>
    </w:body></w:document>'''
    path = tmp_path/'evidence.docx'
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('word/document.xml', xml)
    if fault:
        with pytest.raises(AssertionError):
            inspect_ooxml(path, '[NOTICE] 中文😀', 'NORMAL-TAIL', block=True, table=True)
    else:
        assert inspect_ooxml(path, '[NOTICE] 中文😀', 'NORMAL-TAIL', block=True, table=True)['table_count'] == 1


def test_public_block_activates_exact_owned_window_before_table_creation(monkeypatch):
    """Word -2710 on inactive document must not be mistaken for normal fallback."""
    module = api()
    session, transport = bind(monkeypatch, [checkpoint_rows(), table_ack('[NOTICE] x')])
    native = session._execute
    def inactive_window_transport(lines, **kwargs):
        if any(line.startswith('set noticeTable to make new table') for line in lines):
            creation = next(i for i, line in enumerate(lines)
                            if line.startswith('set noticeTable to make new table'))
            if 'activate object boundWindow' not in lines[:creation]:
                raise RuntimeError('Native Word -2710: inactive owned document')
        return native(lines, **kwargs)
    monkeypatch.setattr(session, '_execute', inactive_window_transport)
    box = session.add_degradation_notice('NOTICE', None, 'x')
    assert box.table_index == 1 and box.Range.Text == '[NOTICE] x\r\x07\r\x07'
    assert not session._quarantined and len(transport.calls) == 2


@pytest.mark.parametrize(('position', 'table_start'), [(0, 0), (11, 12)])
def test_native_word_table_end_markers_do_not_count_as_literal_utf16(monkeypatch, position, table_start):
    """Word 16.112.3 geometry-diagnosis-01: CR/BEL expands native terminators."""
    module = api()
    display = '[GEOMETRY_NOTICE] 中文😀'
    ack = [['degradation-table', position, table_start, table_start+24,
            display+'\r\x07\r\x07', position+1, table_start+25,
            0, 1, 1, 1, table_start, table_start+22, display],
           ['degradation-style', True, True, True],
           ['degradation-paragraph', 0, 3, True, True],
           ['degradation-row', False]]
    assert module._table_ack(ack, display, position) == (
        table_start, table_start+24, display+'\r\x07\r\x07', 1)


@pytest.mark.parametrize(('case', 'valid'), [
    ('native-cell-and-table', True), ('table-only', True),
    ('run-auto', False), ('run-clear', False), ('run-other', False),
    ('paragraph-auto', False), ('cell-other', False),
    ('unrelated-table', False),
])
def test_ooxml_notice_shading_respects_exact_ancestor_and_explicit_override(tmp_path, case, valid):
    """run02 Word XML persists notice background on tblPr/tcPr, not rPr."""
    import zipfile
    from fixtures.microsoft_parity.macos_word_degradation import inspect_ooxml
    native_shading = '<w:shd w:val="clear" w:color="auto" w:fill="FCE8E6"/>'
    overrides = {
        'run-auto': '<w:shd w:val="clear" w:fill="auto"/>',
        'run-clear': '<w:shd w:val="clear"/>',
        'run-other': '<w:shd w:val="clear" w:fill="FFFFFF"/>',
    }
    run_shading = overrides.get(case, '')
    paragraph_shading = '<w:shd w:fill="auto"/>' if case == 'paragraph-auto' else ''
    cell_shading = '' if case == 'table-only' else native_shading
    if case == 'cell-other':
        cell_shading = '<w:shd w:fill="FFFFFF"/>'
    notice = f'''<w:p><w:pPr><w:spacing w:after="60"/><w:keepLines/>{paragraph_shading}</w:pPr>
    <w:r><w:rPr><w:i/><w:color w:val="9C0006"/>{run_shading}</w:rPr><w:t>[NOTICE] 中文😀</w:t></w:r></w:p>'''
    cell_body = '<w:p><w:r><w:t>unrelated</w:t></w:r></w:p>' if case == 'unrelated-table' else notice
    table = f'''<w:tbl><w:tblPr>{native_shading}</w:tblPr><w:tr><w:trPr><w:cantSplit/></w:trPr>
    <w:tc><w:tcPr>{cell_shading}</w:tcPr>{cell_body}</w:tc></w:tr></w:tbl>'''
    body = (notice if case == 'unrelated-table' else '')+table
    xml = f'''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    {body}<w:p><w:r><w:t>NORMAL-TAIL</w:t></w:r></w:p></w:body></w:document>'''
    path = tmp_path/'native-shading.docx'
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('word/document.xml', xml)
    if valid:
        assert inspect_ooxml(path, '[NOTICE] 中文😀', 'NORMAL-TAIL', block=True, table=True)['table_count'] == 1
    else:
        with pytest.raises(AssertionError, match='shading'):
            inspect_ooxml(path, '[NOTICE] 中文😀', 'NORMAL-TAIL', block=True, table=True)
