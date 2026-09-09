"""Protect content until Mac Word can honor nonterminal table positions."""
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import document_api as api, office_engines as engines
from skills.WPSComposer.scripts.msoffice.edit_preflight import supports_edit_ops
from test_edit_preflight import mac_ms, recording_word


NONTERMINAL = ['start', {'before': 'paragraph:2'},
               {'after': 'paragraph:1'}, {'index': 2}]


def table_op(position):
    return {'op': 'insert', 'parent': 'body', 'type': 'table',
            'position': position, 'props': {'rows': 1, 'cols': 1, 'data': [['keep']]}}


@pytest.mark.parametrize('position', NONTERMINAL)
@pytest.mark.parametrize('engine', ['auto', 'msoffice'])
def test_later_nonterminal_table_rejects_entire_edit_before_open(
        mac_ms, monkeypatch, position, engine):
    opened = []
    def forbidden_open(*args, **kwargs):
        opened.append(True)
        raise AssertionError('unsafe batch reached native open')
    monkeypatch.setattr(api, 'open_document', forbidden_open)
    operations = iter([{'op': 'set', 'target': 'paragraph:1', 'text': 'do not write'},
                       table_op(position)])
    expected = engines.EngineUnavailableError if engine == 'auto' else ValueError
    with pytest.raises(expected):
        api.edit('/source.docx', engine=engine, ops=operations, atomic=False,
                 output='/separate-output.docx')
    assert opened == []


@pytest.mark.parametrize('position', NONTERMINAL)
@pytest.mark.parametrize('atomic', [True, False])
def test_later_nonterminal_table_rejects_existing_session_batch_before_any_write(
        mac_ms, recording_word, monkeypatch, position, atomic):
    session, native_calls = recording_word
    writes = []
    monkeypatch.setattr(session, 'apply_format_patch',
                        lambda *a, **k: writes.append(True) or {'accepted': ['text']})
    operations = iter([{'op': 'set', 'target': 'paragraph:1', 'text': 'do not write'},
                       table_op(position)])
    with pytest.raises(ValueError, match='Mac Word.*table.*end'):
        api.apply_ops(session, operations, atomic=atomic)
    assert writes == []
    assert native_calls == []


@pytest.mark.parametrize('position', NONTERMINAL)
def test_direct_nonterminal_table_rejects_before_native_transport(recording_word, position):
    session, native_calls = recording_word
    with pytest.raises(ValueError, match='Mac Word.*table.*end'):
        session.apply_structural_op(table_op(position))
    assert native_calls == []


@pytest.mark.parametrize('position', [None, 'end', 'omitted'])
def test_terminal_tables_keep_public_batch_execution(mac_ms, recording_word, position):
    session, native_calls = recording_word
    operation = table_op(position)
    if position == 'omitted':
        operation.pop('position')
    assert supports_edit_ops('writer', [operation], platform='darwin')
    reports = api.apply_ops(session, iter([operation]))
    assert reports[0]['ok'] is True
    assert len(native_calls) == 1


@pytest.mark.parametrize('platform,engine', [('win32', 'msoffice'), ('win32', 'wps'),
                                             ('darwin', 'wps')])
def test_non_mac_microsoft_batches_keep_existing_execution(monkeypatch, platform, engine):
    monkeypatch.setattr(engines.sys, 'platform', platform)
    applied = []
    composer = SimpleNamespace(kind='writer', engine=engine,
                               apply_structural_op=lambda op: applied.append(op) or {})
    operation = table_op({'before': 'paragraph:2'})
    reports = api.apply_ops(composer, iter([operation]))
    assert reports[0]['ok'] is True
    assert applied == [operation]


@pytest.mark.parametrize('position', NONTERMINAL)
def test_windows_capability_does_not_inherit_mac_table_limit(position):
    operation = table_op(position)
    for engine in ('wps', 'msoffice'):
        assert supports_edit_ops('writer', [operation], platform='win32', engine=engine)


@pytest.mark.parametrize('request_kind', ['patches-only', 'none', 'empty-generator'])
def test_explicit_edit_without_structural_ops_continues_to_open(mac_ms, monkeypatch, request_kind):
    opened = []
    def stop_at_open(*args, **kwargs):
        opened.append(kwargs['engine'])
        raise RuntimeError('reached expected opener')
    monkeypatch.setattr(api, 'open_document', stop_at_open)
    options = ({'patches': [{'target': 'paragraph:1', 'text': 'valid'}]}
               if request_kind == 'patches-only'
               else {'ops': None if request_kind == 'none' else iter(())})
    with pytest.raises(RuntimeError, match='reached expected opener'):
        api.edit('/source.docx', engine='msoffice', output='/separate-output.docx', **options)
    assert opened == ['msoffice']


@pytest.mark.parametrize('atomic', [True, False])
@pytest.mark.parametrize('mixed', [True, False])
def test_patch_override_nonterminal_table_rejects_before_open_or_output_preparation(
        mac_ms, monkeypatch, atomic, mixed):
    reached = []
    def forbidden_stage(*args, **kwargs):
        reached.append(True)
        raise AssertionError('unsafe patch batch reached document preparation')
    monkeypatch.setattr(api, 'open_document', forbidden_stage)
    monkeypatch.setattr(api, 'snapshot_artifact_state', forbidden_stage)
    first = {'target': 'paragraph:1', 'text': 'must not write'}
    patches = [first, table_op('start')] if mixed else [table_op('start')]
    operations = [table_op('end')] if mixed else []
    with pytest.raises(ValueError, match='Mac Word.*table.*end'):
        api.edit('/source.docx', engine='msoffice', patches=iter(patches),
                 ops=iter(operations), atomic=atomic)
    assert reached == []


@pytest.mark.parametrize('engine', ['auto', 'msoffice'])
def test_patch_override_and_ops_iterators_execute_once_in_order(
        mac_ms, recording_word, monkeypatch, engine):
    session, native_calls = recording_word
    monkeypatch.setattr(api, 'open_document', lambda *a, **kw: session)
    yielded = []
    def operations(label):
        yielded.append(label)
        operation = table_op('end')
        operation['props']['data'] = [[label]]
        yield operation
    result = api.edit('/source.docx', engine=engine,
                      patches=operations('from patch'), ops=operations('from ops'))
    assert yielded == ['from patch', 'from ops']
    assert [report['ok'] for report in result['ops']] == [True, True]
    assert len(native_calls) == 2
    assert any('from patch' in line for line in native_calls[0])
    assert any('from ops' in line for line in native_calls[1])
