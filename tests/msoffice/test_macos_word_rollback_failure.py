"""Pure checks for the guarded native failure fixture; never launch Office."""
import hashlib
import importlib.util
from pathlib import Path
import subprocess

import pytest

ROOT = next(parent for parent in Path(__file__).resolve().parents
            if (parent/'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py').is_file())
FIXTURE = ROOT/'fixtures/microsoft_parity/macos_word_rollback_failure.py'
if not FIXTURE.exists():
    FIXTURE = Path(__file__).with_name('macos_word_rollback_failure.py')


def load():
    assert FIXTURE.is_file(), 'guarded fixture is not implemented'
    spec = importlib.util.spec_from_file_location('rollback_failure_fixture', FIXTURE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_import_does_not_launch_native_or_create_outputs(monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail('import performed native execution or output creation')
    monkeypatch.setattr(subprocess, 'run', forbidden)
    monkeypatch.setattr(Path, 'mkdir', forbidden)
    module = load()
    assert module.ROOT == ROOT


def test_cli_without_execute_cannot_create_output_or_run_native(monkeypatch, tmp_path):
    module = load()
    def forbidden(*args, **kwargs):
        pytest.fail('unguarded native execution')
    monkeypatch.setattr(module, 'run', forbidden)
    output = tmp_path/'must-not-exist'
    with pytest.raises(SystemExit) as caught:
        module.main(['--output', str(output)])
    assert caught.value.code == 2 and not output.exists()


def actual_rollback_commands():
    from skills.WPSComposer.scripts.msoffice import macos_word_recovery as r
    header = ('checkpoint-state', 1, 6, 7, 1, 0, 7, hashlib.sha256(b'target').hexdigest(),
              0, 6, 'target', hashlib.sha256(b'target\r').hexdigest())
    state = (header, ('objects', 0, 0, 0, 0, 0), ('checkpoint-end',))
    saved = r.CheckpointSnapshot(
        coordinate=6, state=state, tracked_indexes_prefix=(),
        tracked_references_prefix=(), tracked_numbering_prefix=(),
        observed_field_topology=None,
    )
    field = ('field', 'main', 1, 'REF', ' REF probe ', 8, 15, 16, 19)
    table = ('table', 1, 6, 23, 1, 1)
    current_header = (*header[:2], 23, 24, *header[4:])
    current = (current_header, table, field, ('objects', 0, 0, 1, 1, 0), ('checkpoint-end',))
    return r.rollback_commands(saved, current, {'field':[field], 'table':[table], 'bookmark':[]})


def test_injection_changes_only_unique_real_field_delete_boundary():
    module = load()
    source = actual_rollback_commands()
    original = list(source)
    result = module.inject_after_field_delete(source)
    index = result.index('delete recoveryField')
    assert result[index+1] == 'error "WPSC_INJECTED_ROLLBACK_FAILURE"'
    assert result[:index+1] + result[index+2:] == source == original
    assert index < result.index('delete recoveryTable') < result.index('set content of rollbackRange to ""')
    assert sum('WPSC_INJECTED_ROLLBACK_FAILURE' in line for line in result) == 1


@pytest.mark.parametrize('source', [[], ['delete recoveryTable'],
    ['delete recoveryField', 'delete recoveryField'],
    ['delete recoveryField', 'error "WPSC_INJECTED_ROLLBACK_FAILURE"']])
def test_injection_rejects_missing_duplicate_or_preinjected_boundary(source):
    with pytest.raises(ValueError):
        load().inject_after_field_delete(source)


def test_temporary_compiler_wrapper_restores_real_function_even_on_failure():
    module = load()
    original = module.recovery.rollback_commands
    with pytest.raises(RuntimeError):
        with module.injected_compiler() as records:
            assert module.recovery.rollback_commands is not original
            raise RuntimeError('controlled preparation failure')
    assert module.recovery.rollback_commands is original and records == []


def test_independent_snapshot_is_exact_bound_readonly_and_never_rebinds_session(tmp_path):
    module = load()
    path = str(tmp_path / 'owned-document.docx')
    source = module.snapshot_source(path, 20)
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert 'owned-document.docx' in source and apple_string(path) in source
    assert 'isEqualToString:' in source and 'WPSC_DIAGNOSTIC_BINDING_CHANGED' in source
    assert 'if not application "Microsoft Word" is running' in source
    assert 'window wi of guardedDoc' in source and 'active window' not in source
    assert 'count fields of guardedDoc' in source and 'count tables of guardedDoc' in source
    assert 'close ' not in source and 'delete ' not in source and 'set content of ' not in source
    assert 'active document' not in source


def test_fixture_session_call_signatures_and_lazy_tracking_are_valid():
    import ast
    import inspect
    module = load()
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
    assert module.tracking_state(MacWordSession()) == ((), ())
    for node in ast.walk(ast.parse(FIXTURE.read_text())):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
            continue
        if not isinstance(node.func.value, ast.Name) or node.func.value.id != 'session':
            continue
        method = getattr(MacWordSession, node.func.attr)
        inspect.signature(method).bind(None, *[None for _ in node.args],
                                       **{keyword.arg: None for keyword in node.keywords})


def test_report_captures_full_shipped_digest_even_when_stopped_before_native(monkeypatch, tmp_path):
    module = load()
    digest = module.source_digest(ROOT)
    def stop(*args):
        raise RuntimeError('controlled stop before any native call')
    monkeypatch.setattr(module, 'inventory', stop)
    report = module.run(tmp_path/'read-only-report')
    assert report['status'] == 'FAIL_RETAINED'
    assert report['source_digest_before'] == report['source_digest_after'] == digest
    assert report['checks']['full_source_digest_unchanged']
    assert 'capability_checks' not in report


def test_diagnostic_ack_rejects_unknown_malformed_and_bool_count_rows():
    from copy import deepcopy
    module = load()
    valid = [['owned', '/tmp/owned.docx', 'owned.docx', False, 0, 1, 'a'*64, 'b'*64, [['owned.docx', None]]],
             ['table', 1, 20, 50, 'c'*64, True]]
    assert module.validate_snapshot(valid, '/tmp/owned.docx') == valid
    invalid = [[], [valid[0]], valid+[['unknown']], [['unknown'], *valid[1:]]]
    for index in range(9):
        broken = deepcopy(valid); broken[0][index] = None; invalid.append(broken)
    for index in (2, 4, 5):
        broken = deepcopy(valid); broken[0][index] = True; invalid.append(broken)
    for index in range(6):
        broken = deepcopy(valid); broken[1][index] = None; invalid.append(broken)
    for broken in invalid:
        with pytest.raises(ValueError):
            module.validate_snapshot(broken, '/tmp/owned.docx')


def test_file_owned_missing_window_id_uses_exact_native_document_identity(tmp_path):
    module = load()
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
    session = MacWordSession(); session._owns_doc = True
    session._private_path = tmp_path/'owned.docx'
    session._bind([['binding', None, str(session._private_path)]])
    assert session._window_id is None
    source = module.snapshot_source(session._bound_path, 20)
    assert 'active window' not in source and 'selection' not in source
    assert 'window wi of guardedDoc' in source
    assert 'name of diagnosticWindow' in source and 'id of diagnosticWindow' in source
    assert 'name of guardedDoc as text' in source
    assert 'WPSC_DIAGNOSTIC_BINDING_CHANGED' in source
    rows = [['owned', session._bound_path, 'owned.docx', False, 0, 0, 'a'*64, 'b'*64,
             [['owned.docx', None]]]]
    identity = {'document_name':'owned.docx', 'windows':[['owned.docx', None]]}
    assert module.validate_snapshot(rows, session._bound_path, identity) == rows
    assert session._window_id is None


@pytest.mark.parametrize('windows', [[['owned.docx', None], ['owned.docx', None]],
    [['owned.docx']], [['owned.docx', True]], [['owned.docx', 0]], [['owned.docx', '12']],
    [[12, None]], [[None, None], [None, None]]])
def test_identity_ack_rejects_incomplete_duplicate_or_invalid_windows(windows):
    rows = [['owned', '/tmp/owned.docx', 'owned.docx', False, 0, 0, 'a'*64, 'b'*64, windows]]
    with pytest.raises(ValueError):
        load().validate_snapshot(rows, '/tmp/owned.docx')


def test_identity_ack_preserves_missing_metadata_and_rejects_front_back_drift():
    module = load()
    rows = [['owned', '/tmp/owned.docx', 'owned.docx', False, 0, 0, 'a'*64, 'b'*64, [[None, None]]]]
    baseline = {'document_name':'owned.docx', 'windows':[[None, None]]}
    assert module.validate_snapshot(rows, '/tmp/owned.docx', baseline) == rows
    for wrong in ({'document_name':'OWNED.docx', 'windows':[[None, None]]},
                  {'document_name':'owned.docx', 'windows':[['new window', None]]}):
        with pytest.raises(ValueError):
            module.validate_snapshot(rows, '/tmp/owned.docx', wrong)
    with pytest.raises(ValueError):
        module.validate_snapshot(rows, '/different/owned.docx')
