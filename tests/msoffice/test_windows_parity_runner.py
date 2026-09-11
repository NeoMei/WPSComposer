"""Pure checks for the explicit Windows acceptance runner, never Office."""
from __future__ import annotations

import importlib.util
import io
import json
from pathlib import Path
import subprocess
import sys
import time
from types import SimpleNamespace

import pytest

SOURCE = Path(__file__).resolve().parents[2] / 'fixtures/microsoft_parity/windows_native.py'


def runner():
    assert SOURCE.exists(), 'Windows public native acceptance runner is missing'
    spec = importlib.util.spec_from_file_location('windows_parity_runner', SOURCE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_timeout_kills_only_launched_python_and_retains_partial_logs(tmp_path):
    module = runner()
    result = module.run_bounded([sys.executable, '-u', '-c',
        'import time; print("owned document bound", flush=True); time.sleep(30)'], tmp_path, timeout=.2)
    assert result['status'] == 'timeout'
    assert result['python_pid'] > 0
    assert result['office_termination_attempted'] is False
    assert result['recovery_required'] is True
    assert 'owned document bound' in (tmp_path / 'stdout.log').read_text()
    assert json.loads((tmp_path / 'outer-timeout.json').read_text())['python_pid'] == result['python_pid']


@pytest.mark.parametrize('timeout', [0, -1, True, float('nan'), 4000])
def test_invalid_deadline_rejected_before_output_or_process(tmp_path, timeout):
    module = runner()
    with pytest.raises(ValueError):
        module.run_probe(tmp_path, tmp_path / 'new', timeout=timeout)
    assert not (tmp_path / 'new').exists()


def test_existing_output_never_overwritten(tmp_path):
    module = runner()
    marker = tmp_path / 'keep'
    marker.write_text('original')
    with pytest.raises(FileExistsError):
        module.run_probe(tmp_path, tmp_path)
    assert marker.read_text() == 'original'


def test_successful_child_without_evidence_is_failure_and_fault_is_opt_in(tmp_path):
    module = runner()
    plugin = tmp_path / 'plugin'
    plugin.mkdir()
    calls = []
    def execute(command, output, *, timeout):
        calls.append(command)
        return {'status': 'completed', 'returncode': 0}
    report = module.run_probe(plugin, tmp_path / 'run', kinds=['writer'], execute=execute)
    assert not report['passed']
    assert report['fault_timeout'] == 'not_requested'
    assert '--fault-timeout' not in calls[0]
    assert (tmp_path / 'run' / 'runner-source.py').read_bytes() == SOURCE.read_bytes()


def test_runtime_requires_matching_native_identity_and_final_close_ack(tmp_path):
    module = runner()
    identity = {'office_pid': 12, 'office_executable': r'C:\Office\WINWORD.EXE',
                'staging_path': str(tmp_path / 'private'), 'attached': False}
    rows = [dict(protocol=1, id=1, status='ok', value={'kind':'writer'}, identity=identity),
            dict(protocol=1, id=2, status='ok', value={'closed':True}, identity=identity)]
    (tmp_path / 'worker-stdout.log').write_text(''.join(json.dumps(r)+'\n' for r in rows))
    (tmp_path / 'last-request.json').write_text(json.dumps({'id':2, 'method':'close'}))
    assert module.verify_proxy_runtime(tmp_path, 'writer', returncode=0)['close_ack'] is True
    rows[-1]['identity'] = dict(identity, office_pid=13)
    (tmp_path / 'worker-stdout.log').write_text(''.join(json.dumps(r)+'\n' for r in rows))
    with pytest.raises(AssertionError, match='identity'):
        module.verify_proxy_runtime(tmp_path, 'writer', returncode=0)
    rows[-1]['identity'] = identity
    rows[-1]['value'] = []
    (tmp_path / 'worker-stdout.log').write_text(''.join(json.dumps(r)+'\n' for r in rows))
    with pytest.raises(AssertionError, match='close'):
        module.verify_proxy_runtime(tmp_path, 'writer', returncode=0)


def test_excel_reopen_requires_exact_cells_and_formula_not_any_fifteen():
    module = runner()
    snap = {'sheets': [{'index':1, 'name':'Data', 'cells':[
        {'address':'$A$1','value':'Native Excel'}, {'address':'$A$2','value':5},
        {'address':'$B$2','value':15,'formula':'=A2*3'}]},
        {'index':2,'name':'Details', 'cells':[{'address':'$A$1','value':'Typed worksheet'}]}]}
    assert module.check_snapshot('sheet', snap, edited=True)
    snap['sheets'][0]['cells'][2]['address'] = '$C$2'
    with pytest.raises(AssertionError):
        module.check_snapshot('sheet', snap, edited=True)


@pytest.mark.parametrize('changed', ['content_sha256', 'saved', 'full_name', 'identity', 'extra_document'])
def test_sentinel_baseline_change_is_never_a_pass(changed):
    module = runner()
    before = [{'identity':'token', 'full_name':'unsaved', 'saved':False, 'content_sha256':'abc'}]
    after = [dict(before[0])]
    if changed == 'extra_document':
        after.append(dict(before[0], identity='other'))
    else:
        after[0][changed] = 'changed'
    with pytest.raises(AssertionError, match='baseline'):
        module.assert_baseline(before, after)


def test_office_free_import_and_cli_help():
    module = runner()
    result = subprocess.run([sys.executable, str(SOURCE), '--help'], capture_output=True, text=True, timeout=10)
    assert result.returncode == 0
    assert '--plugin' in result.stdout and '--checkout' in result.stdout
    assert '--fault-timeout' in result.stdout
    assert not any(name.startswith(('win32com', 'pythoncom')) for name in vars(module))


def test_native_case_source_change_invalidates_all_successful_checks(tmp_path, monkeypatch):
    module = runner()
    monkeypatch.setattr(module.sys, 'platform', 'win32')
    monkeypatch.setattr(module, 'import_api', lambda path: SimpleNamespace(__file__=str(tmp_path / 'api.py')))
    values = iter(['before', 'after'])
    monkeypatch.setattr(module, 'bound_source_digest', lambda path: next(values), raising=False)
    class Sentinels:
        def __init__(self, *args): pass
        def start(self): pass
        def verify(self, **kwargs): pass
        def close(self, **kwargs): return True
    monkeypatch.setattr(module, 'NativeSentinels', Sentinels)
    monkeypatch.setattr(module, 'standard_case', lambda *args: dict.fromkeys(module.REQUIRED, True))
    report = module.native_case(tmp_path, tmp_path, 'writer', 100)
    assert report['passed'] is False
    assert report['platform'] == 'win32' and report['engine'] == 'msoffice'
    assert report['source_digest'] == 'before' and report['source_digest_after'] == 'after'
    assert report['checks']['source_unchanged'] is False


def test_populate_arguments_match_actual_composer_signatures():
    module = runner()
    import inspect
    from skills.WPSComposer.scripts.writer import WriterComposer
    from skills.WPSComposer.scripts.sheet import SheetComposer
    from skills.WPSComposer.scripts.slide import SlideComposer
    from skills.WPSComposer.scripts.msoffice.windows_session_proxy import NativeSessionHandle
    class Recorder:
        def __init__(self, cls): self.cls = cls
        def __getattr__(self, name):
            def call(*args, **kwargs):
                inspect.signature(getattr(self.cls, name)).bind(self, *args, **kwargs)
                if name == 'apply_format_patch' and self.cls is WriterComposer:
                    assert not (isinstance(args[0], NativeSessionHandle) and args[0].type == 'table'), 'Writer formatting accepts a table cell, not a whole table'
                kind = {'select_sheet':'sheet', 'add_sheet':'sheet', 'add_blank_slide':'slide',
                        'add_textbox':'shape', 'add_floating_textbox':'shape', 'add_table':'table'}.get(name)
                if kind:
                    handle = NativeSessionHandle(self, '0'*32, kind)
                    return (handle, 1) if name == 'add_blank_slide' else handle
                return {'moved':True, 'to_slide':1}
            return call
    for kind, cls in [('writer', WriterComposer), ('sheet', SheetComposer), ('slide', SlideComposer)]:
        module.populate(Recorder(cls), kind)


def test_installed_smoke_readback_checks_exact_targets_and_cached_formula():
    path = SOURCE.with_name('installed_public.py')
    spec = importlib.util.spec_from_file_location('installed_parity_smoke', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    assert hasattr(module, 'verify_readback'), 'installed smoke still uses broad JSON/any-cell matches'
    word = {'paragraphs':[{'index':1,'text':'Installed Word edited'}, {'index':2,'text':'other'}]}
    slide = {'slides':[{'index':1,'shapes':[{'index':1,'text':'Installed PowerPoint edited'}]}]}
    sheet = {'sheets':[{'index':1,'cells':[{'address':'$A$2','value':5},
                     {'address':'$B$2','value':15,'formula':'=A2*3'}]}]}
    assert module.verify_readback('writer', word, 'paragraph:1')
    assert module.verify_readback('slide', slide, 'slide:1/shape:1')
    assert module.verify_readback('sheet', sheet, 'sheet:1/cell:A2')
    word['paragraphs'].reverse()
    word['paragraphs'][1]['index'] = 3
    assert not module.verify_readback('writer', word, 'paragraph:1')
    slide['slides'][0]['shapes'][0]['index'] = 2
    assert not module.verify_readback('slide', slide, 'slide:1/shape:1')
    sheet['sheets'][0]['cells'][1]['address'] = '$C$2'
    assert not module.verify_readback('sheet', sheet, 'sheet:1/cell:A2')


def test_timeout_quarantine_retains_exact_component_only(tmp_path):
    module = runner()
    root = tmp_path / 'writer'
    root.mkdir()
    marker = root / 'native-office.quarantine.json'
    module.retain_quarantine(root, {'python_pid':12})
    original = marker.read_bytes()
    module.retain_quarantine(root, {'python_pid':99})
    assert marker.read_bytes() == original
    assert json.loads(original)['office_termination_attempted'] is False
    link = tmp_path / 'link'
    link.symlink_to(root, target_is_directory=True)
    with pytest.raises(ValueError):
        module.retain_quarantine(link, {})


def test_pdf_gate_reads_content_and_rejects_valid_but_empty_document(tmp_path):
    module = runner()
    from pypdf import PdfWriter
    path = tmp_path / 'blank.pdf'
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=100)
    with path.open('wb') as stream: writer.write(stream)
    with pytest.raises(AssertionError, match='PDF missing'):
        module.pdf_check(path, ['Native Word'])


def test_fault_evidence_cannot_pass_without_source_binding_or_cleanup(tmp_path):
    module = runner()
    plugin = tmp_path / 'plugin'
    plugin.mkdir()
    calls = []
    def execute(command, output, *, timeout):
        fault = '--fault-timeout' in command
        calls.append(fault)
        checks = ({'deadline_failure':True, 'sentinels_preserved':True, 'baseline_preserved':True}
                  if fault else dict.fromkeys(module.REQUIRED, True))
        (output / 'case-report.json').write_text(json.dumps({'passed':True, 'checks':checks,
            'platform':'win32', 'engine':'msoffice', 'component':'writer',
            'source_digest':'x'*64, 'source_digest_after':'x'*64}))
        return {'status':'completed', 'returncode':0}
    report = module.run_probe(plugin, tmp_path / 'run', kinds=['writer'], fault_timeout=True, execute=execute)
    assert calls == [False, True] and not report['passed']
    assert not report['cases']['writer-fault']['passed']


def test_child_check_booleans_do_not_override_wrong_engine_provenance(tmp_path):
    module = runner()
    plugin = tmp_path / 'plugin'
    plugin.mkdir()
    def execute(command, output, *, timeout):
        (output / 'case-report.json').write_text(json.dumps({'passed':True,
            'checks':dict.fromkeys(module.REQUIRED, True), 'platform':'win32',
            'engine':'wps', 'component':'writer', 'source_digest':'a'*64,
            'source_digest_after':'a'*64}))
        return {'status':'completed', 'returncode':0}
    assert not module.run_probe(plugin, tmp_path / 'run', kinds=['writer'], execute=execute)['passed']


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_generated_gate_requires_source_text_and_exact_table_values(kind):
    module = runner()
    cells = [dict(row=r,column=c,text=text) for r,c,text in
             [(1,1,'Item'),(1,2,'Value'),(2,1,'Native row'),(2,2,'42')]]
    snapshots = {
        'writer':{'paragraphs':[{'text':'Generated Native'}, {'text':'Generated native content.'}], 'tables':[{'cells':cells}]},
        'sheet':{'sheets':[{'name':'Generated Native', 'cells':[
            {'address':address, 'value':value} for address,value in
            [('A1','Item'),('B1','Value'),('A2','Native row'),('B2','42')]]}]},
        'slide':{'slides':[{'shapes':[{'text':'Generated Native'}, {'text':'Generated native content.'},
                                    {'table':{'cells':cells}}]}]},
    }
    snapshot = snapshots[kind]
    assert module.check_generated_snapshot(kind, snapshot)
    if kind == 'sheet': snapshot['sheets'][0]['cells'][-1]['value'] = 'wrong'
    else: cells[-1]['text'] = 'wrong'
    with pytest.raises(AssertionError):
        module.check_generated_snapshot(kind, snapshot)


def test_sentinel_noop_close_cannot_be_reported_closed(tmp_path):
    module = runner()
    state = module.NativeSentinels('sheet', tmp_path)
    doc = SimpleNamespace(Close=lambda **kwargs: None)
    app = object()
    state.deps = SimpleNamespace(identity=lambda obj: obj, uninitialize=lambda:None)
    state.composer = SimpleNamespace(_doc=doc, _app=app, _com_initialized=False,
        _verify_application=lambda:None)
    state.pythoncom = SimpleNamespace(CoUninitialize=lambda:None)
    state.apps = [app]
    token = state._token(doc)
    record = {'identity':token}
    state.owned, state.expected = [(doc, token)], [record]
    state._record = lambda value: record
    state.snapshot = lambda value: [record]  # COM Close returned but document remains.
    with pytest.raises(AssertionError, match='sentinel.*still open'):
        state.close(preserve_host=True)
    assert state.composer._doc is doc


def test_installed_capture_requires_postclose_ack_and_preserves_it(tmp_path):
    spec = importlib.util.spec_from_file_location('installed_capture', SOURCE.with_name('installed_public.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    job, output = tmp_path / 'job', tmp_path / 'output'
    job.mkdir(); output.mkdir()
    response = {'protocol':1, 'id':9, 'status':'ok', 'value':{'closed':True}}
    (job / 'worker-stdout.log').write_text(json.dumps(response)+'\n')
    (job / 'last-request.json').write_text(json.dumps({'id':9, 'method':'close'}))
    session = SimpleNamespace(_child=SimpleNamespace(poll=lambda:0))
    closed, evidence = module.capture_close_evidence(session, job, output)
    assert closed and evidence['response'] == response
    assert (output / 'runtime' / 'worker-stdout.log').read_text() == json.dumps(response)+'\n'
    response['value'] = []
    (job / 'worker-stdout.log').write_text(json.dumps(response)+'\n')
    output2 = tmp_path / 'output2'; output2.mkdir()
    assert not module.capture_close_evidence(session, job, output2)[0]


def test_fault_worker_passes_delayed_factory_explicitly_to_protocol(tmp_path, monkeypatch):
    module = runner()
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker
    monkeypatch.setattr(module, 'import_api', lambda path:None)
    monkeypatch.setattr(module.sys, 'argv', ['fixture'])
    monkeypatch.setattr(module.tempfile, 'tempdir', module.tempfile.tempdir)
    for key in ('TMP', 'TEMP', 'TMPDIR'):
        monkeypatch.setenv(key, module.os.environ.get(key, ''))
    session = SimpleNamespace(inspect_document=lambda:'native snapshot')
    original = session.inspect_document
    monkeypatch.setattr(worker, 'default_factory', lambda *a, **k:session)
    invoked = []
    def serve(incoming, outgoing, job, *, factory=None):
        assert factory is not None, 'delayed factory was not passed to the actual protocol entry'
        assert factory('writer', 'new_document', [], {}, deadline=123) is session
        assert session.inspect_document is not original
        invoked.append(True)
    monkeypatch.setattr(worker, 'serve', serve)
    module.delayed_worker(tmp_path, tmp_path)
    assert invoked == [True]


def test_fault_injection_runs_through_real_worker_protocol_without_com(tmp_path, monkeypatch):
    module = runner()
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker
    monkeypatch.setattr(module, 'import_api', lambda path:None)
    monkeypatch.setattr(module.tempfile, 'tempdir', module.tempfile.tempdir)
    for key in ('TMP', 'TEMP', 'TMPDIR'):
        monkeypatch.setenv(key, module.os.environ.get(key, ''))
    native = SimpleNamespace(inspect_document=lambda:{'native_snapshot':True}, close=lambda **kwargs:None)
    calls, delays = [], []
    def native_factory(kind, method, args, kwargs, *, deadline=None):
        calls.append((kind, method, deadline))
        return native
    monkeypatch.setattr(worker, 'default_factory', native_factory)
    monkeypatch.setattr(module.time, 'sleep', lambda seconds:delays.append(seconds))
    deadline = time.monotonic() + 10
    requests = [worker.encode_frame({'protocol':1, 'id':index, 'kind':'writer', 'method':method,
        'args':[], 'kwargs':{}, 'deadline':deadline, 'remaining_seconds':9})
        for index, method in enumerate(['new_document','inspect_document','close'], 1)]
    output = io.BytesIO()
    monkeypatch.setattr(module.sys, 'stdin', SimpleNamespace(buffer=io.BytesIO(b''.join(requests))))
    monkeypatch.setattr(module.sys, 'stdout', SimpleNamespace(buffer=output))
    module.delayed_worker(tmp_path, tmp_path)
    rows = [json.loads(line) for line in output.getvalue().splitlines()]
    assert calls[0][:2] == ('writer','new_document') and calls[0][2] is not None
    assert delays == [60]  # Real serve invoked the installed per-session stall.
    assert rows[1]['value'] == {'native_snapshot':True}
    assert rows[-1]['value'] == {'closed':True} and all(r['status'] == 'ok' for r in rows)
