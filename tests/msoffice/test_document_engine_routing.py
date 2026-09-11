from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import document_api as api
from skills.WPSComposer.scripts import office_engines


@pytest.mark.parametrize('action', ['open_document', 'inspect', 'edit', 'attach_active'])
def test_invalid_document_engine_rejected_before_native_action(monkeypatch, tmp_path, action):
    monkeypatch.setattr(api, '_com_available', lambda: pytest.fail('native probe'))
    with pytest.raises(ValueError, match='engine'):
        getattr(api, action)(**({'path': str(tmp_path/'x.xlsx')} if action != 'attach_active' else {}), engine='wrong')


def test_microsoft_inspect_uses_bound_session_and_closes(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    events = []
    class Bound:
        def __enter__(self):
            events.append('enter'); return self
        def __exit__(self, *args):
            events.append('close')
        def inspect_document(self, **options):
            events.append(options); return {'kind':'sheet','sheets':[]}
    def open_session(path, **kwargs):
        events.append(kwargs); return Bound()
    monkeypatch.setattr(document_sessions, 'open_session', open_session)
    monkeypatch.setattr(api, '_com_available', lambda: pytest.fail('WPS probe'))
    result = api.inspect(tmp_path/'x.xlsx', engine='msoffice', include_text=False)
    assert result['kind'] == 'sheet'
    assert events == [{'kind':'sheet','read_only':True,'visible':False}, 'enter', {'include_text':False}, 'close']


def test_explicit_office_open_never_uses_legacy_composer(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    calls = []
    monkeypatch.setattr(document_sessions, 'open_session', lambda path, **kw: calls.append(kw) or 'bound')
    monkeypatch.setattr(api.SheetComposer, 'open_document', lambda *a, **kw: pytest.fail('legacy'))
    assert api.open_document(tmp_path/'x.xlsx', engine='msoffice') == 'bound'
    assert calls == [{'kind':'sheet','read_only':False,'visible':False}]


def test_wps_file_open_excludes_microsoft_fallback(monkeypatch, tmp_path):
    def open_file(*args, **kwargs):
        assert office_engines.com_progids(('Ket.Application','Excel.Application')) == ('Ket.Application',)
        return 'wps'
    monkeypatch.setattr(api.SheetComposer, 'open_document', open_file)
    assert api.open_document(tmp_path/'x.xlsx') == 'wps'


def test_microsoft_attached_atomic_batch_fails_before_attachment(monkeypatch):
    from skills.WPSComposer.scripts import document_sessions
    monkeypatch.setattr(document_sessions, 'attach_session', lambda **kw: pytest.fail('attached'))
    result = api.edit(kind='sheet', engine='msoffice', ops=[
        {'op':'insert','target':'sheet:1','type':'row','index':2},
    ])
    assert not result['ok'] and not result['saved']
    assert result['errors'][0]['error']['code'] == 'atomic_attached_batch_unsupported'


def test_auto_resolves_once_before_opening_session(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    detected = []
    monkeypatch.setattr(office_engines, 'engine_executable', lambda e,c: detected.append((e,c)) or ('/Excel' if e=='msoffice' else None))
    monkeypatch.setattr(document_sessions, 'open_session', lambda *a, **kw: 'bound')
    assert api.open_document(tmp_path/'x.xlsx', engine='auto') == 'bound'
    assert detected == [('wps','spreadsheet'), ('msoffice','spreadsheet')]


@pytest.mark.parametrize('fails', [False, True])
def test_attached_microsoft_inspect_releases_session_even_on_failure(monkeypatch, fails):
    from skills.WPSComposer.scripts import document_sessions
    calls = []
    class Bound:
        def inspect_document(self, **options):
            if fails:
                raise RuntimeError('snapshot failed')
            return {'kind': 'sheet'}
        def close(self, *, save_changes=False):
            calls.append(save_changes)
    monkeypatch.setattr(document_sessions, 'attach_session', lambda **kw: Bound())
    if fails:
        with pytest.raises(RuntimeError, match='snapshot failed'):
            api.inspect(kind='sheet', engine='msoffice')
    else:
        assert api.inspect(kind='sheet', engine='msoffice') == {'kind': 'sheet'}
    assert calls == [False]


def test_lazy_wps_composer_keeps_pin_at_real_dispatch(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import _base
    calls = []
    monkeypatch.setattr(_base, '_require', lambda: None)
    monkeypatch.setattr(_base, '_dispatch', lambda progids: calls.append(tuple(progids)) or SimpleNamespace(app=object(), owns_app=False))
    monkeypatch.setattr(api.SheetComposer, '_open_document', staticmethod(lambda *a: object()))
    composer = api.open_document(tmp_path/'book.xlsx', engine='wps')
    composer._pool_app = False
    composer.__enter__()
    assert calls and all('Excel.Application' not in ids for ids in calls)


def test_wps_active_attach_filters_getactiveobject_fallback(monkeypatch):
    import sys
    from skills.WPSComposer.scripts import _base
    calls = []
    monkeypatch.setattr(_base, '_require', lambda: None)
    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    def active(progid):
        calls.append(progid)
        if progid == 'Excel.Application':
            return SimpleNamespace(ActiveWorkbook=object())
        raise RuntimeError('No WPS running')
    client = SimpleNamespace(GetActiveObject=active)
    monkeypatch.setitem(sys.modules, 'pythoncom', pythoncom)
    monkeypatch.setitem(sys.modules, 'win32com', SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, 'win32com.client', client)
    with pytest.raises(Exception):
        api.attach_active('sheet', engine='wps')
    assert 'Excel.Application' not in calls


def test_microsoft_output_format_fails_before_native_attachment(monkeypatch):
    from skills.WPSComposer.scripts import document_sessions
    monkeypatch.setattr(document_sessions, 'attach_session', lambda **kw: pytest.fail('attached before format check'))
    with pytest.raises(ValueError, match='native format'):
        api.edit(kind='sheet', engine='msoffice', output='copy.xls', patches=[{'target':'sheet:1/cell:A1','value':2}])


def test_microsoft_attached_save_preflight_precedes_mutation(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    events = []
    class Bound:
        kind = 'sheet'
        def supports_attached_save_copy(self): return True
        def preflight_save(self, output=None, *, overwrite=False):
            events.append(('preflight', output, overwrite))
            raise ValueError('invalid save destination')
        def apply_format_patch(self, *a, **kw): pytest.fail('mutated before save validation')
        def close(self, **kw): events.append('close')
    monkeypatch.setattr(document_sessions, 'attach_session', lambda **kw: Bound())
    output = str(tmp_path/'copy.xlsx')
    with pytest.raises(ValueError, match='invalid save destination'):
        api.edit(kind='sheet', engine='msoffice', output=output, overwrite=True, patches=[{'target':'sheet:1/cell:A1','value':2}])
    assert events == [('preflight',output,True),'close']


def test_auto_macos_open_skips_installed_wps_without_file_session(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    monkeypatch.setattr(office_engines.sys, 'platform', 'darwin')
    monkeypatch.setattr(office_engines, 'engine_executable', lambda *a: '/installed')
    monkeypatch.setattr(document_sessions, 'open_session', lambda *a, **kw: 'native-bound')
    assert api.open_document(tmp_path/'book.xlsx', engine='auto') == 'native-bound'


def test_auto_macos_structural_edit_skips_wps_before_open(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts import document_sessions
    monkeypatch.setattr(office_engines.sys, 'platform', 'darwin')
    monkeypatch.setattr(office_engines, 'engine_executable', lambda *a: '/installed')
    def native(*a, **kw): raise LookupError('selected native session')
    monkeypatch.setattr(document_sessions, 'open_session', native)
    monkeypatch.setattr(api, '_com_available', lambda: pytest.fail('wrong WPS route'))
    with pytest.raises(LookupError, match='selected native session'):
        api.edit(tmp_path/'slides.pptx', engine='auto', ops=[{'op':'insert','target':'presentation','type':'slide','index':1}])


def test_failed_native_copy_reports_clipboard_side_effect():
    class Session:
        kind = 'slide'
        def apply_structural_op(self, op):
            error = RuntimeError('native paste failed')
            error.clipboard_changed = True
            raise error
    reports = api.apply_ops(Session(), [{'op':'clone','target':'slide:1','to':{'index':2}}], atomic=False)
    assert reports[0]['clipboard_changed'] is True


def test_edited_artifact_validation_keeps_session_deadline(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(api, 'validate_before_deadline', lambda validator, path, deadline: calls.append((validator,path,deadline)), raising=False)
    class Session:
        engine = 'msoffice'
        _deadline = 123.5
    validator = api._session_validator(Session(), api._validate_edited_artifact)
    validator(tmp_path/'stage.docx')
    assert calls[0][1:] == (tmp_path/'stage.docx',123.5)


@pytest.mark.parametrize('kind,name',[('writer','ProxyWordSession'),('sheet','ProxyExcelSession'),('slide','ProxyPowerPointSession')])
def test_windows_sessions_use_deadline_bound_process_proxy(monkeypatch, kind, name):
    from skills.WPSComposer.scripts import document_sessions
    from skills.WPSComposer.scripts.msoffice import windows_session_proxy
    monkeypatch.setattr(document_sessions.sys, 'platform', 'win32')
    assert document_sessions._session_class(kind) is getattr(windows_session_proxy, name)


@pytest.mark.parametrize('kind', ['writer','sheet','slide'])
def test_create_microsoft_document_keeps_native_factory_pin(monkeypatch, kind):
    from skills.WPSComposer.scripts import document_sessions
    events = []
    monkeypatch.setattr(document_sessions, 'new_session', lambda **kwargs: events.append(kwargs) or 'native', raising=False)
    assert api.create_document(kind, visible=True, engine='msoffice') == 'native'
    assert events == [{'kind':kind, 'visible':True}]


def test_create_default_wps_retains_pin_after_factory_returns(monkeypatch):
    composer = api.create_document('sheet')
    assert 'Excel.Application' not in composer._progids


@pytest.mark.parametrize('kwargs',[{'engine':'wrong'},{'kind':'wrong'}])
def test_create_invalid_request_fails_without_native_discovery(monkeypatch, kwargs):
    monkeypatch.setattr(office_engines, 'engine_executable', lambda *a: pytest.fail('native discovery'))
    with pytest.raises(ValueError):
        api.create_document(**kwargs)


@pytest.mark.parametrize('argument', ['ops', 'patches'])
def test_edit_materializes_generator_once_before_action_routing(monkeypatch, tmp_path, argument):
    class Session:
        kind = 'writer'
        def __init__(self):
            self.applied = []
        def __enter__(self):
            return self
        def close(self, save_changes=False):
            pass
        def save_current(self):
            return 'saved-native-document'
        def apply_format_patch(self, target, **values):
            self.applied.append((target, values))
            return {'accepted': list(values), 'rejected': []}

    session = Session()
    monkeypatch.setattr(api, '_com_available', lambda: True)
    monkeypatch.setattr(api, 'open_document', lambda *a, **kw: session)
    operation = {'target': 'paragraph:1', 'text': 'requested edit'}
    if argument == 'ops':
        operation['op'] = 'set'
    result = api.edit(tmp_path/'uncreated-review.docx', engine='wps',
                      **{argument: (dict(operation) for _ in range(1))})
    assert result['ok'] is True
    assert len(result['ops']) == 1
    assert session.applied == [('paragraph:1', {'text': 'requested edit'})]
