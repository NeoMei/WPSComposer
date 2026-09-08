"""Complete native candidate validation without Office, files or live IDs."""
from __future__ import annotations

import pytest

from skills.WPSComposer.scripts import document_api as api, office_engines as engines


@pytest.fixture
def mac_ms(monkeypatch):
    monkeypatch.setattr(engines.sys, 'platform', 'darwin')
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, *args: '/fake/Microsoft.app' if engine == 'msoffice' else None)


@pytest.mark.parametrize('suffix,first,later', [
    ('docx', {'target':'paragraph:1','font':{'bold':True}}, {'target':'paragraph:1','geometry':{'width':200}}),
    ('xlsx', {'target':'sheet:1/cell:A1','value':42}, {'target':'sheet:1/cell:A1','fill':{'transparency':.5}}),
    ('pptx', {'target':'slide:1/shape:1','text':'first'}, {'target':'presentation','page_setup':{'slide_height':540}}),
    ('pptx', {'target':'slide:1/shape:1','text':'first'}, {'target':'slide:1/shape:@id=7','font':{'bold':True}}),
])
def test_auto_rejects_later_backend_capability_before_native_open(mac_ms, monkeypatch, suffix, first, later):
    opened=[]
    def open_native(*args, **kwargs):
        opened.append(True)
        raise AssertionError('Native opened before whole-request validation')
    monkeypatch.setattr(api, 'open_document', open_native)
    with pytest.raises(engines.EngineUnavailableError):
        api.edit('/pure-audit.'+suffix, engine='auto', ops=({'op':'set',**p} for p in (first,later)))
    assert opened == []


@pytest.mark.parametrize('suffix,operation', [
    ('docx', {'op':'set','target':'paragraph:1','font':{'bold':True,'size':-2}}),
    ('xlsx', {'op':'set','target':'sheet:1/cell:A1','value':42,'font':{'size':-2}}),
    ('pptx', {'op':'set','target':'slide:1/shape:1','text':'first','font':{'size':-2}}),
    ('docx', {'op':'move','target':'shape:1','to':'end'}),
    ('xlsx', {'op':'clone','target':'sheet:1/chart:1','to':'end'}),
    ('pptx', {'op':'insert','parent':'slide:1','type':'table','props':{'rows':2,'cols':2}}),
])
def test_auto_rejects_mixed_invalid_values_and_structural_forms(mac_ms, suffix, operation):
    with pytest.raises(engines.EngineUnavailableError):
        api._document_engine('/source.'+suffix,None,'auto', action='edit',operations=[operation])


@pytest.mark.parametrize('suffix,operation', [
    ('docx', {'op':'set','target':'paragraph:@paraId=ABCD1234','font':{'bold':True}}),
    ('xlsx', {'op':'set','target':'sheet:999/cell:A1','value':42}),
    ('pptx', {'op':'set','target':'slide:999/shape:@name=Future','text':'valid form'}),
])
def test_auto_pure_support_does_not_invent_live_existence(mac_ms, suffix, operation):
    assert api._document_engine('/source.'+suffix,None,'auto',action='edit',operations=[operation]) == 'msoffice'


def test_explicit_engine_keeps_best_effort_and_windows_does_not_inherit_mac_limits(mac_ms,monkeypatch):
    operation={'op':'set','target':'slide:1/shape:@id=7','font':{'bold':True}}
    assert api._document_engine('/source.pptx',None,'msoffice',action='edit',operations=[operation])=='msoffice'
    monkeypatch.setattr(engines.sys,'platform','win32')
    assert api._document_engine('/source.pptx',None,'auto',action='edit',operations=[operation])=='msoffice'


@pytest.mark.parametrize('value',[{},None])
def test_word_unknown_empty_property_is_rejected_before_native(mac_ms,value):
    with pytest.raises(engines.EngineUnavailableError):
        api._document_engine('/source.docx',None,'auto',action='edit',operations=[
            {'op':'set','target':'paragraph:1','unknown':value}])


def test_excel_existing_sheet_deletion_is_known_unsupported_without_inserting_sheet(mac_ms):
    with pytest.raises(engines.EngineUnavailableError):
        api._document_engine('/source.xlsx',None,'auto',action='edit',operations=[
            {'op':'set','target':'sheet:1/cell:A1','value':42},
            {'op':'remove','target':'sheet:2'}])


def test_wps_preference_is_preserved_for_supported_request(mac_ms,monkeypatch):
    monkeypatch.setattr(engines,'engine_executable',lambda *args:'/installed/app')
    op={'op':'set','target':'slide:1/shape:@id=7','font':{'bold':True}}
    assert api._document_engine('/source.pptx',None,'auto',action='edit',operations=[op]) == 'wps'


def test_both_installed_skip_wps_and_ms_for_incomplete_request(mac_ms,monkeypatch):
    monkeypatch.setattr(engines,'engine_executable',lambda *args:'/installed/app')
    with pytest.raises(engines.EngineUnavailableError):
        api._document_engine('/source.pptx',None,'auto',action='edit',operations=[
            {'op':'set','target':'slide:1/shape:1','text':'valid'},
            {'op':'set','target':'presentation','page_setup':{'slide_height':540}}])


def test_supported_public_generator_batch_is_selected_once(mac_ms,monkeypatch):
    yielded=[]
    def ops():
        for text in ('one','two'):
            yielded.append(text)
            yield {'op':'set','target':'slide:1/shape:1','text':text}
    def open_native(*args, **kwargs):
        assert kwargs['engine']=='msoffice'
        raise RuntimeError('selected native opener')
    monkeypatch.setattr(api,'open_document',open_native)
    with pytest.raises(RuntimeError,match='selected native opener'):
        api.edit('/source.pptx',engine='auto',ops=ops())
    assert yielded == ['one','two']


def test_explicit_best_effort_still_dispatches_first_supported_operation(mac_ms,monkeypatch):
    from skills.WPSComposer.scripts.msoffice.macos_powerpoint_session import MacPowerPointSession
    session=MacPowerPointSession();calls=[]
    monkeypatch.setattr(session,'_run',lambda *args,**kwargs:calls.append(args[0]))
    monkeypatch.setattr(session,'close',lambda **kwargs:None)
    monkeypatch.setattr(session,'save_current',lambda:'/source.pptx')
    monkeypatch.setattr(MacPowerPointSession,'__enter__',lambda self:self)
    monkeypatch.setattr(api,'open_document',lambda *args,**kwargs:session)
    result=api.edit('/source.pptx',engine='msoffice',atomic=False,ops=[
        {'op':'set','target':'slide:1/shape:1','text':'first'},
        {'op':'set','target':'slide:1/shape:@id=7','font':{'bold':True}}])
    assert len(calls)==1 and result['ops'][0]['ok'] is True and result['ops'][1]['ok'] is False


@pytest.mark.parametrize('module_name,class_name', [('macos_word_session','MacWordSession'),('macos_excel_session','MacExcelSession')])
def test_execution_calls_the_same_pure_patch_compiler(monkeypatch,module_name,class_name):
    import importlib
    from skills.WPSComposer.scripts.msoffice import edit_preflight as pure
    module=importlib.import_module('skills.WPSComposer.scripts.msoffice.'+module_name)
    cls=getattr(module,class_name)
    session=cls.__new__(cls)
    method='compile_word_patch' if class_name=='MacWordSession' else 'compile_excel_patch'
    def reject(*args,**kwargs): raise ValueError('shared compiler rejection')
    monkeypatch.setattr(pure,method,reject)
    monkeypatch.setattr(session,'_writable' if class_name=='MacWordSession' else '_assert_live',lambda *a,**k:None)
    target='paragraph:1' if class_name=='MacWordSession' else 'sheet:1/cell:A1'
    with pytest.raises(ValueError,match='shared compiler rejection'):
        session.apply_format_patch(target,font={'bold':True})


def test_fresh_sheet_possible_does_not_prove_an_arbitrary_target_exists(mac_ms):
    operations=[{'op':'insert','type':'sheet','props':{'name':'New'}},
                {'op':'remove','target':'sheet:999'}]
    assert api._document_engine('/source.xlsx',None,'auto',action='edit',operations=operations)=='msoffice'
    # Native deletion clears its fresh-sheet tracking; a second positional
    # deletion cannot borrow an earlier insertion as a support guarantee.
    with pytest.raises(engines.EngineUnavailableError):
        api._document_engine('/source.xlsx',None,'auto',action='edit',operations=operations+[{'op':'remove','target':'sheet:1'}])


@pytest.mark.parametrize('family,target,extra',[
    ('writer','paragraph:1',{'op':'clone','target':'shape:1','to':'end'}),
    ('sheet','sheet:1/cell:A1',{'op':'insert','parent':'sheet:1','type':'row','position':{'index':0}}),
    ('slide','slide:1/shape:1',{'op':'move','target':'slide:1/shape:1','to':{'slide':False}}),
])
def test_structural_later_rejection_does_not_open_native(mac_ms,monkeypatch,family,target,extra):
    suffix={'writer':'docx','sheet':'xlsx','slide':'pptx'}[family]
    def opened(*args,**kwargs): pytest.fail('native session opened for unsupported batch')
    monkeypatch.setattr(api,'open_document',opened)
    first={'op':'set','target':target,**({'value':42} if family=='sheet' else {'text':'valid'})}
    with pytest.raises(engines.EngineUnavailableError):
        api.edit('/source.'+suffix,engine='auto',ops=[first,extra])


def test_shared_validation_import_has_no_native_or_filesystem_execution():
    import subprocess
    import sys
    source='''import subprocess
def no_native(*args,**kwargs):raise AssertionError('native process launched')
subprocess.run=no_native
subprocess.Popen=no_native
from skills.WPSComposer.scripts.msoffice.edit_preflight import supports_edit_ops
assert supports_edit_ops('writer',[{'op':'set','target':'paragraph:@paraId=ABCD','font':{'bold':True}}],platform='darwin')
assert supports_edit_ops('sheet',[{'op':'set','target':'sheet:1/cell:A1','value':42}],platform='darwin')
assert supports_edit_ops('slide',[{'op':'set','target':'slide:1/shape:1','text':'ok'}],platform='darwin')
'''
    subprocess.run([sys.executable,'-c',source],check=True,timeout=10)


def test_generation_plan_catalog_remains_frozen():
    from skills.WPSComposer.scripts.generation_plan import ALLOWED_OPERATIONS
    assert len(ALLOWED_OPERATIONS['spreadsheet'])==7
    assert len(ALLOWED_OPERATIONS['presentation'])==9


def test_selection_shape_capability_is_deferred_to_live_selection_not_rejected(mac_ms):
    assert api._document_engine(None,'slide','auto',action='edit',operations=[
        {'op':'set','target':'selection','geometry':{'width':200}}]) == 'msoffice'


@pytest.mark.parametrize('kind,props', [
    ('table', {'rows':1,'cols':1,'data':[['bad\0']]}),
    ('textbox', {'text':'bad\0'}),
])
def test_word_invalid_structural_text_rejected_before_auto_open(mac_ms,monkeypatch,kind,props):
    opened=[]
    monkeypatch.setattr(api,'open_document',lambda *a,**kw:opened.append(True))
    with pytest.raises(engines.EngineUnavailableError):
        api.edit('/source.docx',engine='auto',ops=iter([
            {'op':'set','target':'paragraph:1','font':{'bold':True}},
            {'op':'insert','parent':'body','type':kind,'props':props}]))
    assert opened == []


@pytest.fixture
def recording_word(monkeypatch):
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
    session=MacWordSession(); calls=[]
    monkeypatch.setattr(session,'_writable',lambda:None)
    monkeypatch.setattr(session,'_position',lambda *a:[])
    monkeypatch.setattr(session,'_range',lambda *a:([], 'targetRange'))
    monkeypatch.setattr(session,'_execute',lambda lines,**kw:calls.append(lines))
    monkeypatch.setattr(session,'close',lambda **kw:None)
    monkeypatch.setattr(session,'save_current',lambda:'/source.docx')
    monkeypatch.setattr(MacWordSession,'__enter__',lambda self:self)
    return session,calls


@pytest.mark.parametrize('kind,props', [
    ('table', {'rows':1,'cols':1,'data':[['bad\0']]}),
    ('textbox', {'text':'bad\0'}),
])
def test_word_invalid_structural_text_direct_never_transports(recording_word,kind,props):
    session,calls=recording_word
    with pytest.raises(ValueError,match='control characters'):
        session.apply_structural_op({'op':'insert','parent':'body','type':kind,'props':props})
    assert calls == []


@pytest.mark.parametrize('engine',['auto','msoffice'])
def test_word_public_nested_generators_preserve_text_and_best_effort(mac_ms,monkeypatch,recording_word,engine):
    session,calls=recording_word
    monkeypatch.setattr(api,'open_document',lambda *a,**kw:session)
    yielded=[]
    def rows():
        yielded.append('row')
        yield ['keep me']
    def operations():
        yielded.append('op')
        yield {'op':'insert','parent':'body','type':'table','props':{'rows':1,'cols':1,'data':rows()}}
        if engine=='msoffice':
            yield {'op':'insert','parent':'body','type':'textbox','props':{'text':'bad\0'}}
    result=api.edit('/source.docx',engine=engine,atomic=False,ops=operations())
    assert result['ops'][0]['ok'] is True
    if engine=='msoffice': assert result['ops'][1]['ok'] is False
    assert yielded == ['op','row']
    assert len(calls)==1
    assert any('keep me' in line for line in calls[0])


def test_word_direct_table_generator_preserves_text(recording_word):
    session,calls=recording_word
    session.apply_structural_op({'op':'insert','parent':'body','type':'table',
        'props':{'rows':1,'cols':1,'data':(row for row in [['keep me']])}})
    assert any('keep me' in line for line in calls[0])


def test_word_supports_predicate_does_not_consume_unmaterialized_table_generator():
    from skills.WPSComposer.scripts.msoffice.edit_preflight import supports_edit_ops
    data=(row for row in [['keep me']])
    assert not supports_edit_ops('writer',[{'op':'insert','parent':'body','type':'table',
        'props':{'rows':1,'cols':1,'data':data}}],platform='darwin')
    assert list(data)==[['keep me']]


def test_word_row_generator_still_rejected_without_consumption(recording_word):
    session,calls=recording_word
    row=(text for text in ['keep me'])
    with pytest.raises(ValueError,match='Table data exceeds dimensions'):
        session.apply_structural_op({'op':'insert','parent':'body','type':'table',
            'props':{'rows':1,'cols':1,'data':(item for item in [row])}})
    assert calls == []
    assert list(row)==['keep me']


@pytest.mark.parametrize('engine',['auto','msoffice'])
def test_word_public_row_generator_keeps_baseline_rejection(mac_ms,monkeypatch,recording_word,engine):
    session,calls=recording_word; opened=[]
    def open_native(*a,**kw):
        opened.append(True)
        return session
    monkeypatch.setattr(api,'open_document',open_native)
    row=(text for text in ['keep me'])
    operation={'op':'insert','parent':'body','type':'table',
               'props':{'rows':1,'cols':1,'data':(item for item in [row])}}
    if engine=='auto':
        with pytest.raises(engines.EngineUnavailableError):
            api.edit('/source.docx',engine=engine,ops=iter([operation]))
        assert opened == []
    else:
        result=api.edit('/source.docx',engine=engine,atomic=False,ops=iter([operation]))
        assert result['ops'][0]['ok'] is False
    assert calls == []
    assert list(row)==['keep me']


def test_word_materialization_bounded_and_reuses_canonical_data():
    from skills.WPSComposer.scripts.msoffice.edit_preflight import materialize_word_structural, supports_edit_ops
    visited=[]
    def rows():
        for i in range(100):
            visited.append(i)
            yield [i]
    operation={'op':'insert','parent':'body','type':'table','props':{'rows':1,'cols':1,'data':rows()}}
    canonical=materialize_word_structural(operation)
    assert canonical is not operation and canonical['props'] is not operation['props']
    assert visited == [0,1]
    assert materialize_word_structural(canonical) is canonical
    assert not supports_edit_ops('writer',[canonical],platform='darwin')
    assert visited == [0,1]


@pytest.mark.parametrize('kind,props', [
    ('table', {'rows':1,'cols':1,'data':[[42]]}),
    ('textbox', {'text':42}),
])
def test_word_structural_text_conversion_matches_execution(recording_word,kind,props):
    from skills.WPSComposer.scripts.msoffice.edit_preflight import supports_edit_ops
    session,calls=recording_word
    operation={'op':'insert','parent':'body','type':kind,'props':props}
    assert supports_edit_ops('writer',[operation],platform='darwin')
    session.apply_structural_op(operation)
    assert any('to "42"' in line for line in calls[0])


@pytest.fixture
def windows_engines(monkeypatch):
    monkeypatch.setattr(engines.sys, 'platform', 'win32')


@pytest.mark.parametrize('installed', ['msoffice-only', 'both-installed'])
@pytest.mark.parametrize('suffix,operations', [
    ('docx', [
        {'op':'set','target':'paragraph:1','font':{'bold':True}},
        {'op':'set','target':'paragraph:1','geometry':{'width':200}},
    ]),
    ('xlsx', [
        {'op':'set','target':'sheet:1/cell:A1','value':42},
        {'op':'set','target':'sheet:1/cell:A1','fill':{'transparency':.5}},
    ]),
    ('pptx', [
        {'op':'set','target':'slide:1/shape:1','text':'first'},
        {'op':'set','target':'presentation','font':{'bold':True}},
    ]),
])
def test_windows_auto_rejects_supported_first_known_ignored_later_before_native_open(
        windows_engines, monkeypatch, installed, suffix, operations):
    opened=[]
    monkeypatch.setattr(
        engines, 'engine_executable',
        lambda engine, component: ('/installed/' + engine)
        if installed == 'both-installed' or engine == 'msoffice' else None,
    )
    monkeypatch.setattr(
        api, 'open_document',
        lambda *args, **kwargs: opened.append(kwargs.get('engine')),
    )
    with pytest.raises(engines.EngineUnavailableError):
        api.edit('/source.' + suffix, engine='auto', ops=operations)
    assert opened == []


@pytest.mark.parametrize('suffix,operation', [
    ('docx', {'op':'set','target':'paragraph:@paraId=ABCD','font':{'bold':True}}),
    ('xlsx', {'op':'set','target':'sheet:99/shape:@name=Future','geometry':{'width':200}}),
    ('pptx', {'op':'set','target':'slide:99/shape:@id=7/paragraph:2/run:3',
              'font':{'underline':99}}),
])
def test_windows_auto_supported_request_preserves_wps_preference(
        windows_engines, monkeypatch, suffix, operation):
    monkeypatch.setattr(engines, 'engine_executable', lambda *args: '/installed/app')
    assert api._document_engine('/source.' + suffix, None, 'auto', action='edit',
                                operations=[operation]) == 'wps'


@pytest.mark.parametrize('operation', [
    {'op':'move','target':'sheet:2','to':'start'},
    {'op':'clone','target':'sheet:1','to':'end'},
])
def test_windows_auto_can_skip_wps_for_microsoft_only_structural_form(
        windows_engines, monkeypatch, operation):
    monkeypatch.setattr(engines, 'engine_executable', lambda *args: '/installed/app')
    assert api._document_engine('/source.xlsx', None, 'auto', action='edit',
                                operations=[operation]) == 'msoffice'


def test_windows_auto_preserves_wps_for_in_place_worksheet_clone(
        windows_engines, monkeypatch):
    monkeypatch.setattr(engines, 'engine_executable', lambda *args: '/installed/app')
    operation={'op':'clone','target':'sheet:1','to':None}
    assert api._document_engine('/source.xlsx', None, 'auto', action='edit',
                                operations=[operation]) == 'wps'


@pytest.mark.parametrize('family,operation,expected', [
    ('writer', {'op':'set','target':'paragraph:1','geometry':{'width':200}}, 'unsupported'),
    ('sheet', {'op':'set','target':'sheet:1/cell:A1','fill':{'transparency':.5}}, 'unsupported'),
    ('slide', {'op':'set','target':'presentation','font':{'bold':True}}, 'unsupported'),
    ('writer', {'op':'set','target':'paragraph:@paraId=ABCDEF','paragraph':{'line_spacing_rule':99}}, 'supported'),
    ('slide', {'op':'set','target':'slide:99/shape:@id=7/paragraph:2/run:3',
               'font':{'underline':99}}, 'supported'),
    ('slide', {'op':'set','target':'slide:99/shape:@id=7/table/cell:2,3',
               'text_frame':{'vertical_anchor':99}}, 'supported'),
])
def test_windows_confirmed_target_field_cases(family, operation, expected):
    from skills.WPSComposer.scripts.msoffice.edit_preflight import classify_windows_set_op
    assert classify_windows_set_op(family, operation) == expected


@pytest.mark.parametrize('family,operation', [
    ('writer', {'op':'set','target':'paragraph:1','font':{'unknown':1}}),
    ('writer', {'op':'set','target':'shape:1','fill':{'color':'#GG0000'}}),
    ('sheet', {'op':'set','target':'sheet:1/cell:A1','borders':{'x':{'style':1}}}),
    ('slide', {'op':'set','target':'slide:1/shape:1','text_frame':{'unknown':1}}),
    ('slide', {'op':'set','target':'slide:1/shape:1','unknown':{}}),
])
def test_windows_nested_rejection_and_unknown_keyword_are_unsupported(family, operation):
    from skills.WPSComposer.scripts.msoffice.edit_preflight import classify_windows_set_op
    assert classify_windows_set_op(family, operation) == 'unsupported'


def test_windows_pure_validation_does_not_resolve_live_ids_or_launch_native(monkeypatch):
    from skills.WPSComposer.scripts.msoffice import edit_preflight as pure
    import subprocess
    monkeypatch.setattr(subprocess, 'run', lambda *a, **k: pytest.fail('launched process'))
    monkeypatch.setattr(subprocess, 'Popen', lambda *a, **k: pytest.fail('launched process'))
    assert pure.classify_windows_set_op('writer', {
        'op':'set','target':'range:100-200','font':{'bold':True}}) == 'supported'
    assert pure.classify_windows_set_op('sheet', {
        'op':'set','target':'sheet:999/shape:@id=999','line':{'dash_style':99}}) == 'supported'
    assert pure.classify_windows_set_op('slide', {
        'op':'set','target':'selection','geometry':{'width':200}}) == 'supported'
