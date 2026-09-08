"""COM-free checks for the persistent Windows session process boundary."""
from __future__ import annotations

import importlib
import io
import json
from pathlib import Path
import subprocess
import sys
import time
import threading
from types import SimpleNamespace

import pytest

TEST_DEADLINE = time.monotonic() + 500


def modules():
    return (importlib.import_module('skills.WPSComposer.scripts.msoffice.windows_session_proxy'),
            importlib.import_module('skills.WPSComposer.scripts.msoffice.windows_session_worker'))


class Lock:
    def __init__(self, root):
        self.quarantine_path = root / 'native-office.quarantine.json'
        self.closed = False
    def acquire(self, deadline):
        if self.quarantine_path.exists():
            raise RuntimeError('quarantined')
    def quarantine(self, detail):
        self.quarantine_path.write_text(json.dumps(detail))
    def close(self):
        self.closed = True


@pytest.fixture
def transport(monkeypatch, tmp_path):
    m, _ = modules()
    monkeypatch.setattr(m, '_component_root', lambda component: tmp_path / component)
    monkeypatch.setattr(m, 'OfficeJobLock', Lock)
    children = []
    def launch(job, *, bad=None, delay=None):
        code = '''import json,sys,time
for line in sys.stdin:
 r=json.loads(line)
 if DELAY and r['method']==DELAY: time.sleep(10)
 value={'kind':r['kind']} if r['method'] in ('open_document','new_document') else {'text':'中文','clipboard_changed':True}
 if r['method']=='close': value={'closed':True}
 result={'protocol':1,'id':r['id'],'status':'ok','value':value,'identity':{'office_pid':987,'staging_path':'private-doc'}}
 if BAD=='id' and r['method']!='open_document':result['id']+=1
 if BAD=='error' and r['method']=='apply_format_patch':result.update(status='error',error={'type':'PermissionError','message':'read only'})
 if BAD=='close' and r['method']=='close':result['value']={}
 if BAD=='oversize' and r['method']!='open_document':sys.stdout.write('x'*2000+'\\n');sys.stdout.flush();continue
 print(json.dumps(result),flush=True)
 if r['method']=='close':break
'''.replace('DELAY', repr(delay)).replace('BAD', repr(bad))
        child = subprocess.Popen([sys.executable, '-u', '-c', code], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        children.append(child)
        return child
    monkeypatch.setattr(m, '_spawn_worker', lambda job: launch(job))
    yield m, launch, children
    for child in children:
        if child.poll() is None: child.kill()
        child.wait(timeout=2)


def test_persistent_proxy_roundtrip_and_verified_close(transport):
    m, _, children = transport
    with m.ProxyWordSession.open_document('/source.docx') as session:
        deadline = session.publication_deadline
        result = session.inspect_document()
        assert result == {'text': '中文', 'clipboard_changed': True}
        assert session.publication_deadline == deadline
        assert len(children) == 1
        assert session.kind == 'writer' and session.engine == 'msoffice'
    assert children[0].poll() == 0
    assert session._lock.closed
    assert not session._lock.quarantine_path.exists()


@pytest.mark.parametrize('name', ['ProxyWordSession','ProxyExcelSession','ProxyPowerPointSession'])
def test_proxy_new_document_preserves_native_factory_request(transport, name):
    m, _, _=transport
    with getattr(m,name).new_document(visible=True) as session:
        request=json.loads((session.staging_root/'last-request.json').read_text())
        assert request['method']=='new_document'
        assert request['args']==[]
        assert request['kwargs']=={'visible':True}


def test_deadline_kills_only_exact_python_child_and_preserves_recovery(transport, monkeypatch):
    m, launch, children = transport
    monkeypatch.setattr(m, '_spawn_worker', lambda job: launch(job, delay='inspect_document'))
    session = m.ProxyExcelSession.open_document('/source.xlsx')
    session._deadline = time.monotonic() + .1
    before = time.monotonic()
    with pytest.raises(m.NativeOfficeError) as error:
        session.inspect_document()
    assert error.value.code == 'NATIVE_OFFICE_TIMEOUT'
    assert time.monotonic() - before < 2
    assert children[0].poll() is not None
    recovery = json.loads((session.staging_root / 'recovery.json').read_text())
    assert recovery['office_termination_attempted'] is False
    assert recovery['identity']['office_pid'] == 987
    assert recovery['last_request']['method'] == 'inspect_document'
    assert session._lock.quarantine_path.exists()
    with pytest.raises(m.NativeOfficeError): session.close()
    with pytest.raises(RuntimeError, match='quarantined'): m.ProxyExcelSession.open_document('/new.xlsx')
    assert len(children) == 1


@pytest.mark.parametrize('bad', ['id', 'oversize'])
def test_corrupt_response_is_never_success_and_quarantines(transport, monkeypatch, bad):
    m, launch, _ = transport
    monkeypatch.setattr(m, 'MAX_FRAME_BYTES', 1024)
    monkeypatch.setattr(m, '_spawn_worker', lambda job: launch(job, bad=bad))
    session = m.ProxyPowerPointSession.open_document('/source.pptx')
    with pytest.raises(m.NativeOfficeError): session.inspect_document()
    assert session._lock.quarantine_path.exists()
    assert session.staging_root.exists()


def test_remote_permission_error_preserves_type_and_can_close(transport, monkeypatch):
    m, launch, _ = transport
    monkeypatch.setattr(m, '_spawn_worker', lambda job: launch(job, bad='error'))
    with m.ProxyWordSession.open_document('/source.docx') as session:
        with pytest.raises(PermissionError, match='read only'):
            session.apply_format_patch('paragraph:1', text='x')
        assert not session._lock.quarantine_path.exists()


def test_closed_allowlist_rejects_raw_com_and_non_json_before_request(transport):
    m, _, _ = transport
    with m.ProxyWordSession.open_document('/source.docx') as session:
        count = session._sequence
        for name in ('doc', 'app', 'selection', 'open', '_native_eval'):
            with pytest.raises(AttributeError): getattr(session, name)
        with pytest.raises(NotImplementedError): session.apply_design_preset(object())
        with pytest.raises((TypeError, ValueError)): session.apply_format_patch('paragraph:1', text=object())
        assert session._sequence == count


def test_path_arguments_are_plain_json_paths(transport, tmp_path):
    m, _, _ = transport
    with m.ProxyWordSession.open_document(tmp_path/'source.docx') as session:
        session.save(tmp_path/'copy.docx')
        request = json.loads((session.staging_root/'last-request.json').read_text())
        assert request['args'] == [str(tmp_path/'copy.docx')]


def test_close_requires_explicit_ack_and_preserves_original_operation_error(transport, monkeypatch):
    m, launch, _ = transport
    monkeypatch.setattr(m, '_spawn_worker', lambda job: launch(job, bad='close'))
    with pytest.raises(ValueError, match='original error'):
        with m.ProxyWordSession.open_document('/source.docx') as session:
            raise ValueError('original error')
    assert session._lock.quarantine_path.exists()
    assert session._uncertain


def test_blocked_pipe_write_is_bounded_without_windows_select(transport, monkeypatch):
    m, _, children = transport
    session = m.ProxyExcelSession.open_document('/source.xlsx')
    child = children[0]
    original = child.stdin
    unblocked = threading.Event()
    class BlockedPipe:
        def write(self, raw):
            unblocked.wait(3)
            raise BrokenPipeError()
        def flush(self): pass
    child.stdin = BlockedPipe()
    session._deadline = time.monotonic()+.1
    try:
        with pytest.raises(m.NativeOfficeError) as error: session.inspect_document()
        assert error.value.code == 'NATIVE_OFFICE_TIMEOUT'
        assert child.poll() is not None
    finally:
        unblocked.set()
        original.close()


def test_reader_eof_quarantines_unverified_session(transport):
    m, _, children = transport
    session = m.ProxyWordSession.open_document('/source.docx')
    children[0].kill();children[0].wait(timeout=2)
    with pytest.raises(m.NativeWordError): session.inspect_document()
    assert session._lock.quarantine_path.exists()


def native_registry(kind):
    _, w=modules()
    class Collection:
        def __init__(self,values=()):self.values=list(values)
        @property
        def Count(self):return len(self.values)
        def Item(self,index):return self.values[index-1]
    shape=SimpleNamespace(token='shape')
    slide=SimpleNamespace(token='slide',Shapes=Collection([shape]))
    sheet=SimpleNamespace(token='sheet',Shapes=Collection(),ChartObjects=Collection())
    table=SimpleNamespace(token='table')
    doc=SimpleNamespace(Slides=Collection([slide]),Worksheets=Collection([sheet]),Tables=Collection([table]),Shapes=Collection(),InlineShapes=Collection())
    session=SimpleNamespace(kind=kind,_composer=SimpleNamespace(_doc=doc,_deps=SimpleNamespace(identity=lambda obj:obj.token)),_verify=lambda:None)
    return w.NativeHandleRegistry(session),session,slide,sheet,table,shape


def test_native_handles_preserve_created_com_results_and_semantic_arguments():
    registry,session,slide,_,_,shape=native_registry('slide')
    wire=registry.encode_result('add_blank_slide',(slide,1),[],{})
    reference=wire['__wpscomposer_tuple__'][0]
    assert reference['__wpscomposer_handle__']['type']=='slide'
    args,kwargs=registry.prepare_call('add_textbox',[reference,'hello',1,2,3,4],{})
    assert args==[1,'hello',1,2,3,4] and kwargs=={}
    shape_ref=registry.encode_result('add_shape',shape,[],{})
    args,_=registry.prepare_call('apply_format_patch',[shape_ref],{'text':'edited'})
    assert args==['slide:1/shape:1']


def test_sheet_select_and_rename_return_handles_instead_of_serialization_error():
    registry,_,_,sheet,_,_=native_registry('sheet')
    wire=registry.encode_result('select_sheet',sheet,[1],{})
    args,_=registry.prepare_call('rename_sheet',[wire,'renamed'],{})
    assert args==[1,'renamed']
    assert registry.encode_result('rename_sheet',sheet,args,{})['__wpscomposer_handle__']['type']=='sheet'


def test_stale_foreign_and_wrong_type_handles_fail_before_native_call():
    registry,session,slide,_,_,shape=native_registry('slide')
    ref=registry.encode_result('add_shape',shape,[],{})
    with pytest.raises(ValueError,match='type'):registry.prepare_call('add_textbox',[ref,'x',1,2,3,4],{})
    other,*_=native_registry('slide')
    with pytest.raises(ValueError,match='session'):other.prepare_call('apply_format_patch',[ref],{})
    slide.Shapes.values=[]
    with pytest.raises(ValueError,match='stale'):registry.prepare_call('apply_format_patch',[ref],{})


def test_handle_destinations_preserve_component_specific_anchor_contract():
    registry,_,slide,_,_,shape=native_registry('slide')
    slide_ref=registry.register(slide,'slide')
    shape_ref=registry.register(shape,'shape')
    for key in ('before','after'):
        args,_=registry.prepare_call('apply_structural_op',[{'op':'clone','target':slide_ref,'to':{key:slide_ref}}],{})
        assert args[0]=={'op':'clone','target':'slide:1','to':{key:'slide:1'}}
    args,_=registry.prepare_call('apply_structural_op',[{'op':'move','target':shape_ref,'to':{'slide':slide_ref}}],{})
    assert args[0]=={'op':'move','target':'slide:1/shape:1','to':{'slide':1}}
    excel,_,_,sheet,_,_=native_registry('sheet')
    sheet_ref=excel.register(sheet,'sheet')
    args,_=excel.prepare_call('apply_structural_op',[{'op':'clone','target':sheet_ref,'to':{'after':sheet_ref}}],{})
    assert args[0]['to']=={'after':1}


def test_word_range_anchor_requires_exact_live_paragraph_range():
    registry,session,_,_,_,_=native_registry('writer')
    doc=session._composer._doc
    doc.token='doc';doc.Content=SimpleNamespace(End=12)
    first=SimpleNamespace(Start=0,End=6,Document=doc,token='range1')
    second=SimpleNamespace(Start=6,End=12,Document=doc,token='range2')
    collection=type(doc.Tables)
    doc.Paragraphs=collection([SimpleNamespace(Range=first),SimpleNamespace(Range=second)])
    ref=registry.register(second,'range')
    args,_=registry.prepare_call('apply_structural_op',[{'op':'insert','type':'paragraph','position':{'before':ref},'props':{'text':'new'}}],{})
    assert args[0]['position']=={'before':'paragraph:2'}
    doc.Paragraphs.values[1].Range=SimpleNamespace(Start=6,End=12)
    second.End=9
    with pytest.raises(ValueError,match='paragraph'):
        registry.prepare_call('apply_structural_op',[{'op':'insert','position':{'after':ref}}],{})


def test_client_handle_is_opaque_and_rejects_cross_session(transport):
    m, _, _=transport
    with m.ProxyWordSession.open_document('/a.docx') as first:
        handle=m.NativeSessionHandle(first,'a'*32,'table')
        for name in ('Range','Font','Delete','_native'):
            with pytest.raises(AttributeError):getattr(handle,name)
        # Another component uses a different lock, without a second Word app.
        with m.ProxyExcelSession.open_document('/b.xlsx') as second:
            count=second._sequence
            with pytest.raises(ValueError,match='session'):second.apply_format_patch(handle,text='x')
            assert second._sequence==count


@pytest.mark.parametrize('changed', [True, False])
def test_error_clipboard_flag_crosses_worker_and_proxy_strictly(transport, changed):
    m, _, _ = transport
    _, w=modules()
    error=RuntimeError('native move failed');error.clipboard_changed=changed
    detail=w._error(error)
    assert detail['clipboard_changed'] is changed
    with m.ProxyWordSession.open_document('/source.docx') as session:
        assert session._remote_error(detail).clipboard_changed is changed
        detail['clipboard_changed']='true'
        with pytest.raises(ValueError):session._remote_error(detail)


def test_word_identity_uses_captured_identity_dataclass():
    _, w=modules()
    session=SimpleNamespace(staging_root=Path('/private'),_attached=False,
        _composer=SimpleNamespace(identity=SimpleNamespace(pid=314,executable='C:\\Office\\WINWORD.EXE')))
    assert w._identity(session)['office_pid']==314
    assert w._identity(session)['office_executable'].endswith('WINWORD.EXE')


def test_symlink_component_root_rejected_before_worker_launch(transport, tmp_path):
    m, _, children=transport
    real=tmp_path/'elsewhere';real.mkdir()
    (tmp_path/'writer').symlink_to(real,target_is_directory=True)
    with pytest.raises(ValueError,match='symlink'):m.ProxyWordSession.open_document('/source.docx')
    assert children==[]


def test_full_response_queue_does_not_leave_reader_thread(transport, tmp_path):
    m, _, _=transport
    import queue
    session=m.ProxyWordSession.__new__(m.ProxyWordSession)
    session.staging_root=tmp_path
    session._responses=queue.Queue(maxsize=2)
    session._stop=threading.Event()
    session._reader_error=None
    # Use the actual reader loop with a deliberately flooded, bounded stream.
    session._child=SimpleNamespace(stdout=io.BytesIO(b'{}\n'*5))
    reader=threading.Thread(target=session._read_responses,daemon=True)
    reader.start();reader.join(timeout=.2)
    assert not reader.is_alive()


def request(number, method, *, kind='writer', args=None, kwargs=None):
    return {'protocol': 1, 'id': number, 'kind': kind, 'method': method,
            'args': args or [], 'kwargs': kwargs or {}, 'deadline': TEST_DEADLINE,
            'remaining_seconds': 500}


def test_real_worker_protocol_uses_one_injected_session_and_no_raw_objects(tmp_path):
    _, w = modules()
    calls = []
    class Session:
        staging_root = tmp_path
        def inspect_document(self): return {'kind':'writer', 'counts':{'paragraphs':3}}
        def close(self, save_changes=False): calls.append(('close', save_changes))
    def factory(kind, method, args, kwargs):
        calls.append((kind, method, args, kwargs));return Session()
    frames = [request(1, 'open_document', args=['source.docx']), request(2, 'inspect_document'), request(3, 'close')]
    output = io.BytesIO()
    w.serve(io.BytesIO(b''.join(json.dumps(x).encode()+b'\n' for x in frames)), output, tmp_path, factory=factory)
    responses = [json.loads(line) for line in output.getvalue().splitlines()]
    assert [r['id'] for r in responses] == [1,2,3]
    assert responses[1]['value']['counts']['paragraphs'] == 3
    assert responses[2]['value'] == {'closed': True}
    assert calls == [('writer','open_document',['source.docx'],{}), ('close',False)]


def test_worker_rejects_attribute_escape_and_extended_deadline(tmp_path):
    _, w = modules()
    calls=[]
    class Session:
        staging_root=tmp_path
        def close(self,save_changes=False):calls.append('close')
        def __getattr__(self,name):calls.append(name);raise AssertionError('must not resolve arbitrary attributes')
    frames=[request(1,'open_document'), request(2,'__getattribute__',args=['_composer'])]
    output=io.BytesIO()
    w.serve(io.BytesIO(b''.join(json.dumps(x).encode()+b'\n' for x in frames)),output,tmp_path,factory=lambda *a:Session())
    response=json.loads(output.getvalue().splitlines()[-1])
    assert response['status']=='error'
    assert '__getattribute__' not in calls
    output=io.BytesIO();frames[1]=request(2,'inspect_document');frames[1]['deadline']+=1
    w.serve(io.BytesIO(b''.join(json.dumps(x).encode()+b'\n' for x in frames)),output,tmp_path,factory=lambda *a:Session())
    assert json.loads(output.getvalue().splitlines()[-1])['status']=='error'


def test_actual_worker_serve_subprocess_preserves_error_metadata_and_close(transport, monkeypatch):
    m, _, children=transport
    code='''import sys,json
from pathlib import Path
from skills.WPSComposer.scripts.msoffice.windows_session_worker import serve
job=Path(sys.argv[1])
class Session:
 def __init__(self): self.staging_root=job
 def inspect_document(self): return {'kind':'writer','counts':{'paragraphs':2}}
 def apply_structural_op(self,op):
  error=RuntimeError('move failed after native copy');error.clipboard_changed=True;raise error
 def close(self,save_changes=False): (job/'native-close.json').write_text(json.dumps({'save_changes':save_changes}))
def factory(*args):
 print('native diagnostic does not enter protocol stdout')
 return Session()
serve(sys.stdin.buffer,sys.stdout.buffer,job,factory=factory)
'''
    def launch(job):
        child=subprocess.Popen([sys.executable,'-u','-c',code,str(job)],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL)
        children.append(child);return child
    monkeypatch.setattr(m,'_spawn_worker',launch)
    with m.ProxyWordSession.open_document('/source.docx') as session:
        assert session.inspect_document()['counts']['paragraphs']==2
        with pytest.raises(RuntimeError) as error:session.apply_structural_op({'op':'move','target':'shape:1','to':'end'})
        assert error.value.clipboard_changed is True
    assert json.loads((session.staging_root/'native-close.json').read_text())=={'save_changes':False}
    assert not session._reader.is_alive()


def test_default_worker_factory_hands_exact_deadline_to_native_session(monkeypatch):
    _,w=modules()
    from skills.WPSComposer.scripts.msoffice import windows_document_api
    recorded=[]
    monkeypatch.setattr(windows_document_api.WindowsWordSession,'open_document',lambda *a,**k:recorded.append(k) or object())
    w.default_factory('writer','open_document',['source.docx'],{'read_only':True},deadline=12345.0)
    assert recorded==[{'read_only':True,'_deadline':12345.0}]


def test_handles_roundtrip_real_worker_with_com_object_return_doubles(transport, monkeypatch):
    m, _, children=transport
    code='''import sys,json
from pathlib import Path
from types import SimpleNamespace
from skills.WPSComposer.scripts.msoffice.windows_session_worker import serve
job=Path(sys.argv[1])
class Collection:
 def __init__(self,values):self.values=list(values)
 @property
 def Count(self):return len(self.values)
 def Item(self,index):return self.values[index-1]
class Session:
 kind='sheet'
 def __init__(self):
  self.staging_root=job
  self.sheet=SimpleNamespace(token='owned-sheet',Name='Original')
  self._composer=SimpleNamespace(_doc=SimpleNamespace(Worksheets=Collection([self.sheet])),_deps=SimpleNamespace(identity=lambda o:o.token))
 def _verify(self):pass
 def select_sheet(self,index):return self._composer._doc.Worksheets.Item(index)
 def rename_sheet(self,index,name):
  obj=self._composer._doc.Worksheets.Item(index);obj.Name=name
  (job/'renamed.json').write_text(json.dumps({'index':index,'name':name}))
  return obj
 def inspect_document(self):return {'name':self.sheet.Name}
 def apply_structural_op(self,op):self._composer._doc.Worksheets.values=[];return {'removed':op['target']}
 def close(self,save_changes=False):pass
serve(sys.stdin.buffer,sys.stdout.buffer,job,factory=lambda *a:Session())
'''
    def launch(job):
        child=subprocess.Popen([sys.executable,'-u','-c',code,str(job)],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL)
        children.append(child);return child
    monkeypatch.setattr(m,'_spawn_worker',launch)
    with m.ProxyExcelSession.new_document() as session:
        sheet=session.select_sheet(1)
        assert isinstance(sheet,m.NativeSessionHandle) and sheet.type=='sheet'
        renamed=session.rename_sheet(sheet,'Changed')
        assert isinstance(renamed,m.NativeSessionHandle)
        assert session.inspect_document()['name']=='Changed'
        assert session.apply_structural_op({'op':'remove','target':sheet})=={'removed':'sheet:1'}
        with pytest.raises(ValueError,match='stale'):session.rename_sheet(sheet,'Must not apply')
        assert json.loads((session.staging_root/'renamed.json').read_text())=={'index':1,'name':'Changed'}


def test_handle_capacity_and_table_target_preflight():
    registry,_,_,_,table,_=native_registry('writer')
    ref=registry.encode_result('add_table',table,[],{})
    args,_=registry.prepare_call('apply_structural_op',[{'op':'remove','target':ref}],{})
    assert args==[{'op':'remove','target':'table:1'}]
    registry.records.update({str(i):(table,'table') for i in range(1024)})
    with pytest.raises(ValueError,match='before mutation'):registry.prepare_call('add_table',[1,1,[['x']]],{})
