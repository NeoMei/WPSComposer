from __future__ import annotations

import json
import subprocess
import time
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.msoffice import macos_excel_session as mod
from test_macos_excel_session import session, package


class Owner:
    def __init__(self):
        self.events = []
        self.is_private_owned = True
        self.recovery = None
        self.identity = None
        self.owned_path = None
        self.state = "owned"
    def start(self, *, deadline):
        self.events.append("start")
    def reserve_workbook(self, path):
        self.owned_path = path
        self.events.append("reserve")
    def claim_workbook(self, path, *, deadline):
        assert path == self.owned_path
        self.events.append("claim")
    def run(self, path, deadline):
        self.events.append(("run", path.read_text(), deadline))
        return subprocess.CompletedProcess([], 0, '{"ok":true}', '')
    def close(self, *, deadline):
        self.events.append("close")
    def quarantine(self, error):
        self.recovery = {"owned_path":str(self.owned_path),"code":"uncertain"}
        self.is_private_owned = False


def test_owned_transport_bypasses_application_name_executor(tmp_path):
    obj = session(tmp_path)
    obj._process_owner = owner = Owner()
    obj._execute = lambda *a, **k: pytest.fail("no bundle-name fallback")
    assert obj._run('return "{}"') == {"ok":True}
    assert owner.events[0][0] == "run"
    assert owner.events[0][2] <= obj._deadline


def test_private_transport_failure_records_owner_recovery_and_blocks_later_calls(tmp_path):
    obj = session(tmp_path)
    obj._process_owner = owner = Owner()
    def fail(*args):
        raise RuntimeError('uncertain')
    owner.run = fail
    obj._execute = lambda *a, **k: pytest.fail("no bundle-name fallback")
    with pytest.raises(RuntimeError, match='recovery required'):
        obj._run('return "{}"')
    assert not owner.is_private_owned
    assert json.loads((tmp_path/'recovery.json').read_text())['private_process'] == owner.recovery
    with pytest.raises(RuntimeError, match='quarantined'):
        obj._run('return "{}"')


def test_private_close_delegates_exact_close_and_quit_before_releasing_lock(tmp_path):
    obj = session(tmp_path)
    obj._process_owner = owner = Owner()
    obj._run = lambda *a, **k: pytest.fail('owner controls close/quit')
    class Lock:
        def close(self):
            assert owner.events == ['close']
    obj._lock = Lock()
    obj.close()
    assert obj._closed


def test_private_close_failure_quarantines_before_lock_release(tmp_path):
    obj = session(tmp_path)
    obj._process_owner = owner = Owner()
    def fail(**kwargs):
        raise RuntimeError('uncertain quit')
    owner.close = fail
    obj._run = lambda *a, **k: pytest.fail('owner controls close/quit')
    events=[]
    class Lock:
        def quarantine(self, detail):
            events.append('quarantine')
        def close(self):
            events.append('release')
    obj._lock = Lock()
    with pytest.raises(RuntimeError, match='uncertain quit'):
        obj.close()
    assert obj._failed and obj._closed
    assert events == ['quarantine','release']


def test_attached_session_keeps_existing_transport_and_never_closes_owner(tmp_path):
    obj = session(tmp_path)
    obj._attached = True
    obj._execute = lambda *a, **k: subprocess.CompletedProcess([],0,'{}','')
    assert obj._run('return "{}"') == {}
    obj.close()
    assert obj._closed


def test_populated_sheet_delete_requires_owned_process_and_strict_ack(tmp_path):
    obj = session(tmp_path)
    obj._process_owner = Owner()
    obj._process_owner.owned_path = obj._native
    obj._sheet_index = 3
    scripts=[]
    obj._run = lambda body, **kw: scripts.append((body,kw)) or {'removed':'sheet:2','deleted':True}
    result = obj.apply_structural_op({'op':'remove','target':'sheet:2'})
    assert result['removed'] == 'sheet:2'
    assert obj._sheet_index == 2
    assert scripts[0][1]['acknowledge'] == 'deleted'
    assert 'display alerts' in scripts[0][0]
    assert 'name of every worksheet' in scripts[0][0]


def test_attached_populated_sheet_deletion_remains_closed(tmp_path):
    obj = session(tmp_path)
    obj._attached = True
    obj._process_owner = Owner()  # It cannot authorize attached deletion.
    obj._run = lambda *a, **k: pytest.fail('attached cannot suppress alerts')
    with pytest.raises(NotImplementedError):
        obj.apply_structural_op({'op':'remove','target':'sheet:1'})


@pytest.mark.parametrize('create', [False, True])
def test_file_session_launches_and_claims_only_after_staging(tmp_path, monkeypatch, create):
    owner=Owner()
    events=owner.events
    source=package(tmp_path/'source.xlsx','original')
    class Lock:
        def __init__(self, root): pass
        def acquire(self, deadline): events.append('lock')
        def close(self): events.append('release')
    monkeypatch.setattr(mod,'OfficeJobLock',Lock)
    monkeypatch.setattr(mod,'_container_root',lambda component:tmp_path/'container')
    monkeypatch.setattr(mod,'validate_native_input',lambda *a,**kw:None)
    monkeypatch.setattr(mod,'PrivateExcelProcessOwner',lambda **kw:owner,raising=False)
    def run(self, body, **kwargs):
        assert self._process_owner is owner
        assert owner.owned_path==self._native
        assert owner.events == ['lock','start','reserve']
        if not create:assert self._native.read_bytes()==source.read_bytes()
        events.append('native-open')
        return {}
    monkeypatch.setattr(mod.MacExcelSession,'_run',run)
    obj=mod.MacExcelSession.new_document() if create else mod.MacExcelSession.open_document(source)
    assert events == ['lock','start','reserve','native-open','claim']
    assert obj._process_owner is owner


def test_startup_failure_persists_process_identity_before_release(tmp_path,monkeypatch):
    owner=Owner()
    owner.recovery={'identity':{'pid':123},'code':'STARTUP_FAILED'}
    def fail(**kwargs):raise RuntimeError('launch incomplete')
    owner.start=fail
    events=[]
    class Lock:
        def __init__(self,root):pass
        def acquire(self,deadline):pass
        def quarantine(self,detail):
            assert 'private_process' in detail
            events.append('quarantine')
        def close(self):events.append('release')
    monkeypatch.setattr(mod,'OfficeJobLock',Lock)
    monkeypatch.setattr(mod,'_container_root',lambda component:tmp_path)
    monkeypatch.setattr(mod,'PrivateExcelProcessOwner',lambda **kw:owner,raising=False)
    monkeypatch.setattr(mod.MacExcelSession,'_run',lambda *a,**kw:pytest.fail('native startup must not execute in unit test'))
    with pytest.raises(RuntimeError,match='launch incomplete'):
        mod.MacExcelSession.new_document()
    assert events==['quarantine','release']


@pytest.mark.parametrize('state,path', [('ready','match'),('owned','wrong')])
def test_existing_sheet_delete_rejects_unclaimed_or_wrong_workbook(tmp_path,state,path):
    obj=session(tmp_path)
    obj._process_owner=Owner()
    obj._process_owner.state=state
    obj._process_owner.owned_path=obj._native if path=='match' else tmp_path/'other.xlsx'
    obj._run=lambda *a,**kw:pytest.fail('wrong owner cannot delete')
    with pytest.raises(NotImplementedError):obj.apply_structural_op({'op':'remove','target':'sheet:2'})


def test_script_preparation_cannot_extend_step_deadline(tmp_path,monkeypatch):
    obj=session(tmp_path);obj._deadline=1000;obj._process_owner=Owner()
    clock=[10.0]
    monkeypatch.setattr(mod.time,'monotonic',lambda:clock[0])
    real_write=Path.write_text
    def slow_write(path,*args,**kw):
        if path.suffix=='.applescript':clock[0]+=5
        return real_write(path,*args,**kw)
    monkeypatch.setattr(Path,'write_text',slow_write)
    obj._run('return "{}"')
    assert obj._process_owner.events[0][2]==70.0


def test_owner_timeout_preserves_native_timeout_classification(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_osa_transport import OSATransportError
    obj=session(tmp_path);obj._process_owner=Owner()
    def fail(*a):raise OSATransportError('OSA_TIMEOUT',stdout='partial',outcome_uncertain=True)
    obj._process_owner.run=fail
    with pytest.raises(RuntimeError) as error:obj._run('return "{}"')
    assert error.value.code=='NATIVE_OFFICE_TIMEOUT'
    assert (tmp_path/'step-0001.stdout.log').read_text()=='partial'


def test_new_workbook_requires_native_saved_ack_and_package_validation(tmp_path,monkeypatch):
    owner=Owner();events=[]
    class Lock:
        def __init__(self,root):pass
        def acquire(self,deadline):pass
        def close(self):pass
    monkeypatch.setattr(mod,'OfficeJobLock',Lock)
    monkeypatch.setattr(mod,'_container_root',lambda c:tmp_path)
    monkeypatch.setattr(mod,'PrivateExcelProcessOwner',lambda **kw:owner)
    def run(self,body,**kw):
        assert kw['acknowledge']=='created'
        assert 'saved of ownedBook' in body
        events.append('native')
        return {'created':True}
    monkeypatch.setattr(mod.MacExcelSession,'_run',run)
    monkeypatch.setattr(mod,'validate_native_input',lambda *a,**kw:events.append('validated'))
    mod.MacExcelSession.new_document()
    assert events==['native','validated']


def test_owner_quarantine_failure_does_not_skip_session_recovery(tmp_path):
    obj=session(tmp_path);obj._process_owner=Owner()
    def fail(*args):raise OSError('owner diagnostic failed')
    obj._process_owner.quarantine=fail
    events=[]
    class Lock:
        def quarantine(self,detail):events.append('lock-quarantined')
    obj._lock=Lock()
    obj._quarantine({'code':'NATIVE_OFFICE_EXECUTION_FAILED'})
    assert obj._failed
    assert (tmp_path/'recovery.json').is_file()
    assert events==['lock-quarantined']
    assert ('owner_quarantine','OSError') in obj._diagnostic_io_failures


def test_creation_validation_failure_quarantines_owned_process(tmp_path,monkeypatch):
    owner=Owner()
    events=[]
    class Lock:
        def __init__(self,root):pass
        def acquire(self,deadline):pass
        def quarantine(self,detail):events.append('quarantine')
        def close(self):events.append('release')
    monkeypatch.setattr(mod,'OfficeJobLock',Lock)
    monkeypatch.setattr(mod,'_container_root',lambda c:tmp_path)
    monkeypatch.setattr(mod,'PrivateExcelProcessOwner',lambda **kw:owner)
    monkeypatch.setattr(mod.MacExcelSession,'_run',lambda *a,**kw:{'created':True})
    def invalid(*a,**kw):raise ValueError('not a valid native package')
    monkeypatch.setattr(mod,'validate_native_input',invalid)
    with pytest.raises(ValueError,match='native package'):mod.MacExcelSession.new_document()
    assert not owner.is_private_owned
    assert events==['quarantine','release']
