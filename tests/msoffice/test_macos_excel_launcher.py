from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import time

import pytest

ROOT = Path(__file__).resolve().parents[2]
pytestmark = pytest.mark.skipif(sys.platform == 'win32', reason='Inherited socketpair helper transport is POSIX; pure protocol checks run separately')


@pytest.fixture(autouse=True)
def prohibit_native_side_effects(monkeypatch):
    from skills.WPSComposer.scripts.msoffice import macos_excel_process, macos_osa_transport, macos_office_runtime
    def forbidden(*args, **kwargs):
        raise AssertionError('Native access prohibited by test isolation guard')
    monkeypatch.setattr(macos_excel_process.AppKitExcelLauncher, '__init__', forbidden)
    monkeypatch.setattr(macos_osa_transport._DarwinRuntime, '__init__', forbidden)
    monkeypatch.setattr(macos_osa_transport.BoundOSAKitTransport, '__init__', forbidden)
    monkeypatch.setattr(macos_osa_transport.BoundOSAKitTransport, 'run', forbidden)
    monkeypatch.setattr(macos_office_runtime, '_container_root', forbidden)


def api():
    from skills.WPSComposer.scripts.msoffice import macos_excel_launcher
    return macos_excel_launcher


@pytest.fixture
def helper_script(tmp_path):
    path = tmp_path / 'fake_launcher.py'
    path.write_text('''import sys, time, os
from pathlib import Path
sys.path.insert(0, ''' + repr(str(ROOT)) + ''')
from skills.WPSComposer.scripts.msoffice.macos_excel_launcher import helper_main
from skills.WPSComposer.scripts.msoffice.macos_excel_process import LaunchedExcel
from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity
mode=sys.argv.pop(1)
identity=ExcelProcessIdentity(4242,100,250,'/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel','com.microsoft.Excel')
class Native:
 def __init__(self, *, clock, observer=None):
  self.observer=observer
  self.record=None
  if mode=='hello_block':time.sleep(60)
 def existing_pids(self):return {9000}
 def launch(self, path, deadline):
  if self.observer:self.observer('candidate', {'pid':4242})
  if mode=='launch_block':time.sleep(60)
  self.record=LaunchedExcel(identity,99.985,'held-native-handle')
  return self.record
 def snapshot(self,pid):
  if mode=='probe_block':time.sleep(60)
  if mode=='reused':return ExcelProcessIdentity(4242,200,250,identity.executable,identity.bundle_id)
  return None if mode=='exited' else identity
 def finished_launching(self,record):return True
 def has_terminated(self,record):return mode=='exited'
 def wait_until(self,deadline):pass
sys.exit(helper_main(sys.argv[1:], launcher_factory=Native))
''')
    return path


def client(tmp_path, helper_script, mode='normal', seconds=2):
    return api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher',
        deadline=time.monotonic()+seconds,
        helper_command=[sys.executable, str(helper_script), mode])


def test_real_helper_blocks_discovery_but_parent_returns_by_deadline(tmp_path, helper_script):
    started=time.monotonic()
    launch=client(tmp_path,helper_script,'hello_block',seconds=.3)
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
        launch.existing_pids()
    assert time.monotonic()-started < 2
    assert launch.failed
    assert launch.process.poll() is not None
    assert (launch.evidence_dir/'intent.json').is_file()


def test_real_blocked_launch_preserves_candidate_and_does_not_retry_or_touch_other_child(tmp_path, helper_script):
    witness=subprocess.Popen([sys.executable,'-c','import time;time.sleep(60)'])
    launch=client(tmp_path,helper_script,'launch_block',seconds=.5)
    try:
        assert launch.existing_pids()=={9000}
        with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
            launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+.4)
        assert launch.recovery_candidate['pid']==4242
        assert launch.process.poll() is not None
        assert witness.poll() is None
        assert json.loads((launch.evidence_dir/'launching.json').read_text())['operation']=='LAUNCH'
        with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_UNAVAILABLE'):
            launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+5)
    finally:
        witness.kill();witness.wait()
        launch.abort()


def test_real_blocked_probe_is_bounded_and_keeps_exact_identity(tmp_path, helper_script):
    launch=client(tmp_path,helper_script,'probe_block',seconds=2)
    try:
        launch.existing_pids()
        record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
        launch.set_deadline(time.monotonic()+.2)
        with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
            launch.snapshot(record.identity.pid)
        assert launch.recovery_candidate['identity']['start_microseconds']==250
        assert launch.failed
    finally:launch.abort()


def test_real_helper_full_identity_probe_duplicate_launch_and_explicit_release(tmp_path, helper_script):
    launch=client(tmp_path,helper_script)
    try:
        assert launch.existing_pids()=={9000}
        record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
        assert record.handle != 'held-native-handle'
        assert launch.snapshot(4242)==record.identity
        assert launch.finished_launching(record)
        assert not launch.has_terminated(record)
        with pytest.raises(api().LauncherError,match='EXCEL_LAUNCH_ALREADY_REQUESTED'):
            launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
        with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_NOT_RELEASED'):
            launch.release()
    finally:launch.abort()


def test_release_only_after_exact_birth_gone_and_retained_handle_terminated(tmp_path, helper_script):
    launch=client(tmp_path,helper_script,'exited')
    launch.existing_pids()
    record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    assert launch.snapshot(record.identity.pid) is None
    assert launch.has_terminated(record)
    launch.release()
    assert launch.process.poll() is not None


def test_peer_eof_exits_helper_without_application_action(tmp_path, helper_script):
    launch=client(tmp_path,helper_script)
    launch.existing_pids()
    launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    launch.socket.close()
    assert launch.process.wait(timeout=2)==0


def test_local_process_relative_clock_maps_to_shared_clock_for_real_helper(tmp_path, helper_script):
    offset=time.monotonic()
    launch=api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher', deadline=2,
        clock=lambda:time.monotonic()-offset,
        helper_command=[sys.executable,str(helper_script),'normal'])
    try:
        assert launch.existing_pids()=={9000}
        request=json.loads((launch.evidence_dir/'request-1.json').read_text())
        assert request['version']==1
        assert request['deadline'] > 100
    finally:launch.abort()


@pytest.mark.parametrize('mode', ['wrong_nonce','wrong_id','oversized','partial'])
def test_malformed_or_partial_real_helper_reply_quarantines_without_new_request(tmp_path, mode):
    script=tmp_path/'bad.py'
    script.write_text('''import sys,socket,struct,json,time,os
s=socket.socket(fileno=int(sys.argv[2]))
n=struct.unpack('!I',s.recv(4))[0]
b=b''
while len(b)<n:b+=s.recv(n-len(b))
r=json.loads(b)
mode='''+repr(mode)+'''
if mode=='oversized':s.sendall(struct.pack('!I',9999999))
elif mode=='partial':s.sendall(struct.pack('!I',100)+b'{');time.sleep(60)
else:
 r={'version':1,'nonce':r['nonce'],'id':r['id'],'helper_pid':os.getpid(),'status':'ok','data':{'pids':[9000]}}
 if mode=='wrong_nonce':r['nonce']='wrong'
 if mode=='wrong_id':r['id']+=1
 b=json.dumps(r).encode();s.sendall(struct.pack('!I',len(b))+b)
time.sleep(60)
''')
    launch=api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher',deadline=time.monotonic()+.3,
        helper_command=[sys.executable,str(script)])
    with pytest.raises(api().LauncherError):launch.existing_pids()
    assert launch.failed
    assert launch.process.poll() is not None
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_UNAVAILABLE'):launch.existing_pids()
    assert len(list(launch.evidence_dir.glob('request-*.json')))==1


def test_lost_reply_preserves_durable_candidate_without_masking_peer_error(tmp_path, helper_script):
    source=helper_script.read_text().replace("if mode=='launch_block':time.sleep(60)","if mode=='launch_block':time.sleep(60)\n  if mode=='lost_reply':os._exit(0)")
    helper_script.write_text(source)
    launch=client(tmp_path,helper_script,'lost_reply')
    launch.existing_pids()
    with pytest.raises(EOFError):launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    assert launch.failed
    assert launch.recovery_candidate['pid']==4242
    assert (launch.evidence_dir/'launching.json').exists()


def test_failure_reading_candidate_never_masks_original_rpc_timeout(tmp_path, helper_script, monkeypatch):
    launch=client(tmp_path,helper_script,'launch_block')
    launch.existing_pids()
    original=Path.read_text
    def read(path,*args,**kwargs):
        if path.name=='candidate.json':raise PermissionError('evidence unreadable')
        return original(path,*args,**kwargs)
    monkeypatch.setattr(Path,'read_text',read)
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_TIMEOUT'):
        launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+.3)
    assert launch.process.poll() is not None
    assert launch.recovery_candidate['pid']==4242


def test_helper_rejects_foreign_identity_token_before_native_probe(tmp_path, helper_script):
    launch=client(tmp_path,helper_script)
    launch.existing_pids()
    record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_BAD_REQUEST'):
        launch._rpc('PROBE',{'token':'foreign','identity':record.identity.to_request(shared_deadline=1)})
    assert launch.failed


def test_parent_disconnect_while_launch_ffi_blocks_exits_helper_without_office_action(tmp_path, helper_script):
    launch=client(tmp_path,helper_script,'launch_block')
    launch.existing_pids()
    request={'version':1,'nonce':launch.nonce,'id':2,'operation':'LAUNCH',
        'deadline':api().shared_clock()+20,'data':{'app_path':'/Applications/Microsoft Excel.app'}}
    api()._transfer(launch.socket,request,time.monotonic()+1,time.monotonic)
    reply=api()._transfer(launch.socket,None,time.monotonic()+1,time.monotonic,receive=True)
    assert reply['status']=='progress'
    launch.socket.close()
    launch._liveness.close()
    try:
        assert launch.process.wait(timeout=2)==0
    finally:launch.abort()


def test_live_reused_birth_is_rejected_by_helper_probe(tmp_path, helper_script):
    launch=client(tmp_path,helper_script,'reused')
    launch.existing_pids()
    record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    try:
        with pytest.raises(api().LauncherError,match='OSA_PROCESS_IDENTITY_CHANGED'):
            launch.snapshot(record.identity.pid)
        assert launch.failed
    finally:launch.abort()


def test_completed_probe_deadline_does_not_expire_helper_session(tmp_path, helper_script):
    launch=client(tmp_path,helper_script,seconds=3)
    try:
        launch.existing_pids()
        record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+2)
        launch.set_deadline(time.monotonic()+.1)
        assert launch.snapshot(record.identity.pid)==record.identity
        time.sleep(.15)
        launch.set_deadline(time.monotonic()+1)
        assert launch.snapshot(record.identity.pid)==record.identity
    finally:launch.abort()


def test_helper_ipc_descriptors_are_not_inherited_by_subsequent_children(tmp_path,helper_script):
    helper_script.write_text(helper_script.read_text().replace("  self.observer=observer", "  assert not os.get_inheritable(int(sys.argv[2]))\n  assert not os.get_inheritable(int(sys.argv[5]))\n  self.observer=observer"))
    launch=client(tmp_path,helper_script)
    try:assert launch.existing_pids()=={9000}
    finally:launch.abort()


def test_abort_reports_exact_helper_that_remains_alive_after_kill(tmp_path):
    class StuckProcess:
        pid = 777
        def __init__(self):self.killed=False
        def poll(self):return None
        def kill(self):self.killed=True
        def wait(self, timeout):raise subprocess.TimeoutExpired('helper', timeout)

    launch=api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher',
        deadline=time.monotonic()+1,
        helper_command=[sys.executable,'-c','raise SystemExit(0)'])
    launch.process=StuckProcess()
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_HELPER_STILL_RUNNING'):
        launch.abort()
    assert launch.process.killed


def test_abort_accepts_helper_that_exits_at_final_poll(tmp_path):
    class JustExitedProcess:
        pid = 779
        def __init__(self):self.polls=0
        def poll(self):
            self.polls+=1
            return None if self.polls == 1 else -9
        def kill(self):pass
        def wait(self, timeout):raise subprocess.TimeoutExpired('helper', timeout)

    launch=api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher',
        deadline=time.monotonic()+1,
        helper_command=[sys.executable,'-c','raise SystemExit(0)'])
    launch.process=JustExitedProcess()
    launch.abort()
    assert not launch.cleanup_errors


def test_release_waits_for_graceful_helper_exit_without_killing(tmp_path,helper_script):
    launch=client(tmp_path,helper_script,'exited')
    launch.existing_pids()
    record=launch.launch(Path('/Applications/Microsoft Excel.app'),time.monotonic()+1)
    assert launch.snapshot(record.identity.pid) is None
    assert launch.has_terminated(record)
    process=launch.process
    killed=[]
    original_kill=process.kill
    process.kill=lambda: (killed.append(True),original_kill())[1]
    launch.release()
    assert process.returncode == 0
    assert not killed


def test_release_timeout_is_structured_and_kills_only_exact_helper(tmp_path):
    class SlowExitProcess:
        pid = 778
        returncode = None
        def __init__(self):self.killed=False
        def poll(self):return self.returncode
        def kill(self):self.killed=True;self.returncode=-9
        def wait(self, timeout):
            if not self.killed:raise subprocess.TimeoutExpired('helper',timeout)
            return self.returncode

    launch=api().BoundedExcelLauncher(evidence_dir=tmp_path/'launcher',
        deadline=time.monotonic()+1,
        helper_command=[sys.executable,'-c','raise SystemExit(0)'])
    owned=api().ExcelProcessIdentity(4242,100,250,
        '/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel','com.microsoft.Excel')
    launch._record=api().LaunchedExcel(owned,99.985,'token')
    launch.process=SlowExitProcess()
    launch._rpc=lambda *args,**kwargs:{}
    with pytest.raises(api().LauncherError,match='EXCEL_LAUNCHER_HELPER_EXIT_TIMEOUT'):
        launch.release()
    assert launch.process.killed
    assert launch.failed
