"""Try a task-owned ROT registration; never mutate the prior registered app."""
from pathlib import Path
import sys,json,subprocess,traceback,time,datetime
import pythoncom,win32com.client
ROOT=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(ROOT/'fixtures/msoffice_spike'))
from windows_word import word_processes,application_identity
OUT=Path(__file__).resolve().parent
report={'legacy_application_mutated':False,'quit_or_kill_called':False,'errors':[]}
cookie=None
pythoncom.CoInitialize()
try:
    prior=win32com.client.GetActiveObject('Word.Application')
    before=word_processes()
    owned=win32com.client.DispatchEx('Word.Application')
    identity=application_identity(owned,before)
    if owned.Documents.Count!=0 or owned._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==prior._oleobj_.QueryInterface(pythoncom.IID_IUnknown):
        raise RuntimeError('Task application is not empty and distinct')
    report['task_registered_application_candidate']=identity
    cookie=pythoncom.RegisterActiveObject(owned._oleobj_,pythoncom.MakeIID('{000209FF-0000-0000-C000-000000000046}'),pythoncom.ACTIVEOBJECT_WEAK)
    report['task_rot_cookie']=cookie
    active=win32com.client.GetActiveObject('Word.Application')
    target_ok=active._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==owned._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
    report['get_active_object_matches_new_owned_instance']=target_ok
    if not target_ok:
        raise RuntimeError('GetActiveObject did not select the new task instance; sentinel was not launched')
    command=[sys.executable,str(ROOT/'fixtures/msoffice_spike/windows_unsaved_sentinel.py'),'--evidence-dir',str(ROOT/'build/msoffice-spike/windows-sentinel-04'),'--output-dir',str(ROOT/'build/msoffice-spike/windows-run-11')]
    report['command']=command;report['started_utc']=datetime.datetime.now(datetime.timezone.utc).isoformat()
    (OUT/'isolated-registration.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    start=time.monotonic()
    with (OUT/'sentinel-wrapper.stdout.log').open('wb') as stdout,(OUT/'sentinel-wrapper.stderr.log').open('wb') as stderr:
        result=subprocess.run(command,stdout=stdout,stderr=stderr)
    report['wrapper_exit_code']=result.returncode;report['duration_seconds']=round(time.monotonic()-start,3)
    s=json.loads((ROOT/'build/msoffice-spike/windows-sentinel-04/sentinel-result.json').read_text('utf-8'))
    report['sentinel_used_new_owned_instance']=s.get('sentinel_application',{}).get('pid')==identity['pid']
    report['overall']='passed' if result.returncode==0 and report['sentinel_used_new_owned_instance'] else 'failed'
except Exception:
    report['errors'].append(traceback.format_exc());report['overall']='blocked'
finally:
    if cookie is not None:
        pythoncom.RevokeActiveObject(cookie)
        report['own_rot_registration_revoked']=True
    (OUT/'isolated-registration.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
raise SystemExit(0 if report.get('overall')=='passed' else 1)
