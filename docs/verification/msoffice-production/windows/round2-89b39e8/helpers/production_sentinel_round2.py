import sys,subprocess,json,hashlib,uuid,time,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import pythoncom,win32com.client,win32process
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes, _windows_path
root=Path('build/msoffice-production/sentinel-02').resolve();root.mkdir(exist_ok=False)
report={'status':'RUNNING','registered_quit_attempted':False,'sentinel_saved':False,'errors':[]}
def persist(): (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
def state(doc):
    text=doc.Content.Text
    return {'name':doc.Name,'path':doc.FullName,'saved':bool(doc.Saved),'characters':len(text),'text_sha256':hashlib.sha256(text.encode()).hexdigest()}
sentinel=None;initial=None;before=None
try:
    pythoncom.CoInitialize();app=win32com.client.GetActiveObject('Word.Application')
    assert app.Name=='Microsoft Word' and app.Documents.Count==0,'Registered Word must be empty'
    report['registered_documents_before']=0
    sentinel=app.Documents.Add();initial=state(sentinel)
    hwnd=int(sentinel.Windows.Item(1).Hwnd);pid=win32process.GetWindowThreadProcessId(hwnd)[1]
    image=_word_processes().get(pid)
    assert image and _windows_path(Path(image).parent)==_windows_path(app.Path)
    report.update(registered_pid=pid,sentinel_hwnd=hwnd)
    marker='WPSCOMPOSER-PRODUCTION-SENTINEL-'+uuid.uuid4().hex
    sentinel.Content.Text=marker
    assert sentinel.Content.Text==marker+'\r' and not sentinel.Saved
    before=state(sentinel);report['before']=before;persist()
    commands=[['representative-02','fixtures/verify_msoffice_production.py','--output-root','build/msoffice-production/representative-02','--fixture','representative','--engine','msoffice','--timeout','600']]
    report['runs']=[]
    for command in commands:
        result=subprocess.run([sys.executable,'build/msoffice-production/run_logged.py']+command,capture_output=True)
        (root/(command[0]+'.stdout.log')).write_bytes(result.stdout)
        (root/(command[0]+'.stderr.log')).write_bytes(result.stderr)
        after=state(sentinel)
        report['runs'].append({'name':command[0],'exit_code':result.returncode,'after':after,'unchanged':after==before})
        persist()
        assert after==before,'Sentinel changed'
        if result.returncode:raise RuntimeError(command[0]+' failed')
    report['status']='PASS'
except Exception as exc:
    report['errors'].append({'error':repr(exc),'traceback':traceback.format_exc()});report['status']='FAIL'
finally:
    if sentinel is not None:
        try:
            current=state(sentinel)
            if current!=before and current!=initial:raise RuntimeError('Refusing to close changed sentinel')
            sentinel.Close(SaveChanges=0)
            report['sentinel_closed_without_saving']=True
            report['registered_documents_after']=app.Documents.Count
            report['registered_responsive']=app.Name=='Microsoft Word'
            report['registered_pid_retained']=report.get('registered_pid') in _word_processes()
        except Exception as exc:
            report['status']='FAIL';report['errors'].append({'cleanup_error':repr(exc)})
    persist()
print(json.dumps(report,ensure_ascii=False))
raise SystemExit(0 if report['status']=='PASS' else 1)
