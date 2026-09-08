"""COM setup and readback for a CUA audit of a synthetic copy; never Quit."""
import sys, json, hashlib, traceback
from pathlib import Path
import pythoncom, win32com.client, win32process

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / 'fixtures/msoffice_spike'))
from windows_word import word_processes, application_identity
OUT = Path(__file__).resolve().parent
COPY = OUT / 'audit-ui-probe.docx'
pythoncom.CoInitialize()
existing = win32com.client.GetActiveObject('Word.Application')
if existing.Documents.Count != 0:
    raise RuntimeError('Registered instance has documents; refusing audit to avoid reading user data')
before = word_processes()
app = win32com.client.DispatchEx('Word.Application')
identity = application_identity(app, before)
if app._oleobj_.QueryInterface(pythoncom.IID_IUnknown) == existing._oleobj_.QueryInterface(pythoncom.IID_IUnknown) or app.Documents.Count != 0:
    raise RuntimeError('Cannot establish empty separate instance')
doc = app.Documents.Open(FileName=str(COPY), AddToRecentFiles=False, ReadOnly=False)
hwnd = int(doc.Windows.Item(1).Hwnd)
pid = win32process.GetWindowThreadProcessId(hwnd)[1]
if pid != identity['pid'] or Path(doc.FullName) != COPY:
    raise RuntimeError('Owned document identity mismatch; refusing app mutation')
app.Visible = True
report = {'transport': 'COM setup/readback plus native CUA actions recorded separately', 'identity': identity,
          'owned_hwnd': hwnd, 'owned_pid': pid, 'registered_before_count': 0, 'events': [], 'quit_called': False}

def state():
    if app.Documents.Count != 1:
        return {'owned_instance_document_count': app.Documents.Count}
    current = app.Documents.Item(1)
    if Path(current.FullName) != COPY:
        raise RuntimeError('Unexpected document; refusing content read')
    text = current.Content.Text
    tables = []
    for i in range(1,current.Tables.Count+1):
        t=current.Tables.Item(i)
        tables.append({'rows':t.Rows.Count,'columns':t.Columns.Count,'header':bool(t.Rows.Item(1).HeadingFormat)})
    return {'owned_instance_document_count':1,'path':str(COPY),'saved':bool(current.Saved),
            'text_sha256':hashlib.sha256(text.encode('utf-8')).hexdigest(),'characters':len(text),
            'audit_marker_present':'WPS-AUDIT-UI-' in text,'end_marker_present':'MSOFFICE-SPIKE-END' in text,
            'tables':tables,'toc_count':current.TablesOfContents.Count,'field_count':current.Fields.Count}

def record(label):
    data={'label':label,'state':state()};report['events'].append(data)
    (OUT/'com-observations.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(data,ensure_ascii=False),flush=True)

record('opened_copy_before_ui')
for line in sys.stdin:
    try:
        cmd=json.loads(line)
        if cmd['action']=='observe':record(cmd['label'])
        elif cmd['action']=='reopen':
            if app.Documents.Count != 0:raise RuntimeError('Refusing reopen while document remains open')
            doc=app.Documents.Open(FileName=str(COPY),AddToRecentFiles=False,ReadOnly=False)
            if win32process.GetWindowThreadProcessId(int(doc.Windows.Item(1).Hwnd))[1]!=pid:raise RuntimeError('Reopened PID changed')
            record('reopened_copy_via_COM')
        elif cmd['action']=='finish':
            if app.Documents.Count!=0:raise RuntimeError('Expected CUA to close task document first')
            report['registered_after_count']=existing.Documents.Count
            report['owned_documents_at_finish']=app.Documents.Count
            report['quit_called']=False
            (OUT/'com-observations.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
            print('Observer finished; no application Quit performed',flush=True);break
        else:raise ValueError('Unknown observer command')
    except Exception:
        raw=traceback.format_exc();print(raw,flush=True)
        report.setdefault('errors',[]).append(raw)
        (OUT/'com-observations.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
