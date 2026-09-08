import sys, json, traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import pythoncom, win32com.client, win32process
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes
pythoncom.CoInitialize()
before=_word_processes()
app=win32com.client.DispatchEx('Word.Application')
after=_word_processes()
report={'new_processes':{p:v for p,v in after.items() if p not in before},'name':str(app.Name),'documents_before':app.Documents.Count,'chains':[],'documents_added':False,'quit_attempted':False}
for bar_index in [1,2]:
    obj=app.CommandBars.Item(bar_index)
    chain=[]
    for depth in range(12):
        row={'depth':depth}
        try:
            hwnd=obj._oleobj_.QueryInterface(pythoncom.IID_IOleWindow).GetWindow()
            row.update(hwnd=hwnd,pid=win32process.GetWindowThreadProcessId(hwnd)[1])
        except Exception as exc: row['hwnd_error']=repr(exc)
        chain.append(row)
        try:
            obj=obj.accParent
            if obj is None: break
        except Exception as exc:
            row['parent_error']=repr(exc)
            break
    report['chains'].append({'bar_index':bar_index,'chain':chain})
report['documents_after']=app.Documents.Count
Path(__file__).with_name('empty-accessible-probe.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
