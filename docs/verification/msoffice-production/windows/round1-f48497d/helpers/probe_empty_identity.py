import sys, json, traceback
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import pythoncom, win32com.client, win32process
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes
out=Path(__file__).with_name('empty-identity-probe.json')
pythoncom.CoInitialize()
before=_word_processes()
app=win32com.client.DispatchEx('Word.Application')
after=_word_processes()
r={'new_processes':{p:v for p,v in after.items() if p not in before},'name':str(app.Name),'count':int(app.Documents.Count),'path':str(app.Path),'observations':{},'documents_added':False,'quit_attempted':False}
def check(name,call):
    try: r['observations'][name]={'value':call()}
    except Exception as exc: r['observations'][name]={'error':repr(exc),'traceback':traceback.format_exc()}
    out.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
check('Application.Hwnd',lambda:app.Hwnd)
check('Application.IOleWindow',lambda: app._oleobj_.QueryInterface(pythoncom.IID_IOleWindow).GetWindow())
check('WordBasic.AppInfo2',lambda:app.WordBasic.AppInfo(2))
check('CommandBars.Count',lambda:app.CommandBars.Count)
check('CommandBar.IOleWindow',lambda:app.CommandBars.Item(1)._oleobj_.QueryInterface(pythoncom.IID_IOleWindow).GetWindow())
print(out)
