import sys, json, uuid, traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import pythoncom, win32com.client, win32process, win32gui
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes, _windows_path
pythoncom.CoInitialize()
previous=win32com.client.GetActiveObject('Word.Application')
before=_word_processes()
app=win32com.client.DispatchEx('Word.Application')
new={p:v for p,v in _word_processes().items() if p not in before}
assert len(new)==1
pid,path=next(iter(new.items()))
assert app.Name=='Microsoft Word' and _windows_path(app.Path)==_windows_path(str(Path(path).parent)) and app.Documents.Count==0
assert app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)!=previous._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
r={'pid':pid,'documents_before':app.Documents.Count,'old_caption':app.Caption,'new_processes':new,'quit_attempted':False,'documents_added':False}
marker='WPSComposer-identity-'+uuid.uuid4().hex
def windows():
    found=[]
    def observe(hwnd,arg):
        if win32process.GetWindowThreadProcessId(hwnd)[1]==pid:
            found.append({'hwnd':hwnd,'class':win32gui.GetClassName(hwnd),'caption':win32gui.GetWindowText(hwnd)})
    win32gui.EnumWindows(observe,None)
    return found
r['windows_before']=windows()
try:
    app.Caption=marker
    r['challenge']=marker
    r['windows_during']=windows()
except Exception as e:r['error']=repr(e)
finally:
    if app.Caption==marker:app.Caption=r['old_caption']
r['windows_after']=windows();r['caption_after']=app.Caption;r['documents_after']=app.Documents.Count
Path(__file__).with_name('caption-identity-probe.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(r,ensure_ascii=False))
