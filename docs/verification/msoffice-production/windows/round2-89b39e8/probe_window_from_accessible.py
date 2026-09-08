"""Read-only COM capability probe; does not add documents or set app properties."""
import sys,json,ctypes,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
sys.path.insert(0,str(Path(__file__).with_name('deps')))
import comtypes,comtypes.client,comtypes.automation
import win32process
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes
class IAccessible(comtypes.automation.IDispatch):
    _iid_=comtypes.GUID('{618736E0-3C3D-11CF-810C-00AA00389B71}')
    _methods_=[]
def identity(obj):
    pointer=obj.QueryInterface(comtypes.IUnknown)
    return ctypes.cast(pointer,ctypes.c_void_p).value
library=ctypes.OleDLL('oleacc')
library.WindowFromAccessibleObject.argtypes=[ctypes.POINTER(IAccessible),ctypes.POINTER(ctypes.c_void_p)]
library.WindowFromAccessibleObject.restype=ctypes.c_long
before=_word_processes()
app=comtypes.client.CreateObject('Word.Application',dynamic=True)
new={p:v for p,v in _word_processes().items() if p not in before}
report={'new_processes':new,'application_name':app.Name,'documents_before':app.Documents.Count,'visible':bool(app.Visible),'documents_added':False,'caption_or_visibility_changed':False,'quit_attempted':False,'bars':[],'note':'comtypes supplies typed IAccessible pointer for oleacc; same native COM interfaces, no user-home generated wrapper module required'}
assert len(new)==1 and app.Name=='Microsoft Word' and app.Documents.Count==0
pid=next(iter(new))
for name in [1,2,'Ribbon','Menu Bar','Status Bar']:
    row={'bar':name};report['bars'].append(row)
    try:
        bar=app.CommandBars.Item(name)
        row['canonical_application_iunknown_matches']=identity(bar.Application)==identity(app)
        assert row['canonical_application_iunknown_matches']
        accessible=bar.QueryInterface(IAccessible)
        row['iaccessible_query_succeeded']=True
        hwnd=ctypes.c_void_p()
        result=library.WindowFromAccessibleObject(accessible,ctypes.byref(hwnd))
        row.update(hresult=result,hwnd=hwnd.value)
        if hwnd.value:
            row['pid']=win32process.GetWindowThreadProcessId(hwnd.value)[1]
            row['matches_new_pid']=row['pid']==pid
    except Exception as exc:
        row.update(error=repr(exc),traceback=traceback.format_exc())
report['documents_after']=app.Documents.Count
report['usable_binding_found']=any(r.get('matches_new_pid') for r in report['bars'])
out=Path(__file__).with_name('window-from-accessible-probe.json')
out.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False))
