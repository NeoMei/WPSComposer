"""Observe a completed UI reopen synchronously, then await explicit closed-UI signal."""
import json, hashlib, time, traceback
from pathlib import Path
import pythoncom, win32com.client, win32process
root=Path('build/msoffice-production/native-edit-03').resolve()
report_path=root/'ui-reopen-verification.json'
report=json.loads(report_path.read_text())
expected=root/'edited-pagination.docx'; app=None
def persist():report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
try:
    pythoncom.CoInitialize(); rot=pythoncom.GetRunningObjectTable(); ctx=pythoncom.CreateBindCtx(0)
    found=[]
    for moniker in rot.EnumRunning():
        if str(expected).casefold()==moniker.GetDisplayName(ctx,None).casefold():
            found.append(win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch)))
    assert len(found)==1
    doc=found[0]; app=doc.Application
    assert Path(doc.FullName).resolve()==expected
    assert win32process.GetWindowThreadProcessId(doc.Windows.Item(1).Hwnd)[1]==30192
    assert app.Documents.Count==1 and Path(app.Path).name=='Office16'
    token=app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
    report['pid']=30192; report['documents_after_reopen']=1
    baseline=json.loads((root/'report.json').read_text(encoding='utf-8'))
    rows=[]
    for old in baseline['snapshots'][-1]['headings']:
        r=doc.Sections(3).Range.Duplicate
        assert r.Find.Execute(FindText=old['text'],Forward=True,Wrap=0)
        p=r.Paragraphs(1); position=p.Range.Duplicate; position.Collapse(1)
        rows.append({'text':p.Range.Text.strip(),'number':str(p.Range.ListFormat.ListString),
                     'logical_page':int(position.Information(1)),'physical_page':int(position.Information(3))})
    report['headings_after_ui_reopen']=rows
    report['toc_after_ui_reopen']=doc.TablesOfContents.Item(1).Range.Text
    report['text_sha256_after_ui_reopen']=hashlib.sha256(doc.Content.Text.encode()).hexdigest()
    report['disk_sha256_after_ui_reopen']=hashlib.sha256(expected.read_bytes()).hexdigest()
    report['checks']={
        'ui_reopen_uses_original_owned_word_pid':report['pid']==report['closed_window_pid'],
        'ui_reopen_preserves_all_six_heading_numbers_and_pages':rows==baseline['snapshots'][-1]['headings'],
        'ui_reopen_preserves_updated_toc':report['toc_after_ui_reopen']==baseline['snapshots'][-1]['toc_text'],
        'ui_reopen_preserves_document_text':report['text_sha256_after_ui_reopen']==baseline['ui_snapshots'][0]['text_sha256'],
        'ui_reopen_preserves_saved_file_bytes':report['disk_sha256_after_ui_reopen']==report['disk_sha256_before_ui_reopen'],
    }
    assert all(report['checks'].values()),report['checks']
    report['status']='REOPEN_PASS_AWAITING_UI_CLOSE'; persist(); print(report['status'],flush=True)
    deadline=time.monotonic()+600
    while not (root/'finish-ui.json').exists():
        assert time.monotonic()<deadline,'UI close signal timeout'
        time.sleep(.2)
    assert app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==token
    assert app.Documents.Count==0,'Refusing Quit while documents remain open'
    report['documents_after_final_ui_close']=0
    app.Quit(SaveChanges=0)
    report['owned_application_quit_requested']=True
    report['status']='PASS'; persist()
except Exception as exc:
    report['status']='FAIL';report['error']=repr(exc);report['traceback']=traceback.format_exc();persist()
print(report['status'],flush=True)
raise SystemExit(0 if report['status']=='PASS' else 1)
