"""Native capability probe for one shared multilevel template on Word and WPS."""
import sys,json,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import pythoncom,win32com.client
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer
from skills.WPSComposer.scripts.writer import WriterComposer
engine=sys.argv[1];root=Path('build/msoffice-production/linked-style-'+engine+'-01').resolve();root.mkdir(exist_ok=False)
report={'engine':engine,'states':{},'errors':[]}
def persist(): (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
c=None;doc=None
try:
    if engine=='word':
        c=create_dedicated_composer(str(root));doc=c.doc
        report['pid']=c.identity.pid;report['binding']=c.application_binding
    else:
        pythoncom.CoInitialize();app=win32com.client.DispatchEx('KWps.Application')
        assert Path(app.Path).resolve()==Path('C:/Users/1/AppData/Local/12.1.0.28505/office6').resolve() and app.Documents.Count==0
        doc=app.Documents.Add();assert doc.Application._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        c=WriterComposer.__new__(WriterComposer);c._app=app;c._doc=doc
        report['application']={'name':app.Name,'path':app.Path,'version':app.Version}
    template=doc.ListTemplates.Add(True)
    for i,fmt in enumerate(['%1','%1.%2','%1.%2.%3','%1.%2.%3.%4'],1):
        level=template.ListLevels(i);level.NumberFormat=fmt;level.NumberStyle=0
        level.NumberPosition=(i-1)*18;level.TextPosition=i*18;level.ResetOnHigher=0 if i==1 else i-1;level.StartAt=1
        level.LinkedStyle=doc.Styles(-1-i).NameLocal
    # Supply the real native template through the production composer's cache.
    # This probe isolates the proposed binding primitive; public runs follow.
    c._wpsc_heading_templates={'decimal':template}
    for i in range(1,5):c.add_heading_level_native('Original H'+str(i),i,numbering=True,scheme='decimal')
    def state():return [{'text':doc.Paragraphs(i).Range.Text.strip(),'number':str(doc.Paragraphs(i).Range.ListFormat.ListString)} for i in range(1,doc.Paragraphs.Count+1) if doc.Paragraphs(i).Range.Text.strip()]
    def save(name):
        if engine=='word':c.save_docx(root/name)
        else:doc.SaveAs(str(root/name),12)
    report['states']['before']=state();save('before.docx');persist()
    doc.Range(0,0).InsertBefore('Inserted H1\r');doc.Range(0,len('Inserted H1\r')).Style=doc.Styles(-2)
    c.update_fields();report['states']['inserted']=state();save('inserted.docx');persist()
    doc.Paragraphs(1).Range.Delete();c.update_fields();report['states']['deleted']=state();save('deleted.docx')
    report['checks']={'before':[p['number'] for p in report['states']['before']]==['1','1.1','1.1.1','1.1.1.1'],
                      'inserted':[p['number'] for p in report['states']['inserted']]==['1','2','2.1','2.1.1','2.1.1.1'],
                      'deleted':[p['number'] for p in report['states']['deleted']]==['1','1.1','1.1.1','1.1.1.1']}
    report['status']='PASS' if all(report['checks'].values()) else 'FAIL';persist()
except Exception as e:report['status']='FAIL';report['errors'].append({'error':repr(e),'traceback':traceback.format_exc()})
finally:
    if c is not None:
        if engine=='word':c.close(save_changes=False);report['closed']=c._closed
        elif doc is not None:doc.Close(SaveChanges=0);report['owned_document_closed']=True;report['quit_attempted']=False
    persist()
print(json.dumps(report,ensure_ascii=False))
