import sys,json,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import pythoncom,win32com.client
from skills.WPSComposer.scripts.writer import WriterComposer
root=Path('build/msoffice-production/wps-sections-probe-01').resolve();root.mkdir(exist_ok=False)
report={'events':[],'errors':[],'quit_attempted':False}
def persist(): (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
def trace(frame,kind,arg):
    if kind=='exception' and frame.f_code.co_name in ('set_page_numbering','set_header_footer','configure_section'):
        report['errors'].append({'function':frame.f_code.co_name,'line':frame.f_lineno,'error':repr(arg[1])});persist()
    return trace
doc=None
try:
    pythoncom.CoInitialize();app=win32com.client.GetActiveObject('KWps.Application')
    report['application']={'name':app.Name,'path':app.Path,'version':app.Version,'documents_before':app.Documents.Count}
    assert Path(app.Path).resolve()==Path('C:/Users/1/AppData/Local/12.1.0.28505/office6').resolve()
    assert app.Documents.Count==0,'Refusing to use an existing WPS document'
    token=app._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
    doc=app.Documents.Add()
    assert doc.Application._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==token and app.Documents.Count==1
    c=WriterComposer.__new__(WriterComposer);c._app=app;c._doc=doc;c._first_section_configured=False
    sys.settrace(trace)
    for role,fmt in [('cover','none'),('toc','roman'),('body','arabic')]:
        assert app.ActiveDocument._oleobj_.QueryInterface(pythoncom.IID_IUnknown)==doc._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        c.configure_section(role=role,page_number_format=fmt,restart_page_numbering=True,start_page_number=1,header_text='',footer_text='',link_to_previous_header=False,link_to_previous_footer=False)
        c.add_paragraph('Synthetic WPS '+role+' page')
        section=doc.Sections(doc.Sections.Count);numbers=section.Footers(1).PageNumbers
        report['events'].append({'role':role,'restart':bool(numbers.RestartNumberingAtSection),'start':numbers.StartingNumber,'style':numbers.NumberStyle,'footer_fields':section.Footers(1).Range.Fields.Count});persist()
    c.update_fields()
    report['status']='PASS'
except Exception as e:
    report['status']='FAIL';report['errors'].append({'error':repr(e),'traceback':traceback.format_exc()})
finally:
    sys.settrace(None)
    if doc is not None:
        try:
            doc.SaveAs(str(root/'probe.docx'),12)
            doc.Close(SaveChanges=0)
            report['owned_document_closed']=True;report['documents_after']=app.Documents.Count
        except Exception as e:report['errors'].append({'cleanup_error':repr(e)})
    persist()
print(json.dumps(report,ensure_ascii=False))
