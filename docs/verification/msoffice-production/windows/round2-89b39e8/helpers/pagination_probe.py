import sys, json, traceback, shutil
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer, _word_processes
root=Path('build/msoffice-production/pagination-probe-01').resolve();root.mkdir(exist_ok=False)
report={'events':[], 'errors':[]}
def persist(): (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
def event(label, section):
    row={'label':label,'section':int(section.Index)}
    for name,getter in {
        'link':lambda:bool(section.Footers(1).LinkToPrevious),
        'count':lambda:int(section.Footers(1).PageNumbers.Count),
        'restart':lambda:bool(section.Footers(1).PageNumbers.RestartNumberingAtSection),
        'starting_number':lambda:int(section.Footers(1).PageNumbers.StartingNumber),
        'number_style':lambda:int(section.Footers(1).PageNumbers.NumberStyle),
        'field_codes':lambda:[section.Footers(1).Range.Fields(i).Code.Text for i in range(1,section.Footers(1).Range.Fields.Count+1)],
    }.items():
        try:row[name]=getter()
        except Exception as e:row[name]={'error':repr(e)}
    report['events'].append(row);persist()
def change(label,section,fn):
    try:fn()
    except Exception as e:report['errors'].append({'label':label,'error':repr(e),'traceback':traceback.format_exc()})
    event(label,section)
c=None
try:
    c=create_dedicated_composer(str(root));report['pid']=c.identity.pid;report['binding']=c.application_binding
    # First record the inherited production methods and their swallowed errors.
    def trace(frame,kind,arg):
        if kind=='exception' and frame.f_code.co_name in ('set_page_numbering','set_header_footer'):
            report['errors'].append({'function':frame.f_code.co_name,'line':frame.f_lineno,'error':repr(arg[1])});persist()
        return trace
    sys.settrace(trace)
    for role,fmt in [('cover','none'),('toc','roman'),('body','arabic')]:
        c.configure_section(role=role, page_number_format=fmt, restart_page_numbering=True, start_page_number=1,link_to_previous_header=False,link_to_previous_footer=False)
        c.add_paragraph('Synthetic '+role+' page')
        event('production_configure_'+role,c.doc.Sections(c.doc.Sections.Count))
    sys.settrace(None)
    c.update_fields();c.save_docx(root/'before.docx');c.export_pdf(root/'before.pdf')
    for i in range(1,4):event('before_manual',c.doc.Sections(i))
    for i,style in [(1,0),(2,2),(3,0)]:
        section=c.doc.Sections(i);footer=section.Footers(1)
        change('unlink',section,lambda:setattr(footer,'LinkToPrevious',False))
        if i==1:
            change('clear_cover',section,lambda:setattr(footer.Range,'Text',''))
            continue
        numbers=footer.PageNumbers
        change('restart_true',section,lambda:setattr(numbers,'RestartNumberingAtSection',True))
        change('start_one',section,lambda:setattr(numbers,'StartingNumber',1))
        change('style',section,lambda:setattr(numbers,'NumberStyle',style))
    c.update_fields();c.save_docx(root/'after.docx');c.export_pdf(root/'after.pdf')
    for i in range(1,4):event('after_saved',c.doc.Sections(i))
finally:
    sys.settrace(None)
    if c is not None:
        c.close(save_changes=False);report['closed']=c._closed;report['pid_absent']=c.identity.pid not in _word_processes()
    persist()
print(json.dumps(report,ensure_ascii=False))
