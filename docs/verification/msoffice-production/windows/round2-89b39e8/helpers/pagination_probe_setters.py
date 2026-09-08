import sys, json, traceback, shutil
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer, _word_processes
root=Path('build/msoffice-production/pagination-probe-02').resolve();root.mkdir(exist_ok=False)
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
    section=c.doc.Sections(1);footer=section.Footers(1);numbers=footer.PageNumbers
    event('empty_initial',section)
    change('empty_restart_minus_one',section,lambda:setattr(numbers,'RestartNumberingAtSection',-1))
    change('empty_start_one',section,lambda:setattr(numbers,'StartingNumber',1))
    change('empty_style_roman',section,lambda:setattr(numbers,'NumberStyle',2))
    change('add_page_field',section,lambda:footer.Range.Fields.Add(footer.Range,33))
    change('field_restart_minus_one',section,lambda:setattr(numbers,'RestartNumberingAtSection',-1))
    change('field_start_one',section,lambda:setattr(numbers,'StartingNumber',1))
    change('field_restart_false',section,lambda:setattr(numbers,'RestartNumberingAtSection',False))
    change('field_restart_true',section,lambda:setattr(numbers,'RestartNumberingAtSection',True))
    change('field_start_five',section,lambda:setattr(numbers,'StartingNumber',5))
    c.add_paragraph('Synthetic setter sequence');c.save_docx(root/'setters.docx');c.export_pdf(root/'setters.pdf')
finally:
    sys.settrace(None)
    if c is not None:
        c.close(save_changes=False);report['closed']=c._closed;report['pid_absent']=c.identity.pid not in _word_processes()
    persist()
print(json.dumps(report,ensure_ascii=False))
