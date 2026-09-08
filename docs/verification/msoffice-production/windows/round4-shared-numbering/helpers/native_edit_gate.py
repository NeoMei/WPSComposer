"""Bounded native numbering/pagination edits on an owned representative copy."""
import sys,json,hashlib,shutil,time,re,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import fitz
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer
from fixtures.verify_msoffice_production import HEADINGS
root=Path('build/msoffice-production/native-edit-01').resolve();root.mkdir(exist_ok=False)
source=Path('build/msoffice-production/representative-03/generated.docx').resolve()
before=root/'before.docx';shutil.copyfile(source,before)
shutil.copyfile(source.with_name('converted.pdf'),root/'before.pdf')
report={'status':'RUNNING','checks':{},'snapshots':[],'ui_snapshots':[],'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest()}
def persist(): (root/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
c=None;ui_ready=False
def heading(text):
    matches=[c.doc.Paragraphs(i) for i in range(1,c.doc.Paragraphs.Count+1) if c.doc.Paragraphs(i).Range.Text.strip()==text and int(c.doc.Paragraphs(i).OutlineLevel)<=6]
    assert len(matches)==1,(text,len(matches))
    return matches[0]
def refresh():
    for _ in range(2):
        c.doc.Repaginate()
        c.update_fields()  # Production method includes the dedicated TOC.Update call.
    c.doc.Repaginate()
def snapshot(label):
    rows=[]
    for text in [*HEADINGS,'临时新增编号章节']:
        try:p=heading(text)
        except AssertionError:continue
        r=p.Range.Duplicate;r.Collapse(1)
        rows.append({'text':text,'number':str(p.Range.ListFormat.ListString),'logical_page':int(r.Information(1)),'physical_page':int(r.Information(3))})
    toc=str(c.doc.TablesOfContents.Item(1).Range.Text)
    row={'label':label,'headings':rows,'toc_text':toc,'sections':c.doc.Sections.Count,'pages':c.doc.ComputeStatistics(2)}
    report['snapshots'].append(row);persist();return row
def toc_page(row,text):
    matches=[p for p in row['toc_text'].split('\r') if text in p]
    assert len(matches)==1,matches
    return int(matches[0].strip().split('\t')[-1])
try:
    c=create_dedicated_composer(str(root));report['pid']=c.identity.pid;report['binding']=c.application_binding
    # Exercise the final shared section code on real Word after the WPS adjustment.
    for role,fmt in [('cover','none'),('toc','roman'),('body','arabic')]:
        c.configure_section(role=role,page_number_format=fmt,restart_page_numbering=True,start_page_number=1,header_text='',footer_text='',link_to_previous_header=False,link_to_previous_footer=False)
        c.add_paragraph('Final Word section '+role)
    c.save_docx(root/'current-section-probe.docx');c.export_pdf(root/'current-section-probe.pdf')
    with fitz.open(root/'current-section-probe.pdf') as pdf:
        footers=[re.sub(r'\s+','',p.get_text(clip=fitz.Rect(0,p.rect.height-72,p.rect.width,p.rect.height))) for p in pdf]
    report['checks']['final_source_word_section_footers']=footers==['','i','1']
    report['section_probe_footers']=footers
    c.open_owned_document(before,read_only=False)
    initial=snapshot('before')
    marker='临时新增编号章节'
    start=int(heading(HEADINGS[0]).Range.Start)
    c.doc.Range(start,start).InsertBefore(marker+'\r')
    inserted=c.doc.Range(start,start+len(marker)+1)
    inserted.Style=c.doc.Styles(-2)
    refresh();added=snapshot('inserted_heading')
    report['checks']['insertion_automatically_renumbers_h1_and_h2']=(str(heading(marker).Range.ListFormat.ListString)=='1' and str(heading(HEADINGS[0]).Range.ListFormat.ListString)=='2' and str(heading(HEADINGS[1]).Range.ListFormat.ListString)=='2.1')
    report['checks']['toc_includes_inserted_heading_and_updated_number']=marker in added['toc_text'] and '2'+HEADINGS[0] in re.sub(r'\s+','',added['toc_text'])
    c.save_docx(root/'after-insert.docx');c.export_pdf(root/'after-insert.pdf')
    heading(marker).Range.Delete()
    refresh();deleted=snapshot('deleted_heading')
    report['checks']['deletion_restores_original_numbering']=str(heading(HEADINGS[0]).Range.ListFormat.ListString)=='1' and str(heading(HEADINGS[1]).Range.ListFormat.ListString)=='1.1'
    report['checks']['toc_removes_deleted_heading']=marker not in deleted['toc_text']
    c.save_docx(root/'after-delete.docx')
    start=int(heading(HEADINGS[0]).Range.Start)
    spacer='分页移动验收占位段\r'
    c.doc.Range(start,start).InsertBefore(spacer)
    spacer_range=c.doc.Range(start,start+len(spacer))
    spacer_range.Style=c.doc.Styles('Body Text');spacer_range.ListFormat.RemoveNumbers(1)
    c.doc.Range(start+len(spacer),start+len(spacer)).InsertBreak(7)
    refresh();moved=snapshot('chapter_moved_to_next_page')
    original=next(h for h in initial['headings'] if h['text']==HEADINGS[0])
    current=next(h for h in moved['headings'] if h['text']==HEADINGS[0])
    report['checks']['chapter_moves_one_physical_and_logical_page']=current['logical_page']==original['logical_page']+1 and current['physical_page']==original['physical_page']+1
    report['checks']['updated_toc_matches_native_adjusted_page']=all(toc_page(moved,h['text'])==h['logical_page'] for h in moved['headings'][:3])
    final=root/'edited-pagination.docx';c.save_docx(final)
    c.open_owned_document(final,read_only=False)
    reopened=snapshot('saved_and_reopened')
    report['checks']['save_reopen_retains_numbering_and_toc']=reopened['headings']==moved['headings'] and reopened['toc_text']==moved['toc_text']
    c.export_pdf(root/'edited-pagination.pdf')
    with fitz.open(root/'edited-pagination.pdf') as pdf:
        pages=[re.sub(r'\s+','',p.get_text()) for p in pdf]
        footers=[re.sub(r'\s+','',p.get_text(clip=fitz.Rect(0,p.rect.height-72,p.rect.width,p.rect.height))) for p in pdf]
    report['pdf_footer_texts']=footers
    report['checks']['pdf_chapter_on_expected_physical_page']=HEADINGS[0] in pages[current['physical_page']-1]
    report['checks']['pdf_toc_matches_logical_chapter_page']=all(re.search(re.escape(h)+r'\.+2',pages[1]) for h in HEADINGS[:3])
    report['checks']['pdf_footer_sequence_preserved']=footers==['','i']+[str(i) for i in range(1,len(pages)-1)]
    report['checks']['original_source_bytes_unchanged']=hashlib.sha256(source.read_bytes()).hexdigest()==report['source_sha256']
    assert all(report['checks'].values()),report['checks']
    report['status']='NATIVE_PASS_UI_REOPEN_PENDING';persist()
    c.app.Visible=True;c.doc.Activate();ui_ready=True
    print('UI_READY',report['pid'],str(final),flush=True)
    sequence=0;deadline=time.monotonic()+1800
    while time.monotonic()<deadline:
        command=root/'command.json'
        if command.exists():
            data=json.loads(command.read_text())
            if data['sequence']>sequence:
                sequence=data['sequence'];count=c.app.Documents.Count
                row={'label':data.get('label',data['action']),'documents_count':count,'time_unix':time.time()}
                if count==1:
                    doc=c.app.Documents.Item(1);assert Path(doc.FullName).resolve()==final
                    c._verify_document(doc);c._doc=doc
                    row.update(saved=bool(doc.Saved),text_sha256=hashlib.sha256(doc.Content.Text.encode()).hexdigest(),toc_text=doc.TablesOfContents.Item(1).Range.Text)
                row['disk_sha256']=hashlib.sha256(final.read_bytes()).hexdigest()
                report['ui_snapshots'].append(row);persist()
                if data['action']=='finish':
                    assert count==0,'Refusing to Quit with an open document'
                    c._doc=None;c.close(save_changes=False);report['owned_application_closed']=c._closed
                    report['status']='PASS';persist();break
        time.sleep(.2)
except Exception as e:
    report['status']='FAIL';report['error']=repr(e);report['traceback']=traceback.format_exc();persist()
finally:
    if c is not None and not ui_ready:
        try:c.close(save_changes=False);report['owned_application_closed']=c._closed
        except Exception as e:report['cleanup_error']=repr(e)
        persist()
print(json.dumps(report,ensure_ascii=False),flush=True)
