import sys,json,time,hashlib,shutil,traceback
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer
root=Path('build/msoffice-production/ui-01').resolve();root.mkdir(exist_ok=False)
source=Path('build/msoffice-production/representative-01/generated.docx').resolve()
copy=root/'production-ui-edit.docx';shutil.copyfile(source,copy)
report={'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'copy_path':str(copy),'snapshots':[]}
def persist(): (root/'com-observations.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
composer=create_dedicated_composer(str(root));report['binding']=composer.application_binding
report['pid']=composer.identity.pid
composer.open_owned_document(copy,read_only=False)
composer.app.Visible=True
composer.doc.Activate()
def snapshot(label):
    count=composer.app.Documents.Count
    row={'label':label,'documents_count':count,'time_unix':time.time()}
    if count==1:
        doc=composer.app.Documents.Item(1)
        assert Path(doc.FullName).resolve()==copy,'Unexpected document path'
        composer._verify_document(doc)
        composer._doc=doc
        content=doc.Content.Text
        row.update(name=doc.Name,path=doc.FullName,saved=bool(doc.Saved),text_sha256=hashlib.sha256(content.encode()).hexdigest(),characters=len(content),marker_present='WPS-PRODUCTION-UI-MARKER' in content,table_count=doc.Tables.Count,toc_count=doc.TablesOfContents.Count,fields_count=doc.Fields.Count,hwnd=int(doc.Windows.Item(1).Hwnd))
    row['disk_sha256']=hashlib.sha256(copy.read_bytes()).hexdigest()
    report['snapshots'].append(row);persist()
snapshot('opened')
print('UI_READY',report['pid'],flush=True)
seen=0
deadline=time.monotonic()+1800
while time.monotonic()<deadline:
    command=root/'command.json'
    if command.exists():
        data=json.loads(command.read_text('utf-8'))
        if data['sequence']>seen:
            seen=data['sequence']
            try:
                if data['action']=='finish':
                    snapshot('final_closed')
                    assert composer.app.Documents.Count==0,'Refusing Quit with an open document'
                    composer._doc=None
                    composer.close(save_changes=False)
                    report['owned_application_closed']=composer._closed;persist();break
                snapshot(data['label'])
            except Exception as exc:
                report.setdefault('errors',[]).append({'error':repr(exc),'traceback':traceback.format_exc()});persist()
    time.sleep(.2)
