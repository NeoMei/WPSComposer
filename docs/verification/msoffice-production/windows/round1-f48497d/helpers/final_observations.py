import sys,json,hashlib,time,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
import fitz,pythoncom,win32com.client,win32api
from PIL import Image
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes
root=Path('build/msoffice-production')
events=[json.loads(line) for p in (root/'traces').rglob('trace-*.jsonl') for line in p.read_text('utf-8').splitlines()]
created=[e for e in events if e['event']=='composer_created']
closed=[e for e in events if e['event']=='composer_close_return']
ui=json.loads((root/'ui-01/com-observations.json').read_text('utf-8'))
pids=sorted({e['word_pid'] for e in created}|{ui['pid']})
processes=_word_processes()
pythoncom.CoInitialize();registered=win32com.client.GetActiveObject('Word.Application')
post={'task_pids':pids,'task_pids_absent':{str(p):p not in processes for p in pids},'all_task_pids_absent':all(p not in processes for p in pids),'registered_9252_retained':9252 in processes,'registered_name':registered.Name,'registered_documents_count':registered.Documents.Count,'all_native_composers_closed':all(any(c['word_pid']==e['word_pid'] and c['closed'] and c['cleanup_error']=='None' for c in closed) for e in created),'observed_unix':time.time(),'observation_hook_errors':[str(p) for p in (root/'traces').rglob('observer-errors.log')]}
info=win32api.GetFileVersionInfo(created[0]['executable'],'\\')
post['word_file_version']='.'.join(map(str,[info['FileVersionMS']>>16,info['FileVersionMS']&65535,info['FileVersionLS']>>16,info['FileVersionLS']&65535]))
(root/'process-postflight.json').write_text(json.dumps(post,indent=2),encoding='utf-8')
assert post['all_task_pids_absent'] and post['all_native_composers_closed'] and not post['observation_hook_errors']
assert post['registered_9252_retained'] and post['registered_documents_count']==0
def raster(path):
    with fitz.open(path) as pdf:
        return [hashlib.sha256(p.get_pixmap(matrix=fitz.Matrix(1.25,1.25),alpha=False).samples).hexdigest() for p in pdf]
comparison={'converted_pdf_raster_hashes':raster(root/'representative-01/converted.pdf'),'public_generate_pdf_raster_hashes':raster(root/'public-pdf-01/generated.pdf')}
comparison['all_pages_identical']=comparison['converted_pdf_raster_hashes']==comparison['public_generate_pdf_raster_hashes']
(root/'pdf-route-comparison.json').write_text(json.dumps(comparison,indent=2),encoding='utf-8')
assert comparison['all_pages_identical']
for raw in (root/'ui-01').glob('*-raw.png'):
    with Image.open(raw) as image:
        image.crop((260,185,min(1260,image.width),min(738,image.height))).save(raw.with_name(raw.name.replace('-raw.png','-document-only.png')))
snap={s['label']:s for s in ui['snapshots']}
labels=['opened','undone','saved','reopened']
checks={'text_restored_after_undo_and_reopen':len({snap[l]['text_sha256'] for l in labels})==1,'inserted_marker_was_unsaved':snap['inserted']['marker_present'] and not snap['inserted']['saved'],'marker_absent_after_undo_and_reopen':all(not snap[l]['marker_present'] for l in labels),'saved_file_bytes_unchanged':all(snap[l]['disk_sha256']==ui['source_sha256'] for l in labels),'table_and_toc_retained':all(snap[l]['table_count']==1 and snap[l]['toc_count']==1 and snap[l]['fields_count']==4 for l in labels),'closed_then_reopened':snap['closed']['documents_count']==0 and snap['reopened']['documents_count']==1,'final_owned_instance_closed':ui['owned_application_closed']}
result={'checks':checks,'all_passed':all(checks.values()),'source_sha256':ui['source_sha256'],'ui_pid':ui['pid'],'limitations':['CUA focused_element repeatedly reported the navigation search box while the visible caret and actual insertion were in the body.','Several captures returned null accessibility, one element index was unavailable, and a reopened capture returned unrelated browser document_text. Re-observation, window screenshots, retained COM identity and task-document hashes were used; these UIA fields are not treated as authoritative.','Screenshots published as document-only crops exclude account controls, unrelated navigation and file-dialog directory listings. Raw captures remain in ignored build.']}
(root/'ui-01/validation.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
assert result['all_passed']
raw_events=json.loads((root/'ui-01/cua-events.json').read_text('utf-8'))
for event in raw_events:
    value=event.get('document_text') or ''
    if 'NATIVE-PRODUCTION-END' not in value and '原生文档生产验收报告' not in value:
        event.pop('document_text',None)
        event['unreliable_document_text_omitted']=True
(root/'ui-01/cua-events-filtered.json').write_text(json.dumps(raw_events,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'postflight':post,'pdf_routes_identical':comparison['all_pages_identical'],'ui':result},ensure_ascii=False))
