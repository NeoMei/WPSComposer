from __future__ import annotations
from dataclasses import asdict
import hashlib, json, shutil, sys, traceback
from pathlib import Path
from uuid import uuid4

ROOT=Path('/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description')
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_sections import _inventory_commands

OUT=ROOT/'.superpowers/sdd/2026-09-08-microsoft-wps-parity/word-field-topology-native-run-04'
OUT.mkdir(parents=True,exist_ok=False)
sources=[ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_fields.py']
report={'passed':False,'source_hashes':{str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in sources},'checks':{}}
session=None;sentinel=None
try:
  starting=[]
  with create_document('writer',engine='msoffice',visible=False) as session:
    starting=session._execute(_inventory_commands())
    token='Topology sentinel 中文😀 '+uuid4().hex
    sentinel=session._execute(['set d to make new document',f'set content of text object of d to {apple_string(token)}','set nativeRows to {{name of d as text}}'])[0][0]
    before=session._execute(_inventory_commands())
    session.add_heading_level('Topology acceptance heading',1)
    session.add_paragraph('Target paragraph for native reference.')
    bookmark='wpsc_fig_'+'a'*24
    session._execute(['set r to create range boundDoc start 0 end 8',f'make new bookmark at boundDoc with properties {{name:{apple_string(bookmark)},text object:r}}',*session._position('end'),'set f to create range boundDoc start insertionPoint end insertionPoint',f'create new field text range f field type field ref field text {apple_string(bookmark)} preserve formatting true','set nativeRows to {{"ok"}}'])
    initial=session.snapshot_fields(); initial_keys=[x.stable_key for x in initial]
    assert len(initial)==1 and initial_keys[0][1]=='REF'
    session.add_cross_reference_paragraph(runs=[{'type':'reference','bookmarkName':bookmark,'prefix':'Owned ','suffix':' end','fallbackText':'fallback'}],owner_node_id='owner:native')
    after_ref=session.snapshot_fields(); keys_ref=[x.stable_key for x in after_ref]
    assert keys_ref[0]==initial_keys[0] and ('owner:native','REF',0) in keys_ref and len(after_ref)==2
    session.insert_toc('Native Contents')
    after_toc=session.snapshot_fields(); keys_toc=[x.stable_key for x in after_toc]
    assert keys_toc[:2]==keys_ref and ('doc:toc','TOC',0) in keys_toc and len(after_toc)==3
    session.apply_structural_op({'op':'insert','type':'paragraph','position':'start','props':{'text':'Prefix'}})
    after_shift=session.snapshot_fields(); keys_shift=[x.stable_key for x in after_shift]
    assert keys_shift==keys_toc
    report['snapshots']={k:[asdict(x) for x in v] for k,v in [('initial',initial),('after_ref',after_ref),('after_toc',after_toc),('after_shift',after_shift)]}
    report['checks'].update(old_keys_preserved=True,new_ref_count=True,new_toc_count=True,structural_shift_rebased=True)
    sys.modules['__main__'].__file__='<native-topology-probe>'
    session.save_docx(OUT/'topology.docx'); session.export_pdf(OUT/'topology.pdf')
    report['inventory_before']=before;report['inventory_after']=session._execute(_inventory_commands());assert report['inventory_after']==before
    report['checks']['unrelated_preserved']=True
    session._execute([f'set d to document {apple_string(sentinel)}',f'if (content of text object of d as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of d then error "SENTINEL_SAVED"','close d saving no','set nativeRows to {{"ok"}}']);sentinel=None
    shutil.copytree(session.staging_root,OUT/'native-runtime')
  report['after_owned_close']=session._closed
  with open_document(OUT/'topology.docx',engine='msoffice',read_only=True,visible=False) as session:
    reopen=session._execute(['set nativeRows to {{count fields of boundDoc,count tables of contents of boundDoc,content of text object of boundDoc as text}}'])[0]
    assert reopen[0]>=3 and reopen[1]>=1 and 'Prefix' in reopen[2] and 'Owned' in reopen[2]
    report['reopen']=reopen;shutil.copytree(session.staging_root,OUT/'reopen-runtime')
  report['checks']['save_reopen_source']=True
  import pdfplumber
  with pdfplumber.open(OUT/'topology.pdf') as pdf:
    text='\n'.join(p.extract_text() or '' for p in pdf.pages)
    assert 'Prefix' in text and 'Owned' in text
    report['pdf_pages']=len(pdf.pages)
  report['checks']['pdf_read']=True
  report['artifact_hashes']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (OUT/'topology.docx',OUT/'topology.pdf')}
  report['passed']=all(report['checks'].values())
except BaseException as exc:
  report['error']={'type':type(exc).__name__,'message':str(exc)}
  (OUT/'failure.txt').write_text(traceback.format_exc())
  if session and getattr(session,'staging_root',None) and session.staging_root.exists():shutil.copytree(session.staging_root,OUT/'failed-runtime',dirs_exist_ok=True)
  if sentinel:report['retained_sentinel_name']=sentinel
finally:
  report['final_inventory']=session._execute(_inventory_commands()) if session and not session._closed and not session._quarantined else 'session-closed-or-quarantined'
  (OUT/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print(json.dumps({'passed':report['passed'],'error':report.get('error')}));sys.exit(0 if report['passed'] else 1)
