"""Task-specific exact-owned-window table diagnostic; explicit root lease only."""
from pathlib import Path
import hashlib
import json
import shutil
import sys
import traceback
from uuid import uuid4

ROOT=Path.cwd()
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_recovery import inventory,close_sentinel

output=ROOT/'docs/verification/microsoft-parity/macos-word-degradation/geometry-diagnosis-01'
output.mkdir(parents=True,exist_ok=False)
shutil.copy2(__file__,output/'diagnosis.py')
report={'passed':False,'modes':{},'source_hashes':{}}
for name in ('macos_word_session.py','macos_word_degradation.py','macos_word_recovery.py'):
 source=ROOT/'skills/WPSComposer/scripts/msoffice'/name
 shutil.copy2(source,output/name)
 report['source_hashes'][name]=hashlib.sha256(source.read_bytes()).hexdigest()

def flush():
 (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))

flush()
for activate in (False,True):
 case=output/('nonempty' if activate else 'empty');case.mkdir()
 state=report['modes'][case.name]={};sentinel=None;session=None
 try:
  starting=inventory(case,'starting')
  with create_document('writer',engine='msoffice',visible=False) as session:
   try:
    token='Window diagnosis sentinel 中文😀 '+uuid4().hex
    sentinel=session._execute(['set d to make new document',f'set content of text object of d to {apple_string(token)}','set nativeRows to {{name of d as text}}'])[0][0]
    before=next(r for r in inventory(case,'sentinel-before') if r[0]==sentinel)
    state['sentinel_before']=before
    if activate:
     session._execute(['set content of text object of boundDoc to "PREFIX 中文😀"','set nativeRows to {{"prefix"}}'])
    from skills.WPSComposer.scripts.msoffice.macos_word_degradation import table_commands
    prefix='PREFIX 中文😀' if activate else ''
    display='[GEOMETRY_NOTICE] 中文😀'
    position=len(prefix.encode('utf-16-le'))//2
    commands=table_commands(session,display,position)
    catch=commands.index('on error noticeError number noticeNumber')
    commands[catch:]=['on error noticeError number noticeNumber','set nativeRows to {{"private-error",noticeNumber,noticeError}}','end try']
    commands += ['set end of nativeRows to {"full-body",content of text object of boundDoc as text,end of content of text object of boundDoc,count tables of boundDoc}',
     'if (count tables of boundDoc) is 1 then',
     'set t to table 1 of boundDoc','set c to get cell from table t row 1 column 1',
     'set end of nativeRows to {"table",start of content of text object of t,end of content of text object of t,content of text object of t as text}',
     'set end of nativeRows to {"cell",start of content of text object of c,end of content of text object of c,content of text object of c as text}',
     'end if']
    state['native']=session._execute(commands)
    assert next(r for r in inventory(case,'sentinel-after') if r[0]==sentinel)==before
    flush()
   finally:
    if session.staging_root.exists():shutil.copytree(session.staging_root,case/'native-runtime')
  after=inventory(case,'after-owned-close')
  assert next(r for r in after if r[0]==sentinel)==before
  assert [r for r in after if r[0]!=sentinel]==starting
  state['owned_closed']=True
  state['sentinel_close']=close_sentinel(case,sentinel,token);sentinel=None
  state['after']=inventory(case,'after-sentinel-close');assert state['after']==starting
  state['cleanup_verified']=True
 except BaseException as e:
  state['error']={'type':type(e).__name__,'message':str(e),'quarantined':getattr(session,'_quarantined',False)}
  state['retained_sentinel']=sentinel
  (case/'failure.txt').write_text(traceback.format_exc())
  if session and session.staging_root.exists():shutil.copytree(session.staging_root,case/'failure-runtime',dirs_exist_ok=True)
  flush();raise
 flush()
report['passed']=True;flush()
print(json.dumps(report,ensure_ascii=False))
