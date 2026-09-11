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

output=ROOT/'docs/verification/microsoft-parity/macos-word-degradation/window-diagnosis-01'
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
 case=output/('active' if activate else 'inactive');case.mkdir()
 state=report['modes'][case.name]={};sentinel=None;session=None
 try:
  starting=inventory(case,'starting')
  with create_document('writer',engine='msoffice',visible=False) as session:
   try:
    token='Window diagnosis sentinel 中文😀 '+uuid4().hex
    sentinel=session._execute(['set d to make new document',f'set content of text object of d to {apple_string(token)}','set nativeRows to {{name of d as text}}'])[0][0]
    before=next(r for r in inventory(case,'sentinel-before') if r[0]==sentinel)
    state['sentinel_before']=before
    commands=['set beforeText to content of text object of boundDoc as text','set beforeEnd to end of content of text object of boundDoc','set beforeCount to count tables of boundDoc','set probeResult to {"not-run"}','try']
    if activate:commands+=['activate object boundWindow']
    commands+=['set p to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start p end p','set t to make new table at boundDoc with properties {text object:r,number of rows:1,number of columns:1}','set probeResult to {"created",start of content of text object of t,end of content of text object of t,count rows of t,count columns of t}','on error probeError number probeNumber','set probeResult to {"error",probeNumber,probeError}','end try','set nativeRows to {{"before",beforeText,beforeEnd,beforeCount},probeResult,{"after",content of text object of boundDoc as text,end of content of text object of boundDoc,count tables of boundDoc}}']
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
