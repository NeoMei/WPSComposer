# coding: utf-8
from pathlib import Path
from types import SimpleNamespace
import time,shutil,json,hashlib
from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
from skills.WPSComposer.scripts.msoffice.macos_script import wrap_owned,apple_string,refresh_source
root=Path(__file__).resolve().parent
source=Path('build/msoffice-production-mac/public-representative-shared-outline/generated.docx').resolve()
adapter=MacWordAdapter(SimpleNamespace(plan=SimpleNamespace(operations=[])))
deadline=time.monotonic()+180
helper='''on findHeading(d, labelText)
 tell application "Microsoft Word"
  repeat with pi from 1 to (count of paragraphs of d)
   set rr to text object of paragraph pi of d
   if content of rr is labelText & return then return rr
  end repeat
  error "Task heading not found"
 end tell
end findHeading
on reportHeading(d, labelText, phaseText)
 tell application "Microsoft Word"
  set rr to my findHeading(d, labelText)
  set pp to (get range information rr information type active end adjusted page number) as integer
  log "WPSC_HEADING" & tab & phaseText & tab & labelText & tab & (list string of list format of rr) & tab & (list value of list format of rr) & tab & pp
  if labelText is "一级章节验收" then
   repeat with childLabel in {"二级章节验收", "三级章节验收", "四级章节验收"}
    set cr to my findHeading(d, childLabel as text)
    set cp to (get range information cr information type active end adjusted page number) as integer
    log "WPSC_CHILD" & tab & phaseText & tab & (childLabel as text) & tab & (list string of list format of cr) & tab & cp
   end repeat
  end if
 end tell
end reportHeading
'''
def run(phase,body,target):
 raw=adapter._run(helper+wrap_owned(body,target,50,source=target),deadline)
 (root/(phase+'.log')).write_text(raw)
 print(phase,raw,flush=True)
 shutil.copy2(target,root/(phase+'.docx'))
 return raw
try:
 adapter._ensure_started(deadline)
 target=adapter.staging_root/'structural-edit.docx';shutil.copy2(source,target)
 save='save as ownedDoc file name '+apple_string(str(target))+' file format format document default add to recent files false'
 before=run('before','my reportHeading(ownedDoc, "一级章节验收", "before")',target)
 body='''set originalRange to my findHeading(ownedDoc, "一级章节验收")
set p to start of content of originalRange
set r to create range ownedDoc start p end p
set content of r to "新增原生编号章节" & return
set r to my findHeading(ownedDoc, "新增原生编号章节")
set style of r to style heading1
set r to my findHeading(ownedDoc, "一级章节验收")
set page break before of paragraph format of r to true
'''+refresh_source()+'''
my reportHeading(ownedDoc, "新增原生编号章节", "inserted")
my reportHeading(ownedDoc, "一级章节验收", "inserted")
'''+save
 inserted=run('inserted',body,target)
 reopened=run('reopened','my reportHeading(ownedDoc, "新增原生编号章节", "reopened")\nmy reportHeading(ownedDoc, "一级章节验收", "reopened")',target)
 body='''set r to my findHeading(ownedDoc, "新增原生编号章节")
set content of r to ""
set r to my findHeading(ownedDoc, "一级章节验收")
set page break before of paragraph format of r to false
'''+refresh_source()+'''
my reportHeading(ownedDoc, "一级章节验收", "deleted")
'''+save
 deleted=run('deleted',body,target)
 deleted_reopened=run('deleted-reopened','my reportHeading(ownedDoc, "一级章节验收", "deleted-reopened")',target)
 (root/'report.json').write_text(json.dumps({'status':'NATIVE_OPERATIONS_COMPLETE','source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'staging_root':str(adapter.staging_root),'method':'Native Word object-model edit with production ownership/sentinel/lock/deadline wrapper; save, close, reopen each phase','phases':['before','inserted','reopened','deleted','deleted-reopened']},ensure_ascii=False,indent=2))
finally:
 adapter.close()
