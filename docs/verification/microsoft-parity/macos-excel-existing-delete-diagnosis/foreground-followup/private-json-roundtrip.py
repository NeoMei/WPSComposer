from pathlib import Path
import sys,os,json,subprocess,time,hashlib,shutil,traceback
from unittest.mock import patch
root=Path.cwd();sys.path.insert(0,str(root))
from skills.WPSComposer.scripts.msoffice.macos_excel_session import MacExcelSession,_JSON,_object,_quote
out=root/'build/excel-command-boundary-20260910/private-json-roundtrip';out.mkdir(exist_ok=False)
helper=root/'build/excel-command-boundary-20260910/osakit-instance-ctypes.py'; real_run=subprocess.run
source=root/'docs/verification/microsoft-parity/macos-excel-sessions/business-04/business.xlsx'
report={'passed':False,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'helper_sha256':hashlib.sha256(helper.read_bytes()).hexdigest(),'steps':[]};pid=None;identity=None;session=None

def flush(): (out/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False))
def pids():return set(real_run(['pgrep','-x','Microsoft Excel'],capture_output=True,text=True).stdout.split())
def ident(target):return real_run(['ps','-p',str(target),'-o','lstart=','-o','comm='],capture_output=True,text=True,check=True).stdout.strip()
def call(target,label,body):
 script=out/(label+'.applescript');script.write_text(_JSON+'\ntell application "/Applications/Microsoft Excel.app"\n'+body+'\nend tell\n')
 r=real_run([sys.executable,str(helper),str(target),str(script)],capture_output=True,text=True,timeout=30)
 script.with_suffix('.stdout.log').write_text(r.stdout);script.with_suffix('.stderr.log').write_text(r.stderr)
 if r.returncode:raise RuntimeError(label+' failed')
 return json.loads(r.stdout)
def inventory(target,label):
 return call(target,label,'set rows to {}\nrepeat with wb in workbooks\nset end of rows to {name of wb,full name of wb,saved of wb,value of range "A1" of worksheet 1 of wb}\nend repeat\nreturn '+_object({'books':'rows','alerts':'display alerts'}))
def wrapped(args,**kw):
 if args[0]=='/usr/bin/osascript':
  if ident(pid)!=identity:raise RuntimeError('owned PID identity changed')
  return real_run([sys.executable,str(helper),str(pid),args[1]],**kw)
 return real_run(args,**kw)
try:
 before=pids();assert before=={'9212'},before
 report['original_before']=inventory(9212,'original-before');flush()
 real_run(['open','-n','-a','/Applications/Microsoft Excel.app'],check=True)
 end=time.monotonic()+25
 while time.monotonic()<end:
  current=pids();new=current-before
  if len(new)==1:pid=int(next(iter(new)));break
  time.sleep(.2)
 if pid is None or pids()!=before|{str(pid)}:raise RuntimeError('unambiguous new PID missing')
 identity=ident(pid);report.update(pid=pid,process_identity=identity);flush()
 time.sleep(3)
 report['private_before']=inventory(pid,'private-before');flush()
 assert report['private_before']['alerts'] is True
 books=report['private_before']['books'];assert not books or (len(books)==1 and books[0][2] is True and books[0][3] in ('',None)),books
 with patch('subprocess.run',wrapped):
  session=MacExcelSession.open_document(source)
  report['owned_path']=str(session._native);report['runtime']=str(session._job);flush()
  report['steps'].append(session._run('return '+_object({'names':'name of every worksheet of ownedBook'})))
  assert report['steps'][-1]['names']==['Data','Results','Business report']
  session._run('set value of range "Z40" of worksheet 1 of ownedBook to "PRIVATE-JSON-ROUNDTRIP"\nreturn "{}"')
  deletion='''if (count of workbooks) is not 1 then error "foreign workbook in private instance"
if display alerts is not true then error "unexpected alerts"
set display alerts to false
if display alerts is not false then error "alerts not suppressed"
try
 delete worksheet "Business report" of ownedBook
 set display alerts to true
on error msg number n
 set display alerts to true
 error msg number n
end try
if name of every worksheet of ownedBook is not {"Data","Results"} then error "delete not acknowledged"
if display alerts is not true then error "alerts not restored"
return '''+_object({'deleted':'true'})
  report['steps'].append(session._run(deletion,acknowledge='deleted'));flush()
  session.save(out/'result.xlsx');session.close();shutil.copytree(session._job,out/'first-runtime');session=None
  reopened=MacExcelSession.open_document(out/'result.xlsx',read_only=True);session=reopened
  got=reopened._run('return '+_object({'names':'name of every worksheet of ownedBook','marker':'value of range "Z40" of worksheet 1 of ownedBook','alerts':'display alerts'}));report['reopened']=got;flush()
  assert got=={'names':['Data','Results'],'marker':'PRIVATE-JSON-ROUNDTRIP','alerts':True},got
  reopened.close();shutil.copytree(reopened._job,out/'reopen-runtime');session=None
 report['private_final']=inventory(pid,'private-final');assert report['private_final']=={'books':[],'alerts':True}
 assert ident(pid)==identity
 report['quit']=call(pid,'quit-owned','if (count of workbooks) is not 0 then error "not empty"\nif display alerts is not true then error "alerts changed"\nquit saving no\nreturn "{\\\"quit\\\":true}"')
 end=time.monotonic()+15
 while time.monotonic()<end and str(pid) in pids():time.sleep(.2)
 assert pids()==before,pids()
 report['original_after']=inventory(9212,'original-after');assert report['original_after']==report['original_before']
 assert hashlib.sha256(source.read_bytes()).hexdigest()==report['source_sha256']
 report['passed']=True
except BaseException as exc:
 report['error']=str(exc);(out/'failure.txt').write_text(traceback.format_exc())
 if session is not None:report['remaining_owned_path']=str(session._native)
finally:flush()
print(json.dumps(report,indent=2,ensure_ascii=False))
raise SystemExit(0 if report['passed'] else 1)
