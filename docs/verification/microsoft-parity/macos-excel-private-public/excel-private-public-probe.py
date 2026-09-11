from pathlib import Path
import json,sys,time,hashlib,traceback,zipfile,subprocess,shutil
from dataclasses import asdict
ROOT=Path.cwd();sys.path.insert(0,str(ROOT))

def main():
 from skills.WPSComposer import open_document
 from skills.WPSComposer.scripts.msoffice import macos_osa_transport as osa
 from skills.WPSComposer.scripts.msoffice.macos_excel_session import _JSON,_object
 out=Path(sys.argv[1]).resolve();out.mkdir(parents=True,exist_ok=False)
 source=ROOT/'docs/verification/microsoft-parity/macos-excel-sessions/business-04/business.xlsx'
 files=[ROOT/'skills/WPSComposer/scripts/msoffice'/n for n in ['macos_excel_session.py','macos_excel_process.py','macos_osa_transport.py','macos_office_runtime.py']]
 hashes={}
 report={'passed':False,'jobs':[],'owners':[]}
 session=None;r=osa._DarwinRuntime();pool=r.message(r.objc_class('NSAutoreleasePool'),'new');lookup=osa._DarwinProcessLookup(r)
 def flush(): (out/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n')
 original=None
 def original_state(label):
  p=out/(label+'.applescript');p.write_text(_JSON+'''\ntell application "/Applications/Microsoft Excel.app"
set probeRows to {}
repeat with n from 1 to count of workbooks
set b to workbook n
set end of probeRows to {name of b,full name of b,saved of b,value of range "A1" of worksheet 1 of b}
end repeat
return '''+_object({'books':'probeRows','alerts':'display alerts'})+'\nend tell')
  a=osa.BoundOSAKitTransport(original).run(p,time.monotonic()+25)
  p.with_suffix('.stdout.log').write_text(a.stdout);p.with_suffix('.stderr.log').write_text(a.stderr)
  assert a.returncode==0,(a.returncode,a.stderr)
  return json.loads(a.stdout)
 try:
  hashes={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
  report.update(source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),source_hashes=hashes,runner_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest())
  expected=json.loads((ROOT/'build/excel-transport-20260910/readonly-01/identity.json').read_text())
  expected.pop('deadline',None)
  original=osa.ExcelProcessIdentity(**expected)
  osa.require_same_process(original,lookup.snapshot(original.pid))
  report['original_before']=original_state('original-before');flush()
  assert report['original_before']=={'books':[['工作簿5','工作簿5',False,'sentinel-58d9ceb7df8e'],['工作簿7','工作簿7',False,'sentinel-e9a7d3dce1ad']],'alerts':True}
  for phase,path in [('edit',source),('reopen',out/'result.xlsx')]:
   session=open_document(path,engine='msoffice',read_only=phase=='reopen')
   report['jobs'].append(str(session._job));ident=session._process_owner.identity;report['owners'].append(asdict(ident));flush()
   assert ident.pid != original.pid
   osa.require_excel_process(ident)
   assert ident.executable==original.executable
   osa.require_same_process(original,lookup.snapshot(original.pid))
   if phase=='edit':
    snapshot=session.inspect_document(include_cells=False);assert [s['name'] for s in snapshot['sheets']]==['Data','Results','Business report']
    session.select_sheet(1);session.write_cell(40,26,'PRIVATE-PUBLIC-SESSION')
    report['delete']=session.apply_structural_op({'op':'remove','target':'sheet:3'});flush()
    session.save(out/'result.xlsx')
    session.export_pdf(out/'result.pdf')
   else:
    check=session._run('return '+_object({'names':'name of every worksheet of ownedBook','marker':'value of range "Z40" of worksheet 1 of ownedBook','alerts':'display alerts','saved':'saved of ownedBook'}));report['reopen']=check;flush()
    assert check=={'names':['Data','Results'],'marker':'PRIVATE-PUBLIC-SESSION','alerts':True,'saved':True},check
   job=session._job;session.close();assert session._process_owner.state=='closed'
   assert lookup.snapshot(ident.pid) is None
   osa.require_same_process(original,lookup.snapshot(original.pid))
   shutil.copytree(job,out/(phase+'-runtime'));session=None
  report['original_after']=original_state('original-after');assert report['original_before']==report['original_after']
  assert hashlib.sha256(source.read_bytes()).hexdigest()==report['source_sha256']
  assert all(hashlib.sha256((ROOT/p).read_bytes()).hexdigest()==h for p,h in hashes.items())
  with zipfile.ZipFile(out/'result.xlsx') as z:
   assert z.testzip() is None
   assert any(b'PRIVATE-PUBLIC-SESSION' in z.read(n) for n in z.namelist() if n.endswith('.xml'))
  report['passed']=True
 except BaseException as exc:
  report['error']={'type':type(exc).__name__,'message':str(exc)};(out/'failure.txt').write_text(traceback.format_exc())
  if session is not None:
   report['remaining_job']=str(session._job);report['remaining_path']=str(session._native)
   report['owner_recovery']=session._process_owner.recovery
 finally:
  try:
   if original is None:raise RuntimeError('original identity preflight incomplete; no native preservation read')
   report['final_original_state']=original_state('original-final')
   report['original_preserved']=report['final_original_state']==report.get('original_before')
   report['source_preserved']=hashlib.sha256(source.read_bytes()).hexdigest()==report.get('source_sha256')
   report['runner_preserved']=hashlib.sha256(Path(__file__).read_bytes()).hexdigest()==report.get('runner_sha256')
   report['code_preserved']=all(hashlib.sha256((ROOT/p).read_bytes()).hexdigest()==h for p,h in hashes.items())
   if not all(report[k] for k in ['original_preserved','source_preserved','code_preserved','runner_preserved']):report['passed']=False
  except BaseException as check_error:
   report['final_preservation_error']=str(check_error);report['passed']=False
  flush();r.message(pool,'drain')
 print(json.dumps(report,indent=2,ensure_ascii=False));return 0 if report['passed'] else 1

if __name__=='__main__':raise SystemExit(main())
