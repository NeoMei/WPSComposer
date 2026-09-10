from pathlib import Path
import hashlib,json,sys,subprocess,time
root=Path.cwd();sys.path.insert(0,str(root))
from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity,BoundOSAKitTransport
folder=root/'build/installed-checkpoint433';info=json.loads((folder/'install.json').read_text());plugin=Path(info['plugin'])
paths=[p for p in plugin.rglob('*') if p.is_file() and 'node_modules' not in p.parts and '__pycache__' not in p.parts]
source={str(p.relative_to(plugin)):hashlib.sha256(p.read_bytes()).hexdigest() for p in paths if (root/p.relative_to(plugin)).is_file()}
assert source and all(hashlib.sha256((root/p).read_bytes()).hexdigest()==h for p,h in source.items())
assert 'skills/WPSComposer/scripts/msoffice/macos_excel_launcher.py' in source
expected=json.loads((root/'build/excel-bounded-public-20260910/run-01/report.json').read_text())
raw=json.loads((root/'build/excel-transport-20260910/readonly-01/identity.json').read_text());raw.pop('deadline',None);identity=ExcelProcessIdentity(**raw)
def inventory(label):
 script=root/'build/excel-bounded-public-20260910/run-01/original-final.applescript'
 result=BoundOSAKitTransport(identity).run(script,time.monotonic()+25)
 (folder/(label+'.stdout')).write_text(result.stdout);(folder/(label+'.stderr')).write_text(result.stderr)
 assert result.returncode==0
 return json.loads(result.stdout)
report={'passed':False,'source_hashes':source,'original_before':inventory('before')}
assert report['original_before']==expected['original_before']
command=[sys.executable,str(root/'fixtures/microsoft_parity/installed_public.py'),'--plugin',str(plugin),'--output-dir',str(folder/'sheet'),'--kind','sheet']
with (folder/'native.log').open('w') as log: result=subprocess.run(command,cwd='/tmp',stdout=log,stderr=subprocess.STDOUT)
report['native_exitcode']=result.returncode;report['original_after']=inventory('after')
report['source_equal_after']=all(hashlib.sha256((root/p).read_bytes()).hexdigest()==h and hashlib.sha256((plugin/p).read_bytes()).hexdigest()==h for p,h in source.items())
report['passed']=result.returncode==0 and report['source_equal_after'] and report['original_before']==report['original_after']
(folder/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n');print(json.dumps({k:v for k,v in report.items() if k!='source_hashes'},ensure_ascii=False));raise SystemExit(0 if report['passed'] else 1)
