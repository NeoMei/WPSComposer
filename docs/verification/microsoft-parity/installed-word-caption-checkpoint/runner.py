from pathlib import Path
import hashlib,json,sys,subprocess,tempfile
root=Path('/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description');sys.path.insert(0,str(root))
from fixtures.microsoft_parity.macos_word_recovery import inventory
out=root/'build/installed-caption-20260911';out.mkdir(exist_ok=False)
base=Path(tempfile.mkdtemp(prefix='wpscomposer-installed-caption-'));plugin=base/'codex/plugins/wps-composer'
command=[sys.executable,str(root/'install.py'),'--codex-home',str(base/'codex'),'--marketplace-root',str(base)]
with (out/'install.log').open('w') as log:r=subprocess.run(command,cwd=root,stdout=log,stderr=subprocess.STDOUT)
(out/'install.json').write_text(json.dumps({'returncode':r.returncode,'plugin':str(plugin),'command':command},indent=2));assert r.returncode==0
hashes={str(p.relative_to(plugin)):hashlib.sha256(p.read_bytes()).hexdigest() for p in plugin.rglob('*') if p.is_file() and not {'node_modules','__pycache__'} & set(p.parts) and (root/p.relative_to(plugin)).is_file()}
def eq():return all(hashlib.sha256((root/p).read_bytes()).hexdigest()==h and hashlib.sha256((plugin/p).read_bytes()).hexdigest()==h for p,h in hashes.items())
assert eq()
report={'passed':False,'commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=root,text=True).strip(),'source_hashes':hashes,'inventory_before':inventory(out,'before')}
assert report['inventory_before']==[]
cmd=[sys.executable,str(root/'fixtures/microsoft_parity/installed_public.py'),'--plugin',str(plugin),'--output-dir',str(out/'writer'),'--kind','writer']
with (out/'writer.log').open('w') as log:r=subprocess.run(cmd,cwd='/tmp',stdout=log,stderr=subprocess.STDOUT)
report.update(native_returncode=r.returncode,inventory_after=inventory(out,'after'),source_equal_after=eq(),native=json.loads((out/'writer/report.json').read_text()))
report['passed']=r.returncode==0 and report['native']['passed'] and report['source_equal_after'] and report['inventory_before']==report['inventory_after']
(out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps({k:v for k,v in report.items() if k not in ('source_hashes','native')},ensure_ascii=False));print('source_count',len(hashes));raise SystemExit(0 if report['passed'] else 1)
