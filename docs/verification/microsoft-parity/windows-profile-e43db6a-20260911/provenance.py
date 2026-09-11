from pathlib import Path
import datetime,hashlib,importlib,json,subprocess,sys
root=Path('D:/wpsce43');out=Path(__file__).parent;phase=sys.argv[1]
sys.path.insert(0,str(root))
module=importlib.import_module('skills.WPSComposer.scripts.macos_probe.profile_server')
assert Path(module.__file__).resolve()==(root/'skills/WPSComposer/scripts/macos_probe/profile_server.py').resolve()
def git(*args):return subprocess.check_output(['git','-C',str(root),*args]).decode()
entries={}
for row in git('ls-tree','-r','-z','HEAD').split('\0'):
 if row:
  meta,name=row.split('\t',1)
  if meta.split()[1]=='blob':entries[name]=meta.split()[2]
files=[]
for row in git('ls-files','-v','-z').split('\0'):
 if not row or row.startswith('S '):continue
 name=row[2:];raw=(root/name).read_bytes()
 blob=hashlib.sha1(b'blob '+str(len(raw)).encode()+b'\0'+raw).hexdigest()
 files.append({'path':name,'sha256':hashlib.sha256(raw).hexdigest(),'git_blob':blob,'matches_commit':blob==entries[name]})
markers={}
for component in ('writer','spreadsheet','presentation'):
 p=Path('C:/Users/1/AppData/Local/WPSComposer/native-office')/component/'native-office.quarantine.json'
 markers[component]={'exists':p.exists(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest() if p.exists() else None}
ps="Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(WINWORD|EXCEL|POWERPNT|wps|et|wpp)\\.exe$' } | Select-Object ProcessId,Name,ExecutablePath,CreationDate | ConvertTo-Json"
p=subprocess.run(['pwsh','-NoProfile','-Command',ps],capture_output=True);assert p.returncode==0
report={'time':datetime.datetime.now(datetime.timezone.utc).isoformat(),'python':sys.executable,'python_version':sys.version,'module_import_path':module.__file__,'handler_timeout':module.ProfileRequestHandler.timeout,'head':git('rev-parse','HEAD').strip(),'status':git('status','--short'),'files':files,'quarantines':markers,'office_processes':json.loads(p.stdout.decode('utf-8-sig') or '[]'),'process_query_exit_code':p.returncode}
assert report['head']=='e43db6a391bbf8a7f78e6d5f2da46c4e68b4ffea'
assert all(f['matches_commit'] for f in files)
if phase=='after':
 before=json.loads((out/'before.json').read_text())
 report['sources_unchanged']=before['files']==files
 report['quarantines_unchanged']=before['quarantines']==markers
 report['office_processes_unchanged']=before['office_processes']==report['office_processes']
 assert report['sources_unchanged'] and report['quarantines_unchanged']
target=out/(phase+'.json');assert not target.exists();target.write_text(json.dumps(report,indent=2,ensure_ascii=False),encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k not in ('files','office_processes')}|{'materialized_files':len(files)},ensure_ascii=False))
