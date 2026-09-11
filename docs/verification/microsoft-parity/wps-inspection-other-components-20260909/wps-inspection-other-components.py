from pathlib import Path
import hashlib,json,sys,traceback
ROOT=Path('/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description');sys.path.insert(0,str(ROOT))
from skills.WPSComposer import inspect
from skills.WPSComposer.scripts.macos_probe.runtime import list_wps_pids

def main():
 out=ROOT/'docs/verification/microsoft-parity/wps-inspection-other-components-20260909';out.mkdir(exist_ok=False)
 results={'passed':False,'components':{},'wps_pids_before':sorted(list_wps_pids(Path('/Applications/wpsoffice.app')))}
 for kind,suffix,marker in [('writer','docx','Installed Word edited'),('sheet','xlsx','Installed native Excel')]:
  source=ROOT/'docs/verification/microsoft-parity/installed-audit-04'/kind/('edited.'+suffix); before=hashlib.sha256(source.read_bytes()).hexdigest()
  try:
   snapshot=inspect(source,engine='wps',timeout=60);(out/(kind+'.json')).write_text(json.dumps(snapshot,ensure_ascii=False,indent=2))
   assert marker in json.dumps(snapshot,ensure_ascii=False);assert hashlib.sha256(source.read_bytes()).hexdigest()==before
   results['components'][kind]={'passed':True,'source_sha256':before}
  except BaseException as e:
   results['components'][kind]={'passed':False,'error':str(e)};(out/(kind+'-error.txt')).write_text(traceback.format_exc())
 results['wps_pids_after']=sorted(list_wps_pids(Path('/Applications/wpsoffice.app')))
 results['passed']=all(c['passed'] for c in results['components'].values()) and results['wps_pids_after']==results['wps_pids_before']
 (out/'report.json').write_text(json.dumps(results,indent=2));print(json.dumps(results));return 0 if results['passed'] else 1
if __name__=='__main__':raise SystemExit(main())
