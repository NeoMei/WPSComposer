from pathlib import Path
import json,subprocess,hashlib
from fixtures.microsoft_parity.macos_word_recovery import inventory
from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands
from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_runtime import recover_quarantine
from fixtures.microsoft_parity.macos_word_heading_import import SEED
p=Path('build/word-heading-import-20260910/run-01-parent-cleanup');before=json.loads((p/'before.log').read_text()); report={}
owned=[r for r in before if '/wpscomposer-session-b03g3kic/' in r[1]];assert len(owned)==1 and owned[0][2] is True and owned[0][3]==hashlib.sha256((SEED+'\r').encode()).hexdigest()
sentinel=json.loads(Path('build/word-heading-import-20260910/run-01/report.json').read_text())['sentinel_before'];assert sentinel in before and len(before)==2
for label,row in [('owned',owned[0]),('sentinel',sentinel)]:
 name,path,saved,digest=row
 lines=['tell application "Microsoft Word"','set matches to {}','repeat with di from 1 to count documents','set d to document di',f'if (posix full name of d as text) is {apple_string(path)} then set end of matches to di','end repeat','if (count of matches) is not 1 then error "CLEANUP_IDENTITY"','set d to document (item 1 of matches)',f'if (name of d as text) is not {apple_string(name)} then error "CLEANUP_NAME"',f'if saved of d is not {str(saved).lower()} then error "CLEANUP_SAVED"',*hash_commands('content of text object of d as text','actualHash'),f'if actualHash is not {apple_string(digest)} then error "CLEANUP_CONTENT"','close d saving no',f'set nativeRows to {{{{"closed",{apple_string(path)}}}}}','end tell','return my jsonRows(nativeRows)']
 f=p/(label+'.applescript');f.write_text(_JSON+'\n'.join(lines));r=subprocess.run(['/usr/bin/osascript',str(f)],capture_output=True,text=True,timeout=30);(p/(label+'.log')).write_text(r.stdout+r.stderr);assert r.returncode==0 and json.loads(r.stdout)==[['closed',path]]
 report[label+'_closed']=True
report['inventory_after']=inventory(p,'after');assert report['inventory_after']==[]
report['quarantine_recovered']=recover_quarantine(timeout=30);report['inventory_final']=inventory(p,'final');assert report['inventory_final']==[]
report['source_preimage_sha256']=hashlib.sha256(Path('build/word-heading-import-20260910/run-01/source-preimage.docx').read_bytes()).hexdigest();report['passed']=True
(p/'report.json').write_text(json.dumps(report,indent=2)+'\n');print(report)
