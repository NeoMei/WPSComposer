from pathlib import Path
import json,time,shutil,sys,hashlib
from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity,BoundOSAKitTransport
from skills.WPSComposer.scripts.msoffice.macos_office_runtime import _read_excel_process_identity,recover_quarantine
out=Path(sys.argv[1]);expected_pid=int(sys.argv[2]);assert expected_pid!=9212
report=json.loads((out/'report.json').read_text());assert not report['passed'] and not report['jobs']
marker=Path('/Users/neomei/Library/Containers/com.microsoft.Excel/Data/tmp/wpscomposer/native-office.quarantine.json');d=json.loads(marker.read_text());raw=d['private_process']['identity'];assert raw['pid']==expected_pid
assert d['private_process']['owned_path'] is None
(out/'original-quarantine.json').write_bytes(marker.read_bytes());shutil.copytree(d['staging_path'],out/'failed-runtime')
p=out/'parent-exact-cleanup.applescript';p.write_bytes(Path('build/excel-private-public-20260910/run-01/parent-exact-cleanup.applescript').read_bytes())
r=BoundOSAKitTransport(ExcelProcessIdentity(**raw)).run(p,time.monotonic()+30)
p.with_suffix('.stdout.log').write_text(r.stdout);p.with_suffix('.stderr.log').write_text(r.stderr)
assert r.returncode==0 and r.stdout.strip()=='EXACT_BOOTSTRAP_CLEANUP_PASS'
end=time.monotonic()+10
while _read_excel_process_identity(expected_pid) is not None and time.monotonic()<end:time.sleep(.2)
assert _read_excel_process_identity(expected_pid) is None
assert recover_quarantine('spreadsheet')
(out/'parent-cleanup-result.json').write_text(json.dumps({'cleanup_ack':True,'identity':raw,'exact_pid_gone':True,'recovered':True,'runner_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}));print('exact private bootstrap cleaned; quarantine recovered')
