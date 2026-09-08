from pathlib import Path
import subprocess, sys, json, hashlib, datetime, time
import pythoncom, win32com.client

root = Path.cwd()
evidence = root / 'build/msoffice-spike/audit-native-round3'
evidence.mkdir(exist_ok=False)
candidate = 'd4884a55fe8a0db48d1f2f48543948959aa08dbf'
def write(name, value):
    (evidence / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
source = {}
for name in ['windows_word.py', 'windows_unsaved_sentinel.py', 'validate_windows_artifacts.py']:
    path = 'fixtures/msoffice_spike/' + name
    raw = Path(path).read_bytes()
    blob = subprocess.check_output(['git', 'show', candidate + ':' + path])
    assert raw.replace(b'\r\n', b'\n') == blob, name
    (evidence / ('executed-' + name + '.txt')).write_bytes(raw)
    source[name] = {'raw_sha256': hashlib.sha256(raw).hexdigest(), 'lf_sha256': hashlib.sha256(blob).hexdigest(), 'matches_candidate': True}
write('source-provenance.json', {'candidate': candidate, 'execution_head': subprocess.check_output(['git','rev-parse','HEAD'], text=True).strip(), 'files': source})
pythoncom.CoInitialize()
app = win32com.client.GetActiveObject('Word.Application')
preflight = {'name': str(app.Name), 'documents_count': int(app.Documents.Count), 'path': str(app.Path), 'version': str(app.Version), 'build': str(app.Build), 'observed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat()}
write('registered-preflight.json', preflight)
assert preflight['name'] == 'Microsoft Word' and preflight['documents_count'] == 0, 'Registered instance not empty Microsoft Word; no document created'
app = None
def run(label, args):
    start = time.monotonic()
    stamp = datetime.datetime.now(datetime.timezone.utc).isoformat()
    completed = subprocess.run([sys.executable] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    (evidence / (label + '.stdout.log')).write_bytes(completed.stdout)
    (evidence / (label + '.stderr.log')).write_bytes(completed.stderr)
    write(label + '-command.json', {'command': [sys.executable] + args, 'exit_code': completed.returncode, 'started_utc': stamp, 'completed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'duration_seconds': round(time.monotonic()-start,3)})
    print(label, completed.returncode, flush=True)
    return completed.returncode
status = run('sentinel', ['fixtures/msoffice_spike/windows_unsaved_sentinel.py','--evidence-dir','build/msoffice-spike/windows-sentinel-04','--output-dir','build/msoffice-spike/windows-run-11'])
if status:
    raise SystemExit(status)
status = run('validator', ['fixtures/msoffice_spike/validate_windows_artifacts.py','build/msoffice-spike/windows-run-11'])
sys.path.insert(0, str(root / 'fixtures/msoffice_spike'))
from windows_word import word_processes
sentinel = json.loads((root / 'build/msoffice-spike/windows-sentinel-04/sentinel-result.json').read_text('utf-8'))
processes = word_processes()
app = win32com.client.GetActiveObject('Word.Application')
write('postflight.json', {'registered_name': str(app.Name), 'registered_documents_count': int(app.Documents.Count), 'sentinel_pid': sentinel.get('sentinel_application', {}).get('pid'), 'sentinel_pid_still_present': sentinel.get('sentinel_application', {}).get('pid') in processes, 'runner_pid': sentinel.get('runner_application_pid'), 'runner_pid_absent': sentinel.get('runner_application_pid') not in processes, 'no_quit_or_kill_in_postflight': True, 'observed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat()})
raise SystemExit(status)
