import json, hashlib, subprocess, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from skills.WPSComposer.scripts.msoffice.windows_host import _word_processes

root = Path('build/msoffice-production')
labels = ['smoke-05','representative-03','numbering-03','representative-04',
          'schemes-word-01','native-edit-01','native-edit-02','native-edit-03','linked-style-word-01']
events = []
for label in labels:
    for path in (root/'traces'/label).glob('*.jsonl'):
        for line in path.read_text(encoding='utf-8').splitlines():
            event = json.loads(line)
            if 'word_pid' in event:
                events.append({'label':label, **event})
owned = sorted({e['word_pid'] for e in events if e['event']=='composer_created'})
observed = _word_processes()
for _ in range(10):
    if not any(pid in observed for pid in owned): break
    time.sleep(.5); observed = _word_processes()
preserved = [9252,19732,14432,20960,15152,26492,14848,6332,22736,23588]
checks = {'all_recorded_success_or_cancelled_owned_word_pids_absent':all(pid not in observed for pid in owned),
          'preexisting_registered_word_9252_still_present':9252 in observed}
report = {'head':subprocess.check_output(['git','rev-parse','HEAD']).decode().strip(),
          'checks':checks,'owned_pids':owned,'owned_still_running':[pid for pid in owned if pid in observed],
          'preserved_process_observations':{str(pid):'present' if pid in observed else 'absent' for pid in preserved},
          'notes':['No retained or user Word process was killed or Quit by this postflight.',
                   'native-edit-02 was cancelled through its exact ROT document, verified PID 23856, then zero-doc dedicated application Quit; raw interrupted helper error is retained.',
                   'Old empty probes and the earlier uncertain timeout host are reported separately, never classified as cleaned up.',
                   'User-document content isolation is the prior sentinel-03 native observation; this final process read does not re-read user document text.'],
          'status':'PASS' if all(checks.values()) else 'FAIL','events':events}
(root/'final-postflight-round4.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'status':report['status'],'checks':checks,'owned_pids':owned,'owned_still_running':report['owned_still_running']}))
raise SystemExit(0 if report['status']=='PASS' else 1)
