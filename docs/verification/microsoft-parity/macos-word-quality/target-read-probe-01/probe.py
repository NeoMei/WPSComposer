"""One leased read-only comparison probe. Never mutates document contents."""
from pathlib import Path
import hashlib
import json
import shutil
import sys
import traceback

ROOT=Path.cwd();sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice import macos_word_quality as q
from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands
from fixtures.microsoft_parity.macos_word_recovery import inventory

OUT=Path(__file__).resolve().parent
SOURCE=ROOT/'docs/verification/microsoft-parity/macos-word-quality/run-01/before.docx'
SOURCES=[Path(__file__),SOURCE,*[ROOT/'skills/WPSComposer/scripts/msoffice'/n for n in ('macos_word_quality.py','macos_word_session.py','macos_word_recovery.py','macos_runtime.py','macos_script.py')]]
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
report={'status':'FAIL','scope':'fresh owned read-only copy; expression comparison only; no quality insertion','checks':{},'steps':[],'source_hashes':{str(p.relative_to(ROOT)):sha(p) for p in SOURCES}}
s=None

def flush():
    (OUT/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))

def commands(mode):
    lines=['set qualityPoint to 111','set probePhase to "point"','set nativeRows to {}','try',
           'set probeLower to qualityPoint - 1','set probeUpper to qualityPoint + 1',
           'set end of nativeRows to {"bounds",qualityPoint,probeLower,probeUpper}']
    if mode in ('original','foreach-atomic'):
        lines+=['set probeOrdinal to 0','repeat with qt in tables of boundDoc','set probeOrdinal to probeOrdinal + 1']
    else:
        lines+=['repeat with probeOrdinal from 1 to count tables of boundDoc','set qt to table probeOrdinal of boundDoc']
    if mode in ('original','indexed-original'):
        lines+=['set probePhase to "compound-comparison"',
                'set probeHit to false',
                'if (start of content of text object of qt) <= qualityPoint + 1 and (end of content of text object of qt) >= qualityPoint - 1 then set probeHit to true',
                'set end of nativeRows to {"compound",probeOrdinal as integer,probeHit}']
    else:
        lines+=['set probePhase to "range"','set probeRange to text object of qt',
                'set probePhase to "start"','set probeStart to (start of content of probeRange) as integer',
                'set probePhase to "end"','set probeEnd to (end of content of probeRange) as integer',
                'set probePhase to "left-comparison"','set probeLeft to (probeStart <= probeUpper)',
                'set probePhase to "right-comparison"','set probeRight to (probeEnd >= probeLower)',
                'set end of nativeRows to {"atomic",probeOrdinal as integer,probeStart,probeEnd,probeLeft,probeRight,(probeLeft and probeRight)}']
    lines+=['end repeat','set end of nativeRows to {"done"}',
            'on error probeError number probeNumber',
            # The expected ordinary read error is acknowledged as data. Unknown
            # completion is never caught/continued, and no further AE follows it.
            'if probeNumber is not -2763 then error probeError number probeNumber',
            'set end of nativeRows to {"expected-read-error",probePhase,probeNumber}',
            'end try']
    return lines

try:
    report['inventory_before']=inventory(OUT,'before');flush()
    if report['inventory_before']!=[]:raise RuntimeError('Expected empty starting inventory; lease must reconcile')
    s=MacWordSession.open_document(SOURCE,read_only=True,visible=False);s._retain_evidence=True
    report['owned_path']=s._bound_path;report['runtime']=str(s.staging_root);flush()
    original_execute=s._execute
    def guarded_execute(lines,**kwargs):return original_execute(q._guard(s)+list(lines),**kwargs)
    s._execute=guarded_execute
    rows=s._execute(hash_commands('content of text object of boundDoc as text','probeBodyHash')+['set nativeRows to {{"before",end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,version as text}}'])
    report['before_body']=rows;flush()
    if not (len(rows)==1 and rows[0][0:2]==['before',220] and rows[0][3] is True):raise RuntimeError('Readonly fixture baseline mismatch')
    for mode in ('original','atomic','indexed-original','foreach-atomic'):
        report['current_phase']=mode;flush()
        rows=s._execute(commands(mode));report['steps'].append({'mode':mode,'rows':rows});flush()
        if mode in ('original','indexed-original'):
            expected=[['bounds',111,110,112],['expected-read-error','compound-comparison',-2763]]
        else:
            expected=[['bounds',111,110,112],['atomic',1,59,77,True,False,False],['atomic',2,173,191,False,True,False],['done']]
        if rows!=expected:raise RuntimeError('Unexpected typed probe outcome; no further native calls')
    rows=s._execute(hash_commands('content of text object of boundDoc as text','probeBodyHash')+['set nativeRows to {{"after",end of content of text object of boundDoc,probeBodyHash,read only of boundDoc}}'])
    report['after_body']=rows
    report['checks']['body_unchanged']=rows==[['after',220,report['before_body'][0][2],True]]
    if not report['checks']['body_unchanged']:raise RuntimeError('Body changed')
    report['checks']['private_copy_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    if not report['checks']['private_copy_file_unchanged']:raise RuntimeError('Private copy changed')
    s.close();report['owned_closed']=s._closed
    report['inventory_final']=inventory(OUT,'final')
    report['checks']['inventory_empty']=report['inventory_final']==[]
    report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
    report['status']='PASS' if all(report['checks'].values()) else 'FAIL'
except BaseException as exc:
    report['error']={'type':type(exc).__name__,'message':str(exc)}
    (OUT/'failure.txt').write_text(traceback.format_exc())
    if s and not s._closed:
        s._retain('Read-only quality probe outcome unverified; parent reconciliation required')
        report['quarantined']=s._quarantined
        report['remaining_owned_path']=s._bound_path
finally:
    if s and s.staging_root.exists():shutil.copytree(s.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
    flush()
print(json.dumps({'status':report['status'],'steps':report['steps'],'checks':report['checks'],'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False))
raise SystemExit(0 if report['status']=='PASS' else 1)
