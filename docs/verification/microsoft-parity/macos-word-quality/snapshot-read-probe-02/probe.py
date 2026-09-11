"""Leased read-only diagnosis with per-block state and native stderr phases."""
from pathlib import Path
import hashlib,json,shutil,sys,traceback,time
ROOT=Path.cwd();sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice import macos_word_quality as q
from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands
from fixtures.microsoft_parity.macos_word_recovery import inventory
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
OUT=Path(__file__).resolve().parent
SOURCE=ROOT/'docs/verification/microsoft-parity/macos-word-quality/run-02/before.docx'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
SOURCES=[Path(__file__),SOURCE,*[ROOT/'skills/WPSComposer/scripts/msoffice'/n for n in ('macos_word_quality.py','macos_word_session.py','macos_word_recovery.py','macos_runtime.py','macos_script.py')]]
report={'status':'FAIL','scope':'one fresh read-only owned copy; no body writes/save/export; ordinary errors are observations','steps':[],'source_hashes':{str(p.relative_to(ROOT)):sha(p) for p in SOURCES},'checks':{}}
s=None

def flush():
    (OUT/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))

def traced(lines,label):
    out=[]
    for i,line in enumerate(lines):
        if not line.startswith(('end ', 'else', 'on error')):
            out+=['set probePhase to '+apple_string(label+':'+str(i)+': '+line[:140]),'log probePhase']
        out.append(line)
    return out

def state(label):
    commands=['log '+apple_string(label),'set probeSavedBefore to saved of boundDoc',
              *hash_commands('content of text object of boundDoc as text','probeBodyHash'),
              'set nativeRows to {{"state",probeSavedBefore,saved of boundDoc,end of content of text object of boundDoc,probeBodyHash,read only of boundDoc}}']
    rows=s._execute(commands)
    if not (len(rows)==1 and len(rows[0])==6 and rows[0][0]=='state' and type(rows[0][1]) is bool and type(rows[0][2]) is bool and type(rows[0][3]) is int and isinstance(rows[0][4],str) and len(rows[0][4])==64 and rows[0][5] is True):
        raise RuntimeError('Unknown state ACK')
    return rows[0]

def read_block(label,lines,kind='rows'):
    report['current_phase']=label;flush()
    step={'label':label,'before':state(label+':before')};report['steps'].append(step);flush()
    if step['before'][3:5]!=report['baseline'][3:5]:raise RuntimeError('Body changed before block')
    commands=['set qualityPoint to 111','set qualityDelta to 0','set qualityState to {}','set qualityLayout to {}','set probeRows to {}','set probePhase to "start"','try',*traced(lines,label)]
    commands+=['set nativeRows to {{"read-ok",'+('qualityState' if kind=='state' else 'probeRows')+'}}',
               'on error probeError number probeNumber',
               'if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber',
               'log {probePhase,probeNumber}',
               'set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}','end try']
    started=time.monotonic();step['rows']=s._execute(commands);step['seconds']=round(time.monotonic()-started,3);flush()
    step['after']=state(label+':after');step['body_unchanged']=step['after'][3:5]==step['before'][3:5];step['saved_transition']=[step['before'][2],step['after'][1]];flush()
    if not step['body_unchanged']:raise RuntimeError('Native body changed; no more AppleEvents')
    print(json.dumps({'label':label,'rows':step['rows'],'saved':step['saved_transition'],'seconds':step['seconds']},ensure_ascii=False),flush=True)
    return step['rows']

try:
    report['inventory_before']=inventory(OUT,'before');flush()
    if report['inventory_before']!=[]:raise RuntimeError('Nonempty inventory; not opening')
    s=MacWordSession.open_document(SOURCE,read_only=True,visible=False);s._retain_evidence=True
    original_execute=s._execute
    s._execute=lambda lines,**kw:original_execute(q._guard(s)+list(lines),**kw)
    report['owned_path']=s._bound_path;report['runtime']=str(s.staging_root);report['baseline']=state('baseline');flush()
    for i in (1,2):
        prefix=[f'set qt to table {i} of boundDoc']
        read_block(f'table-{i}-bounds',prefix+['set qr to text object of qt','set probeRows to {{"bounds",start of content of qr,end of content of qr}}'])
        read_block(f'table-{i}-content-hash',prefix+hash_commands('content of text object of qt as text','probeTableHash')+['set probeRows to {{"table-hash",probeTableHash}}'])
        read_block(f'table-{i}-dimensions',prefix+['set probeRows to {{"dimensions",(count rows of qt) as integer,(count columns of qt) as integer}}'])
        read_block(f'table-{i}-ordinal-rows',prefix+['repeat with probeRowOrdinal from 1 to count rows of qt','set qrow to row probeRowOrdinal of qt','set probeFlag to allow break across pages of qrow','set end of probeRows to {"row",probeRowOrdinal as integer,probeFlag}','end repeat'])
    names=read_block('style-name-list',['set probeNames to get name local of every Word style of boundDoc','set probeRows to {{"names",probeNames}}'])
    styles_ok=False
    style_replacement=[]
    if names[0][0]=='read-ok' and names[0][1] and names[0][1][0][0]=='names' and isinstance(names[0][1][0][1],list):
        style_replacement=['set qualityStyleNames to get name local of every Word style of boundDoc','repeat with qualityStyleOrdinal from 1 to count qualityStyleNames','set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text','set qstyle to Word style qualityStyleName of boundDoc']
        styles=read_block('style-named-full-properties',style_replacement+['set end of probeRows to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}','end repeat'])
        styles_ok=styles[0][0]=='read-ok'
    else:
        got=read_block('style-explicit-get-list',['set probeStyles to get Word styles of boundDoc','set probeRows to {{"local-count",count probeStyles}}'])
        if got[0][0]=='read-ok':
            style_replacement=['set qualityStyles to get Word styles of boundDoc','repeat with qualityStyleOrdinal from 1 to count qualityStyles','set qstyle to contents of item qualityStyleOrdinal of qualityStyles']
            styles=read_block('style-dereferenced-list-properties',style_replacement+['set end of probeRows to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}','end repeat'])
            styles_ok=styles[0][0]=='read-ok'
    snapshot=q._snapshot_commands()
    bookmark_start=snapshot.index('repeat with qi from 1 to count bookmarks of boundDoc')
    read_block('bookmarks',snapshot[bookmark_start:],'state')
    if styles_ok:
        revised=[]
        for line in snapshot:
            if line=='repeat with qstyle in Word styles of boundDoc':revised.extend(style_replacement)
            elif line=='repeat with qrow in rows of qt':revised.extend(['repeat with qualityRowOrdinal from 1 to count rows of qt','set qrow to row qualityRowOrdinal of qt'])
            else:revised.append(line)
        (OUT/'snapshot-candidate-lines.json').write_text(json.dumps(revised,indent=2))
        full=read_block('complete-snapshot-candidate',revised,'state')
        report['checks']['complete_snapshot_valid']=full[0][0]=='read-ok' and q._valid_state(full[0][1])
    else:
        report['blocker']='All-style enumeration failed; complete snapshot remains blocked for all quality table/block paths. Reservation unaffected.'
    report['after']=state('final-state');report['checks']['body_unchanged']=report['after'][3:5]==report['baseline'][3:5]
    report['checks']['private_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    # Guarded close of this exact owned document uses saving no; no save/export.
    s.close();report['owned_closed']=s._closed
    report['inventory_final']=inventory(OUT,'final');report['checks']['inventory_empty']=report['inventory_final']==[]
    report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
    report['status']='DIAGNOSED' if all(v for k,v in report['checks'].items() if k!='complete_snapshot_valid') else 'FAIL'
except BaseException as exc:
    report['error']={'type':type(exc).__name__,'message':str(exc)};(OUT/'failure.txt').write_text(traceback.format_exc())
    if s and not s._closed:
        s._retain('Read-only snapshot probe uncertainty; parent reconciliation required');report['quarantined']=s._quarantined;report['remaining_owned_path']=s._bound_path
finally:
    if s and s.staging_root.exists():shutil.copytree(s.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
    flush()
print(json.dumps({'status':report['status'],'checks':report['checks'],'blocker':report.get('blocker'),'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False),flush=True)
raise SystemExit(0 if report['status']=='DIAGNOSED' else 1)
