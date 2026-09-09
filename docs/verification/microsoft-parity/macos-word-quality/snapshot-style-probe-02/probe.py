"""Leased style classification diagnostic; bounded fresh copy, no document writes."""
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
              'set nativeRows to {{"state",probeSavedBefore,saved of boundDoc,end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,(count list templates of boundDoc) as integer}}']
    rows=s._execute(commands)
    if not (len(rows)==1 and len(rows[0])==7 and rows[0][0]=='state' and type(rows[0][1]) is bool and type(rows[0][2]) is bool and type(rows[0][3]) is int and isinstance(rows[0][4],str) and len(rows[0][4])==64 and rows[0][5] is True):
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
    print(json.dumps({'label':label,'row_kind':step['rows'][0][0],'saved':step['saved_transition'],'seconds':step['seconds'],'list_templates':[step['before'][6],step['after'][6]]},ensure_ascii=False),flush=True)
    return step['rows']

try:
    report['inventory_before']=inventory(OUT,'before');flush()
    if report['inventory_before']!=[]:raise RuntimeError('Nonempty inventory; not opening')
    s=MacWordSession.open_document(SOURCE,read_only=True,visible=False);s._retain_evidence=True
    original_execute=s._execute
    s._execute=lambda lines,**kw:original_execute(q._guard(s)+list(lines),**kw)
    report['owned_path']=s._bound_path;report['runtime']=str(s.staging_root);report['baseline']=state('baseline');flush()
    vectors={};safe=True
    def stable(step):return step['after'][1:3]==[True,True] and step['after'][6]==report['baseline'][6]
    for label,prop in [('names','name local'),('in-use','in use'),('built-in','built in')]:
        rows=read_block('style-bulk-'+label,[f'set probeValues to get {prop} of every Word style of boundDoc','set probeRows to {{"values",probeValues}}'])
        if rows[0][0]=='read-ok' and len(rows[0][1])==1 and rows[0][1][0][0]=='values' and isinstance(rows[0][1][0][1],list):vectors[label]=rows[0][1][0][1]
        else:safe=False
        if not stable(report['steps'][-1]):safe=False
        if not safe:break
    report['vector_lengths']={k:len(v) for k,v in vectors.items()};report['checks']['bulk_getters_stable']=safe;flush()
    classified=set(vectors)=={'names','in-use','built-in'} and len(set(map(len,vectors.values())))==1 and all(type(x) is bool for k in ('in-use','built-in') for x in vectors[k])
    if safe and classified:
        defined=[name for name,used,builtin in zip(vectors['names'],vectors['in-use'],vectors['built-in']) if used or not builtin]
        report['defined_candidates']=defined;flush()
        for label,prop,cast in [('description','description',' as text'),('auto-update','automatically update','')]:
            lines=[]
            for name in defined:
                lines+=['set qstyle to Word style '+apple_string(name)+' of boundDoc',
                        'set probeValue to '+prop+' of qstyle'+cast,
                        'set end of probeRows to {'+apple_string(label)+','+apple_string(name)+',probeValue}']
            rows=read_block('defined-'+label,lines)
            safe=rows[0][0]=='read-ok' and stable(report['steps'][-1])
            report['checks']['defined_'+label+'_stable']=safe;flush()
            if not safe:break
        if safe:
            replacement=[
                'set qualityStyleNames to get name local of every Word style of boundDoc',
                'set qualityStyleInUse to get in use of every Word style of boundDoc',
                'set qualityStyleBuiltIn to get built in of every Word style of boundDoc',
                'if (count qualityStyleNames) is not (count qualityStyleInUse) or (count qualityStyleNames) is not (count qualityStyleBuiltIn) then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"',
                'repeat with qualityStyleOrdinal from 1 to count qualityStyleNames',
                'set qualityStyleName to item qualityStyleOrdinal of qualityStyleNames as text',
                'set qualityStyleUsed to item qualityStyleOrdinal of qualityStyleInUse',
                'set qualityStyleBuiltin to item qualityStyleOrdinal of qualityStyleBuiltIn',
                'set end of qualityLayout to {"style-state",qualityStyleName,qualityStyleUsed,qualityStyleBuiltin}',
                'if qualityStyleUsed or not qualityStyleBuiltin then',
                'set qstyle to Word style qualityStyleName of boundDoc',
                'if (name local of qstyle as text) is not qualityStyleName then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"',
                'set end of qualityLayout to {"style",name local of qstyle as text,description of qstyle as text,automatically update of qstyle}',
                'end if','end repeat',
                'if (get name local of every Word style of boundDoc) is not qualityStyleNames then error "WPSC_QUALITY_STYLE_IDENTITY_UNVERIFIED"']
            snapshot=q._snapshot_commands();start=snapshot.index('repeat with qstyle in Word styles of boundDoc')
            revised=snapshot[:start]+replacement+snapshot[start+3:]
            (OUT/'snapshot-candidate-lines.json').write_text(json.dumps(revised,indent=2))
            full=read_block('complete-snapshot-defined-styles',revised,'state')
            report['checks']['complete_snapshot_valid']=full[0][0]=='read-ok' and q._valid_state(full[0][1]);flush()
            if report['checks']['complete_snapshot_valid'] and stable(report['steps'][-1]):
                second=read_block('complete-snapshot-repeat',revised,'state')
                report['checks']['repeat_snapshot_equal']=second==full
                report['checks']['repeat_snapshot_stable']=stable(report['steps'][-1])
    else:report['blocker']='Style classification unavailable, malformed, or dirty; no definition filtering performed.'
    report['after']=state('final-state');report['checks']['body_unchanged']=report['after'][3:5]==report['baseline'][3:5]
    report['checks']['private_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    s.close();report['owned_closed']=s._closed
    report['inventory_final']=inventory(OUT,'final');report['checks']['inventory_empty']=report['inventory_final']==[]
    report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
    report['status']='DIAGNOSED' if report['checks']['body_unchanged'] and report['checks']['private_file_unchanged'] and report['checks']['inventory_empty'] and report['checks']['sources_unchanged'] else 'FAIL'
except BaseException as exc:
    report['error']={'type':type(exc).__name__,'message':str(exc)};(OUT/'failure.txt').write_text(traceback.format_exc())
    if s and not s._closed:
        s._retain('Read-only style probe uncertainty; parent reconciliation required');report['quarantined']=s._quarantined;report['remaining_owned_path']=s._bound_path
finally:
    if s and s.staging_root.exists():shutil.copytree(s.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
    flush()
print(json.dumps({'status':report['status'],'checks':report['checks'],'blocker':report.get('blocker'),'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False),flush=True)
raise SystemExit(0 if report['status']=='DIAGNOSED' else 1)
