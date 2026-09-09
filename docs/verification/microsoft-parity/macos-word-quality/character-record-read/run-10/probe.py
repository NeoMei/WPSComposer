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
SOURCE=OUT/'format-variants.docx'
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
              'set nativeRows to {{"state",probeSavedBefore,saved of boundDoc,end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,(count list templates of boundDoc) as integer,(count sections of boundDoc),(count paragraphs of boundDoc),(count fields of boundDoc),(count tables of boundDoc),(count bookmarks of boundDoc)}}']
    rows=s._execute(commands)
    if not (len(rows)==1 and len(rows[0])==12 and rows[0][0]=='state' and type(rows[0][1]) is bool and type(rows[0][2]) is bool and type(rows[0][3]) is int and isinstance(rows[0][4],str) and len(rows[0][4])==64 and rows[0][5] is True):
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
    if step['after'][1:3] != [True,True]:
        report['first_dirty_phase']=label;flush();raise DirtyRead(label)
    return step['rows']

class DirtyRead(Exception): pass
try:
    report['inventory_before']=inventory(OUT,'before');flush()
    if report['inventory_before']!=[]:raise RuntimeError('Nonempty inventory; not opening')
    s=MacWordSession.open_document(SOURCE,read_only=True,visible=False);s._retain_evidence=True
    original_execute=s._execute
    s._execute=lambda lines,**kw:original_execute(q._guard(s)+list(lines),**kw)
    report['owned_path']=s._bound_path;report['runtime']=str(s.staging_root);report['baseline']=state('baseline');flush()
    blocks=[]
    for begin in range(0,220,10):
        stop=min(begin+10,220)
        lines=[f'set probeFirst to {begin}',f'set probeLast to {stop-1}',
            'repeat with probePosition from probeFirst to probeLast',
            'set probeRange to create range boundDoc start probePosition end (probePosition+1)',
            'set probeFontReference to font object of probeRange',
            'set probeLegacyRow to {name of probeFontReference as text,font size of probeFontReference,bold of probeFontReference,italic of probeFontReference,underline of probeFontReference as text,color of probeFontReference,background pattern color of shading of probeRange}',
            'set probeFontRecord to (get properties of font object of probeRange) as record',
            'set probeShadingRecord to (get properties of shading of probeRange) as record',
            'set probeRecordRow to {name of probeFontRecord as text,font size of probeFontRecord,bold of probeFontRecord,italic of probeFontRecord,underline of probeFontRecord as text,color of probeFontRecord,background pattern color of probeShadingRecord}',
            'set probeLegacyJson to my jsonRows({probeLegacyRow})','set probeRecordJson to my jsonRows({probeRecordRow})',
            "set probeEqual to (current application's NSString's stringWithString:probeLegacyJson)'s isEqualToString:probeRecordJson",
            'if not (probeEqual as boolean) then error "RECORD_VALUE_MISMATCH"',
            'set end of probeRows to {"same-coordinate",probePosition as integer,probeEqual as boolean,probeLegacyRow}', 'end repeat']
        blocks.append(('record-vs-legacy-'+str(begin)+'-'+str(stop),lines))
    # Execute the real production-generated character loop first on this fresh copy.
    generated=q._snapshot_commands()
    first=generated.index('repeat with qc from qa to qz - 1')
    last=generated.index('end repeat',first)
    actual_loop=generated[first:last+1]
    prefix=['set qa to 0','set qz to 220','set qualityFormats to {}']
    full_lines=prefix+actual_loop+['set probeRows to qualityFormats']
    blocks=[('production-first-format-'+str(i),full_lines) for i in range(2)]+blocks
    try:
        for label,lines in blocks:
            result=read_block(label,lines)
            if result[0][0]!='read-ok': raise RuntimeError('Native format read did not succeed')
        full=[v['rows'] for v in report['steps'] if v['label'].startswith('production-first-format')]
        report['checks']['all_coordinates_equal']=sum(len(v['rows'][0][1]) for v in report['steps'] if v['label'].startswith('record-vs-legacy'))==220
        report['checks']['full_format_repeat_equal']=len(full)==2 and full[0]==full[1] and len(full[0][0][1])==220
        report['checks']['explicit_format_variant_observed']=['Arial',17,True,True,'underline single',[17,34,51],[238,232,170]] in full[0][0][1]
        expected={position:row for position,row in enumerate(full[0][0][1])}
        # Every later per-coordinate legacy comparison must equal the first production read.
        report['checks']['production_matches_legacy']=all(step['rows'][0][1]==[["same-coordinate",pos,True,expected[pos]] for pos in range(int(step['label'].split('-')[-2]),int(step['label'].split('-')[-1]))] for step in report['steps'] if step['label'].startswith('record-vs-legacy'))

    except DirtyRead as exc: report['blocker']='Native getter dirtied copy at '+str(exc)
    report['after']=state('final-state');report['checks']['body_unchanged']=report['after'][3:5]==report['baseline'][3:5]
    report['checks']['private_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    s.close();report['owned_closed']=s._closed
    report['inventory_final']=inventory(OUT,'final');report['checks']['inventory_empty']=report['inventory_final']==[]
    report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
    report['status']='PASS' if all(report['checks'].get(k) is True for k in ('all_coordinates_equal','full_format_repeat_equal','explicit_format_variant_observed','production_matches_legacy','body_unchanged','private_file_unchanged','inventory_empty','sources_unchanged')) and not any(report.get(k) for k in ('blocker','error','first_dirty_phase')) else 'FAIL'
except BaseException as exc:
    report['error']={'type':type(exc).__name__,'message':str(exc)};(OUT/'failure.txt').write_text(traceback.format_exc())
    if s and not s._closed:
        s._retain('Read-only snapshot probe uncertainty; parent reconciliation required');report['quarantined']=s._quarantined;report['remaining_owned_path']=s._bound_path
finally:
    if s and s.staging_root.exists():shutil.copytree(s.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
    flush()
print(json.dumps({'status':report['status'],'checks':report['checks'],'blocker':report.get('blocker'),'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False),flush=True)
raise SystemExit(0 if report['status']=='PASS' else 1)
