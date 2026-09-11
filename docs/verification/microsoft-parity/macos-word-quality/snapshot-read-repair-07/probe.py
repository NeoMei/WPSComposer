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
    blocks=[('character-count',['set probeRows to {{"count",(count characters of text object of boundDoc) as integer}}'])]
    vectors=[('start','start of content of every character of text object of boundDoc','integer'),('end','end of content of every character of text object of boundDoc','integer'),('font-name','name of font object of every character of text object of boundDoc','text'),('font-size','font size of font object of every character of text object of boundDoc','real'),('bold','bold of font object of every character of text object of boundDoc','boolean'),('italic','italic of font object of every character of text object of boundDoc','boolean'),('underline','underline of font object of every character of text object of boundDoc','text'),('color','color of font object of every character of text object of boundDoc','rgb'),('shading','background pattern color of shading of every character of text object of boundDoc','rgb')]
    for label,expression,kind in vectors:
        lines=['set probeValues to "unresolved-sentinel"','set probeValues to get '+expression,'set probeVectorClass to (class of probeValues) as text','if class of probeValues is not list then error "BULK_VALUES_NOT_LIST"','set probeScalarValues to {}','repeat with probeValueOrdinal from 1 to count probeValues','set probeValue to item probeValueOrdinal of probeValues']
        if kind=='rgb':
            lines+=['if class of probeValue is not list then error "BULK_RGB_NOT_LIST"','set probeRgbValues to {}','repeat with probeRgbOrdinal from 1 to count probeValue','set end of probeRgbValues to item probeRgbOrdinal of probeValue as integer','end repeat','set end of probeScalarValues to probeRgbValues']
        else:lines+=['set end of probeScalarValues to probeValue as '+kind]
        lines+=['end repeat','set probeRows to {{"vector",probeVectorClass,(count probeScalarValues) as integer,probeScalarValues}}']
        blocks.append(('bulk-'+label,lines))
    for position in (0,59,110):
        blocks.append(('font-record-'+str(position),[
            f'set probeRange to create range boundDoc start {position} end {position+1}',
            'set probeFontRecord to get properties of font object of probeRange',
            'set probeRecordClass to class of probeFontRecord as text',
            'set probeRows to {{"font-record",probeRecordClass,name of probeFontRecord as text,font size of probeFontRecord as real,bold of probeFontRecord as boolean,italic of probeFontRecord as boolean,underline of probeFontRecord as text}}']))
    try:
        for label,lines in blocks: read_block(label,lines)
    except DirtyRead as exc: report['blocker']='Native getter dirtied copy at '+str(exc)
    report['after']=state('final-state');report['checks']['body_unchanged']=report['after'][3:5]==report['baseline'][3:5]
    report['checks']['private_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    s.close();report['owned_closed']=s._closed
    report['inventory_final']=inventory(OUT,'final');report['checks']['inventory_empty']=report['inventory_final']==[]
    report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
    report['status']='DIAGNOSED' if all(report['checks'].values()) else 'FAIL'
except BaseException as exc:
    report['error']={'type':type(exc).__name__,'message':str(exc)};(OUT/'failure.txt').write_text(traceback.format_exc())
    if s and not s._closed:
        s._retain('Read-only snapshot probe uncertainty; parent reconciliation required');report['quarantined']=s._quarantined;report['remaining_owned_path']=s._bound_path
finally:
    if s and s.staging_root.exists():shutil.copytree(s.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
    flush()
print(json.dumps({'status':report['status'],'checks':report['checks'],'blocker':report.get('blocker'),'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False),flush=True)
raise SystemExit(0 if report['status']=='DIAGNOSED' else 1)
