"""Leased source-bound read-only snapshot diagnostics; no document writes."""
from pathlib import Path
import hashlib,json,shutil,sys,traceback
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
report={'status':'FAIL','scope':'read-only fresh owned copy; capture ordinary getter errors and continue independent read blocks','steps':[],'checks':{},'source_hashes':{str(p.relative_to(ROOT)):sha(p) for p in SOURCES}}
s=None

def flush():(OUT/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))

def traced(lines):
    out=[]
    for i,line in enumerate(lines):
        if not line.startswith(('end ', 'else', 'on error', 'repeat ')):
            out.append('set probePhase to '+apple_string(str(i)+': '+line[:220]))
        out.append(line)
    return out

def read_block(label,lines,kind='state'):
    report['current_phase']=label;flush()
    commands=['set qualityPoint to 111','set qualityDelta to 0','set qualityState to {}','set qualityLayout to {}',
              'set qualityPrefixHash to ""','set qualitySuffixHash to ""','set probePhase to "start"','try',*traced(lines)]
    if kind=='state':commands+=['set nativeRows to {{"read-ok",qualityState}}']
    elif kind=='layout':commands+=hash_commands('my jsonRows(qualityLayout)','probeHash')+['set nativeRows to {{"read-ok",count qualityLayout,probeHash}}']
    else:commands+=['set nativeRows to {{"read-ok",probeRows}}']
    commands+=['on error probeError number probeNumber',
               'if probeNumber is -1712 or probeNumber is -609 or probeNumber is -128 then error probeError number probeNumber',
               # This contains only source/probe phases. Binding guard lives
               # outside try; owned identity errors can never be normalized.
               'set nativeRows to {{"ordinary-read-error",probePhase,probeNumber,probeError as text}}','end try']
    rows=s._execute(commands);report['steps'].append({'label':label,'rows':rows});flush();return rows

try:
    report['inventory_before']=inventory(OUT,'before');flush()
    if report['inventory_before']!=[]:raise RuntimeError('Nonempty inventory; no document opened')
    s=MacWordSession.open_document(SOURCE,read_only=True,visible=False);s._retain_evidence=True
    original_execute=s._execute
    s._execute=lambda lines,**kw:original_execute(q._guard(s)+list(lines),**kw)
    report['owned_path']=s._bound_path;report['runtime']=str(s.staging_root);flush()
    rows=s._execute(hash_commands('content of text object of boundDoc as text','probeBodyHash')+['set nativeRows to {{"before",end of content of text object of boundDoc,probeBodyHash,read only of boundDoc,version as text}}'])
    report['before_body']=rows;flush()
    if not (len(rows)==1 and rows[0][:2]==['before',220] and rows[0][3] is True):raise RuntimeError('Readonly baseline mismatch')
    prefix=['set probeRows to {}','set qsection to section 1 of boundDoc','set qheader to get header qsection index header footer primary',
            'set qfooter to get footer qsection index header footer primary','repeat with qpart in {qheader,qfooter}']
    original=['if count shapes of qpart is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"']
    parens=['if (count shapes of qpart) is not 0 then error "WPSC_QUALITY_DRAWING_UNVERIFIED"']
    atomic=['set probeCount to (count shapes of qpart) as integer','set probeReject to (probeCount is not 0)',
            'set end of probeRows to {"part",is header of qpart,probeCount,probeReject}']
    for label,part in [('count-original',original),('count-parenthesized',parens),('count-atomic',atomic)]:
        read_block(label,prefix+part+['end repeat'],'rows')
    all_lines=q._snapshot_commands()
    corrected=[line.replace('if count shapes of qpart is not 0','if (count shapes of qpart) is not 0') for line in all_lines]
    read_block('complete-snapshot-parenthesized-count',corrected)
    layout=q._layout_commands();layout=[line.replace('if count shapes of qpart is not 0','if (count shapes of qpart) is not 0') for line in layout]
    li=layout.index('repeat with qi from 1 to count list templates of boundDoc')
    st=layout.index('repeat with qstyle in Word styles of boundDoc')
    tail=layout.index("set recoveryTask to current application's NSTask's alloc()'s init()")
    for label,block in [('section-and-page-part',layout[:li]),('list-definitions',layout[li:st]),('styles',layout[st:tail])]:
        read_block(label,block,'layout')
    markers=['repeat with qi from 1 to count paragraphs of boundDoc','repeat with qi from 1 to count fields of boundDoc',
             'set qualityOldOrdinal to 0','repeat with qi from 1 to count bookmarks of boundDoc']
    indexes=[corrected.index(v) for v in markers]+[len(corrected)]
    for i,label in enumerate(('paragraph-and-character','main-fields','tables-and-row-flags','bookmarks')):
        read_block(label,corrected[indexes[i]:indexes[i+1]])
    rows=s._execute(hash_commands('content of text object of boundDoc as text','probeBodyHash')+['set nativeRows to {{"after",end of content of text object of boundDoc,probeBodyHash,read only of boundDoc}}'])
    report['after_body']=rows
    report['checks']['body_unchanged']=rows==[['after',220,report['before_body'][0][2],True]]
    if not report['checks']['body_unchanged']:raise RuntimeError('Unexpected native body change')
    report['checks']['private_file_unchanged']=sha(Path(s._bound_path))==sha(SOURCE)
    if not report['checks']['private_file_unchanged']:raise RuntimeError('Private file changed')
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
print(json.dumps({'status':report['status'],'steps':report['steps'],'checks':report['checks'],'error':report.get('error'),'remaining_owned_path':report.get('remaining_owned_path')},ensure_ascii=False))
raise SystemExit(0 if report['status']=='DIAGNOSED' else 1)
