"""Matrix02 preparation: separate constructor success from insertion placement.

Each owned private copy receives exactly one target make-table request. Only
native -2710 is caught as an ordinary error; all other errors/unknown ACKs stop.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
from uuid import uuid4

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession,_temporary_root
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_semantic_table import observe_native,retain_runtime
from fixtures.microsoft_parity.macos_word_quality import quarantine,record_failure,same
from fixtures.microsoft_parity.macos_word_quality_feasibility import bound_guard,sentinel_guard,exact,sha,source_pair_matches
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity import macos_word_inline_rule_feasibility as safe

CASES=(('A-reselect-document-location',4,3,19,True,'boundDoc'),
       ('B-reselect-range-location',4,3,19,True,'diagRange'))
SOURCES=[*sorted((ROOT/'skills/WPSComposer').rglob('*.py')),*sorted((ROOT/'fixtures/microsoft_parity').glob('*.py')),ROOT/'tests/msoffice/test_semantic_table_matrix02_preparation.py']


def commands(path,name,token,nrows,ncols,end,active,*,location="boundDoc"):
    if (nrows,ncols,end,active,location) not in [(r,c,e,a,l) for _,r,c,e,a,l in CASES]:raise ValueError('Closed matrix only')
    lines=bound_guard(path)+sentinel_guard(name,token)+[
        'activate object boundWindow',
        'set selection start of selection of boundWindow to 12',
        f'set selection end of selection of boundWindow to {end}',
        f'set diagActiveOwned to {exact("posix full name of document of active window as text",path)} as boolean',
        f'if diagActiveOwned is not {str(active).lower()} then error "DIAG_ACTIVE_OWNER_CHANGED"',
        'set diagSelection to selection of boundWindow',
        f'if not {exact("posix full name of document of diagSelection as text",path)} then error "DIAG_SELECTION_OWNER_CHANGED"',
        'if story type of diagSelection is not main text story then error "DIAG_STORY_CHANGED"',
        'set diagRange to text object of diagSelection',
        f'if start of content of diagRange is not 12 or end of content of diagRange is not {end} then error "DIAG_RANGE_CHANGED"',
        'if (count tables of boundDoc) is not 1 then error "DIAG_SEED_CHANGED"',
        'set diagPrefix to create range boundDoc start 0 end 12',
        'set diagFullMarker to create range boundDoc start 12 end 19',
        'if (content of diagFullMarker as text) is not "REPLACE" then error "DIAG_MARKER_CHANGED"',
        'set nativeRows to {{"before",posix full name of boundDoc as text,12,end of content of diagRange,count tables of boundDoc,'
        'content of text object of boundDoc as text,content of text object of bookmark "semantic_old_table" of boundDoc as text,'
        'content of diagPrefix as text,content of text object of bookmark "semantic_suffix" of boundDoc as text,'
        'saved of boundDoc,name of document of active window as text,diagActiveOwned}}',
        'set diagResult to "created"','set diagCode to 0','set diagMessage to ""',
        'try',
        f'set diagTable to make new table at {location} with properties {{text object:diagRange,number of rows:{nrows},number of columns:{ncols}}}',
        'on error diagCaughtMessage number diagCaughtCode',
        'if diagCaughtCode is not -2710 then error diagCaughtMessage number diagCaughtCode',
        'set diagResult to "ordinary-error"',
        'set diagCode to diagCaughtCode',
        'set diagMessage to diagCaughtMessage',
        'end try',
        *bound_guard(path),*sentinel_guard(name,token),
        'set end of nativeRows to {"attempt",diagResult,diagCode,diagMessage}',
        'set end of nativeRows to {"after",count tables of boundDoc,content of text object of boundDoc as text,saved of boundDoc}',
        'repeat with diagIndex from 1 to (count tables of boundDoc)',
        'set diagObserved to table diagIndex of boundDoc',
        'set end of nativeRows to {"table",diagIndex,start of content of text object of diagObserved,end of content of text object of diagObserved,'
        'number of rows of diagObserved,number of columns of diagObserved,content of text object of diagObserved as text}',
        'end repeat',
    ]
    return lines


def valid(rows,path,end,active):
    if not isinstance(rows,list) or len(rows)<4 or any(not isinstance(r,list) for r in rows):return False
    before,attempt,after=rows[:3]
    if (len(before)!=12 or before[:3]!=['before',path,12] or type(before[2]) is not int
            or type(before[3]) is not int or before[3]!=end or type(before[4]) is not int or before[4]!=1
            or not all(isinstance(v,str) for v in before[5:9]) or before[7]!='PREFIX 中文😀\r'
            or before[8]!='SUFFIX preserved 中文😀\r' or 'EXISTING TABLE 原样' not in before[6]
            or type(before[9]) is not bool or not isinstance(before[10],str) or before[11] is not active):return False
    if (len(attempt)!=4 or attempt[0]!='attempt' or type(attempt[2]) is not int or not isinstance(attempt[3],str)
            or (attempt[1],attempt[2]) not in (('created',0),('ordinary-error',-2710))):return False
    if (len(after)!=4 or after[0]!='after' or type(after[1]) is not int or after[1]<1
            or not isinstance(after[2],str) or type(after[3]) is not bool or len(rows)!=after[1]+3):return False
    for index,row in enumerate(rows[3:],1):
        if (len(row)!=7 or row[0]!='table' or any(type(v) is not int for v in row[1:6]) or row[1]!=index
                or not 0<=row[2]<row[3] or not row[4]>0 or not row[5]>0 or not isinstance(row[6],str)):return False
    return True


def assess_placement(native,nrows,ncols,end):
    """Interpret already validated observations without turning append into parity."""
    created=native[1][1:3]==['created',0]
    matches=[row for row in native[3:] if row[4:6]==[nrows,ncols]
             and 'EXISTING TABLE 原样' not in row[6]] if created else []
    bounds=matches[0][2:4] if len(matches)==1 else None
    return {'constructor_succeeded':created,'created_table_bounds':bounds,
            'requested_insertion_observed':bounds is not None and bounds[0]==12,
            'replacement_removed':None if end==12 else 'REPLACE' not in native[2][2]}


def run(output,source):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    source=Path(source).resolve(strict=True)
    report={'status':'RUNNING','scope':'constructor diagnostics only; successful creation is not insertion-location or replacement parity','cases':[],'checks':{},'steps':[],'confirmed':[],'source_input':str(source),'source_input_sha256':sha(source)}
    session=None
    def flush():
        pending=output/'report.pending';pending.write_text(json.dumps(report,ensure_ascii=False,indent=2));pending.replace(output/'report.json')
    def call(label,fn,owner=None):
        if report.get('quarantined') or report.get('native_uncertainty'):raise RuntimeError('Native followup refused')
        report['current_step']=label;flush()
        try:return fn()
        except BaseException as error:quarantine(output,report,label,error,owner);raise
    try:
        marker=_temporary_root()/'wpscomposer-native-word.lock.quarantine'
        if marker.exists():raise RuntimeError('Existing quarantine blocks matrix')
        report['source_hashes']=retain_sources(output,SOURCES)
        dictionary=Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef');shutil.copy2(dictionary,output/'localWord.sdef');report['dictionary_sha256']=sha(dictionary)
        report['inventory_before']=safe.independent_inventory(output,'before',report)
        if report['inventory_before']!=[]:raise RuntimeError('Expected initial empty inventory')
        for label,nrows,ncols,end,active,location in CASES:
            case={'label':label,'requested_dimensions':[nrows,ncols],'range':[12,end],'active_owned':active,'make_location':location};report['cases'].append(case);flush()
            case_dir=output/label;case_dir.mkdir()
            session=call(label+'-open',lambda:MacWordSession.open_document(source,read_only=False,visible=False));session._retain_evidence=True
            case['owned_path']=session._bound_path;flush()
            if 'sentinel_name' not in report:
                token='SEMANTIC MATRIX SENTINEL 中文😀 '+uuid4().hex;report['sentinel_token']=token
                observed=observe_native(session,output,report,'sentinel-created',[
                    'set qualitySentinel to make new document',f'set content of text object of qualitySentinel to {apple_string(token)}',
                    'set nativeRows to {{name of qualitySentinel as text,version as text}}'],
                    lambda r:isinstance(r,list) and len(r)==1 and isinstance(r[0],list) and len(r[0])==2 and all(isinstance(v,str) and v for v in r[0]))
                report['sentinel_name'],report['word_version']=observed[0]
                report['sentinel_before']=sentinel_preimage(safe.independent_inventory(output,'sentinel-before',report,session),report['sentinel_name'],token);flush()
            native=observe_native(session,output,report,label+'-one-constructor',commands(session._bound_path,report['sentinel_name'],report['sentinel_token'],nrows,ncols,end,active,location=location),
                lambda r:valid(r,session._bound_path,end,active),mutate=True)
            case['native_rows']=native;case['outcome']=native[1][1];case['error_code']=native[1][2]
            case['body_unchanged']=native[0][5]==native[2][2];case['table_count_delta']=native[2][1]-native[0][4]
            case.update(assess_placement(native,nrows,ncols,end));flush()
            call(label+'-save',lambda:session.save_docx(case_dir/'after.docx'),session)
            safe.retain_xml(case_dir,'after')
            case['saved_docx_sha256']=sha(case_dir/'after.docx')
            call(label+'-close',session.close,session)
            if not session._closed or session._quarantined:raise RuntimeError('Owned close unverified')
            observed=safe.independent_inventory(output,label+'-after-close',report,session)
            if observed!=[report['sentinel_before']]:raise RuntimeError('Owned inventory or sentinel changed')
            case['owned_closed_with_sentinel_preserved']=True
            retain_runtime(session,case_dir,report,label+'-runtime');flush()
        safe.close_sentinel_after_owned(session,output,report,report['sentinel_name'],report['sentinel_token'])
        report['inventory_final']=safe.independent_inventory(output,'final',report,session)
        report['checks']['inventory_restored']=report['inventory_final']==report['inventory_before']
        report['checks']['source_unchanged']=sha(source)==report['source_input_sha256']
        report['checks']['sources_unchanged']=all(source_pair_matches(ROOT/p,output/'source'/p,value) for p,value in report['source_hashes'].items())
        report['checks']['dictionary_unchanged']=sha(dictionary)==report['dictionary_sha256']
        if all(report['checks'].values()) and len(report['cases'])==2 and report.get('sentinel_closed'):
            report['status']='DIAGNOSTIC_COMPLETE'
    except BaseException as error:
        record_failure(output,report,error)
        if session and not session._closed:quarantine(output,report,report.get('current_step','matrix'),error,session)
    finally:
        if report['status']!='DIAGNOSTIC_COMPLETE':
            retain_runtime(session,output,report,'failed-runtime')
        report['remaining_sentinel']=None if report.get('sentinel_closed') else report.get('sentinel_name')
        report['artifacts']={str(p.relative_to(output)):sha(p) for p in output.rglob('*') if p.is_file() and 'source' not in p.relative_to(output).parts and p.name not in ('report.json','report.pending')}
        flush()
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',required=True);parser.add_argument('--source',required=True);args=parser.parse_args()
    if not args.execute:parser.error('--execute and root Word lease required')
    result=run(args.output,args.source)
    print(json.dumps({'status':result['status'],'cases':[(c['label'],c.get('outcome')) for c in result['cases']],'remaining_sentinel':result['remaining_sentinel']},ensure_ascii=False))
    raise SystemExit(0 if result['status']=='DIAGNOSTIC_COMPLETE' else 1)
