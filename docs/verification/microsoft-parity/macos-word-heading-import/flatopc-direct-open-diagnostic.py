"""One direct-file open diagnostic, no SaveAs/import/repair/retry.

Same immutable Flat OPC input and converter/read-only/confirmation flags as
flatopc-open-01. Only file argument dispatch changes from optional file-name
text to the Standard Suite's direct POSIX file argument. Raw return kind and
inventory are persisted before Python binding gates, including zero-doc result.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession,apple_string,sha,inventory,retain_sources,SOURCES,
)
INPUT=Path(__file__).parent/'run-01/fragment.xml'
INPUT_SHA='56d2c52396c7ce74bbdcb201a31d7cb18f17e62b6edee77b3a59fe0125dce4b4'


def commands(path):
    return ['if (count documents) is not 0 then error "WPSC_DIRECT_OPEN_REQUIRES_EMPTY"',
            f'set openedValue to open (POSIX file {apple_string(str(path))}) file converter open format xmldocument serialized read only true add to recent files false confirm conversions false',
            'set resultKind to "missing value"','set resultCount to 0',
            'if openedValue is not missing value then',
            'set resultKind to (class of openedValue) as text',
            'set resultCount to 1',
            'if class of openedValue is list then set resultCount to count openedValue','end if',
            'set nativeRows to {{"open-result",resultKind,resultCount}}',
            'repeat with nativeDocument in documents',
            'set end of nativeRows to {"document",name of nativeDocument as text,posix full name of nativeDocument as text,saved of nativeDocument,id of active window of nativeDocument}',
            'end repeat']


def observation_valid(rows):
    if (not isinstance(rows,list) or not rows or not isinstance(rows[0],list)
            or len(rows[0])!=3 or rows[0][0]!='open-result'
            or not isinstance(rows[0][1],str) or not rows[0][1]
            or type(rows[0][2]) is not int or rows[0][2]<0):return False
    return all(isinstance(r,list) and len(r)==5 and r[0]=='document'
               and isinstance(r[1],str) and bool(r[1]) and isinstance(r[2],str)
               and type(r[3]) is bool and type(r[4]) is int and r[4]>=0 for r in rows[1:])


def owned_binding(rows,path):
    if not observation_valid(rows) or len(rows)!=2:return None
    record=rows[1]
    if record[1:3]!=[path.name,str(path)]:return None
    return [['binding',record[4],record[2],record[1]]]


def run(output):
    report={'status':'FAIL','scope':'Diagnostic only; never format/style acceptance',
            'checks':{},'journal':[],
            'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_flatopc_direct_open_diagnostic.py'),*SOURCES])}
    owner=None;uncertain=False
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('Direct open diagnostic failed: '+label)
    try:
        check('immutable_input',sha(INPUT)==INPUT_SHA)
        report['inventory_before']=inventory(output,'before')
        check('requires_empty',report['inventory_before']==[])
        owner=MacWordSession();owner._prepare();owner._retain_evidence=True
        path=owner.staging_root/('direct-open-'+uuid4().hex+'.xml')
        path.write_bytes(INPUT.read_bytes());report['input_path']=str(path)
        report['dictionary_sha256']=sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['journal'].append({'operation':'direct_file_open','path':str(path),'before':[],'status':'submitted'})
        report['current_phase']='direct_file_open';flush();uncertain=True
        rows=owner._execute(commands(path),bind=False)
        # Persist complete raw observation before schema/binding acceptance.
        report['raw_open_observation']=rows;flush()
        check('observation_schema',observation_valid(rows))
        report['inventory_after_open']=inventory(output,'after-open');flush()
        if len(rows)==1:
            check('zero_document_crosscheck',report['inventory_after_open']==[])
            report['outcome']='NO_DOCUMENTS_CREATED'
            uncertain=False
        else:
            binding=owned_binding(rows,path)
            check('one_exact_owned_path',binding is not None)
            observed=report['inventory_after_open']
            check('independent_owned_inventory',len(observed)==1 and observed[0][:2]==[path.name,str(path)] and observed[0][2] is rows[1][3])
            owner._private_path=path;owner._owns_doc=True;owner._bind(binding)
            uncertain=False;report['outcome']='EXACT_PRIVATE_DOCUMENT_CREATED'
        report['journal'][-1]['status']='observed';flush()
        report['journal'].append({'operation':'close_owned_if_bound','path':str(path) if owner._bound_path else None,'status':'submitted'});flush()
        owner.close()
        check('owner_closed',owner._closed and not owner._quarantined)
        check('empty_after_close',inventory(output,'after-close')==[])
        report['journal'][-1]['status']='acknowledged'
        check('private_input_preserved',sha(path)==INPUT_SHA)
        check('original_input_preserved',sha(INPUT)==INPUT_SHA)
        check('sources_unchanged',all(sha(ROOT/name)==digest for name,digest in report['source_hashes'].items()))
        report['status']='DIAGNOSTIC_COMPLETE'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and not owner._closed:
            if uncertain:owner._retain('Direct-file open completion or identity unverified; inspect raw observation before cleanup')
            try:owner.close()
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        if owner and owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/'native-runtime',dirs_exist_ok=True)
        if report.get('cleanup_failure'):report['status']='FAIL'
        report['identity_uncertain']=uncertain;flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Direct-file diagnostic requires explicit --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='DIAGNOSTIC_COMPLETE' else 1


if __name__=='__main__':raise SystemExit(main())
