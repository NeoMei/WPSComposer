"""Source-bound real controller recovery acceptance, task-owned Word only."""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import traceback
from uuid import uuid4

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document,open_document
from skills.WPSComposer.scripts.generation_plan import GenerationOperation
from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands,state_commands
from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.writer import NativeWriterObjectError
from fixtures.microsoft_parity.macos_word_fields import retain_sources

SOURCES=[Path(__file__),*[ROOT/'skills/WPSComposer/scripts'/p for p in (
    'msoffice/macos_word_recovery.py','msoffice/macos_word_session.py','msoffice/macos_word_fields.py',
    'msoffice/macos_word_references.py','msoffice/macos_runtime.py','msoffice/macos_script.py',
    'writer.py','longform/windows_executor.py','longform/recovery.py') if (ROOT/'skills/WPSComposer/scripts'/p).exists()]]
BOOKMARK='wpsc_fig_'+'a'*24


class Journal:
    def __init__(self,output):
        self.output=output
        self.data={'passed':False,'source_hashes':retain_sources(output,SOURCES),'steps':[],'checks':{}}
        self.flush()
    def flush(self):
        path=self.output/'report.pending';path.write_text(json.dumps(self.data,ensure_ascii=False,indent=2));path.replace(self.output/'report.json')
    def add(self,label,rows):
        self.data['steps'].append({'label':label,'rows':rows});self.flush();return rows
    def check(self,name):self.data['checks'][name]=True;self.flush()


def inventory(output,label):
    lines=['set nativeRows to {}','if application "Microsoft Word" is running then',
           'tell application "Microsoft Word"','repeat with di from 1 to count documents',
           'set inventoryDoc to document di',*hash_commands('content of text object of inventoryDoc as text','inventoryHash'),
           'set end of nativeRows to {name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,inventoryHash}',
           'end repeat','end tell','end if','return my jsonRows(nativeRows)']
    path=output/(label+'.applescript');path.write_text(_JSON+'\n'.join(lines))
    result=subprocess.run(['/usr/bin/osascript',str(path)],capture_output=True,text=True,timeout=30)
    path.with_suffix('.log').write_text(result.stdout+'\n'+result.stderr)
    if result.returncode:raise RuntimeError('inventory failed')
    return sorted(json.loads(result.stdout))


def close_sentinel(output,name,token):
    """Close the exact synthetic sentinel only after the owned context exited."""
    lines=['if not application "Microsoft Word" is running then error "SENTINEL_UNAVAILABLE"',
           'tell application "Microsoft Word"',
           f'set sentinelDocument to document {apple_string(name)}',
           f'if (name of sentinelDocument as text) is not {apple_string(name)} then error "SENTINEL_IDENTITY_CHANGED"',
           'if (path of sentinelDocument as text) is not "" then error "SENTINEL_SAVED"',
           f'if (content of text object of sentinelDocument as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"',
           'if saved of sentinelDocument then error "SENTINEL_SAVED"',
           'close sentinelDocument saving no',
           f'set nativeRows to {{{{"sentinel-closed",{apple_string(name)}}}}}',
           'end tell','return my jsonRows(nativeRows)']
    source=output/'independent-sentinel-close.applescript';source.write_text(_JSON+'\n'.join(lines))
    result=subprocess.run(['/usr/bin/osascript',str(source)],capture_output=True,text=True,timeout=30)
    source.with_suffix('.log').write_text(result.stdout+'\n'+result.stderr)
    if result.returncode:raise RuntimeError('Independent sentinel close failed')
    rows=json.loads(result.stdout)
    if rows!=[['sentinel-closed',name]]:raise RuntimeError('Independent sentinel close acknowledgement invalid')
    return rows


def main(output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    journal=Journal(output);session=None;sentinel=None;sentinel_token=None
    try:
        starting=journal.add('starting-inventory',inventory(output,'starting-inventory'))
        with create_document('writer',engine='msoffice',visible=False) as session:
            try:
                sentinel_token='Recovery sentinel 中文😀 '+uuid4().hex
                sentinel=journal.add('sentinel-created',session._execute(['set d to make new document',f'set content of text object of d to {apple_string(sentinel_token)}','set nativeRows to {{name of d as text}}']))[0][0]
                sentinel_before=next(r for r in inventory(output,'before-owned-write') if r[0]==sentinel)
                journal.data['sentinel_before']=sentinel_before;journal.flush()
                journal.add('word-version',session._execute(['set nativeRows to {{version as text}}']))
                # Empty native document; plain integer/no-op/repeated coercion.
                token=session.degradation_checkpoint();assert type(token) is int and token==0
                session.rollback_degradation_checkpoint('0');session.rollback_degradation_checkpoint(0)
                journal.check('empty_and_repeated_noop')
                session.add_paragraph('Recovery prefix 中文😀')
                session.add_heading_level('Chapter',1)
                old_index=session.insert_toc('Contents')
                journal.add('seed-target',session._execute(['set r to create range boundDoc start 0 end 8',f'make new bookmark at boundDoc with properties {{name:{apple_string(BOOKMARK)},text object:r}}','set nativeRows to {{"seed-ack"}}']))
                runs=[{'type':'text','text':'Existing '},{'type':'reference','bookmarkName':BOOKMARK,'prefix':'','suffix':'','fallbackText':'EXISTING'}]
                session.add_cross_reference_paragraph(runs=runs,owner_node_id='old-ref')
                old_ref=session._tracked_references[0][0]
                original_handles=(tuple(session._tracked_indexes),tuple(session._tracked_references))
                before_fields=session.snapshot_fields()
                # Mixed appended topology: real REF in a new table and another
                # outside, then an acknowledged partial native operation failure.
                original_reference=session.add_cross_reference_paragraph
                original_fallback=session.add_cross_reference_fallback
                fallback_calls=[];partial_calls=[]
                def partial_reference(**kwargs):
                    partial_calls.append(1)
                    original_reference(runs=[{'type':'text','text':'PARTIAL-TRACKED '},{'type':'reference','bookmarkName':BOOKMARK,'prefix':'','suffix':'','fallbackText':'PARTIAL'}],owner_node_id='partial-tracked')
                    lines=['activate object boundWindow','set p to (end of content of text object of boundDoc) - 1',
                           'set r to create range boundDoc start p end p',
                           'set ownTable to make new table at boundDoc with properties {text object:r,number of rows:1,number of columns:1}',
                           'set p to start of content of text object of cell 1 of row 1 of ownTable',
                           'set r to create range boundDoc start p end p',
                           f'create new field text range r field type field ref field text {apple_string(BOOKMARK)} preserve formatting true',
                           'set p to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start p end p',
                           'set content of r to "PARTIAL-OUTSIDE "','set p to (end of content of text object of boundDoc) - 1',
                           'set r to create range boundDoc start p end p',
                           f'create new field text range r field type field ref field text {apple_string(BOOKMARK)} preserve formatting true',
                           'set nativeRows to {{"partial-native-ack",count fields of boundDoc,count tables of boundDoc}}']
                    journal.add('partial-native-ack',session._execute_structural(lines))
                    raise NativeWriterObjectError('CROSS_REFERENCE_FAILED','controlled partial failure')
                def fallback(**kwargs):
                    fallback_calls.append(1);return original_fallback(**kwargs)
                session.add_cross_reference_paragraph=partial_reference
                session.add_cross_reference_fallback=fallback
                executor=WindowsLongformExecutor()
                operation=GenerationOperation(op='writer.add_cross_reference',args={'runs':[{'type':'text','text':'FALLBACK-ONCE'}]},node_id='recovery-partial',failure_policy={'mode':'degrade','recoverableCodes':['CROSS_REFERENCE_FAILED'],'fallback':'inline-fallback'})
                try:executor._run_op(session,operation)
                finally:
                    session.add_cross_reference_paragraph=original_reference;session.add_cross_reference_fallback=original_fallback
                assert partial_calls==[1] and fallback_calls==[1] and len(executor._issues)==1
                assert tuple(session._tracked_indexes)==original_handles[0] and tuple(session._tracked_references)==original_handles[1]
                assert session._tracked_indexes[0][0] is old_index and session._tracked_references[0][0] is old_ref
                journal.data['issues']=[asdict(item) for item in executor._issues]
                journal.check('actual_controller_checkpoint_partial_native_rollback_fallback_once_issue_once')
                journal.check('original_ref_index_handle_instances_preserved')
                after_fields=session.snapshot_fields()
                assert [x.stable_key for x in before_fields]==[x.stable_key for x in after_fields]
                journal.check('field_snapshot_after_rollback_and_fallback')
                body=journal.add('body',session._execute(['set nativeRows to {{content of text object of boundDoc as text,count tables of boundDoc}}']))
                assert body[0][0].count('FALLBACK-ONCE')==1 and 'PARTIAL-' not in body[0][0] and body[0][1]==0
                token=session.degradation_checkpoint()
                for invalid in (-1,token+100):
                    try:session.rollback_degradation_checkpoint(invalid)
                    except NativeWriterObjectError as exc:assert exc.code=='LOCAL_MUTATION_ROLLBACK_FAILED'
                    else:raise AssertionError('invalid coordinate accepted')
                assert session._execute(['set nativeRows to {{content of text object of boundDoc as text,count tables of boundDoc}}'])==body
                session.rollback_degradation_checkpoint(token);session.rollback_degradation_checkpoint(str(token))
                journal.check('invalid_bounds_no_mutation_and_repeat')
                session.save_docx(output/'recovery.docx');session.export_pdf(output/'recovery.pdf')
                journal.check('native_docx_save_and_pdf_export')
                assert next(r for r in inventory(output,'after-save') if r[0]==sentinel)==sentinel_before
                journal.check('unsaved_sentinel_exact_hash_unchanged')
            finally:
                if session.staging_root.exists():shutil.copytree(session.staging_root,output/'native-runtime',dirs_exist_ok=True)
        after_owned=journal.add('after-owned-close',inventory(output,'after-owned-close'))
        assert next(r for r in after_owned if r[0]==sentinel)==sentinel_before
        assert [r for r in after_owned if r[0]!=sentinel]==starting
        journal.check('owned_context_closed_with_unsaved_sentinel_preserved')
        journal.add('sentinel-close',close_sentinel(output,sentinel,sentinel_token));sentinel=None
        assert journal.add('after-sentinel-close',inventory(output,'after-sentinel-close'))==starting
        with create_document('writer',engine='msoffice',visible=False) as session:
            try:
                session.add_paragraph('Shape guard prefix')
                session.add_image(ROOT/'tests/longform_m3/fixtures/media/oriented.jpg',width=20,height=20,inline=True)
                session.add_floating_textbox('Existing box',left=20,top=20,width=100,height=40)
                token=session.degradation_checkpoint()
                session.rollback_degradation_checkpoint(token)
                journal.check('unchanged_preexisting_textbox_and_inline_picture_exact_noop')
                session.add_floating_textbox('New box',left=20,top=80,width=100,height=40)
                try:session.rollback_degradation_checkpoint(token)
                except NativeWriterObjectError as exc:assert exc.code=='LOCAL_MUTATION_ROLLBACK_FAILED'
                else:raise AssertionError('new floating object accepted')
                assert not session._quarantined
                native=journal.add('shape-guard-after-rejection',session._execute(['set nativeRows to {{count shapes of boundDoc}}']))
                assert native==[[2]]
                journal.check('new_textbox_blocks_rollback_before_deletion')
            finally:
                if session.staging_root.exists():shutil.copytree(session.staging_root,output/'shape-guard-runtime',dirs_exist_ok=True)
        assert journal.add('after-shape-guard-discard',inventory(output,'after-shape-guard-discard'))==starting
        digest=hashlib.sha256((output/'recovery.docx').read_bytes()).hexdigest()
        with open_document(output/'recovery.docx',engine='msoffice',read_only=False,visible=False) as session:
            try:
                observed=journal.add('reopen-body',session._execute(['set nativeRows to {{content of text object of boundDoc as text,count tables of boundDoc}}']))
                assert observed==body
                token=session.degradation_checkpoint();session.add_paragraph('REOPEN-EDIT');session.rollback_degradation_checkpoint(token)
                assert session._execute(['set nativeRows to {{content of text object of boundDoc as text,count tables of boundDoc}}'])==body
                journal.check('native_reopen_edit_and_actual_rollback')
            finally:
                if session.staging_root.exists():shutil.copytree(session.staging_root,output/'reopen-runtime',dirs_exist_ok=True)
        assert hashlib.sha256((output/'recovery.docx').read_bytes()).hexdigest()==digest
        assert journal.add('after-reopen-close',inventory(output,'after-reopen-close'))==starting
        journal.check('source_unchanged_and_owned_documents_closed')
        import pdfplumber
        with pdfplumber.open(output/'recovery.pdf') as pdf:
            text='\n'.join(page.extract_text() or '' for page in pdf.pages)
            assert text.count('FALLBACK-ONCE')==1 and 'PARTIAL-' not in text
            journal.data['pdf_text']=text
        journal.check('pdf_content_verified')
        journal.data['passed']=True
    except BaseException as error:
        journal.data['error']={'type':type(error).__name__,'message':str(error),'quarantined':getattr(session,'_quarantined',False)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
        if sentinel:journal.data['retained_sentinel']=sentinel
    journal.flush();return journal.data


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',required=True);args=parser.parse_args()
    if not args.execute:parser.error('--execute required')
    report=main(args.output);print(json.dumps({'passed':report['passed'],'error':report.get('error')}));sys.exit(0 if report['passed'] else 1)
