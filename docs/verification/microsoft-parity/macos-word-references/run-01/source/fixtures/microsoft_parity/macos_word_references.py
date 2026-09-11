"""Guarded source-bound direct Mac Word reference probe; --execute is required.

The real controller checkpoint chain is a separate dependency, never simulated.
All package access in this fixture is read-only verification, never repair.
"""
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

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.writer import NativeWriterObjectError
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_sections import _inventory_commands

BOOKMARKS = ['wpsc_' + kind + '_' + char * 24 for kind,char in [('fig','a'),('tab','b'),('eq','c')]]
SOURCES = [Path(__file__), ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
           ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_references.py',
           ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_fields.py',
           ROOT/'skills/WPSComposer/scripts/writer.py',
           ROOT/'skills/WPSComposer/scripts/artifact_transport.py',
           ROOT/'skills/WPSComposer/scripts/msoffice/macos_runtime.py',
           ROOT/'skills/WPSComposer/scripts/msoffice/macos_script.py',
           ROOT/'skills/WPSComposer/scripts/msoffice/errors.py',
           ROOT/'skills/WPSComposer/scripts/longform/windows_executor.py',
           ROOT/'fixtures/microsoft_parity/macos_word_fields.py',
           ROOT/'fixtures/microsoft_parity/macos_word_sections.py']


def inventory_without_launch(output, label):
    """Read every native document identity/state hash, even after owned close."""
    from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
    lines=['set nativeRows to {}',
           'if application "Microsoft Word" is running then',
           'tell application "Microsoft Word"',
           'repeat with di from 1 to count documents',
           'set inventoryDoc to document di',
           'set documentText to content of text object of inventoryDoc as text',
           'set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")',
           'set end of nativeRows to {name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}',
           'end repeat', 'end tell', 'end if', 'return my jsonRows(nativeRows)']
    source=Path(output)/(label+'.applescript');source.write_text(_JSON+'\n'.join(lines))
    result=subprocess.run(['/usr/bin/osascript',str(source)],capture_output=True,text=True,timeout=30)
    source.with_suffix('.log').write_text(result.stdout+'\n'+result.stderr)
    if result.returncode:raise RuntimeError('Native inventory failed')
    rows=json.loads(result.stdout)
    assert isinstance(rows,list) and all(isinstance(r,list) and len(r)==4 and type(r[2]) is bool for r in rows)
    return sorted(rows)


def seeded_targets(session):
    lines=[]
    for bookmark,sequence,number in zip(BOOKMARKS,['WPSC_FIG','WPSC_TAB','WPSC_EQ'],[7,11,13]):
        lines += session._position('end') + session._paragraph_boundary() + [
            'set r to create range boundDoc start insertionPoint end insertionPoint',
            f'create new field text range r field type field sequence field text {apple_string(sequence + " " + chr(92) + "r " + str(number))} preserve formatting true',
            'set ownField to field (count fields of boundDoc) of boundDoc',
            'if (update field ownField) is false then error "TARGET_UPDATE_FAILED"',
            f'make new bookmark at boundDoc with properties {{name:{apple_string(bookmark)},text object:result range of ownField}}',
            'set p to (end of content of text object of boundDoc) - 1',
            'set r to create range boundDoc start p end p', 'set content of r to return']
    assert session._execute_structural(lines+['set nativeRows to {{"ok"}}']) == [['ok']]


def readback(session):
    return session._execute(['set nativeRows to {}',
        'repeat with pi from 1 to count paragraphs of boundDoc',
        'set r to text object of paragraph pi of boundDoc',
        'set end of nativeRows to {"paragraph",pi,content of r as text,paragraph format left indent of paragraph format of r,first line indent of paragraph format of r,space before of paragraph format of r,space after of paragraph format of r,keep together of paragraph format of r,start of content of r,end of content of r}',
        'end repeat',
        'repeat with fi from 1 to count fields of boundDoc',
        'set f to field fi of boundDoc',
        'if field type of f is field ref then set end of nativeRows to {"reference",content of field code of f as text,content of result range of f as text}',
        'end repeat'])


def inspect_rows(rows):
    paragraphs=[r for r in rows if r[0]=='paragraph']
    refs=[r for r in rows if r[0]=='reference']
    assert len(refs)==4
    assert [r[1].strip().removesuffix(' \\* MERGEFORMAT') for r in refs] == ['REF '+b+' \\h' for b in BOOKMARKS] + ['REF '+BOOKMARKS[0]+' \\h']
    assert [r[2] for r in refs]==['8','11','13','8']
    by_text={r[2]:r for r in paragraphs}
    assert by_text['[3] 中文😀 已引\r'][3:8]==[21.5,-12.25,0,4.5,True]
    assert by_text['[9] 未引\r'][3:8]==[21.5,-12.25,0,4.5,True]
    assert '[8] 原样条目\r' in by_text and '7\r' in by_text and 'False\r' in by_text and 'None\r' in by_text
    bullet=next(r for r in paragraphs if r[2].startswith('•\t'))
    assert bullet[3:7]==[24,-24,0,3]
    literal=''.join(r[2] for r in paragraphs)
    assert '前中文😀见8尾；11；(13)后\r' in literal
    assert literal.count('FALLBACK-ONLY')==1
    assert literal.count('LOCAL-FALLBACK')==1 and 'PARTIAL-NATIVE-FAILURE' not in literal
    assert literal.count('[MISS 缺失]')==2


def inspect_native_styles(session, rows):
    """Read exact ranges, including both list kinds and each notice occurrence."""
    commands=['set nativeRows to {}'];expected=[]
    for row in rows:
        if row[0]!='paragraph':continue
        text=row[2];start=row[8];end=row[9]
        if text.startswith('•\t') or text.startswith('1.\t'):
            indent=24 if text.startswith('•\t') else 31.5
            commands += [f'set r to create range boundDoc start {start} end {end}',
                         'set tabVerified to false',
                         'repeat with ownTab in tab stops of paragraph 1 of r',
                         f'if tab stop position of ownTab is {indent} then set tabVerified to true', 'end repeat',
                         'set end of nativeRows to {"list",paragraph format left indent of paragraph format of r,first line indent of paragraph format of r,(line spacing rule of paragraph format of r is line space1 pt5),space before of paragraph format of r,space after of paragraph format of r,tabVerified}']
            expected.append(['list',indent,-indent,True,0,3,True])
        cursor=0
        while True:
            position=text.find('[MISS 缺失]',cursor)
            if position<0:break
            notice_start=start+len(text[:position].encode('utf-16-le'))//2
            notice_end=notice_start+9
            commands += [f'set r to create range boundDoc start {notice_start} end {notice_end}',
                         'set end of nativeRows to {"degradation",start of content of r,end of content of r,content of r as text,italic of font object of r,(color of font object of r is {40092, 0, 1542}),(background pattern color of shading of r is {64764, 59624, 59110})}']
            expected.append(['degradation',notice_start,notice_end,'[MISS 缺失]',True,True,True])
            cursor=position+9
    assert len([r for r in expected if r[0]=='list'])==2
    assert len([r for r in expected if r[0]=='degradation'])==2
    observed=session._execute(commands)
    assert observed==expected
    return observed


def run(output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'passed':False,'source_hashes':retain_sources(output,SOURCES),'checks':{},
            'dependencies':{'real_controller_checkpoint_recovery':'NOT_RUN: checkpoint/rollback methods outside this five-method slice'}}
    session=None;sentinel=None
    try:
        starting_inventory=inventory_without_launch(output,'starting-inventory')
        report['starting_inventory']=starting_inventory
        with create_document('writer',engine='msoffice',visible=False) as session:
            token='References sentinel 中文😀 '+uuid4().hex
            sentinel=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
            before=session._execute(_inventory_commands());report['inventory_before']=before
            report['word_version']=session._execute(['set nativeRows to {{version as text}}'])[0][0]
            try:
                session.add_paragraph('References direct native acceptance')
                seeded_targets(session)
                session.add_bibliography_native(entries=[dict(number=3,text='中文😀 已引'),dict(number=9,text='未引')],style='numeric',hangingIndentPt=12.25,leftIndentPt=21.5,spaceAfterPt=4.5)
                session.add_bibliography_legacy(entries=['[8] 原样条目',7,False,None])
                runs=[{'type':'text','text':'前中文😀'},
                      dict(type='reference',bookmarkName=BOOKMARKS[0],prefix='见',suffix='尾；',fallbackText='7'),
                      dict(type='reference',bookmarkName=BOOKMARKS[1],prefix='',suffix='；',fallbackText='11'),
                      dict(type='reference',bookmarkName=BOOKMARKS[2],prefix='(',suffix=')后',fallbackText='13')]
                assert session.add_cross_reference_paragraph(runs=runs,owner_node_id='references:mixed')=={'issues':[]}
                assert session.add_cross_reference_paragraph(runs=[dict(type='text',text='•\t'),dict(runs[1],prefix='',suffix='')],listFormatting={'indentPt':24},owner_node_id='references:bullet')=={'issues':[]}
                degraded=[dict(type='degradation',fallbackText='[MISS 缺失]',code='MISS',nodeId=node) for node in ('missing:1','missing:2')]
                result=session.add_citation_paragraph(runs=[dict(type='text',text='1.\t'),dict(type='citation',fallbackText='[3]'),*degraded],listFormatting={'indentPt':31.5})
                assert [r['nodeId'] for r in result['issues']]==['missing:1','missing:2']
                report['occurrence_issues']=result['issues']
                assert session.add_cross_reference_fallback(runs=[dict(type='text',text='Static '),dict(type='reference',bookmarkName=BOOKMARKS[0],prefix='(',fallbackText='FALLBACK-ONLY',suffix=')')]) is None
                # Real native partial append plus injected creation failure; restore
                # the original transport immediately. This tests local recovery only.
                execute=session._execute
                def forced_failure(lines,**kwargs):
                    replacement=[]
                    for line in lines:
                        if 'create new field text range ownRange field type field ref' in line:
                            replacement += ['set content of ownRange to "PARTIAL-NATIVE-FAILURE"','error "CONTROLLED_REFERENCE_CREATION_FAILURE"']
                        else:replacement.append(line)
                    return execute(replacement,**kwargs)
                session._execute=forced_failure
                try:
                    result=session.add_cross_reference_paragraph(runs=[dict(runs[1],prefix='Local ',suffix='',fallbackText='LOCAL-FALLBACK')])
                    assert len(result['issues'])==1 and result['issues'][0]['code']=='CROSS_REFERENCE_FAILED'
                finally:session._execute=execute
                report['checks']['native_operation_local_partial_rollback']=True
                unchanged=session._execute(['set nativeRows to {{content of text object of boundDoc as text}}'])
                try:session.add_cross_reference_paragraph(runs=[dict(runs[1],bookmarkName='invalid')])
                except ValueError:pass
                else:raise AssertionError('Invalid bookmark accepted')
                assert unchanged==session._execute(['set nativeRows to {{content of text object of boundDoc as text}}'])
                report['checks']['invalid_input_document_unchanged']=True
                session.refresh_bookmarks_and_references()
                before_fields=session.snapshot_fields()
                session._execute(['set f to field 1 of boundDoc','set content of field code of f to '+apple_string(' SEQ WPSC_FIG '+chr(92)+'r 8 '),'if (update field f) is false then error "TARGET_UPDATE_FAILED"','set nativeRows to {{"ok"}}'])
                session.refresh_bookmarks_and_references()
                # Changing a SEQ code intentionally invalidates untracked topology;
                # read native REF rows directly instead of masking the stale state.
                report['snapshot_before_target_edit']=[asdict(x) for x in before_fields]
                report['native_rows']=readback(session);inspect_rows(report['native_rows'])
                report['native_styles']=inspect_native_styles(session,report['native_rows'])
                report['checks']['true_ref_refresh_and_exact_paragraph_geometry']=True
                session.save_docx(output/'references.docx');session.export_pdf(output/'references.pdf')
                report['inventory_after']=session._execute(_inventory_commands())
                assert report['inventory_after']==before
                report['checks']['unsaved_sentinel_and_unrelated_preserved']=True
            finally:
                if sentinel and not session._quarantined:
                    assert session._execute(_inventory_commands())==before
                    session._execute([f'set sentinelDoc to document {apple_string(sentinel)}',f'if (content of text object of sentinelDoc as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of sentinelDoc then error "SENTINEL_SAVED"','close sentinelDoc saving no','set nativeRows to {{"ok"}}']);sentinel=None
                if session.staging_root.exists():shutil.copytree(session.staging_root,output/'native-runtime',dirs_exist_ok=True)
        report['after_owned_close']=inventory_without_launch(output,'after-owned-close')
        assert report['after_owned_close']==starting_inventory
        report['checks']['owned_document_native_closed']=True
        source_digest=hashlib.sha256((output/'references.docx').read_bytes()).hexdigest()
        with open_document(output/'references.docx',engine='msoffice',read_only=False,visible=False) as session:
            session.refresh_bookmarks_and_references()
            report['reopen_rows']=readback(session);inspect_rows(report['reopen_rows'])
            report['reopen_styles']=inspect_native_styles(session,report['reopen_rows'])
            shutil.copytree(session.staging_root,output/'reopen-runtime')
        report['after_reopen_close']=inventory_without_launch(output,'after-reopen-close')
        assert report['after_reopen_close']==starting_inventory
        report['checks']['owned_reopen_native_closed']=True
        assert hashlib.sha256((output/'references.docx').read_bytes()).hexdigest()==source_digest
        report['checks']['native_reopen_source_immutable']=True
        import pdfplumber
        with pdfplumber.open(output/'references.pdf') as pdf:
            report['pdf_text']='\n'.join(p.extract_text() or '' for p in pdf.pages)
            assert 'LOCAL-FALLBACK' in report['pdf_text'] and 'PARTIAL-NATIVE-FAILURE' not in report['pdf_text']
            assert '[MISS' in report['pdf_text']
        report['checks']['pdf_text_read']=True
        report['artifact_hashes']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (output/'references.docx',output/'references.pdf')}
        report['passed']=all(report['checks'].values())
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
        if sentinel:report['retained_sentinel_name']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',required=True);args=parser.parse_args()
    if not args.execute:parser.error('--execute required; native Word is opt-in')
    result=run(args.output);print(json.dumps({'passed':result['passed'],'error':result.get('error')}));sys.exit(0 if result['passed'] else 1)
