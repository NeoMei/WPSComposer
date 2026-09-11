"""Guarded direct numbering acceptance. Requires --execute and root Word lease."""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import re
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage
from fixtures.microsoft_parity.macos_word_quality_feasibility import (
    close_sentinel_after_owned, sentinel_guard, source_pair_matches,
)

BOOKMARKS = ['wpsc_eq_'+c*24 for c in 'abc']
GLOBAL = dict(sequenceId='WPSC_EQ',mode='global',prefix='(',suffix=')')
CHAPTER = dict(sequenceId='WPSC_FIG',mode='chapter',chapterStyleLevel=1,resetLevel=1,prefix='[',suffix=']')
SOURCES = [Path(__file__),ROOT/'tests/msoffice/test_macos_word_numbering.py',
           *sorted((ROOT/'skills/WPSComposer').rglob('*.py')),
           *[ROOT/('fixtures/microsoft_parity/'+name) for name in
             ('macos_word_fields.py','macos_word_recovery.py','macos_word_sections.py',
              'macos_word_paragraph_rule.py','macos_word_quality_feasibility.py')]]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def seed_commands():
    return [
        'activate object boundWindow',
        'set content of text object of boundDoc to "Chapter" & return & "left REPLACE right" & return & "TAIL" & return',
        'set numberList to make new list template at boundDoc with properties {name:"WPSCNumberingFixture",outline numbered:true}',
        'set numberLevel to list level 1 of numberList',
        'set number style of numberLevel to list number style arabic',
        'set number format of numberLevel to "%1"',
        'set start at of numberLevel to 1',
        'set linked style of numberLevel to name local of Word style (style heading1) of boundDoc as text',
        'set style of text object of paragraph 1 of boundDoc to style heading1',
        'set selection start of selection of boundWindow to 13',
        'set selection end of selection of boundWindow to 20',
        'set nativeRows to {{"seed",list string of list format of text object of paragraph 1 of boundDoc as text}}',
    ]


def readback_commands():
    lines = ['set nativeRows to {{"body",content of text object of boundDoc as text}}',
             'repeat with ni from 1 to count fields of boundDoc',
             'set nf to field ni of boundDoc',
             'set end of nativeRows to {"field",content of field code of nf as text,content of result range of nf as text}',
             'end repeat']
    for name in BOOKMARKS:
        lines += [f'set nb to bookmark {apple_string(name)} of boundDoc',
                  'set end of nativeRows to {"bookmark",name of nb,start of bookmark of nb,end of bookmark of nb,content of text object of nb as text}']
    return lines + ['set end of nativeRows to {"counts",count tables of boundDoc,count inline shapes of boundDoc,count shapes of boundDoc}',
        'repeat with ni from 1 to count paragraphs of boundDoc',
        'set nr to text object of paragraph ni of boundDoc',
        'if (content of nr as text) contains tab then set end of nativeRows to {"paragraph",content of nr as text,(alignment of paragraph format of nr is align paragraph right),keep together of paragraph format of nr}',
        'end repeat']


def validate_readback(rows):
    try:
        body = next(r[1] for r in rows if r[0]=='body')
        fields = [(r[1].strip().removesuffix(' \\* MERGEFORMAT'),r[2]) for r in rows if r[0]=='field']
        numbering = [(code,result) for code,result in fields if code.startswith(('SEQ ','STYLEREF '))]
        bookmarks = {r[1]:r[4] for r in rows if r[0]=='bookmark' and r[2]<r[3]}
        paragraphs = [r for r in rows if r[0]=='paragraph']
        return (len(numbering)==4 and numbering[0]==('SEQ WPSC_EQ \\* ARABIC','1')
                and re.fullmatch(r'STYLEREF "[^"]+" \\s',numbering[1][0]) is not None
                and numbering[1][1]=='1'
                and numbering[2]==('SEQ WPSC_FIG \\* ARABIC \\s 1','1')
                and numbering[3]==('SEQ WPSC_EQ \\* ARABIC','2')
                and bookmarks==dict(zip(BOOKMARKS,['1','1-1','2']))
                and 'left x=1\t😀(1)尾\r' in body and 'fallback\t[1-1]\r' in body
                and '7\t(2)\r right\rTAIL\r' in body and 'REPLACE' not in body
                and ['counts',0,0,0] in rows and bool(paragraphs)
                and all(r[2:] == [True,True] for r in paragraphs))
    except (TypeError,ValueError,IndexError,StopIteration):
        return False


def run(output, *, execute=False):
    if not execute: raise ValueError('--execute and the root Word lease required')
    output = Path(output).resolve()
    output.mkdir(parents=True,exist_ok=False)
    report = {'status':'FAIL','checks':{},'scope':'direct native numbering only; no OMath/table/figure/heading method or arbitrary middle rollback claim',
              'source_hashes':retain_sources(output,SOURCES)}
    session = reopened = None
    name = token = None
    def flush(): (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def stage(label): report['current_step']=label; flush()
    try:
        report['inventory_before']=inventory(output,'before')
        with create_document('writer',engine='msoffice',visible=False) as session:
            session._retain_evidence=True
            stage('seed')
            report['seed']=session._execute(seed_commands())
            assert report['seed']==[['seed','1']]
            token='NUMBERING SENTINEL 中文😀 '+uuid4().hex
            name=session._execute(['set qualitySentinel to make new document',
                f'set content of text object of qualitySentinel to {apple_string(token)}',
                'activate object (active window of qualitySentinel)',
                'set nativeRows to {{name of qualitySentinel as text}}'])[0][0]
            report['sentinel_name']=name
            report['sentinel_before']=sentinel_preimage(inventory(output,'sentinel-before'),name,token)
            stage('noncollapsed-middle-global')
            assert session.add_equation_number_native(source='x=1',numbering=dict(GLOBAL,prefix='😀(',suffix=')尾'),
                bookmarkName=BOOKMARKS[0],fallbackText='',owner_node_id='eq:shared')=={'issues':[]}
            stage('middle-chapter')
            assert session.add_equation_number_native(source='',numbering=CHAPTER,
                bookmarkName=BOOKMARKS[1],fallbackText='fallback',owner_node_id='eq:shared')=={'issues':[]}
            stage('middle-global-coercion')
            assert session.add_equation_number_native(source=0,numbering=GLOBAL,
                bookmarkName=BOOKMARKS[2],fallbackText=7,owner_node_id='eq:shared')=={'issues':[]}
            stage('native-number-update')
            session.repaginate_and_update_numbering()
            report['native_rows']=session._execute(readback_commands())
            assert validate_readback(report['native_rows'])
            report['checks']['middle_selection_native_numbering_bookmarks_geometry']=True
            report['field_snapshots']=[asdict(x) for x in session.snapshot_fields()]
            assert [tuple(x['stable_key']) for x in report['field_snapshots']]==[
                ('eq:shared','SEQ_EQ',0),('eq:shared','STYLEREF',0),('eq:shared','SEQ_FIG',0),('eq:shared','SEQ_EQ',1)]
            assert all(x['field_category']=='numbering' for x in report['field_snapshots'])
            report['checks']['numbering_owner_category_ordinals']=True
            stage('refs')
            session.add_cross_reference_paragraph(runs=[dict(type='reference',bookmarkName=b,prefix='',suffix=';',fallbackText='BAD-REF') for b in BOOKMARKS],owner_node_id='references')
            session.refresh_bookmarks_and_references()
            rows=session._execute(readback_commands())
            assert [r[2] for r in rows if r[0]=='field' and r[1].strip().startswith('REF ')]==['1','1-1','2']
            report['checks']['number_only_ref_update']=True
            stage('append-checkpoint-and-late-bookmark-error')
            old_tracking=list(session._tracked_numbering)
            before=session._execute(readback_commands())
            checkpoint=session.degradation_checkpoint()
            session._execute(['set np to (end of content of text object of boundDoc) - 1',
                'set selection start of selection of boundWindow to np','set selection end of selection of boundWindow to np',
                'set nativeRows to {{"point",np}}'])
            try:
                session.add_equation_number_native(source='REMOVED',numbering=GLOBAL,bookmarkName='invalid',fallbackText='')
            except ValueError as error: assert str(error)=='invalid native bookmark'
            else: raise AssertionError('Expected late bookmark failure')
            assert len(session._tracked_numbering)==len(old_tracking)+1
            session.rollback_degradation_checkpoint(checkpoint)
            assert session._tracked_numbering==old_tracking
            assert session._execute(readback_commands())==before
            report['checks']['append_rollback_preserves_prior_native_fields_and_tracking']=True
            stage('save-and-pdf')
            session.save_docx(output/'numbering.docx')
            session.export_pdf(output/'numbering.pdf')
            report['sentinel_after']=sentinel_preimage(inventory(output,'sentinel-after'),name,token)
            assert report['sentinel_after']==report['sentinel_before']
            report['checks']['unsaved_sentinel_preserved']=True
        close_sentinel_after_owned(session,output,report,name,token)
        name=None
        stage('owned-reopen')
        digest=sha(output/'numbering.docx')
        with open_document(output/'numbering.docx',engine='msoffice',read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True
            report['reopen_rows']=reopened._execute(readback_commands())
            assert validate_readback(report['reopen_rows'])
        report['checks']['native_reopen_source_preserved']=sha(output/'numbering.docx')==digest
        with ZipFile(output/'numbering.docx') as z:
            doc=z.read('word/document.xml'); (output/'document.xml').write_bytes(doc)
            root=ET.fromstring(doc)
            w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
            codes=[e.text or '' for e in root.iter(w+'instrText')]+[e.get(w+'instr','') for e in root.iter(w+'fldSimple')]
            report['checks']['editable_native_fields_no_omath_table_drawing']=(sum('SEQ WPSC_' in c for c in codes)==3
                and sum('STYLEREF ' in c for c in codes)==1 and not list(root.iter(m+'oMath'))
                and not list(root.iter(w+'tbl')) and not list(root.iter(w+'drawing')))
        import pdfplumber
        with pdfplumber.open(output/'numbering.pdf') as pdf:
            report['pdf_text']='\n'.join(p.extract_text() or '' for p in pdf.pages)
        report['checks']['pdf_readable_formula_and_affixes']=all(t in report['pdf_text'] for t in ('x=1','fallback','TAIL','1-1')) and 'REMOVED' not in report['pdf_text']
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and name:
            try: close_sentinel_after_owned(session,output,report,name,token); name=None
            except BaseException: report['cleanup_failure']=traceback.format_exc()
        for owner,label in ((session,'native-runtime'),(reopened,'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root,output/label,dirs_exist_ok=True)
        report['remaining_sentinel']=name
        report['checks']['sources_unchanged']=all(source_pair_matches(ROOT/p,output/'source'/p,value) for p,value in report['source_hashes'].items())
        if not name and not report.get('cleanup_failure'):
            try:
                report['inventory_final']=inventory(output,'final')
                report['checks']['inventory_preserved']=report['inventory_final']==report.get('inventory_before')
            except BaseException: report['inventory_failure']=traceback.format_exc()
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.pdf','.xml')}
        if report['checks'] and all(report['checks'].values()) and not any(k in report for k in ('error','cleanup_failure','inventory_failure')) and not name:
            report['status']='PASS'
        flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser()
    parser.add_argument('--execute',action='store_true')
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Native numbering requires --execute and the root Word lease',file=sys.stderr)
        return 2
    return 0 if run(args.output,execute=True)['status']=='PASS' else 1


if __name__=='__main__': raise SystemExit(main())
