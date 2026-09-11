"""Guarded Mac Word fields/index probe; requires explicit --execute."""
from __future__ import annotations
import argparse
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document,open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_sections import _inventory_commands


def retain_sources(output,sources):
    hashes={}
    for source in sources:
        relative=source.relative_to(ROOT)
        target=Path(output)/'source'/relative;target.parent.mkdir(parents=True,exist_ok=True)
        shutil.copy2(source,target)
        digest=hashlib.sha256(source.read_bytes()).hexdigest()
        assert hashlib.sha256(target.read_bytes()).hexdigest()==digest
        hashes[str(relative)]=digest
    return hashes


def validate_snapshots(rows):
    from collections import Counter
    expected=Counter({'STYLEREF':1,'SEQ_FIG':1,'SEQ_TAB':1,'REF':1,'TOC':2,'TOF_FIG':1,'TOF_TAB':1,'PAGE':1,'NUMPAGES':1})
    assert Counter(x['stable_key'][1] for x in rows)==expected
    assert len({tuple(x['stable_key']) for x in rows})==len(rows)
    for row in rows:
        owner,kind,ordinal=row['stable_key']
        expected_owner='doc:toc' if kind=='TOC' else 'figures' if kind=='TOF_FIG' else 'tables' if kind=='TOF_TAB' else 'story:primary footer/chain:1' if kind in {'PAGE','NUMPAGES'} else 'story:main text/chain:1'
        assert owner==expected_owner and ordinal in (range(2) if kind=='TOC' else range(1))
        assert row['field_category']==('index' if kind in {'TOC','TOF_FIG','TOF_TAB'} else 'page' if kind in {'PAGE','NUMPAGES'} else 'field')


def seeded_fields(session):
    # Only the exactly bound synthetic document is modified by this probe.
    lines=[]
    for label,typ,code in [('Figure ','field sequence','WPSC_FIG'),('Table ','field sequence','WPSC_TAB')]:
        lines+=session._business_paragraph(label+'😀','Body Text')
        lines+=session._position('end')+['set r to create range boundDoc start insertionPoint end insertionPoint',f'create new field text range r field type {typ} field text {apple_string(code)} preserve formatting true','set insertionPoint to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start insertionPoint end insertionPoint',f'set content of r to {apple_string(" "+label+"Caption 😀")} & return']
    lines+=['set r to create range boundDoc start 0 end 8','make new bookmark at boundDoc with properties {name:"WPSC_RefProbe",text object:r}']
    lines+=session._position('end')+['set r to create range boundDoc start insertionPoint end insertionPoint','create new field text range r field type field ref field text "WPSC_RefProbe" preserve formatting true','set insertionPoint to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start insertionPoint end insertionPoint','set content of r to return']
    lines+=session._position('end')+['set r to create range boundDoc start insertionPoint end insertionPoint','create new field text range r field type field style ref field text "1" preserve formatting true','set insertionPoint to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start insertionPoint end insertionPoint','set content of r to return']
    session._execute_structural(lines+['set nativeRows to {{"ok"}}'])
    session.set_footer('Retain PAGE NUMPAGES PAGEREF ')
    session.set_page_numbering('arabic',1,True)
    session._execute(['set ownFooter to get footer (section 1 of boundDoc) index header footer primary',
                      'set footerRange to text object of ownFooter',
                      'set lastCharacter to count characters of footerRange',
                      'create new field text range (character lastCharacter of text object of ownFooter) field type field num pages preserve formatting true',
                      'set nativeRows to {{"ok"}}'])


def readback(session):
    return session._execute(['set nativeRows to {{"counts",count tables of contents of boundDoc,count tables of figures of boundDoc,count bookmarks of boundDoc}}','repeat with fi from 1 to count fields of boundDoc','set f to field fi of boundDoc','set end of nativeRows to {"field",content of field code of f as text,content of result range of f as text}','end repeat','set ownFooter to get footer (section 1 of boundDoc) index header footer primary','set end of nativeRows to {"footer",content of text object of ownFooter as text}'])


def growth_acceptance(output,report):
    """A refreshed TOC changes native length before an otherwise unchanged REF."""
    with create_document('writer',engine='msoffice',visible=False) as session:
        token='Growth sentinel '+uuid4().hex+' 😀'
        sentinel=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
        before_inventory=session._execute(_inventory_commands())
        try:
            session.add_heading_level('Initial heading 😀',1)
            session._execute(['set r to create range boundDoc start 0 end 7','make new bookmark at boundDoc with properties {name:"WPSC_GrowthRef",text object:r}','set nativeRows to {{"ok"}}'])
            handle=session.insert_toc('Growth Contents')
            session._execute_structural(session._position('end')+['set r to create range boundDoc start insertionPoint end insertionPoint','create new field text range r field type field ref field text "WPSC_GrowthRef" preserve formatting true','set insertionPoint to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start insertionPoint end insertionPoint','set content of r to return','set nativeRows to {{"ok"}}'])
            session.add_heading_level('Added after TOC 😀',2)
            session.refresh_bookmarks_and_references()
            before=session.snapshot_fields()
            session.refresh_indexes()
            after=session.snapshot_fields()
            assert [x.stable_key for x in before]==[x.stable_key for x in after]
            assert before[0].field_category=='index' and before[0].result_hash!=after[0].result_hash
            assert before[1].stable_key[1]=='REF' and before[1].result_hash==after[1].result_hash
            report['growth_snapshots']={'before':[asdict(x) for x in before],'after':[asdict(x) for x in after]}
            session.save_docx(output/'growth.docx');session.export_pdf(output/'growth.pdf')
            assert session.snapshot_fields()==after
            assert session._execute(_inventory_commands())==before_inventory
            report['checks']['tracked_toc_growth_preserves_following_ref_identity']=True
        except BaseException:
            if not session._quarantined:
                session.save_docx(output/'failed-growth.docx');session.export_pdf(output/'failed-growth.pdf')
            raise
        finally:
            if not session._quarantined:
                session._execute([f'set d to document {apple_string(sentinel)}',f'if (content of text object of d as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of d then error "SENTINEL_SAVED"','close d saving no','set nativeRows to {{"ok"}}'])
            if session.staging_root.exists():shutil.copytree(session.staging_root,output/'growth-runtime',dirs_exist_ok=True)
    report['checks']['growth_exact_owned_close']=session._closed
    import pdfplumber
    with pdfplumber.open(output/'growth.pdf') as pdf:
        assert 'Added after TOC' in '\n'.join(p.extract_text() or '' for p in pdf.pages)


def run(output,mode):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    sources=[Path(__file__),ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_fields.py',ROOT/'skills/WPSComposer/scripts/msoffice/errors.py']
    report={'passed':False,'mode':mode,'source_hashes':retain_sources(output,sources),'checks':{}}
    session=None;sentinel=None
    try:
        with create_document('writer',engine='msoffice',visible=False) as session:
            token='Fields sentinel '+uuid4().hex+' 中文😀'
            sentinel=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
            before=session._execute(_inventory_commands());report['inventory_before']=before
            assert any(r[0]==sentinel and r[2] is False for r in before)
            try:
                session.add_heading_level('Heading 😀 Duplicate',1)
                session.add_paragraph('Probe body preserves existing text.')
                if mode=='accept':
                    # Enough native headings to force the index itself across pages.
                    session._business_commit(sum((session._business_paragraph('Heading 😀 Duplicate '+str(i),'Heading 2') for i in range(72)),[]),structural=True)
                    session.add_page_break();session.add_heading_level('Heading 😀 Duplicate',3)
                seeded_fields(session)
                handle=session.insert_toc('目录😀')
                session.insert_toc_with_styles('Styled Contents',{'minFontSizePt':{'toc1':11,'toc2':10,'toc3':9},'minSpaceAfterPt':{'toc2':2}})
                session.insert_caption_index_native(title='Figures😀',sequence_id='WPSC_FIG',title_style_id='Missing WPSC style',owner_node_id='figures')
                session.insert_caption_index_native(title='Tables😀',sequence_id='WPSC_TAB',title_style_id='Body Text',owner_node_id='tables')
                session.insert_figure_index('Literal figures');session.insert_table_index('Literal tables')
                session.repaginate_and_update_numbering();session.refresh_bookmarks_and_references();session.refresh_indexes();session.repaginate_and_update_page_fields()
                report['snapshot']=[asdict(x) for x in session.snapshot_fields()]
                assert [asdict(x) for x in session.snapshot_fields()]==report['snapshot']
                report['checks']['repeat_snapshot_stable']=True
                report['native_rows']=readback(session)
                assert report['native_rows'][0][1:3]==[2,2]
                index_rows=[r for r in report['native_rows'] if r[0]=='field' and r[1].strip().startswith('TOC ')]
                assert len(index_rows)==4
                assert all('目录' not in r[2] and 'Styled Contents' not in r[2] for r in index_rows[:2])
                assert 'Figure Caption' in index_rows[2][2] and 'Table Caption' not in index_rows[2][2]
                assert 'Table Caption' in index_rows[3][2] and 'Figure Caption' not in index_rows[3][2]
                validate_snapshots(report['snapshot'])
                report['checks']['exact_sequence_filtered_native_indexes_and_field_kinds']=True
                report['checks']['native_toc_title_exclusion']=True
                if mode=='accept':
                    assert report['snapshot'][0]['toc_page_count']>=4
                    assert all(r[2].count('Heading')>=74 for r in index_rows[:2])
                    report['checks']['multi_page_native_toc_ranges']=True
                session.save_docx(output/'fields.docx');session.export_pdf(output/'fields.pdf')
                after=session._execute(_inventory_commands());report['inventory_after']=after
                assert before==after
                report['checks']['unsaved_sentinel_and_unrelated_documents_preserved']=True
                session._execute([f'delete bookmark {apple_string(handle.bookmark)} of boundDoc','set nativeRows to {{"ok"}}'])
                try:session.snapshot_fields()
                except Exception as exc:
                    assert getattr(exc,'code',None)=='NATIVE_WORD_FIELD_IDENTITY_STALE'
                    report['checks']['deleted_owned_index_identity_rejected']=True
                else:raise AssertionError('Deleted bookmark was silently rebound')
                shutil.copytree(session.staging_root,output/'native-runtime')
            except BaseException:
                if not session._quarantined:
                    report['failure_inventory_after']=session._execute(_inventory_commands())
                    report['checks']['failure_preserves_unrelated']=before==report['failure_inventory_after']
                    session.save_docx(output/'failed-fields.docx');session.export_pdf(output/'failed-fields.pdf')
                raise
            finally:
                if sentinel and not session._quarantined:
                    session._execute([f'set sentinelDoc to document {apple_string(sentinel)}',f'if (content of text object of sentinelDoc as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of sentinelDoc then error "SENTINEL_SAVED"','close sentinelDoc saving no','set nativeRows to {{"ok"}}']);sentinel=None
                if session.staging_root.exists():shutil.copytree(session.staging_root,output/'final-runtime',dirs_exist_ok=True)
        report['checks']['exact_owned_close']=session._closed
        digest=hashlib.sha256((output/'fields.docx').read_bytes()).hexdigest()
        with open_document(output/'fields.docx',engine='msoffice',read_only=False,visible=False) as session:
            session.repaginate_and_update_page_fields()
            report['reopen_rows']=readback(session)
            def canonical(rows):
                return [[r[0],r[1].strip(),r[2]] if r[0]=='field' else r for r in rows]
            assert canonical(report['reopen_rows'])==canonical(report['native_rows'])
            report['reopen_snapshot']=[asdict(x) for x in session.snapshot_fields()]
            shutil.copytree(session.staging_root,output/'reopen-runtime')
        assert hashlib.sha256((output/'fields.docx').read_bytes()).hexdigest()==digest
        report['checks']['native_reopen_source_preserved']=True
        report['docx_sha256_before_reopen']=digest
        report['docx_sha256_after_reopen']=hashlib.sha256((output/'fields.docx').read_bytes()).hexdigest()
        report['checks']['exact_owned_reopen_close']=session._closed
        import zipfile,xml.etree.ElementTree as E
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        with zipfile.ZipFile(output/'fields.docx') as z:
            d=E.fromstring(z.read('word/document.xml'));codes=[x.text for x in d.findall('.//w:instrText',ns)]+[x.get('{'+ns['w']+'}instr') for x in d.findall('.//w:fldSimple',ns)]
            report['docx_field_codes']=codes
            assert any('SEQ WPSC_FIG' in x for x in codes) and any('SEQ WPSC_TAB' in x for x in codes)
            assert any('TOC' in x and '\\c "WPSC_FIG"' in x for x in codes)
        import pdfplumber
        with pdfplumber.open(output/'fields.pdf') as pdf:
            texts=[p.extract_text() or '' for p in pdf.pages];report['pdf_pages']=len(texts)
            assert any('Figure index placeholder' in t for t in texts)
            assert any('Table index placeholder' in t for t in texts)
            report['pdf_text']='\n\n'.join(texts)
        report['checks']['independent_docx_and_pdf_read']=True
        growth_acceptance(output,report)
        report['artifact_hashes']={name:hashlib.sha256((output/name).read_bytes()).hexdigest() for name in ('fields.docx','fields.pdf','growth.docx','growth.pdf')}
        report['passed']=all(report['checks'].values())
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
        if sentinel:report['retained_sentinel_name']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',required=True);p.add_argument('--mode',choices=['probe','accept'],default='probe');args=p.parse_args()
    if not args.execute:p.error('--execute required')
    r=run(args.output,args.mode);print(json.dumps({'passed':r['passed'],'error':r.get('error')}));sys.exit(0 if r['passed'] else 1)
