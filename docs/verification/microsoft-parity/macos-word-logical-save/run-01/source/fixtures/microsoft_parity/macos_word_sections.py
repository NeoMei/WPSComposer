"""Guarded section/metadata probe and direct-method acceptance; no run on import."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
def _inventory_commands():
    # Hash all native UTF-8 text before any result/log is emitted. The saved
    # script contains only this variable expression, never document content.
    # AppleScript quoted form prevents document text from becoming shell code.
    return ['set nativeRows to {}', 'repeat with di from 1 to count documents',
            'set inventoryDoc to document di',
            'if (posix full name of inventoryDoc as text) is not (posix full name of boundDoc as text) then',
            'set documentText to content of text object of inventoryDoc as text',
            'set textHash to do shell script ("/usr/bin/printf %s " & quoted form of documentText & " | /usr/bin/shasum -a 256")',
            'set end of nativeRows to {name of inventoryDoc as text,posix full name of inventoryDoc as text,saved of inventoryDoc,textHash}',
            'end if','end repeat']



def primitive_commands():
    return [
        'set content of text object of boundDoc to "Probe body 中文😀" & return',
        'make new variable at boundDoc with properties {name:"WpsComposerSectionRole_1", variable value:"封面😀"}',
        'set variable value of variable "WpsComposerSectionRole_1" of boundDoc to "正文😀"',
        'set value of document property "Title" of boundDoc to "标题😀"',
        'set value of document property "Author" of boundDoc to "作者😀"',
        'set ownFooter to get footer (section 1 of boundDoc) index header footer primary',
        'create new field text range (text object of ownFooter) field type field page preserve formatting true',
        'insert text "Retain footer " at beginning of text object of ownFooter',
        'set number style of page number options of ownFooter to page number style lowercase roman',
        'set restart numbering at section of page number options of ownFooter to true',
        'set starting number of page number options of ownFooter to 3',
        'set footerRange to text object of ownFooter',
        'delete field 1 of footerRange',
        'if (content of text object of ownFooter as text) does not start with "Retain footer " then error "FOOTER_ERASED"',
        'set content of text object of ownFooter to "Retain footer PAGE NUMPAGES PAGEREF "',
        'set footerRange to text object of ownFooter',
        'set lastCharacter to count characters of footerRange',
        'create new field text range (character lastCharacter of text object of ownFooter) field type field num pages preserve formatting true',
        'set footerRange to text object of ownFooter',
        'set lastCharacter to count characters of footerRange',
        'create new field text range (character lastCharacter of text object of ownFooter) field type field page preserve formatting true',
        *readback_commands(),
    ]


def readback_commands():
    return [
        'set ownFooter to get footer (section 1 of boundDoc) index header footer primary',
        'set nativeRows to {{"version", version}, {"role", variable value of variable "WpsComposerSectionRole_1" of boundDoc}, {"title", value of document property "Title" of boundDoc}, {"author", value of document property "Author" of boundDoc}, {"footer", content of text object of ownFooter}, {"number", number style of page number options of ownFooter as text, restart numbering at section of page number options of ownFooter, starting number of page number options of ownFooter}}',
        'set footerRange to text object of ownFooter',
        'repeat with fieldIndex from 1 to (count fields of footerRange)',
        'set ownField to field fieldIndex of footerRange',
        'set end of nativeRows to {"field", field type of ownField as text, content of field code of ownField as text}',
        'end repeat',
    ]


def run(output, mode):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'passed':False,'engine':'msoffice','platform':'darwin','component':'writer','mode':mode,'checks':{},'source_hashes':{str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in [Path(__file__),ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py']}}
    session=None;sentinel=None
    try:
        from skills.WPSComposer import create_document
        with create_document('writer',engine='msoffice',visible=False) as session:
            token='Section sentinel '+uuid4().hex+' 中文😀'
            sentinel=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
            before=session._execute(_inventory_commands());report['inventory_before']=before
            assert any(row[0]==sentinel and row[2] is False for row in before)
            assert session._execute([f'set nativeRows to {{{{(content of text object of document {apple_string(sentinel)} as text) is {apple_string(token)} & return}}}}']) == [[True]]
            try:
                if mode=='probe':
                    report['native_rows']=session._execute(primitive_commands())
                elif mode=='variants':
                    exercise_variants(session,report)
                else:
                    exercise(session,report,output)
                session.save_docx(output/'sections.docx');session.export_pdf(output/'sections.pdf')
                after=session._execute(_inventory_commands());report['inventory_after']=after
                assert before==after
                report['checks']['unsaved_sentinel_and_unrelated_documents_preserved']=True
                shutil.copytree(session.staging_root,output/'native-runtime')
            except BaseException:
                if not session._quarantined:
                    after=session._execute(_inventory_commands());report['failure_inventory_after']=after
                    report['checks']['failure_preserves_sentinel_and_unrelated_documents']=before==after
                    session.save_docx(output/'failed-sections.docx')
                    session.export_pdf(output/'failed-sections.pdf')
                raise
            finally:
                if sentinel and not session._quarantined:
                    session._execute([f'set sentinelDoc to document {apple_string(sentinel)}',f'if (content of text object of sentinelDoc as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of sentinelDoc then error "SENTINEL_SAVED"','close sentinelDoc saving no','set nativeRows to {{"ok"}}']);sentinel=None
                if session.staging_root.exists():
                    shutil.copytree(session.staging_root,output/'final-runtime',dirs_exist_ok=True)
        report['checks']['exact_owned_document_close']=session._closed
        source=output/'sections.docx';digest=hashlib.sha256(source.read_bytes()).hexdigest()
        with MacWordSession.open_document(source) as session:
            report['reopen_rows']=session._execute(readback_commands())
            if mode=='probe':
                assert report['reopen_rows']==report['native_rows']
            elif mode=='variants':
                verify_variants(session,output,report)
            else:
                verify_reopened(session,output,report)
            shutil.copytree(session.staging_root,output/'reopen-runtime')
        assert hashlib.sha256(source.read_bytes()).hexdigest()==digest
        report['checks']['native_reopen_source_preserved']=True
        report['source_docx_sha256_before_reopen']=digest
        report['source_docx_sha256_after_reopen']=hashlib.sha256(source.read_bytes()).hexdigest()
        report['artifact_hashes']={name:hashlib.sha256((output/name).read_bytes()).hexdigest() for name in ('sections.docx','sections.pdf')}
        report['checks']['exact_owned_reopen_close']=session._closed
        report['limitations']=['Representative checks are listed per mode; separate acceptance and variants reports are required for the complete exercised branch set.','No attached session or user UI edit/Undo acceptance is claimed.','Runtime timeout/quarantine checks are separate unit tests; this fixture does not deliberately cause uncertain native completion.']
        if mode=='acceptance':
            scalar_source=output/'scalar-metadata.docx'
            scalar_digest=hashlib.sha256(scalar_source.read_bytes()).hexdigest()
            with MacWordSession.open_document(scalar_source) as scalar_session:
                scalar_rows=scalar_session._execute(['set nativeRows to {{value of document property "Title" of boundDoc, value of document property "Author" of boundDoc}}'])
                assert scalar_rows==[['7','']]
                report['metadata_scalar_reopen_rows']=scalar_rows
                shutil.copytree(scalar_session.staging_root,output/'scalar-reopen-runtime')
            assert scalar_session._closed
            assert hashlib.sha256(scalar_source.read_bytes()).hexdigest()==scalar_digest
            report['scalar_metadata_sha256_before_reopen']=scalar_digest
            report['scalar_metadata_sha256_after_reopen']=hashlib.sha256(scalar_source.read_bytes()).hexdigest()
            report['checks']['scalar_metadata_7_false_saved_native_reopen_and_source_preserved']=True
        report['passed']=all(report['checks'].values())
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
        if sentinel: report['retained_sentinel_name']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report

def exercise(session,report,output):
    session.configure_section(role='封面😀', landscape=False, page_size='a4', margins={'top':60,'bottom':60,'left':66,'right':66},
                              page_number_format='roman', start_page_number=3,restart_page_numbering=True,
                              header_text='Section One',footer_text='Keep footer 1 ',link_to_previous_header=False,link_to_previous_footer=False)
    assert session._first_section_configured
    assert session.inspect_document()['counts']['sections']==1
    session.set_page_role('正文😀')
    session.set_document_metadata(title=7,author=False)
    report['metadata_scalar_readback']=session._execute(['set nativeRows to {{value of document property "Title" of boundDoc, value of document property "Author" of boundDoc}}'])
    assert report['metadata_scalar_readback']==[['7','']]
    session.save_docx(output/'scalar-metadata.docx')
    session.set_document_metadata(title=None,author=None)
    assert session._execute(['set nativeRows to {{value of document property "Title" of boundDoc, value of document property "Author" of boundDoc}}']) == [['','']]
    session.set_document_metadata(title='标题😀',author='作者😀')
    session.add_heading_level('重复标题😀',2)
    session.add_paragraph('First duplicate remains portrait 中文😀')
    session.configure_section(role='landscape😀',landscape=True,page_size='letter',
                              margins={'top':54,'bottom':54,'left':60,'right':60},
                              page_number_format='roman-upper',start_page_number=5,restart_page_numbering=True,
                              header_text='Section Two',footer_text='Keep footer 2 ',link_to_previous_header=False,link_to_previous_footer=False)
    session.add_paragraph('Landscape configured section 中文😀')
    session.configure_section(role='body😀',landscape=False,page_size='a4',page_number_format='arabic',start_page_number=7,restart_page_numbering=True,
                              header_text='Section Three',footer_text='Keep footer 3 PAGE NUMPAGES PAGEREF ',link_to_previous_header=False,link_to_previous_footer=False)
    session._execute(['set ownFooter to get footer (section 3 of boundDoc) index header footer primary',
                      'set footerRange to text object of ownFooter','set lastCharacter to count characters of footerRange',
                      'create new field text range (character lastCharacter of text object of ownFooter) field type field num pages preserve formatting true','set nativeRows to {{"ok"}}'])
    session.set_page_numbering('none')
    none=session._execute(['set ownFooter to get footer (section 3 of boundDoc) index header footer primary',
                          'set footerRange to text object of ownFooter',
                          'set nativeRows to {{content of footerRange as text, count fields of footerRange, field type of field 1 of footerRange as text}}'])
    assert none[0][0].startswith('Keep footer 3 PAGE NUMPAGES PAGEREF ') and none[0][1:]==[1,'field num pages'],none
    report['none_readback']=none
    session.set_page_numbering('arabic',7,True);session.set_page_numbering('arabic',7,True)
    session.add_paragraph('Third portrait before pending heading 中文😀')
    session.add_heading_level('重复标题😀',2)
    pending=session._pending_heading;report['pending_native_position']=pending
    session.add_landscape_section_before_pending_heading()
    session.add_table(2,2,[['Landscape object 中文😀','Value'],['Tail','42']])
    session.set_page_role('media😀')
    count=session.inspect_document()['counts']['sections']
    assert count==4
    session.add_heading_level('Invalidated pending 😀',2);session.add_paragraph('Intervening body invalidates pending 中文😀')
    try:session.add_landscape_section_before_pending_heading()
    except ValueError:pass
    else:raise AssertionError('Stale pending heading accepted')
    assert session.inspect_document()['counts']['sections']==count
    session._execute(['repaginate boundDoc','set nativeRows to {{"ok"}}'])
    report['checks']['configure_three_sections_and_pending_fourth']=True
    report['checks']['none_preserves_plain_text_and_numpages']=True
    report['checks']['intervening_content_invalidates_pending_without_section_change']=True
    report['method_checks']={
        'configure_section':['first call has one section','two later calls each add one next-page section','a4/letter dimensions and portrait-landscape-portrait geometry','margins and detached header/footer'],
        'set_page_numbering':['roman lower/upper/arabic','restart true with 3/5/7 start','none retains NUMPAGES and plain PAGE text','repeated calls maintain one PAGE'],
        'set_page_role':['create/update document-local section variables','Chinese emoji after reopen'],
        'set_document_metadata':['Title Author Chinese emoji after reopen','title=7 and author=False produce 7 and empty with saved native reopen'],
        'add_landscape_section_before_pending_heading':['native UTF16 start','duplicate heading second occurrence only','intervening-content negative'],
    }


def verify_reopened(session,output,report):
    from zipfile import ZipFile
    import xml.etree.ElementTree as ET
    import fitz
    snap=session.inspect_document()
    assert snap['counts']['sections']==4
    rows=session._execute(['set nativeRows to {}','repeat with si from 1 to count sections of boundDoc',
                          'set ownSection to section si of boundDoc','set ownFooter to get footer ownSection index header footer primary',
                          'set ownHeader to get header ownSection index header footer primary',
                          'set end of nativeRows to {si,orientation of page setup of ownSection as text,page width of page setup of ownSection,page height of page setup of ownSection,top margin of page setup of ownSection,bottom margin of page setup of ownSection,left margin of page setup of ownSection,right margin of page setup of ownSection,link to previous of ownHeader,link to previous of ownFooter,number style of page number options of ownFooter as text,restart numbering at section of page number options of ownFooter,starting number of page number options of ownFooter,variable value of variable ("WpsComposerSectionRole_" & si) of boundDoc}',
                          'end repeat'])
    report['section_reopen_rows']=rows
    assert [r[1] for r in rows]==['orient portrait','orient landscape','orient portrait','orient landscape']
    assert [r[13] for r in rows]==['正文😀','landscape😀','body😀','media😀']
    assert [r[10] for r in rows[:3]]==['page number style lowercase roman','page number style uppercase roman','page number style arabic']
    assert [r[12] for r in rows[:3]]==[3,5,7]
    assert all(r[11] for r in rows[:3])
    assert rows[3][11] is False
    assert rows[1][2:4]==[792,612]
    assert rows[0][4:8]==[60,60,66,66]
    assert rows[1][4:8]==[54,54,60,60]
    assert all(not r[8] and not r[9] for r in rows[1:3])
    w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    from fixtures.microsoft_parity.macos_word_business_followup import _verify_page_fields,_field_commands
    with ZipFile(output/'sections.docx') as package:
        body=ET.fromstring(package.read('word/document.xml'))
        footer_parts=_verify_page_fields(package,body)
        for part in footer_parts:
            assert _field_commands(ET.fromstring(package.read(part))).count('PAGE')==1
        assert _field_commands(ET.fromstring(package.read(footer_parts[2]))).count('NUMPAGES')==1
        section_texts=[];texts=[]
        for element in body.find(w+'body'):
            texts += [t.text or '' for t in element.iter(w+'t')]
            if element.tag==w+'sectPr' or element.find(w+'pPr/'+w+'sectPr') is not None:
                section_texts.append(''.join(texts));texts=[]
        assert len(section_texts)==4
        assert '重复标题😀' in section_texts[0] and '重复标题😀' in section_texts[3]
        assert '重复标题😀' not in section_texts[2]
        assert 'Landscape object 中文😀' in section_texts[3]
        core=ET.fromstring(package.read('docProps/core.xml'))
        assert core.find('{http://purl.org/dc/elements/1.1/}title').text=='标题😀'
        assert core.find('{http://purl.org/dc/elements/1.1/}creator').text=='作者😀'
    with fitz.open(output/'sections.pdf') as pdf:
        assert len(pdf)==4,len(pdf)
        text=''.join(page.get_text() for page in pdf)
        assert 'First duplicate remains portrait' in text and 'Landscape object' in text
        assert pdf[0].rect.width<pdf[0].rect.height and pdf[1].rect.width>pdf[1].rect.height
        assert pdf[2].rect.width<pdf[2].rect.height and pdf[3].rect.width>pdf[3].rect.height
        report['pdf_text']=[page.get_text() for page in pdf]
        assert 'Keep footer 1 iii' in report['pdf_text'][0]
        assert 'Keep footer 2 V' in report['pdf_text'][1]
        assert 'Keep footer 3 PAGE NUMPAGES PAGEREF 47' in report['pdf_text'][2]
        assert 'Keep footer 3 PAGE NUMPAGES PAGEREF 48' in report['pdf_text'][3]
        report['checks']['rendered_PAGE_results_iii_V_7_8_and_NUMPAGES_4']=True
    report['checks']['native_reopen_metadata_roles_numbering_dimensions']=True
    report['checks']['saved_footer_true_PAGE_and_NUMPAGES']=True
    report['checks']['pending_heading_and_table_same_landscape_section']=True
    report['checks']['four_page_pdf_orientation_and_text']=True


def exercise_variants(session,report):
    session.configure_section(role='正文😀',page_size='legal',landscape=False,margins={'top':40,'bottom':42,'left':50,'right':52},
                              header_text='Continuation Header',footer_text='Variant ',page_number_format='arabic',restart_page_numbering=True,start_page_number=9)
    session.set_document_metadata(title='标题😀',author='作者😀')
    session.add_paragraph('Legal portrait start 9 中文😀')
    session.configure_section(role='a3 landscape',page_size='a3',landscape=True,page_number_format='continue',restart_page_numbering=False,start_page_number=99,
                              link_to_previous_header=True,link_to_previous_footer=True)
    session.add_paragraph('A3 landscape continues 10 中文😀')
    session.configure_section(role='inherited landscape',page_size='legal',margins={'top':50,'bottom':52,'left':70,'right':72},
                              page_number_format='continue',restart_page_numbering=False,start_page_number=99,link_to_previous_header=True,link_to_previous_footer=True)
    session.add_paragraph('Legal inherits landscape and continues 11 中文😀')
    session.configure_section(role='a3 portrait',page_size='a3',landscape=False,page_number_format='continue',restart_page_numbering=False,start_page_number=99,
                              link_to_previous_header=True,link_to_previous_footer=True)
    session.add_paragraph('A3 portrait continues 12 中文😀')
    session._execute(['repaginate boundDoc','set nativeRows to {{"ok"}}'])
    report['method_checks']={'configure_section':['legal/A3 native sizes','inherited landscape with explicit margins','linked header/footer true'],
                             'set_page_numbering':['continue with restart false and supplied start ignored natively','Arabic rendered 9/10/11/12']}


def verify_variants(session,output,report):
    import fitz
    rows=session._execute(['set nativeRows to {}','repeat with si from 1 to count sections of boundDoc','set ownSection to section si of boundDoc',
                          'set ownFooter to get footer ownSection index header footer primary','set ownHeader to get header ownSection index header footer primary',
                          'set end of nativeRows to {page width of page setup of ownSection,page height of page setup of ownSection,orientation of page setup of ownSection as text,top margin of page setup of ownSection,bottom margin of page setup of ownSection,left margin of page setup of ownSection,right margin of page setup of ownSection,link to previous of ownHeader,link to previous of ownFooter,restart numbering at section of page number options of ownFooter,starting number of page number options of ownFooter}', 'end repeat'])
    report['section_reopen_rows']=rows
    expected=[(612,1008),(1190.55,841.89),(1008,612),(841.89,1190.55)]
    assert len(rows)==4
    for row,(width,height) in zip(rows,expected):
        assert abs(row[0]-width)<0.1 and abs(row[1]-height)<0.1
    assert rows[2][3:7]==[50,52,70,72]
    assert all(row[7:10]==[True,True,False] for row in rows[1:])
    assert all(row[10]==0 for row in rows[1:])
    with fitz.open(output/'sections.pdf') as pdf:
        assert len(pdf)==4
        report['pdf_text']=[page.get_text() for page in pdf]
        for i,page in enumerate(pdf):
            assert f'Variant {9+i}' in page.get_text()
            assert 'Continuation Header' in page.get_text()
    report['checks']['native_legal_a3_inherited_orientation_margins_links']=True
    report['checks']['native_continue_ignores_start_and_renders_9_10_11_12']=True


def cleanup_failed_sentinels(output,reports):
    """Close only failed-run sentinels whose complete recorded identity matches."""
    import re
    from skills.WPSComposer import create_document
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    owned=[]
    for path in reports:
        original=json.loads(Path(path).read_text())
        name=original.get('retained_sentinel_name')
        if not name:continue
        matches=[row for row in original['inventory_before'] if row[0]==name]
        if len(matches)!=1:raise ValueError('Ambiguous retained sentinel')
        row=matches[0]
        if row[0]!=row[1] or row[2] is not False or not re.fullmatch(r'Section sentinel [0-9a-f]{32} 中文😀\r',row[3]):
            raise ValueError('Report does not prove owned synthetic unsaved sentinel')
        if row not in owned:owned.append(row)
    result={'passed':False,'engine':'msoffice','platform':'darwin','component':'writer','source_reports':[str(p) for p in reports]}
    with create_document('writer',engine='msoffice',visible=False) as session:
        result['inventory_before']=session._execute(_inventory_commands())
        for name,path,saved,text in owned:
            lines=[f'set sentinelDoc to document {apple_string(name)}',
                   f'if (posix full name of sentinelDoc as text) is not {apple_string(path)} then error "SENTINEL_PATH_CHANGED"',
                   'if saved of sentinelDoc then error "SENTINEL_SAVED"',
                   f'if (content of text object of sentinelDoc as text) is not {apple_string(text)} then error "SENTINEL_TEXT_CHANGED"',
                   'close sentinelDoc saving no','set nativeRows to {{"ok"}}']
            assert session._execute(lines)==[['ok']]
        result['inventory_after']=session._execute(_inventory_commands())
        assert result['inventory_after']==[row for row in result['inventory_before'] if row[0] not in [o[0] for o in owned]]
        result['closed_exact_sentinel_names']=[row[0] for row in owned]
        shutil.copytree(session.staging_root,output/'native-runtime')
    result['passed']=session._closed
    (output/'report.json').write_text(json.dumps(result,ensure_ascii=False,indent=2))
    return result


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--output-dir',required=True,type=Path);p.add_argument('--mode',choices=['probe','acceptance','variants'],default='probe');a=p.parse_args()
    result=run(a.output_dir,a.mode);print(json.dumps(result,ensure_ascii=False));raise SystemExit(0 if result['passed'] else 1)
