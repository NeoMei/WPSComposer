"""Guarded DRAFT acceptance for the second direct Word slice.

This is not evidence until explicitly run on an available native host. The
existing output directory is never overwritten, and unrelated documents are
not edited or closed. Only a task-owned new document is used.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))


def run(output):
    from skills.WPSComposer import create_document,inspect
    from skills.WPSComposer.scripts.reference_styles import STYLES,get_heading_style
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'passed':False,'checks':{},'limitations':[
        'Custom character styles remain unsupported.',
        'Index refresh with nonempty TOC/TOF requires a separate native seeded fixture.',
        'Fresh generation proof does not establish attached-session behavior.',
    ]}
    session=None
    inventory=['set nativeRows to {}','repeat with di from 1 to count documents','set d to document di',
       'if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text,posix full name of d as text,saved of d}','end repeat']
    try:
        with create_document('writer',engine='msoffice',visible=False) as session:
            before=session._execute(inventory)
            assert session.reset() is None
            session.ensure_styles({'SourceCode':STYLES['SourceCode'],'CustomBody':{
                'name':'Custom Body','based_on':'Body Text','font_size':12,'color':'#123456',
                'indent_first':24,'space_after':4}})
            session.ensure_heading_styles({level:get_heading_style(level) for level in range(1,7)})
            session.apply_heading_text_color('#224466')
            session.add_styled_paragraph('Followup native body','Custom Body')
            session.add_code_lines(['native_code = 1','','native_code += 1'])
            session.set_columns(2)
            session.add_paragraph('Two-column native text ' * 80)
            session.set_page_number_in_footer()
            session.add_section(True,continuous=False)
            session.add_paragraph('Landscape followup section')
            session.add_section(False,continuous=True)
            session.add_paragraph('Portrait followup section')
            session.compact_terminal_paragraph()
            session.refresh_indexes();session.update_fields();session.finalize_fields(max_rounds=3)
            snapshot=session.refresh_fields(0)
            assert len(snapshot)==1 and snapshot[0].total_pages>=1
            report['field_snapshot']=snapshot[0].to_dict()
            report['checks']['refresh_return_contract']=True
            session.save_docx(output/'followup.docx');session.export_pdf(output/'followup.pdf')
            report['checks']['unrelated_bindings_saved_state_unchanged']=before==session._execute(inventory)
            shutil.copytree(session.staging_root,output/'native-runtime')
        report['checks']['exact_owned_close']=session._closed;session=None
        snap=inspect(output/'followup.docx',engine='msoffice')
        (output/'reopened.json').write_text(json.dumps(snap,ensure_ascii=False,indent=2))
        texts=[p.get('text','') for p in snap['paragraphs']]
        assert 'native_code = 1' in texts and 'Landscape followup section' in texts
        assert snap['counts']['sections']==3
        report['checks']['native_reopen_text_sections']=True
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        with ZipFile(output/'followup.docx') as package:
            footer_xml=[ET.fromstring(package.read(name)) for name in package.namelist() if name.startswith('word/footer') and name.endswith('.xml')]
            assert any('PAGE' in ''.join(node.itertext()) for node in footer_xml)
            body=ET.fromstring(package.read('word/document.xml'))
            columns=body.findall('.//w:sectPr/w:cols',ns)
            assert len(columns)==3 and all(node.get('{'+ns['w']+'}num')=='2' for node in columns)
            report['checks']['native_two_column_sections']=True
            assert not any('PAGE' in (node.text or '') for node in body.findall('.//w:instrText',ns))
            styles=ET.fromstring(package.read('word/styles.xml'))
            assert any(node.get('{'+ns['w']+'}val')=='Custom Body' for node in styles.findall('.//w:style/w:name',ns))
        report['checks']['native_field_story_and_named_style']=True
        import pdfplumber
        with pdfplumber.open(output/'followup.pdf') as pdf:
            text='\n'.join(page.extract_text() or '' for page in pdf.pages)
            assert 'Followup native body' in text and 'Page' in text
            report['pdf_pages']=len(pdf.pages)
        report['checks']['native_pdf_text']=True
        report['hashes']={name:hashlib.sha256((output/name).read_bytes()).hexdigest() for name in ('followup.docx','followup.pdf')}
        report['passed']=all(report['checks'].values())
        report['acceptance_status']='BASIC_CHECKS_ONLY_PENDING_STYLE_COLUMN_FIELD_REVIEW'
        report['requires_artifact_review']=['Styles font/indent/heading definitions','Code empty line and trailing spacer','Portrait/landscape sections','Two-column text geometry','PAGE footer visible results','Terminal empty paragraph compacted']
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--output-dir',type=Path,required=True)
    result=run(parser.parse_args().output_dir);print(json.dumps(result,ensure_ascii=False))
    return 0 if result['passed'] else 1


if __name__=='__main__':
    raise SystemExit(main())
