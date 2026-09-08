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




def _field_commands(root):
    """Return exact instruction command tokens, requiring balanced complex fields."""
    w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    commands=[];stack=[]
    def command(instruction):
        tokens=instruction.split()
        assert tokens, 'Field instruction is empty'
        return tokens[0].upper()
    for node in root.iter():
        if node.tag==w+'fldSimple':
            commands.append(command(node.get(w+'instr','')))
        elif node.tag==w+'fldChar':
            kind=node.get(w+'fldCharType')
            if kind=='begin':
                stack.append({'instruction':[],'separated':False})
            elif kind=='separate':
                assert stack and not stack[-1]['separated'], 'Unbalanced field separator'
                stack[-1]['separated']=True
            elif kind=='end':
                assert stack and stack[-1]['separated'], 'Unbalanced field end'
                field=stack.pop()
                commands.append(command(''.join(field['instruction'])))
            else:
                raise AssertionError('Unknown field delimiter')
        elif node.tag==w+'instrText':
            assert stack and not stack[-1]['separated'], 'Unbalanced field instruction'
            stack[-1]['instruction'].append(node.text or '')
    assert not stack, 'Unbalanced field begin'
    return commands


def _verify_page_fields(package, body):
    """Require PAGE in every effective default footer, following section inheritance."""
    import posixpath
    w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    rel_type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
    assert 'PAGE' not in _field_commands(body), 'PAGE field must not enter body story'
    relationships=ET.fromstring(package.read('word/_rels/document.xml.rels'))
    byid={}
    for rel in relationships:
        key=rel.get('Id')
        assert key and key not in byid, 'Invalid duplicate document relationship'
        byid[key]=rel
    effective=None;resolved=[]
    for section in body.iter(w+'sectPr'):
        references=[ref for ref in section.findall(w+'footerReference') if ref.get(w+'type')=='default']
        assert len(references)<=1, 'Multiple default footer references'
        if references:
            effective=references[0].get(r+'id')
        assert effective in byid, 'Effective default footer relationship missing'
        rel=byid[effective]
        assert rel.get('Type')==rel_type and rel.get('TargetMode','Internal')=='Internal', 'Invalid default footer relationship'
        target=rel.get('Target','')
        assert target and ':' not in target and '\\' not in target, 'Invalid footer target'
        target=posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('word',target))
        assert target.startswith('word/') and target in package.namelist(), 'Default footer part missing'
        footer=ET.fromstring(package.read(target))
        assert footer.tag==w+'ftr', 'Default footer target is not a footer'
        assert 'PAGE' in _field_commands(footer), 'Effective default footer requires a PAGE field'
        resolved.append(target)
    assert resolved, 'Document sections missing'
    return resolved


def verify_artifacts(output):
    """Check native style/section/field structures and rendered column geometry."""
    import fitz
    output=Path(output);ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};a='{'+ns['w']+'}'
    with ZipFile(output/'followup.docx') as package:
        body=ET.fromstring(package.read('word/document.xml'))
        _verify_page_fields(package, body)
        assert 'Followup native body' in ''.join(body.itertext())
        sections=body.findall('.//w:sectPr',ns)
        assert len(sections)==3
        assert [s.find('w:pgSz',ns).get(a+'orient','portrait') for s in sections]==['portrait','landscape','portrait']
        assert all(s.find('w:cols',ns).get(a+'num')=='2' for s in sections)
        styles=ET.fromstring(package.read('word/styles.xml'))
        byname={s.find('w:name',ns).get(a+'val'):s for s in styles.findall('w:style',ns)}
        for level,size in enumerate([16,15,15,14,14,12],1):
            h=byname[f'heading {level}']
            assert h.find('w:rPr/w:sz',ns).get(a+'val')==str(size*2)
            assert h.find('w:rPr/w:color',ns).get(a+'val')=='224466'
            assert h.find('w:pPr/w:keepNext',ns) is not None
        code=byname['Source Code']
        assert code.find('w:rPr/w:rFonts',ns).get(a+'ascii')=='Consolas'
        assert code.find('w:rPr/w:sz',ns).get(a+'val')=='18'
        assert code.find('w:pPr/w:ind',ns).get(a+'left')=='720'
        assert code.find('w:pPr/w:shd',ns).get(a+'fill')=='F5F5F5'
        custom=byname['Custom Body']
        assert custom.find('w:pPr/w:ind',ns).get(a+'firstLine')=='480'
        assert custom.find('w:rPr/w:color',ns).get(a+'val')=='123456'
        tail=body.findall('w:body/w:p',ns)[-1]
        assert tail.find('w:pPr/w:rPr/w:sz',ns).get(a+'val')=='2'
        assert tail.find('w:pPr/w:spacing',ns).get(a+'line')=='20'
        assert tail.find('w:pPr/w:spacing',ns).get(a+'lineRule')=='exact'
        # The third section inherits section two's header by omitting a new reference.
        assert sections[1].find('w:headerReference',ns) is not None
        assert sections[2].find('w:headerReference',ns) is None
    with fitz.open(output/'followup.pdf') as pdf:
        assert len(pdf)==3
        for index,page in enumerate(pdf,1):
            assert f'Page {index}' in page.get_text(), (index,page.get_text())
            assert ('Section One' if index==1 else 'Section Two') in page.get_text()
        words=pdf[0].get_text('words')
        assert any(w[4]=='Two-column' and w[0]<150 for w in words)
        assert any(w[4]=='Two-column' and w[0]>300 for w in words)
        spans=[s for b in pdf[0].get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans']]
        assert any('native_code' in s['text'] and s['font']=='Consolas' and abs(s['size']-9)<0.2 for s in spans)
    return {'native_style_columns_sections_compaction':True,'native_page_field_results_and_header_linkage':True}


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
            session.set_header('Section One')
            session.add_section(True,continuous=False)
            session.set_header_footer(header='Section Two',link_to_previous_header=False,link_to_previous_footer=True)
            session.add_paragraph('Landscape followup section')
            session.add_section(False,continuous=True)
            session.set_header_footer(link_to_previous_header=True,link_to_previous_footer=True)
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
        source_hash=hashlib.sha256((output/'followup.docx').read_bytes()).hexdigest()
        snap=inspect(output/'followup.docx',engine='msoffice')
        assert hashlib.sha256((output/'followup.docx').read_bytes()).hexdigest()==source_hash
        report['checks']['native_reopen_preserves_source_hash']=True
        (output/'reopened.json').write_text(json.dumps(snap,ensure_ascii=False,indent=2))
        texts=[p.get('text','') for p in snap['paragraphs']]
        assert 'native_code = 1' in texts and 'Landscape followup section' in texts
        assert snap['counts']['sections']==3
        report['checks']['native_reopen_text_sections']=True
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        with ZipFile(output/'followup.docx') as package:
            body=ET.fromstring(package.read('word/document.xml'))
            columns=body.findall('.//w:sectPr/w:cols',ns)
            assert len(columns)==3 and all(node.get('{'+ns['w']+'}num')=='2' for node in columns)
            report['checks']['native_two_column_sections']=True
            _verify_page_fields(package, body)
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
        report['checks'].update(verify_artifacts(output))
        report['passed']=all(report['checks'].values())
        report['acceptance_status']='NATIVE_ARTIFACT_CHECKS_PASSED_WITH_EXPLICIT_LIMITS'
        report['requires_artifact_review']=['Nonempty TOC/TOF refresh requires a separate seeded document']
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
