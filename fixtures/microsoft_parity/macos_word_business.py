"""DRAFT native acceptance for 17 direct Writer methods; never runs on import.

Run only after the host is unlocked and the Office UI/session state is verified.
This draft is not native acceptance evidence. Every output directory must be new.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))



def verify_artifacts(output):
    """Read native DOCX/PDF outputs; never generate or rewrite their contents."""
    from zipfile import ZipFile
    import xml.etree.ElementTree as ET
    import fitz
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    attr='{'+ns['w']+'}'
    with ZipFile(Path(output)/'direct.docx') as package:
        body=ET.fromstring(package.read('word/document.xml'))
        styles=ET.fromstring(package.read('word/styles.xml'))
        paragraphs={''.join(p.itertext()):p for p in body.findall('.//w:body/w:p',ns)}
        byid={s.get(attr+'styleId'):s for s in styles.findall('w:style',ns)}
        for text,level,size in [('Direct Heading One',1,16),('Direct Heading Two',2,15),('Direct Heading Six',6,12)]:
            style=byid[paragraphs[text].find('w:pPr/w:pStyle',ns).get(attr+'val')]
            assert style.find('w:rPr/w:sz',ns).get(attr+'val')==str(size*2)
            assert style.find('w:pPr/w:outlineLvl',ns).get(attr+'val')==str(level-1)
            assert style.find('w:pPr/w:keepNext',ns) is not None
            assert style.find('w:rPr/w:b',ns) is not None
        p=paragraphs['Direct paragraph 中文 😀']
        assert p.find('w:pPr/w:ind',ns).get(attr+'firstLine')=='480'
        assert p.find('w:pPr/w:spacing',ns).get(attr+'line')=='360'
        fonts=p.find('w:pPr/w:rPr/w:rFonts',ns)
        assert fonts.get(attr+'eastAsia')=='仿宋' and fonts.get(attr+'ascii')=='Times New Roman'
        for text,p in paragraphs.items():
            if text.startswith(('→','1.','2.')):
                assert p.find('w:pPr/w:tabs/w:tab',ns).get(attr+'pos')=='480'
                assert p.find('w:pPr/w:ind',ns).get(attr+'hanging')=='480'
        rich=paragraphs['Plain 😀 Bold Code']
        assert any('Bold' in ''.join(r.itertext()) and r.find('w:rPr/w:b',ns) is not None for r in rich.findall('w:r',ns))
        assert any('Code' in ''.join(r.itertext()) and r.find('w:rPr/w:strike',ns) is not None for r in rich.findall('w:r',ns))
        headers=[ET.fromstring(package.read(n)) for n in package.namelist() if n.startswith('word/header') and n.endswith('.xml')]
        assert any(n.get(attr+'sz')=='6' for h in headers for n in h.findall('.//w:pBdr/w:bottom',ns))
        margins=body.find('.//w:sectPr/w:pgMar',ns)
        assert [margins.get(attr+x) for x in ('top','bottom','left','right')]==['1200','1200','1440','1440']
    with fitz.open(Path(output)/'direct.pdf') as pdf:
        spans=[s for p in pdf for b in p.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans']]
        for text,font,size in [('中文','FangSong',12),('Direct paragraph','TimesNewRoman',12),('Code','Consolas',9),('Bold','Bold',12)]:
            assert any(text in s['text'] and font in s['font'] and abs(s['size']-size)<0.2 for s in spans),(text,font)
        first=pdf[0].get_text('words')
        assert any(w[4]=='First' and abs(w[0]-96)<1 for w in first)
        assert any(w[4]=='bullet' and abs(w[0]-96)<1 for w in first)
    return {'native_heading_font_indent_span_list_border':True,'native_pdf_font_and_wrapped_list_geometry':True}


def run(output):
    from skills.WPSComposer import create_document, inspect
    from skills.WPSComposer.scripts.document_model import Span
    from skills.WPSComposer.scripts.msoffice.macos_word_session import NativeWordCapabilityError
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'passed': False, 'checks': {}, 'limitations': [
        'Rich hyperlink spans are explicitly unsupported.',
        'Two-section header linkage is covered separately by the followup fixture.',
        'Consolas uses Latin slots; its CJK glyphs retain the inherited East Asian font.',
    ]}
    session = None
    inventory = ['set nativeRows to {}', 'repeat with di from 1 to count documents',
        'set d to document di',
        'if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text, posix full name of d as text, saved of d}', 'end repeat']
    try:
        with create_document('writer', engine='msoffice', visible=False) as session:
            before = session._execute(inventory)
            session.set_margins(60, 60, 72, 72)
            session.set_page_size(595.3, 841.9)
            session.set_orientation(False)
            session.set_header('First header')
            session.set_footer('First footer')
            session.add_heading('Direct Heading One')
            session.add_heading2('Direct Heading Two')
            session.add_heading_level('Direct Heading Six', level=6)
            session.add_paragraph('Direct paragraph 中文 😀', size=12, font_name='仿宋',
                                  font_name_ascii='Times New Roman', indent_first=24,
                                  line_spacing_rule='one_and_half', space_after=4)
            session.add_centered('Centered direct paragraph')
            session.add_styled_paragraph('Styled normal paragraph', 'Normal')
            session.add_rich_paragraph([Span('Plain 😀 '), Span('Bold', bold=True),
                                       Span(' Code', code=True, strikethrough=True)], 'Body Text')
            session.add_rich_paragraph([Span('First body paragraph')], 'First Paragraph')
            session.add_bullet_list(['First literal bullet', 'Second wrapped literal bullet ' * 5], glyph='→')
            session.add_numbered_list(['First literal number', 'Second literal number'])
            session.add_page_break()
            session.add_paragraph('After direct page break')
            session.set_header_footer(header='Final centered header', footer='Final footer',
                                      link_to_previous_header=False, link_to_previous_footer=False)
            scripts_before = len(list(session.staging_root.glob('*.applescript')))
            try:
                session.add_rich_paragraph([Span('unsupported link', link='https://example.invalid')], 'Body Text')
            except NativeWordCapabilityError:
                report['checks']['unsupported_hyperlink_preflight'] = len(list(session.staging_root.glob('*.applescript'))) == scripts_before
            else:
                raise AssertionError('Unsupported hyperlink was accepted')
            session.save_docx(output/'direct.docx')
            session.export_pdf(output/'direct.pdf')
            report['checks']['unrelated_bindings_saved_state_unchanged'] = before == session._execute(inventory)
            shutil.copytree(session.staging_root, output/'native-runtime')
        report['checks']['exact_owned_close'] = session._closed
        session = None
        source_hash = hashlib.sha256((output/'direct.docx').read_bytes()).hexdigest()
        snapshot = inspect(output/'direct.docx', engine='msoffice')
        assert hashlib.sha256((output/'direct.docx').read_bytes()).hexdigest() == source_hash
        report['checks']['native_reopen_preserves_source_hash'] = True
        (output/'reopened.json').write_text(json.dumps(snapshot, ensure_ascii=False, indent=2))
        texts = [p.get('text', '') for p in snapshot['paragraphs']]
        for text in ['Direct Heading One', 'Direct Heading Two', 'Direct Heading Six',
                     'Direct paragraph 中文 😀', 'Centered direct paragraph', 'Styled normal paragraph',
                     'Plain 😀 Bold Code', 'First body paragraph', 'After direct page break']:
            assert text in texts, text
        assert any(text.startswith('→\tFirst literal bullet') for text in texts)
        assert any(text.startswith('1.\tFirst literal number') for text in texts)
        report['checks']['native_reopen_text_and_list_prefixes'] = True
        import pdfplumber
        with pdfplumber.open(output/'direct.pdf') as pdf:
            report['checks']['native_page_break'] = len(pdf.pages) >= 2
            pdf_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
            assert 'Final centered header' in pdf_text and 'Final footer' in pdf_text
            report['checks']['native_header_footer_pdf'] = True
        report['hashes'] = {name: hashlib.sha256((output/name).read_bytes()).hexdigest() for name in ('direct.docx','direct.pdf')}
        report['checks'].update(verify_artifacts(output))
        # Separate multisection linkage still needs native acceptance.
        report['requires_manual_artifact_checks'] = ['Two-section header/footer linkage']
        report['passed'] = all(report['checks'].values())
        report['acceptance_status'] = 'NATIVE_ARTIFACT_CHECKS_PASSED_WITH_EXPLICIT_LIMITS'
    except BaseException as error:
        report['error'] = {'type': type(error).__name__, 'message': str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root, output/'failed-runtime', dirs_exist_ok=True)
    (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    return report


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path, required=True)
    result = run(parser.parse_args().output_dir)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result['passed'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
