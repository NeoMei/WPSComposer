"""Public native Word creation gate; preserves unrelated document state."""
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
from skills.WPSComposer import create_document, inspect
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession


def run(output):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'passed': False, 'checks': {}}
    session = None
    try:
        with create_document('writer', engine='msoffice', visible=True) as session:
            report['private_path'] = session._bound_path
            before = session._execute(['set nativeRows to {}',
                'repeat with di from 1 to count documents',
                'set d to document di',
                'if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text, posix full name of d as text, saved of d}',
                'end repeat'])
            session.apply_format_patch('paragraph:1', text='Native public Word creation 中文', font={'size':18,'bold':True})
            session.apply_structural_op({'op':'insert','type':'paragraph','props':{'text':'Editable paragraph created through the public API.'}})
            session.save(output/'created.docx')
            session.export_pdf(output/'created.pdf')
            after = session._execute(['set nativeRows to {}',
                'repeat with di from 1 to count documents',
                'set d to document di',
                'if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text, posix full name of d as text, saved of d}',
                'end repeat'])
            # Never persist unrelated document content in evidence.
            report['checks']['unrelated_document_bindings_and_saved_state_unchanged'] = before == after
            report['checks']['public_native_create_save_export'] = True
            shutil.copytree(session.staging_root, output/'runtime')
        report['checks']['exact_owned_close'] = session._closed
        session = None
        snap = inspect(output/'created.docx', engine='msoffice')
        report['checks']['native_reopen_content'] = [p['text'] for p in snap['paragraphs'] if p['text']] == ['Native public Word creation 中文', 'Editable paragraph created through the public API.']
        report['checks']['native_reopen_font'] = snap['paragraphs'][0]['font']['size'] == 18 and snap['paragraphs'][0]['font']['bold'] is True
        import pdfplumber
        with pdfplumber.open(output/'created.pdf') as pdf:
            report['checks']['native_pdf_text'] = 'Native public Word creation' in '\n'.join(p.extract_text() or '' for p in pdf.pages)
            report['pdf_pages'] = len(pdf.pages)
        report['hashes'] = {name:hashlib.sha256((output/name).read_bytes()).hexdigest() for name in ('created.docx','created.pdf')}
        report['passed'] = all(report['checks'].values())
    except BaseException as error:
        report['error'] = {'type':type(error).__name__, 'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root.exists():
            shutil.copytree(session.staging_root, output/'failed-runtime', dirs_exist_ok=True)
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', required=True, type=Path)
    result = run(parser.parse_args().output_dir)
    print(json.dumps(result,ensure_ascii=False))
    raise SystemExit(0 if result['passed'] else 1)
