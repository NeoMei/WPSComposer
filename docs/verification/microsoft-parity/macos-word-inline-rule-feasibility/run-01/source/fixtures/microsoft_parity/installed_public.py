"""Run a public native smoke gate from an explicitly isolated installed bundle.

Execute this script outside the repository with --plugin pointing to the test
installation. No real personal marketplace or installed plugin is modified.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback


def verify_readback(kind, snapshot, target):
    if kind == 'writer':
        index = int(target.split(':')[1])
        return any(p.get('index') == index and p.get('text') == 'Installed Word edited'
                   for p in snapshot.get('paragraphs', []))
    if kind == 'slide':
        slide, shape = [int(part.split(':')[1]) for part in target.split('/')]
        return any(s.get('index') == slide and any(
            item.get('index') == shape and item.get('text') == 'Installed PowerPoint edited'
            for item in s.get('shapes', [])) for s in snapshot.get('slides', []))
    sheets = [s for s in snapshot.get('sheets', []) if s.get('index') == 1]
    if len(sheets) != 1: return False
    cells = {str(c.get('address', '')).replace('$', '').upper():c for c in sheets[0].get('cells', [])}
    return (cells.get('A2', {}).get('value') == 5 and cells.get('B2', {}).get('value') == 15
            and cells.get('B2', {}).get('formula') == '=A2*3')


def verify_pdf(path, marker):
    from pypdf import PdfReader
    path = Path(path)
    if not path.is_file(): return False
    reader = PdfReader(path)
    return bool(reader.pages) and marker in '\n'.join(page.extract_text() or '' for page in reader.pages)


def capture_close_evidence(session, private, output):
    if private and Path(private).is_dir():
        shutil.copytree(private, Path(output)/'runtime')
    child = getattr(session, '_child', None)
    if child is None:
        # Mac may delete its successful job. This is a returned context exit,
        # never a fabricated worker ACK log.
        return True, {'method':'context_exit_returned', 'native_ack_log_available':False}
    rows = [json.loads(line) for line in (Path(private)/'worker-stdout.log').read_text(encoding='utf-8').splitlines()]
    request = json.loads((Path(private)/'last-request.json').read_text(encoding='utf-8'))
    closed = (bool(rows) and child.poll() == 0 and rows[-1].get('protocol') == 1
        and rows[-1].get('status') == 'ok' and rows[-1].get('value') == {'closed':True}
        and request.get('method') == 'close' and rows[-1].get('id') == request.get('id'))
    return closed, {'python_returncode':child.poll(), 'response':rows[-1] if rows else None}


def run(plugin, output, kind):
    plugin, output = Path(plugin).resolve(strict=True), Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    sys.path.insert(0, str(plugin))
    import skills.WPSComposer as api
    assert Path(api.__file__).resolve().is_relative_to(plugin)
    suffix = {'writer':'docx','sheet':'xlsx','slide':'pptx'}[kind]
    source = output/('installed.'+suffix)
    edited = output/('edited.'+suffix)
    report = {'passed':False,'kind':kind,'installed_plugin':str(plugin),'checks':{}}
    session = None
    try:
        with api.create_document(kind, engine='msoffice') as session:
            if kind == 'writer':
                session.apply_format_patch('paragraph:1', text='Installed native Word', font={'size':16,'bold':True})
                target, patch = 'paragraph:1', {'text':'Installed Word edited'}
            elif kind == 'sheet':
                session.write_cell(1,1,'Installed native Excel')
                session.write_cell(2,1,4)
                session.set_formula(2,2,'=A2*3')
                target, patch = 'sheet:1/cell:A2', {'value':5}
            else:
                index = session.add_title_slide('Installed native PowerPoint','Public plugin API')
                target, patch = f'slide:{index}/shape:1', {'text':'Installed PowerPoint edited'}
            session.save(source)
            session.export_pdf(output/'installed.pdf')
            private = getattr(session,'staging_root',None) or getattr(session,'_job',None)
        closed, close_evidence = capture_close_evidence(session, private, output)
        report['checks']['create_session_close_completed'] = closed
        report['create_session_close_evidence'] = close_evidence
        session = None
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        report['checks']['public_native_created'] = True
        report['snapshot'] = api.inspect(source,engine='msoffice')
        result = api.edit(source,engine='msoffice',patches=[{'target':target,**patch}],output=edited,export_pdf=output/'edited.pdf')
        assert result['ok'] and result['saved'], result
        report['checks']['public_native_edit_saved'] = True
        report['checks']['source_preserved'] = hashlib.sha256(source.read_bytes()).hexdigest() == before
        report['reopened'] = api.inspect(edited,engine='msoffice')
        report['checks']['edited_value_after_native_reopen'] = verify_readback(kind, report['reopened'], target)
        report['checks']['native_pdf_contents'] = (
            verify_pdf(output/'installed.pdf', {'writer':'Installed native Word', 'sheet':'Installed native Excel', 'slide':'Installed native PowerPoint'}[kind])
            and verify_pdf(output/'edited.pdf', {'writer':'Installed Word edited', 'sheet':'Installed native Excel', 'slide':'Installed PowerPoint edited'}[kind]))
        report['hashes'] = {p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in output.iterdir() if p.is_file()}
        report['passed'] = all(report['checks'].values())
    except BaseException as exc:
        report['error'] = {'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session:
            private = getattr(session,'staging_root',None) or getattr(session,'_job',None)
            if private and Path(private).is_dir(): shutil.copytree(private,output/'failed-runtime',dirs_exist_ok=True)
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--plugin',required=True,type=Path)
    parser.add_argument('--output-dir',required=True,type=Path)
    parser.add_argument('--kind',required=True,choices=['writer','sheet','slide'])
    args=parser.parse_args()
    result=run(args.plugin,args.output_dir,args.kind)
    print(json.dumps({'passed':result['passed'],'checks':result['checks'],'error':result.get('error')},ensure_ascii=False))
    raise SystemExit(0 if result['passed'] else 1)
