"""Guarded public Word logical save acceptance; never runs on import."""
from __future__ import annotations
import argparse,hashlib,json,shutil,sys,traceback
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document,open_document
from skills.WPSComposer.scripts.msoffice import macos_word_session as module
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_sections import _inventory_commands


def inventory(session):return session._execute(_inventory_commands())
def text(session):return session._execute(['set nativeRows to {{content of text object of boundDoc as text}}'])[0][0]
def copy_job(session,output,name):
    if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/name,dirs_exist_ok=True)


def run(output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    sources=[Path(__file__),ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',ROOT/'skills/WPSComposer/scripts/artifact_transport.py',ROOT/'fixtures/microsoft_parity/macos_word_sections.py']
    report={'passed':False,'checks':{},'source_hashes':{str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}}
    for p in sources:
        target=output/'source'/p.relative_to(ROOT);target.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(p,target)
        assert hashlib.sha256(target.read_bytes()).hexdigest()==report['source_hashes'][str(p.relative_to(ROOT))]
    session=None;sentinel=None
    try:
        with create_document('writer',engine='msoffice',visible=False) as session:
            token='Logical Word sentinel '+uuid4().hex+' 中文😀'
            sentinel=session._execute(['set d to make new document',f'set content of text object of d to {apple_string(token)}','set nativeRows to {{name of d as text}}'])[0][0]
            before=inventory(session);report['inventory_before']=before
            assert any(r[0]==sentinel and r[2] is False for r in before)
            seq=len(list(session.staging_root.glob('*.applescript')))
            try:session.save_current()
            except ValueError as e:assert 'explicit' in str(e)
            else:raise AssertionError('New current save accepted without destination')
            assert len(list(session.staging_root.glob('*.applescript')))==seq
            session.add_paragraph('NEW INITIAL 😀')
            private=session._bound_path
            session.save(output/'new-B.docx');session.save_copy(output/'new-C.docx');session.export_pdf(output/'copy.pdf')
            session.add_paragraph('NEW LATER EDIT')
            assert session.save_current()==str(output/'new-B.docx') and session._bound_path==private
            session.save_copy(output/'A.docx')
            assert session._logical_path==output/'new-B.docx'
            report['checks']['new_save_as_copy_pdf_then_current']=True
            assert inventory(session)==before;copy_job(session,output,'new-job')
        report['checks']['new_exact_close']=session._closed
        a=output/'A.docx';original=hashlib.sha256(a.read_bytes()).hexdigest()
        with open_document(a,engine='msoffice',visible=False) as session:
            private=session._bound_path
            session.save(output/'B.docx');session.save_copy(output/'C.docx');session.add_paragraph('OPEN LATER EDIT')
            assert session.save_current()==str(output/'B.docx') and session._bound_path==private
            assert hashlib.sha256(a.read_bytes()).hexdigest()==original
            report['checks']['open_save_as_copy_then_current']=True
            assert inventory(session)==before;copy_job(session,output,'open-job')
        report['checks']['open_exact_close']=session._closed
        expected={'A.docx':['NEW INITIAL 😀','NEW LATER EDIT'],'B.docx':['NEW INITIAL 😀','NEW LATER EDIT','OPEN LATER EDIT'],'C.docx':['NEW INITIAL 😀','NEW LATER EDIT'],'new-B.docx':['NEW INITIAL 😀','NEW LATER EDIT'],'new-C.docx':['NEW INITIAL 😀']}
        report['reopen_text']={}
        for name,wanted in expected.items():
            p=output/name;digest=hashlib.sha256(p.read_bytes()).hexdigest()
            with open_document(p,engine='msoffice',read_only=True,visible=False) as session:
                actual=[x for x in text(session).split('\r') if x];assert actual==wanted;report['reopen_text'][name]=actual
                assert inventory(session)==before;copy_job(session,output,'reopen-'+p.stem)
            assert session._closed and hashlib.sha256(p.read_bytes()).hexdigest()==digest
        report['checks']['native_reopen_A_B_C_new_B_new_C']=True
        report['checks']['original_A_preserved']=hashlib.sha256(a.read_bytes()).hexdigest()==original
        with open_document(a,engine='msoffice',visible=False) as session:
            session.save(output/'conflict-B.docx');session.add_paragraph('UNPUBLISHED EDIT')
            conflict=output/'conflict-B.docx';shutil.copyfile(output/'new-C.docx',conflict);changed=hashlib.sha256(conflict.read_bytes()).hexdigest();seq=len(list(session.staging_root.glob('*.applescript')))
            try:session.save_current()
            except ValueError as e:assert 'changed' in str(e)
            else:raise AssertionError('Concurrent destination change overwritten')
            assert seq==len(list(session.staging_root.glob('*.applescript'))) and hashlib.sha256(conflict.read_bytes()).hexdigest()==changed
            report['checks']['conflict_before_native_preserves_changed_bytes']=True
            session.save(output/'race-B.docx');session.add_paragraph('RACING UNPUBLISHED EDIT');race=output/'race-B.docx';prior=session._logical_state
            real_validate=module.validate_before_deadline
            def racing_validate(spec,path,deadline):
                real_validate(spec,path,deadline)
                if Path(path)!=Path(session._bound_path):shutil.copyfile(output/'new-C.docx',race)
            module.validate_before_deadline=racing_validate
            try:
                try:session.save_current()
                except RuntimeError as e:assert 'changed' in str(e)
                else:raise AssertionError('Post-staging destination race overwritten')
            finally:module.validate_before_deadline=real_validate
            assert race.read_bytes()==(output/'new-C.docx').read_bytes() and session._logical_state==prior and session._retain_evidence
            report['checks']['conflict_after_staging_preserves_changed_bytes_and_recovery']=True
            assert inventory(session)==before;report['checks']['nonempty_unsaved_sentinel_and_unrelated_docs_preserved']=True
            session._execute([f'set d to document {apple_string(sentinel)}',f'if (content of text object of d as text) is not {apple_string(token)} & return then error "SENTINEL_CHANGED"','if saved of d then error "SENTINEL_SAVED"','close d saving no','set nativeRows to {{"ok"}}']);sentinel=None
            copy_job(session,output,'conflict-job')
        report['checks']['conflict_exact_close']=session._closed
        report['artifact_hashes']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in output.glob('*') if p.is_file() and p.suffix in {'.docx','.pdf'}}
        report['passed']=all(report['checks'].values())
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)};(output/'failure.txt').write_text(traceback.format_exc());copy_job(session,output,'failed-job')
        if sentinel:report['retained_sentinel']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',required=True);a=p.parse_args()
    if not a.execute:p.error('--execute required')
    r=run(a.output);print(json.dumps({'passed':r['passed'],'error':r.get('error')}));sys.exit(0 if r['passed'] else 1)
