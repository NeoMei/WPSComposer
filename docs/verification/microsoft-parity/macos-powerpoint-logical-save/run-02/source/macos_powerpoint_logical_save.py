"""Guarded public PowerPoint logical-save acceptance; --execute required."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import time
import traceback
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from skills.WPSComposer import create_document,open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string


def inventory(session):
    # Only hashes of unrelated document text leave native process.
    body='''set inventoryRows to {}
repeat with di from 1 to count presentations
 set d to presentation di
 if full name of d is not full name of ownedDoc then
  set allText to ""
  repeat with si from 1 to count slides of d
   repeat with sh from 1 to count shapes of slide si of d
    if has text frame of shape sh of slide si of d then set allText to allText & (content of text range of text frame of shape sh of slide si of d as text) & return
   end repeat
  end repeat
  set textHash to do shell script ("/usr/bin/printf %s " & quoted form of allText & " | /usr/bin/shasum -a 256")
  set end of inventoryRows to {name of d,full name of d,saved of d,count slides of d,textHash}
 end if
end repeat
return my encodeJSON(inventoryRows)'''
    return json.loads(session._run(body))


def text(session):
    snapshot=session.inspect_document()
    return [shape.get('text','') for slide in snapshot['slides'] for shape in slide['shapes']]


def copy_job(session,output,name):
    if session and session._job and session._job.exists():shutil.copytree(session._job,output/name,dirs_exist_ok=True)


def opened(path,*,create=False):
    s=create_document('slide',engine='msoffice',visible=False) if create else open_document(path,engine='msoffice',visible=False)
    s.timeout=180
    return s


def run(output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    sources=[Path(__file__),ROOT/'skills/WPSComposer/scripts/msoffice/macos_powerpoint_session.py',ROOT/'skills/WPSComposer/scripts/artifact_transport.py']
    report={'passed':False,'checks':{},'source_hashes':{str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}}
    (output/'source').mkdir()
    for p in sources:shutil.copy2(p,output/'source'/p.name)
    session=None;sentinel=None
    try:
        with opened(None,create=True) as session:
            token='Logical save sentinel '+uuid4().hex+' 中文😀'
            sentinel=session._run('''set sentinelDoc to make new presentation
set sentinelSlide to make new slide at end of sentinelDoc with properties {layout:slide layout blank}
set sentinelShape to make new text box at end of sentinelSlide with properties {left position:30,top:30,width:500,height:50}
set content of text range of text frame of sentinelShape to '''+apple_string(token)+'''
return name of sentinelDoc''',mutation=True)
            before=inventory(session);assert any(r[0]==sentinel and r[2] is False for r in before)
            report['inventory_before']=before
            try:
                try:session.save_current()
                except ValueError as e:assert 'explicit' in str(e)
                else:raise AssertionError('New save_current accepted without destination')
                _,slide=session.add_blank_slide();target=session.add_textbox(slide,'NEW INITIAL',30,30,400,50)
                private=session._path
                session.save(output/'new-B.pptx');session.save_copy(output/'new-C.pptx')
                session.apply_format_patch(target,text='NEW LATER EDIT')
                assert session.save_current()==str(output/'new-B.pptx')
                assert session._path==private
                session.save_copy(output/'A.pptx')
                assert session._logical==output/'new-B.pptx'
                report['checks']['new_save_as_copy_then_current']=True
                assert inventory(session)==before
                copy_job(session,output,'new-job')
            finally:
                pass  # Synthetic unsaved sentinel spans every lifecycle case.
        report['checks']['new_exact_close']=not session._entered
        a=output/'A.pptx';digest=hashlib.sha256(a.read_bytes()).hexdigest()
        with opened(a) as session:
            private=session._path
            session.save(output/'B.pptx');session.save_copy(output/'C.pptx')
            session.apply_format_patch('slide:1/shape:1',text='OPEN LATER EDIT')
            assert session.save_current()==str(output/'B.pptx')
            assert session._path==private
            assert hashlib.sha256(a.read_bytes()).hexdigest()==digest
            report['checks']['open_save_as_copy_then_current']=True
            assert inventory(session)==before
            copy_job(session,output,'open-job')
        report['checks']['open_exact_close']=not session._entered
        expected={'A.pptx':'NEW LATER EDIT','B.pptx':'OPEN LATER EDIT','C.pptx':'NEW LATER EDIT','new-B.pptx':'NEW LATER EDIT','new-C.pptx':'NEW INITIAL'}
        report['reopen_text']={}
        for name,expected_text in expected.items():
            path=output/name;original=hashlib.sha256(path.read_bytes()).hexdigest()
            with opened(path) as session:
                actual=text(session);assert actual==[expected_text];report['reopen_text'][name]=actual
                assert inventory(session)==before
                copy_job(session,output,'reopen-'+name)
            assert not session._entered and hashlib.sha256(path.read_bytes()).hexdigest()==original
        report['checks']['native_reopen_A_B_C_new_B_new_C']=True
        report['checks']['original_A_preserved']=hashlib.sha256(a.read_bytes()).hexdigest()==digest
        with opened(a) as session:
            session.save(output/'conflict-B.pptx');session.apply_format_patch('slide:1/shape:1',text='UNPUBLISHED EDIT')
            conflict=output/'conflict-B.pptx';shutil.copyfile(output/'new-C.pptx',conflict);changed=hashlib.sha256(conflict.read_bytes()).hexdigest();seq=session._sequence
            try:session.save_current()
            except ValueError as e:assert 'changed' in str(e)
            else:raise AssertionError('Concurrent destination edit overwritten')
            assert seq==session._sequence and hashlib.sha256(conflict.read_bytes()).hexdigest()==changed
            report['checks']['conflict_before_native_preserves_changed_destination']=True
            session.save(output/'race-B.pptx')
            session.apply_format_patch('slide:1/shape:1',text='RACING UNPUBLISHED EDIT')
            race=output/'race-B.pptx';prior_state=session._logical_state
            real_validate=session._validate_artifact
            def racing_validate(path,fmt):
                real_validate(path,fmt)
                if Path(path)!=Path(session._path):shutil.copyfile(output/'new-C.pptx',race)
            session._validate_artifact=racing_validate
            try:
                try:session.save_current()
                except RuntimeError as e:assert 'changed' in str(e)
                else:raise AssertionError('Post-staging destination race overwritten')
            finally:session._validate_artifact=real_validate
            assert race.read_bytes()==(output/'new-C.pptx').read_bytes() and session._logical_state==prior_state
            report['checks']['conflict_after_native_staging_preserves_changed_destination']=True
            assert inventory(session)==before
            report['checks']['sentinel_and_unrelated_documents_preserved_across_cases']=True
            session._run('set d to presentation '+apple_string(sentinel)+'\nif saved of d then error "SENTINEL_SAVED"\nif content of text range of text frame of shape 1 of slide 1 of d is not '+apple_string(token)+' then error "SENTINEL_CHANGED"\nclose d saving no\nreturn "CLOSED"',mutation=True);sentinel=None
            copy_job(session,output,'conflict-job')
        report['checks']['conflict_exact_close']=not session._entered
        report['artifact_hashes']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in output.glob('*.pptx')}
        report['passed']=all(report['checks'].values())
    except BaseException as error:
        report['error']={'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc());copy_job(session,output,'failed-job')
        if sentinel:report['retained_sentinel']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',required=True);a=p.parse_args()
    if not a.execute:p.error('--execute required')
    r=run(a.output);print(json.dumps({'passed':r['passed'],'error':r.get('error')}));sys.exit(0 if r['passed'] else 1)
