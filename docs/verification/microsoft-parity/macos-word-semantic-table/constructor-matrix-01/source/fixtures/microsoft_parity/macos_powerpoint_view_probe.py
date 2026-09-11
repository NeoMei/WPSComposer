"""Bound native view primitive probe; no enabling assumptions or GUI automation."""
from __future__ import annotations
import argparse,hashlib,json,shutil,sys,traceback
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity.macos_powerpoint_logical_save import opened,inventory,copy_job
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string


def run(output):
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    report={'passed':False,'source_hashes':{str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in [Path(__file__),ROOT/'skills/WPSComposer/scripts/msoffice/macos_powerpoint_session.py']}}
    session=None;sentinel=None
    try:
        with opened(None,create=True) as session:
            session._run('make new slide at end of ownedDoc with properties {layout:slide layout blank}\nmake new slide at end of ownedDoc with properties {layout:slide layout blank}\nreturn "CREATED"',mutation=True)
            token='View sentinel '+uuid4().hex+' 中文😀'
            sentinel=session._run('set d to make new presentation\nset s to make new slide at end of d with properties {layout:slide layout blank}\nset t to make new text box at end of s with properties {left position:30,top:30,width:500,height:40}\nset content of text range of text frame of t to '+apple_string(token)+'\nreturn name of d',mutation=True)
            before=inventory(session);report['inventory_before']=before
            assert any(r[0]==sentinel and r[2] is False for r in before)
            report['sentinel_initially_active']=session._run('return name of active presentation')==sentinel
            body='set results to {}\n'
            for label,command in [
                ('window-index','set index of window (name of ownedDoc) to 1'),
                ('window-visible','set visible of window (name of ownedDoc) to true'),
                ('activate-presentation','activate ownedDoc'),
                ('set-window-active','set active of document window 1 of ownedDoc to true'),
                ('open-exact-private','open (POSIX file '+apple_string(session._path)+' as alias)'),
            ]:
                body+='try\n'+command+'\nset end of results to {'+apple_string(label)+',true,full name of active presentation is full name of ownedDoc}\non error errorText number errorNumber\nset end of results to {'+apple_string(label)+',false,errorNumber}\nend try\n'
            body+='try\ngo to slide (view of document window 1 of ownedDoc) number 2\nset end of results to {"goto-owned",true,slide index of slide of view of document window 1 of ownedDoc}\non error errorText number errorNumber\nset end of results to {"goto-owned",false,errorNumber}\nend try\nreturn my encodeJSON(results)'
            report['results']=json.loads(session._run(body))
            after=inventory(session);report['inventory_after']=after;assert before==after
            session._run('set d to presentation '+apple_string(sentinel)+'\nif saved of d then error "SENTINEL_SAVED"\nif content of text range of text frame of shape 1 of slide 1 of d is not '+apple_string(token)+' then error "SENTINEL_CHANGED"\nclose d saving no\nreturn "CLOSED"',mutation=True);sentinel=None
            copy_job(session,output,'native-job')
            report['passed']=report['sentinel_initially_active'] and any(r[:3]==['goto-owned',True,2] for r in report['results'])
        report['exact_owned_close']=not session._entered
    except BaseException as e:
        report['error']={'type':type(e).__name__,'message':str(e)};(output/'failure.txt').write_text(traceback.format_exc());copy_job(session,output,'failed-job')
        if sentinel:report['retained_sentinel']=sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',required=True);a=p.parse_args()
    if not a.execute:p.error('--execute required')
    r=run(a.output);print(json.dumps({'passed':r['passed'],'results':r.get('results'),'error':r.get('error')}));sys.exit(0 if r['passed'] else 1)
