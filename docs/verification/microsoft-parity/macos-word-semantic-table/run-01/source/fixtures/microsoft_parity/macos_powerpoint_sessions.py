"""Bounded native PowerPoint acceptance; explicit clipboard side effects disclosed."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_powerpoint_session import MacPowerPointSession,_JSON


def digest(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def application_snapshot(output):
    script=_JSON+'''with timeout of 20 seconds
 tell application "Microsoft PowerPoint"
  set inventoryRows to {}
  repeat with di from 1 to count of presentations
   set allText to {}
   repeat with si from 1 to count of slides of presentation di
    repeat with sh from 1 to count of shapes of slide si of presentation di
     if has text frame of shape sh of slide si of presentation di then set end of allText to content of text range of text frame of shape sh of slide si of presentation di
    end repeat
   end repeat
   set end of inventoryRows to {name of presentation di,full name of presentation di,saved of presentation di,count of slides of presentation di,allText}
  end repeat
  return my encodeJSON(inventoryRows)
 end tell
end timeout'''
    output.with_suffix('.applescript').write_text(script)
    r=subprocess.run(['/usr/bin/osascript',str(output.with_suffix('.applescript'))],capture_output=True,text=True,timeout=25)
    output.with_suffix('.stdout').write_text(r.stdout);output.with_suffix('.stderr').write_text(r.stderr)
    if r.returncode:raise RuntimeError('Native application inventory failed; raw diagnostics retained')
    data=json.loads(r.stdout);output.write_text(json.dumps(data,ensure_ascii=False,indent=2));return sorted(data,key=lambda d:d[0])


def run(source,output):
    source=Path(source).resolve();output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    before=digest(source);baseline=application_snapshot(output/'before.json');report={'source':str(source),'source_sha256_before':before,'operations':[]}
    session=MacPowerPointSession(source,timeout=120)
    try:
        with session:
            (output/'initial.json').write_text(json.dumps(session.inspect_document(),ensure_ascii=False,indent=2))
            actions=[
              lambda:session.apply_format_patch('slide:1/shape:@name=parity-title',text='Session native edit',font={'name':'Arial','size':26,'bold':True,'italic':True,'underline':True,'color':'#123456'},paragraph={'alignment':2,'space_before':3,'space_after':6,'line_spacing':1.2},geometry={'left':45,'top':50,'width':550,'height':70},fill={'visible':True,'color':'#F1E2D3','transparency':0},line={'color':'#445566','weight':2,'transparency':0},text_frame={'margin_left':8,'margin_right':9,'margin_top':10,'margin_bottom':11,'word_wrap':True,'auto_size':0,'vertical_anchor':3}),
              lambda:session.apply_format_patch('slide:2/shape:@name=parity-table/table/cell:2,2',text='99',font={'bold':True,'color':'#AA1122'}),
              lambda:session.apply_format_patch('slide:1',follow_master_background=False,background={'color':'#EEFFFF','visible':True}),
              lambda:session.apply_format_patch('slide:1/shape:@name=parity-image',geometry={'left':300,'width':120}),
              lambda:session.apply_structural_op({'op':'insert','type':'slide','position':'end','props':{'layout':12}}),
              lambda:session.apply_structural_op({'op':'insert','type':'textbox','parent':'slide:3','props':{'text':'Inserted native session text','left':60,'top':100,'width':500,'height':80}}),
              lambda:session.apply_structural_op({'op':'clone','target':'slide:3/shape:1'}),
              lambda:session.apply_structural_op({'op':'remove','target':'slide:3/shape:2'}),
              lambda:session.apply_structural_op({'op':'move','target':'slide:3','to':'start'}),
              lambda:session.apply_structural_op({'op':'move','target':'slide:1','to':'end'}),
            ]
            for i,action in enumerate(actions):
                result=action();report['operations'].append({'index':i,'result':result});(output/'progress.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
            edited=session.inspect_document();(output/'edited.json').write_text(json.dumps(edited,ensure_ascii=False,indent=2))
            session.save(output/'edited.pptx');session.export_pdf(output/'edited.pdf')
            report['bound_after_pdf']=session.is_bound_to(session._path)
            report['selection']=session.inspect_selection()
        shutil.copytree(session._job,output/'native-job')
        with MacPowerPointSession(output/'edited.pptx',read_only=True,timeout=60) as reopened:
            final=reopened.inspect_document();(output/'reopened.json').write_text(json.dumps(final,ensure_ascii=False,indent=2))
        shutil.copytree(reopened._job,output/'reopen-job')
        assert final['slide_count']==3
        assert final['slides'][0]['shapes'][0]['text']=='Session native edit'
        assert final['slides'][0]['shapes'][0]['font']['size']==26
        title=final['slides'][0]['shapes'][0]
        assert abs(title['geometry']['height']-70)<0.1
        assert title['line']['color']=='#445566'
        assert final['slides'][0]['background']['color']=='#EEFFFF'
        assert final['slides'][1]['shapes'][0]['table']['cells'][3]['text']=='99'
        assert final['slides'][2]['shape_count']==1
        with ZipFile(output/'edited.pptx') as z:
            texts=[''.join(ET.fromstring(z.read(n)).itertext()) for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
        assert any('Session native edit' in t for t in texts)
        assert any('Inserted native session text' in t for t in texts)
        import fitz
        with fitz.open(output/'edited.pdf') as pdf:
            assert len(pdf)==3
            pdf_text='\n'.join(p.get_text() for p in pdf)
            assert 'Session native edit' in pdf_text
            assert 'Inserted native session text' in pdf_text
            for i,p in enumerate(pdf):p.get_pixmap(matrix=fitz.Matrix(1,1)).save(output/f'page-{i+1}.png')
        report['artifact_checks']='PASS';report['status']='PASS'
    except BaseException as e:
        report['status']='FAIL';report['error']=str(e);report['diagnostic_path']=getattr(e,'diagnostic_path',None)
        if session._job and not (output/'native-job').exists():shutil.copytree(session._job,output/'native-job')
        raise
    finally:
        report['source_sha256_after']=digest(source);report['source_unchanged']=before==report['source_sha256_after']
        after=application_snapshot(output/'after.json');report['preexisting_documents_unchanged']=baseline==after
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    assert report['source_unchanged'] and report['preexisting_documents_unchanged']
    return report


def run_semantic(output):
    """Public factory → semantic API → native save/reopen/PDF acceptance."""
    from skills.WPSComposer import create_document
    from skills.WPSComposer.scripts.design_presets import PRESETS
    from skills.WPSComposer.scripts.layout_templates import LayoutTemplate
    output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    baseline=application_snapshot(output/'before.json');report={}
    session=create_document('slide',engine='msoffice',visible=False);session.timeout=180
    try:
        with session:
            assert session.slide_count==0
            session.set_slide_size(960,540);session.apply_design_preset(PRESETS['consultant'])
            assert session.add_title_slide('Semantic native title','Native subtitle')==1
            assert session.add_section_slide('Semantic section')==2
            assert session.add_text_slide('Semantic text','Plain paragraph',bullets=False)==3
            assert session.add_bullets_slide('Semantic bullets',['First native bullet','Second native bullet'])==4
            _,index=session.add_blank_slide();assert index==5
            session.set_background_color(5,'#EEF4FA')
            refs=[session.add_textbox(5,'Native textbox',40,30,300,55,size=24,bold=True,color='#123456',align=2),session.add_shape(5,5,40,115,180,80,fill_color='#226699',line_color='#FFAA00',text='Rounded shape',size=18),session.add_shape(5,9,250,115,90,80,fill_color='#33AA88'),session.add_image(5,ROOT/'docs/verification/microsoft-parity/macos-powerpoint/run-19/source.png',380,115,width=120),session.add_table(5,2,2,40,260,450,130,[['Native','Table'],['Value',42]],header_shade='#334455',header_font='#FFFFFF',font_size=15)]
            session.set_notes(5,'Semantic native notes')
            layout=LayoutTemplate('native','content','static',[{'type':'line','x':40,'y':410,'w':450,'h':4,'color':'#AA4422'},{'type':'box','x':560,'y':80,'w':220,'h':120,'color':'#EECC88','text':'Native box'},{'type':'text','x':560,'y':240,'w':300,'h':65,'text':'Native layout text','fs':22,'color':'#123456','bold':True,'align':2}])
            assert session.apply_layout_template(layout)['ignored_compound_elements']==[]
            (output/'snapshot.json').write_text(json.dumps(session.inspect_document(),ensure_ascii=False,indent=2))
            session.save_pptx(output/'semantic.pptx');session.export_pdf(output/'semantic.pdf');report['returned_refs']=refs
        shutil.copytree(session._job,output/'native-job')
        with MacPowerPointSession(output/'semantic.pptx',read_only=True,timeout=80) as reopened:
            snapshot=reopened.inspect_document();(output/'reopened.json').write_text(json.dumps(snapshot,ensure_ascii=False,indent=2))
            assert snapshot['slide_count']==5 and snapshot['slides'][0]['shapes'][0]['font']['size']==36
            assert snapshot['slides'][4]['notes']=='Semantic native notes' and len(snapshot['slides'][4]['shapes'])==8
        shutil.copytree(reopened._job,output/'reopen-job')
        import fitz
        with fitz.open(output/'semantic.pdf') as pdf:
            assert len(pdf)==5
            pdf_text='\n'.join(page.get_text() for page in pdf)
            for value in ['Semantic native title','Semantic section','Semantic text','Semantic bullets','Native textbox','Native box','Native layout text']:assert value in pdf_text
            for i,page in enumerate(pdf):page.get_pixmap(matrix=fitz.Matrix(1,1)).save(output/f'page-{i+1}.png')
        with ZipFile(output/'semantic.pptx') as archive:
            root=ET.fromstring(archive.read('ppt/slides/slide5.xml'));ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
            assert len(root.findall('.//p:cxnSp',ns))==1 and len(root.findall('.//a:tbl',ns))==1 and len(root.findall('.//p:pic',ns))==1
        report['status']='PASS'
    except BaseException as exc:
        report['status']='FAIL';report['error']=str(exc);report['diagnostic']=getattr(exc,'diagnostic_path',None)
        raise
    finally:
        if session._job and not (output/'native-job').exists():shutil.copytree(session._job,output/'native-job')
        report['preexisting_documents_unchanged']=baseline==application_snapshot(output/'after.json')
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    assert report['preexisting_documents_unchanged']
    return report



def run_save_current(source,output):
    """Explicit original-file saves and hash-conflict rejection on a synthetic copy."""
    source=Path(source).resolve();output=Path(output).resolve();output.mkdir(parents=True,exist_ok=False)
    original_hash=digest(source);baseline=application_snapshot(output/'before.json');report={}
    current=output/'current.pptx';shutil.copy2(source,current)
    sessions=[]
    try:
        session=MacPowerPointSession(current,timeout=120);sessions.append(session)
        with session:
            session.apply_format_patch('slide:1/shape:1',text='Explicit save current native')
            assert session.save_current()==str(current)
            report['first_save_changed_copy']=digest(current)!=original_hash
            session.apply_format_patch('slide:1/shape:1',text='Explicit close save native')
            session.close(save_changes=True)
        with MacPowerPointSession(current,read_only=True,timeout=80) as reopened:
            sessions.append(reopened);snapshot=reopened.inspect_document()
            (output/'reopened.json').write_text(json.dumps(snapshot,ensure_ascii=False,indent=2))
            assert snapshot['slides'][0]['shapes'][0]['text']=='Explicit close save native'
            reopened.export_pdf(output/'current.pdf')
        shutil.copy2(current,output/'saved-current.pptx')
        conflict=MacPowerPointSession(current,timeout=80);sessions.append(conflict)
        with conflict:
            shutil.copy2(source,current);external_hash=digest(current)
            count=conflict._sequence
            try:conflict.preflight_save()
            except ValueError as exc:report['preflight_conflict']=str(exc)
            else:raise AssertionError('Conflict preflight unexpectedly passed')
            try:conflict.save_current()
            except ValueError as exc:report['save_conflict']=str(exc)
            else:raise AssertionError('Conflict save unexpectedly passed')
            assert conflict._sequence==count and digest(current)==external_hash
            report['conflict_rejected_before_native']=True
        assert report['first_save_changed_copy'];report['status']='PASS'
    except BaseException as exc:
        report['status']='FAIL';report['error']=str(exc);raise
    finally:
        for i,session in enumerate(sessions):
            if session._job:shutil.copytree(session._job,output/f'native-job-{i+1}')
        report['original_source_unchanged']=digest(source)==original_hash
        report['preexisting_documents_unchanged']=baseline==application_snapshot(output/'after.json')
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    assert report['original_source_unchanged'] and report['preexisting_documents_unchanged']
    return report


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args()
    print(json.dumps(run(a.source,a.output),ensure_ascii=False,indent=2))
