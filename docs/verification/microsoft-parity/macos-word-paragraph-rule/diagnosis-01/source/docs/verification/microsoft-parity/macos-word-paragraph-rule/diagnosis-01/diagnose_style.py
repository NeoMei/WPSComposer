"""One-off source-bound native style representation diagnosis; no production edits."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT=Path.cwd()
sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage,close_after_owned,PREFIX
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_fields import retain_sources
OUT=Path(__file__).resolve().parent
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SOURCES=[Path(__file__),*[ROOT/p for p in (
    'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_rules.py',
    'fixtures/microsoft_parity/macos_word_paragraph_rule.py',
    'fixtures/microsoft_parity/macos_word_recovery.py',
    'fixtures/microsoft_parity/macos_word_fields.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py')]]

def capture(stage):
    lines=['set nativeRows to {{"stage",'+apple_string(stage)+'}}',
           'repeat with paragraphIndex from 1 to count paragraphs of boundDoc',
           'set captureRange to text object of paragraph paragraphIndex of boundDoc']
    for label,expression in [
        ('content','content of captureRange as text'),
        ('style_equals_object','style of captureRange is Word style (style body text) of boundDoc'),
        ('style_as_text','style of captureRange as text'),
        ('style_name_local','name local of style of captureRange as text'),
        ('bodytext_name_local','name local of Word style (style body text) of boundDoc as text'),
        ('bodytext_as_text','Word style (style body text) of boundDoc as text'),
        ('centered','alignment of paragraph format of captureRange is align paragraph center')]:
        lines += ['try',f'set captureValue to ({expression})',
                  f'set end of nativeRows to {{paragraphIndex,{apple_string(label)},captureValue}}',
                  'on error detail number errorNumber',
                  f'set end of nativeRows to {{paragraphIndex,{apple_string(label)},"ERROR",errorNumber,detail}}','end try']
    return lines+['end repeat']

def saved_xml(name):
    path=OUT/name
    with ZipFile(path) as z:
        xml=z.read('word/document.xml');styles=z.read('word/styles.xml')
    path.with_suffix('.document.xml').write_bytes(xml)
    path.with_suffix('.styles.xml').write_bytes(styles)
    result=[]
    for p in ET.fromstring(xml).findall('./'+W+'body/'+W+'p'):
        prop=p.find(W+'pPr')
        result.append({'text':''.join(n.text or '' for n in p.iter(W+'t')),
                       'pPr':ET.tostring(prop,encoding='unicode') if prop is not None else None})
    return {'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'paragraphs':result}

def main():
    report={'status':'FAIL','checks':{},'source_hashes':retain_sources(OUT,SOURCES),'stages':{}}
    session=None;name=None;token=None
    try:
        report['inventory_before']=inventory(OUT,'before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence=True
            token='Paragraph style diagnosis sentinel 中文😀 '+uuid4().hex
            name=session._execute(['set d to make new document',f'set content of text object of d to {apple_string(token)}','set nativeRows to {{name of d as text}}'])[0][0]
            report['sentinel_before']=sentinel_preimage(inventory(OUT,'sentinel-created'),name,token)
            report['word_version']=session._execute(['set nativeRows to {{version as text}}'])[0][0]
            report['stages']['before_setup']=session._execute(capture('before_setup'))
            session._execute_structural([
                f'set content of text object of boundDoc to {apple_string(PREFIX)}',
                'set beforeRange to text object of paragraph 1 of boundDoc',
                'set style of beforeRange to Word style (style body text) of boundDoc',
                'set space before of paragraph format of beforeRange to 6',
                'set space after of paragraph format of beforeRange to 9',
                'set first line indent of paragraph format of beforeRange to 12',
                'set line spacing rule of paragraph format of beforeRange to line space exactly',
                'set line spacing of paragraph format of beforeRange to 18','set nativeRows to {{"setup",true}}'])
            report['stages']['after_setup']=session._execute(capture('after_setup'))
            session.save_docx(OUT/'setup.docx')
            report['public_result']=session.add_paragraph_horizontal_line()
            report['stages']['after_public_rule']=session._execute(capture('after_public_rule'))
            session.save_docx(OUT/'after-rule.docx')
            session._execute_structural(session._position('end')+[
                'set targetRange to create range boundDoc start insertionPoint end insertionPoint',
                'set content of targetRange to "FOLLOWING"','set nativeRows to {{"following",true}}'])
            report['stages']['after_following_text']=session._execute(capture('after_following_text'))
            session.save_docx(OUT/'after-following.docx')
        close_after_owned(session,OUT,report,name,token);name=None
        report['artifacts']={n:saved_xml(n) for n in ('setup.docx','after-rule.docx','after-following.docx')}
        report['status']='DIAGNOSIS_COMPLETE'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (OUT/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and name and not report.get('sentinel_cleanup_attempted'):
            try:close_after_owned(session,OUT,report,name,token);name=None
            except BaseException:report['cleanup_error']=traceback.format_exc()
        if report.get('sentinel_closed'):name=None
        if session:
            report['quarantined']=session._quarantined
            report['retained_sentinel']=name
            if session.staging_root.exists():shutil.copytree(session.staging_root,OUT/'native-runtime',dirs_exist_ok=True)
        (OUT/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps({'status':report['status'],'error':report.get('error'),'retained_sentinel':name},ensure_ascii=False))
    return 0 if report['status']=='DIAGNOSIS_COMPLETE' else 1

if __name__=='__main__':raise SystemExit(main())
