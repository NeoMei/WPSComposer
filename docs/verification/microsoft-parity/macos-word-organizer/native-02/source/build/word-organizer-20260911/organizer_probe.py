"""Owned native single-style copy probe; no public capability or generic flattening."""
from __future__ import annotations
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile,ZIP_DEFLATED
import xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer import open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage
from fixtures.microsoft_parity.macos_word_quality_feasibility import close_sentinel_after_owned
SOURCE=ROOT/'build/word-figure-rollback-20260911/native-v3-01/figure-rollback-source.docx'
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NAME='WPSC Organizer Heading';ID='WPSCOrganizerHeading'

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def canonical(node):
    return node.tag,tuple(sorted(node.attrib.items())),node.text or '',tuple(canonical(c) for c in node)
def target_style(root):
    nodes=[s for s in root.findall(W+'style') if s.get(W+'styleId')==ID or (s.find(W+'name') is not None and s.find(W+'name').get(W+'val')==NAME)]
    if len(nodes)>1:raise ValueError('duplicate target styles')
    return nodes[0] if nodes else None

def make_input(source,target,size):
    if size not in (17,23):raise ValueError('closed fixture size')
    with ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
    root=ET.fromstring(parts['word/styles.xml'])
    if target_style(root) is not None:raise ValueError('source already has target')
    normal=[s for s in root.findall(W+'style') if s.get(W+'type')=='paragraph' and s.get(W+'default')=='1']
    if len(normal)!=1:raise ValueError('expected one Normal paragraph style')
    node=ET.SubElement(root,W+'style',{W+'type':'paragraph',W+'customStyle':'1',W+'styleId':ID})
    for tag,value in [('name',NAME),('basedOn',normal[0].get(W+'styleId')),('next',normal[0].get(W+'styleId'))]:ET.SubElement(node,W+tag,{W+'val':value})
    p=ET.SubElement(node,W+'pPr');ET.SubElement(p,W+'keepNext');ET.SubElement(p,W+'keepLines')
    ET.SubElement(p,W+'spacing',{W+'before':str(size*10),W+'after':'80'})
    ET.SubElement(p,W+'ind',{W+'left':'340'});ET.SubElement(p,W+'outlineLvl',{W+'val':'0'})
    r=ET.SubElement(node,W+'rPr');ET.SubElement(r,W+'b');ET.SubElement(r,W+'sz',{W+'val':str(size*2)});ET.SubElement(r,W+'szCs',{W+'val':str(size*2)})
    original_styles=parts['word/styles.xml']
    if original_styles.count(b'</w:styles>') != 1:
        raise ValueError('Unexpected owned native styles root encoding')
    # Preserve all source namespace declarations, including mc:Ignorable prefixes.
    parts['word/styles.xml']=original_styles.replace(b'</w:styles>',ET.tostring(node)+b'</w:styles>')
    with ZipFile(target,'x',compression=ZIP_DEFLATED) as z:
        for n,data in parts.items():z.writestr(n,data)

def package(path):
    with ZipFile(path) as z:return z.read('word/document.xml'),z.read('word/styles.xml'),z.read('word/theme/theme1.xml')

def style_signature(node):
    node=deepcopy(node)
    for child in node.findall(W+'rsid'):node.remove(child)
    return canonical(node)

def other_styles_equal(before,after):
    roots=[ET.fromstring(x) for x in (before,after)]
    for root in roots:
        target=target_style(root)
        if target is not None:root.remove(target)
    return canonical(roots[0])==canonical(roots[1])

def state_commands():
    return [f'set probeStyle to Word style {apple_string(NAME)} of boundDoc','set nativeRows to {{"style",name local of probeStyle as text,font size of font object of probeStyle,paragraph format left indent of paragraph format of probeStyle,space before of paragraph format of probeStyle}}']

def copy_commands(source):
    return [f'organizer copy source {apple_string(str(source))} destination (posix full name of boundDoc as text) name {apple_string(NAME)} organizer object type organizer object styles',*state_commands()]

def run(out):
    out=Path(out).resolve();out.mkdir(exist_ok=False)
    sources=[Path(__file__),Path(__file__).with_name('test_organizer_probe.py'),ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',ROOT/'fixtures/microsoft_parity/macos_word_fields.py',ROOT/'fixtures/microsoft_parity/macos_word_recovery.py',ROOT/'fixtures/microsoft_parity/macos_word_paragraph_rule.py',ROOT/'fixtures/microsoft_parity/macos_word_quality_feasibility.py']
    report={'status':'FAIL','checks':{},'scope':'single named style install and same-name update on owned native DOCX only','source_hashes':retain_sources(out,sources),'input_source_sha256':sha(SOURCE)}
    owners=[];owner=None;sentinel=token=None
    def flush():(out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    def check(name,value):
        report['checks'][name]=bool(value);flush()
        if not value:raise AssertionError(name)
    try:
        report['inventory_before']=inventory(out,'before');flush()
        donors=[]
        for size in (17,23):
            input_path=out/f'input-{size}.docx';make_input(SOURCE,input_path,size);digest=sha(input_path)
            with open_document(input_path,engine='msoffice',visible=False) as donor:
                donor._retain_evidence=True;owners.append((donor,f'donor-{size}-runtime'))
                rows=donor._execute(state_commands());report[f'donor_{size}_rows']=rows
                check(f'donor_{size}_native',[['style',NAME,size,17,size/2]]==rows)
                saved=out/f'donor-{size}.docx';donor.save_copy(saved)
                check(f'donor_{size}_input_preserved',sha(input_path)==digest)
                donors.append(saved)
        with open_document(SOURCE,engine='msoffice',visible=False) as owner:
            owner._retain_evidence=True;owners.append((owner,'recipient-runtime'))
            token='ORGANIZER SENTINEL '+uuid4().hex
            rows=owner._execute(['set ownSentinel to make new document',f'set content of text object of ownSentinel to {apple_string(token)}','set nativeRows to {{name of ownSentinel as text}}'])
            sentinel=rows[0][0];report['sentinel_name']=sentinel
            report['sentinel_before']=sentinel_preimage(inventory(out,'sentinel-before'),sentinel,token);flush()
            owner._execute(['set selection start of selection of boundWindow to 10','set selection end of selection of boundWindow to 21','set nativeRows to {{"positioned",true}}'])
            before=out/'recipient-before.docx';owner.save_copy(before);before_parts=package(before)
            for saved,size in zip(donors,(17,23)):
                staged=owner.staging_root/saved.name;shutil.copy2(saved,staged)
                report['current_step']=f'organizer-{size}';flush()
                report[f'copy_{size}_rows']=owner._execute(copy_commands(staged));flush()
                check(f'copy_{size}_live_style',report[f'copy_{size}_rows']==[['style',NAME,size,17,size/2]])
                # Read the private file before the explicit native save separately.
                disk=out/f'recipient-{size}-before-save.docx';shutil.copy2(owner._private_path,disk)
                disk_style=target_style(ET.fromstring(package(disk)[1]));report[f'disk_before_save_{size}_has_style']=disk_style is not None
                target=out/f'recipient-{size}.docx';owner.save_copy(target);parts=package(target)
                check(f'copy_{size}_body_exact',parts[0]==before_parts[0])
                check(f'copy_{size}_other_styles',other_styles_equal(before_parts[1],parts[1]))
                check(f'copy_{size}_theme',parts[2]==before_parts[2])
                expected=target_style(ET.fromstring(package(saved)[1]));actual=target_style(ET.fromstring(parts[1]))
                check(f'copy_{size}_definition',actual is not None and style_signature(actual)==style_signature(expected))
                selection=owner._execute(['set r to text object of selection of boundWindow','set nativeRows to {{start of content of r,end of content of r,content of r as text}}'])
                check(f'copy_{size}_selection',selection==[[10,21,'REPLACE-长😀']])
            report['sentinel_after']=sentinel_preimage(inventory(out,'sentinel-after'),sentinel,token)
            check('sentinel_preserved',report['sentinel_before']==report['sentinel_after'])
        close_sentinel_after_owned(owner,out,report,sentinel,token);sentinel=None
        with open_document(out/'recipient-23.docx',engine='msoffice',read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True;owners.append((reopened,'reopen-runtime'))
            check('reopen_style',reopened._execute(state_commands())==[['style',NAME,23,17,11.5]])
        check('source_preserved',sha(SOURCE)==report['input_source_sha256'])
        report['status']='PASS'
    except BaseException as e:
        report['error']={'type':type(e).__name__,'message':str(e)}
        (out/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and sentinel:
            try:close_sentinel_after_owned(owner,out,report,sentinel,token);sentinel=None
            except BaseException:report['cleanup_error']=traceback.format_exc()
        for session,label in owners:
            if session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,out/label)
        report['remaining_sentinel']=sentinel
        report['source_unchanged']=all(sha(ROOT/n)==h for n,h in report['source_hashes'].items())
        report['input_source_unchanged']=sha(SOURCE)==report['input_source_sha256']
        report['inventory_final']=inventory(out,'final')
        if report['inventory_final']!=report.get('inventory_before') or not report['source_unchanged'] or not report['input_source_unchanged'] or sentinel or report.get('cleanup_error'):report['status']='FAIL'
        report['artifact_hashes']={p.name:sha(p) for p in out.glob('*.docx')}
        flush()
    return report

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
    if not args.execute:raise SystemExit('Requires --execute and root native Word lease')
    raise SystemExit(0 if run(args.output)['status']=='PASS' else 1)
