"""Owned source Heading1 italic-only trial; no font replacement/parity claim.
--compile-only emits AppleScript sources without compiling/executing them.
Root may compile those sources, then separately opt into --execute.
"""
from __future__ import annotations
import argparse
from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET
BASE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('retained_v5_italic',BASE/'native-donor-heading-import-v5.py')
v5=importlib.util.module_from_spec(spec);spec.loader.exec_module(v5)
W=v5.W;ROOT=v5.ROOT;sha=v5.sha;apple_string=v5.apple_string
INPUT=BASE/'native-donor-import-05/heading.docx'
INPUT_SHA='1655223f01b9e500790c2f8dc3fcba0960628e46185da0a9e57f27ec4bd23071'


def paired_styles(root):
    styles=root.findall(W+'style');ids=[s.get(W+'styleId') for s in styles]
    if any(not i for i in ids) or len(set(ids))!=len(ids):raise ValueError('Duplicate/missing style identities')
    source=v5.paragraph_style_by_name(styles,'heading 1')
    if source is None or source.get(W+'type')!='paragraph':raise ValueError('Source missing')
    links=source.findall(W+'link')
    if len(links)!=1 or links[0].attrib!={W+'val':'10'} or len(links[0]):raise ValueError('Wrong source link')
    pair=next(s for s in styles if s.get(W+'styleId')=='10')
    if source.get(W+'styleId')!='1' or pair.get(W+'type')!='character' or pair.get(W+'customStyle')!='1':raise ValueError('Wrong pair type')
    names=pair.findall(W+'name');back=pair.findall(W+'link')
    if len(names)!=1 or names[0].attrib!={W+'val':'标题 1 字符'} or len(back)!=1 or back[0].attrib!={W+'val':'1'}:raise ValueError('Wrong pair identity')
    return source,pair


def italic_off_valid(before,after):
    try:
        expected=ET.fromstring(before);actual=ET.fromstring(after)
        originals=paired_styles(expected);observed=paired_styles(actual)
        for original,changed in zip(originals,observed):
            old=original.findall('./'+W+'rPr/'+W+'i');new=changed.findall('./'+W+'rPr/'+W+'i')
            if len(old)!=1 or old[0].attrib or len(old[0]) or old[0].text or old[0].tail:return False
            # Literal false serialization only. Do not infer omission/inheritance.
            if len(new)!=1 or new[0].attrib not in ({W+'val':'0'},{W+'val':'false'}) or len(new[0]) or new[0].text or new[0].tail:return False
            old[0].set(W+'val',new[0].get(W+'val'))
        return v5.complete_styles_equal(ET.tostring(expected),after)
    except (ValueError,StopIteration,ET.ParseError):return False


def document_signature(xml):
    """Whole body: strict properties, scalar runs, paragraph-local spell markers."""
    root=ET.fromstring(xml);body=root.find(W+'body');result=[]
    if body is None:raise ValueError('Missing body')
    for child in body:
        if child.tag!=W+'p':result.append(v5.property_signature(child));continue
        props=child.findall(W+'pPr')
        if len(props)>1 or (child.text and child.text.strip()):raise ValueError('Malformed paragraph')
        scalars=[];spelling_open=False
        for run in child:
            if run.tail and run.tail.strip():raise ValueError('Hidden paragraph content')
            if run.tag==W+'pPr':continue
            if run.tag==W+'proofErr':
                if len(run) or run.text or set(run.attrib)!={W+'type'}:raise ValueError('Invalid proofing metadata')
                kind=run.get(W+'type')
                if kind=='spellStart' and not spelling_open:spelling_open=True
                elif kind=='spellEnd' and spelling_open:spelling_open=False
                else:raise ValueError('Unbalanced proofing metadata')
                continue
            if run.tag!=W+'r' or (run.text and run.text.strip()) or set(run.attrib)-{W+'rsidR',W+'rsidRPr'}:raise ValueError('Unknown run content')
            rpr=run.findall(W+'rPr')
            if len(rpr)>1:raise ValueError('Duplicate run properties')
            signature=v5.property_signature(rpr[0]) if rpr else None
            for node in run:
                if node.tail and node.tail.strip():raise ValueError('Hidden run content')
                if node.tag==W+'rPr':continue
                if node.tag!=W+'t' or len(node) or set(node.attrib)-{'{http://www.w3.org/XML/1998/namespace}space'}:raise ValueError('Unsupported run node')
                scalars.extend((c,signature) for c in node.text or '')
        if spelling_open:raise ValueError('Unclosed proofing metadata')
        result.append((tuple(sorted((k,v) for k,v in child.attrib.items() if k not in (W+'rsidR',W+'rsidRPr',W+'rsidP'))),
                       v5.property_signature(props[0]) if props else None,tuple(scalars)))
    return (tuple(sorted(body.attrib.items())),body.text or '',tuple(result))


def state_commands():
    return ['set sourceStyle to Word style (style heading1) of boundDoc',
            'set linkedStyle to Word style "标题 1 字符" of boundDoc',
            f'set cloneStyle to Word style {apple_string(v5.CLONE)} of boundDoc',
            'set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc,italic of font object of sourceStyle,italic of font object of linkedStyle,italic of font object of cloneStyle}}']


def valid_state(rows,body,italic):
    return (isinstance(rows,list) and len(rows)==1 and len(rows[0])==9 and rows[0][0]=='state'
            and rows[0][1] is True and rows[0][2]==body and all(type(x) is int for x in rows[0][3:6])
            and rows[0][3:6]==[4,0,0] and rows[0][6] is italic and rows[0][7] is italic and rows[0][8] is True)


def toggle_commands(body,value):
    if type(value) is not bool:raise ValueError('Boolean required')
    prior='false' if value else 'true';target='true' if value else 'false'
    return [f'if not (((current application\'s NSString\'s stringWithString:(content of text object of boundDoc as text))\'s isEqualToString:{apple_string(body)}) as boolean) then error "WPSC_ITALIC_BODY"',
            'set sourceStyle to Word style (style heading1) of boundDoc',
            'set linkedStyle to Word style "标题 1 字符" of boundDoc',
            f'if italic of font object of sourceStyle is not {prior} then error "WPSC_ITALIC_PREIMAGE"',
            f'if italic of font object of linkedStyle is not {prior} then error "WPSC_ITALIC_PAIR_PREIMAGE"',
            f'set italic of font object of sourceStyle to {target}',
            'set nativeRows to {{"italic",italic of font object of sourceStyle,italic of font object of linkedStyle}}']


def package(path):
    with ZipFile(path) as z:return z.read('word/styles.xml'),z.read('word/document.xml')


def run(output):
    sources=[Path(__file__),BASE/'test_italic_face_diagnostic.py',*v5.SOURCES]
    report={'status':'FAIL','scope':'Italic face diagnosis only; visual results are observations, not parity', 'checks':{},'source_hashes':v5.retain_sources(output,sources)}
    owner=reopened=None;sentinel=None;token=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(name,value):
        report['checks'][name]=bool(value);flush()
        if not value:raise AssertionError(name)
    def execute(label,lines):
        report['current_phase']=label;flush();rows=owner._execute(lines);report.setdefault('raw_rows',{})[label]=rows;flush();return rows
    try:
        check('pinned_input',sha(INPUT)==INPUT_SHA);styles,xml=package(INPUT);signature=document_signature(xml)
        paired_styles(ET.fromstring(styles))
        body='\r'.join(''.join(t.text or '' for t in p.iter(W+'t')) for p in ET.fromstring(xml).find(W+'body').findall(W+'p'))+'\r'
        report['expected_body']=body;report['inventory_before']=v5.inventory(output,'before');check('empty_inventory',report['inventory_before']==[])
        report['open_journal']={'source':str(INPUT),'status':'submitted'};flush()
        owner=v5.MacWordSession.open_document(INPUT,read_only=False,visible=False);owner._retain_evidence=True
        report['open_journal'].update(status='acknowledged',private_path=str(owner._private_path));flush()
        check('private_input_exact',sha(owner._private_path)==INPUT_SHA)
        rows=v5.inventory(output,'opened');check('exact_opened_inventory',len(rows)==1 and rows[0][1]==str(owner._private_path))
        token='ITALIC SENTINEL '+uuid4().hex+' 中文😀'
        ack=execute('sentinel',['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])
        if len(ack)!=1 or len(ack[0])!=1 or not isinstance(ack[0][0],str):owner._retain('Uncertain sentinel creation');raise ValueError('Sentinel ACK')
        sentinel=ack[0][0];report['sentinel_before']=v5.sentinel_preimage(v5.inventory(output,'sentinel-before'),sentinel,token);report['sentinel_name']=sentinel
        check('baseline_state',valid_state(execute('baseline',state_commands()),body,True))
        for label,value in [('italic-off',False),('restored',True)]:
            ack=execute(label,toggle_commands(body,value));owner.save_docx(output/(label+'.docx'))
            check(label+'_ack',ack==[['italic',value,value]] and all(type(v) is bool for v in ack[0][1:]))
            current,currentxml=package(output/(label+'.docx'))
            check(label+'_style_tree',v5.complete_styles_equal(styles,current) if value else italic_off_valid(styles,current))
            check(label+'_document_semantics',document_signature(currentxml)==signature)
            check(label+'_theme',v5.theme_hash(output/(label+'.docx'))==v5.theme_hash(INPUT))
            check(label+'_state',valid_state(execute(label+'_state',state_commands()),body,value))
            owner.export_pdf(output/(label+'.pdf'))
            check(label+'_after_pdf_state',valid_state(execute(label+'_after_pdf',state_commands()),body,value))
        owner.close();check('owned_closed',owner._closed and not owner._quarantined)
        v5.close_after_owned(owner,output,report,sentinel,token);sentinel=None
        restored=output/'restored.docx';digest=sha(restored)
        with v5.MacWordSession.open_document(restored,read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True
            check('reopen_state',valid_state(reopened._execute(state_commands()),body,True))
            current,currentxml=package(reopened._private_path)
            check('reopen_styles',v5.complete_styles_equal(styles,current));check('reopen_document',document_signature(currentxml)==signature)
        check('reopen_closed',reopened._closed and not reopened._quarantined);check('restored_bytes',sha(restored)==digest)
        check('original_input_preserved',sha(INPUT)==INPUT_SHA);check('final_empty',v5.inventory(output,'final')==[])
        check('sources_unchanged',all(sha(ROOT/p)==h for p,h in report['source_hashes'].items()))
        report['status']='DIAGNOSTIC_COMPLETE'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)};(output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and not owner._closed:
            try:owner.close()
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        if owner and sentinel and not report.get('sentinel_cleanup_attempted'):
            try:v5.close_after_owned(owner,output,report,sentinel,token);sentinel=None
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        for session,label in ((owner,'native-runtime'),(reopened,'reopen-runtime')):
            if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/label,dirs_exist_ok=True)
        report['remaining_sentinel']=sentinel
        if sentinel or report.get('cleanup_failure'):report['status']='FAIL'
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.pdf')};flush()
    return report


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--compile-only',action='store_true');parser.add_argument('--output',required=True,type=Path);args=parser.parse_args()
    if args.execute==args.compile_only:parser.error('Choose exactly one mode')
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    if args.compile_only:
        from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
        _,xml=package(INPUT);body='\r'.join(''.join(t.text or '' for t in p.iter(W+'t')) for p in ET.fromstring(xml).find(W+'body').findall(W+'p'))+'\r'
        for label,lines in [('off',toggle_commands(body,False)),('restore',toggle_commands(body,True)),('state',state_commands())]:
            (output/(label+'.applescript')).write_text(_JSON+'\ntell application "Microsoft Word"\nset boundDoc to document "COMPILE_ONLY.docx"\n'+'\n'.join(lines)+'\nend tell\nreturn my jsonRows(nativeRows)\n')
        return 0
    return 0 if run(output)['status']=='DIAGNOSTIC_COMPLETE' else 1

if __name__=='__main__':raise SystemExit(main())
