"""Gate A only: explicit Flat OPC converter open -> Word-saved DOCX.

Requires empty Word inventory. Every native document path is UUID-private and
journaled before open or SaveAs. Uncertain identity retains/quarantines; no
active-document guesses, retries, input rewriting or recipient import.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession, apple_string, sha, inventory, retain_sources,
    package_styles, canonical, W, PKG, CLONE, IMPORTED, SOURCES,
    paragraph_style_by_name,
)
FRAGMENT = Path(__file__).parent/'run-01/fragment.xml'
SOURCE_STYLES = Path(__file__).parent/'run-01/source-styles.xml'
PINNED = {FRAGMENT: '56d2c52396c7ce74bbdcb201a31d7cb18f17e62b6edee77b3a59fe0125dce4b4',
          SOURCE_STYLES: '365a3b1103c9194f22b3b1ba9b7b22e5a3386c7b6421da00f5645c117eb02852'}


def materialized_style_valid(source_xml, saved_xml):
    source = paragraph_style_by_name(ET.fromstring(source_xml).findall(W+'style'), 'heading 1')
    styles = ET.fromstring(saved_xml).findall(W+'style')
    clone = paragraph_style_by_name(styles, CLONE)
    normal = paragraph_style_by_name(styles, 'Normal')
    if source is None or clone is None or normal is None or clone is normal:
        return False
    bases = clone.findall(W+'basedOn')
    if (len(bases) != 1 or bases[0].get(W+'val') != normal.get(W+'styleId')
            or clone.find(W+'link') is not None or clone.get(W+'customStyle') not in ('1','true')):
        return False
    return all(source.find(W+tag) is not None and clone.find(W+tag) is not None
               and canonical(source.find(W+tag)) == canonical(clone.find(W+tag))
               for tag in ('rPr','pPr'))


def binding_valid(rows, path):
    return (isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list)
            and len(rows[0])==4 and rows[0][0]=='binding'
            and type(rows[0][1]) is int and rows[0][1]>=0
            and rows[0][2:]==[str(path),path.name])


def path_binding(path):
    literal = apple_string(str(path))
    return ['set matches to {}','repeat with nativeDocument in documents',
            f'if (((current application\'s NSString\'s stringWithString:(posix full name of nativeDocument as text))\'s isEqualToString:{literal}) as boolean) then set end of matches to nativeDocument',
            'end repeat','if (count matches) is not 1 or (count documents) is not 1 then error "WPSC_FLATOPC_IDENTITY_DELTA"',
            'set boundDoc to item 1 of matches','set boundWindow to active window of boundDoc',
            f'if not (((current application\'s NSString\'s stringWithString:(name of boundDoc as text))\'s isEqualToString:{apple_string(path.name)}) as boolean) then error "WPSC_FLATOPC_NAME"',
            'set nativeRows to {{"binding",id of boundWindow,posix full name of boundDoc as text,name of boundDoc as text}}']


def open_commands(path):
    return ['if (count documents) is not 0 then error "WPSC_FLATOPC_REQUIRES_EMPTY"',
            f'open file name {apple_string(str(path))} file converter open format xmldocument serialized read only true add to recent files false confirm conversions false',
            *path_binding(path)]


def save_commands(path):
    return [f'save as boundDoc file name {apple_string(str(path))} file format format document default add to recent files false',
            *path_binding(path)]


def native_readback():
    return [f'set cloneStyle to Word style {apple_string(CLONE)} of boundDoc',
            'set normalName to name local of Word style (style normal) of boundDoc as text',
            'set cloneBase to name local of Word style (base style of cloneStyle) of boundDoc as text',
            'set nativeRows to {{"materialized",content of text object of boundDoc as text,name local of style of text object of paragraph 1 of boundDoc as text,name local of cloneStyle as text,normalName,cloneBase}}']


def native_valid(rows):
    return (isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list)
            and len(rows[0])==6 and rows[0][:4]==['materialized',IMPORTED+'\r',CLONE,CLONE]
            and isinstance(rows[0][4],str) and bool(rows[0][4]) and rows[0][4]==rows[0][5])


def package_valid(path, source_xml):
    with ZipFile(path) as package:
        styles_xml=package.read('word/styles.xml')
        if not materialized_style_valid(source_xml,styles_xml):return False
        clone=paragraph_style_by_name(ET.fromstring(styles_xml).findall(W+'style'),CLONE)
        body=ET.fromstring(package.read('word/document.xml')).find(W+'body')
        paras=body.findall(W+'p')
        if len(paras)!=1:return False
        p=paras[0]
        if ''.join(t.text or '' for t in p.iter(W+'t'))!=IMPORTED:return False
        props=p.findall(W+'pPr')
        if len(props)!=1 or len(props[0])!=1:return False
        pstyle=props[0][0]
        return (pstyle.tag==W+'pStyle' and pstyle.get(W+'val')==clone.get(W+'styleId')
                and all(len(rpr)==0 and not rpr.attrib for rpr in p.iter(W+'rPr')))


def run(output):
    report={'status':'FAIL','scope':'Explicit converter materialization only; no import or product parity',
            'checks':{},'journal':[],'raw_rows':{},
            'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_flatopc_open_control.py'),*SOURCES])}
    owner=None;reopened=None;identity_uncertain=False
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('FlatOPC control failed: '+label)
    def execute(session,label,commands,bind=True):
        report['current_phase']=label;flush()
        rows=session._execute(commands,bind=bind)
        report['raw_rows'][label]=rows;flush();return rows
    try:
        check('immutable_inputs',all(sha(path)==digest for path,digest in PINNED.items()))
        source_xml=SOURCE_STYLES.read_bytes()
        report['inventory_before']=inventory(output,'before')
        check('empty_inventory',report['inventory_before']==[])
        owner=MacWordSession();owner._prepare();owner._retain_evidence=True
        input_path=owner.staging_root/('flatopc-'+uuid4().hex+'.xml')
        output_path=owner.staging_root/('converted-'+uuid4().hex+'.docx')
        input_path.write_bytes(FRAGMENT.read_bytes())
        report['native_paths']={'input':str(input_path),'docx':str(output_path)}
        report['dictionary_sha256']=sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['journal'].append({'operation':'open','expected_before':[],'expected_after':[str(input_path)],'status':'submitted'})
        flush();identity_uncertain=True
        rows=execute(owner,'explicit_converter_open',open_commands(input_path),bind=False)
        check('exact_open_binding',binding_valid(rows,input_path))
        owner._private_path=input_path;owner._owns_doc=True;owner._bind(rows)
        identity_uncertain=False
        report['journal'][-1]['status']='acknowledged';flush()
        report['word_version']=execute(owner,'version',['set nativeRows to {{version as text}}'])
        check('native_materialized',native_valid(execute(owner,'native_readback',native_readback())))
        report['inventory_after_open']=inventory(output,'after-open')
        check('open_inventory_delta',len(report['inventory_after_open'])==1 and report['inventory_after_open'][0][1]==str(input_path))
        report['journal'].append({'operation':'save_as','expected_before':[str(input_path)],'expected_after':[str(output_path)],'status':'submitted'})
        flush();identity_uncertain=True
        rows=execute(owner,'native_save_as',save_commands(output_path))
        check('exact_save_binding',binding_valid(rows,output_path))
        owner._private_path=output_path;owner._bind(rows);identity_uncertain=False
        report['journal'][-1]['status']='acknowledged';flush()
        check('word_saved_complete_style',package_valid(output_path,source_xml))
        check('native_after_save',native_valid(execute(owner,'readback_after_save',native_readback())))
        shutil.copyfile(output_path,output/'native-materialized.docx')
        check('copied_native_bytes',sha(output_path)==sha(output/'native-materialized.docx'))
        (output/'native-styles.xml').write_bytes(package_styles(output_path))
        report['journal'].append({'operation':'close','expected_before':[str(output_path)],'expected_after':[],'status':'submitted'});flush()
        owner.close()
        check('closed',owner._closed and not owner._quarantined)
        check('empty_after_close',inventory(output,'after-close')==[])
        report['journal'][-1]['status']='acknowledged';flush()
        digest=sha(output/'native-materialized.docx')
        with MacWordSession.open_document(output/'native-materialized.docx',read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True
            check('reopen_native',native_valid(execute(reopened,'reopen_readback',native_readback())))
        check('reopen_closed',reopened._closed and not reopened._quarantined)
        check('reopen_bytes',sha(output/'native-materialized.docx')==digest)
        check('complete_reopened_package',package_valid(output/'native-materialized.docx',source_xml))
        check('input_private_preserved',sha(input_path)==PINNED[FRAGMENT])
        check('input_originals_preserved',all(sha(path)==value for path,value in PINNED.items()))
        report['inventory_final']=inventory(output,'final');check('empty_final',report['inventory_final']==[])
        check('sources_unchanged',all(sha(ROOT/path)==digest for path,digest in report['source_hashes'].items()))
        report['status']='PASS'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and not owner._closed:
            if identity_uncertain:
                owner._retain('Flat OPC open or SaveAs identity unverified; inspect journal before cleanup')
            try:owner.close()
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        for session,label in ((owner,'native-runtime'),(reopened,'reopen-runtime')):
            if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/label,dirs_exist_ok=True)
        if report.get('cleanup_failure'):report['status']='FAIL'
        report['identity_uncertain']=identity_uncertain
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.xml')}
        flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Flat OPC control requires explicit --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='PASS' else 1


if __name__=='__main__':raise SystemExit(main())
