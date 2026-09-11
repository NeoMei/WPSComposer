"""Gate A: immutable four-part DOCX INPUT -> native Word save and reopen.

Only complete named rPr/pPr materialization is tested. The four-part input has
no source theme: resolved appearance, independence, numbering and selection
remain unproved. This is not a production writer or final document artifact.
"""
from __future__ import annotations
import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

BASE = Path(__file__).with_name('flatopc-open-control.py')
spec = importlib.util.spec_from_file_location('retained_flatopc_gate', BASE)
gate = importlib.util.module_from_spec(spec); spec.loader.exec_module(gate)
ROOT=gate.ROOT
MacWordSession=gate.MacWordSession
sha=gate.sha;inventory=gate.inventory;retain_sources=gate.retain_sources
W=gate.W;PKG=gate.PKG;CLONE=gate.CLONE;IMPORTED=gate.IMPORTED
FRAGMENT=gate.FRAGMENT;SOURCE_STYLES=gate.SOURCE_STYLES;PINNED=gate.PINNED
paragraph_style_by_name=gate.paragraph_style_by_name
package_valid=gate.package_valid
CT='{http://schemas.openxmlformats.org/package/2006/content-types}'
REL='{http://schemas.openxmlformats.org/package/2006/relationships}'
REL_TYPE='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
PARTS={
    '_rels/.rels':'application/vnd.openxmlformats-package.relationships+xml',
    'word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'word/styles.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    'word/_rels/document.xml.rels':'application/vnd.openxmlformats-package.relationships+xml',
}


def input_parts(fragment):
    root=ET.fromstring(fragment)
    if root.tag!=PKG+'package' or len(root)!=4:raise ValueError('Expected four immutable XML input parts')
    parts={}
    for part in root:
        name=part.get(PKG+'name','')
        if (part.tag!=PKG+'part' or name not in {'/'+p for p in PARTS}
                or name[1:] in parts or part.get(PKG+'contentType')!=PARTS[name[1:]]
                or len(part)!=1 or part[0].tag!=PKG+'xmlData' or len(part[0])!=1):
            raise ValueError('Unexpected input part')
        parts[name[1:]]=part[0][0]
    for path,kind,target in (('_rels/.rels','officeDocument','word/document.xml'),
                             ('word/_rels/document.xml.rels','styles','styles.xml')):
        rels=parts[path]
        if (rels.tag!=REL+'Relationships' or len(rels)!=1 or rels[0].tag!=REL+'Relationship'
                or rels[0].attrib!={'Id':'rId1','Type':REL_TYPE+kind,'Target':target}):
            raise ValueError('Unexpected relationship graph')
    if parts['word/document.xml'].tag!=W+'document' or parts['word/styles.xml'].tag!=W+'styles':
        raise ValueError('Unexpected Word input part roots')
    return parts


def package_input(fragment,target):
    parts=input_parts(fragment)
    types=ET.Element(CT+'Types')
    for name,kind in PARTS.items():ET.SubElement(types,CT+'Override',{'PartName':'/'+name,'ContentType':kind})
    with ZipFile(target,'x',compression=ZIP_DEFLATED) as package:
        package.writestr('[Content_Types].xml',ET.tostring(types,encoding='utf-8',xml_declaration=True))
        for name,root in parts.items():package.writestr(name,ET.tostring(root,encoding='utf-8',xml_declaration=True))


def input_valid(path,fragment):
    parts=input_parts(fragment)
    with ZipFile(path) as package:
        names=package.namelist()
        if len(names)!=5 or set(names)!=set(PARTS)|{'[Content_Types].xml'}:return False
        types=ET.fromstring(package.read('[Content_Types].xml'))
        if (types.tag!=CT+'Types' or len(types)!=4 or any(x.tag!=CT+'Override' for x in types)
                or {x.get('PartName'):x.get('ContentType') for x in types}!={'/'+k:v for k,v in PARTS.items()}):return False
        return all(gate.canonical(root)==gate.canonical(ET.fromstring(package.read(name))) for name,root in parts.items())


def native_serialized_valid(source,saved):
    return saved.is_file() and sha(source)!=sha(saved)


def run(output):
    report={'status':'FAIL','scope':'Complete named rPr/pPr and style-only body materialization only; theme appearance and product parity unproved',
            'checks':{},'journal':[],'raw_rows':{},
            'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_docx_package_materialization.py'),BASE,*gate.SOURCES])}
    owner=None;reopened=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('DOCX materialization gate failed: '+label)
    def execute(session,label,commands):
        report['current_phase']=label;flush()
        rows=session._execute(commands)
        report['raw_rows'][label]=rows;flush();return rows
    def inventory_check(session,label):
        rows=inventory(output,label);report[label]=rows;flush()
        check(label+'_exact_private',len(rows)==1 and rows[0][1]==str(session._private_path))
    def journal(operation,path):
        report['journal'].append({'operation':operation,'source':str(path),'status':'submitted'});flush()
    def ack():report['journal'][-1]['status']='acknowledged';flush()
    try:
        check('immutable_inputs',all(sha(p)==digest for p,digest in PINNED.items()))
        source_xml=SOURCE_STYLES.read_bytes()
        input_path=output/('packaged-input-'+uuid4().hex+'.docx')
        package_input(FRAGMENT.read_bytes(),input_path)
        report['input_path']=str(input_path);report['input_sha256']=sha(input_path)
        check('complete_four_part_input',input_valid(input_path,FRAGMENT.read_bytes()))
        check('complete_input_named_style',package_valid(input_path,source_xml))
        check('empty_inventory',inventory(output,'before')==[])
        journal('open_document_private_copy',input_path)
        owner=MacWordSession.open_document(input_path,read_only=False,visible=False)
        owner._retain_evidence=True
        report['private_path']=str(owner._private_path);ack()
        inventory_check(owner,'after_open')
        check('native_open_materialized',gate.native_valid(execute(owner,'open_readback',gate.native_readback())))
        rows=execute(owner,'mark_owned_unsaved',['set saved of boundDoc to false','set nativeRows to {{"saved",saved of boundDoc}}'])
        check('owned_unsaved_ack',rows==[['saved',False]])
        native_path=output/'native-materialized.docx'
        journal('save_docx',native_path);owner.save_docx(native_path);ack()
        rows=execute(owner,'saved_readback',['set nativeRows to {{"saved",saved of boundDoc}}'])
        check('native_saved_ack',rows==[['saved',True]])
        check('actual_native_serialization',native_serialized_valid(input_path,native_path))
        check('complete_word_saved_style_body',package_valid(native_path,source_xml))
        check('native_after_save',gate.native_valid(execute(owner,'after_save_readback',gate.native_readback())))
        inventory_check(owner,'after_save')
        (output/'native-styles.xml').write_bytes(gate.package_styles(native_path))
        journal('close_owned',owner._private_path);owner.close();ack()
        check('closed',owner._closed and not owner._quarantined)
        check('empty_after_close',inventory(output,'after-close')==[])
        native_digest=sha(native_path)
        journal('reopen_document_private_copy',native_path)
        reopened=MacWordSession.open_document(native_path,read_only=True,visible=False)
        reopened._retain_evidence=True
        report['reopen_private_path']=str(reopened._private_path);ack()
        inventory_check(reopened,'after_reopen')
        check('reopened_native',gate.native_valid(execute(reopened,'reopen_readback',gate.native_readback())))
        check('complete_reopened_private_package',package_valid(reopened._private_path,source_xml))
        journal('close_reopened_owned',reopened._private_path);reopened.close();ack()
        check('reopen_closed',reopened._closed and not reopened._quarantined)
        check('native_output_preserved',sha(native_path)==native_digest)
        check('packaged_input_preserved',sha(input_path)==report['input_sha256'])
        check('original_inputs_preserved',all(sha(p)==digest for p,digest in PINNED.items()))
        check('empty_final',inventory(output,'final')==[])
        check('sources_unchanged',all(sha(ROOT/p)==digest for p,digest in report['source_hashes'].items()))
        report['status']='PASS'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc),'staging_path':str(getattr(exc,'staging_path',None))}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        for session,label in ((owner,'native-runtime'),(reopened,'reopen-runtime')):
            if session and not session._closed:
                try:session.close()
                except BaseException:report[label+'_cleanup_failure']=traceback.format_exc();report['status']='FAIL'
            if session and session.staging_root and session.staging_root.exists():shutil.copytree(session.staging_root,output/label,dirs_exist_ok=True)
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.xml')};flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('DOCX materialization requires explicit --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='PASS' else 1


if __name__=='__main__':raise SystemExit(main())
