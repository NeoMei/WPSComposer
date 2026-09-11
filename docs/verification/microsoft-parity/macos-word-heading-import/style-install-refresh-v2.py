"""Direct DOCX-input style-definition installation/refresh in one private recipient.

One specifically verified inherited property (left indent) only. All package
writes are input fragments; Word alone imports/edits/saves actual output.
Definition-only: no existing/fresh target-style paragraph behavior is proved.
No generic style-flattening, rollback, rendered-Arabic or public parity claim.
"""
from __future__ import annotations
import argparse
from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import shutil
import traceback
from uuid import uuid4
from zipfile import ZipFile,ZIP_DEFLATED
import xml.etree.ElementTree as ET
BASE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('install_retained_italic',BASE/'italic-face-diagnostic.py');base=importlib.util.module_from_spec(spec);spec.loader.exec_module(base)
v5=base.v5;W=v5.W;ROOT=v5.ROOT;sha=v5.sha;apple_string=v5.apple_string;units=v5.collapsed.units
INPUT=base.INPUT;INPUT_SHA=base.INPUT_SHA
NAME='WPSC Unnumbered Heading 1';STYLE_ID='WPSCUnnumberedHeading1';PARENT='WPSC Install Parent';BOOKMARK='WPSC_InstallHeading';CARRIER='WPSC INSTALL CARRIER 中文😀'
CT='{http://schemas.openxmlformats.org/package/2006/content-types}';REL='{http://schemas.openxmlformats.org/package/2006/relationships}';RT='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'


def expected_clone(styles):
    root=ET.fromstring(styles);nodes=root.findall(W+'style');source=v5.paragraph_style_by_name(nodes,'heading 1');parent=v5.paragraph_style_by_name(nodes,PARENT);normal=v5.paragraph_style_by_name(nodes,'Normal')
    if source is None or parent is None or normal is None:raise ValueError('Missing explicit source dependency')
    if source.find(W+'basedOn').get(W+'val')!=parent.get(W+'styleId') or parent.find(W+'basedOn').get(W+'val')!=normal.get(W+'styleId'):raise ValueError('Unsupported source chain')
    ppr=source.find(W+'pPr');rpr=source.find(W+'rPr');pp=parent.find(W+'pPr');pr=parent.find(W+'rPr')
    if ppr is None or rpr is None or ppr.find(W+'ind') is not None:raise ValueError('Expected source to inherit indent')
    if pp is None or len(pp)!=1 or pp[0].tag!=W+'ind' or pp[0].attrib!={W+'left':'340'} or (pr is not None and (len(pr) or pr.attrib)):raise ValueError('Only inherited left indent is supported')
    clone=ET.Element(W+'style',{W+'type':'paragraph',W+'customStyle':'1',W+'styleId':STYLE_ID})
    for tag,value in [('name',NAME),('basedOn',normal.get(W+'styleId')),('next',normal.get(W+'styleId'))]:ET.SubElement(clone,W+tag,{W+'val':value})
    clone_ppr=deepcopy(ppr)
    # OOXML pPr order: ind precedes outlineLvl. Do not reorder source properties.
    outline=clone_ppr.find(W+'outlineLvl');at=list(clone_ppr).index(outline) if outline is not None else len(clone_ppr)
    clone_ppr.insert(at,deepcopy(pp[0]));clone.append(clone_ppr);clone.append(deepcopy(rpr));return clone


def target_style(root):
    matches=[s for s in root.findall(W+'style') if s.get(W+'styleId')==STYLE_ID or (s.find(W+'name') is not None and s.find(W+'name').get(W+'val')==NAME)]
    if len(matches)>1:raise ValueError('Ambiguous target identity')
    if matches and (matches[0].get(W+'styleId')!=STYLE_ID or matches[0].find(W+'name').get(W+'val')!=NAME):raise ValueError('Renamed/colliding style')
    return matches[0] if matches else None


def other_styles_preserved(before,after):
    try:
        roots=[ET.fromstring(x) for x in (before,after)]
        for root in roots:
            ids=[n.get(W+'styleId') for n in root.findall(W+'style')]
            if len(ids)!=len(set(ids)):return False
            node=target_style(root)
            if node is not None:root.remove(node)
        return v5.complete_styles_equal(*(ET.tostring(x) for x in roots))
    except (ValueError,ET.ParseError):return False


def clone_matches(styles,expected):
    try:
        actual=target_style(ET.fromstring(styles))
        if actual is None:return False
        actual=deepcopy(actual)
        for child in actual.findall(W+'rsid'):actual.remove(child)
        return v5.canonical(actual)==v5.canonical(expected)
    except (ValueError,ET.ParseError):return False


def input_fragment(styles,theme,target):
    root=ET.fromstring(styles);clone=expected_clone(styles);normal=v5.paragraph_style_by_name(root.findall(W+'style'),'Normal')
    payload=ET.Element(W+'styles');payload.append(deepcopy(root.find(W+'docDefaults')));payload.append(deepcopy(normal));payload.append(clone)
    doc=ET.Element(W+'document');p=ET.SubElement(ET.SubElement(doc,W+'body'),W+'p');ET.SubElement(ET.SubElement(p,W+'pPr'),W+'pStyle',{W+'val':STYLE_ID});ET.SubElement(ET.SubElement(p,W+'r'),W+'t').text=CARRIER
    rels=ET.Element(REL+'Relationships');ET.SubElement(rels,REL+'Relationship',{'Id':'rId1','Type':RT+'officeDocument','Target':'word/document.xml'})
    docrels=ET.Element(REL+'Relationships')
    for i,kind,name in [(1,'styles','styles.xml'),(2,'theme','theme/theme1.xml')]:ET.SubElement(docrels,REL+'Relationship',{'Id':'rId'+str(i),'Type':RT+kind,'Target':name})
    parts={'_rels/.rels':ET.tostring(rels),'word/document.xml':ET.tostring(doc),'word/styles.xml':ET.tostring(payload),'word/_rels/document.xml.rels':ET.tostring(docrels),'word/theme/theme1.xml':theme}
    types=ET.Element(CT+'Types')
    for name,kind in [('word/document.xml','document.main'),('word/styles.xml','styles')]:ET.SubElement(types,CT+'Override',{'PartName':'/'+name,'ContentType':'application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml'})
    ET.SubElement(types,CT+'Override',{'PartName':'/word/theme/theme1.xml','ContentType':'application/vnd.openxmlformats-officedocument.theme+xml'})
    ET.SubElement(types,CT+'Default',{'Extension':'rels','ContentType':'application/vnd.openxmlformats-package.relationships+xml'})
    with ZipFile(target,'x',compression=ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ET.tostring(types))
        for name,data in parts.items():z.writestr(name,data)
    return clone


def clear_commands(text,start,end):
    raw=text.encode('utf-16-le')
    if type(start) is not int or type(end) is not int or not 0<=start<end<=len(raw)//2 or raw[start*2:end*2].decode('utf-16-le')!=CARRIER+'\r':raise ValueError('Not exact owned carrier')
    return v5.collapsed.full_guard(text)+[f'set carrierRange to create range boundDoc start {start} end {end}',f'if content of carrierRange is not {apple_string(CARRIER)} & return then error "WPSC_CARRIER_TEXT"',f'if start of content of carrierRange is not {start} or end of content of carrierRange is not {end} then error "WPSC_CARRIER_BOUNDS"','set content of carrierRange to ""','set ownSelection to selection of boundWindow','set selection start of ownSelection to 0','set selection end of ownSelection to 5','set nativeRows to {{"clear",content of text object of boundDoc as text}}']


def carrier_valid(xml):
    paragraphs=ET.fromstring(xml).find(W+'body').findall(W+'p')
    matches=[p for p in paragraphs if ''.join(t.text or '' for t in p.iter(W+'t'))==CARRIER]
    if len(paragraphs)!=5 or len(matches)!=1:return False
    pstyle=matches[0].findall('./'+W+'pPr/'+W+'pStyle')
    return len(pstyle)==1 and pstyle[0].attrib=={W+'val':STYLE_ID} and not len(pstyle[0]) and not pstyle[0].text


def document_signature(xml):
    root=ET.fromstring(xml);starts=list(root.iter(W+'bookmarkStart'));ends=list(root.iter(W+'bookmarkEnd'))
    if len(starts)!=1 or len(ends)!=1 or set(starts[0].attrib)!={W+'name',W+'id'} or set(ends[0].attrib)!={W+'id'} or starts[0].get(W+'name')!=BOOKMARK or starts[0].get(W+'id')!=ends[0].get(W+'id'):raise ValueError('Owned bookmark graph changed')
    for parent in root.iter():
        for child in list(parent):
            if child in starts+ends:parent.remove(child)
    return base.document_signature(ET.tostring(root))


def chain_readback(label):
    return ['set sourceStyle to Word style (style heading1) of boundDoc',
            f'set parentStyle to Word style {apple_string(PARENT)} of boundDoc',
            'set normalName to name local of Word style (style normal) of boundDoc as text',
            'set sourceName to name local of sourceStyle as text',
            'set parentName to name local of parentStyle as text',
            'set parentBase to name local of Word style (base style of parentStyle) of boundDoc as text',
            'set sourceBase to name local of Word style (base style of sourceStyle) of boundDoc as text',
            f'set nativeRows to {{{{{apple_string(label)},normalName,sourceName,parentName,parentBase,sourceBase,paragraph format left indent of paragraph format of parentStyle,paragraph format left indent of paragraph format of sourceStyle,built in of sourceStyle,built in of parentStyle}}}}']


def chain_valid(rows,label):
    if not isinstance(rows,list) or len(rows)!=1 or not isinstance(rows[0],list) or len(rows[0])!=10:return False
    r=rows[0]
    return (r[0]==label and all(type(x) is str and x for x in r[1:6]) and r[3]==PARENT and r[2]!=r[3]
            and r[4]==r[1] and r[5]==(r[1] if label=='parent' else PARENT)
            and type(r[6]) in (int,float) and r[6]==17 and type(r[7]) in (int,float)
            and r[7]==(0 if label=='parent' else 17) and r[8] is True and r[9] is False)


def parent_setup_commands():
    # Ignore make's object-valued return, then resolve the immutable exact name.
    return [f'make new Word style at boundDoc with properties {{name local:{apple_string(PARENT)}}}',
            f'set base style of Word style {apple_string(PARENT)} of boundDoc to style normal',
            f'set paragraph format left indent of paragraph format of Word style {apple_string(PARENT)} of boundDoc to 17',
            *chain_readback('parent')]


def setup_commands():
    # Dictionary declares WdBuiltinStyle; this custom-name text assignment is
    # explicitly a final bounded native hypothesis, never assumed support.
    return [f'set base style of Word style (style heading1) of boundDoc to {apple_string(PARENT)}',
            f'set headingRange to create range boundDoc start 0 end {units(v5.FRESH)}',
            f'make new bookmark at boundDoc with properties {{name:{apple_string(BOOKMARK)},text object:headingRange}}',
            'set ownSelection to selection of boundWindow','set selection start of ownSelection to 0',
            'set selection end of ownSelection to 5',*chain_readback('source')]


def parent_chain_xml_valid(styles):
    root=ET.fromstring(styles);nodes=root.findall(W+'style')
    normal=v5.paragraph_style_by_name(nodes,'Normal');parent=v5.paragraph_style_by_name(nodes,PARENT);source=v5.paragraph_style_by_name(nodes,'heading 1')
    if any(node is None for node in (normal,parent,source)):return False
    return all(len(node.findall(W+'basedOn'))==1 and node.find(W+'basedOn').attrib=={W+'val':normal.get(W+'styleId')} for node in (parent,source))


def state_commands():
    return ['set ownSelection to selection of boundWindow','if story type of ownSelection is not main text story then error "WPSC_INSTALL_STORY"',f'set headingBookmark to bookmark {apple_string(BOOKMARK)} of boundDoc','set sourceStyle to Word style (style heading1) of boundDoc','set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of text object of ownSelection,end of content of text object of ownSelection,start of bookmark of headingBookmark,end of bookmark of headingBookmark,content of text object of headingBookmark as text,count fields of boundDoc,count tables of boundDoc,font size of font object of sourceStyle,paragraph format left indent of paragraph format of sourceStyle}}']


def state_valid(rows,body,size,selection=True):
    if not isinstance(rows,list) or len(rows)!=1 or len(rows[0])!=12:return False
    r=rows[0]
    return (r[0]=='state' and r[1] is True and r[2]==body and all(type(r[i]) is int for i in (3,4,5,6,8,9)) and (not selection or r[3:5]==[0,5]) and r[5:10]==[0,units(v5.FRESH),v5.FRESH,0,0] and type(r[10]) in (int,float) and r[10]==size and type(r[11]) in (int,float) and r[11]==17)


def source_change_valid(before,after):
    root=ET.fromstring(before)
    for style in base.paired_styles(root):
        size=style.find('./'+W+'rPr/'+W+'sz')
        if size is None or size.attrib!={W+'val':'34'}:return False
        size.set(W+'val','46')
    return v5.complete_styles_equal(ET.tostring(root),after)


def run(output):
    report={'status':'FAIL','scope':'Persisted style definitions only: same-name install/refresh and one inherited indent; no existing/fresh target-style paragraph behavior; known Arabic rendering failure separate','checks':{},'trials':{},'source_hashes':v5.retain_sources(output,[Path(__file__),BASE/'test_style_install_refresh_v2.py',Path(base.__file__),*v5.SOURCES])};owner=reopened=None;sentinel=token=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(k,v):
        report['checks'][k]=bool(v);flush()
        if not v:raise AssertionError(k)
    def execute(k,lines):
        report['current_phase']=k;flush();rows=owner._execute(lines);report.setdefault('raw_rows',{})[k]=rows;flush();return rows
    try:
        check('pinned_input',sha(INPUT)==INPUT_SHA);report['inventory_before']=v5.inventory(output,'before');check('empty_inventory',report['inventory_before']==[])
        report['open_journal']={'source':str(INPUT),'state':'submitted'};flush();owner=v5.MacWordSession.open_document(INPUT,read_only=False,visible=False);owner._retain_evidence=True;report['open_journal'].update(state='acknowledged',private_path=str(owner._private_path));flush()
        opened=v5.inventory(output,'opened');check('private_identity',len(opened)==1 and opened[0][1]==str(owner._private_path) and sha(owner._private_path)==INPUT_SHA)
        token='INSTALL SENTINEL '+uuid4().hex+' 中文😀';rows=execute('sentinel',['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])
        if len(rows)!=1 or len(rows[0])!=1 or not isinstance(rows[0][0],str):owner._retain('Uncertain sentinel creation');raise ValueError('sentinel ACK')
        sentinel=rows[0][0];report['sentinel_before']=v5.sentinel_preimage(v5.inventory(output,'sentinel-before'),sentinel,token);report['sentinel_name']=sentinel
        ack=execute('parent_setup',parent_setup_commands());owner.save_docx(output/'parent-setup.docx');check('parent_native_chain',chain_valid(ack,'parent'));check('parent_saved_chain',parent_chain_xml_valid(base.package(output/'parent-setup.docx')[0]))
        ack=execute('setup',setup_commands());owner.save_docx(output/'setup.docx');check('source_native_chain',chain_valid(ack,'source'))
        styles,xml=base.package(output/'setup.docx');signature=document_signature(xml);body='\r'.join(''.join(t.text or '' for t in p.iter(W+'t')) for p in ET.fromstring(xml).find(W+'body').findall(W+'p'))+'\r';report['body']=body;check('source_inherited_case',expected_clone(styles) is not None)
        start=units(v5.FRESH+'\r'+v5.IMPORTED+'\r');raw=body.encode('utf-16-le');inserted=(raw[:start*2]+(CARRIER+'\r').encode('utf-16-le')+raw[start*2:]).decode('utf-16-le')
        with ZipFile(INPUT) as z:theme=z.read('word/theme/theme1.xml')
        for attempt,size in [(1,17),(2,23)]:
            label='attempt-'+str(attempt);trial=report['trials'][label]={};check(label+'_pre_state',state_valid(execute(label+'_before',state_commands()),body,size))
            donor=output/(label+'-input-'+uuid4().hex+'.docx');expected=input_fragment(styles,theme,donor)
            from skills.WPSComposer.scripts.msoffice.input_validation import validate_native_input
            validate_native_input(donor,'writer')
            digest=sha(donor);trial['input_sha256']=digest
            before_inventory=v5.inventory(output,label+'-before-insert');ack=execute(label+'_import',v5.collapsed.collapsed_commands(donor,start,body));owner.save_docx(output/(label+'-inserted.docx'))
            check(label+'_insert_ack',v5.collapsed.exact_ack(ack,'insert',start,start,inserted));current,currentxml=base.package(output/(label+'-inserted.docx'))
            check(label+'_other_styles_preserved',other_styles_preserved(styles,current));check(label+'_carrier_identity',carrier_valid(currentxml));check(label+'_inventory',v5.collapsed.unrelated_inventory_preserved(before_inventory,v5.inventory(output,label+'-after-insert'),owner._private_path))
            check(label+'_bookmark_preserved',state_valid(execute(label+'_inserted_state',state_commands()),inserted,size,selection=False))
            # A stale target is a diagnostic result; remove only the exact owned
            # carrier before asserting install/refresh success.
            trial['target_style_matches']=clone_matches(current,expected);flush()
            ack=execute(label+'_clear',clear_commands(inserted,start,start+units(CARRIER+'\r')));owner.save_docx(output/(label+'-clean.docx'));check(label+'_clear_ack',ack==[['clear',body]])
            cleaned,cleanxml=base.package(output/(label+'-clean.docx'));check(label+'_clean_styles',v5.complete_styles_equal(current,cleaned));check(label+'_exact_document',document_signature(cleanxml)==signature);check(label+'_selection_bookmark',state_valid(execute(label+'_clean_state',state_commands()),body,size));check(label+'_input_preserved',sha(donor)==digest)
            check(label+'_theme_preserved',v5.theme_hash(output/(label+'-clean.docx'))==v5.theme_hash(INPUT));check(label+'_target_installed_or_refreshed',trial['target_style_matches'])
            if attempt==1:
                ack=execute('change_source',['set sourceStyle to Word style (style heading1) of boundDoc','if font size of font object of sourceStyle is not 17 then error "WPSC_SOURCE_SIZE"','set font size of font object of sourceStyle to 23','set nativeRows to {{"size",font size of font object of sourceStyle}}']);owner.save_docx(output/'source-changed.docx');check('source_change_ack',ack==[['size',23]] and type(ack[0][1]) in (int,float));changed,changedxml=base.package(output/'source-changed.docx');check('source_change_exact_pair',source_change_valid(cleaned,changed));check('source_change_document',document_signature(changedxml)==signature);styles=changed
        owner.close();check('owned_closed',owner._closed and not owner._quarantined);v5.close_after_owned(owner,output,report,sentinel,token);sentinel=None
        final=output/'attempt-2-clean.docx';digest=sha(final)
        with v5.MacWordSession.open_document(final,read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True
            finalstyles,finalxml=base.package(reopened._private_path);check('reopen_style',clone_matches(finalstyles,expected));check('reopen_document',document_signature(finalxml)==signature)
        check('reopen_preserved',sha(final)==digest and reopened._closed and not reopened._quarantined);check('original_preserved',sha(INPUT)==INPUT_SHA);check('final_empty',v5.inventory(output,'final')==[]);check('sources_unchanged',all(sha(ROOT/p)==h for p,h in report['source_hashes'].items()));report['status']='PASS'
    except BaseException as exc:report['error']={'type':type(exc).__name__,'message':str(exc)};(output/'failure.txt').write_text(traceback.format_exc())
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
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix=='.docx'};flush()
    return report


def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--compile-only',action='store_true');p.add_argument('--output',required=True,type=Path);a=p.parse_args()
    if a.execute==a.compile_only:p.error('Choose exactly one mode')
    output=a.output.resolve();output.mkdir(parents=True,exist_ok=False)
    if a.compile_only:
        from skills.WPSComposer.scripts.msoffice.macos_word_session import _JSON
        body=v5.FRESH+'\r'+v5.IMPORTED+'\rSUFFIX\r'+v5.FRESH+'\r';start=units(v5.FRESH+'\r'+v5.IMPORTED+'\r');raw=body.encode('utf-16-le');inserted=(raw[:start*2]+(CARRIER+'\r').encode('utf-16-le')+raw[start*2:]).decode('utf-16-le')
        for label,lines in [('parent-setup',parent_setup_commands()),('source-setup',setup_commands()),('state',state_commands()),('import',v5.collapsed.collapsed_commands(Path('/tmp/COMPILE_ONLY.docx'),start,body)),('clear',clear_commands(inserted,start,start+units(CARRIER+'\r')))]:
            (output/(label+'.applescript')).write_text(_JSON+'\ntell application "Microsoft Word"\nset boundDoc to document "COMPILE_ONLY.docx"\nset boundWindow to active window of boundDoc\n'+'\n'.join(lines)+'\nend tell\nreturn my jsonRows(nativeRows)\n')
        return 0
    return 0 if run(output)['status']=='PASS' else 1

if __name__=='__main__':raise SystemExit(main())
