"""Controlled owned-copy fresh paragraph order A/B; diagnostic, not parity.

A inserts text before clone style/reset. B styles/resets an empty paragraph
before inserting identical text. Save each native package before observations
or mutation ACK assertions. Source paragraph direct overrides are not a style
appearance oracle; compare fresh anchors with existing style-only import.
"""
from __future__ import annotations
import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import sys
import traceback
from zipfile import ZipFile
import xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession,apple_string,sha,inventory,retain_sources,CLONE,FRESH,IMPORTED,W,
    SOURCES,package_styles,existing_styles_preserved,clone_xml_valid,fresh_package_valid,
)
HELPER=Path(__file__).with_name('font-anchor-read-diagnostic.py')
spec=importlib.util.spec_from_file_location('retained_font_anchor',HELPER)
font=importlib.util.module_from_spec(spec);spec.loader.exec_module(font)
INPUT=font.INPUT;INPUT_SHA=font.INPUT_SHA


def creation_commands(mode,preimage,p):
    if mode not in ('A','B') or type(p) is not int or p<0:raise ValueError('Invalid trial')
    raw=preimage.encode('utf-16-le')
    if p!=len(raw)//2-1 or raw[p*2:].decode('utf-16-le')!='\r':raise ValueError('Expected terminal paragraph')
    lines=[f'if not (((current application\'s NSString\'s stringWithString:(content of text object of boundDoc as text))\'s isEqualToString:{apple_string(preimage)}) as boolean) then error "WPSC_FRESH_PREIMAGE"',
           'set terminalRange to text object of paragraph 4 of boundDoc',
           f'if start of content of terminalRange is not {p} or content of terminalRange is not return then error "WPSC_FRESH_TERMINAL"',
           f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
           f'set freshRange to create range boundDoc start {p} end {p}',
           (f'set content of freshRange to {apple_string(FRESH)} & return' if mode=='A' else 'set content of freshRange to return'),
           'set freshRange to text object of paragraph 4 of boundDoc',
           'set style of freshRange to detachedStyle','reset font object of freshRange','reset paragraph format of freshRange']
    if mode=='B':
        lines += ['set freshStart to start of content of freshRange',
                  'set freshInsert to create range boundDoc start freshStart end freshStart',
                  f'set content of freshInsert to {apple_string(FRESH)}']
    return lines+['set nativeRows to {{"stage",true}}']


def save_after_creation(owner,path,mutation):
    rows=mutation()
    owner.save_docx(path)
    return rows


def range_anchors(body,start,end,marker):
    if type(start) is not int or type(end) is not int:raise ValueError('Invalid anchor bounds')
    raw=body.encode('utf-16-le')
    if not 0<=start<end<=len(raw)//2:raise ValueError('Invalid anchor bounds')
    try:text=raw[start*2:end*2].decode('utf-16-le')
    except UnicodeDecodeError:raise ValueError('Split surrogate') from None
    if text!=marker+'\r':raise ValueError('Wrong paragraph preimage')
    result=[]
    for script,scalar in [('latin','E'),('cjk','中'),('arabic','ا'),('emoji','😀')]:
        offset=text.index(scalar);left=start+len(text[:offset].encode('utf-16-le'))//2
        result.append({'script':script,'scalar':scalar,'start':left,'end':left+len(scalar.encode('utf-16-le'))//2})
    return result


def anchor_map(body,s1,e1,s2,e2):
    return {owner:[{'script':'paragraph','scalar':marker+'\r','start':start,'end':end},*range_anchors(body,start,end,marker)]
            for owner,start,end,marker in [('source',s1,e1,IMPORTED),('clone',s2,e2,FRESH)]}


def fresh_font_commands(body,anchors):
    commands=font.font_commands(body,anchors)
    old='set sourceStyle to Word style (style heading1) of boundDoc'
    commands[commands.index(old)]=f'set sourceStyle to Word style {apple_string(CLONE)} of boundDoc'
    return commands


def paragraph_xml_observation(xml,marker):
    paragraphs=[p for p in ET.fromstring(xml).iter(W+'p') if ''.join(t.text or '' for t in p.iter(W+'t'))==marker]
    if len(paragraphs)!=1:return {'unique':False,'count':len(paragraphs)}
    p=paragraphs[0]
    return {'unique':True,'pPr_xml':[ET.tostring(x,encoding='unicode') for x in p.findall(W+'pPr')],
            'runs':[{'text':''.join(t.text or '' for t in run.iter(W+'t')),
                     'rPr_xml':''.join(ET.tostring(x,encoding='unicode') for x in run.findall(W+'rPr'))}
                    for run in p.iter(W+'r')]}


def state_commands():
    return ['set referenceRange to text object of paragraph 2 of boundDoc','set freshRange to text object of paragraph 4 of boundDoc',
            'set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of referenceRange,end of content of referenceRange,start of content of freshRange,end of content of freshRange,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}']


def run(output):
    report={'status':'FAIL','scope':'Controlled fresh-order A/B diagnostic; no fresh appearance or public parity acceptance',
            'font_owner_roles':{'source':'existing style-only imported paragraph 2','clone':'new fresh paragraph 4'},
            'checks':{},'trials':{},'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_fresh_order_diagnostic.py'),HELPER,*SOURCES])}
    owner=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('Fresh-order diagnostic failed: '+label)
    def execute(mode,label,commands):
        report['current_phase']=mode+'_'+label;flush();rows=owner._execute(commands)
        report['trials'][mode].setdefault('raw_rows',{})[label]=rows;flush();return rows
    try:
        check('pinned_input',sha(INPUT)==INPUT_SHA)
        source_styles=package_styles(INPUT)
        check('complete_input_clone',clone_xml_valid(source_styles,'detached-properties'))
        for mode in ('A','B'):
            trial=report['trials'][mode]={'status':'FAIL'};flush()
            check(mode+'_empty_inventory_before',inventory(output,mode+'-before')==[])
            trial['open_journal']={'source':str(INPUT),'status':'submitted'};flush()
            owner=MacWordSession.open_document(INPUT,read_only=False,visible=False);owner._retain_evidence=True
            trial['open_journal'].update(status='acknowledged',private_path=str(owner._private_path));flush()
            trial['inventory_after_open']=inventory(output,mode+'-after-open')
            check(mode+'_exact_private_inventory',len(trial['inventory_after_open'])==1 and trial['inventory_after_open'][0][1]==str(owner._private_path))
            before=execute(mode,'preimage',state_commands());check(mode+'_preimage_shape',font.state_valid(before))
            check(mode+'_private_input_identical',sha(owner._private_path)==INPUT_SHA)
            body=before[0][2];p=before[0][5]
            raw=body.encode('utf-16-le');expected=raw[:p*2].decode('utf-16-le')+FRESH+'\r'+raw[p*2:].decode('utf-16-le')
            trial['expected_text']=expected;flush()
            native_path=output/(mode+'-fresh.docx')
            rows=save_after_creation(owner,native_path,lambda:execute(mode,'create_fresh',creation_commands(mode,body,p)))
            # Saved package and raw XML are retained before any acceptance-like
            # observation. A negative fresh style-only verdict stays diagnostic.
            trial['native_saved_path']=str(native_path);trial['native_saved_sha256']=sha(native_path)
            with ZipFile(native_path) as z:
                xml=z.read('word/document.xml');styles=z.read('word/styles.xml')
                (output/(mode+'-document.xml')).write_bytes(xml);(output/(mode+'-styles.xml')).write_bytes(styles)
            trial['fresh_xml']=paragraph_xml_observation(xml,FRESH)
            trial['reference_imported_xml']=paragraph_xml_observation(xml,IMPORTED)
            trial['fresh_style_only_xml']=fresh_package_valid(native_path);flush()
            check(mode+'_mutation_ack',rows==[['stage',True]])
            check(mode+'_all_style_definitions_preserved',existing_styles_preserved(source_styles,styles) and existing_styles_preserved(styles,source_styles))
            check(mode+'_complete_clone_xml',clone_xml_valid(styles,'detached-properties'))
            after=execute(mode,'saved_state',state_commands());check(mode+'_saved_state_shape',font.state_valid(after))
            check(mode+'_exact_inserted_text',after[0][2]==expected)
            check(mode+'_unchanged_fields_tables',after[0][8:]==before[0][8:])
            check(mode+'_one_added_paragraph',after[0][7]==before[0][7]+1)
            anchors=anchor_map(after[0][2],*after[0][3:7]);trial['anchors']=anchors;flush()
            font_rows=execute(mode,'font_anchors',fresh_font_commands(after[0][2],anchors))
            check(mode+'_typed_font_observations',font.rows_valid(font_rows,anchors))
            trial['style_comparisons']=font.observations(font_rows)
            groups={side:{r[2]:r[7:] for r in font_rows[1:] if r[1]==side} for side in ('source','clone')}
            trial['fresh_vs_existing_imported']={scope:font.compare_fields(groups['source'][scope],groups['clone'][scope]) for scope in groups['source']};flush()
            check(mode+'_reads_preserve_state',execute(mode,'after_read_state',state_commands())==after)
            current_inventory=inventory(output,mode+'-after-read');trial['inventory_after_read']=current_inventory
            check(mode+'_owned_inventory_after_read',len(current_inventory)==1 and current_inventory[0][1]==str(owner._private_path) and current_inventory[0][2] is True)
            check(mode+'_saved_private_matches_output',sha(owner._private_path)==sha(native_path))
            check(mode+'_immutable_input',sha(INPUT)==INPUT_SHA)
            trial['close_journal']={'private_path':str(owner._private_path),'status':'submitted'};flush()
            owner.close();trial['close_journal']['status']='acknowledged'
            check(mode+'_owned_closed',owner._closed and not owner._quarantined)
            check(mode+'_final_empty',inventory(output,mode+'-final')==[])
            if owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/(mode+'-native-runtime'),dirs_exist_ok=True)
            trial['status']='DIAGNOSTIC_COMPLETE';owner=None;flush()
        check('original_input_final_preserved',sha(INPUT)==INPUT_SHA)
        check('sources_unchanged',all(sha(ROOT/p)==digest for p,digest in report['source_hashes'].items()))
        report['status']='DIAGNOSTIC_COMPLETE'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)};(output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and not owner._closed:
            try:owner.close()
            except BaseException:report['cleanup_failure']=traceback.format_exc();report['status']='FAIL'
        if owner and owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/'failure-native-runtime',dirs_exist_ok=True)
        flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Fresh order diagnostic requires --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='DIAGNOSTIC_COMPLETE' else 1


if __name__=='__main__':raise SystemExit(main())
