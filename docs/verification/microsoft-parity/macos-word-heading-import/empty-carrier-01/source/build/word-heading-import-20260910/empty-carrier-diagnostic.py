"""B-only controlled empty carrier: apply clone, insert text, never reset.

The pinned input's fourth and terminal paragraph must have no content, runs,
or paragraph properties. No extra paragraph is created. All native output and
raw XML are saved before diagnostic assertions; no style visibility exception,
font rewrite, arbitrary existing-paragraph support, or public parity claim.
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
PRIOR=Path(__file__).with_name('fresh-order-diagnostic.py')
spec=importlib.util.spec_from_file_location('retained_fresh_order',PRIOR)
prior=importlib.util.module_from_spec(spec);spec.loader.exec_module(prior)
ROOT=prior.ROOT;MacWordSession=prior.MacWordSession;apple_string=prior.apple_string
sha=prior.sha;inventory=prior.inventory;retain_sources=prior.retain_sources
CLONE=prior.CLONE;FRESH=prior.FRESH;IMPORTED=prior.IMPORTED;W=prior.W
SOURCES=prior.SOURCES;package_styles=prior.package_styles
existing_styles_preserved=prior.existing_styles_preserved;clone_xml_valid=prior.clone_xml_valid
fresh_package_valid=prior.fresh_package_valid;HELPER=prior.HELPER;font=prior.font
INPUT=prior.INPUT;INPUT_SHA=prior.INPUT_SHA
save_after_creation=prior.save_after_creation;anchor_map=prior.anchor_map
fresh_font_commands=prior.fresh_font_commands;paragraph_xml_observation=prior.paragraph_xml_observation
state_commands=prior.state_commands


def empty_carrier_xml_valid(xml):
    body=ET.fromstring(xml).find(W+'body')
    if body is None or any(child.tag not in (W+'p',W+'sectPr') for child in body):return False
    paragraphs=body.findall(W+'p')
    return len(paragraphs)==4 and len(paragraphs[3])==0 and not (paragraphs[3].text or '')


def creation_commands(mode,preimage,p):
    if mode!='B' or type(p) is not int or p<0:raise ValueError('Only existing empty carrier B is supported')
    raw=preimage.encode('utf-16-le')
    if p!=len(raw)//2-1 or raw[p*2:].decode('utf-16-le')!='\r':raise ValueError('Expected terminal paragraph')
    return [f'if not (((current application\'s NSString\'s stringWithString:(content of text object of boundDoc as text))\'s isEqualToString:{apple_string(preimage)}) as boolean) then error "WPSC_EMPTY_CARRIER_PREIMAGE"',
            'set carrierRange to text object of paragraph 4 of boundDoc',
            f'if start of content of carrierRange is not {p} or end of content of carrierRange is not {p+1} then error "WPSC_EMPTY_CARRIER_BOUNDS"',
            'if content of carrierRange is not return then error "WPSC_EMPTY_CARRIER_TEXT"',
            f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
            'set style of carrierRange to detachedStyle',
            f'set freshInsert to create range boundDoc start {p} end {p}',
            f'set content of freshInsert to {apple_string(FRESH)}',
            'set nativeRows to {{"stage",true}}']


def run(output):
    report={'status':'FAIL','scope':'Existing empty carrier without resets; diagnostic only, no appearance or public parity acceptance',
            'font_owner_roles':{'source':'existing style-only imported paragraph 2','clone':'new fresh paragraph 4'},
            'checks':{},'trials':{},'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_empty_carrier_diagnostic.py'),PRIOR,HELPER,*SOURCES])}
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
        with ZipFile(INPUT) as package:
            check('pinned_empty_carrier_xml',empty_carrier_xml_valid(package.read('word/document.xml')))
        source_styles=package_styles(INPUT)
        check('complete_input_clone',clone_xml_valid(source_styles,'detached-properties'))
        for mode in ('B',):
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
            raw=body.encode('utf-16-le');expected=raw[:p*2].decode('utf-16-le')+FRESH+raw[p*2:].decode('utf-16-le')
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
            check(mode+'_paragraph_count_preserved',after[0][7]==before[0][7])
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
        print('Empty carrier diagnostic requires --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='DIAGNOSTIC_COMPLETE' else 1


if __name__=='__main__':raise SystemExit(main())
