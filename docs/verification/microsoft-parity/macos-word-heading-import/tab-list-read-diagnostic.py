"""Read-only tab collection diagnostic on the exact saved import snapshot.

A: explicit get of style paragraph-format tab list, then local list count.
B: previously exercised paragraph tab collection access on style-only paragraphs.
B cannot establish unused-style collection access. Neither block substitutes
empty values for failed reads. No import, edits, save, retry, or public enable.
"""
from __future__ import annotations
import argparse
import json
import math
from pathlib import Path
import shutil
import sys
import traceback

ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession,apple_string,sha,inventory,retain_sources,CLONE,SOURCES,
)
INPUT=Path(__file__).parent/'native-donor-import-01/before-delete.docx'
INPUT_SHA='1629dbce65928e4694822753ab44583a9a6825dab8b210f308b797f8d933600b'


def tab_commands(mode):
    if mode not in ('style','paragraph'):raise ValueError('Unknown tab diagnostic')
    lines=['set sourceStyle to Word style (style heading1) of boundDoc',
           f'set cloneStyle to Word style {apple_string(CLONE)} of boundDoc',
           'set sourceName to name local of sourceStyle as text',
           'set cloneName to name local of cloneStyle as text']
    if mode=='style':
        lines += ['set sourceParagraph to paragraph format of sourceStyle',
                  'set cloneParagraph to paragraph format of cloneStyle',
                  'set sourceTabs to (get every tab stop of sourceParagraph)',
                  'set cloneTabs to (get every tab stop of cloneParagraph)',
                  'if class of sourceTabs is not list or class of cloneTabs is not list then error "WPSC_TAB_LIST_TYPE"',
                  'set sourceCount to count sourceTabs','set cloneCount to count cloneTabs']
    else:
        lines += ['set sourceRange to text object of paragraph 1 of boundDoc',
                  'set cloneRange to text object of paragraph 2 of boundDoc',
                  'if name local of style of sourceRange is not sourceName or name local of style of cloneRange is not cloneName then error "WPSC_TAB_PARAGRAPH_STYLE"',
                  'set sourceParagraph to paragraph 1 of sourceRange',
                  'set cloneParagraph to paragraph 1 of cloneRange',
                  'set sourceCount to count tab stops of sourceParagraph',
                  'set cloneCount to count tab stops of cloneParagraph']
    lines += [f'set nativeRows to {{{{"tabs",{apple_string(mode)},sourceName,cloneName,sourceCount,cloneCount}}}}']
    for label,prefix in [('source','source'),('clone','clone')]:
        lines += [f'repeat with tabIndex from 1 to {prefix}Count',
                  f'set ownTab to '+(f'item tabIndex of {prefix}Tabs' if mode=='style' else f'tab stop tabIndex of {prefix}Paragraph'),
                  f'set end of nativeRows to {{{apple_string(label)},tabIndex,tab stop position of ownTab,alignment of ownTab as text,tab leader of ownTab as text,custom tab of ownTab}}',
                  'end repeat']
    return lines


def tabs_valid(rows,mode):
    if not isinstance(rows,list) or not rows or not isinstance(rows[0],list) or len(rows[0])!=6:return False
    header=rows[0]
    if (header[:2]!=['tabs',mode] or not isinstance(header[2],str) or not header[2]
            or header[3]!=CLONE or any(type(n) is not int or not 0<=n<=10000 for n in header[4:])):return False
    count,other=header[4:]
    if count!=other or len(rows)!=1+count+other:return False
    observed={}
    offset=1
    for label,length in [('source',count),('clone',other)]:
        values=[]
        for i,row in enumerate(rows[offset:offset+length],1):
            if (not isinstance(row,list) or len(row)!=6 or row[:2]!=[label,i]
                    or type(row[1]) is not int or type(row[2]) not in (int,float) or not math.isfinite(row[2])
                    or any(not isinstance(v,str) or not v for v in row[3:5]) or type(row[5]) is not bool):return False
            values.append(row[2:])
        observed[label]=values;offset+=length
    return observed['source']==observed['clone']


def state_commands():
    return ['set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}']


def state_valid(rows):
    return (isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list) and len(rows[0])==6
            and rows[0][:2]==['state',True] and isinstance(rows[0][2],str)
            and all(type(v) is int and v>=0 for v in rows[0][3:]))


def run(output):
    report={'status':'FAIL','scope':'Read-only tab diagnostics; empty-list result proves only this fixture; no full import or public parity',
            'checks':{},'blocks':{},'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_tab_list_read_diagnostic.py'),*SOURCES])}
    owner=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('Tab read diagnostic failed: '+label)
    def execute(label,commands):
        report['current_phase']=label;flush();rows=owner._execute(commands)
        report.setdefault('raw_rows',{})[label]=rows;flush();return rows
    try:
        check('pinned_input',sha(INPUT)==INPUT_SHA)
        report['inventory_before']=inventory(output,'before');check('empty_inventory',report['inventory_before']==[])
        report['open_journal']={'source':str(INPUT),'operation':'open_document_read_only_private_copy','status':'submitted'};flush()
        owner=MacWordSession.open_document(INPUT,read_only=True,visible=False);owner._retain_evidence=True
        report['open_journal'].update(status='acknowledged',private_path=str(owner._private_path));flush()
        initial_inventory=inventory(output,'after-open');report['inventory_after_open']=initial_inventory
        check('exact_owned_inventory',len(initial_inventory)==1 and initial_inventory[0][1]==str(owner._private_path))
        initial=execute('initial_state',state_commands());check('initial_saved_state',state_valid(initial))
        check('initial_private_hash',sha(owner._private_path)==INPUT_SHA)
        for mode in ('style','paragraph'):
            # Independent, attributed reads. A failed block is retained and cannot
            # make the overall probe PASS; B is never an empty-style fallback.
            report['blocks'][mode]={'status':'FAIL'};flush()
            try:
                rows=execute(mode+'_tabs',tab_commands(mode))
                valid=tabs_valid(rows,mode)
                report['blocks'][mode].update(status='PASS' if valid else 'FAIL',valid=valid,
                    tab_count=rows[0][4] if valid else None)
            except BaseException as exc:
                report['blocks'][mode]['error']={'type':type(exc).__name__,'message':str(exc)}
                (output/(mode+'-failure.txt')).write_text(traceback.format_exc())
                if owner._quarantined:raise
            flush()
            check(mode+'_saved_text_topology_unchanged',execute(mode+'_after_state',state_commands())==initial)
            check(mode+'_inventory_hash_unchanged',inventory(output,mode+'-after')==initial_inventory)
            check(mode+'_private_bytes_unchanged',sha(owner._private_path)==INPUT_SHA)
            check(mode+'_input_bytes_unchanged',sha(INPUT)==INPUT_SHA)
        report['close_journal']={'private_path':str(owner._private_path),'status':'submitted'};flush()
        owner.close();report['close_journal']['status']='acknowledged'
        check('owned_closed',owner._closed and not owner._quarantined)
        check('final_inventory_empty',inventory(output,'final')==[])
        check('final_input_preserved',sha(INPUT)==INPUT_SHA)
        check('sources_unchanged',all(sha(ROOT/p)==digest for p,digest in report['source_hashes'].items()))
        report['status']='PASS' if all(b['status']=='PASS' for b in report['blocks'].values()) else 'FAIL'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and not owner._closed:
            try:owner.close()
            except BaseException:report['cleanup_failure']=traceback.format_exc();report['status']='FAIL'
        if owner and owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/'native-runtime',dirs_exist_ok=True)
        flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true');parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Read-only tab diagnostic requires --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='PASS' else 1


if __name__=='__main__':raise SystemExit(main())
