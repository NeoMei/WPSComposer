"""Read-only mixed-script Font aggregation diagnostic; never appearance acceptance.

Read exact source/imported style, whole paragraph and one Unicode scalar per
Latin/CJK/Arabic/emoji script. Empty family names remain recorded differences.
No fresh paragraph is created; the failed v2 fixture/evidence is preserved.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import shutil
import sys
import traceback
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession,apple_string,sha,inventory,retain_sources,CLONE,SOURCE,IMPORTED,SOURCES,
)
INPUT=Path(__file__).parent/'native-donor-import-02/after-import.docx'
INPUT_SHA='52d724bb135c5f1cf7e4d4d41472341010a8e6e8b91cca4a17d13a4c75e63c39'
FIELDS=('ascii name','complex script name','name','other name','east asian name')


def anchor_ranges(body,start,end,marker):
    if not isinstance(body,str) or type(start) is not int or type(end) is not int:raise ValueError('Invalid native preimage')
    raw=body.encode('utf-16-le')
    if not 0<=start<end<=len(raw)//2:raise ValueError('Invalid native bounds')
    try:paragraph=raw[start*2:end*2].decode('utf-16-le')
    except UnicodeDecodeError:raise ValueError('Native range splits surrogate') from None
    if paragraph!=marker+'\r':raise ValueError('Unexpected native paragraph')
    result=[]
    for label,scalar in [('latin','A'),('cjk','中'),('arabic','ا'),('emoji','😀')]:
        offset=paragraph.index(scalar)
        left=start+len(paragraph[:offset].encode('utf-16-le'))//2
        right=left+len(scalar.encode('utf-16-le'))//2
        result.append({'script':label,'scalar':scalar,'start':left,'end':right})
    return result


def all_anchors(body,s1,e1,s2,e2):
    return {owner:[{'script':'paragraph','scalar':marker+'\r','start':start,'end':end},
                    *anchor_ranges(body,start,end,marker)]
            for owner,start,end,marker in [('source',s1,e1,SOURCE),('clone',s2,e2,IMPORTED)]}


def specs(anchors):
    return [(owner,'style','',-1,-1) if row is None else (owner,row['script'],row['scalar'],row['start'],row['end'])
            for owner in ('source','clone') for row in [None,*anchors[owner]]]


def font_commands(body,anchors):
    lines=[f'if not (((current application\'s NSString\'s stringWithString:(content of text object of boundDoc as text))\'s isEqualToString:{apple_string(body)}) as boolean) then error "WPSC_FONT_FULL_PREIMAGE"',
           'set sourceStyle to Word style (style heading1) of boundDoc',
           f'set cloneStyle to Word style {apple_string(CLONE)} of boundDoc',
           'set sourceName to name local of sourceStyle as text','set cloneName to name local of cloneStyle as text',
           'set nativeRows to {{"styles",sourceName,cloneName}}']
    for owner,scope,scalar,start,end in specs(anchors):
        if scope=='style':
            lines += [f'set observedFont to font object of {owner}Style',f'set observedName to {owner}Name',
                      'set observedText to ""','set observedStart to -1','set observedEnd to -1']
        else:
            lines += [f'set targetRange to create range boundDoc start {start} end {end}',
                      'set observedStart to start of content of targetRange','set observedEnd to end of content of targetRange',
                      f'if observedStart is not {start} or observedEnd is not {end} then error "WPSC_FONT_ANCHOR_BOUNDS"',
                      'set observedText to content of targetRange as text',
                      f'if not (((current application\'s NSString\'s stringWithString:observedText)\'s isEqualToString:{apple_string(scalar)}) as boolean) then error "WPSC_FONT_ANCHOR_TEXT"',
                      'set observedName to name local of style of targetRange as text',
                      f'if observedName is not {owner}Name then error "WPSC_FONT_ANCHOR_STYLE"',
                      'set observedFont to font object of targetRange']
        values=','.join(field+' of observedFont as text' for field in FIELDS)
        lines.append(f'set end of nativeRows to {{"font",{apple_string(owner)},{apple_string(scope)},observedText,observedStart,observedEnd,observedName,{values}}}')
    return lines


def rows_valid(rows,anchors):
    expected=specs(anchors)
    if (not isinstance(rows,list) or len(rows)!=1+len(expected) or not isinstance(rows[0],list)
            or len(rows[0])!=3 or rows[0][0]!='styles' or not isinstance(rows[0][1],str)
            or not rows[0][1] or rows[0][2]!=CLONE):return False
    names={'source':rows[0][1],'clone':rows[0][2]}
    for row,(owner,scope,scalar,start,end) in zip(rows[1:],expected):
        if (not isinstance(row,list) or len(row)!=12
                or row[:7]!=['font',owner,scope,scalar,start,end,names[owner]]
                or type(row[4]) is not int or type(row[5]) is not int
                or any(not isinstance(v,str) for v in row[7:])):return False
    return True


def compare_fields(style,actual):
    diffs={key:{'style':left,'range':right} for key,left,right in zip(FIELDS,style,actual) if left!=right}
    return {'equal':not diffs,'differences':diffs}


def observations(rows):
    by_owner={owner:{} for owner in ('source','clone')}
    for row in rows[1:]:by_owner[row[1]][row[2]]=row[7:]
    return {owner:{scope:compare_fields(values['style'],actual) for scope,actual in values.items() if scope!='style'}
            for owner,values in by_owner.items()}


def state_commands():
    return ['set firstRange to text object of paragraph 1 of boundDoc','set secondRange to text object of paragraph 2 of boundDoc',
            'set nativeRows to {{"state",saved of boundDoc,content of text object of boundDoc as text,start of content of firstRange,end of content of firstRange,start of content of secondRange,end of content of secondRange,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}']


def state_valid(rows):
    return (isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list) and len(rows[0])==10
            and rows[0][0]=='state' and type(rows[0][1]) is bool and rows[0][1]
            and isinstance(rows[0][2],str) and all(type(v) is int and v>=0 for v in rows[0][3:]))


def run(output):
    report={'status':'FAIL','scope':'Read-only font aggregation observations; no appearance acceptance or fresh-paragraph proof',
            'checks':{},'source_hashes':retain_sources(output,[Path(__file__),Path(__file__).with_name('test_font_anchor_read_diagnostic.py'),*SOURCES])}
    owner=None
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def check(label,value):
        report['checks'][label]=bool(value);flush()
        if not value:raise AssertionError('Font anchor diagnostic failed: '+label)
    def execute(label,commands):
        report['current_phase']=label;flush();rows=owner._execute(commands)
        report.setdefault('raw_rows',{})[label]=rows;flush();return rows
    try:
        check('pinned_input',sha(INPUT)==INPUT_SHA)
        check('empty_inventory',inventory(output,'before')==[])
        report['open_journal']={'source':str(INPUT),'status':'submitted'};flush()
        owner=MacWordSession.open_document(INPUT,read_only=True,visible=False);owner._retain_evidence=True
        report['open_journal'].update(status='acknowledged',private_path=str(owner._private_path));flush()
        before_inventory=inventory(output,'after-open');report['inventory_after_open']=before_inventory
        check('exact_private_inventory',len(before_inventory)==1 and before_inventory[0][1]==str(owner._private_path))
        before=execute('before_state',state_commands());check('valid_saved_preimage',state_valid(before))
        check('private_preimage_bytes',sha(owner._private_path)==INPUT_SHA)
        anchors=all_anchors(before[0][2],*before[0][3:7]);report['anchors']=anchors;flush()
        rows=execute('font_ranges',font_commands(before[0][2],anchors))
        check('typed_complete_readback',rows_valid(rows,anchors))
        report['observations']=observations(rows);flush()
        check('saved_text_ranges_topology_unchanged',execute('after_state',state_commands())==before)
        check('inventory_saved_hash_unchanged',inventory(output,'after-read')==before_inventory)
        check('private_bytes_unchanged',sha(owner._private_path)==INPUT_SHA)
        report['close_journal']={'private_path':str(owner._private_path),'status':'submitted'};flush()
        owner.close();report['close_journal']['status']='acknowledged'
        check('owned_closed',owner._closed and not owner._quarantined)
        check('final_inventory_empty',inventory(output,'final')==[])
        check('original_input_unchanged',sha(INPUT)==INPUT_SHA)
        check('sources_unchanged',all(sha(ROOT/p)==digest for p,digest in report['source_hashes'].items()))
        report['status']='DIAGNOSTIC_COMPLETE'
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
        print('Read-only font diagnostic requires --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    return 0 if run(output)['status']=='DIAGNOSTIC_COMPLETE' else 1


if __name__=='__main__':raise SystemExit(main())
