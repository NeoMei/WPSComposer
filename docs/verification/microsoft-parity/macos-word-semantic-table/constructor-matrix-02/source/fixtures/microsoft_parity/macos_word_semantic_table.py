"""Preparation-only portrait semantic-table primitive native acceptance.

Requires --execute and a separately granted exclusive Word lease. Uses unchanged
pure primitive commands, exact bound-document sessions and reviewed uncertainty
helpers. No retry, repair, rollback or public semantic-table parity is claimed.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import shutil
import sys
from uuid import uuid4
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.macos_word_semantic_table import (
    build_semantic_table_commands, render_semantic_table_commands,
)
from fixtures.microsoft_parity import macos_word_inline_rule_feasibility as safe
from fixtures.microsoft_parity.macos_word_quality import cleanup, quarantine, record_failure, same
from fixtures.microsoft_parity.macos_word_quality_feasibility import bound_guard, sha, source_pair_matches
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage

DICTIONARY = Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
PREFIX = 'PREFIX 中文😀\r'
SUFFIX = 'SUFFIX preserved 中文😀\r'
OLD = 'EXISTING TABLE 原样'
HEADERS = ['编号 ID', 'Narrative 内容', 'Flag']
ROWS = [['A1', '正文 中文😀', 'V-MERGE'], ['A2', None, ''], ['H-MERGE', '', True]]
# Independent frozen-6dd3a00 oracle values for the above data at 480pt.
WIDTHS = [135.927094400347, 220.8142163128442, 123.25868928680886]
BORDERS = {'top':1.5, 'left':0, 'bottom':.75, 'right':.5,
           'insideHorizontal':0, 'insideVertical':2, 'headerBottom':.75}
NATIVE_BORDERS = [('top',True,1.5),('left',False,0),('bottom',True,.75),
                  ('right',True,.25),('insideHorizontal',False,0),
                  ('insideVertical',True,.25),('headerBottom',True,.75)]
NEW_BOOKMARK = 'wpsc_semantic_fixture_table'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SOURCES = [*sorted((ROOT/'skills/WPSComposer').rglob('*.py')),
           *sorted((ROOT/'fixtures/microsoft_parity').glob('*.py')),
           ROOT/'tests/msoffice/test_macos_word_semantic_table_fixture.py', ROOT/'pyproject.toml']


def _number(value):
    return type(value) in (int, float) and math.isfinite(value)


def _near(value, expected, tolerance=.1):
    return _number(value) and abs(value-expected) <= tolerance


def seed_commands():
    text = PREFIX+'REPLACE\r'+SUFFIX+'OLD TABLE SLOT\r'
    return [
        'activate object boundWindow',
        f'set content of text object of boundDoc to {apple_string(text)}',
        'set orientation of page setup of boundDoc to orient portrait',
        'set page width of page setup of boundDoc to 600',
        'set page height of page setup of boundDoc to 780',
        'set left margin of page setup of boundDoc to 60',
        'set right margin of page setup of boundDoc to 60',
        'set top margin of page setup of boundDoc to 60',
        'set bottom margin of page setup of boundDoc to 60',
        'set fixtureRange to text object of paragraph 2 of boundDoc',
        'set fixtureRange to create range boundDoc start (start of content of fixtureRange) end ((end of content of fixtureRange) - 1)',
        'make new bookmark at boundDoc with properties {name:"semantic_replace",text object:fixtureRange}',
        'make new bookmark at boundDoc with properties {name:"semantic_suffix",text object:text object of paragraph 3 of boundDoc}',
        'set fixtureRange to text object of paragraph 4 of boundDoc',
        'set fixtureRange to create range boundDoc start (start of content of fixtureRange) end ((end of content of fixtureRange) - 1)',
        'set fixtureOldTable to make new table at boundDoc with properties {text object:fixtureRange,number of rows:1,number of columns:1}',
        f'set content of text object of (get cell from table fixtureOldTable row 1 column 1) to {apple_string(OLD)}',
        'make new bookmark at boundDoc with properties {name:"semantic_old_table",text object:text object of fixtureOldTable}',
        'set fixtureRange to text object of bookmark "semantic_replace" of boundDoc',
        'set selection start of selection of boundWindow to start of content of fixtureRange',
        'set selection end of selection of boundWindow to end of content of fixtureRange',
        'set nativeRows to {{"seed",start of content of fixtureRange,end of content of fixtureRange,count tables of boundDoc,'
        'content of text object of bookmark "semantic_old_table" of boundDoc as text}}',
    ]


def _style_commands():
    lines = ['set nativeRows to {{"style",preferred width of semanticTable,allow auto fit of semanticTable,'
             'heading format of row 1 of semanticTable,allow break across pages of row options of semanticTable,'
             '{width of column 1 of semanticTable,width of column 2 of semanticTable,width of column 3 of semanticTable}}}']
    for row in range(1,5):
        for column in range(1,4):
            lines += [f'set fixtureCell to get cell from table semanticTable row {row} column {column}',
                      'set fixtureRange to text object of fixtureCell',
                      f'set end of nativeRows to {{"cell",{row},{column},content of fixtureRange as text,'
                      'first line indent of paragraph format of fixtureRange,paragraph format left indent of paragraph format of fixtureRange,'
                      'paragraph format right indent of paragraph format of fixtureRange,'
                      'my enumIndex(alignment of paragraph format of fixtureRange,{align paragraph left,align paragraph center,align paragraph right})}']
    for key, side in [('top','top'),('left','left'),('bottom','bottom'),('right','right'),
                      ('insideHorizontal','horizontal'),('insideVertical','vertical'),('headerBottom','bottom')]:
        target = 'row 1 of semanticTable' if key == 'headerBottom' else 'semanticTable'
        lines += [f'set fixtureBorder to get border {target} which border border {side}',
                  'set fixtureWidth to 0',
                  'if line style of fixtureBorder is not line style none then',
                  'set fixtureWidth to -1',
                  'if line width of fixtureBorder is line width25 point then set fixtureWidth to 0.25',
                  'if line width of fixtureBorder is line width75 point then set fixtureWidth to 0.75',
                  'if line width of fixtureBorder is line width150 point then set fixtureWidth to 1.5',
                  'end if',
                  f'set end of nativeRows to {{"border","{key}",(line style of fixtureBorder is line style single),fixtureWidth}}']
    return lines


def mutation_commands(path, start, end):
    if any(type(v) is not int or v < 0 for v in (start,end)) or end <= start:
        raise ValueError('Expected explicit noncollapsed native range')
    commands = build_semantic_table_commands(start=start,end=end,headers=HEADERS,rows=ROWS,
        alignments=['left','center','right'],border_spec=BORDERS,repeat_header=True,
        allow_row_split=False,cell_indent_pt=2.5,available_width_pt=480.0,
        merges=[dict(top=2,left=3,bottom=3,right=3),dict(top=4,left=1,bottom=4,right=2)])
    split = next(i for i,c in enumerate(commands) if c.action=='merge')
    return bound_guard(path) + [
        'set fixtureSelection to selection of boundWindow',
        'if story type of fixtureSelection is not main text story then error "SEMANTIC_WRONG_STORY"',
        'set fixtureRange to text object of fixtureSelection',
        f'if start of content of fixtureRange is not {start} or end of content of fixtureRange is not {end} then error "SEMANTIC_RANGE_CHANGED"',
        f'set fixtureRange to create range boundDoc start {start} end {end}',
        'if (content of fixtureRange as text) is not "REPLACE" then error "SEMANTIC_RANGE_TEXT_CHANGED"',
        'if (count sections of boundDoc) is not 1 then error "SEMANTIC_SECTION_CHANGED"',
        'set fixturePage to page setup of section 1 of boundDoc',
        'if orientation of fixturePage is not orient portrait or page width of fixturePage is not 600 or left margin of fixturePage is not 60 or right margin of fixturePage is not 60 then error "SEMANTIC_PAGE_CHANGED"',
        'if (count tables of boundDoc) is not 1 then error "SEMANTIC_TABLE_BASELINE_CHANGED"',
        *render_semantic_table_commands(commands[:split]), *_style_commands(),
        *render_semantic_table_commands(commands[split:]),
        f'make new bookmark at boundDoc with properties {{name:"{NEW_BOOKMARK}",text object:text object of semanticTable}}',
        'set fixtureStyleRows to nativeRows',
        *result_commands(start),
        'set nativeRows to fixtureStyleRows & nativeRows',
    ]


def result_commands(start):
    if type(start) is not int or start < 0: raise ValueError('Expected native start')
    return [
        f'set fixtureBookmark to text object of bookmark "{NEW_BOOKMARK}" of boundDoc',
        'set fixtureMatches to 0',
        'repeat with fixtureIndex from 1 to count tables of boundDoc',
        'set fixtureCandidate to table fixtureIndex of boundDoc',
        'set fixtureRange to text object of fixtureCandidate',
        'if start of content of fixtureRange is start of content of fixtureBookmark and end of content of fixtureRange is end of content of fixtureBookmark then',
        'set fixtureMatches to fixtureMatches + 1', 'set fixtureTable to fixtureCandidate', 'end if', 'end repeat',
        'if fixtureMatches is not 1 then error "SEMANTIC_TABLE_IDENTITY"',
        'set fixtureEnd to end of content of text object of fixtureTable',
        'set fixtureTail to create range boundDoc start fixtureEnd end (fixtureEnd + 1)',
        'set fixtureSelectionRange to text object of selection of boundWindow',
        'set nativeRows to {{"result",posix full name of boundDoc as text,start of content of text object of fixtureTable,fixtureEnd,'
        'end of content of text object of boundDoc,count tables of boundDoc,start of content of fixtureSelectionRange,'
        'end of content of fixtureSelectionRange,content of fixtureTail as text,(story type of fixtureSelectionRange is main text story)}}',
        f'set fixturePrefix to create range boundDoc start 0 end {start}',
        'set fixtureOld to text object of bookmark "semantic_old_table" of boundDoc',
        'set end of nativeRows to {"outside",content of fixturePrefix as text,'
        'content of text object of bookmark "semantic_suffix" of boundDoc as text,content of fixtureOld as text,start of content of fixtureOld}',
        'set end of nativeRows to {"table-text",content of text object of fixtureTable as text}',
    ]


def style_valid(rows):
    try:
        if not isinstance(rows,list) or len(rows)!=20: return False
        style=rows[0]
        if (not isinstance(style,list) or len(style)!=6 or style[0]!='style'
                or not _near(style[1],480) or style[2] is not False or style[3] is not True
                or style[4] is not False or not isinstance(style[5],list) or len(style[5])!=3
                or not all(_near(v,w) for v,w in zip(style[5],WIDTHS))): return False
        expected=[(ri,ci,str(value)) for ri,row in enumerate([HEADERS,*ROWS],1) for ci,value in enumerate(row,1)]
        for observed,(row,column,text) in zip(rows[1:13],expected):
            if (not isinstance(observed,list) or len(observed)!=8 or observed[:3]!=['cell',row,column]
                    or type(observed[1]) is not int or type(observed[2]) is not int
                    or not isinstance(observed[3],str) or observed[3].rstrip('\r\x07')!=text
                    or not _near(observed[4],2.5) or not _near(observed[5],0)
                    or not _near(observed[6],0) or type(observed[7]) is not int or observed[7]!=column-1): return False
        for observed,(key,enabled,width) in zip(rows[13:],NATIVE_BORDERS):
            if (not isinstance(observed,list) or len(observed)!=4 or observed[:2]!=['border',key]
                    or observed[2] is not enabled or not _near(observed[3],width,.01)): return False
        return True
    except (TypeError,IndexError,ValueError): return False


def result_valid(rows,path,start,old_text,*,require_cursor=True):
    try:
        if not isinstance(rows,list) or len(rows)!=3 or any(not isinstance(r,list) for r in rows): return False
        r,out,text=rows
        if len(r)!=10 or len(out)!=5 or len(text)!=2: return False
        if (r[:2]!=['result',path] or any(type(r[i]) is not int for i in range(2,8))
                or not (r[2]==start<r[3]<r[4]) or r[5]!=2 or r[8]!='\r' or r[9] is not True): return False
        if require_cursor and r[6:8]!=[r[3]+1,r[3]+1]: return False
        if (out[:4]!=['outside',PREFIX,SUFFIX,old_text] or type(out[4]) is not int
                or not r[3]<out[4]<r[4] or text[0]!='table-text' or not isinstance(text[1],str)): return False
        return 'REPLACE' not in text[1] and all(text[1].count(value)==1 for value in (
            *HEADERS,'A1','正文 中文😀','V-MERGE','A2','None','H-MERGE','True'))
    except (TypeError,ValueError,IndexError): return False


def _flush(output,report):
    path=output/'report.pending'
    path.write_text(json.dumps(report,ensure_ascii=False,indent=2))
    path.replace(output/'report.json')


def observe_native(owner,output,report,label,commands,validator,*,mutate=False):
    if report.get('native_uncertainty') or report.get('quarantined') or owner._quarantined:
        raise RuntimeError('Native completion uncertain; followup refused')
    report['current_step']=label; _flush(output,report)
    try:
        if mutate: owner._mutation_preflight()
        rows=(owner._execute_structural if mutate else owner._execute)(bound_guard(owner._bound_path)+commands)
        report['steps'].append({'label':label,'rows':rows}); _flush(output,report)
        if not validator(rows): raise AssertionError(label+' native observation invalid')
    except BaseException as error:
        quarantine(output,report,label,error,owner)
        raise
    report['confirmed'].append(label); _flush(output,report)
    return rows


def xml_checks(before,after):
    old,new=ET.fromstring(before),ET.fromstring(after)
    text=lambda n: ''.join(e.text or '' for e in n.iter(W+'t'))
    tables=list(new.iter(W+'tbl'))
    candidates=[t for t in tables if 'Narrative 内容' in text(t)]
    checks=dict(native_table=False,grid_widths=False,borders=False,merge_topology=False,
                cell_paragraphs=False,row_flags=False,outside_preserved=False)
    if len(candidates)!=1: return checks
    table=candidates[0]
    checks['native_table']=len(tables)==2 and not any(list(new.iter(W+t)) for t in ('drawing','pict','object','instrText','fldSimple'))
    try:
        grid=[int(e.get(W+'w')) for e in table.findall('./'+W+'tblGrid/'+W+'gridCol')]
        checks['grid_widths']=len(grid)==3 and all(abs(v/20-w)<=.1 for v,w in zip(grid,WIDTHS))
        borders=table.find('./'+W+'tblPr/'+W+'tblBorders')
        def border_ok(node,width):
            if node is None: return False
            return node.get(W+'val') in ('nil','none') if width==0 else node.get(W+'val')=='single' and int(node.get(W+'sz','-1'))==width
        checks['borders']=borders is not None and all(border_ok(borders.find(W+k),v) for k,v in
            [('top',12),('left',0),('bottom',6),('right',2),('insideH',0),('insideV',2)])
        rows=table.findall(W+'tr'); cells=[row.findall(W+'tc') for row in rows]
        def enabled(node): return node is not None and node.get(W+'val','1') in ('1','true','on')
        checks['row_flags']=(len(rows)==4 and enabled(rows[0].find('./'+W+'trPr/'+W+'tblHeader'))
            and all(enabled(row.find('./'+W+'trPr/'+W+'cantSplit')) for row in rows)
            and all(not enabled(row.find('./'+W+'trPr/'+W+'tblHeader')) for row in rows[1:]))
        checks['borders'] = checks['borders'] and all(border_ok(cell.find('./'+W+'tcPr/'+W+'tcBorders/'+W+'bottom'),6) for cell in cells[0])
        merges=[]
        for ri,row in enumerate(cells):
            for ci,cell in enumerate(row):
                for tag in ('gridSpan','vMerge'):
                    for element in cell.findall('./'+W+'tcPr/'+W+tag): merges.append((ri,ci,tag,element.get(W+'val')))
        checks['merge_topology']=([len(row) for row in cells]==[3,3,3,2]
            and merges==[(1,2,'vMerge','restart'),(2,2,'vMerge',None),(3,0,'gridSpan','2')]
            and [[text(c) for c in row] for row in cells]==[
                HEADERS,['A1','正文 中文😀','V-MERGE'],['A2','None',''],['H-MERGE','True']])
        formats=[]
        for ri,row in enumerate(cells):
            for ci,cell in enumerate(row):
                column=2 if ri==3 and ci==1 else ci
                for paragraph in cell.findall(W+'p'):
                    props=paragraph.find(W+'pPr')
                    ind=props.find(W+'ind') if props is not None else None
                    align=props.find(W+'jc') if props is not None else None
                    formats.append(ind is not None and ind.get(W+'firstLine')=='50'
                        and ind.get(W+'left',ind.get(W+'start'))=='0' and ind.get(W+'right',ind.get(W+'end'))=='0'
                        and align is not None and align.get(W+'val')==['left','center','right'][column])
        checks['cell_paragraphs']=bool(formats) and all(formats)
        def normalized(node):
            return (node.tag,tuple(sorted((k,v) for k,v in node.attrib.items() if not k.split('}')[-1].startswith('rsid')
                and k.split('}')[-1] not in ('paraId','textId'))),node.text or '',tuple(normalized(c) for c in node
                if c.tag not in (W+'bookmarkStart',W+'bookmarkEnd')))
        def terminal_paragraph(original, tail):
            # The one allowed tail is an empty paragraph with the replacement
            # paragraph's exact properties. Empty unformatted runs are equivalent
            # to no run; breaks, fields, section/page properties or other content
            # are never discarded merely because there is no w:t text.
            if tail.tag != W+'p' or normalized(original)[1] != normalized(tail)[1]:
                return False
            def properties(paragraph):
                found=paragraph.findall(W+'pPr')
                if len(found)>1: return ('invalid-duplicate-properties',)
                return normalized(found[0]) if found else None
            if properties(original) != properties(tail): return False
            for child in original:
                if child.tag in (W+'pPr',W+'bookmarkStart',W+'bookmarkEnd'): continue
                if child.tag != W+'r' or any(part.tag not in (W+'rPr',W+'t') for part in child):
                    return False
            for child in tail:
                if child.tag in (W+'pPr',W+'bookmarkStart',W+'bookmarkEnd'): continue
                if child.tag != W+'r' or normalized(child)[1]: return False
                for part in child:
                    if part.tag != W+'t' or part.text or len(part): return False
                    if any(key != '{http://www.w3.org/XML/1998/namespace}space' for key in part.attrib):
                        return False
            return True
        def exact_splice():
            old_body,new_body=old.find(W+'body'),new.find(W+'body')
            if old_body is None or new_body is None: return False
            original,current=list(old_body),list(new_body)
            slots=[i for i,node in enumerate(original) if node.tag==W+'p' and text(node)=='REPLACE']
            if len(slots)!=1 or len(current)!=len(original)+1: return False
            index=slots[0]
            if current[index] is not table or not terminal_paragraph(original[index],current[index+1]):
                return False
            # Preserve every other body node in order, including all existing
            # empty paragraphs and their layout/section/run-level structure.
            before=original[:index]+original[index+1:]
            after=current[:index]+current[index+2:]
            return (normalized(old_body)[1:3]==normalized(new_body)[1:3]
                    and [normalized(node) for node in before]==[normalized(node) for node in after])
        checks['outside_preserved']=(exact_splice()
            and text(new).count('PREFIX')==text(new).count('SUFFIX')==text(new).count(OLD)==1
            and 'REPLACE' not in text(new) and 'OLD TABLE SLOT' not in text(new))
    except (AttributeError,TypeError,ValueError,IndexError): pass
    return checks


def pdf_checks(path,images):
    import fitz
    images=Path(images); images.mkdir(parents=True,exist_ok=False)
    words=[]; drawings=[]; image_count=0; bounds_ok=True
    with fitz.open(path) as pdf:
        for index,page in enumerate(pdf):
            page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(images/f'page-{index+1}.png')
            for word in page.get_text('words'):
                words.append((index,*word[:5])); bounds_ok &= page.rect.contains(fitz.Rect(word[:4]))
            image_count+=len(page.get_images())
            for drawing in page.get_drawings(): drawings.append((index,drawing))
    text=' '.join(str(w[5]) for w in words)
    def first(token): return next((w for w in words if token in w[5]),None)
    prefix,suffix,old=first('PREFIX'),first('SUFFIX'),first('EXISTING')
    order=bool(prefix and suffix and old and (prefix[0],prefix[2])<(suffix[0],suffix[2])<(old[0],old[2]))
    required=('PREFIX','SUFFIX','EXISTING','ID','Narrative','Flag','A1','A2','V-MERGE','H-MERGE','None','True')
    tokens_ok=all(text.count(token)==1 for token in required) and 'REPLACE' not in text
    if order:
        order=all(first(token) is not None and (prefix[0],prefix[2]) < (first(token)[0],first(token)[2]) < (suffix[0],suffix[2]) for token in required[3:])
    geometry=False
    if prefix and suffix and prefix[0]==suffix[0]:
        for page,drawing in drawings:
            rect=drawing['rect']; thickness=drawing.get('width',0) if drawing.get('type')=='s' else rect.height
            if (page==prefix[0] and prefix[4]<rect.y0<=rect.y1<suffix[2]
                    and _near(rect.width,480,2) and any(_near(thickness,w,.1) for w in (1.5,.75))): geometry=True
    (images/'observations.json').write_text(json.dumps({'text':text,'drawings':drawings},default=str,ensure_ascii=False,indent=2))
    return {'pdf_text_order':order and tokens_ok,'pdf_table_geometry':geometry,
            'pdf_no_images':image_count==0,'pdf_text_inside_pages':bool(words) and bounds_ok}


def retain_runtime(owner,output,report,label):
    if owner is None: return
    report[label+'-identity']={'path':owner._bound_path,'closed':owner._closed,'quarantined':owner._quarantined,'staging_root':str(owner.staging_root)}
    try:
        if owner.staging_root.exists(): shutil.copytree(owner.staging_root,output/label,dirs_exist_ok=True)
    except BaseException as error:
        record_failure(output,report,error,label+'-retention')


def run(output,*,execute=False):
    if execute is not True: raise ValueError('Explicit execute=True and separate root Word lease required')
    output=Path(output).resolve(); output.mkdir(parents=True,exist_ok=False)
    report={'status':'FAIL','scope':'portrait middle-range semantic table primitive only','checks':{},'steps':[],'confirmed':[],
            'pending':['native execution','visual PNG inspection','fault rollback','landscape','controller fallback','public semantic method parity']}
    session=reopened=None
    def call(label,fn,owner=None):
        if report.get('quarantined') or report.get('native_uncertainty'): raise RuntimeError('Native followup refused')
        report['current_step']=label; _flush(output,report)
        try: return fn()
        except BaseException as error:
            quarantine(output,report,label,error,owner); raise
    try:
        report['source_hashes']=retain_sources(output,SOURCES)
        report['dictionary_sha256']=sha(DICTIONARY); shutil.copy2(DICTIONARY,output/'localWord.sdef')
        report['inventory_before']=safe.independent_inventory(output,'before',report)
        session=call('open-owned',lambda:MacWordSession.new_document(visible=False)); session._retain_evidence=True
        start=len(PREFIX.encode('utf-16-le'))//2; end=start+7
        seed=observe_native(session,output,report,'seed',seed_commands(),lambda r:
            isinstance(r,list) and len(r)==1 and isinstance(r[0],list) and len(r[0])==5
            and same(r[0][:4],['seed',start,end,1]) and isinstance(r[0][4],str) and OLD in r[0][4],mutate=True)
        old_text=seed[0][4]
        call('save-before',lambda:session.save_docx(output/'before.docx'),session)
        report['before_sha256']=sha(output/'before.docx')
        token='SEMANTIC TABLE SENTINEL 中文😀 '+uuid4().hex; report['sentinel_token']=token
        sentinel=observe_native(session,output,report,'sentinel-create',[
            'set qualitySentinel to make new document',f'set content of text object of qualitySentinel to {apple_string(token)}',
            'set nativeRows to {{name of qualitySentinel as text,version as text}}'],lambda r:
            isinstance(r,list) and len(r)==1 and isinstance(r[0],list) and len(r[0])==2 and all(isinstance(v,str) and v for v in r[0]))
        report['sentinel_name'],report['word_version']=sentinel[0]
        report['sentinel_before']=sentinel_preimage(safe.independent_inventory(output,'sentinel-before',report,session),report['sentinel_name'],token)
        # Explicitly restore only the bound document's selection after native Save.
        observe_native(session,output,report,'select-middle',[
            f'set selection start of selection of boundWindow to {start}',f'set selection end of selection of boundWindow to {end}',
            'set nativeRows to {{"selected",true}}'],lambda r:same(r,[['selected',True]]))
        rows=observe_native(session,output,report,'semantic-table',mutation_commands(session._bound_path,start,end),
            lambda r:style_valid(r[:20]) and result_valid(r[20:],session._bound_path,start,old_text),mutate=True)
        report['checks']['native_style_and_middle_bounds']=True
        call('save-after',lambda:session.save_docx(output/'after.docx'),session)
        call('export-pdf',lambda:session.export_pdf(output/'after.pdf'),session)
        report['checks']['sentinel_unchanged']=same(report['sentinel_before'],sentinel_preimage(
            safe.independent_inventory(output,'sentinel-after',report,session),report['sentinel_name'],token))
        if not report['checks']['sentinel_unchanged']: raise AssertionError('Sentinel changed')
        call('close-owned',session.close,session)
        safe.close_sentinel_after_owned(session,output,report,report['sentinel_name'],token)
        digest=sha(output/'after.docx')
        report['reopen_requested_path']=str(output/'after.docx')
        reopened=call('open-readonly',lambda:MacWordSession.open_document(output/'after.docx',read_only=True,visible=False))
        reopened._retain_evidence=True
        again=observe_native(reopened,output,report,'reopen-readback',result_commands(start),
            lambda r:result_valid(r,reopened._bound_path,start,old_text,require_cursor=False))
        original=rows[20:]
        report['checks']['native_reopen_preserved']=same([again[0][2:6],again[0][8:],*again[1:]],
            [original[0][2:6],original[0][8:],*original[1:]])
        call('close-reopened',reopened.close,reopened)
        report['checks']['reopen_source_preserved']=sha(output/'after.docx')==digest and reopened._closed and not reopened._quarantined
        report['checks'].update(xml_checks(safe.retain_xml(output,'before'),safe.retain_xml(output,'after')))
        report['checks'].update(pdf_checks(output/'after.pdf',output/'pdf-pages'))
    except BaseException as error:
        record_failure(output,report,error)
        if (reopened or session) and not (reopened or session)._closed:
            quarantine(output,report,report.get('current_step','run'),error,reopened or session)
    finally:
        cleanup(reopened or session,output,report)
        for owner,label in ((session,'native-runtime'),(reopened,'reopen-runtime')):
            retain_runtime(owner,output,report,label)
        try:
            if 'source_hashes' in report: report['checks']['sources_unchanged']=all(source_pair_matches(ROOT/p,output/'source'/p,v) for p,v in report['source_hashes'].items())
            if 'dictionary_sha256' in report: report['checks']['dictionary_unchanged']=source_pair_matches(DICTIONARY,output/'localWord.sdef',report['dictionary_sha256'])
            if 'before_sha256' in report: report['checks']['before_file_preserved']=sha(output/'before.docx')==report['before_sha256']
            report['artifact_hashes']={str(p.relative_to(output)):sha(p) for p in output.rglob('*') if p.is_file()
                and 'source' not in p.relative_to(output).parts and p.name not in ('report.json','report.pending')}
        except BaseException as error: record_failure(output,report,error,'evidence')
        required=('native_style_and_middle_bounds','sentinel_unchanged','owned_closed_with_sentinel_preserved',
            'native_reopen_preserved','reopen_source_preserved','native_table','grid_widths','borders','merge_topology',
            'cell_paragraphs','row_flags','outside_preserved','pdf_text_order','pdf_table_geometry','pdf_no_images','pdf_text_inside_pages',
            'sources_unchanged','dictionary_unchanged','before_file_preserved','inventory_preserved')
        if (all(report['checks'].get(k) is True for k in required) and not report.get('quarantined') and not report.get('remaining_sentinel')
                and not any(k=='error' or k.endswith(('_error','_failure')) for k in report)):
            report['status']='PASS_PRIMITIVE_AUTOMATED'; report['pending'].remove('native execution')
        _flush(output,report)
    return report


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('--execute',action='store_true'); parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Requires --execute and a separately granted root Word lease',file=sys.stderr); return 2
    return 0 if run(args.output,execute=True)['status']=='PASS_PRIMITIVE_AUTOMATED' else 1


if __name__=='__main__': raise SystemExit(main())
