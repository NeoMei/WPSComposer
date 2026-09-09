"""Isolated heading primitives, not production support; native requires --execute.

Each mode uses a fresh owned document. Whole-property and paragraph-object copy
are unproven hypotheses; failure is retained, never masked by scalar fallback.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage, close_after_owned

MODES=('font-properties','paragraph-object','detached-properties','list-identity')
CLONE='WPSC Heading Clone'
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SOURCES=[Path(__file__),ROOT/'tests/msoffice/test_macos_word_heading_feasibility.py',
         *[ROOT/('skills/WPSComposer/scripts/msoffice/'+p) for p in
           ('macos_word_session.py','macos_script.py','macos_runtime.py','macos_word_recovery.py')],
         *[ROOT/('fixtures/microsoft_parity/'+p) for p in
           ('macos_word_paragraph_rule.py','macos_word_recovery.py','macos_word_fields.py')],
         ROOT/'skills/WPSComposer/scripts/writer.py']

FONT=('all caps', 'animation', 'ascii name', 'bold', 'bold bi', 'color', 'color index', 'color theme index', 'complex script name', 'disable character space grid', 'double strike through', 'east asian name', 'emboss', 'emphasis mark', 'engrave', 'font position', 'font size', 'hidden', 'italic', 'italic bi', 'kerning', 'name', 'other name', 'outline', 'scaling', 'shadow', 'small caps', 'spacing', 'strike through', 'subscript', 'superscript', 'underline', 'underline color', 'underline color theme index')

PARAGRAPH=('add space between east asian and alpha', 'add space between east asian and digit', 'alignment', 'auto adjust right indent', 'base line alignment', 'character unit first line indent', 'character unit left indent', 'character unit right indent', 'disable line height grid', 'east asian line break control', 'first line indent', 'half width punctuation on top of line', 'hanging punctuation', 'hyphenation', 'keep together', 'keep with next', 'line spacing', 'line spacing rule', 'line unit after', 'line unit before', 'no line number', 'outline level', 'page break before', 'paragraph format left indent', 'paragraph format right indent', 'space after', 'space after auto', 'space before', 'space before auto', 'style', 'widow control', 'word wrap')

SHADING=('background pattern color', 'background pattern color index', 'background pattern color theme index', 'foreground pattern color', 'foreground pattern color index', 'foreground pattern color theme index', 'texture')

BORDERS=('always in front', 'distance from', 'distance from bottom', 'distance from left', 'distance from right', 'distance from top', 'enable borders', 'enable first page in section', 'enable other pages in section', 'inside color', 'inside color index', 'inside color theme index', 'inside line style', 'inside line width', 'join borders', 'outside color', 'outside color index', 'outside color theme index', 'outside line style', 'outside line width', 'shadow', 'surround footer', 'surround header')


def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def bindings():
    return ['set sourceStyle to Word style (style heading1) of boundDoc',
            f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
            'set sourceFont to font object of sourceStyle','set destinationFont to font object of detachedStyle',
            'set sourceParagraph to paragraph format of sourceStyle','set destinationParagraph to paragraph format of detachedStyle']


def seed_commands(mode):
    if mode not in MODES:raise ValueError('Unknown mode')
    lines=['activate object boundWindow',
           'set sourceStyle to Word style (style heading1) of boundDoc',
           f'set detachedStyle to make new Word style at boundDoc with properties {{name local:{apple_string(CLONE)}}}',
           'set base style of detachedStyle to style normal',
           'set sourceFont to font object of sourceStyle','set sourceParagraph to paragraph format of sourceStyle']
    if mode in ('font-properties','detached-properties'):
        lines += ['set name of sourceFont to "Arial"','set ascii name of sourceFont to "Arial"',
                  'set font size of sourceFont to 17','set bold of sourceFont to true','set italic of sourceFont to true',
                  'set scaling of sourceFont to 110','set spacing of sourceFont to 1.25','set kerning of sourceFont to 12',
                  'set color of sourceFont to {39936,0,1536}',
                  'set background pattern color of shading of sourceFont to {64512,59392,58880}',
                  'set outside line style of border options of sourceFont to line style single',
                  'set outside color of border options of sourceFont to {49344,49344,49344}']
    if mode in ('paragraph-object','detached-properties'):
        lines += ['set alignment of sourceParagraph to align paragraph center','set space before of sourceParagraph to 14',
                  'set space after of sourceParagraph to 7','set first line indent of sourceParagraph to 4',
                  'set keep with next of sourceParagraph to true','set widow control of sourceParagraph to false',
                  'set line spacing rule of sourceParagraph to line space exactly','set line spacing of sourceParagraph to 23',
                  'set background pattern color of shading of sourceParagraph to {58880,62208,65280}',
                  'set outside line style of border options of sourceParagraph to line style single',
                  'set outside line width of border options of sourceParagraph to line width75 point',
                  'set outside color of border options of sourceParagraph to {32768,16384,8192}',
                  'make new tab stop at sourceParagraph with properties {tab stop position:40,alignment:align tab right}']
    lines += ['set content of text object of boundDoc to "SOURCE APPEARANCE 中文😀" & return & "DETACHED APPEARANCE 中文😀" & return',
              'set style of text object of paragraph 1 of boundDoc to sourceStyle',
              'set style of text object of paragraph 2 of boundDoc to detachedStyle',
              'set nativeRows to {{"stage",true}}']
    return lines


def clone_commands(mode):
    if mode not in MODES:raise ValueError('Unknown mode')
    lines=bindings()
    if mode in ('font-properties','detached-properties'):
        lines.append('set properties of destinationFont to properties of sourceFont')
    if mode in ('paragraph-object','detached-properties'):
        lines.append('set paragraph format of detachedStyle to paragraph format of sourceStyle')
    lines += ['set base style of detachedStyle to style normal','set nativeRows to {{"stage",true}}']
    return lines


def dimensions(mode):
    result=[]
    if mode in ('font-properties','detached-properties'):
        for domain,left,right,props in [('font','sourceFont','destinationFont',FONT),
            ('font-shading','shading of sourceFont','shading of destinationFont',SHADING),
            ('font-borders','border options of sourceFont','border options of destinationFont',BORDERS)]:
            result += [(domain+':'+p,p+' of '+left,p+' of '+right) for p in props]
    if mode in ('paragraph-object','detached-properties'):
        for domain,left,right,props in [('paragraph','sourceParagraph','destinationParagraph',PARAGRAPH),
            ('paragraph-shading','shading of sourceParagraph','shading of destinationParagraph',SHADING),
            ('paragraph-borders','border options of sourceParagraph','border options of destinationParagraph',BORDERS)]:
            # Style identity intentionally differs; every other writable scalar is compared.
            result += [(domain+':'+p,p+' of '+left,p+' of '+right) for p in props if p!='style']
    return result


def expected_labels(mode):
    if mode not in MODES:raise ValueError('Unknown mode')
    if mode=='list-identity':return ['list-unique','list-outline']+['level-'+str(i) for i in range(1,5)]
    return [v[0] for v in dimensions(mode)]+(['paragraph:tabs'] if mode in ('paragraph-object','detached-properties') else [])+['clone-base-normal']


def native_valid(mode,rows):return valid_labels(expected_labels(mode),rows)


def valid_labels(labels,rows):
    return (isinstance(rows,list) and len(rows)==len(labels)
            and all(isinstance(row,list) and len(row)==2 and row[0]==label and row[1] is True for label,row in zip(labels,rows)))


def clone_readback(mode):
    lines=bindings()+['set nativeRows to {}']
    for label,left,right in dimensions(mode):
        # Text values must not use AppleScript's default case-insensitive comparison.
        lines += [f'set leftValue to {left}',f'set rightValue to {right}',
                  'if class of leftValue is text then',
                  "set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean",
                  'else','set sameValue to (leftValue is equal to rightValue)','end if',
                  f'set end of nativeRows to {{{apple_string(label)},sameValue}}']
    if mode in ('paragraph-object','detached-properties'):
        lines += ['set tabsEqual to ((count tab stops of sourceParagraph) is (count tab stops of destinationParagraph))',
                  'if tabsEqual then','repeat with i from 1 to count tab stops of sourceParagraph',
                  'set sourceTab to tab stop i of sourceParagraph', 'set destinationTab to tab stop i of destinationParagraph',
                  'if (alignment of sourceTab is not alignment of destinationTab) or (tab stop position of sourceTab is not tab stop position of destinationTab) or (tab leader of sourceTab is not tab leader of destinationTab) or (custom tab of sourceTab is not custom tab of destinationTab) then set tabsEqual to false',
                  'end repeat','end if','set end of nativeRows to {"paragraph:tabs",tabsEqual}']
    lines += ['set expectedBase to name local of Word style (style normal) of boundDoc as text',
              'set observedBase to name local of Word style (base style of detachedStyle) of boundDoc as text',
              "set baseEqual to ((current application's NSString's stringWithString:observedBase)'s isEqualToString:expectedBase) as boolean",
              'set end of nativeRows to {"clone-base-normal",baseEqual}']
    return lines


def check_name(name):
    if not isinstance(name,str) or not re.fullmatch(r'WPSCHeading_[0-9a-f]{32}',name):
        raise ValueError('Template name must be owned UUID')


def resolve_list(name):
    check_name(name)
    return ['set matches to 0','repeat with i from 1 to count list templates of boundDoc',
            'set candidateList to list template i of boundDoc',
            f"if ((current application's NSString's stringWithString:(name of candidateList as text))'s isEqualToString:{apple_string(name)}) as boolean then",
            'set matches to matches + 1','set ownList to candidateList','end if','end repeat',
            'if matches is not 1 then error "WPSC_LIST_IDENTITY_NOT_UNIQUE"']


def list_commands(name):
    check_name(name)
    lines=[f'set ownList to make new list template at boundDoc with properties {{name:{apple_string(name)},outline numbered:true}}']
    for level in range(1,5):
        pattern='.'.join('%'+str(i) for i in range(1,level+1))
        lines += [f'set ownLevel to list level {level} of ownList','set number style of ownLevel to list number style arabic',
                  f'set number format of ownLevel to {apple_string(pattern)}',f'set number position of ownLevel to {(level-1)*18}',
                  f'set text position of ownLevel to {level*18}',f'set reset on higher of ownLevel to {level-1}',
                  'set start at of ownLevel to 1',f'set linked style of ownLevel to name local of Word style (style heading{level}) of boundDoc as text']
    return lines+['set nativeRows to {{"stage",true}}']


def list_readback(name):
    lines=resolve_list(name)+['set nativeRows to {{"list-unique",matches is 1},{"list-outline",outline numbered of ownList}}']
    for level in range(1,5):
        pattern='.'.join('%'+str(i) for i in range(1,level+1))
        lines += [f'set ownLevel to list level {level} of ownList',
                  f'set numberedRange to text object of paragraph {level+2} of boundDoc',
                  'set observedListString to list string of list format of numberedRange as text',
                  f'set actualNumberedLevel to (list level number of list format of numberedRange is {level})',
                  f'set namedStyle to Word style (style heading{level}) of boundDoc',
                  'set styleListName to name of list template of namedStyle as text',
                  f"set identityEqual to ((current application's NSString's stringWithString:styleListName)'s isEqualToString:{apple_string(name)}) as boolean",
                  "set linkedEqual to ((current application's NSString's stringWithString:(linked style of ownLevel as text))'s isEqualToString:(name local of namedStyle as text)) as boolean",
                  f"set formatEqual to ((current application's NSString's stringWithString:(number format of ownLevel as text))'s isEqualToString:{apple_string(pattern)}) as boolean",
                  f'set end of nativeRows to {{"level-{level}",identityEqual and linkedEqual and formatEqual and actualNumberedLevel and (observedListString is not "") and (number style of ownLevel is list number style arabic) and (number position of ownLevel is {(level-1)*18}) and (text position of ownLevel is {level*18}) and (reset on higher of ownLevel is {level-1}) and (start at of ownLevel is 1)}}']
    return lines


def challenge_commands(name):
    # A reversible nonce observed through an independently resolved style path.
    return resolve_list(name)+['set number format of list level 4 of ownList to "PROBE-%1.%2.%3.%4"',
        'set linkedTemplate to list template of Word style (style heading4) of boundDoc',
        "set challengeEqual to ((current application's NSString's stringWithString:(number format of list level 4 of linkedTemplate as text))'s isEqualToString:\"PROBE-%1.%2.%3.%4\") as boolean",
        'set number format of list level 4 of ownList to "%1.%2.%3.%4"',
        'set nativeRows to {{"same-native-list-challenge",challengeEqual}}']


def list_paragraphs():
    lines=[]
    for level in range(1,5):
        lines += ['set p to (end of content of text object of boundDoc) - 1','set r to create range boundDoc start p end p',
                  f'set content of r to "NUMBERED L{level}" & return',
                  f'set r to create range boundDoc start p end (p + 12)',
                  f'set style of r to style heading{level}']
    return lines+['set nativeRows to {{"stage",true}}']


def detached_readback():
    return [f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
            'set detachedRange to text object of paragraph 2 of boundDoc',
            "set detachedTextEqual to ((current application's NSString's stringWithString:(content of detachedRange as text))'s isEqualToString:(\"DETACHED APPEARANCE 中文😀\" & return)) as boolean",
            'set nativeRows to {{"detached-no-number",(list string of list format of detachedRange as text) is ""},'
            '{"detached-outline",outline level of paragraph format of detachedRange is outline level1},'
            '{"detached-text",detachedTextEqual}}']


def phases(mode,name):
    if mode not in MODES:raise ValueError('Unknown mode')
    if mode=='list-identity':return [seed_commands(mode),list_commands(name),list_paragraphs(),list_readback(name),challenge_commands(name)]
    result=[seed_commands(mode),clone_commands(mode),clone_readback(mode)]
    if mode=='detached-properties':result += [list_commands(name),list_paragraphs(),detached_readback(),list_readback(name)]
    return result


def canonical(element):
    if element is None:return None
    return (element.tag,tuple(sorted(element.attrib.items())),element.text or '',tuple(canonical(c) for c in element))


def clone_xml_valid(xml,mode):
    if mode not in ('font-properties','paragraph-object','detached-properties'):return False
    root=ET.fromstring(xml);styles=root.findall(W+'style')
    source=[s for s in styles if s.get(W+'styleId')=='Heading1']
    clone=[s for s in styles if s.find(W+'name') is not None and s.find(W+'name').get(W+'val')==CLONE]
    if len(source)!=1 or len(clone)!=1:return False
    base=clone[0].find(W+'basedOn')
    if base is None or base.get(W+'val')!='Normal':return False
    tags=['rPr'] if mode=='font-properties' else ['pPr'] if mode=='paragraph-object' else ['rPr','pPr']
    # Exact full structures, including nested shading, borders and tabs; no subset pass.
    return all(source[0].find(W+tag) is not None and canonical(source[0].find(W+tag))==canonical(clone[0].find(W+tag)) for tag in tags)


def list_xml_valid(styles_xml,numbering_xml):
    styles=ET.fromstring(styles_xml);numbering=ET.fromstring(numbering_xml)
    ids=[]
    for level in range(1,5):
        matches=[s for s in styles.findall(W+'style') if s.get(W+'styleId')=='Heading'+str(level)]
        if len(matches)!=1:return False
        num=matches[0].find('./'+W+'pPr/'+W+'numPr/'+W+'numId')
        if num is None:return False
        ids.append(num.get(W+'val'))
    if len(set(ids))!=1:return False
    nums=[n for n in numbering.findall(W+'num') if n.get(W+'numId')==ids[0]]
    if len(nums)!=1:return False
    abstract=nums[0].find(W+'abstractNumId')
    if abstract is None:return False
    definitions=[n for n in numbering.findall(W+'abstractNum') if n.get(W+'abstractNumId')==abstract.get(W+'val')]
    if len(definitions)!=1:return False
    for level in range(4):
        matches=[x for x in definitions[0].findall(W+'lvl') if x.get(W+'ilvl')==str(level)]
        if len(matches)!=1:return False
        lvl=matches[0]
        expected={'pStyle':'Heading'+str(level+1),'lvlText':'.'.join('%'+str(i) for i in range(1,level+2)),'start':'1','numFmt':'decimal'}
        if any(lvl.find(W+k) is None or lvl.find(W+k).get(W+'val')!=v for k,v in expected.items()):return False
    return True


def artifact_checks(output,mode):
    with ZipFile(output/'heading.docx') as z:
        styles=z.read('word/styles.xml');(output/'styles.xml').write_bytes(styles)
        doc=z.read('word/document.xml');(output/'document.xml').write_bytes(doc)
        numbers=z.read('word/numbering.xml') if 'word/numbering.xml' in z.namelist() else None
    checks={}
    if mode!='list-identity':
        with ZipFile(output/'pre-list.docx') as z:before=z.read('word/styles.xml')
        checks['complete_clone_xml_before_list']=clone_xml_valid(before,mode)
        before_root=ET.fromstring(before);after_root=ET.fromstring(styles)
        def named(root):return [s for s in root.findall(W+'style') if s.find(W+'name') is not None and s.find(W+'name').get(W+'val')==CLONE]
        a,b=named(before_root),named(after_root)
        checks['detached_style_persisted_unchanged']=len(a)==len(b)==1 and canonical(a[0])==canonical(b[0])
    if mode in ('detached-properties','list-identity'):
        checks['four_levels_share_numbering_xml']=numbers is not None and list_xml_valid(styles,numbers)
        if numbers is not None:(output/'numbering.xml').write_bytes(numbers)
    import fitz
    with fitz.open(output/'heading.pdf') as pdf:
        text='\n'.join(p.get_text() for p in pdf)
        for i,p in enumerate(pdf):p.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(output/('page-%s.png'%(i+1)))
    (output/'pdf-text.txt').write_text(text)
    checks['pdf_probe_text_visible']='SOURCE APPEARANCE' in text and 'DETACHED APPEARANCE' in text
    if mode in ('detached-properties','list-identity'):checks['pdf_numbered_paragraphs_visible']=all('NUMBERED L'+str(i) in text for i in range(1,5))
    return checks


def run(output,mode):
    if mode not in MODES:raise ValueError('Unknown mode')
    report={'status':'FAIL','mode':mode,'checks':{},'scope':'native feasibility only; full direct method and UI parity unproven',
            'source_hashes':retain_sources(output,SOURCES),'list_name':'WPSCHeading_'+uuid4().hex}
    session=None;reopened=None;sentinel_name=None;token=None
    try:
        report['dictionary_sha256']=sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before']=inventory(output,'before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence=True
            token='HEADING FEASIBILITY SENTINEL '+uuid4().hex+' 中文😀'
            rows=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])
            if len(rows)!=1 or len(rows[0])!=1 or not isinstance(rows[0][0],str):raise RuntimeError('Sentinel creation ACK invalid')
            sentinel_name=rows[0][0];report['sentinel_name']=sentinel_name
            report['sentinel_before']=sentinel_preimage(inventory(output,'sentinel-before'),sentinel_name,token)
            report['word_version']=session._execute(['set nativeRows to {{version as text}}'])
            commands=phases(mode,report['list_name'])
            report['phase_rows']=[]
            for i,phase in enumerate(commands):
                report['current_phase']=i
                (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
                # Each bounded phase retains exact script/result. No fallback/retry on failure.
                rows=session._execute(phase);report['phase_rows'].append(rows)
                if mode=='list-identity':
                    valid=native_valid(mode,rows) if i==3 else valid_labels(['same-native-list-challenge'],rows) if i==4 else rows==[['stage',True]]
                else:
                    valid=native_valid(mode,rows) if i==2 else valid_labels(['detached-no-number','detached-outline','detached-text'],rows) if i==5 else native_valid('list-identity',rows) if i==6 else rows==[['stage',True]]
                if not valid:raise AssertionError('Native phase acknowledgement failed: '+str(i))
                report['checks']['phase-'+str(i)]=True
                if mode!='list-identity' and i==2:session.save_docx(output/'pre-list.docx')
            session.save_docx(output/'heading.docx');session.export_pdf(output/'heading.pdf')
        close_after_owned(session,output,report,sentinel_name,token);sentinel_name=None
        digest=sha(output/'heading.docx')
        with MacWordSession.open_document(output/'heading.docx',read_only=True,visible=False) as reopened:
            reopened._retain_evidence=True
            if mode in ('list-identity','detached-properties'):
                rows=reopened._execute(list_readback(report['list_name']));report['reopen_list_rows']=rows
                report['checks']['reopen_list_identity']=native_valid('list-identity',rows)
            else:
                rows=reopened._execute(clone_readback(mode));report['reopen_clone_rows']=rows
                report['checks']['reopen_clone']=native_valid(mode,rows)
            if mode=='detached-properties':
                rows=reopened._execute(detached_readback());report['reopen_detached_rows']=rows
                report['checks']['reopen_detached']=valid_labels(['detached-no-number','detached-outline','detached-text'],rows)
        report['checks']['reopen_source_unchanged']=digest==sha(output/'heading.docx')
        report['inventory_final']=inventory(output,'after-reopen')
        report['checks']['final_inventory_preserved']=report['inventory_final']==report['inventory_before']
        report['checks'].update(artifact_checks(output,mode))
        report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
        report['status']='PASS' if all(report['checks'].values()) else 'FAIL'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:close_after_owned(session,output,report,sentinel_name,token);sentinel_name=None
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        for owner,label in ((session,'native-runtime'),(reopened,'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/label,dirs_exist_ok=True)
        if sentinel_name or report.get('cleanup_failure'):report['status']='FAIL'
        report['remaining_sentinel']=sentinel_name
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.pdf','.xml','.png')}
        (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report



def diagnostic_value(prefix):
    """Capture actual scalar values in JSON-safe, type-labelled native rows."""
    return [f'if class of {prefix}Value is text then',f'set {prefix}Kind to "text"',
            f'else if class of {prefix}Value is boolean then',f'set {prefix}Kind to "boolean"',
            f'else if (class of {prefix}Value is integer) or (class of {prefix}Value is real) then',f'set {prefix}Kind to "number"',
            f'else if class of {prefix}Value is list then',f'set {prefix}Kind to "number-list"',
            'else',f'set {prefix}Kind to "native-enum"',f'set {prefix}Value to {prefix}Value as text','end if']


def diagnostic_readback():
    lines=bindings()+['set nativeRows to {}']
    values=dimensions('font-properties')+[
        ('clone-base-normal','name local of Word style (style normal) of boundDoc as text',
         'name local of Word style (base style of detachedStyle) of boundDoc as text')]
    for label,left,right in values:
        lines += [f'set leftValue to {left}',f'set rightValue to {right}',
                  'if class of leftValue is text then',
                  "set sameValue to ((current application's NSString's stringWithString:leftValue)'s isEqualToString:rightValue) as boolean",
                  'else','set sameValue to (leftValue is equal to rightValue)','end if']
        lines += diagnostic_value('left')+diagnostic_value('right')
        lines += [f'set end of nativeRows to {{{apple_string(label)},{{leftKind,leftValue}},{{rightKind,rightValue}},sameValue}}']
    return lines


def diagnostic_phases():
    # Original acceptance modes remain byte-for-byte in behavior; this candidate
    # isolates the repeated base assignment rather than silently changing them.
    copy=bindings()+['set base style of detachedStyle to style normal',
                     'set properties of destinationFont to properties of sourceFont','set nativeRows to {{"stage",true}}']
    reset=bindings()+['set base style of detachedStyle to style normal','set nativeRows to {{"stage",true}}']
    return [seed_commands('font-properties'),copy,diagnostic_readback(),reset,diagnostic_readback()]


def diagnostic_rows_valid(rows):
    import math
    def number(value):return type(value) in (int,float) and math.isfinite(value)
    def actual(value):
        if not isinstance(value,list) or len(value)!=2:return False
        kind,observed=value
        if kind=='number':return number(observed)
        if kind=='boolean':return type(observed) is bool
        if kind in ('text','native-enum'):return isinstance(observed,str) and (kind!='native-enum' or bool(observed))
        if kind=='number-list':return isinstance(observed,list) and bool(observed) and all(number(v) for v in observed)
        return False
    labels=expected_labels('font-properties')
    if not isinstance(rows,list) or len(rows)!=len(labels):return False
    for label,row in zip(labels,rows):
        if not isinstance(row,list) or len(row)!=4 or row[0]!=label or type(row[3]) is not bool:return False
        if not actual(row[1]) or not actual(row[2]):return False
        # Require the supplied native equality bit to agree with retained actual
        # values; do not accept a claimed True for different typed observations.
        if row[3] != (row[1]==row[2]):return False
    return True


def record_observation(report,label,rows):
    if not diagnostic_rows_valid(rows):raise AssertionError('Malformed diagnostic observation: '+label)
    report['diagnostic'][label]={'rows':rows,'mismatch_labels':[r[0] for r in rows if not r[3]],
                                 'appearance_equal':all(r[3] for r in rows)}


def diagnostic_artifacts(output,report):
    """Record complete appearance verdicts as observations, never weaken them."""
    appearances={}
    for label in ('before-base-reset','after-base-reset'):
        with ZipFile(output/(label+'.docx')) as z:
            xml=z.read('word/styles.xml');doc=z.read('word/document.xml')
        (output/(label+'-styles.xml')).write_bytes(xml)
        (output/(label+'-document.xml')).write_bytes(doc)
        appearances[label]={'complete_font_xml_equal':clone_xml_valid(xml,'font-properties')}
    report['artifact_appearance_observations']=appearances
    import fitz
    with fitz.open(output/'after-base-reset.pdf') as pdf:
        text='\n'.join(p.get_text() for p in pdf)
        for i,p in enumerate(pdf):p.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(output/('page-%s.png'%(i+1)))
    (output/'pdf-text.txt').write_text(text)
    report['checks']['pdf_probe_text_visible']='SOURCE APPEARANCE' in text and 'DETACHED APPEARANCE' in text


def run_order_diagnostic(output):
    report={'status':'FAIL','mode':'font-order-diagnostic','checks':{},'diagnostic':{},
            'scope':'ordering diagnosis only; mismatches are observations, never native clone acceptance',
            'source_hashes':retain_sources(output,SOURCES)}
    session=None;reopened=None;sentinel_name=None;token=None;snapshots={}
    def flush():(output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    def stage(owner,label,commands):
        report['current_phase']=label;flush()
        rows=owner._execute(commands)
        report.setdefault('stage_rows',{})[label]=rows
        if rows!=[['stage',True]]:raise AssertionError('Diagnostic stage ACK invalid: '+label)
        report['checks'][label]=True;flush()
    def observe(owner,label,commands):
        report['current_phase']=label;flush()
        rows=owner._execute(commands)
        # Retain malformed rows too, before strict validation raises.
        report.setdefault('raw_observation_rows',{})[label]=rows;flush()
        record_observation(report,label,rows);flush()
    try:
        report['dictionary_sha256']=sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before']=inventory(output,'before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence=True
            token='HEADING ORDER SENTINEL '+uuid4().hex+' 中文😀'
            rows=session._execute(['set sentinelDoc to make new document',f'set content of text object of sentinelDoc to {apple_string(token)}','set nativeRows to {{name of sentinelDoc as text}}'])
            if len(rows)!=1 or len(rows[0])!=1 or not isinstance(rows[0][0],str):raise RuntimeError('Sentinel creation ACK invalid')
            sentinel_name=rows[0][0];report['sentinel_name']=sentinel_name
            report['sentinel_before']=sentinel_preimage(inventory(output,'sentinel-before'),sentinel_name,token)
            report['word_version']=session._execute(['set nativeRows to {{version as text}}'])
            phases=diagnostic_phases()
            stage(session,'seed',phases[0]);stage(session,'copy_after_normal_base',phases[1])
            observe(session,'immediate_after_copy',phases[2])
            session.save_docx(output/'before-base-reset.docx')
            snapshots['before-base-reset']=sha(output/'before-base-reset.docx')
            # A save itself can normalize Word style state. Retain that separate
            # observation before attributing later differences to the base setter.
            observe(session,'before_reset_after_save',phases[2])
            stage(session,'isolated_normal_base_reset',phases[3])
            observe(session,'after_base_reset',phases[4])
            session.save_docx(output/'after-base-reset.docx')
            snapshots['after-base-reset']=sha(output/'after-base-reset.docx')
            session.export_pdf(output/'after-base-reset.pdf')
        close_after_owned(session,output,report,sentinel_name,token);sentinel_name=None
        for label in ('before-base-reset','after-base-reset'):
            with MacWordSession.open_document(output/(label+'.docx'),read_only=True,visible=False) as reopened:
                reopened._retain_evidence=True
                observe(reopened,'reopen_'+label,diagnostic_readback())
            shutil.copytree(reopened.staging_root,output/('reopen-runtime-'+label),dirs_exist_ok=True)
            report['checks']['reopen_'+label+'_closed']=reopened._closed and not reopened._quarantined
            report['checks']['reopen_'+label+'_inventory']=inventory(output,'after-reopen-'+label)==report['inventory_before']
        report['snapshot_hashes']=snapshots
        report['checks']['snapshot_bytes_preserved']=all(sha(output/(k+'.docx'))==v for k,v in snapshots.items())
        report['inventory_final']=inventory(output,'final')
        report['checks']['final_inventory_preserved']=report['inventory_final']==report['inventory_before']
        diagnostic_artifacts(output,report)
        report['checks']['sources_unchanged']=all(sha(ROOT/p)==v for p,v in report['source_hashes'].items())
        report['diagnostic_observation_complete']=all(report['checks'].values())
        # Never return acceptance PASS merely because the diagnostic ran to end.
        report['status']='DIAGNOSTIC_COMPLETE' if report['diagnostic_observation_complete'] else 'FAIL'
    except BaseException as exc:
        report['error']={'type':type(exc).__name__,'message':str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:close_after_owned(session,output,report,sentinel_name,token);sentinel_name=None
            except BaseException:report['cleanup_failure']=traceback.format_exc()
        for owner,label in ((session,'native-runtime'),(reopened,'last-reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():shutil.copytree(owner.staging_root,output/label,dirs_exist_ok=True)
        if sentinel_name or report.get('cleanup_failure'):report['status']='FAIL'
        report['remaining_sentinel']=sentinel_name
        report['artifact_hashes']={p.name:sha(p) for p in output.iterdir() if p.suffix in ('.docx','.pdf','.xml','.png')}
        flush()
    return report


def main(argv=None):
    parser=argparse.ArgumentParser();parser.add_argument('--execute',action='store_true')
    parser.add_argument('--mode',choices=MODES+('font-order-diagnostic',),required=True);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args(argv)
    if not args.execute:
        print('Native feasibility requires explicit --execute',file=sys.stderr);return 2
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    if args.mode=='font-order-diagnostic':
        return 0 if run_order_diagnostic(output)['status']=='DIAGNOSTIC_COMPLETE' else 1
    return 0 if run(output,args.mode)['status']=='PASS' else 1


if __name__=='__main__':raise SystemExit(main())
