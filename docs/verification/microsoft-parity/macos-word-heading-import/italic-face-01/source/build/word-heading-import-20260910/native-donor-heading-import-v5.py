"""V5 bounded donor import acceptance with same-text native style controls.

Two indexed, proven empty carriers receive identical text via style then insert,
without reset. Direct run XML must match exactly; no font fallback whitelist.
This remains owned-fixture acceptance, not selection/numbering/public parity.
"""
from __future__ import annotations
import argparse
from copy import deepcopy
import importlib.util
import json
import math
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
def load(filename,name):
    spec=importlib.util.spec_from_file_location(name,Path(__file__).with_name(filename))
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module
v2=load('native-donor-heading-import-v2.py','retained_heading_v2')
order=load('fresh-order-diagnostic.py','retained_fresh_order_v3')
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession,apple_string,sha,inventory,retain_sources,sentinel_preimage,close_after_owned,
    seed,SOURCES as OLD_SOURCES,package_styles,canonical,W,CLONE,IMPORTED,FRESH,
    paragraph_style_by_name,diagnostic_readback,diagnostic_rows_valid,scalar_copy_verdict,
    clone_xml_valid,replacement_snapshot,existing_styles_preserved,style_definition,
)
collapsed=v2.collapsed;gate=v2.gate;heading=v2.heading;font=order.font
DONOR=v2.DONOR;DONOR_SHA=v2.DONOR_SHA;SOURCE_STYLES=v2.SOURCE_STYLES;THEME_SHA=v2.THEME_SHA
theme_hash=v2.theme_hash;recipient_matches_donor=v2.recipient_matches_donor
only_clone_added=v2.only_clone_added;imported_package_valid=v2.imported_package_valid;deletion_gate=v2.deletion_gate
SOURCES=[Path(__file__),Path(__file__).with_name('test_native_donor_heading_import_v5.py'),
         Path(v2.__file__),Path(order.__file__),Path(font.__file__),Path(v2.tab_diagnostic.__file__),
         Path(collapsed.__file__),Path(gate.__file__),*OLD_SOURCES]


def seed_v3():
    return [line for line in seed() if not line.startswith(('set content of text object of boundDoc','set style of text object of paragraph 1','set nativeRows'))]+[
        f'set content of text object of boundDoc to {apple_string(chr(13)+"REPLACE"+chr(13)+"SUFFIX"+chr(13))}',
        'set nativeRows to {{"stage",true}}']


def carriers_empty(xml,indices):
    body=ET.fromstring(xml).find(W+'body')
    if body is None:return False
    paragraphs=body.findall(W+'p')
    return len(paragraphs)==4 and all(len(paragraphs[i-1])==0 and not (paragraphs[i-1].text or '') for i in indices)


def control_constructor(body,index,start,kind):
    if kind not in ('source','clone') or index not in (1,4) or type(start) is not int:raise ValueError('Invalid control carrier')
    raw=body.encode('utf-16-le')
    if raw[start*2:start*2+2].decode('utf-16-le')!='\r':raise ValueError('Not empty carrier')
    named='Word style (style heading1)' if kind=='source' else f'Word style {apple_string(CLONE)}'
    return collapsed.full_guard(body)+[f'set carrierRange to text object of paragraph {index} of boundDoc',
        f'if start of content of carrierRange is not {start} or end of content of carrierRange is not {start+1} or content of carrierRange is not return then error "WPSC_V3_EMPTY_CARRIER"',
        f'set controlStyle to {named} of boundDoc','set style of carrierRange to controlStyle',
        f'set freshInsert to create range boundDoc start {start} end {start}',
        f'set content of freshInsert to {apple_string(FRESH)}','set nativeRows to {{"stage",true}}']


def control_anchors(body):
    paras=body.split('\r')
    if len(paras)!=5 or paras[0]!=FRESH or paras[3]!=FRESH or paras[-1]!='':raise ValueError('Unexpected indexed controls')
    first_end=collapsed.units(FRESH+'\r');last_start=collapsed.units('\r'.join(paras[:3])+'\r')
    return {owner:[{'script':'paragraph','scalar':FRESH+'\r','start':start,'end':end},
                   *order.range_anchors(body,start,end,FRESH)]
            for owner,start,end in [('source',0,first_end),('clone',last_start,collapsed.units(body))]}


def property_signature(element):
    """Retain attributes, element order, text and tails in complete properties."""
    if element is None:return None
    return (element.tag,tuple(sorted(element.attrib.items())),element.text or '',element.tail or '',
            tuple(property_signature(child) for child in element))


def control_signatures(xml,source_id,clone_id):
    """Exact scalar text/rPr formatting; validate non-formatting spelling marks."""
    paragraphs=ET.fromstring(xml).find(W+'body').findall(W+'p')
    if len(paragraphs)!=4:raise ValueError('Unexpected paragraph count')
    result=[]
    for index,identifier in [(1,source_id),(4,clone_id)]:
        p=paragraphs[index-1]
        props=p.findall(W+'pPr')
        if len(props)!=1:raise ValueError('Ambiguous control paragraph properties')
        styles=props[0].findall(W+'pStyle')
        if (len(styles)!=1 or styles[0].attrib!={W+'val':identifier} or len(styles[0]) or styles[0].text or styles[0].tail):raise ValueError('Wrong control style identity or hidden content')
        normalized_props=deepcopy(props[0])
        normalized_props.find(W+'pStyle').set(W+'val','WPSC_CONTROL_STYLE_IDENTITY')
        if p.text and p.text.strip():raise ValueError('Unwrapped control text')
        scalars=[];spelling_open=False
        for child in p:
            if child.tail and child.tail.strip():raise ValueError('Unwrapped control tail text')
            if child.tag==W+'pPr':continue
            if child.tag==W+'proofErr':
                if len(child) or child.text or set(child.attrib)!={W+'type'}:raise ValueError('Invalid proofErr content or attributes')
                kind=child.get(W+'type')
                if kind=='spellStart' and not spelling_open:spelling_open=True
                elif kind=='spellEnd' and spelling_open:spelling_open=False
                else:raise ValueError('Unsupported or unbalanced proofErr')
                continue
            if child.tag!=W+'r' or (child.text and child.text.strip()):raise ValueError('Unsupported control content')
            # Only these two non-formatting Word revision markers are omitted
            # from run identity. Unknown run attributes fail closed.
            if set(child.attrib)-{W+'rsidR',W+'rsidRPr'}:raise ValueError('Unknown run attributes')
            run_properties=child.findall(W+'rPr')
            if len(run_properties)>1 or any(node.tag not in (W+'rPr',W+'t') for node in child):raise ValueError('Unsupported run content')
            signature=property_signature(run_properties[0]) if run_properties else None
            for node in child:
                if node.tail and node.tail.strip():raise ValueError('Unwrapped run tail text')
                if node.tag==W+'rPr':continue
                if len(node) or any(key!='{http://www.w3.org/XML/1998/namespace}space' for key in node.attrib):raise ValueError('Unsupported text node content')
                # Preserve exact Unicode scalar order and every direct rPr field.
                # Only run boundaries and validated spelling annotation nodes
                # are omitted; family names, hints and szCs are never rewritten.
                scalars.extend((character,signature) for character in (node.text or ''))
        if spelling_open:raise ValueError('Unclosed proofErr')
        if ''.join(character for character,_ in scalars)!=FRESH:raise ValueError('Wrong indexed control text')
        result.append((property_signature(normalized_props),tuple(scalars)))
    return tuple(result)


def control_fonts():
    return [line.replace('set sourceFont to font object of sourceStyle','set sourceFont to font object of text object of paragraph 1 of boundDoc')
            .replace('set destinationFont to font object of detachedStyle','set destinationFont to font object of text object of paragraph 4 of boundDoc')
            for line in diagnostic_readback()]


def control_tabs():
    return [line.replace('paragraph 2 of boundDoc','paragraph 4 of boundDoc') for line in v2.effective_tab_readback()]


def native_body():
    return ['set nativeRows to {{"body",content of text object of boundDoc as text,saved of boundDoc,count paragraphs of boundDoc,count fields of boundDoc,count tables of boundDoc}}']


def anchors_nonempty(rows,anchors):
    return font.rows_valid(rows,anchors) and all(all(value for value in r[7:]) for r in rows[1:] if r[2] not in ('style','paragraph'))


def anchors_equal(rows,anchors):
    if not anchors_nonempty(rows,anchors):return False
    def normalized(owner):
        origin=anchors[owner][0]['start']
        return [r[2:4]+[r[4]-origin,r[5]-origin]+r[7:] for r in rows[1:] if r[1]==owner and r[2] not in ('style','paragraph')]
    return normalized('source')==normalized('clone')


def challenge_readback():
    return ['set sourceStyle to Word style (style heading1) of boundDoc',f'set cloneStyle to Word style {apple_string(CLONE)} of boundDoc',
        'set nativeRows to {{"challenge",font size of font object of sourceStyle,space before of paragraph format of sourceStyle,font size of font object of text object of paragraph 1 of boundDoc,space before of paragraph format of text object of paragraph 1 of boundDoc,font size of font object of cloneStyle,space before of paragraph format of cloneStyle,font size of font object of text object of paragraph 4 of boundDoc,space before of paragraph format of text object of paragraph 4 of boundDoc}}']


def challenge_rows_valid(rows):
    return isinstance(rows,list) and len(rows)==1 and isinstance(rows[0],list) and len(rows[0])==9 and rows[0][0]=='challenge' and all(type(x) in (int,float) and math.isfinite(x) for x in rows[0][1:])


def challenge_changed(before,after):
    return (challenge_rows_valid(before) and challenge_rows_valid(after)
        and after[0][1:5]==[before[0][1]+11,before[0][2]+7,before[0][3]+11,before[0][4]+7]
        and after[0][5:]==before[0][5:])


def clone_observations_stable(before,after):
    return (diagnostic_rows_valid(before['fonts']) and diagnostic_rows_valid(after['fonts'])
        and [r[2] for r in before['fonts']]==[r[2] for r in after['fonts']]
        and [r for r in before['anchors'] if len(r)>1 and r[1]=='clone']==[r for r in after['anchors'] if len(r)>1 and r[1]=='clone']
        and [r for r in before['tabs'] if r[0]=='clone']==[r for r in after['tabs'] if r[0]=='clone'])


def complete_styles_equal(before,after):
    """Entire styles tree; retain the existing style-rsid revision exclusion only."""
    def tree(xml):
        root=ET.fromstring(xml)
        # Same explicitly scoped revision marker exclusion as the retained
        # existing_styles_preserved oracle. Defaults, latent styles, root attrs,
        # all other nodes/attributes/properties remain in the exact comparison.
        for style in root.findall(W+'style'):
            for marker in style.findall(W+'rsid'):style.remove(marker)
        return canonical(root)
    return tree(before)==tree(after)


def source_only_challenge_xml(before,after):
    """Expect three exact attributes on the observed reciprocal source pair."""
    expected=ET.fromstring(before);styles=expected.findall(W+'style')
    identifiers=[style.get(W+'styleId') for style in styles]
    if any(not identifier for identifier in identifiers) or len(set(identifiers))!=len(identifiers):return False
    source=paragraph_style_by_name(styles,'heading 1')
    if source is None:return False
    links=source.findall(W+'link')
    if len(links)!=1 or set(links[0].attrib)!={W+'val'}:return False
    paired=[style for style in styles if style.get(W+'styleId')==links[0].get(W+'val')]
    if len(paired)!=1:return False
    pair=paired[0];names=pair.findall(W+'name');back=pair.findall(W+'link')
    if (pair is source or pair.get(W+'type')!='character' or pair.get(W+'customStyle')!='1'
            or len(names)!=1 or names[0].attrib!={W+'val':'标题 1 字符'}
            or len(back)!=1 or back[0].attrib!={W+'val':source.get(W+'styleId')}):return False
    sizes=source.findall('./'+W+'rPr/'+W+'sz');paired_sizes=pair.findall('./'+W+'rPr/'+W+'sz')
    spacings=source.findall('./'+W+'pPr/'+W+'spacing')
    if (len(sizes)!=1 or len(paired_sizes)!=1 or len(spacings)!=1
            or sizes[0].attrib!=paired_sizes[0].attrib or set(sizes[0].attrib)!={W+'val'}):return False
    size=sizes[0];paired_size=paired_sizes[0];spacing=spacings[0]
    try:
        original_size=int(size.get(W+'val'));original_spacing=int(spacing.get(W+'before'))
    except (ValueError,TypeError):return False
    if original_size<=0:return False
    size.set(W+'val',str(original_size+22));paired_size.set(W+'val',str(original_size+22))
    spacing.set(W+'before',str(original_spacing+140))
    # Entire root remains strict, including defaults, links, szCs and clone.
    # The existing narrow style-rsid revision-marker policy is unchanged.
    return complete_styles_equal(ET.tostring(expected),after)


def fixture_tabs_valid(rows):
    return v2.tab_diagnostic.tabs_valid(rows,'paragraph') and rows[0][4:]==[19,19]


def run(output):
    report = {'status': 'FAIL', 'scope': 'Owned same-text control acceptance; native fallback must match exactly; unused-style tabs and public heading parity remain unproved',
              'checks': {}, 'source_hashes': retain_sources(output, SOURCES)}
    session = reopened = None
    sentinel_name = token = None
    def flush():
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    def check(name, value):
        report['checks'][name] = bool(value)
        flush()
        if not value:
            raise AssertionError('Native import check failed: '+name)
    def execute(owner, label, lines):
        report['current_phase'] = label
        flush()
        rows = owner._execute(lines)
        report.setdefault('raw_rows', {})[label] = rows
        flush()
        return rows
    def observe_styles(owner,label):
        rows=execute(owner,label+'_style_fonts',diagnostic_readback())
        for key,value in scalar_copy_verdict(font_preimage,rows).items():check(label+'_'+key,value)
        check(label+'_all_style_scalars',v2.style_scalars_valid(execute(owner,label+'_style_scalars',v2.style_scalar_readback())))
    def observe_controls(owner,label,body,equal=True):
        anchors=control_anchors(body)
        values=execute(owner,label+'_font65',control_fonts())
        check(label+'_font65_typed',diagnostic_rows_valid(values))
        if equal:check(label+'_font65_exact_equal',all(r[3] for r in values))
        ranges=execute(owner,label+'_font_anchors',font.font_commands(body,anchors))
        check(label+'_anchors_typed',font.rows_valid(ranges,anchors))
        check(label+'_anchors_nonempty',anchors_nonempty(ranges,anchors))
        if equal:check(label+'_anchors_exact_equal',anchors_equal(ranges,anchors))
        tabs=execute(owner,label+'_effective_tabs',control_tabs())
        check(label+'_effective_tabs_equal',fixture_tabs_valid(tabs))
        challenge=execute(owner,label+'_effective_size_spacing',challenge_readback())
        check(label+'_size_spacing_typed',challenge_rows_valid(challenge))
        return {'fonts':values,'anchors':ranges,'tabs':tabs,'challenge':challenge}
    def package_controls(path):
        with ZipFile(path) as package:
            styles=package.read('word/styles.xml');xml=package.read('word/document.xml')
        root=ET.fromstring(styles);source=paragraph_style_by_name(root.findall(W+'style'),'heading 1');clone=paragraph_style_by_name(root.findall(W+'style'),CLONE)
        if source is None or clone is None:raise AssertionError('Control style identity absent')
        return control_signatures(xml,source.get(W+'styleId'),clone.get(W+'styleId'))
    try:
        check('pinned_word_saved_donor', sha(DONOR) == DONOR_SHA)
        check('pinned_original_source', sha(SOURCE_STYLES) == gate.PINNED[SOURCE_STYLES])
        donor_styles = package_styles(DONOR)
        check('complete_pinned_donor', gate.package_valid(DONOR, SOURCE_STYLES.read_bytes()))
        check('pinned_donor_theme', theme_hash(DONOR) == THEME_SHA)
        report['donor_sha256'] = DONOR_SHA
        report['theme_sha256'] = THEME_SHA
        report['dictionary_sha256'] = sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before'] = inventory(output, 'before')
        check('empty_inventory_precondition', report['inventory_before'] == [])
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence = True
            token = 'IMPORT SENTINEL '+uuid4().hex+' 中文😀'
            rows = execute(session, 'sentinel', ['set sentinelDoc to make new document', f'set content of text object of sentinelDoc to {apple_string(token)}', 'set nativeRows to {{name of sentinelDoc as text}}'])
            if len(rows) != 1 or len(rows[0]) != 1 or not isinstance(rows[0][0], str):
                raise ValueError('Invalid sentinel acknowledgement')
            sentinel_name = rows[0][0]
            report['sentinel_name'] = sentinel_name
            report['sentinel_before'] = sentinel_preimage(inventory(output, 'sentinel-before'), sentinel_name, token)
            report['word_version'] = execute(session, 'version', ['set nativeRows to {{version as text}}'])
            check('seed_ack', execute(session, 'seed', seed_v3()) == [['stage', True]])
            session.save_docx(output/'empty-carriers.docx')
            with ZipFile(output/'empty-carriers.docx') as package:
                check('two_empty_carriers_xml',carriers_empty(package.read('word/document.xml'),(1,4)))
            empty_styles=package_styles(output/'empty-carriers.docx')
            body_rows=execute(session,'empty_body',native_body())
            empty_body=body_rows[0][1]
            check('exact_empty_body',empty_body=='\rREPLACE\rSUFFIX\r\r')
            source_ack=execute(session,'construct_source_control',control_constructor(empty_body,1,0,'source'))
            session.save_docx(output/'source-control.docx')
            check('source_constructor_ack',source_ack==[['stage',True]])
            check('source_constructor_styles_preserved',existing_styles_preserved(empty_styles,package_styles(output/'source-control.docx')) and existing_styles_preserved(package_styles(output/'source-control.docx'),empty_styles))
            session.save_docx(output/'source-preimage.docx')
            preimage_hash = sha(output/'source-preimage.docx')
            preimage = package_styles(output/'source-preimage.docx')
            (output/'source-styles.xml').write_bytes(preimage)
            check('fresh_source_complete_match', recipient_matches_donor(preimage, donor_styles))
            check('fresh_source_theme_match', theme_hash(output/'source-preimage.docx') == THEME_SHA)
            native_donor = session.staging_root/('heading-input-'+uuid4().hex+'.docx')
            shutil.copyfile(DONOR, native_donor)
            check('private_donor_identical', sha(native_donor) == DONOR_SHA)
            # Native font preimage uses Heading1 for both columns; no clone yet.
            before_commands = [line.replace(f'Word style {apple_string(CLONE)}', 'Word style (style heading1)') for line in diagnostic_readback()]
            font_preimage = execute(session, 'source_font_preimage', before_commands)
            check('source_font_preimage_valid', diagnostic_rows_valid(font_preimage) and all(r[3] for r in font_preimage))
            before = execute(session, 'replacement_preimage', replacement_snapshot())
            check('replacement_preimage_shape', len(before) == 1 and len(before[0]) == 4 and before[0][0] == 'preimage')
            _, native_text, start, end = before[0]
            inserted, expected, shifted_start, shifted_end = collapsed.expected_stages(native_text, start, end, IMPORTED+'\r')
            report['expected'] = {'original':native_text,'inserted':inserted,'final':expected,
                                  'start':start,'end':end,'shifted_start':shifted_start,'shifted_end':shifted_end}
            before_inventory = inventory(output, 'before-insert')
            report['inventory_before_insert'] = before_inventory
            rows = execute(session, 'native_collapsed_import', collapsed.collapsed_commands(native_donor, start, native_text))
            check('exact_collapsed_insert', collapsed.exact_ack(rows, 'insert', start, start, inserted))
            after_inventory = inventory(output, 'after-insert')
            report['inventory_after_insert'] = after_inventory
            check('insert_inventory_preserved', collapsed.unrelated_inventory_preserved(before_inventory, after_inventory, session._bound_path))
            session.save_docx(output/'before-delete.docx')
            inserted_styles = package_styles(output/'before-delete.docx')
            check('only_complete_clone_added', only_clone_added(preimage, inserted_styles, donor_styles))
            check('imported_style_only_before_delete', imported_package_valid(output/'before-delete.docx'))
            check('theme_preserved_before_delete', theme_hash(output/'before-delete.docx') == THEME_SHA)
            with ZipFile(output/'before-delete.docx') as package:
                inserted_document = package.read('word/document.xml')
            check('all_deletion_prerequisites', deletion_gate(rows,start,inserted,before_inventory,after_inventory,
                  session._bound_path,preimage,inserted_styles,donor_styles,inserted_document))
            # All ACKs above must pass. This setter can only clear the proven shifted old target.
            rows = execute(session, 'clear_shifted_original', collapsed.clear_commands(inserted, shifted_start, shifted_end))
            check('exact_guarded_replacement', collapsed.exact_ack(rows, 'clear', shifted_start, shifted_end, expected))
            after_clear_inventory = inventory(output, 'after-clear')
            report['inventory_after_clear'] = after_clear_inventory
            check('clear_inventory_preserved', collapsed.unrelated_inventory_preserved(before_inventory, after_clear_inventory, session._bound_path))
            session.save_docx(output/'after-import.docx')
            styles = package_styles(output/'after-import.docx')
            check('all_existing_styles_preserved', only_clone_added(preimage, styles, donor_styles))
            check('inserted_styles_survive_clear', existing_styles_preserved(inserted_styles, styles))
            check('imported_style_only_after_clear', imported_package_valid(output/'after-import.docx'))
            check('complete_imported_style_xml', clone_xml_valid(styles, 'detached-properties'))
            with ZipFile(output/'after-import.docx') as package:
                check('clone_carrier_still_empty',carriers_empty(package.read('word/document.xml'),(4,)))
            fresh_start=collapsed.units(expected)-1
            fresh_ack=execute(session,'construct_clone_control',control_constructor(expected,4,fresh_start,'clone'))
            session.save_docx(output/'before-challenge.docx')
            check('clone_constructor_ack',fresh_ack==[['stage',True]])
            control_body=expected[:-1]+FRESH+'\r'
            check('exact_control_body',execute(session,'control_body',native_body())[0][1]==control_body)
            baseline_styles=package_styles(output/'before-challenge.docx')
            check('all_source_styles_and_complete_clone',only_clone_added(preimage,baseline_styles,donor_styles))
            check('baseline_theme',theme_hash(output/'before-challenge.docx')==THEME_SHA)
            signatures_before=package_controls(output/'before-challenge.docx')
            check('same_text_run_rPr_exact_equal',signatures_before[0]==signatures_before[1])
            observe_styles(session,'before_challenge')
            baseline=observe_controls(session,'before_challenge',control_body)
            size=baseline['challenge'][0][1];spacing=baseline['challenge'][0][2]
            changes=['set sourceStyle to Word style (style heading1) of boundDoc',
                     f'set font size of font object of sourceStyle to {size+11}',
                     f'set space before of paragraph format of sourceStyle to {spacing+7}',
                     'set nativeRows to {{"stage",true}}']
            changed_ack=execute(session,'source_challenge',changes)
            session.save_docx(output/'source-challenged.docx')
            check('challenge_ack',changed_ack==[['stage',True]])
            changed=observe_controls(session,'challenged',control_body,equal=False)
            check('source_control_size_and_spacing_changed',challenge_changed(baseline['challenge'],changed['challenge']))
            check('clone_control_native_stable',clone_observations_stable(baseline,changed))
            challenged_styles=package_styles(output/'source-challenged.docx')
            check('only_exact_source_linked_pair_challenge',source_only_challenge_xml(baseline_styles,challenged_styles))
            check('control_direct_format_unchanged_under_challenge',package_controls(output/'source-challenged.docx')==signatures_before)
            changes[1]=f'set font size of font object of sourceStyle to {size}'
            changes[2]=f'set space before of paragraph format of sourceStyle to {spacing}'
            restore_ack=execute(session,'restore_source',changes)
            session.save_docx(output/'heading.docx')
            check('restore_ack',restore_ack==[['stage',True]])
            final_styles=package_styles(output/'heading.docx')
            (output/'styles.xml').write_bytes(final_styles)
            check('source_fully_restored',only_clone_added(preimage,final_styles,donor_styles))
            check('complete_baseline_style_tree_restored',complete_styles_equal(baseline_styles,final_styles))
            check('complete_final_style_xml',clone_xml_valid(final_styles,'detached-properties'))
            check('final_controls_xml_unchanged',package_controls(output/'heading.docx')==signatures_before)
            check('final_theme_preserved',theme_hash(output/'heading.docx')==THEME_SHA)
            observe_styles(session,'restored')
            restored=observe_controls(session,'restored',control_body)
            check('all_control_native_observations_restored',restored==baseline)
            session.export_pdf(output/'heading.pdf')
        close_after_owned(session, output, report, sentinel_name, token)
        sentinel_name = None
        digest = sha(output/'heading.docx')
        with MacWordSession.open_document(output/'heading.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            observe_styles(reopened,'reopen')
            reopened_controls=observe_controls(reopened,'reopen',control_body)
            check('reopen_controls_exact',reopened_controls==baseline)
            check('reopen_private_controls_xml',package_controls(reopened._private_path)==signatures_before)
            check('reopen_complete_style_tree',complete_styles_equal(baseline_styles,package_styles(reopened._private_path)))
        check('reopen_bytes_preserved', digest == sha(output/'heading.docx'))
        check('reopen_controls_xml',package_controls(output/'heading.docx')==signatures_before)
        check('reopen_closed', reopened._closed and not reopened._quarantined)
        check('owned_closed', session._closed and not session._quarantined)
        check('preimage_bytes_preserved', preimage_hash == sha(output/'source-preimage.docx'))
        check('input_bytes_preserved', sha(DONOR) == DONOR_SHA and sha(native_donor) == DONOR_SHA)
        check('original_source_preserved', sha(SOURCE_STYLES) == gate.PINNED[SOURCE_STYLES])
        import fitz
        with fitz.open(output/'heading.pdf') as pdf:
            text = '\n'.join(page.get_text() for page in pdf)
            for i, page in enumerate(pdf):
                page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output/f'page-{i+1}.png')
        (output/'pdf-text.txt').write_text(text)
        check('pdf_text_present', all(marker in text for marker in ('IMPORTED APPEARANCE', 'FRESH STYLE', 'SUFFIX')))
        check('pdf_both_control_texts_present',text.count('FRESH STYLE')==2)
        report['inventory_final'] = inventory(output, 'final')
        check('inventory_preserved', report['inventory_final'] == report['inventory_before'])
        check('sources_unchanged', all(sha(ROOT/path) == value for path, value in report['source_hashes'].items()))
        report['status'] = 'PASS'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:
                close_after_owned(session, output, report, sentinel_name, token)
                sentinel_name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, output/label, dirs_exist_ok=True)
        if sentinel_name or report.get('cleanup_failure'):
            report['status'] = 'FAIL'
        report['remaining_sentinel'] = sentinel_name
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix in ('.docx', '.pdf', '.xml', '.png')}
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Native import probe requires explicit --execute', file=sys.stderr)
        return 2
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    return 0 if run(output)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
