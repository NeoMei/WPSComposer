"""V2 owned-recipient probe: style scalars and effective paragraph tabs attributed separately.

Fixed donor and matching source rPr/pPr, Normal/docDefaults/theme only. This
probe does not support arbitrary themes, attached rollback, numbering, or the
full public direct-heading contract. Word performs every edit and final save.
Style-level tab collection access remains unproved; the native paragraph path
returned 19 matching tabs on the retained fixture without changing its state.
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
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from fixtures.microsoft_parity import macos_word_heading_feasibility as heading
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession, apple_string, sha, inventory, retain_sources,
    sentinel_preimage, close_after_owned, seed, SOURCES as ORIGINAL_SOURCES,
    package_styles, canonical, W, CLONE, IMPORTED, SOURCE, FRESH,
    paragraph_style_by_name, diagnostic_readback, diagnostic_rows_valid,
    scalar_copy_verdict, native_valid, clone_readback, clone_xml_valid,
    replacement_snapshot, existing_styles_preserved, fresh_package_valid,
    style_definition,
)

def load_helper(filename,name):
    spec=importlib.util.spec_from_file_location(name,Path(__file__).with_name(filename))
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module

collapsed=load_helper('docx-collapsed-control.py','retained_collapsed_control')
gate=load_helper('flatopc-open-control.py','retained_materialization_oracle')
tab_diagnostic=load_helper('tab-list-read-diagnostic.py','retained_tab_read_diagnostic')
DONOR=Path(__file__).parent/'docx-package-01/native-materialized.docx'
DONOR_SHA='fa7d6dd3060e67e6dedc61d37515ace3bd854624382378192e7cd59d8819ed7c'
SOURCE_STYLES=gate.SOURCE_STYLES
THEME_SHA='7017cbe8b7512b56e04e6b3e3579aa862a35b6f9f31e0bf48328f6e5c074ebbd'
SOURCES=[Path(__file__),Path(__file__).with_name('test_native_donor_heading_import_v2.py'),
         Path(collapsed.__file__),Path(gate.__file__),Path(tab_diagnostic.__file__),*ORIGINAL_SOURCES]


def recipient_matches_donor(before,donor):
    old=ET.fromstring(before)
    return (paragraph_style_by_name(old.findall(W+'style'),CLONE) is None
            and gate.materialized_style_valid(before,donor)
            and dependencies_equal(before,donor))


def dependencies_equal(left,right):
    a,b=ET.fromstring(left),ET.fromstring(right)
    an=paragraph_style_by_name(a.findall(W+'style'),'Normal')
    bn=paragraph_style_by_name(b.findall(W+'style'),'Normal')
    return (an is not None and bn is not None and canonical(an)==canonical(bn)
            and a.find(W+'docDefaults') is not None
            and canonical(a.find(W+'docDefaults'))==canonical(b.find(W+'docDefaults')))


def only_clone_added(before,after,donor):
    if not existing_styles_preserved(before,after) or not dependencies_equal(before,after):return False
    old=ET.fromstring(before).findall(W+'style');new=ET.fromstring(after).findall(W+'style')
    old_ids={s.get(W+'styleId') for s in old};new_ids=[s.get(W+'styleId') for s in new]
    clone=paragraph_style_by_name(new,CLONE)
    if (clone is None or clone.find(W+'link') is not None or len(set(new_ids))!=len(new_ids)
            or any(not i for i in new_ids) or set(new_ids)-old_ids!={clone.get(W+'styleId')}):return False
    expected=paragraph_style_by_name(ET.fromstring(donor).findall(W+'style'),CLONE)
    def stable(s):
        return tuple(sorted(s.attrib.items())),tuple(canonical(c) for c in s if c.tag!=W+'rsid')
    return expected is not None and stable(clone)==stable(expected) and clone_xml_valid(after,'detached-properties')


def imported_paragraph_valid(document_xml,styles_xml):
    styles=ET.fromstring(styles_xml).findall(W+'style')
    clone=paragraph_style_by_name(styles,CLONE)
    if clone is None:return False
    paragraphs=[p for p in ET.fromstring(document_xml).iter(W+'p')
                if ''.join(t.text or '' for t in p.iter(W+'t'))==IMPORTED]
    if len(paragraphs)!=1:return False
    p=paragraphs[0];props=p.findall(W+'pPr')
    return (len(props)==1 and len(props[0])==1 and props[0][0].tag==W+'pStyle'
            and props[0][0].get(W+'val')==clone.get(W+'styleId')
            and all(len(rpr)==0 and not rpr.attrib for rpr in p.iter(W+'rPr')))


def imported_package_valid(path):
    with ZipFile(path) as z:return imported_paragraph_valid(z.read('word/document.xml'),z.read('word/styles.xml'))


def theme_hash(path):
    import hashlib
    with ZipFile(path) as z:return hashlib.sha256(z.read('word/theme/theme1.xml')).hexdigest()


def deletion_gate(rows,start,text,before_inventory,after_inventory,owned_path,before_styles,after_styles,donor_styles,document_xml):
    return (collapsed.exact_ack(rows,'insert',start,start,text)
            and collapsed.unrelated_inventory_preserved(before_inventory,after_inventory,owned_path)
            and only_clone_added(before_styles,after_styles,donor_styles)
            and imported_paragraph_valid(document_xml,after_styles))


def style_scalar_readback():
    """Preserve every style scalar; omit only the unsupported style-tab block."""
    lines=clone_readback('detached-properties')
    first=lines.index('set tabsEqual to ((count tab stops of sourceParagraph) is (count tab stops of destinationParagraph))')
    last=lines.index('set expectedBase to name local of Word style (style normal) of boundDoc as text')
    return lines[:first]+lines[last:]


def style_scalar_labels():
    return [label for label in heading.expected_labels('detached-properties') if label!='paragraph:tabs']


def style_scalars_valid(rows):
    return heading.valid_labels(style_scalar_labels(),rows)


def effective_tab_readback():
    """Effective tabs of exact source/imported paragraphs, not style collection."""
    return tab_diagnostic.tab_commands('paragraph')


def run(output):
    report = {'status': 'FAIL', 'scope': 'Owned matching-source/theme import; complete style scalars/XML plus effective paragraph tabs; unused-style tab access and public method remain unproved',
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
    def observe(owner, label):
        rows = execute(owner, label, diagnostic_readback())
        for key, value in scalar_copy_verdict(font_preimage, rows).items():
            check(label+'_'+key, value)
        check(label+'_complete_style_scalars', style_scalars_valid(execute(owner, label+'_style_scalars', style_scalar_readback())))
        check(label+'_effective_paragraph_tabs', tab_diagnostic.tabs_valid(execute(owner, label+'_effective_paragraph_tabs', effective_tab_readback()), 'paragraph'))
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
            check('seed_ack', execute(session, 'seed', seed()) == [['stage', True]])
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
            observe(session, 'after_import')
            session.save_docx(output/'after-import.docx')
            styles = package_styles(output/'after-import.docx')
            check('all_existing_styles_preserved', only_clone_added(preimage, styles, donor_styles))
            check('inserted_styles_survive_clear', existing_styles_preserved(inserted_styles, styles))
            check('imported_style_only_after_clear', imported_package_valid(output/'after-import.docx'))
            check('complete_imported_style_xml', clone_xml_valid(styles, 'detached-properties'))
            # A new plain paragraph must pick up the imported style without copied runs.
            lines = [f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
                     'set p to (end of content of text object of boundDoc) - 1',
                     'set freshRange to create range boundDoc start p end p',
                     f'set content of freshRange to {apple_string(FRESH)} & return',
                     'set freshRange to text object of paragraph 4 of boundDoc',
                     'set style of freshRange to detachedStyle', 'reset font object of freshRange',
                     'reset paragraph format of freshRange', 'set nativeRows to {{"stage",true}}']
            check('fresh_ack', execute(session, 'fresh', lines) == [['stage', True]])
            fresh_commands = [line.replace('set sourceFont to font object of sourceStyle', 'set sourceFont to font object of text object of paragraph 4 of boundDoc') for line in diagnostic_readback()]
            rows = execute(session, 'fresh_font', fresh_commands)
            check('fresh_font_equal', all(scalar_copy_verdict(font_preimage, rows).values()))
            session.save_docx(output/'before-challenge.docx')
            check('fresh_style_only_xml', fresh_package_valid(output/'before-challenge.docx'))
            clone_before = style_definition(package_styles(output/'before-challenge.docx'), CLONE)
            source_size = next(r[1][1] for r in font_preimage if r[0] == 'font:font size')
            challenge = ['set sourceStyle to Word style (style heading1) of boundDoc', f'set font size of font object of sourceStyle to {source_size+11}', 'set nativeRows to {{"stage",true}}']
            check('challenge_ack', execute(session, 'source_challenge', challenge) == [['stage', True]])
            session.save_docx(output/'source-challenged.docx')
            check('fresh_style_only_after_challenge', fresh_package_valid(output/'source-challenged.docx'))
            check('clone_independent', style_definition(package_styles(output/'source-challenged.docx'), CLONE) == clone_before)
            rows = execute(session, 'fresh_after_challenge', fresh_commands)
            check('fresh_independent', all(scalar_copy_verdict(font_preimage, rows).values()))
            challenge[1] = f'set font size of font object of sourceStyle to {source_size}'
            check('restore_ack', execute(session, 'restore_source', challenge) == [['stage', True]])
            observe(session, 'after_restore')
            session.save_docx(output/'heading.docx')
            session.export_pdf(output/'heading.pdf')
            final_styles = package_styles(output/'heading.docx')
            (output/'styles.xml').write_bytes(final_styles)
            check('source_fully_restored', only_clone_added(preimage, final_styles, donor_styles))
            check('final_theme_preserved', theme_hash(output/'heading.docx') == THEME_SHA)
            check('imported_style_only_final', imported_package_valid(output/'heading.docx'))
            check('complete_final_style_xml', clone_xml_valid(final_styles, 'detached-properties'))
            check('fresh_style_only_final', fresh_package_valid(output/'heading.docx'))
        close_after_owned(session, output, report, sentinel_name, token)
        sentinel_name = None
        digest = sha(output/'heading.docx')
        with MacWordSession.open_document(output/'heading.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            observe(reopened, 'reopen')
            rows = execute(reopened, 'reopen_fresh', fresh_commands)
            check('reopen_fresh_equal', all(scalar_copy_verdict(font_preimage, rows).values()))
        check('reopen_bytes_preserved', digest == sha(output/'heading.docx'))
        check('fresh_style_only_reopened_package', fresh_package_valid(output/'heading.docx'))
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
        check('pdf_text_present', all(marker in text for marker in ('SOURCE APPEARANCE', 'IMPORTED APPEARANCE', 'FRESH STYLE', 'SUFFIX')))
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
