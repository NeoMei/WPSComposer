"""Guarded private Word caption acceptance. Requires --execute and root lease."""
from __future__ import annotations

import argparse
from dataclasses import asdict
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

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage
from fixtures.microsoft_parity.macos_word_quality_feasibility import (
    close_sentinel_after_owned, source_pair_matches,
)

BOOKMARKS = ['wpsc_fig_' + 'c' * 24, 'wpsc_tab_' + 'd' * 24]
GLOBAL = dict(
    sequenceId='WPSC_FIG', mode='global', prefix='图😀(', suffix=')尾')
CHAPTER = dict(
    sequenceId='WPSC_TAB', mode='chapter', chapterStyleLevel=1, resetLevel=1,
    prefix='表😀[', suffix=']尾')
SOURCES = [
    Path(__file__), ROOT / 'tests/msoffice/test_macos_word_numbering.py',
    ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_numbering.py',
    ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
    ROOT / 'skills/WPSComposer/scripts/writer.py',
    *[ROOT / ('fixtures/microsoft_parity/' + name) for name in (
        'macos_word_fields.py', 'macos_word_recovery.py',
        'macos_word_paragraph_rule.py', 'macos_word_quality_feasibility.py')],
]


def _units(value):
    return len(value.encode('utf-16-le')) // 2


LONG_SELECTION = 'REPLACE-' + '长😀' * 40
SEED = f'Chapter\rleft {LONG_SELECTION} right\r\rTAIL\r'
LONG_START = _units('Chapter\rleft ')
LONG_END = LONG_START + _units(LONG_SELECTION)
SECOND_START = LONG_END + _units(' ')
SECOND_END = SECOND_START + _units('right')


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def seed_commands():
    return [
        'activate object boundWindow',
        f'set content of text object of boundDoc to {apple_string(SEED)}',
        'set captionList to make new list template at boundDoc with properties {name:"WPSCCaptionFixture",outline numbered:true}',
        'set captionLevel to list level 1 of captionList',
        'set number style of captionLevel to list number style arabic',
        'set number format of captionLevel to "%1"',
        'set start at of captionLevel to 1',
        'set linked style of captionLevel to name local of Word style (style heading1) of boundDoc as text',
        'set style of text object of paragraph 1 of boundDoc to style heading1',
        f'set secondCaptionRange to create range boundDoc start {SECOND_START} end {SECOND_END}',
        'make new bookmark at boundDoc with properties {name:"caption_second_slot",text object:secondCaptionRange}',
        f'set selection start of selection of boundWindow to {LONG_START}',
        f'set selection end of selection of boundWindow to {LONG_END}',
        'set nativeRows to {{"seed",list string of list format of text object of paragraph 1 of boundDoc as text}}',
    ]


def readback_commands(bookmarks=BOOKMARKS):
    lines = [
        'set nativeRows to {{"body",content of text object of boundDoc as text}}',
        'repeat with captionIndex from 1 to count fields of boundDoc',
        'set captionField to field captionIndex of boundDoc',
        'set end of nativeRows to {"field",content of field code of captionField as text,content of result range of captionField as text}',
        'end repeat',
    ]
    for name in bookmarks:
        lines += [
            f'set captionBookmark to bookmark {apple_string(name)} of boundDoc',
            'set end of nativeRows to {"bookmark",name of captionBookmark,start of bookmark of captionBookmark,end of bookmark of captionBookmark,content of text object of captionBookmark as text}',
        ]
    lines += [
        'repeat with captionIndex from 1 to count paragraphs of boundDoc',
        'set captionRange to text object of paragraph captionIndex of boundDoc',
        'set captionText to content of captionRange as text',
        'if captionText contains "图😀(" or captionText contains "表😀[" then',
        'set end of nativeRows to {"paragraph",start of content of captionRange,end of content of captionRange,captionText,(alignment of paragraph format of captionRange is align paragraph center),keep together of paragraph format of captionRange,keep with next of paragraph format of captionRange}',
        'end if',
        'end repeat',
        'set end of nativeRows to {"counts",count tables of boundDoc,count inline shapes of boundDoc,count shapes of boundDoc}',
    ]
    return lines


def validate_readback(rows):
    try:
        body = next(row[1] for row in rows if row[0] == 'body')
        fields = [
            (row[1].strip().removesuffix(' \\* MERGEFORMAT'), row[2])
            for row in rows if row[0] == 'field'
        ]
        numbering = [
            value for value in fields
            if value[0].startswith(('SEQ ', 'STYLEREF '))
        ]
        bookmarks = {
            row[1]: (row[2], row[3], row[4])
            for row in rows if row[0] == 'bookmark'
        }
        paragraphs = [row for row in rows if row[0] == 'paragraph']
        global_paragraph = next(row for row in paragraphs if '图😀(' in row[3])
        chapter_paragraph = next(row for row in paragraphs if '表😀[' in row[3])
        return (
            len(numbering) == 3
            and numbering[0] == ('SEQ WPSC_FIG \\* ARABIC', '1')
            and re.fullmatch(r'STYLEREF "[^"]+" \\s', numbering[1][0]) is not None
            and numbering[1][1] == '1'
            and numbering[2] == ('SEQ WPSC_TAB \\* ARABIC \\s 1', '1')
            and set(bookmarks) == set(BOOKMARKS)
            and all(start < end and text in {'1', '1-1'}
                    for start, end, text in bookmarks.values())
            and bookmarks[BOOKMARKS[0]][2] == '1'
            and bookmarks[BOOKMARKS[1]][2] == '1-1'
            and '图😀(1)尾 说明😀' in body
            and '表😀[1-1]尾 True' in body
            and 'REPLACE' not in body and 'right' not in body and 'TAIL' in body
            and global_paragraph[4:] == [True, True, False]
            and chapter_paragraph[4:] == [True, True, True]
            and ['counts', 0, 0, 0] in rows
        )
    except (TypeError, ValueError, IndexError, StopIteration):
        return False


def run(output, *, execute=False):
    if not execute:
        raise ValueError('--execute and the root Word lease required')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {
        'status': 'FAIL', 'checks': {},
        'scope': 'private bound native caption primitive only; no public figure/equation declaration or arbitrary rollback claim',
        'source_hashes': retain_sources(output, SOURCES),
    }
    session = reopened = None
    sentinel_name = sentinel_token = None

    def flush():
        (output / 'report.json').write_text(
            json.dumps(report, ensure_ascii=False, indent=2))

    def stage(label):
        report['current_step'] = label
        flush()

    try:
        report['inventory_before'] = inventory(output, 'before')
        with create_document('writer', engine='msoffice', visible=False) as session:
            session._retain_evidence = True
            stage('seed-owned-middle-selection')
            report['seed'] = session._execute(seed_commands())
            assert report['seed'] == [['seed', '1']]

            sentinel_token = 'CAPTION SENTINEL 中文😀 ' + uuid4().hex
            report['current_step'] = 'sentinel-create-submit'
            flush()
            sentinel_rows = session._execute([
                'set captionSentinel to make new document',
                f'set content of text object of captionSentinel to {apple_string(sentinel_token)}',
                'set captionSentinelWindow to active window of captionSentinel',
                'activate object captionSentinelWindow',
                'set nativeRows to {{name of captionSentinel as text}}',
            ])
            report['sentinel_create_rows'] = sentinel_rows
            flush()
            if not (isinstance(sentinel_rows, list) and len(sentinel_rows) == 1
                    and isinstance(sentinel_rows[0], list)
                    and len(sentinel_rows[0]) == 1
                    and isinstance(sentinel_rows[0][0], str)
                    and sentinel_rows[0][0]):
                session._retain('Native caption sentinel acknowledgement invalid')
                raise ValueError('Invalid sentinel acknowledgement')
            sentinel_name = sentinel_rows[0][0]
            report['sentinel_name'] = sentinel_name
            report['sentinel_before'] = sentinel_preimage(
                inventory(output, 'sentinel-before'), sentinel_name, sentinel_token)

            stage('global-caption-noncollapsed-middle')
            global_range = session._add_native_caption(
                '说明😀', GLOBAL, BOOKMARKS[0], 'fig:caption',
                keep_with_next=False)
            assert (global_range.Start == LONG_START
                    and global_range.Start < global_range.End < LONG_END)

            stage('global-caption-readback-before-second-replacement')
            first_rows = session._execute(readback_commands(BOOKMARKS[:1]))
            first_body = next(row[1] for row in first_rows if row[0] == 'body')
            assert ('left 图😀(1)尾 说明😀' in first_body
                    and 'REPLACE' not in first_body and 'right' in first_body)
            report['checks']['long_selection_replaced_from_start_with_suffix_preserved'] = True

            stage('chapter-caption-replaces-second-slot-after-first-readback')
            point = session._execute([
                'set secondCaptionRange to text object of bookmark "caption_second_slot" of boundDoc',
                'set selection start of selection of boundWindow to start of content of secondCaptionRange',
                'set selection end of selection of boundWindow to end of content of secondCaptionRange',
                'set captionSelection to text object of selection of boundWindow',
                'set nativeRows to {{"selection",start of content of captionSelection,end of content of captionSelection,content of captionSelection as text}}',
            ])
            assert (len(point) == 1 and point[0][0] == 'selection'
                    and type(point[0][1]) is int and type(point[0][2]) is int
                    and point[0][1] < point[0][2] and point[0][3] == 'right')
            chapter_range = session._add_native_caption(
                True, CHAPTER, BOOKMARKS[1], 'tab:caption', keep_with_next=True)
            assert chapter_range.Start == point[0][1] and chapter_range.Start < chapter_range.End

            stage('refresh-and-readback')
            session.repaginate_and_update_numbering()
            report['native_rows'] = session._execute(readback_commands())
            assert validate_readback(report['native_rows'])
            report['checks']['bound_middle_caption_fields_bookmarks_ranges_and_format'] = True

            report['field_snapshots'] = [asdict(value) for value in session.snapshot_fields()]
            assert [tuple(value['stable_key']) for value in report['field_snapshots']] == [
                ('fig:caption', 'SEQ_FIG', 0),
                ('tab:caption', 'STYLEREF', 0),
                ('tab:caption', 'SEQ_TAB', 0),
            ]
            assert all(value['field_category'] == 'numbering'
                       for value in report['field_snapshots'])
            report['checks']['field_owner_category_and_ordinals'] = True

            stage('save-and-pdf')
            session.save_docx(output / 'caption.docx')
            session.export_pdf(output / 'caption.pdf')
            report['sentinel_after'] = sentinel_preimage(
                inventory(output, 'sentinel-after'), sentinel_name, sentinel_token)
            assert report['sentinel_after'] == report['sentinel_before']
            report['checks']['exact_unsaved_sentinel_preserved'] = True

        close_sentinel_after_owned(
            session, output, report, sentinel_name, sentinel_token)
        sentinel_name = None

        stage('owned-reopen')
        digest = sha(output / 'caption.docx')
        with open_document(
                output / 'caption.docx', engine='msoffice', read_only=True,
                visible=False) as reopened:
            reopened._retain_evidence = True
            report['reopen_rows'] = reopened._execute(readback_commands())
            assert validate_readback(report['reopen_rows'])
        report['checks']['native_reopen_source_preserved'] = (
            sha(output / 'caption.docx') == digest)

        with ZipFile(output / 'caption.docx') as archive:
            document = archive.read('word/document.xml')
            (output / 'document.xml').write_bytes(document)
            root = ET.fromstring(document)
            word = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            codes = [element.text or '' for element in root.iter(word + 'instrText')]
            codes += [element.get(word + 'instr', '')
                      for element in root.iter(word + 'fldSimple')]
            report['checks']['editable_native_fields_without_object_side_effects'] = (
                sum('SEQ WPSC_' in code for code in codes) == 2
                and sum('STYLEREF ' in code for code in codes) == 1
                and not list(root.iter(word + 'tbl'))
                and not list(root.iter(word + 'drawing'))
            )

        import pdfplumber
        with pdfplumber.open(output / 'caption.pdf') as pdf:
            report['pdf_text'] = '\n'.join(
                page.extract_text() or '' for page in pdf.pages)
        report['checks']['pdf_preserves_caption_text_unicode_affixes_and_suffix'] = all(
            text in report['pdf_text']
            for text in ('说明', 'True', 'TAIL', '尾', '1-1'))
    except BaseException as error:
        report['error'] = {'type': type(error).__name__, 'message': str(error)}
        (output / 'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and sentinel_name:
            try:
                close_sentinel_after_owned(
                    session, output, report, sentinel_name, sentinel_token)
                sentinel_name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(
                    owner.staging_root, output / label, dirs_exist_ok=True)
        report['remaining_sentinel'] = sentinel_name
        report['checks']['sources_unchanged'] = all(
            source_pair_matches(ROOT / path, output / 'source' / path, value)
            for path, value in report['source_hashes'].items())
        if not sentinel_name and not report.get('cleanup_failure'):
            try:
                report['inventory_final'] = inventory(output, 'final')
                report['checks']['inventory_preserved'] = (
                    report['inventory_final'] == report.get('inventory_before'))
            except BaseException:
                report['inventory_failure'] = traceback.format_exc()
        report['artifact_hashes'] = {
            path.name: sha(path) for path in output.iterdir()
            if path.suffix in ('.docx', '.pdf', '.xml')
        }
        if (report['checks'] and all(report['checks'].values())
                and not any(key in report for key in (
                    'error', 'cleanup_failure', 'inventory_failure'))
                and not sentinel_name):
            report['status'] = 'PASS'
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Native caption requires --execute and the root Word lease', file=sys.stderr)
        return 2
    return 0 if run(args.output, execute=True)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
