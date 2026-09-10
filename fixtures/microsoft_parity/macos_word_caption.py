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
FORMAT_FACTS = [17, 3, 5, 11, 7, 21, True, 14]


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


def apply_nondefault_format_commands(start, end):
    return [
        f'set captionPreimage to create range boundDoc start {start} end {end}',
        'set first line indent of paragraph format of captionPreimage to 17',
        'set paragraph format left indent of paragraph format of captionPreimage to 3',
        'set paragraph format right indent of paragraph format of captionPreimage to 5',
        'set space before of paragraph format of captionPreimage to 11',
        'set space after of paragraph format of captionPreimage to 7',
        'set line spacing rule of paragraph format of captionPreimage to line space exactly',
        'set line spacing of paragraph format of captionPreimage to 21',
        'set italic of font object of captionPreimage to true',
        'set font size of font object of captionPreimage to 14',
        'set nativeRows to {{"format-seed",true}}',
    ]


def format_readback_commands(label, start, end):
    return [
        f'set captionDetail to create range boundDoc start {start} end {end}',
        f'set captionMark to create range boundDoc start {end} end {end + 1}',
        f'set followingDetail to create range boundDoc start {end + 1} end {end + 2}',
        f'set nativeRows to {{{{"caption-format-detail",{apple_string(label)},start of content of captionDetail,end of content of captionDetail,content of captionDetail as text,'
        '(alignment of paragraph format of captionDetail is align paragraph center),keep together of paragraph format of captionDetail,keep with next of paragraph format of captionDetail,'
        'first line indent of paragraph format of captionDetail,paragraph format left indent of paragraph format of captionDetail,paragraph format right indent of paragraph format of captionDetail,'
        'space before of paragraph format of captionDetail,space after of paragraph format of captionDetail,line spacing of paragraph format of captionDetail,'
        'italic of font object of captionDetail,font size of font object of captionDetail},'
        f'{{"caption-mark-detail",{apple_string(label)},start of content of captionMark,end of content of captionMark,content of captionMark as text}},'
        f'{{"following-text-detail",{apple_string(label)},start of content of followingDetail,end of content of followingDetail,content of followingDetail as text}}}}',
    ]


def validate_format_readback(rows, label, *, start, end, following):
    try:
        if (not isinstance(rows, list) or len(rows) != 3
                or not isinstance(rows[0], list) or len(rows[0]) != 16
                or any(not isinstance(row, list) or len(row) != 5
                       for row in rows[1:])):
            return False
        caption, mark, after = rows
        caption_text = ('图😀(1)尾 说明😀' if label == 'global'
                        else '表😀[1-1]尾 True')
        keep = label == 'chapter'
        return (
            caption[:2] == ['caption-format-detail', label]
            and mark[:2] == ['caption-mark-detail', label]
            and after[:2] == ['following-text-detail', label]
            and all(type(value) is int for row in rows for value in row[2:4])
            and type(start) is int and type(end) is int and 0 <= start < end
            and caption[2:4] == [start, end]
            and mark[2:4] == [end, end + 1]
            and after[2:4] == [end + 1, end + 2]
            and caption[4] == caption_text
            and mark[4] == '\r' and after[4] == following
            and caption[5:8] == [True, True, keep]
            and all(type(value) is bool for value in caption[5:8])
            and all(type(value) is int for value in caption[8:14])
            and type(caption[14]) is bool and type(caption[15]) is int
            and caption[8:] == FORMAT_FACTS
        )
    except (TypeError, ValueError, IndexError):
        return False


def validate_pdf_text(text):
    """Accept pdfplumber's local emoji/CJK order ambiguity, not content drift."""
    if not isinstance(text, str):
        return False
    lines = [''.join(line.split()) for line in text.splitlines() if line.strip()]
    if len(lines) != 4:
        return False
    emoji_counts = [line.count('😀') for line in lines]
    without_emoji = [line.replace('😀', '') for line in lines]
    return (
        emoji_counts == [0, 2, 1, 0]
        and without_emoji == [
            '1Chapter', 'left图(1)尾说明', '表[1-1]尾True', 'TAIL',
        ]
    )


def validate_readback(rows):
    try:
        if (not isinstance(rows, list) or len(rows) != 9
                or [row[0] if isinstance(row, list) and row else None for row in rows]
                != ['body', 'field', 'field', 'field', 'bookmark', 'bookmark',
                    'paragraph', 'paragraph', 'counts']):
            return False
        body_row, *tail = rows
        field_rows = tail[:3]
        bookmark_rows = tail[3:5]
        global_paragraph, chapter_paragraph, counts = tail[5:]
        if (len(body_row) != 2 or not all(isinstance(value, str) for value in body_row)
                or any(len(row) != 3 or not all(isinstance(value, str) for value in row)
                       for row in field_rows)
                or any(len(row) != 5 or not isinstance(row[1], str)
                       or any(type(value) is not int for value in row[2:4])
                       or not isinstance(row[4], str) for row in bookmark_rows)
                or any(len(row) != 7 or any(type(value) is not int for value in row[1:3])
                       or not isinstance(row[3], str)
                       or any(type(value) is not bool for value in row[4:])
                       for row in (global_paragraph, chapter_paragraph))
                or len(counts) != 4
                or any(type(value) is not int for value in counts[1:])):
            return False
        body = body_row[1]
        fields = [(row[1].strip().removesuffix(' \\* MERGEFORMAT'), row[2])
                  for row in field_rows]
        bookmarks = {row[1]: (row[2], row[3], row[4]) for row in bookmark_rows}
        expected_body = 'Chapter\rleft 图😀(1)尾 说明😀\r 表😀[1-1]尾 True\r\r\rTAIL\r\r'
        return (
            body == expected_body
            and fields[0] == ('SEQ WPSC_FIG \\* ARABIC', '1')
            and re.fullmatch(r'STYLEREF "[^"]+" \\s', fields[1][0]) is not None
            and fields[1][1] == '1'
            and fields[2] == ('SEQ WPSC_TAB \\* ARABIC \\s 1', '1')
            and [row[1] for row in bookmark_rows] == BOOKMARKS
            and all(0 <= start < end and text in {'1', '1-1'}
                    for start, end, text in bookmarks.values())
            and bookmarks[BOOKMARKS[0]][2] == '1'
            and bookmarks[BOOKMARKS[1]][2] == '1-1'
            and global_paragraph[3] == 'left 图😀(1)尾 说明😀\r'
            and chapter_paragraph[3] == ' 表😀[1-1]尾 True\r'
            and 0 <= global_paragraph[1] < global_paragraph[2]
            and global_paragraph[2] == chapter_paragraph[1]
            and chapter_paragraph[1] < chapter_paragraph[2]
            and global_paragraph[1] <= bookmarks[BOOKMARKS[0]][0]
            and bookmarks[BOOKMARKS[0]][1] <= global_paragraph[2]
            and chapter_paragraph[1] <= bookmarks[BOOKMARKS[1]][0]
            and bookmarks[BOOKMARKS[1]][1] <= chapter_paragraph[2]
            and global_paragraph[4:] == [True, True, False]
            and chapter_paragraph[4:] == [True, True, True]
            and counts == ['counts', 0, 0, 0]
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
            assert session._execute(apply_nondefault_format_commands(
                LONG_START, LONG_END)) == [['format-seed', True]]
            global_range = session._add_native_caption(
                '说明😀', GLOBAL, BOOKMARKS[0], 'fig:caption',
                keep_with_next=False)
            assert (global_range.Start == LONG_START
                    and global_range.Start < global_range.End < LONG_END)
            report['global_caption_range'] = [global_range.Start, global_range.End]
            report['global_format_rows'] = session._execute(
                format_readback_commands(
                    'global', global_range.Start, global_range.End))
            assert validate_format_readback(
                report['global_format_rows'], 'global',
                start=global_range.Start, end=global_range.End, following=' ')
            report['checks']['global_handle_nondefault_format_and_trailing_paragraph'] = True

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
            assert session._execute(apply_nondefault_format_commands(
                point[0][1], point[0][2])) == [['format-seed', True]]
            chapter_range = session._add_native_caption(
                True, CHAPTER, BOOKMARKS[1], 'tab:caption', keep_with_next=True)
            assert chapter_range.Start == point[0][1] and chapter_range.Start < chapter_range.End
            report['chapter_caption_range'] = [chapter_range.Start, chapter_range.End]
            report['chapter_format_rows'] = session._execute(
                format_readback_commands(
                    'chapter', chapter_range.Start, chapter_range.End))
            assert validate_format_readback(
                report['chapter_format_rows'], 'chapter',
                start=chapter_range.Start, end=chapter_range.End,
                following='\r')
            report['checks']['chapter_handle_nondefault_format_and_trailing_paragraph'] = True

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
        report['checks']['pdf_preserves_caption_text_unicode_affixes_and_suffix'] = (
            validate_pdf_text(report['pdf_text']))
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
