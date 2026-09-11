"""Guarded build-only Word inline-picture rollback probe. Root lease required."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import shutil
import struct
import sys
import traceback
from uuid import uuid4
import zlib
import zipfile
from xml.etree import ElementTree as ET

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

TARGET_BOOKMARK = 'figure_probe_target'
SUFFIX_BOOKMARK = 'figure_probe_suffix'
SELECTED = 'REPLACE-长😀'
SEED = f'HEAD\rleft {SELECTED} right\rTAIL\r'
FORMAT = [17, 3, 5, 11, 7, 21, True, 14, True, False, False]
SOURCES = [
    Path(__file__),
    ROOT / 'build/word-figure-rollback-20260911/test_macos_word_figure_rollback_probe_v3.py',
    ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
    ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py',
    ROOT / 'fixtures/microsoft_parity/macos_word_fields.py',
    ROOT / 'fixtures/microsoft_parity/macos_word_recovery.py',
    ROOT / 'fixtures/microsoft_parity/macos_word_paragraph_rule.py',
    ROOT / 'fixtures/microsoft_parity/macos_word_quality_feasibility.py',
]


def _units(value):
    return len(value.encode('utf-16-le')) // 2


START = _units('HEAD\rleft ')
END = START + _units(SELECTED)
SUFFIX_END = END + _units(' right')


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_probe_png(path):
    """Write a deterministic 64x40 RGB PNG without optional dependencies."""
    width, height = 64, 40
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            row += bytes((30 + x * 2, 80 + y * 3, 170))
        rows.append(bytes(row))
    def chunk(kind, payload):
        return (struct.pack('>I', len(payload)) + kind + payload
                + struct.pack('>I', zlib.crc32(kind + payload) & 0xffffffff))
    data = (b'\x89PNG\r\n\x1a\n'
            + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(b''.join(rows), 9))
            + chunk(b'IEND', b''))
    Path(path).write_bytes(data)


def seed_commands():
    return [
        'activate object boundWindow',
        f'set content of text object of boundDoc to {apple_string(SEED)}',
        f'set probeTarget to create range boundDoc start {START} end {END}',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(TARGET_BOOKMARK)},text object:probeTarget}}',
        f'set probeSuffix to create range boundDoc start {END} end {SUFFIX_END}',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(SUFFIX_BOOKMARK)},text object:probeSuffix}}',
        'set first line indent of paragraph format of probeTarget to 17',
        'set paragraph format left indent of paragraph format of probeTarget to 3',
        'set paragraph format right indent of paragraph format of probeTarget to 5',
        'set space before of paragraph format of probeTarget to 11',
        'set space after of paragraph format of probeTarget to 7',
        'set line spacing rule of paragraph format of probeTarget to line space exactly',
        'set line spacing of paragraph format of probeTarget to 21',
        'set italic of font object of probeTarget to true',
        'set font size of font object of probeTarget to 14',
        'set alignment of paragraph format of probeTarget to align paragraph right',
        'set keep together of paragraph format of probeTarget to false',
        'set keep with next of paragraph format of probeTarget to false',
        f'set selection start of selection of boundWindow to {START}',
        f'set selection end of selection of boundWindow to {END}',
        'set nativeRows to {{"seed",true}}',
    ]


def state_commands(*, reset=True):
    lines = [
        'set probeSelection to text object of selection of boundWindow',
        f'set probeTargetBookmark to bookmark {apple_string(TARGET_BOOKMARK)} of boundDoc',
        f'set probeSuffixBookmark to bookmark {apple_string(SUFFIX_BOOKMARK)} of boundDoc',
        'set probeFormatRange to text object of probeTargetBookmark',
    ]
    row_lines = [
        'set end of nativeRows to {"body",content of text object of boundDoc as text}',
        'set end of nativeRows to {"selection",start of content of probeSelection,end of content of probeSelection,content of probeSelection as text}',
        'set end of nativeRows to {"bookmark",name of probeTargetBookmark,start of bookmark of probeTargetBookmark,end of bookmark of probeTargetBookmark,content of text object of probeTargetBookmark as text}',
        'set end of nativeRows to {"bookmark",name of probeSuffixBookmark,start of bookmark of probeSuffixBookmark,end of bookmark of probeSuffixBookmark,content of text object of probeSuffixBookmark as text}',
        'set end of nativeRows to {"format",first line indent of paragraph format of probeFormatRange,paragraph format left indent of paragraph format of probeFormatRange,paragraph format right indent of paragraph format of probeFormatRange,space before of paragraph format of probeFormatRange,space after of paragraph format of probeFormatRange,line spacing of paragraph format of probeFormatRange,italic of font object of probeFormatRange,font size of font object of probeFormatRange,(alignment of paragraph format of probeFormatRange is align paragraph right),keep together of paragraph format of probeFormatRange,keep with next of paragraph format of probeFormatRange}',
        'set end of nativeRows to {"counts",count inline pictures of boundDoc,count shapes of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}',
    ]
    if reset:
        lines.append('set nativeRows to {}')
    return lines + row_lines


def _preflight_commands():
    return state_commands() + [
        f'if not ((current application\'s NSArray\'s arrayWithArray:nativeRows)\'s isEqualToArray:{_state_literal()}) then error "WPSC_FIGURE_PROBE_PREIMAGE_CHANGED"',
        'set nativeRows to {}',
        f'set probeStart to {START}',
        f'set probeOriginalText to {apple_string(SELECTED)}',
        f'set probeTarget to text object of bookmark {apple_string(TARGET_BOOKMARK)} of boundDoc',
        'set probeOriginalNoProofing to no proofing of probeTarget',
        'if probeOriginalNoProofing is not false then error "WPSC_FIGURE_PROBE_PROOFING_PREIMAGE"',
        'set probeMutation to create range boundDoc start probeStart end (end of content of probeTarget)',
        'set content of probeMutation to ""',
        'set probeInsertion to create range boundDoc start probeStart end probeStart',
        'set probeBeforePictures to count inline pictures of boundDoc',
    ]


def _insert_commands(path, alt):
    return [
        f'set probePicture to make new inline picture at probeInsertion with properties {{file name:{apple_string(str(path))},link to file:false,save with document:true}}',
        'set probePictureIndex to count inline pictures of boundDoc',
        'set probePicture to inline picture probePictureIndex of boundDoc',
        'set lock aspect ratio of probePicture to true',
        'set width of probePicture to 80',
        f'set alternative text of probePicture to {apple_string(alt)}',
        'set probePictureRange to text object of probePicture',
        'set alignment of paragraph format of probePictureRange to align paragraph center',
        'set keep together of paragraph format of probePictureRange to true',
        'set keep with next of paragraph format of probePictureRange to true',
        'set probePictureStart to start of content of probePictureRange',
        'set probePictureEnd to end of content of probePictureRange',
        'set probeBreak to create range boundDoc start probePictureEnd end probePictureEnd',
        'set content of probeBreak to return',
        'set probeMutationEnd to probePictureEnd + 1',
        'set selection start of selection of boundWindow to probeMutationEnd',
        'set selection end of selection of boundWindow to probeMutationEnd',
    ]


def _rollback_commands(label, alt, failure_observed):
    # Word expands the suffix bookmark over insertion at its left edge. Its
    # start therefore cannot identify the end of the newly inserted objects.
    # This controlled probe admits only the observed one-picture + CR state.
    mutated_body = SEED.replace(SELECTED, '/\r', 1) + '\r'
    suffix_end = START + 2 + _units(' right')
    return [
        f'set probeSuffixBookmark to bookmark {apple_string(SUFFIX_BOOKMARK)} of boundDoc',
        'set probeRollbackEnd to probeMutationEnd',
        f'if probePictureStart is not {START} or probePictureEnd is not {START + 1} or probeRollbackEnd is not {START + 2} then error "WPSC_FIGURE_PROBE_BOUND_CHANGED"',
        f'if content of text object of boundDoc is not {apple_string(mutated_body)} then error "WPSC_FIGURE_PROBE_BODY_CHANGED"',
        f'if start of bookmark of probeSuffixBookmark is not {START} or end of bookmark of probeSuffixBookmark is not {suffix_end} then error "WPSC_FIGURE_PROBE_SUFFIX_BOUNDS"',
        f'if content of text object of probeSuffixBookmark is not {apple_string("/" + chr(13) + " right")} then error "WPSC_FIGURE_PROBE_SUFFIX_CHANGED"',
        'if (count inline pictures of boundDoc) is not 1 then error "WPSC_FIGURE_PROBE_PICTURE_COUNT"',
        f'if alternative text of inline picture 1 of boundDoc is not {apple_string(alt)} then error "WPSC_FIGURE_PROBE_PICTURE_IDENTITY"',
        ('set nativeRows to {{"inserted",'
         + apple_string(label)
         + ',probePictureStart,probePictureEnd,probePictureIndex,probeBeforePictures,'
           'alternative text of probePicture as text,width of probePicture,height of probePicture,'
           '(alignment of paragraph format of probePictureRange is align paragraph center),'
           'keep together of paragraph format of probePictureRange,'
           'keep with next of paragraph format of probePictureRange,'
           'probeMutationEnd,probeRollbackEnd,'
         + str(failure_observed).lower() + '}}'),
        'set probeRollback to create range boundDoc start probeStart end probeRollbackEnd',
        'set content of probeRollback to probeOriginalText',
        f'if exists bookmark {apple_string(TARGET_BOOKMARK)} of boundDoc then delete bookmark {apple_string(TARGET_BOOKMARK)} of boundDoc',
        f'set probeRestored to create range boundDoc start {START} end {END}',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(TARGET_BOOKMARK)},text object:probeRestored}}',
        f'if exists bookmark {apple_string(SUFFIX_BOOKMARK)} of boundDoc then delete bookmark {apple_string(SUFFIX_BOOKMARK)} of boundDoc',
        f'set probeRestoredSuffix to create range boundDoc start {END} end {SUFFIX_END}',
        f'make new bookmark at boundDoc with properties {{name:{apple_string(SUFFIX_BOOKMARK)},text object:probeRestoredSuffix}}',
        'set first line indent of paragraph format of probeRestored to 17',
        'set paragraph format left indent of paragraph format of probeRestored to 3',
        'set paragraph format right indent of paragraph format of probeRestored to 5',
        'set space before of paragraph format of probeRestored to 11',
        'set space after of paragraph format of probeRestored to 7',
        'set line spacing rule of paragraph format of probeRestored to line space exactly',
        'set line spacing of paragraph format of probeRestored to 21',
        'set italic of font object of probeRestored to true',
        'set font size of font object of probeRestored to 14',
        'set no proofing of probeRestored to probeOriginalNoProofing',
        'if no proofing of probeRestored is not probeOriginalNoProofing then error "WPSC_FIGURE_PROBE_PROOFING_RESTORE"',
        'set alignment of paragraph format of probeRestored to align paragraph right',
        'set keep together of paragraph format of probeRestored to false',
        'set keep with next of paragraph format of probeRestored to false',
        f'set selection start of selection of boundWindow to {START}',
        f'set selection end of selection of boundWindow to {END}',
    ] + state_commands(reset=False)


def scenario_commands(label, image, *, missing=None):
    alt = f'figure:probe/{label}'
    lines = _preflight_commands() + _insert_commands(image, alt)
    if missing is not None:
        lines += [
            'set probeSecondFailed to false',
            'try',
            'set probeSecondInsertion to create range boundDoc start probeMutationEnd end probeMutationEnd',
            f'make new inline picture at probeSecondInsertion with properties {{file name:{apple_string(str(missing))},link to file:false,save with document:true}}',
            'on error',
            'set probeSecondFailed to true',
            'end try',
            'if probeSecondFailed is false then error "WPSC_FIGURE_PROBE_EXPECTED_SECOND_FAILURE"',
            'if (count inline pictures of boundDoc) is not probeBeforePictures + 1 then error "WPSC_FIGURE_PROBE_SECOND_FAILURE_SIDE_EFFECT"',
        ]
    return lines + _rollback_commands(label, alt, missing is not None)


def _state_literal():
    def q(value):
        if isinstance(value, str):
            return apple_string(value)
        if type(value) is bool:
            return str(value).lower()
        if type(value) in (int, float):
            return str(value)
        return '{' + ','.join(q(item) for item in value) + '}'
    return q(expected_state())


def expected_state():
    return [
        ['body', SEED + '\r'],
        ['selection', START, END, SELECTED],
        ['bookmark', TARGET_BOOKMARK, START, END, SELECTED],
        ['bookmark', SUFFIX_BOOKMARK, END, SUFFIX_END, ' right'],
        ['format', *FORMAT],
        ['counts', 0, 0, 0, 0, 2],
    ]


def validate_state(rows):
    try:
        return (rows == expected_state()
                and all(type(value) is int for value in rows[1][1:3])
                and all(type(value) is int for row in rows[2:4]
                        for value in row[2:4])
                and all(type(value) is int for value in rows[4][1:7])
                and type(rows[4][7]) is bool and type(rows[4][8]) is int
                and all(type(value) is bool for value in rows[4][9:12])
                and all(type(value) is int for value in rows[5][1:]))
    except (TypeError, IndexError):
        return False


def validate_scenario(rows, label, *, failed):
    try:
        if not isinstance(rows, list) or len(rows) != 7:
            return False
        inserted = rows[0]
        if (not isinstance(inserted, list) or len(inserted) != 15
                or inserted[:2] != ['inserted', label]
                or any(type(inserted[i]) is not int for i in (2, 3, 4, 5, 12, 13))
                or inserted[2:4] != [START, START + 1]
                or inserted[4:6] != [1, 0]
                or inserted[6] != f'figure:probe/{label}'
                or any(type(inserted[i]) not in (int, float) or type(inserted[i]) is bool
                       for i in (7, 8))
                or any(not math.isfinite(inserted[i]) for i in (7, 8))
                or abs(float(inserted[7]) - 80.0) > 1.0
                or abs(float(inserted[8]) - 50.0) > 1.0
                or inserted[9:12] != [True, True, True]
                or any(type(inserted[i]) is not bool for i in (9, 10, 11, 14))
                or inserted[12] != inserted[3] + 1
                or inserted[13] != inserted[12]
                or inserted[14] is not failed):
            return False
        return validate_state(rows[1:])
    except (TypeError, ValueError, IndexError):
        return False


def document_signature(path):
    """Full fixture body properties, bookmarks, styles and theme; no cosmetic loss.

    Word changes edit-session rsids and textId when content is restored. These
    are the only omitted document attributes. Paragraph identity is retained.
    This is deliberately strict and specific to the controlled probe, not a
    general DOCX equivalence implementation.
    """
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    w14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
    volatile = {w + key for key in ('rsidR', 'rsidRPr', 'rsidRDefault', 'rsidP', 'rsidDel')}
    volatile.add(w14 + 'textId')
    def node(element):
        return (element.tag,
                tuple(sorted((key, value) for key, value in element.attrib.items()
                             if key not in volatile)),
                element.text or '', tuple(node(child) for child in element))
    with zipfile.ZipFile(path) as archive:
        return (node(ET.fromstring(archive.read('word/document.xml'))),
                archive.read('word/styles.xml'), archive.read('word/theme/theme1.xml'))


def run(output, *, execute=False):
    if not execute:
        raise ValueError('--execute and the root Word lease required')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {
        'status': 'FAIL', 'checks': {},
        'scope': 'build-only inline-picture rollback probe; no production method or capability enablement',
        'source_hashes': retain_sources(output, SOURCES),
    }
    source_image = output / 'probe.png'
    write_probe_png(source_image)
    image_digest = sha(source_image)
    creator = session = reopened = None
    sentinel_name = sentinel_token = None
    source_document = output / 'figure-rollback-source.docx'
    output_document = output / 'figure-rollback.docx'

    def flush():
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))

    def stage(label):
        report['current_step'] = label
        flush()

    try:
        report['inventory_before'] = inventory(output, 'before')
        stage('create-owned-source-preimage')
        with create_document('writer', engine='msoffice', visible=False) as creator:
            creator._retain_evidence = True
            assert creator._execute(seed_commands()) == [['seed', True]]
            report['source_preimage_rows'] = creator._execute(state_commands())
            assert validate_state(report['source_preimage_rows'])
            creator.save_docx(source_document)
        source_digest = sha(source_document)

        with open_document(
                source_document, engine='msoffice', read_only=False,
                visible=False) as session:
            session._retain_evidence = True
            stage('bind-owned-source-middle-selection')
            report['position_rows'] = session._execute([
                f'set selection start of selection of boundWindow to {START}',
                f'set selection end of selection of boundWindow to {END}',
                'set nativeRows to {{"positioned",true}}',
            ])
            assert report['position_rows'] == [['positioned', True]]
            report['preimage_rows'] = session._execute(state_commands())
            assert validate_state(report['preimage_rows'])

            sentinel_token = 'FIGURE ROLLBACK SENTINEL 中文😀 ' + uuid4().hex
            stage('sentinel-create-submit')
            sentinel_rows = session._execute([
                'set figureSentinel to make new document',
                f'set content of text object of figureSentinel to {apple_string(sentinel_token)}',
                'set figureSentinelWindow to active window of figureSentinel',
                'activate object figureSentinelWindow',
                'set nativeRows to {{name of figureSentinel as text}}',
            ])
            report['sentinel_create_rows'] = sentinel_rows
            flush()
            if not (isinstance(sentinel_rows, list) and len(sentinel_rows) == 1
                    and isinstance(sentinel_rows[0], list) and len(sentinel_rows[0]) == 1
                    and isinstance(sentinel_rows[0][0], str) and sentinel_rows[0][0]):
                session._retain('Figure rollback sentinel acknowledgement invalid')
                raise ValueError('Invalid sentinel acknowledgement')
            sentinel_name = sentinel_rows[0][0]
            report['sentinel_name'] = sentinel_name
            report['sentinel_before'] = sentinel_preimage(
                inventory(output, 'sentinel-before'), sentinel_name, sentinel_token)

            staged_image = session._stage_image(source_image)
            missing_image = session.staging_root / 'definitely-missing-figure-probe.png'
            assert not missing_image.exists()
            assert sha(source_image) == image_digest

            stage('one-image-success-then-exact-rollback')
            report['one_image_rows'] = session._execute(
                scenario_commands('one-image', staged_image))
            assert validate_scenario(report['one_image_rows'], 'one-image', failed=False)
            report['checks']['one_image_success_then_exact_local_rollback'] = True

            stage('first-image-then-second-native-failure-and-exact-rollback')
            report['second_failure_rows'] = session._execute(
                scenario_commands('second-failure', staged_image, missing=missing_image))
            assert validate_scenario(
                report['second_failure_rows'], 'second-failure', failed=True)
            report['checks']['second_image_failure_then_exact_local_rollback'] = True
            assert sha(source_image) == image_digest
            report['checks']['resource_source_preserved'] = True

            stage('save-and-pdf-restored-preimage')
            session.save_copy(output_document)
            session.export_pdf(output / 'figure-rollback.pdf')
            assert sha(source_document) == source_digest
            report['checks']['owned_document_source_preserved'] = True
            report['sentinel_after'] = sentinel_preimage(
                inventory(output, 'sentinel-after'), sentinel_name, sentinel_token)
            assert report['sentinel_after'] == report['sentinel_before']
            report['checks']['exact_unsaved_sentinel_preserved'] = True

        close_sentinel_after_owned(session, output, report, sentinel_name, sentinel_token)
        sentinel_name = None

        stage('owned-reopen-restored-preimage')
        digest = sha(output_document)
        with open_document(
                output_document, engine='msoffice', read_only=True,
                visible=False) as reopened:
            reopened._retain_evidence = True
            report['reopen_position_rows'] = reopened._execute([
                f'set selection start of selection of boundWindow to {START}',
                f'set selection end of selection of boundWindow to {END}',
                'set nativeRows to {{"positioned",true}}',
            ])
            assert report['reopen_position_rows'] == [['positioned', True]]
            report['reopen_rows'] = reopened._execute(state_commands())
            assert validate_state(report['reopen_rows'])
        report['checks']['native_save_reopen_preserves_exact_preimage'] = (
            sha(output_document) == digest)
        report['checks']['owned_document_source_still_preserved_after_reopen'] = (
            sha(source_document) == source_digest)
        report['checks']['complete_body_style_theme_restored'] = (
            document_signature(source_document) == document_signature(output_document))
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
        for owner, label in ((creator, 'source-runtime'), (session, 'native-runtime'),
                             (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, output / label, dirs_exist_ok=True)
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
            if path.suffix in ('.docx', '.pdf', '.png')
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
        print('Native figure rollback probe requires --execute and the root Word lease', file=sys.stderr)
        return 2
    return 0 if run(args.output, execute=True)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
