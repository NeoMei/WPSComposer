"""Two isolated native hypotheses, never production support or automatic repair.

Word.sdef supplies window.selection, selection start/end/document, create range,
create new field, bookmark/table text object, and table deletion. Dictionary
presence and compilation are not native proof. Run only with an exclusive Word
lease and --execute. A failed Delete observation is retained without CR repair.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, _JSON
from skills.WPSComposer.scripts.msoffice.macos_word_recovery import hash_commands
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage

MODES = ('bound-selection', 'middle-table-recovery')
DICTIONARY = Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
MARKER = 'native-table-marker'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SOURCES = [Path(__file__), ROOT/'tests/msoffice/test_macos_word_quality_feasibility.py',
           *sorted((ROOT/'skills/WPSComposer').rglob('*.py')),
           *[ROOT/('fixtures/microsoft_parity/'+p) for p in
             ('macos_word_recovery.py', 'macos_word_fields.py', 'macos_word_paragraph_rule.py', 'macos_word_sections.py')],
           ROOT/'pyproject.toml']


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def source_pair_matches(live, retained, digest):
    return sha(live) == sha(retained) == digest


def exact(expression, literal):
    return f"((current application's NSString's stringWithString:({expression}))'s isEqualToString:{apple_string(literal)})"


def bound_guard(path):
    return [f'if not {exact("posix full name of boundDoc as text", path)} then error "QUALITY_BOUND_PATH_CHANGED"',
            f'if not {exact("posix full name of document of boundWindow as text", path)} then error "QUALITY_WINDOW_CHANGED"']


def selection_commands(path, window_id, sentinel_id):
    # IDs are diagnostic only when native Word returns missing value. Exact
    # document paths plus the synthetic sentinel identity are the binding proof.
    return bound_guard(path) + [
        'set qualitySelection to selection of boundWindow',
        'set qualityRange to text object of qualitySelection',
        'set qualityStart to start of content of qualityRange',
        'set qualityEnd to end of content of qualityRange',
        'set activeRange to text object of selection of active window',
        f'set selectionOwned to {exact("posix full name of document of qualitySelection as text", path)} as boolean',
        'if not selectionOwned then error "QUALITY_SELECTION_FOREIGN"',
        'if story type of qualitySelection is not main text story then error "QUALITY_SELECTION_WRONG_STORY"',
        'if qualityStart is not 9 or qualityEnd is not 9 then error "QUALITY_SELECTION_MOVED"',
        'set beforeEnd to end of content of text object of boundDoc',
        *hash_commands('content of text object of boundDoc as text', 'beforeHash'),
        'set qualityAnchor to create range boundDoc start qualityEnd end qualityEnd',
        'make new bookmark at boundDoc with properties {name:"wpsc_document_quality_anchor",text object:qualityAnchor}',
        'set qualityBookmark to bookmark "wpsc_document_quality_anchor" of boundDoc',
        *hash_commands('content of text object of boundDoc as text', 'afterHash'),
        'set nativeRows to {{"selection",posix full name of boundDoc as text,id of boundWindow,id of active window,'
        'qualityEnd,start of content of activeRange,end of content of activeRange,beforeEnd,'
        'start of bookmark of qualityBookmark,end of bookmark of qualityBookmark,empty of qualityBookmark,'
        'selectionOwned,activeSentinelVerified,beforeHash,afterHash}}',
    ]


def seed_commands(mode):
    if mode not in MODES:
        raise ValueError('Unknown mode')
    if mode == 'bound-selection':
        return ['activate object boundWindow',
                'set content of text object of boundDoc to "PREFIX 中文😀" & return & "MIDDLE QUALITY ANCHOR" & return & "SUFFIX remains" & return',
                'set selection start of selection of boundWindow to 9',
                'set selection end of selection of boundWindow to 9',
                'set nativeRows to {{"seed",true}}']
    texts = ['PREFIX 中文😀', 'PREFIX FIELD ', 'PREFIX TABLE SLOT', 'BEFORE BOUNDARY',
             'AFTER BOUNDARY', 'SUFFIX FIELD ', 'SUFFIX TABLE SLOT', 'TAIL 中文😀']
    lines = ['activate object boundWindow', f'set content of text object of boundDoc to {apple_string(chr(13).join(texts)+chr(13))}']
    for index in range(1, 9):
        lines += [f'set seedRange to text object of paragraph {index} of boundDoc',
                  f'make new bookmark at boundDoc with properties {{name:"quality_seed_{index}",text object:seedRange}}']
    for index in (6, 2):
        lines += [f'set seedPoint to (end of content of text object of bookmark "quality_seed_{index}" of boundDoc) - 1',
                  'set seedRange to create range boundDoc start seedPoint end seedPoint',
                  f'create new field text range seedRange field type field sequence field text "Quality{index}" preserve formatting true']
    for index in (7, 3):
        lines += [f'set seedPoint to start of content of text object of bookmark "quality_seed_{index}" of boundDoc',
                  'set seedRange to create range boundDoc start seedPoint end seedPoint',
                  'set seedTable to make new table at boundDoc with properties {text object:seedRange,number of rows:1,number of columns:1}',
                  f'set content of text object of (get cell from table seedTable row 1 column 1) to "EXISTING TABLE {index}"']
    for index, indent, before, after in ((4, 17, 11, 7), (5, 23, 13, 9)):
        lines += [f'set seedRange to text object of bookmark "quality_seed_{index}" of boundDoc',
                  'set seedFormat to paragraph format of seedRange',
                  f'set first line indent of seedFormat to {indent}',
                  f'set space before of seedFormat to {before}', f'set space after of seedFormat to {after}',
                  'set line spacing rule of seedFormat to line space exactly', 'set line spacing of seedFormat to 21',
                  'set keep with next of seedFormat to true', 'set widow control of seedFormat to false',
                  'set italic of font object of seedRange to true', 'set font size of font object of seedRange to 14']
    return lines + ['set seedPoint to start of content of text object of bookmark "quality_seed_5" of boundDoc',
                    'set seedRange to create range boundDoc start seedPoint end seedPoint',
                    'make new bookmark at boundDoc with properties {name:"wpsc_quality_splice",text object:seedRange}',
                    'set nativeRows to {{"seed",true}}']


def snapshot_commands():
    """Full synthetic body and all paragraphs/fields/tables/bookmarks; no repair."""
    lines = [*hash_commands('content of text object of boundDoc as text', 'bodyHash'),
             'set nativeRows to {{"body",end of content of text object of boundDoc,bodyHash,count paragraphs of boundDoc,count tables of boundDoc,count fields of boundDoc,count bookmarks of boundDoc}}',
             'repeat with qi from 1 to count paragraphs of boundDoc',
             'set qr to text object of paragraph qi of boundDoc',
             *hash_commands('content of qr as text', 'paragraphHash'),
             'set qt to content of qr as text', 'set terminalIDs to {}',
             'if (length of qt) > 0 then set terminalIDs to {id of character -1 of qt}',
             'if (length of qt) > 1 then set terminalIDs to {id of character -2 of qt} & terminalIDs',
             'set qf to paragraph format of qr',
             'set qfmt to {name local of style of qr as text,first line indent of qf,space before of qf,space after of qf,'
             'line spacing of qf,line spacing rule of qf as text,alignment of qf as text,keep with next of qf,keep together of qf,'
             'widow control of qf,outline level of qf as text,italic of font object of qr,font size of font object of qr}',
             'set end of nativeRows to {"paragraph",qi as integer,start of content of qr,end of content of qr,paragraphHash,terminalIDs,qfmt}',
             'end repeat',
             'repeat with qi from 1 to count fields of boundDoc', 'set qfield to field qi of boundDoc',
             *hash_commands('content of field code of qfield as text', 'codeHash'),
             *hash_commands('content of result range of qfield as text', 'resultHash'),
             'set end of nativeRows to {"field",qi as integer,field type of qfield as text,start of content of field code of qfield,'
             'end of content of field code of qfield,start of content of result range of qfield,end of content of result range of qfield,codeHash,resultHash,locked of qfield}',
             'end repeat', 'repeat with qi from 1 to count tables of boundDoc', 'set qtable to table qi of boundDoc',
             *hash_commands('content of text object of qtable as text', 'tableHash'),
             'set end of nativeRows to {"table",qi as integer,start of content of text object of qtable,end of content of text object of qtable,count rows of qtable,count columns of qtable,tableHash}',
             'end repeat', 'repeat with qi from 1 to count bookmarks of boundDoc', 'set qb to bookmark qi of boundDoc',
             *hash_commands('content of text object of qb as text', 'bookmarkHash'),
             'set end of nativeRows to {"bookmark",name of qb as text,start of bookmark of qb,end of bookmark of qb,bookmarkHash}',
             'end repeat', 'set end of nativeRows to {"end"}']
    return lines


def middle_commands():
    return ['activate object boundWindow',
            'set qualityPoint to start of bookmark of bookmark "wpsc_quality_splice" of boundDoc',
            'set qualityBeforeEnd to end of content of text object of boundDoc',
            'if qualityPoint <= 0 or qualityPoint >= qualityBeforeEnd - 1 then error "QUALITY_NOT_MIDDLE"',
            *snapshot_commands(), 'set beforeRows to nativeRows',
            'if (count fields of boundDoc) is not 2 or (count tables of boundDoc) is not 2 then error "QUALITY_SEED_TOPOLOGY"',
            'set qualityTarget to create range boundDoc start qualityPoint end qualityPoint',
            'set qualityNewTable to make new table at boundDoc with properties {text object:qualityTarget,number of rows:1,number of columns:1}',
            f'set content of text object of (get cell from table qualityNewTable row 1 column 1) to {apple_string(MARKER)}',
            'set qualityCreatedStart to start of content of text object of qualityNewTable',
            'set qualityCreatedEnd to end of content of text object of qualityNewTable',
            'set qualityCellStart to start of content of text object of (get cell from table qualityNewTable row 1 column 1)',
            f'set qualityDisplay to create range boundDoc start qualityCellStart end (qualityCellStart + {len(MARKER.encode("utf-16-le"))//2})',
            'set qualityCreatedText to content of qualityDisplay as text',
            'try', 'error "QUALITY_CONTROLLED_ORDINARY_FAILURE" number -2700',
            'on error qualityError number qualityNumber',
            'if qualityNumber is not -2700 then error qualityError number qualityNumber',
            'set injectedNumber to qualityNumber', 'end try',
            *snapshot_commands(), 'set partialRows to nativeRows',
            'set deleteNumber to 0', 'try', 'delete qualityNewTable',
            'on error deleteError number caughtDeleteNumber',
            'if caughtDeleteNumber is -1712 or caughtDeleteNumber is -609 or caughtDeleteNumber is -128 then error deleteError number caughtDeleteNumber',
            'set deleteNumber to caughtDeleteNumber', 'end try',
            *snapshot_commands(), 'set afterRows to nativeRows',
            'set nativeRows to {{"middle",qualityPoint,qualityBeforeEnd,qualityCreatedStart,qualityCreatedEnd,qualityCreatedText,injectedNumber,deleteNumber},'
            '{"before",beforeRows},{"partial",partialRows},{"after-delete",afterRows}}']


def digest(value):
    return isinstance(value, str) and re.fullmatch('[0-9a-f]{64}', value) is not None


def state_valid(rows):
    if not isinstance(rows, list) or len(rows) < 2 or rows[-1] != ['end']:
        return False
    head = rows[0]
    if not (isinstance(head, list) and len(head) == 7 and head[0] == 'body'
            and digest(head[2]) and all(type(head[i]) is int and head[i] >= 0 for i in (1, 3, 4, 5, 6))):
        return False
    sizes = {'paragraph': 7, 'field': 10, 'table': 7, 'bookmark': 5}
    counts = {kind: 0 for kind in sizes}
    names = set()
    for row in rows[1:-1]:
        if not isinstance(row, list) or not row or len(row) != sizes.get(row[0]):
            return False
        kind = row[0]
        counts[kind] += 1
        if kind != 'bookmark' and (type(row[1]) is not int or row[1] != counts[kind]):
            return False
        positions = (3, 4, 5, 6) if kind == 'field' else (2, 3)
        if any(type(row[i]) is not int or not 0 <= row[i] <= head[1] for i in positions):
            return False
        if any(row[a] > row[b] for a, b in zip(positions, positions[1:])):
            return False
        if row[0] == 'paragraph':
            if not (digest(row[4]) and isinstance(row[5], list) and isinstance(row[6], list)
                    and all(type(v) is int and 0 <= v <= 0x10ffff for v in row[5])):
                return False
            if any(type(v) is float and not math.isfinite(v) for v in row[6]):
                return False
        elif row[0] == 'field':
            if not (isinstance(row[2], str) and digest(row[7]) and digest(row[8]) and type(row[9]) is bool):
                return False
        elif not digest(row[-1]):
            return False
        if kind == 'table' and any(type(row[i]) is not int or row[i] < 1 for i in (4, 5)):
            return False
        if kind == 'bookmark':
            if not isinstance(row[1], str) or not row[1] or row[1] in names:
                return False
            names.add(row[1])
    return [counts[kind] for kind in ('paragraph', 'table', 'field', 'bookmark')] == head[3:]


def assess_middle(rows):
    result = {'observed_middle': False, 'restored': False}
    if not (isinstance(rows, list) and len(rows) == 4 and isinstance(rows[0], list) and len(rows[0]) == 8):
        return result
    h = rows[0]
    if not (h[0] == 'middle' and all(type(h[i]) is int for i in (1, 2, 3, 4, 6, 7))
            and h[5] == MARKER and h[6] == -2700):
        return result
    if any(not isinstance(r, list) or len(r) != 2 or r[0] != label or not state_valid(r[1])
           for r, label in zip(rows[1:], ('before', 'partial', 'after-delete'))):
        return result
    before, partial, after = [r[1] for r in rows[1:]]
    result['observed_middle'] = (0 < h[1] < h[2] - 1 and h[2] == before[0][1]
        and h[1] <= h[3] < h[4] and partial[0][1] > before[0][1]
        and partial[0][4] == before[0][4] + 1 and partial[0][2] != before[0][2])
    # Comparing serialized typed data avoids Python treating True == 1.
    result['restored'] = (result['observed_middle'] and h[7] == 0
        and json.dumps(before, ensure_ascii=False) == json.dumps(after, ensure_ascii=False))
    result['delete_error_number'] = h[7]
    return result


def assess_selection(rows, path, window_id, sentinel_id):
    ok = isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list) and len(rows[0]) == 15
    if ok:
        r = rows[0]
        ok = (r[:4] == ['selection', path, window_id, sentinel_id]
              and all(type(r[i]) is int for i in (4, 5, 6, 7, 8, 9))
              and r[4] == 9 and r[5:7] == [2, 2] and r[7] > 10 and r[8:10] == [9, 9]
              and all(v is True for v in r[10:13]) and digest(r[13]) and r[13] == r[14])
    return {'reserved': bool(ok)}


def xml_observation(before, after):
    root = ET.fromstring(after)
    starts = [e for e in root.iter(W+'bookmarkStart') if e.get(W+'name') == 'wpsc_document_quality_anchor']
    collapsed = False
    if len(starts) == 1:
        for parent in root.iter():
            children = list(parent)
            if starts[0] in children:
                index = children.index(starts[0])
                collapsed = (index+1 < len(children) and children[index+1].tag == W+'bookmarkEnd'
                             and children[index+1].get(W+'id') == starts[0].get(W+'id'))
    return {'document_xml_exact': before == after, 'collapsed_quality_bookmark': collapsed}


def sentinel_guard(name, token):
    return [f'set qualitySentinel to document {apple_string(name)}',
            f'if not {exact("name of qualitySentinel as text", name)} then error "QUALITY_SENTINEL_NAME"',
            f'if not {exact("path of qualitySentinel as text", "")} then error "QUALITY_SENTINEL_SAVED"',
            f'if not {exact("content of text object of qualitySentinel as text", token+chr(13))} then error "QUALITY_SENTINEL_TEXT"',
            'if saved of qualitySentinel then error "QUALITY_SENTINEL_SAVED"']


def close_sentinel_after_owned(session, output, report, name, token):
    if not session._closed or session._quarantined:
        raise RuntimeError('Owned close unverified; sentinel retained')
    observed = inventory(output, 'after-owned-close')
    if ([r for r in observed if r[0] == name] != [report['sentinel_before']]
            or [r for r in observed if r[0] != name] != report['inventory_before']):
        raise RuntimeError('Inventory or sentinel changed; retained')
    script = output/'sentinel-close.applescript'
    script.write_text(_JSON + '\ntell application "Microsoft Word"\n' + '\n'.join(
        sentinel_guard(name, token) + ['close qualitySentinel saving no', 'set nativeRows to {{"sentinel-closed"}}'])
        + '\nend tell\nreturn my jsonRows(nativeRows)\n')
    result = subprocess.run(['/usr/bin/osascript', str(script)], capture_output=True, text=True, timeout=30)
    script.with_suffix('.log').write_text(result.stdout+'\n'+result.stderr)
    if result.returncode or json.loads(result.stdout) != [['sentinel-closed']]:
        raise RuntimeError('Sentinel close acknowledgement invalid')
    report['sentinel_closed'] = True


def run(output, mode, *, execute=False):
    if not execute:
        raise ValueError('Explicit execute is required')
    if mode not in MODES:
        raise ValueError('Unknown mode')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'status': 'FAIL', 'mode': mode, 'checks': {}, 'steps': [],
              'scope': 'feasibility only; no production helper, fallback, CR repair, UI or full parity proof',
              'source_hashes': retain_sources(output, SOURCES)}
    session = reopened = None
    name = token = None

    def flush():
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))

    def step(label, commands):
        report['current_step'] = label
        flush()
        rows = session._execute(bound_guard(session._bound_path) + commands)
        report['steps'].append({'label': label, 'rows': rows})
        flush()
        return rows

    try:
        report['dictionary_sha256'] = sha(DICTIONARY)
        shutil.copy2(DICTIONARY, output/'localWord.sdef')
        report['inventory_before'] = inventory(output, 'before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence = True
            if step('seed', seed_commands(mode)) != [['seed', True]]:
                raise RuntimeError('Seed acknowledgement invalid')
            token = 'QUALITY SENTINEL 中文😀 '+uuid4().hex
            rows = step('sentinel-create', ['set qualitySentinel to make new document',
                f'set content of text object of qualitySentinel to {apple_string(token)}',
                'set sentinelWindow to active window of qualitySentinel', 'activate object sentinelWindow',
                'set selection start of selection of sentinelWindow to 2',
                'set selection end of selection of sentinelWindow to 2',
                'set nativeRows to {{name of qualitySentinel as text,id of sentinelWindow,version as text}}'])
            if not (len(rows) == 1 and len(rows[0]) == 3 and isinstance(rows[0][0], str)):
                raise RuntimeError('Sentinel creation unverified')
            name, sentinel_id, report['word_version'] = rows[0]
            report['sentinel_name'] = name
            report['sentinel_before'] = sentinel_preimage(inventory(output, 'sentinel-before'), name, token)
            # Native saves are evidence checkpoints only, never replacement/repair.
            session.save_docx(output/'before.docx')
            if mode == 'bound-selection':
                active_guard = sentinel_guard(name, token) + [
                    'activate object (active window of qualitySentinel)',
                    f'set activeSentinelVerified to {exact("name of document of active window as text", name)} as boolean',
                    'if not activeSentinelVerified then error "QUALITY_SENTINEL_NOT_ACTIVE"']
                rows = step('bound-selection', active_guard + selection_commands(session._bound_path, session._window_id, sentinel_id))
                report['checks'].update(assess_selection(rows, session._bound_path, session._window_id, sentinel_id))
            else:
                rows = step('middle-table-delete-only', middle_commands())
                report['checks'].update(assess_middle(rows))
                before = rows[1][1]
                point = rows[0][1]
                fields = [r for r in before if r[0] == 'field']
                tables = [r for r in before if r[0] == 'table']
                report['checks']['seed_has_objects_both_sides'] = (len(fields) == len(tables) == 2
                    and any(r[3] < point for r in fields) and any(r[3] > point for r in fields)
                    and any(r[2] < point for r in tables) and any(r[2] > point for r in tables))
            # Always retain the raw Delete result, even when restoration is false.
            session.save_docx(output/'after.docx')
            report['sentinel_after_observation'] = sentinel_preimage(inventory(output, 'sentinel-after-observation'), name, token)
            report['checks']['sentinel_unchanged'] = report['sentinel_after_observation'] == report['sentinel_before']
        close_sentinel_after_owned(session, output, report, name, token)
        name = None
        before_xml = after_xml = None
        for label in ('before', 'after'):
            with ZipFile(output/(label+'.docx')) as package:
                for member in ('word/document.xml', 'word/styles.xml'):
                    data = package.read(member)
                    (output/(label+'-'+Path(member).name)).write_bytes(data)
                    if member == 'word/document.xml':
                        if label == 'before': before_xml = data
                        else: after_xml = data
        report['xml_observation'] = xml_observation(before_xml, after_xml)
        if mode == 'middle-table-recovery':
            report['checks']['document_xml_exact'] = report['xml_observation']['document_xml_exact']
            report['checks']['styles_xml_exact'] = (output/'before-styles.xml').read_bytes() == (output/'after-styles.xml').read_bytes()
        else:
            report['checks']['collapsed_anchor_xml'] = report['xml_observation']['collapsed_quality_bookmark']
        source_digest = sha(output/'after.docx')
        with MacWordSession.open_document(output/'after.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            rows = reopened._execute(bound_guard(reopened._bound_path) + snapshot_commands())
            report['reopen_snapshot'] = rows
            if mode == 'bound-selection':
                anchor = [r for r in rows if r[0] == 'bookmark' and r[1] == 'wpsc_document_quality_anchor']
                report['checks']['reopened_collapsed_anchor'] = len(anchor) == 1 and anchor[0][2:4] == [9, 9]
            else:
                report['checks']['reopen_matches_after_delete'] = rows == report['steps'][-1]['rows'][3][1]
        report['checks']['reopen_source_unchanged'] = sha(output/'after.docx') == source_digest
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and name:
            try:
                close_sentinel_after_owned(session, output, report, name, token)
                name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, output/label, dirs_exist_ok=True)
        report['remaining_sentinel'] = name
        report['checks']['sources_unchanged'] = all(source_pair_matches(ROOT/p, output/'source'/p, value)
                                                   for p, value in report['source_hashes'].items())
        if 'dictionary_sha256' in report:
            report['checks']['dictionary_unchanged'] = source_pair_matches(DICTIONARY, output/'localWord.sdef', report['dictionary_sha256'])
        if not name and not report.get('cleanup_failure'):
            try:
                report['inventory_final'] = inventory(output, 'final')
                report['checks']['inventory_preserved'] = report['inventory_final'] == report.get('inventory_before')
            except BaseException:
                report['inventory_failure'] = traceback.format_exc()
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix in ('.docx', '.xml', '.sdef')}
        # A diagnostic error number of zero is not itself a false pass check.
        checks = {k: v for k, v in report['checks'].items() if k != 'delete_error_number'}
        if checks and all(v is True for v in checks.values()) and not any(k in report for k in ('error', 'cleanup_failure', 'inventory_failure')) and not name:
            report['status'] = 'PASS'
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--mode', choices=MODES, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Native feasibility requires --execute and the root Word lease', file=sys.stderr)
        return 2
    return 0 if run(args.output, args.mode, execute=True)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
