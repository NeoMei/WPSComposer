#!/usr/bin/env python3
"""Bounded native PowerPoint feasibility probe, deliberately not a product backend.

Creates an owned presentation plus a synthetic unsaved sentinel. Never quits the
application, uses the clipboard, or closes unrelated presentations. On uncertainty
it retains documents for recovery. Each run requires a new output directory.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import plistlib
import subprocess
import sys
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile
import shutil
import tempfile
import struct
import zlib

APP = Path('/Applications/Microsoft PowerPoint.app')
STAGING_PARENT = Path.home() / 'Library/Containers/com.microsoft.Powerpoint/Data/Documents'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def quote(value):
    return '"' + str(value).replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n').replace('\r', '\\r') + '"'


def native_script(output, token, timeout):
    """Compile a closed synthetic fixture, never arbitrary user code."""
    source = '''with timeout of %d seconds
 tell application "Microsoft PowerPoint"
  set beforeState to {}
  repeat with priorIndex from 1 to (count of presentations)
   set end of beforeState to {name of presentation priorIndex, full name of presentation priorIndex, saved of presentation priorIndex, count of slides of presentation priorIndex}
  end repeat
  log "PROBE|baseline|" & (count of beforeState)
  log "PROBE|version|" & Version
  set sentinel to make new presentation
  set sentSlide to make new slide at end of sentinel with properties {layout:slide layout blank}
  set sentText to make new text box at end of sentSlide with properties {left position:40, top:40, width:500, height:80}
  set content of text range of text frame of sentText to %s
  set sentinelName to name of sentinel
  log "PROBE|sentinel_name|" & sentinelName
  set owned to make new presentation
  set ownedName to name of owned
  set slide width of page setup of owned to 720
  log "PROBE|owned_name|" & ownedName
  set s1 to make new slide at end of owned with properties {layout:slide layout blank}
  set s2 to make new slide at end of owned with properties {layout:slide layout blank}
  set titleBox to make new text box at end of s1 with properties {name:"parity-title", left position:40, top:40, width:560, height:60}
  set content of text range of text frame of titleBox to "Native editable PowerPoint"
  set font size of font of text range of text frame of titleBox to 28
  set bold of font of text range of text frame of titleBox to true
  set box to make new shape at end of s1 with properties {name:"parity-shape", auto shape type:autoshape rectangle, left position:40, top:140, width:200, height:90}
  log "PROBE|shape|PASS"
  set fore color of fill format of box to {32, 96, 160}
  log "PROBE|fill|PASS"
  set pic to make new picture at end of s1 with properties {name:"parity-image", file name:%s, link to file:false, save with document:true, left position:280, top:140, width:100, height:90}
  log "PROBE|picture|PASS"
  set tb to make new shape table at end of s2 with properties {name:"parity-table", number of rows:2, number of columns:2, left position:40, top:100, width:500, height:120}
  log "PROBE|table_created|PASS"
  set c to get cell from table object of tb row 1 column 1
  set content of text range of text frame of shape of c to "Native table"
  set c to get cell from table object of tb row 2 column 2
  set content of text range of text frame of shape of c to "42"
  set content of text range of text frame of shape 2 of notes page of s1 to "Native speaker note"
  log "PROBE|page_width|" & slide width of page setup of owned
  set firstId to slide ID of s1
  set secondId to slide ID of s2
  try
   set clonedSlide to duplicate s1
   if (count of slides of owned) is not 3 then error "Slide clone count mismatch"
   log clonedSlide
   if class of clonedSlide is list then set clonedSlide to item 1 of clonedSlide
   delete clonedSlide
   if (count of slides of owned) is not 2 then error "Slide remove count mismatch"
   log "PROBE|slide_clone_remove|PASS"
  on error em number en
   log "PROBE|slide_clone_remove|FAIL|" & en & "|" & em
  end try
  try
   move s2 to before s1
   if slide ID of slide 1 of owned is not secondId then error "Slide move order mismatch"
   move slide 1 of owned to after slide 2 of owned
   if slide ID of slide 1 of owned is not firstId then error "Slide move restore mismatch"
   log "PROBE|slide_move|PASS"
  on error em number en
   log "PROBE|slide_move|FAIL|" & en & "|" & em
  end try
  try
   set clonedShape to duplicate titleBox
   set name of clonedShape to "parity-clone"
   if (count of shapes of s1) is not 4 then error "Shape clone count mismatch"
   delete clonedShape
   if (count of shapes of s1) is not 3 then error "Shape remove count mismatch"
   log "PROBE|shape_clone_remove|PASS"
  on error em number en
   log "PROBE|shape_clone_remove|FAIL|" & en & "|" & em
  end try
  set left position of box to 60
  if left position of box is not 60 then error "Shape geometry edit failed"
  log "PROBE|shape_geometry_edit|PASS"
  log "PROBE|undo|BLOCKED_PRIOR_DESTRUCTIVE_UNDO"
  try
   set currentSelection to selection of document window 1 of owned
   log "PROBE|selection_type|" & (selection type of currentSelection as text)
  on error em number en
   log "PROBE|selection|FAIL|" & en & "|" & em
  end try
  log "PROBE|generation|PASS"
  save owned in ((POSIX file %s) as text) as save as Open XML presentation
  set owned to presentation "native.pptx"
  log "PROBE|saved_path|" & full name of owned
  if saved of owned is false then error "Save did not mark owned presentation saved"
  if full name of owned does not end with "native.pptx" then error "Save did not bind owned output path"
  save owned in pdfPath as save as PDF
  log "PROBE|pdf_before_reopen|PASS"
  if saved of sentinel then error "Sentinel unexpectedly saved"
  log "PROBE|sentinel_unsaved_before_reopen|PASS"
  close owned saving no
  open (POSIX file %s)
  set owned to presentation %s
  if full name of owned does not end with %s then error "Owned reopen identity mismatch"
  log "PROBE|reopened_count|" & (count of slides of owned)
  log "PROBE|reopened_text|" & content of text range of text frame of shape "parity-title" of slide 1 of owned
  save owned in ((POSIX file %s) as text) as save as PDF
  log "PROBE|pdf|PASS"
  close owned saving no
  if saved of sentinel then error "Sentinel unexpectedly saved"
  if content of text range of text frame of shape 1 of slide 1 of sentinel is not %s then error "Sentinel content changed"
  repeat with priorIndex from 1 to (count of beforeState)
   set afterRow to {name of presentation priorIndex, full name of presentation priorIndex, saved of presentation priorIndex, count of slides of presentation priorIndex}
   if afterRow is not item priorIndex of beforeState then
    log item priorIndex of beforeState
    log afterRow
    error "Unrelated presentation changed"
   end if
  end repeat
  log "PROBE|preservation|PASS"
  return "DONE"
 end tell
end timeout
''' % (max(1, int(timeout)), quote('SENTINEL-' + token), quote(output / 'source.png'),
       quote(output / 'native.pptx'), quote(output / 'native.pptx'), quote('native.pptx'),
       quote('native.pptx'), quote(output / 'native.pdf'), quote('SENTINEL-' + token))
    source = source.replace('((POSIX file ' + quote(output / 'native.pptx') + ') as text)', 'pptPath').replace('((POSIX file ' + quote(output / 'native.pdf') + ') as text)', 'pdfPath').replace('(POSIX file ' + quote(output / 'native.pptx') + ')', 'pptAlias')
    source = source.replace('open pptAlias', 'end tell\n set reopenedAlias to pptAlias as alias\n tell application \"Microsoft PowerPoint\"\n  open reopenedAlias')
    source = source.replace('native.pptx', 'native-' + token + '.pptx')
    return 'set pptAlias to POSIX file ' + quote(output / ('native-' + token + '.pptx')) + '\nset pptPath to pptAlias as text\nset pdfPath to (POSIX file ' + quote(output / 'native.pdf') + ') as text\n' + source


def write_source_image(path):
    """Small RGB fixture image, not a rendered slide or an Office generator."""
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    data = b''.join(b'\0' + bytes([32, 160, 96]) * 32 for _ in range(32))
    Path(path).write_bytes(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 32, 32, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(data)) + chunk(b'IEND', b''))


def inspect_artifacts(output):
    checks = {"native_pptx_exists": (output / "native.pptx").is_file(), "native_pdf_exists": (output / "native.pdf").is_file()}
    detail = {}
    path = output / 'native.pptx'
    if path.exists():
        with zipfile.ZipFile(path) as package:
            slides = sorted(n for n in package.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
            roots = [ET.fromstring(package.read(n)) for n in slides]
            texts = ['|'.join(r.itertext()) for r in roots]
            checks['two_native_slides'] = len(roots) == 2
            checks['editable_text'] = any('Native editable PowerPoint' in t for t in texts)
            checks['native_picture'] = any(r.find('.//p:pic', NS) is not None for r in roots)
            checks['native_table'] = any(r.find('.//a:tbl', NS) is not None for r in roots)
            notes = [n for n in package.namelist() if n.startswith('ppt/notesSlides/notesSlide') and n.endswith('.xml')]
            checks['native_notes'] = any(b'Native speaker note' in package.read(n) for n in notes)
            shapes = {}
            for root in roots:
                for shape in root.findall('.//p:sp', NS):
                    identity = shape.find('p:nvSpPr/p:cNvPr', NS)
                    if identity is not None:
                        shapes[identity.get('name')] = shape
            title = shapes.get('parity-title')
            run = title.find('.//a:rPr', NS) if title is not None else None
            checks['title_font_28pt_bold'] = run is not None and run.get('sz') == '2800' and run.get('b') == '1'
            box = shapes.get('parity-shape')
            off = box.find('p:spPr/a:xfrm/a:off', NS) if box is not None else None
            extent = box.find('p:spPr/a:xfrm/a:ext', NS) if box is not None else None
            fill = box.find('p:spPr/a:solidFill/a:srgbClr', NS) if box is not None else None
            checks['shape_geometry_and_fill'] = (off is not None and off.attrib == {'x':'762000','y':'1778000'} and extent is not None and extent.attrib == {'cx':'2540000','cy':'1143000'} and fill is not None and fill.get('val') == '2060A0')
            size = ET.fromstring(package.read('ppt/presentation.xml')).find('p:sldSz', NS)
            checks['page_size_720x540'] = size is not None and size.get('cx') == '9144000' and size.get('cy') == '6858000'
            detail['slide_text'] = texts
    if (output / 'native.pdf').exists():
        try:
            import fitz
            with fitz.open(output / 'native.pdf') as pdf:
                detail['pdf_text'] = [page.get_text() for page in pdf]
                checks['pdf_two_pages'] = len(pdf) == 2
                checks['pdf_content'] = 'Native editable PowerPoint' in '\n'.join(detail['pdf_text']) and 'Native table' in '\n'.join(detail['pdf_text'])
                for i, page in enumerate(pdf):
                    page.get_pixmap(matrix=fitz.Matrix(1, 1)).save(output / ('page-%d.png' % (i + 1)))
        except ImportError:
            checks['pdf_inspection_available'] = False
    return {'checks': checks, **detail}


def run_probe(output_dir, timeout=90):
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('timeout must be finite and positive')
    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=False)
    token = uuid.uuid4().hex
    write_source_image(output / 'source.png')
    script = output / 'native.applescript'
    staging_parent = STAGING_PARENT if sys.platform == 'darwin' else output
    staging_parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix='wpscomposer-parity-', dir=staging_parent))
    shutil.copy2(output / 'source.png', staging / 'source.png')
    script.write_text(native_script(staging, token, timeout), encoding='utf-8')
    report = {'status': 'FAIL', 'token': token, 'source_sha256': sha256(__file__),
              'native_ui_acceptance': 'NOT_RUN', 'recovery': {'may_have_open_owned_documents': True},
              'application': str(APP), 'timeout_seconds': timeout, 'staging_dir': str(staging)}
    dictionary = APP / 'Contents/Resources/PowerPoint.sdef'
    if dictionary.exists():
        report['dictionary_sha256'] = sha256(dictionary)
    info = APP / 'Contents/Info.plist'
    if info.exists():
        report['application_version'] = plistlib.loads(info.read_bytes()).get('CFBundleShortVersionString')
    started = time.monotonic()
    stdout, stderr = '', ''
    try:
        if sys.platform != 'darwin':
            raise RuntimeError('Native PowerPoint probe requires macOS')
        result = subprocess.run(['/usr/bin/osascript', str(script)], capture_output=True, text=True, timeout=timeout)
        stdout, stderr = result.stdout, result.stderr
        report['returncode'] = result.returncode
        for artifact in ('native.pptx', 'native.pdf'):
            staged_file = staging / (('native-' + token + '.pptx') if artifact == 'native.pptx' else artifact)
            if staged_file.is_file():
                shutil.copy2(staged_file, output / artifact)
        report['artifacts'] = inspect_artifacts(output)
        if result.returncode == 0 and stdout.strip() == 'DONE' and report['artifacts']['checks'] and all(report['artifacts']['checks'].values()):
            report['status'] = 'CORE_PASS_PARITY_GAPS'
            report['recovery']['may_have_open_owned_documents'] = False
    except subprocess.TimeoutExpired as exc:
        report['status'] = 'UNCERTAIN_TIMEOUT'
        stdout = exc.stdout or ''
        stderr = exc.stderr or ''
    except Exception as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
    for name, value in [('native.stdout.txt', stdout), ('native.stderr.txt', stderr)]:
        (output / name).write_text(value.decode('utf-8', errors='replace') if isinstance(value, bytes) else value, encoding='utf-8')
    native_log = stderr.decode('utf-8', errors='replace') if isinstance(stderr, bytes) else stderr
    report['native_steps'] = [line.split('|')[1:] for line in native_log.splitlines() if line.startswith('PROBE|')]
    report['parity_gaps'] = ['Native slide clone command failed -50', 'Script text Undo is not a verified rollback', 'Native UI edit/undo acceptance not run']
    report['elapsed_seconds'] = time.monotonic() - started
    report['checksums'] = {p.name: sha256(p) for p in output.iterdir() if p.is_file()}
    (output / 'report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', required=True, type=Path)
    parser.add_argument('--timeout', type=float, default=90)
    args = parser.parse_args()
    result = run_probe(args.output_dir, timeout=args.timeout)
    print(json.dumps(result, indent=2, ensure_ascii=False))
    raise SystemExit(0 if result['status'] == 'PASS' else 1)
