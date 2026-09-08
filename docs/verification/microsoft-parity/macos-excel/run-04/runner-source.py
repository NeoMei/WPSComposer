"""Bounded, native AppleScript Excel probe; this is not a production adapter.

Runs only against synthetic workbooks and retains raw diagnostics on failure.
No static workbook writer, GUI automation, document macros or global settings.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import plistlib
import subprocess
import shutil
import tempfile
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET

APP = Path('/Applications/Microsoft Excel.app')


def literal(value):
    return '"' + str(value).replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n').replace('\r', '\\r') + '"'


def compile_probe(output_dir, token, timeout):
    output_dir = Path(output_dir)
    xlsx = literal(output_dir / 'native.xlsx')
    pdf = literal(output_dir / 'native.pdf')
    return f'''with timeout of {timeout} seconds
 tell application "{APP}"
  set previousBooks to {{}}
  repeat with existingBook in workbooks
   set end of previousBooks to {{name of existingBook, full name of existingBook, saved of existingBook}}
  end repeat
  log "before_books=" & (count of previousBooks)
  set sentinelToken to {literal('sentinel-' + token)}
  set ownedToken to {literal('owned-' + token)}
  set sentinelBook to make new workbook
  set sentinelName to name of sentinelBook
  set value of range "A1" of worksheet 1 of sentinelBook to sentinelToken
  log "sentinel_name=" & sentinelName
  set sentinelPath to full name of sentinelBook
  set ownedBook to make new workbook
  set ownedName to name of ownedBook
  set dataSheet to worksheet 1 of ownedBook
  set name of dataSheet to "Data"
  set value of range "D20" of dataSheet to ownedToken
  log "owned_name=" & ownedName
  log "stage=data"
  set value of range "A1:B4" of dataSheet to {{{{"Month", "Revenue"}}, {{"Jan", 10}}, {{"Feb", 20}}, {{"Mar", 30}}}}
  set value of range "A6" of dataSheet to "Total"
  set formula of range "B6" of dataSheet to "=SUM(B2:B4)"
  tell ownedBook
   make new worksheet at end with properties {{name:"Summary"}}
  end tell
  set summarySheet to worksheet "Summary" of ownedBook
  set value of range "A1" of summarySheet to "Cross sheet total"
  set formula of range "B1" of summarySheet to "=Data!B6*2"
  set column width of range "A:A" of summarySheet to 24
  set column width of range "B:B" of summarySheet to 18
  set value of range "A8" of dataSheet to "Microsoft native parity"
  merge range "A8:D8" of dataSheet
  set bold of font object of range "A1:B1" of dataSheet to true
  set color of interior object of range "A1:B1" of dataSheet to {{220, 235, 250}}
  set number format of range "B2:B6" of dataSheet to "0.00"
  set column width of range "A:D" of dataSheet to 18
  set row height of range "1:1" of dataSheet to 28
  calculate dataSheet
  calculate summarySheet
  if value of range "B6" of dataSheet is not 60 then error "Formula value mismatch"
  if value of range "B1" of summarySheet is not 120 then error "Cross sheet value mismatch"
  log "formula_calculation=pass"
  log "stage=chart"
  tell dataSheet
   make new chart object at end with properties {{left position:20, top:160, width:400, height:220}}
  end tell
  set chartContainer to chart object 1 of dataSheet
  set nativeChart to chart of chartContainer
  set chart type of nativeChart to column clustered
  set source data nativeChart source range "A1:B4" of dataSheet plot by columns
  set has title of nativeChart to true
  set chart title text of chart title of nativeChart to "Revenue by month"
  log "chart=pass"
  log "stage=structure"
  try
   insert into range (range "10:10" of dataSheet)
   set value of range "A10" of dataSheet to "Inserted row"
   if value of range "D21" of dataSheet is not ownedToken then error "Row insertion lost marker"
   delete range (range "10:10" of dataSheet)
   if value of range "D20" of dataSheet is not ownedToken then error "Row deletion lost marker"
   log "row_insert_delete=pass"
  on error messageText number errorNumber
   log "row_insert_delete=failed:" & errorNumber & ":" & messageText
  end try
  try
   tell ownedBook
    make new worksheet at end with properties {{name:"Temporary"}}
   end tell
   set temporarySheet to worksheet "Temporary" of ownedBook
   set name of temporarySheet to "Renamed"
   set temporarySheet to worksheet "Renamed" of ownedBook
   copy worksheet temporarySheet after temporarySheet
   if count of worksheets of ownedBook is not 4 then error "Sheet clone count mismatch"
   delete worksheet "Renamed (2)" of ownedBook
   delete worksheet "Renamed" of ownedBook
   if count of worksheets of ownedBook is not 2 then error "Sheet removal count mismatch"
   log "sheet_create_rename_clone_delete=pass"
  on error messageText number errorNumber
   log "sheet_create_rename_clone_delete=failed:" & errorNumber & ":" & messageText
  end try
  try
   activate object dataSheet
   select range "A1:B4" of dataSheet
   set selectionAddress to get address selection
   if selectionAddress is not "$A$1:$B$4" then error "Selection mismatch"
   log "selection=pass:" & selectionAddress
  on error messageText number errorNumber
   log "selection=failed:" & errorNumber & ":" & messageText
  end try
  set print area of page setup object of dataSheet to "$A$1:$F$28"
  set zoom of page setup object of dataSheet to false
  set fit to pages wide of page setup object of dataSheet to 1
  set fit to pages tall of page setup object of dataSheet to 1
  set print area of page setup object of summarySheet to "$A$1:$D$4"
  log "stage=save"
  if value of range "D20" of dataSheet is not ownedToken then error "Owned marker missing before save"
  save workbook as ownedBook filename {xlsx} file format Excel XML file format
  set ownedBook to workbook "native.xlsx"
  if full name of ownedBook is not {xlsx} then error "Saved path identity mismatch"
  close ownedBook saving no
  log "stage=reopen"
  set ownedBook to open workbook workbook file name {xlsx} with read only
  if full name of ownedBook is not {xlsx} then error "Reopen path identity mismatch"
  set dataSheet to worksheet "Data" of ownedBook
  set summarySheet to worksheet "Summary" of ownedBook
  if value of range "D20" of dataSheet is not ownedToken then error "Reopen owner marker missing"
  if formula of range "B6" of dataSheet is not "=SUM(B2:B4)" then error "Reopen formula mismatch"
  if value of range "B6" of dataSheet is not 60 then error "Reopen calculated value mismatch"
  if value of range "B1" of summarySheet is not 120 then error "Reopen cross sheet value mismatch"
  if merge cells of range "A8:D8" of dataSheet is not true then error "Reopen merged range mismatch"
  if bold of font object of range "A1:B1" of dataSheet is not true then error "Reopen font mismatch"
  if count of chart objects of dataSheet is not 1 then error "Reopen native chart missing"
  log "reopen_semantics=pass"
  log "stage=pdf"
  save workbook as ownedBook filename {pdf} file format PDF file format
  log "pdf_export=pass"
  if value of range "A1" of worksheet 1 of sentinelBook is not sentinelToken then error "Sentinel content changed"
  if saved of sentinelBook is not false then error "Sentinel was saved"
  if full name of sentinelBook is not sentinelPath then error "Sentinel path changed"
  repeat with originalState in previousBooks
   set originalBook to workbook (item 1 of originalState)
   if full name of originalBook is not item 2 of originalState then error "Unrelated path changed"
   if saved of originalBook is not item 3 of originalState then error "Unrelated saved state changed"
  end repeat
  log "unrelated_and_unsaved_sentinel=pass"
  if value of range "D20" of worksheet "Data" of ownedBook is not ownedToken then error "Owned marker missing before close"
  close ownedBook saving no
  log "owned_closed=pass"
  -- Sentinel intentionally remains open and unsaved as preservation evidence.
  return "native_probe=pass"
 end tell
end timeout
'''


def inspect_artifacts(output_dir):
    """Read native output without rewriting it or synthesizing missing content."""
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    with zipfile.ZipFile(output_dir / 'native.xlsx') as archive:
        wb = ET.fromstring(archive.read('xl/workbook.xml'))
        sheets = [e.attrib['name'] for e in wb.findall('s:sheets/s:sheet', ns)]
        data = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        summary = ET.fromstring(archive.read('xl/worksheets/sheet2.xml'))
        formula = data.find(".//s:c[@r='B6']/s:f", ns).text
        calculated = float(data.find(".//s:c[@r='B6']/s:v", ns).text)
        cross_formula = summary.find(".//s:c[@r='B1']/s:f", ns).text
        cross_value = float(summary.find(".//s:c[@r='B1']/s:v", ns).text)
        merges = [e.attrib['ref'] for e in data.findall('s:mergeCells/s:mergeCell', ns)]
        chart_names = [n for n in archive.namelist() if n.startswith('xl/charts/chart') and n.endswith('.xml')]
        chart_sources = []
        for name in chart_names:
            chart_sources.extend(e.text for e in ET.fromstring(archive.read(name)).findall('.//c:f', ns))
        assert sheets == ['Data', 'Summary'], sheets
        assert formula == 'SUM(B2:B4)' and calculated == 60
        assert cross_formula == 'Data!B6*2' and cross_value == 120
        assert 'A8:D8' in merges
        assert len(chart_names) == 1 and any('B$2:$B$4' in s for s in chart_sources), chart_sources
    import fitz
    pdf = fitz.open(output_dir / 'native.pdf')
    text = '\n'.join(page.get_text() for page in pdf)
    assert len(pdf) == 2, len(pdf)
    for expected in ('Month', 'Revenue', 'Revenue by month', 'Cross sheet total', '120'):
        assert expected in text, (expected, text)
    for index, page in enumerate(pdf):
        page.get_pixmap(matrix=fitz.Matrix(1.25, 1.25)).save(output_dir / f'page-{index + 1}.png')
    (output_dir / 'pdf-text.txt').write_text(text)
    return {'sheets': sheets, 'formula': formula, 'calculated_value': calculated,
            'cross_sheet_formula': cross_formula, 'cross_sheet_value': cross_value,
            'merged_ranges': merges, 'native_charts': chart_names,
            'chart_sources': chart_sources, 'pdf_pages': len(pdf)}


def run_probe(output_dir, timeout=90, execute=subprocess.run, native_staging_root=None):
    if type(timeout) is not int or not 1 <= timeout <= 900:
        raise ValueError('timeout must be an integer between 1 and 900 seconds')
    output_dir = Path(output_dir).absolute()
    output_dir.mkdir(parents=True, exist_ok=False)
    token = uuid.uuid4().hex[:12]
    source = Path(__file__).read_bytes()
    (output_dir / 'runner-source.py').write_bytes(source)
    native_dir = output_dir
    if native_staging_root is not None:
        native_dir = Path(tempfile.mkdtemp(prefix='wpscomposer-parity-' + token + '-',
                                         dir=Path(native_staging_root).absolute()))
    script = compile_probe(native_dir, token, max(1, timeout - 2))
    (output_dir / 'probe.applescript').write_text(script)
    report = {'component': 'spreadsheet', 'platform': 'darwin', 'engine': 'msoffice',
              'token': token, 'native_directory': str(native_dir), 'app_path': str(APP), 'source_sha256': hashlib.sha256(source).hexdigest(),
              'status': 'started', 'verified': False, 'recovery_required': True, 'timeout': timeout}
    if (APP / 'Contents/Info.plist').is_file():
        with (APP / 'Contents/Info.plist').open('rb') as stream:
            info = plistlib.load(stream)
        report['app_version'] = info.get('CFBundleShortVersionString')
        report['bundle_id'] = info.get('CFBundleIdentifier')
        report['dictionary_sha256'] = hashlib.sha256((APP / 'Contents/Resources/Excel.sdef').read_bytes()).hexdigest()
    started = time.monotonic()
    stdout, stderr = '', ''
    try:
        result = execute(['/usr/bin/osascript', str(output_dir / 'probe.applescript')],
                         capture_output=True, text=True, timeout=timeout)
        stdout, stderr = result.stdout, result.stderr
        report['returncode'] = result.returncode
        report['status'] = 'failed'
        if result.returncode == 0 and 'native_probe=pass' in stdout:
            if native_dir != output_dir:
                for name in ('native.xlsx', 'native.pdf'):
                    if (native_dir / name).is_file():
                        shutil.copyfile(native_dir / name, output_dir / name)
            report['artifacts'] = inspect_artifacts(output_dir)
            report.update(status='passed', verified=True, recovery_required=False)
    except subprocess.TimeoutExpired as error:
        stdout, stderr = error.stdout or '', error.stderr or ''
        report.update(status='timeout', error=str(error))
    except Exception as error:
        report.update(status='failed', error=f'{type(error).__name__}: {error}')
    finally:
        for filename, content in [('stdout.log', stdout), ('stderr.log', stderr)]:
            (output_dir / filename).write_text(content.decode(errors='replace') if isinstance(content, bytes) else content)
        log_text = stderr.decode(errors='replace') if isinstance(stderr, bytes) else stderr
        report['primitives'] = {line.split('=', 1)[0]: line.split('=', 1)[1]
                                for line in log_text.splitlines()
                                if '=pass' in line or '=failed:' in line}
        report['elapsed_seconds'] = round(time.monotonic() - started, 3)
        report['checksums'] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                               for p in output_dir.iterdir() if p.is_file()}
        (output_dir / 'report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', required=True, type=Path)
    parser.add_argument('--timeout', default=90, type=int)
    parser.add_argument('--native-staging-root', type=Path,
                        default=Path.home() / 'Library/Containers/com.microsoft.Excel/Data/Documents',
                        help='Existing Excel-owned container directory; a unique child is created.')
    args = parser.parse_args()
    result = run_probe(args.output_dir, timeout=args.timeout,
                       native_staging_root=args.native_staging_root)
    print(json.dumps(result, indent=2, ensure_ascii=False))
    raise SystemExit(0 if result['verified'] else 1)
