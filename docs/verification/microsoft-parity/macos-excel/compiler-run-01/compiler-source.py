"""Closed spreadsheet plans compiled to native Microsoft Excel AppleScript.

The caller owns deadlines, app serialization, sandbox staging, macro screening,
artifact validation and publication. Compilation neither starts Excel nor reads
source files. Native errors retain the owned workbook for caller recovery.
"""
from __future__ import annotations

import math
from pathlib import Path
import re
from typing import Mapping, Optional

from ..generation_plan import GenerationPlan, validate_generation_plan

SUCCESS_MARKER = 'WPSCOMPOSER_MS_OFFICE_OK:spreadsheet'
APP_PATH = '/Applications/Microsoft Excel.app'


def _quote(value: str) -> str:
    if not isinstance(value, str) or any(ord(c) < 32 and c not in '\n\r\t' for c in value):
        raise ValueError('Excel text contains unsupported control characters')
    value.encode('utf-16-le')
    value = value.replace('\\', '\\\\').replace('"', '\\"')
    for char, expression in (('\n', 'linefeed'), ('\r', 'return'), ('\t', 'tab')):
        value = value.replace(char, '" & ' + expression + ' & "')
    return '"' + value + '"'


def _path(value, suffixes) -> Path:
    path = Path(value)
    if not path.is_absolute() or '..' in path.parts or path.suffix.lower() not in suffixes:
        raise ValueError('Expected an absolute Excel artifact path with a supported extension')
    if any(ord(c) < 32 for c in str(path)):
        raise ValueError('Artifact path contains control characters')
    _quote(str(path))
    return path


def _seconds(timeout) -> int:
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or not 0 < timeout <= 900:
        raise ValueError('timeout must be finite and between 0 and 900 seconds')
    return max(1, math.ceil(timeout))


def _column_number(column: str) -> int:
    value = 0
    for character in column:
        value = value * 26 + ord(character) - ord('A') + 1
    return value


def _column_name(number: int) -> str:
    result = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _column_range(value: str) -> str:
    value = value.upper()
    if not re.fullmatch(r'[A-Z]{1,3}(?::[A-Z]{1,3})?', value):
        raise ValueError('Column must be an Excel letter or whole-column range')
    ends = value.split(':')
    if any(_column_number(c) > 16384 for c in ends) or _column_number(ends[0]) > _column_number(ends[-1]):
        raise ValueError('Column range lies outside Excel bounds or is reversed')
    return ':'.join((ends[0], ends[-1]))


def _sheet_name(value: str) -> str:
    if not value or len(value.encode('utf-16-le')) // 2 > 31 or re.search(r'[\[\]:*?/\\]', value) or value.startswith("'") or value.endswith("'"):
        raise ValueError('Invalid Excel worksheet name')
    return _quote(value)


def _cell(value) -> str:
    if value is None:
        return '""'
    if isinstance(value, str):
        if len(value.encode('utf-16-le')) // 2 > 32767:
            raise ValueError('Excel cell text exceeds 32767 UTF-16 code units')
        return _quote(value)
    if isinstance(value, bool):
        return 'true' if value else 'false'
    return str(value)


def _rgb(value: str) -> str:
    return '{' + ', '.join(str(int(value[i:i + 2], 16)) for i in (1, 3, 5)) + '}'


def _snapshot() -> list[str]:
    return [
        'set beforeBooks to {}',
        'set beforeNames to name of every workbook',
        'repeat with beforeName in beforeNames',
        ' set priorBook to workbook (contents of beforeName)',
        ' set end of beforeBooks to {name of priorBook, full name of priorBook, saved of priorBook}',
        'end repeat',
    ]


def _finish() -> list[str]:
    return [
        'close ownedBook saving no',
        'set ownedBook to missing value',
        'if (count of workbooks) is not (count of beforeBooks) then error "Workbook count changed"',
        'repeat with beforeState in beforeBooks',
        ' set priorBook to workbook (item 1 of beforeState)',
        ' if full name of priorBook is not item 2 of beforeState then error "Unrelated workbook path changed"',
        ' if saved of priorBook is not item 3 of beforeState then error "Unrelated workbook saved state changed"',
        'end repeat',
    ]


def _wrap(lines: list[str], timeout) -> str:
    seconds = _seconds(timeout)
    return '\n'.join([f'with timeout of {seconds} seconds', f' tell application {_quote(APP_PATH)}'] +
                     ['  ' + line for line in lines] +
                     [' end tell', 'end timeout', f'return {_quote(SUCCESS_MARKER)}', ''])


def compile_plan(plan: GenerationPlan, resources: Mapping[str, Path], native_path: Path,
                 pdf_path: Optional[Path] = None, timeout=60) -> str:
    """Compile the existing closed spreadsheet grammar before native mutation."""
    _seconds(timeout)
    native_path = _path(native_path, {'.xlsx'})
    if pdf_path is not None:
        pdf_path = _path(pdf_path, {'.pdf'})
    if resources:
        raise ValueError('The current spreadsheet plan grammar has no external resources')
    validated = validate_generation_plan(plan.to_dict(), 'spreadsheet')
    names = ['Sheet1']
    selected = 1
    body = []
    for position, operation in enumerate(validated.operations):
        args = operation.args
        if operation.op == 'sheet.reset':
            if position != 0:
                raise ValueError('sheet.reset must occur only at the start of a plan')
        elif operation.op == 'sheet.rename':
            index, name = args['index'], args['name']
            quoted = _sheet_name(name)
            if index > len(names) or any(n.casefold() == name.casefold() for i, n in enumerate(names, 1) if i != index):
                raise ValueError('Worksheet index is invalid or its name is duplicated')
            names[index - 1] = name
            body += [f'set name of worksheet {index} of ownedBook to {quoted}',
                     f'set ownedSheet to worksheet {selected} of ownedBook']
        elif operation.op == 'sheet.add':
            name = args['name']
            quoted = _sheet_name(name)
            if any(n.casefold() == name.casefold() for n in names):
                raise ValueError('Duplicate worksheet name')
            names.append(name)
            selected = len(names)
            body += ['tell ownedBook', f' make new worksheet at end with properties {{name:{quoted}}}',
                     'end tell', f'set ownedSheet to worksheet {selected} of ownedBook']
        elif operation.op == 'sheet.select':
            selected = args['index']
            if selected > len(names):
                raise ValueError('Worksheet index is outside the generated workbook')
            body += [f'set ownedSheet to worksheet {selected} of ownedBook']
        elif operation.op == 'sheet.write_table':
            values = args['values']
            if not values or not values[0]:
                continue
            top, left = args['startRow'], args['startCol']
            bottom, right = top + len(values) - 1, left + len(values[0]) - 1
            if bottom > 1048576 or right > 16384:
                raise ValueError('Table extends beyond Excel row or column limits')
            rectangle = f'{_column_name(left)}{top}:{_column_name(right)}{bottom}'
            header = f'{_column_name(left)}{top}:{_column_name(right)}{top}'
            rows = '{' + ', '.join('{' + ', '.join(_cell(v) for v in row) + '}' for row in values) + '}'
            font_size = args.get('fontSize', 11)
            if font_size > 409:
                raise ValueError('Excel font size exceeds 409 points')
            body += [f'set value of range {_quote(rectangle)} of ownedSheet to {rows}',
                     f'set font size of font object of range {_quote(rectangle)} of ownedSheet to {font_size}']
            if args.get('headerBold', True):
                body += [f'set bold of font object of range {_quote(header)} of ownedSheet to true']
            body += [f'set color of font object of range {_quote(header)} of ownedSheet to {_rgb(args.get("headerFontColor", "#FFFFFF"))}']
            shade = args.get('headerShade', '#4472C4')
            if shade is not False:
                body += [f'set color of interior object of range {_quote(header)} of ownedSheet to {_rgb(shade)}']
        elif operation.op == 'sheet.set_column_width':
            column = _column_range(args['column'])
            if args['width'] > 255:
                raise ValueError('Excel column width exceeds 255')
            body += [f'set column width of range {_quote(column)} of ownedSheet to {args["width"]}']
        elif operation.op == 'sheet.autofit':
            body += ['autofit entire column of used range of ownedSheet']
        else:
            raise ValueError('Unsupported Microsoft Excel operation: ' + operation.op)
    lines = _snapshot() + [
        f'if {_quote(native_path.name)} is in beforeNames then error "Target workbook name is already open"',
        'set ownedBook to make new workbook',
        'set ownedName to name of ownedBook',
        'if ownedName is in beforeNames then error "New workbook identity collision"',
        'set ownedBook to workbook ownedName',
        'repeat while (count of worksheets of ownedBook) > 1',
        ' delete worksheet 2 of ownedBook',
        'end repeat',
        'set name of worksheet 1 of ownedBook to "Sheet1"',
        'set ownedSheet to worksheet 1 of ownedBook',
    ] + body + [
        'repeat with sheetIndex from 1 to (count of worksheets of ownedBook)',
        ' calculate worksheet sheetIndex of ownedBook',
        'end repeat',
        f'save workbook as ownedBook filename {_quote(str(native_path))} file format Excel XML file format',
        f'set ownedBook to workbook {_quote(native_path.name)}',
        f'if full name of ownedBook is not {_quote(str(native_path))} then error "Owned saved path mismatch"',
    ]
    if pdf_path is not None:
        lines += [f'save workbook as ownedBook filename {_quote(str(pdf_path))} file format PDF file format',
                  f'if full name of ownedBook is not {_quote(str(native_path))} then error "PDF export changed workbook identity"']
    return _wrap(lines + _finish(), timeout)


def compile_conversion(source_path: Path, pdf_path: Path, timeout=60) -> str:
    """Compile read-only native PDF conversion; caller must screen source macros."""
    source_path = _path(source_path, {'.xlsx', '.xls'})
    pdf_path = _path(pdf_path, {'.pdf'})
    lines = _snapshot() + [
        f'if {_quote(source_path.name)} is in beforeNames then error "Source workbook is already open"',
        f'set ownedBook to open workbook workbook file name {_quote(str(source_path))} with read only and editable',
        f'if full name of ownedBook is not {_quote(str(source_path))} then error "Owned source path mismatch"',
        f'save workbook as ownedBook filename {_quote(str(pdf_path))} file format PDF file format',
        f'if full name of ownedBook is not {_quote(str(source_path))} then error "PDF export changed source identity"',
    ]
    return _wrap(lines + _finish(), timeout)
