"""Closed presentation plans compiled into native macOS PowerPoint operations.

The caller stages macro-free files/resources in PowerPoint's own container and
owns deadlines, locking, validation and publication. No application-wide cleanup.
"""
from __future__ import annotations

import math
from pathlib import Path

from ..generation_plan import validate_generation_plan
from .macos_script import apple_string

SUCCESS_MARKER = 'WPSCOMPOSER_MS_OFFICE_OK:presentation'


class MacPowerPointCapabilityError(ValueError):
    """A requested PowerPoint primitive is not yet natively verified."""


def _dimension(value):
    if (isinstance(value, bool) or not isinstance(value, (int, float))
            or not math.isfinite(value) or value <= 0):
        raise ValueError('PowerPoint dimensions must be finite positive numbers')
    return format(value, '.15g')


def compile_initial_slide_size(width, height):
    """Compile the verified arbitrary-size sequence for an empty presentation."""
    width_text, height_text = _dimension(width), _dimension(height)
    desired = 'horizontal orientation' if width >= height else 'vertical orientation'
    opposite = 'vertical orientation' if width >= height else 'horizontal orientation'
    return '\n'.join([
        'if (count of slides of ownedDoc) is not 0 then return "EXISTING_SLIDES"',
        f'set slide orientation of page setup of ownedDoc to {opposite}',
        f'set slide width of page setup of ownedDoc to {height_text}',
        f'set slide orientation of page setup of ownedDoc to {desired}',
        f'set slide width of page setup of ownedDoc to {width_text}',
    ])


def _path(value):
    path = Path(value)
    if not path.is_absolute():
        raise ValueError('Native PowerPoint paths must be absolute')
    apple_string(str(path))
    return path


def _timeout(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise ValueError('timeout must be finite and positive')
    return max(1, math.ceil(value))


def _rgb(value):
    if isinstance(value, int):
        colors = (value & 255, (value >> 8) & 255, (value >> 16) & 255)
    else:
        colors = tuple(int(value[i:i+2], 16) for i in (1, 3, 5))
    return '{' + ', '.join(str(c) for c in colors) + '}'


def _font(target, size, color, family=None, bold=False):
    lines = [f'set font size of font of {target} to {size}',
             f'set font color of font of {target} to {_rgb(color)}',
             f'set bold of font of {target} to {str(bold).lower()}']
    if family:
        lines += [f'set font name of font of {target} to {apple_string(family)}',
                  f'set east asian name of font of {target} to {apple_string(family)}']
    return lines


def _snapshot():
    return '''set beforeDocs to {}
repeat with snapshotIndex from 1 to (count of presentations)
 set end of beforeDocs to {name of presentation snapshotIndex, full name of presentation snapshotIndex, saved of presentation snapshotIndex, count of slides of presentation snapshotIndex}
end repeat'''


def _preserve():
    return '''if (count of presentations) is not (count of beforeDocs) then error "Unrelated presentation count changed"
repeat with snapshotIndex from 1 to (count of beforeDocs)
 set afterRow to {name of presentation snapshotIndex, full name of presentation snapshotIndex, saved of presentation snapshotIndex, count of slides of presentation snapshotIndex}
 if afterRow is not item snapshotIndex of beforeDocs then
  log item snapshotIndex of beforeDocs
  log afterRow
  error "Unrelated presentation state changed"
 end if
end repeat'''


def _wrapper(body, native, pdf, timeout, source=None):
    declarations = f'''set nativePath to {apple_string(str(native))}
set nativeHFS to (POSIX file nativePath) as text
set ownedDoc to missing value
'''
    if pdf:
        declarations += f'set pdfPath to {apple_string(str(pdf))}\nset pdfHFS to (POSIX file pdfPath) as text\n'
    collision_name = source.name if source else native.name
    before = _snapshot() + f'''
repeat with prior in beforeDocs
 if item 1 of prior is {apple_string(collision_name)} then error "Existing presentation collision"
end repeat'''
    if source:
        declarations += f'set sourcePath to {apple_string(str(source))}\nset sourceFile to (POSIX file sourcePath) as alias\n'
        acquire = f'''open sourceFile
set openDeadline to (current date) + {timeout}
repeat until exists presentation {apple_string(source.name)}
 if (current date) > openDeadline then error "PowerPoint open deadline exceeded"
 delay 0.1
end repeat
set ownedDoc to presentation {apple_string(source.name)}
if full name of ownedDoc is not sourcePath then error "Owned source path mismatch"'''
    else:
        acquire = '''set ownedDoc to make new presentation
set createdName to name of ownedDoc
repeat with prior in beforeDocs
 if item 1 of prior is createdName then error "New presentation collision"
end repeat'''
    save = '' if source else f'''save ownedDoc in nativeHFS as save as Open XML presentation
set ownedDoc to presentation {apple_string(native.name)}
if full name of ownedDoc is not nativePath then error "Owned output path mismatch"
if saved of ownedDoc is false then error "Owned presentation was not saved"
'''
    if pdf:
        save += 'save ownedDoc in pdfHFS as save as PDF\n'
    # Failed operations deliberately retain their owned document. Runtime quarantines
    # the staging directory; never guess which document should be closed after errors.
    return declarations + f'''with timeout of {timeout} seconds
 tell application "Microsoft PowerPoint"
  {before}
  try
   {acquire}
   {body}
   {save}
   close ownedDoc saving no
   set ownedDoc to missing value
   {_preserve()}
   return "{SUCCESS_MARKER}"
  on error errorText number errorNumber
   log "WPSCOMPOSER_MS_OFFICE_ERROR:presentation" & tab & errorNumber & tab & errorText
   error errorText number errorNumber
  end try
 end tell
end timeout
'''


def compile_plan(plan, resources, native_path, pdf_path=None, timeout=60):
    """Return AppleScript without starting Office or requiring staged files to exist."""
    seconds = _timeout(timeout)
    native = _path(native_path)
    pdf = _path(pdf_path) if pdf_path is not None else None
    if native.suffix.lower() != '.pptx' or (pdf and pdf.suffix.lower() != '.pdf'):
        raise ValueError('PowerPoint generation requires .pptx and optional .pdf')
    validated = validate_generation_plan(plan.to_dict() if hasattr(plan, 'to_dict') else plan, 'presentation')
    bound = {key: _path(value) for key, value in resources.items()}
    lines = []
    count = 0
    preset = {}
    width = 960
    for operation in validated.operations:
        op, a = operation.op, operation.args
        if op == 'slide.reset':
            lines.append('repeat while (count of slides of ownedDoc) > 0\n delete slide 1 of ownedDoc\nend repeat')
            count = 0
        elif op == 'slide.set_size':
            width = a['width']
            if a['height'] == 540:
                lines += ['set slide size of page setup of ownedDoc to slide size on screen',
                          f'set slide width of page setup of ownedDoc to {width}']
            else:
                if count != 0:
                    raise MacPowerPointCapabilityError(
                        'Arbitrary PowerPoint slide size must be set before adding slides')
                lines.append(compile_initial_slide_size(width, a['height']))
        elif op == 'slide.apply_preset':
            preset = a['preset']
            lines.append(f'set fore color of fill format of background of slide master of ownedDoc to {_rgb(preset["colors"]["background"])}')
        elif op in {'slide.add_title', 'slide.add_section', 'slide.add_bullets', 'slide.add_blank'}:
            count += 1
            layouts = {'slide.add_title':'slide layout title slide', 'slide.add_section':'slide layout title only',
                       'slide.add_bullets':'slide layout text slide', 'slide.add_blank':'slide layout blank'}
            lines.append(f'set currentSlide to make new slide at end of ownedDoc with properties {{layout:{layouts[op]}}}')
            colors = preset.get('colors', {})
            fonts = preset.get('fonts', {})
            if colors.get('background'):
                lines += ['set follow master background of currentSlide to false',
                          f'set fore color of fill format of background of currentSlide to {_rgb(colors["background"])}']
            if op != 'slide.add_blank':
                role = fonts.get('title', {})
                target = 'text range of text frame of shape 1 of currentSlide'
                lines.append(f'set content of {target} to {apple_string(a["title"])}')
                title_size = a.get('titleSize', 40 if op == 'slide.add_title' else 32)
                if title_size in (40, 32) and role.get('size'):
                    title_size = role['size']
                lines += _font(target, title_size, a.get('titleColor') or role.get('color') or colors.get('text', '#1A1A1A'), role.get('family'))
            if op == 'slide.add_title' and a.get('subtitle'):
                role = fonts.get('subtitle', {})
                target = 'text range of text frame of shape 2 of currentSlide'
                lines.append(f'set content of {target} to {apple_string(a["subtitle"])}')
                lines += _font(target, (role.get('size', 20) if a.get('subtitleSize', 20) == 20 else a['subtitleSize']), role.get('color', '#666666'), role.get('family'))
            if op == 'slide.add_bullets':
                role = fonts.get('body', {})
                target = 'text range of text frame of shape 2 of currentSlide'
                lines.append(f'set content of {target} to {apple_string(chr(13).join(a["items"]))}')
                lines += _font(target, (role.get('size', 18) if a.get('bodySize', 18) == 18 else a['bodySize']), role.get('color', '#1A1A1A'), role.get('family'))
        elif op == 'slide.add_image':
            if a['slide'] > count:
                raise ValueError('Image refers to a missing or reset slide')
            if a['imageId'] not in bound:
                raise ValueError('Unbound PowerPoint image resource: ' + a['imageId'])
            lines += [f'set targetSlide to slide {a["slide"]} of ownedDoc',
                      f'set currentPicture to make new picture at end of targetSlide with properties {{file name:{apple_string(str(bound[a["imageId"]]))}, link to file:false, save with document:true, left position:{a["left"]}, top:{a["top"]}}}']
            if 'width' in a or 'height' in a:
                lines.append('set lock aspect ratio of currentPicture to ' + ('false' if 'width' in a and 'height' in a else 'true'))
                for dimension in ('width', 'height'):
                    if dimension in a:
                        lines.append(f'set {dimension} of currentPicture to {a[dimension]}')
        elif op == 'slide.add_table':
            if a['slide'] > count:
                raise ValueError('Table refers to a missing or reset slide')
            lines += [f'set targetSlide to slide {a["slide"]} of ownedDoc',
                      f'set currentTable to make new shape table at end of targetSlide with properties {{number of rows:{a["rows"]}, number of columns:{a["cols"]}, left position:{a["left"]}, top:{a["top"]}, width:{a["width"]}, height:{a["height"]}}}']
            for r, row in enumerate(a['data'], 1):
                for col, value in enumerate(row, 1):
                    if r > a['rows'] or col > a['cols']:
                        raise ValueError('Table data exceeds native table bounds')
                    lines += [f'set currentCell to get cell from table object of currentTable row {r} column {col}',
                              f'set content of text range of text frame of shape of currentCell to {apple_string(str(value))}']
                    lines += _font('text range of text frame of shape of currentCell', a.get('fontSize',11), a.get('headerFont','#FFFFFF') if r == 1 else '#000000', bold=r == 1)
                    if r == 1:
                        lines.append(f'set fore color of fill format of shape of currentCell to {_rgb(a.get("headerShade", "#4472C4"))}')
        else:
            raise MacPowerPointCapabilityError('Unsupported native PowerPoint operation: ' + op)
    if count == 0:
        raise ValueError('A PowerPoint plan must contain at least one slide')
    return _wrapper('\n'.join(lines), native, pdf, seconds)


def compile_conversion(source_path, pdf_path, timeout=60):
    """Open an owned macro-free staging copy through native AppleScript."""
    seconds = _timeout(timeout)
    source, pdf = _path(source_path), _path(pdf_path)
    if source.suffix.lower() not in {'.pptx', '.ppsx', '.potx', '.odp'}:
        raise MacPowerPointCapabilityError('Native conversion of macro-capable or unverified PowerPoint formats is not enabled')
    if pdf.suffix.lower() != '.pdf' or source == pdf:
        raise ValueError('Conversion requires a distinct PDF output')
    return _wrapper('', source, pdf, seconds, source=source)
