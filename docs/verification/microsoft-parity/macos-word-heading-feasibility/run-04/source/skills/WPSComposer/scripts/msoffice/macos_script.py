"""Closed GenerationPlan -> native Microsoft Word AppleScript.

This compiler emits object-model operations only. Document text is quoted data;
unsupported semantics are rejected before a Word document can be created.
"""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
import math
from pathlib import Path
import re
from typing import Mapping

from ..generation_plan import GenerationPlan, validate_generation_plan
from ..longform.executor import ExecutionOutcome, ExecutionIssue, PaginationMap, PaginationNode, PaginationFragment
from .errors import NativeWordCapabilityError


class MacWordCapabilityError(NativeWordCapabilityError):
    """The native Mac engine cannot faithfully execute a requested operation."""


def apple_string(value: str) -> str:
    if not isinstance(value, str) or any(ord(c) < 32 and c not in '\n\r\t' for c in value):
        raise ValueError('AppleScript text contains unsupported control characters')
    value = value.replace('\\', '\\\\').replace('"', '\\"')
    for char, expression in (('\n', 'linefeed'), ('\r', 'return'), ('\t', 'tab')):
        value = value.replace(char, '" & ' + expression + ' & "')
    return '"' + value + '"'


def _num(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError('Expected finite Word measurement')
    return str(value)


def _bool(value):
    return 'true' if value else 'false'


def _style(name, declared=()):
    if name == 'Body Text':
        return 'style body text'
    if name == 'Title':
        return 'style title'
    if re.fullmatch(r'Heading [1-6]', name):
        return 'style heading' + name[-1]
    if name in declared:
        return apple_string(name)
    raise MacWordCapabilityError('Unsupported native Word style: ' + str(name))


@dataclass(frozen=True)
class CompiledScript:
    source: str
    nodes: Mapping[str, tuple[str, str]]
    operations: int
    issues: tuple[ExecutionIssue, ...] = ()


_APPEND = '''on appendText(ownedDoc, valueText)
 tell application "Microsoft Word"
  set p to (end of content of text object of ownedDoc) - 1
  set r to create range ownedDoc start p end p
  set content of r to valueText
  set e to (end of content of text object of ownedDoc) - 1
  return create range ownedDoc start p end e
 end tell
end appendText
'''


def wrap_owned(body: str, target: Path, timeout: float, *, source: Path | None = None) -> str:
    """Acquire exact ownership, close only that document and verify sentinels."""
    seconds = max(1, int(timeout))
    target_expr = apple_string(str(target))
    create = f'''set ownedDoc to make new document
set ownedName to name of ownedDoc
repeat with prior in beforeDocs
 if item 1 of prior is ownedName then
  set ownedDoc to missing value
  error "New document identity collision"
 end if
end repeat
set ownedDoc to document ownedName
save as ownedDoc file name {target_expr} file format format document default add to recent files false
set ownedDoc to document {apple_string(target.name)}
if posix full name of ownedDoc is not {target_expr} then error "Owned document path mismatch"'''
    if source is not None:
        create = f'''open file name {apple_string(str(source))} add to recent files false
set ownedDoc to document {apple_string(source.name)}
if posix full name of ownedDoc is not {apple_string(str(source))} then error "Owned source path mismatch"'''
    return _APPEND + f'''set ownedDoc to missing value
set beforeDocs to {{}}
set failureText to ""
with timeout of {seconds} seconds
 tell application "Microsoft Word"
  repeat with documentIndex from 1 to (count of documents)
   set d to document documentIndex
   set end of beforeDocs to {{name of d, posix full name of d, saved of d, content of text object of d}}
  end repeat
  try
   {create}
   {body}
  on error errText number errNumber
   log "WPSC_ERROR" & tab & errNumber & tab & errText
   set failureText to "Native Word operation failed (" & errNumber & "): " & errText
  end try
  try
   if ownedDoc is not missing value then
    set closePath to posix full name of ownedDoc
    set closeName to name of ownedDoc
    if closePath is not {target_expr} and closePath is not {apple_string(str(source)) if source else target_expr} then error "Cleanup identity is outside this operation"
    close document closeName saving no
   end if
   set ownedDoc to missing value
  on error
   error "Native Word owned-document cleanup failed; staging quarantined"
  end try
  if (count of documents) is not (count of beforeDocs) then error "Native Word document count changed; staging quarantined"
  repeat with prior in beforeDocs
   set d to document (item 1 of prior)
   if posix full name of d is not item 2 of prior then error "Preexisting Word path changed"
   if saved of d is not item 3 of prior then error "Preexisting Word saved state changed"
   if content of text object of d is not item 4 of prior then error "Preexisting Word text changed"
  end repeat
  log "WPSC_CLEAN"
  if failureText is not "" then error failureText
 end tell
end timeout
return "WPSC_OK" & tab & (count of beforeDocs)
'''


def pagination_source(nodes):
    lines = ['repaginate ownedDoc']
    for node, (bookmark, _) in nodes.items():
        lines += [
            f'set r to text object of bookmark {apple_string(bookmark)} of ownedDoc',
            'set s to start of content of r',
            'set e to end of content of r',
            'set sr to create range ownedDoc start s end s',
            'set ep to e',
            'if ep > s then set ep to ep - 1',
            'set er to create range ownedDoc start ep end ep',
            'set sp to (get range information sr information type active end page number) as integer',
            'set pe to (get range information er information type active end page number) as integer',
            f'log "WPSC_NODE" & tab & {apple_string(node)} & tab & sp & tab & pe & tab & s & tab & e',
        ]
    return '\n'.join(lines)


def refresh_source(rounds=3):
    # Word's generic field update can succeed with a stale TOC entry cache.
    # Rebuild native index objects each round before observing field results.
    return f'''set priorFields to missing value
set fieldsStable to false
repeat with refreshRound from 1 to {max(2, rounds)}
 repaginate ownedDoc
 repeat with tocIndex from 1 to (count of tables of contents of ownedDoc)
  update (table of contents tocIndex of ownedDoc)
 end repeat
 repeat with figuresIndex from 1 to (count of tables of figures of ownedDoc)
  update (table of figures figuresIndex of ownedDoc)
 end repeat
 repeat with nativeIndex from 1 to (count of indexes of ownedDoc)
  update (index nativeIndex of ownedDoc)
 end repeat
 repeat with fieldIndex from 1 to (count of fields of ownedDoc)
  if (update field (field fieldIndex of ownedDoc)) is false then error "Native Word field update failed"
 end repeat
 repaginate ownedDoc
 set fieldState to {{compute statistics ownedDoc statistic statistic pages}}
 repeat with fieldIndex from 1 to (count of fields of ownedDoc)
  set ownField to field fieldIndex of ownedDoc
  set end of fieldState to content of result range of ownField
 end repeat
 if fieldState is priorFields then
  set fieldsStable to true
  exit repeat
 end if
 set priorFields to fieldState
end repeat
if fieldsStable is false then error "Native Word fields did not converge"
'''


def _rgb(value):
    if not isinstance(value, str) or not re.fullmatch(r'#[0-9a-fA-F]{6}', value):
        raise MacWordCapabilityError('Mac Word requires a hexadecimal color')
    return '{' + ', '.join(str(int(value[i:i+2], 16) * 257) for i in (1, 3, 5)) + '}'


def _format(target, args):
    lines = []
    font_keys = {'fontName': 'east asian name', 'fontNameAscii': 'name', 'fontSize': 'font size', 'size': 'font size', 'bold': 'bold', 'italic': 'italic', 'strikethrough': 'strike through'}
    paragraph_keys = {'indentFirst': 'first line indent', 'spaceBefore': 'space before', 'spaceAfter': 'space after', 'keepWithNext': 'keep with next', 'keepTogether': 'keep together', 'leftIndent': 'paragraph format left indent', 'rightIndent': 'paragraph format right indent'}
    for key, prop in font_keys.items():
        if key in args:
            value = args[key]
            expression = apple_string(value) if isinstance(value, str) else _bool(value) if isinstance(value, bool) else _num(value)
            lines.append(f'set {prop} of font object of {target} to {expression}')
    for key, prop in paragraph_keys.items():
        if key in args:
            value = args[key]
            lines.append(f'set {prop} of paragraph format of {target} to {_bool(value) if isinstance(value, bool) else _num(value)}')
    if 'align' in args:
        if args['align'] not in (0,1,2,3):
            raise MacWordCapabilityError('Mac Word unsupported alignment')
        align = ('left', 'center', 'right', 'justify')[int(args['align'])]
        lines.append(f'set alignment of paragraph format of {target} to align paragraph {align}')
    if 'outlineLevel' in args:
        level = int(args["outlineLevel"])
        value = "outline level body text" if level == 10 else f"outline level{level}"
        lines.append(f"set outline level of paragraph format of {target} to {value}")
    if 'color' in args:
        lines.append(f'set color of font object of {target} to {_rgb(args["color"])}')
    if 'underline' in args:
        lines.append(f'set underline of font object of {target} to underline {"single" if args["underline"] else "none"}')
    if 'shading' in args:
        lines.append(f'set background pattern color of shading of {target} to {_rgb(args["shading"])}')
    if 'leftBorder' in args:
        lines += [f'set ownBorder to get border (paragraph format of {target}) which border border left', f'set line style of ownBorder to line style {"single" if args["leftBorder"] else "none"}']
        if args.get('leftBorder'):
            lines.append('set line width of ownBorder to line width150 point')
    if 'borderColor' in args:
        if not args.get('leftBorder'):
            raise MacWordCapabilityError('Mac Word borderColor requires leftBorder')
        lines.append(f'set color of ownBorder to {_rgb(args["borderColor"])}')
    if 'lineSpacing' in args or 'lineSpacingRule' in args:
        value = args.get('lineSpacing', 1)
        rule = args.get('lineSpacingRule')
        rules = {'single':'line space single', 'one_and_half':'line space1 pt5', 'double':'line space double', 'exact':'line space exactly', 'exactly':'line space exactly', 'at_least':'line space at least', 'multiple':'line space multiple'}
        if rule is None:
            rule = 'one_and_half' if value == 1.5 else 'single' if value == 1 else 'multiple'
        if rule not in rules:
            raise MacWordCapabilityError('Mac native Word unsupported lineSpacingRule')
        lines.append(f'set line spacing rule of paragraph format of {target} to {rules[rule]}')
        if rule in ('exact', 'exactly', 'at_least', 'multiple'):
            lines.append(f'set line spacing of paragraph format of {target} to {_num(value * 12 if rule == "multiple" else value)}')
    return lines


def _caption(args, *, before):
    descriptor = args['numbering']
    lines = [f'set r to my appendText(ownedDoc, {apple_string(descriptor["prefix"])})', 'set captionStart to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start captionStart end captionStart']
    seq_code = descriptor['sequenceId'] + ' ' + chr(92) + '* ARABIC'
    if descriptor['mode'] == 'chapter':
        style_code = str(descriptor['chapterStyleLevel']) + ' ' + chr(92) + 's'
        lines += [f'create new field text range r field type field style ref field text {apple_string(style_code)} preserve formatting true', 'set r to my appendText(ownedDoc, "-")', 'set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p']
        seq_code += ' ' + chr(92) + 's ' + str(descriptor['resetLevel'])
    lines += [f'create new field text range r field type field sequence field text {apple_string(seq_code)} preserve formatting true', 'set captionEnd to (end of content of text object of ownedDoc) - 1']
    if args.get('bookmarkName'):
        lines += ['set r to create range ownedDoc start captionStart end captionEnd', f'make new bookmark at ownedDoc with properties {{name:{apple_string(args["bookmarkName"])}, text object:r}}']
    text = descriptor['suffix'] + ' ' + args['caption']
    lines += [f'set r to my appendText(ownedDoc, {apple_string(text)} & return)', 'set font size of font object of r to 10', 'set first line indent of paragraph format of r to 0', 'set alignment of paragraph format of r to align paragraph center', f'set keep with next of paragraph format of r to {_bool(before and args.get("keepCaptionWithFirstRow"))}']
    return lines


def _equation(operation, resources, issues):
    """Native formula shell keeps math, number, and body in distinct ranges."""
    from ..longform.native_math import convert_restricted_latex, NativeMathConversionError
    a = operation.args
    content = a.get('content', {})
    degradation = content.get('plannedDegradation')
    fallback = a.get('fallbackResource')
    if fallback:
        # Explicit image recovery needs a separately verified native path.
        raise MacWordCapabilityError('Mac native Word unsupported equation fallback resource')
    if degradation:
        linear = a['fallbackText']
        issues.append(ExecutionIssue(degradation['code'], degradation['reason'], node_id=operation.node_id))
    else:
        try:
            linear = content['nativeMath']['linearText'] if content else convert_restricted_latex(a['source']).linear_text
        except NativeMathConversionError:
            raise MacWordCapabilityError('Mac native Word unsupported equation source') from None
    lines = ['set p to (end of content of text object of ownedDoc) - 1',
             'set r to create range ownedDoc start p end p',
             'set formulaTable to make new table at ownedDoc with properties {text object:r, number of rows:1, number of columns:3}',
             'set allow break across pages of row 1 of formulaTable to false',
             'set formulaWidth to (width of column 1 of formulaTable) + (width of column 2 of formulaTable) + (width of column 3 of formulaTable)',
             'if formulaWidth <= 72 then error "Native formula container is too narrow"',
             'set width of column 1 of formulaTable to 36',
             'set width of column 2 of formulaTable to formulaWidth - 72',
             'set width of column 3 of formulaTable to 36']
    for border in ('top','bottom','left','right','horizontal','vertical'):
        lines += [f'set ownBorder to get border formulaTable which border border {border}', 'set line style of ownBorder to line style none']
    lines += ['set formulaCell to get cell from table formulaTable row 1 column 2',
              f'set content of text object of formulaCell to {apple_string(linear)}',
              'set formulaFailed to false']
    if not degradation:
        recover = operation.failure_policy.get('mode') == 'degrade'
        if recover:
            lines.append('try')
        lines += ['set s to start of content of text object of formulaCell',
                  'set e to (end of content of text object of formulaCell) - 1',
                  'set r to create range ownedDoc start s end e',
                  'set nativeMathCount to count of math objects of ownedDoc',
                  'create new equation from range r in document ownedDoc',
                  'if (count of math objects of ownedDoc) is not nativeMathCount + 1 then error "Native equation creation failed"',
                  'set ownMath to math object (nativeMathCount + 1) of ownedDoc',
                  'build up ownMath']
        if recover:
            lines += ['on error',
                      'set s to start of content of text object of formulaCell',
                      'set e to (end of content of text object of formulaCell) - 1',
                      'set failedMathRange to create range ownedDoc start s end e',
                      'delete failedMathRange',
                      f'set content of text object of formulaCell to {apple_string(a["fallbackText"])}',
                      'set formulaFailed to true',
                      f'log "WPSC_EQUATION_FALLBACK" & tab & {apple_string(operation.node_id)}',
                      'end try']
    lines += ['set alignment of paragraph format of text object of formulaCell to align paragraph center',
              'set formulaNumberCell to get cell from table formulaTable row 1 column 3']
    number = a['numbering']
    lines += [f'set content of text object of formulaNumberCell to {apple_string(number["prefix"] + number["suffix"])}', f'set numberStart to (start of content of text object of formulaNumberCell) + {len(number["prefix"])}', 'set r to create range ownedDoc start numberStart end numberStart']
    seq_code = number['sequenceId'] + ' ' + chr(92) + '* ARABIC'
    if number['mode'] == 'chapter':
        style_code = str(number['chapterStyleLevel']) + ' ' + chr(92) + 's'
        lines += [f'create new field text range r field type field style ref field text {apple_string(style_code)} preserve formatting true', f'set p to (end of content of text object of formulaNumberCell) - {len(number["suffix"]) + 1}', 'set r to create range ownedDoc start p end p', 'set content of r to "-"', 'set r to create range ownedDoc start (p + 1) end (p + 1)']
        seq_code += ' ' + chr(92) + 's ' + str(number['resetLevel'])
    lines += [f'create new field text range r field type field sequence field text {apple_string(seq_code)} preserve formatting true', f'set numberEnd to (end of content of text object of formulaNumberCell) - {len(number["suffix"]) + 1}', 'set r to create range ownedDoc start numberStart end numberEnd', f'make new bookmark at ownedDoc with properties {{name:{apple_string(a["bookmarkName"])}, text object:r}}', 'set alignment of paragraph format of text object of formulaNumberCell to align paragraph right', 'set first line indent of paragraph format of text object of formulaTable to 0', 'set keep together of paragraph format of text object of formulaTable to true', 'set r to my appendText(ownedDoc, return)', 'set style of r to style body text']
    if degradation:
        notice = '[' + degradation['code'] + '] ' + degradation['reason']
        lines.append(f'set r to my appendText(ownedDoc, {apple_string(notice)} & return)')
    else:
        lines += ['if formulaFailed then', 'set r to my appendText(ownedDoc, "[EQUATION_INSERT_FAILED] Native equation failed; readable source retained." & return)', 'end if']
    return lines


_ALLOWED_ARGS = {
 'reset': set(),
 'configure_page': {'pageWidth','pageHeight','landscape','columns','header','footer','marginTop','marginBottom','marginLeft','marginRight'},
 'ensure_styles': {'styles'},
 'configure_front_matter': {'title','shortTitle','author','date','header','titlePage'},
 'configure_toc_styles': {'tocTitle','levels','includeFigureIndex','includeTableIndex','figureIndexTitle','tableIndexTitle','minFontSizePt','minSpaceBeforePt','minSpaceAfterPt'},
 'configure_section': {'pageSize','margins','role','pageNumberFormat','restartPageNumbering','headerText','footerText','linkToPreviousHeader','linkToPreviousFooter','startPageNumber','landscape'},
 'reserve_document_quality_anchor': {'title','notices'},
 'add_paragraph': {'color','spans','text','style','fontName','fontNameAscii','size','fontSize','bold','italic','align','indentFirst','spaceBefore','spaceAfter','keepWithNext','keepTogether','lineSpacing','lineSpacingRule'},
 'add_heading': {'size','bold','color','text','level','bookmarkName','numbering','numberingScheme','keepWithNext','sequenceTransparent'},
 'add_page_break': set(),
 'insert_toc': {'title'},
 'finalize_fields': {'maxRounds','compactTerminalParagraph'},
 'add_inline_degradation': {'code','fallbackText','message','text','placement','stage','fallback','recoverable'},
 'add_degradation_notice': {'code','fallbackText','message','text','placement','stage','fallback','recoverable'},
 'add_list': {'items','ordered','glyph'},
 'add_semantic_table': {'includePreviousHeading','continuousExit','cellCitations','cellDegradations','caption','numbering','indexable','referenceable','headers','rows','alignments','style','orientation','borderSpec','merges','repeatHeader','allowRowSplit','cellIndentPt','plannedDegradation','keepCaptionWithFirstRow','bookmarkName','m5Relayout'},
 'add_captioned_figure': {'columns','includePreviousHeading','continuousExit','caption','numbering','indexable','referenceable','widthMode','orientation','kind','children','layout','keepWithCaption','bookmarkName','explicitWidthPt'},
 'insert_figure_index': {'title','sequenceId','titleStyleId'},
 'insert_table_index': {'title','sequenceId','titleStyleId'},
 'add_equation': {'source','numbering','bookmarkName','fallbackText','renderMode','content','fallbackResource'},
 'add_cross_reference': {'runs','listFormatting'},
 'add_bibliography': {'schemaVersion','entries','style','hangingIndentPt','leftIndentPt','spaceAfterPt'},
}

def compile_plan(plan: GenerationPlan, resources: Mapping[str, Path], target: Path, *, timeout: float) -> CompiledScript:
    validate_generation_plan(plan.to_dict(), component='writer')
    supported = {'add_equation', 'reset', 'configure_page', 'ensure_styles', 'configure_front_matter', 'configure_toc_styles', 'configure_section', 'reserve_document_quality_anchor', 'add_paragraph', 'add_heading', 'add_page_break', 'insert_toc', 'finalize_fields', 'add_inline_degradation', 'add_degradation_notice', 'add_list', 'add_semantic_table', 'add_captioned_figure', 'insert_figure_index', 'insert_table_index', 'add_cross_reference', 'add_bibliography'}
    for operation in plan.operations:
        op = operation.op.removeprefix('writer.')
        if op not in supported:
            raise MacWordCapabilityError('Mac native Word unsupported operation: ' + op)
        extra = set(operation.args) - _ALLOWED_ARGS[op]
        if extra:
            raise MacWordCapabilityError('Mac native Word unsupported ' + op + ' attributes: ' + ', '.join(sorted(extra)))
        if operation.node_id and any(c in operation.node_id for c in '\t\n\r'):
            raise MacWordCapabilityError('Mac native Word node identity contains a record separator')
        if op == 'add_list' and operation.args.get('glyph','•') != '•':
            raise MacWordCapabilityError('Mac native Word unsupported custom bullet glyph')
        if op == 'ensure_styles':
            for style in operation.args['styles']:
                if style.get('type', 'paragraph') != 'paragraph':
                    raise MacWordCapabilityError('Mac Word unsupported character style')
        if op == 'add_paragraph' and 'spans' in operation.args and ''.join(span['text'] for span in operation.args['spans']) != operation.args['text']:
            raise MacWordCapabilityError('Mac Word span text must match paragraph text')
        if op == 'add_paragraph' and any(span.get('link') for span in operation.args.get('spans', ())):
            raise MacWordCapabilityError('Mac Word unsupported hyperlink span')
        if op == 'add_captioned_figure':
            for child in operation.args.get('children', ()):
                if child.get('plannedDegradation'):
                    continue
                if child.get('resourceId') not in resources:
                    raise MacWordCapabilityError('Mac native Word requires a bound normalized resource')
                if child.get('mediaType') not in ('image/png','image/jpeg'):
                    raise MacWordCapabilityError('Mac native Word unsupported native image resource type')
        if op == 'add_heading' and operation.args.get('numbering') and operation.args.get('numberingScheme') not in ('decimal','chinese-formal','hybrid-bid'):
            raise MacWordCapabilityError('Mac native Word unsupported heading numbering scheme')
    lines, nodes = [], {}
    issues = []
    section_seen, role = False, 'body'
    section_commands = []
    section_count = 0
    page_config = next((o.args for o in plan.operations if o.op == 'writer.configure_page'), {})
    current_landscape = bool(page_config.get('landscape'))
    numbered = any(o.op == 'writer.add_heading' and o.args.get('numbering') for o in plan.operations)
    numbering_schemes = {o.args.get('numberingScheme') for o in plan.operations if o.op == 'writer.add_heading' and o.args.get('numbering')}
    if len(numbering_schemes) > 1:
        raise MacWordCapabilityError('Mac Word unsupported mixed heading numbering schemes')
    scheme = next(iter(numbering_schemes), 'decimal')
    if numbered:
        lines += ['set ownList to make new list template at ownedDoc with properties {name:"WPSCNativeOutline", outline numbered:true}']
        numbered_levels = sorted({int(o.args['level']) for o in plan.operations if o.op == 'writer.add_heading' and o.args.get('numbering')})
        for level in range(1, max(numbered_levels) + 1):
            if scheme == 'chinese-formal':
                pattern = {1:'第%1章', 2:'第%2节', 3:'%3、', 4:'（%4）'}[level]
                number_style = 'simp chin num1'
            elif scheme == 'hybrid-bid':
                pattern = {1:'第%1章', 2:'%1.%2', 3:'%1.%2.%3', 4:'关键工法%4：'}[level]
                # Legal numbering renders inherited chapter counters as Arabic
                # in dotted subheadings while the chapter itself stays Chinese.
                number_style = 'simp chin num1' if level == 1 else 'arabic lz' if level == 4 else 'legal'
            else:
                pattern = '.'.join('%' + str(i) for i in range(1, level+1))
                number_style = 'arabic'
            # Linking from each style makes Word clone the outline per level.
            # Bind the localized style from the shared level instead, so new
            # chapters update every descendant without changing paragraph layout.
            lines += [f'set lvl to list level {level} of ownList', f'set number style of lvl to list number style {number_style}', 'set start at of lvl to 1', f'set reset on higher of lvl to {level-1}', f'set number format of lvl to {apple_string(pattern)}', f'set headingName to (name local of (Word style (style heading{level}) of ownedDoc)) as text', 'set linked style of lvl to headingName']
    style_defaults = {style['name']: style for o in plan.operations if o.op == 'writer.ensure_styles' for style in o.args['styles']}
    toc_levels = 3
    lines.append('set pendingHeadingStart to missing value')
    for operation in plan.operations:
        op, a = operation.op.removeprefix('writer.'), operation.args
        lines += ['set ownedDoc to document ' + apple_string(target.name), 'log ' + apple_string('WPSC_OP:' + operation.op)]
        is_content = op in {'add_paragraph', 'add_heading', 'insert_toc', 'add_inline_degradation', 'add_degradation_notice', 'configure_front_matter', 'reserve_document_quality_anchor', 'add_list', 'add_semantic_table', 'add_captioned_figure', 'insert_figure_index', 'insert_table_index', 'add_cross_reference', 'add_bibliography', 'add_equation'}
        media_landscape = op in ('add_semantic_table', 'add_captioned_figure') and a.get('orientation') == 'landscape'
        if media_landscape:
            lines += ['set p to (end of content of text object of ownedDoc) - 1']
            if a.get('includePreviousHeading'):
                lines += ['if pendingHeadingStart is not missing value then set p to pendingHeadingStart']
            lines += ['set r to create range ownedDoc start p end p', 'insert break at r break type section break next page']
            section_count = max(1, section_count) + 1
            section_seen = True
            section_commands += [f'set ownSection to section {section_count} of ownedDoc', 'set orientation of page setup of ownSection to orient landscape', 'set mediaFooter to get footer ownSection index header footer primary', 'set link to previous of mediaFooter to true', 'set restart numbering at section of page number options of mediaFooter to false']
        if is_content:
            lines.append('set nodeStart to (end of content of text object of ownedDoc) - 1')
        if op == 'configure_page':
            for key, prop in [('marginTop','top margin'), ('marginBottom','bottom margin'), ('marginLeft','left margin'), ('marginRight','right margin')]:
                lines.append(f'set {prop} of page setup of ownedDoc to {_num(a[key])}')
            lines += [f'set page width of page setup of ownedDoc to {_num(a.get("pageWidth",595.28))}', f'set page height of page setup of ownedDoc to {_num(a.get("pageHeight",841.89))}']
            if 'landscape' in a:
                lines.append(f'set orientation of page setup of ownedDoc to orient {"landscape" if a["landscape"] else "portrait"}')
            if 'columns' in a:
                lines.append(f'set number of text columns (page setup of ownedDoc) number of columns {int(a["columns"])}')
            for name in ('header','footer'):
                if name in a:
                    lines += [f'set pagePart to get {name} (section 1 of ownedDoc) index header footer primary', f'set content of text object of pagePart to {apple_string(a[name])}']
        elif op == 'ensure_styles':
            for style in a['styles']:
                try:
                    builtin = _style(style['name'])
                except MacWordCapabilityError:
                    lines.append(f'set ownStyle to make new Word style at ownedDoc with properties {{name local:{apple_string(style["name"])}}}')
                else:
                    lines.append(f'set ownStyle to Word style ({builtin}) of ownedDoc')
                if style.get('basedOn'):
                    lines.append(f'set base style of ownStyle to {_style(style["basedOn"], style_defaults)}')
                lines += _format('ownStyle', style)
        elif op == 'configure_front_matter':
            if a.get('titlePage') and a.get('title'):
                role = 'cover'
                for text, style in [(a['title'], 'Title'), (a.get('author',''), 'Body Text'), (a.get('date',''), 'Body Text')]:
                    if text:
                        lines += [f'set r to my appendText(ownedDoc, {apple_string(text)} & return)', f'set style of r to {_style(style)}', 'set alignment of paragraph format of r to align paragraph center', 'set first line indent of paragraph format of r to 0']
        elif op == 'configure_toc_styles':
            toc_levels = a.get('levels', 3)
            for level in range(1, toc_levels + 1):
                lines.append(f'set ownStyle to Word style (style toc{level}) of ownedDoc')
                lines += _format('ownStyle', {'fontSize':(a.get('minFontSizePt', {}).get('toc' + str(level), 10) if isinstance(a.get('minFontSizePt'), Mapping) else a.get('minFontSizePt', 10)), 'spaceBefore':(a.get('minSpaceBeforePt', {}).get('toc' + str(level), 0) if isinstance(a.get('minSpaceBeforePt'), Mapping) else a.get('minSpaceBeforePt', 0)), 'spaceAfter':(a.get('minSpaceAfterPt', {}).get('toc' + str(level), 0) if isinstance(a.get('minSpaceAfterPt'), Mapping) else a.get('minSpaceAfterPt', 0))})
        elif op == 'configure_section':
            role = a.get('role', role)
            current_landscape = a.get('landscape', current_landscape)
            if section_seen:
                lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', 'insert break at r break type section break next page', 'repaginate ownedDoc']
            section_seen = True
            section_count += 1
            section_start = len(lines)
            lines += [f'set ownSection to section {section_count} of ownedDoc', 'set ownHeader to get header ownSection index header footer primary', f'set link to previous of ownHeader to {_bool(a.get("linkToPreviousHeader"))}']
            if not a.get('linkToPreviousHeader'):
                lines.append(f'set content of text object of ownHeader to {apple_string(a.get("headerText", ""))}')
            lines += ['set ownFooter to get footer ownSection index header footer primary', f'set link to previous of ownFooter to {_bool(a.get("linkToPreviousFooter"))}']
            if not a.get('linkToPreviousFooter'):
                lines.append(f'set content of text object of ownFooter to {apple_string(a.get("footerText", ""))}')
            fmt = a.get('pageNumberFormat')
            if fmt not in (None, 'none', 'continue'):
                enum = {'roman':'lowercase roman','arabic':'arabic','roman-lower':'lowercase roman','roman-upper':'uppercase roman','lowerRoman':'lowercase roman','upperRoman':'uppercase roman'}.get(fmt)
                if enum is None:
                    raise MacWordCapabilityError('Mac native Word unsupported page-number format: ' + str(fmt))
                lines += [f'set number style of page number options of ownFooter to page number style {enum}', f'set restart numbering at section of page number options of ownFooter to {_bool(a.get("restartPageNumbering"))}']
                if 'startPageNumber' in a:
                    lines.append(f'set starting number of page number options of ownFooter to {int(a["startPageNumber"])}')
                if not a.get('linkToPreviousFooter'):
                    lines += ['set alignment of paragraph format of text object of ownFooter to align paragraph center', f'set content of text object of ownFooter to {apple_string(a.get("footerText", "") + " ")}', f'set r to character {len(a.get("footerText", "")) + 1} of text object of ownFooter', 'create new field text range r field type field page preserve formatting true']
            if 'pageSize' in a:
                sizes = {'a4':(595.28,841.89),'letter':(612,792),'legal':(612,1008),'a3':(841.89,1190.55)}
                if a['pageSize'].lower() not in sizes:
                    raise MacWordCapabilityError('Mac Word unsupported section pageSize')
                width,height = sizes[a['pageSize'].lower()]
                lines += [f'set page width of page setup of ownSection to {width}', f'set page height of page setup of ownSection to {height}']
            if 'landscape' in a:
                lines.append(f'set orientation of page setup of ownSection to orient {"landscape" if a["landscape"] else "portrait"}')
            for side,value in a.get('margins',{}).items():
                lines.append(f'set {side} margin of page setup of ownSection to {_num(value)}')
            section_commands.extend(lines[section_start:])
            del lines[section_start:]
        elif op in ('add_paragraph','add_heading'):
            lines.append(f'set r to my appendText(ownedDoc, {apple_string(a["text"])} & return)')
            heading_style = 'Heading ' + str(a['level']) if op == 'add_heading' else None
            lines.append(f'set style of r to {_style("Body Text" if a.get("sequenceTransparent") else heading_style or a.get("style", "Body Text"), style_defaults)}')
            if a.get('sequenceTransparent'):
                lines += _format('r', style_defaults[heading_style])
                lines += ['set outline level of paragraph format of r to outline level body text', 'set first line indent of paragraph format of r to 0']
            lines += _format('r', a)
            if a.get('spans'):
                offset = 0
                for span in a['spans']:
                    size = len(span['text'].encode('utf-16-le')) // 2
                    lines += [f'set spanRange to create range ownedDoc start (nodeStart + {offset}) end (nodeStart + {offset + size})']
                    lines += _format('spanRange', {k:span[k] for k in ('bold','italic','strikethrough') if k in span})
                    if span.get('code'):
                        lines += _format('spanRange', {'fontNameAscii':'Courier New'})
                    offset += size
            if op == 'add_heading' and numbered and not a.get('numbering'):
                lines.append('remove numbers (list format of r)')
        elif op == 'add_list':
            items = a['items']
            text = '\r'.join(items) + '\r'
            lines += [f'set r to my appendText(ownedDoc, {apple_string(text)})', 'set style of r to style body text', 'set first line indent of paragraph format of r to 0']
            lines.append(('apply number default' if a.get('ordered') else 'apply bullet default') + ' (list format of r)')
        elif op == 'add_semantic_table':
            # Citation descriptors have already rendered their resolved text in
            # the validated grid. M5 relayout is carried by merges/allowRowSplit.
            # Neither descriptor should append a second copy of cell content.
            if a.get('caption'):
                lines += _caption(a, before=True)
            row_values = [a['headers'], *a['rows']]
            columns = len(a['headers'])
            lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'set ownTable to make new table at ownedDoc with properties {{text object:r, number of rows:{len(row_values)}, number of columns:{columns}}}', 'set allow page breaks of ownTable to true', f'set heading format of row 1 of ownTable to {_bool(a.get("repeatHeader"))}', f'set left padding of ownTable to {_num(a.get("cellIndentPt",0))}', f'set right padding of ownTable to {_num(a.get("cellIndentPt",0))}']
            for row_index, row in enumerate(row_values, 1):
                lines.append(f'set allow break across pages of row {row_index} of ownTable to {_bool(a.get("allowRowSplit"))}')
                for col_index, text in enumerate(row, 1):
                    align = a['alignments'][col_index-1] if a.get('alignments') else 'left'
                    align = align if align in ('left','center','right') else 'left'
                    lines += [f'set ownCell to get cell from table ownTable row {row_index} column {col_index}', f'set content of text object of ownCell to {apple_string(text)}', f'set alignment of paragraph format of text object of ownCell to align paragraph {align}']
            lines += ['set r to text object of ownTable', 'set style of r to style body text', 'set first line indent of paragraph format of r to 0', 'set character unit first line indent of paragraph format of r to 0', 'set font size of font object of r to 10', 'set east asian name of font object of r to "宋体"']
            for key, border in [('top','top'),('bottom','bottom'),('left','left'),('right','right'),('insideHorizontal','horizontal'),('insideVertical','vertical')]:
                width = a['borderSpec'][key]
                lines += [f'set ownBorder to get border ownTable which border border {border}', f'set line style of ownBorder to line style {"single" if width else "none"}']
                if width:
                    lines.append(f'set line width of ownBorder to line width{int(width*100)} point')
            width = a['borderSpec']['headerBottom']
            for col in range(1, columns+1):
                lines += [f'set ownCell to get cell from table ownTable row 1 column {col}', 'set ownBorder to get border ownCell which border border bottom', f'set line style of ownBorder to line style {"single" if width else "none"}']
                if width:
                    lines.append(f'set line width of ownBorder to line width{int(width*100)} point')
            for degradation in a.get('cellDegradations', ()):
                lines += [f'set ownCell to get cell from table ownTable row {degradation["row"]} column {degradation["column"]}', 'set background pattern color of shading of text object of ownCell to {65535, 60138, 60138}']
                issues.append(ExecutionIssue(degradation['code'], degradation['fallbackText'], placement='inline', node_id=operation.node_id))
            for merge in sorted(a.get('merges', ()), key=lambda m:(m['top'],m['left']), reverse=True):
                lines += [f'set mergeStart to get cell from table ownTable row {merge["top"]} column {merge["left"]}', f'set mergeEnd to get cell from table ownTable row {merge["bottom"]} column {merge["right"]}', 'merge cell mergeStart with mergeEnd']
            lines += ['set r to my appendText(ownedDoc, return)']
            if a.get('plannedDegradation'):
                degradation = a['plannedDegradation']
                text = '[' + degradation['code'] + '] ' + degradation['message']
                lines.append(f'set r to my appendText(ownedDoc, {apple_string(text)} & return)')
                issues.append(ExecutionIssue(degradation['code'], degradation['message'], node_id=operation.node_id))
        elif op == 'add_captioned_figure':
            columns = a.get('layout') == 'columns'
            if columns:
                lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', 'set pictureTable to make new table at ownedDoc with properties {text object:r, number of rows:1, number of columns:2}']
                for border in ('top','bottom','left','right','horizontal','vertical'):
                    lines += [f'set ownBorder to get border pictureTable which border border {border}', 'set line style of ownBorder to line style none']
            for child_index, child in enumerate(a['children'], 1):
                if columns:
                    lines.append(f'set pictureCell to get cell from table pictureTable row 1 column {child_index}')
                if child.get('plannedDegradation'):
                    degradation = child['plannedDegradation']
                    text = degradation.get('fallback') or '[' + degradation['code'] + ']'
                    issues.append(ExecutionIssue(degradation['code'], degradation.get('message', text), placement='block', node_id=operation.node_id))
                    lines += [f'set content of text object of pictureCell to {apple_string(text)}'] if columns else [f'set r to my appendText(ownedDoc, {apple_string(text)} & return)']
                    continue
                lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'set ownPicture to make new inline picture at {"beginning" if columns else "end"} of text object of {"pictureCell" if columns else "ownedDoc"} with properties {{file name:{apple_string(str(resources[child["resourceId"]]))}, link to file:false, save with document:true}}', 'set ownPicture to inline picture (count of inline pictures of ownedDoc) of ownedDoc', f'set width of ownPicture to {_num(child["displayWidthPt"])}', f'set height of ownPicture to {_num(child["displayHeightPt"])}', 'set r to text object of ownPicture', 'set first line indent of paragraph format of r to 0', 'set alignment of paragraph format of r to align paragraph center', f'set keep with next of paragraph format of r to {_bool(a.get("keepWithCaption"))}', *([] if columns else ['set r to my appendText(ownedDoc, return)'])]
            if columns:
                lines += ['set r to my appendText(ownedDoc, return)']
            if a.get('caption'):
                lines += _caption(a, before=False)
        elif op == 'add_equation':
            lines += _equation(operation, resources, issues)
        elif op in ('insert_figure_index','insert_table_index'):
            code = '\\c "' + a['sequenceId'] + '" \\h \\z'
            lines += [f'set r to my appendText(ownedDoc, {apple_string(a["title"])} & return)', 'set style of r to style toc heading', 'set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'create new field text range r field type field toc field text {apple_string(code)} preserve formatting true', 'set r to my appendText(ownedDoc, return)']
        elif op == 'add_bibliography':
            for index, entry in enumerate(a['entries'], 1):
                text = ('[' + str(entry['number']) + '] ' + entry['text']) if isinstance(entry, Mapping) else ('[' + str(index) + '] ' + entry)
                lines += [f'set r to my appendText(ownedDoc, {apple_string(text)} & return)', 'set style of r to style body text', 'set alignment of paragraph format of r to align paragraph left', f'set paragraph format left indent of paragraph format of r to {_num(a.get("leftIndentPt",18))}', f'set first line indent of paragraph format of r to {_num(-a.get("hangingIndentPt",18))}', f'set space after of paragraph format of r to {_num(a.get("spaceAfterPt",6))}', 'set keep together of paragraph format of r to true']
        elif op == 'add_cross_reference':
            for run in a['runs']:
                if run['type'] == 'reference':
                    lines += [f'set r to my appendText(ownedDoc, {apple_string(run.get("prefix",""))})', 'set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'create new field text range r field type field ref field text {apple_string(run["bookmarkName"] + " " + chr(92) + "h")} preserve formatting true', f'set r to my appendText(ownedDoc, {apple_string(run.get("suffix",""))})']
                elif run['type'] == 'citation':
                    text = '[' + str(run['number']) + ']' if run.get('number') is not None else run['fallbackText']
                    lines.append(f'set r to my appendText(ownedDoc, {apple_string(text)})')
                else:
                    text = run.get('text',run.get('fallbackText',''))
                    if run['type'] == 'degradation':
                        issues.append(ExecutionIssue(run['code'],text,placement='inline',node_id=run.get('nodeId',operation.node_id)))
                    lines.append(f'set r to my appendText(ownedDoc, {apple_string(text)})')
            lines += ['set r to my appendText(ownedDoc, return)', 'set nodeEnd to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start nodeStart end nodeEnd', 'set style of r to style body text']
        elif op == 'insert_toc':
            lines += [f'set r to my appendText(ownedDoc, {apple_string(a["title"])} & return)', 'set style of r to style toc heading', 'set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'create new field text range r field type field toc field text {apple_string(chr(92)+"o " + chr(34) + "1-" + str(toc_levels) + chr(34) + " " + chr(92)+"h " + chr(92)+"z " + chr(92)+"u")} preserve formatting true', 'set r to my appendText(ownedDoc, return)']
        elif op == 'add_page_break':
            lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', 'insert break at r break type page break']
        elif op == 'reserve_document_quality_anchor':
            for notice in a.get('notices', ()):
                issues.append(ExecutionIssue(notice['code'],notice.get('message',notice['code']),node_id=operation.node_id))
                text = '[' + notice['code'] + '] ' + notice.get('message', '')
                lines.append(f'set r to my appendText(ownedDoc, {apple_string(text)} & return)')
        elif op in ('add_inline_degradation','add_degradation_notice'):
            text = a.get('fallbackText') or a.get('text') or '[' + a.get('code','DEGRADATION') + '] ' + a.get('message','')
            lines.append(f'set r to my appendText(ownedDoc, {apple_string(text)} & return)')
            issues.append(ExecutionIssue(a.get('code','DEGRADATION'), a.get('message',text), placement=a.get('placement','block'), node_id=operation.node_id))
        elif op == 'finalize_fields':
            if a.get('compactTerminalParagraph'):
                lines += ['set terminalRange to text object of last paragraph of ownedDoc', 'if content of terminalRange is return or content of terminalRange is "" then', 'set font size of font object of terminalRange to 1', 'set space before of paragraph format of terminalRange to 0', 'set space after of paragraph format of terminalRange to 0', 'set line spacing rule of paragraph format of terminalRange to line space exactly', 'set line spacing of paragraph format of terminalRange to 1', 'set keep with next of paragraph format of terminalRange to false', 'set keep together of paragraph format of terminalRange to false', 'end if']
        if op == 'add_cross_reference' and a.get('listFormatting'):
            lines += _format('r', {'leftIndent':a['listFormatting']['indentPt'], 'indentFirst':-a['listFormatting']['indentPt']})
        if is_content and operation.node_id and operation.node_id not in nodes:
            bookmark = (a.get('bookmarkName') if op not in ('add_captioned_figure','add_semantic_table','add_equation') else None) or 'wpsc_m5_' + hashlib.sha256(operation.node_id.encode()).hexdigest()[:24]
            nodes[operation.node_id] = (bookmark, role)
            lines += ['set nodeEnd to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start nodeStart end nodeEnd', f'make new bookmark at ownedDoc with properties {{name:{apple_string(bookmark)}, text object:r}}']
        if media_landscape:
            break_type = 'continuous' if a.get('continuousExit') else 'next page'
            lines += ['set p to (end of content of text object of ownedDoc) - 1', 'set r to create range ownedDoc start p end p', f'insert break at r break type section break {break_type}']
            section_count += 1
            section_commands += [f'set ownSection to section {section_count} of ownedDoc', f'set orientation of page setup of ownSection to orient {"landscape" if current_landscape else "portrait"}', 'set mediaFooter to get footer ownSection index header footer primary', 'set link to previous of mediaFooter to true', 'set restart numbering at section of page number options of mediaFooter to false']
        if op == 'add_heading':
            lines.append('set pendingHeadingStart to nodeStart')
        elif is_content:
            lines.append('set pendingHeadingStart to missing value')
    for command_index, command in enumerate(section_commands):
        lines += [f'log "WPSC_SECTION_COMMAND:{command_index}"', command]
    lines += [refresh_source(), pagination_source(nodes), f'save as ownedDoc file name {apple_string(str(target))} file format format document default add to recent files false', f'set ownedDoc to document {apple_string(target.name)}']
    return CompiledScript(wrap_owned('\n'.join(lines), target, timeout), nodes, len(plan.operations), tuple(issues))


def parse_result(raw: str, nodes, artifact: Path, operations: int, issues=()) -> ExecutionOutcome:
    observed, completed = {}, False
    runtime_issues = []
    fallback_nodes = set()
    for line in raw.splitlines():
        if line.startswith('WPSC_OK\t'):
            if completed or not re.fullmatch(r'WPSC_OK\t[0-9]+', line):
                raise ValueError('Malformed native completion record')
            completed = True
        elif line.startswith('WPSC_EQUATION_FALLBACK\t'):
            values = line.split('\t')
            if len(values) != 2 or values[1] not in nodes or values[1] in fallback_nodes:
                raise ValueError('Unexpected native equation fallback record')
            fallback_nodes.add(values[1])
            runtime_issues.append(ExecutionIssue('EQUATION_INSERT_FAILED', 'Native equation failed; readable source retained.', placement='block', node_id=values[1], stage='execute', fallback='source', recoverable=True))
        elif line.startswith('WPSC_NODE\t'):
            values = line.split('\t')
            if len(values) != 6 or values[1] not in nodes or values[1] in observed:
                raise ValueError('Unexpected native pagination record')
            if any(not re.fullmatch(r'[0-9]+', value) for value in values[2:]):
                raise ValueError('Invalid native pagination number')
            start_page, end_page, start, end = map(int, values[2:])
            if not 1 <= start_page <= end_page <= 100000 or start > end:
                raise ValueError('Invalid native pagination range')
            observed[values[1]] = PaginationNode(node_id=values[1], story='main', sections=(nodes[values[1]][1],), page_start=start_page, page_end=end_page, range=f'{start}:{end}', fragments=tuple(PaginationFragment(page=p) for p in range(start_page, end_page+1)))
    if not completed or set(observed) != set(nodes):
        raise ValueError('Native result incomplete; artifact cannot be published')
    return ExecutionOutcome(str(artifact), issues=tuple(issues) + tuple(runtime_issues), pagination_map=PaginationMap(version='M5-v1', nodes=tuple(observed[n] for n in nodes)), applied_operations=operations)
