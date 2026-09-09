"""Pure construction of the frozen Writer semantic-table primitive.

This is an internal, unconnected command builder, not a native session or a
recoverable public method. The caller supplies an exact main-story range and
last-section body width from the same guarded bound document. ``None`` width
means both section/document PageSetup reads failed, matching the frozen skip.

Commands preserve mutation order and carry the frozen operation-local failure
code. Those codes are *not* permission to convert transport exceptions into a
fallback: the session must establish ownership, deadline, submission state and
validated native completion itself. This module neither submits nor ACKs work.
"""
from __future__ import annotations

from dataclasses import dataclass
import math

from ..writer import NativeWriterObjectError, _content_column_widths
from .macos_word_recovery import _literal


@dataclass(frozen=True)
class SemanticTableCommand:
    """An ordered native request and its confirmed-operation failure category."""

    action: str
    args: tuple
    error_code: str


def _finite(value):
    number = float(value)
    if not math.isfinite(number):
        raise ValueError('Expected finite table measurement')
    return number


def _coordinate(value):
    if type(value) is not int or value < 0:
        raise ValueError('Expected nonnegative bound range coordinate')
    return value


def iter_semantic_table_commands(*, start, end, headers, rows, alignments,
                                 border_spec, repeat_header, allow_row_split,
                                 cell_indent_pt, merges, available_width_pt):
    """Yield construction requests in frozen order, without native execution.

    Descriptor errors are raised at their construction phase. Consumers needing
    their exact prefix can iterate; the eager builder validates everything
    before submission. Neither mode is the public direct-method error ordering.
    Ragged rows are not padded and merge order is not sorted or rewritten.
    """
    start, end = _coordinate(start), _coordinate(end)
    if end < start:
        raise ValueError('Bound range end precedes start')
    # The frozen materialization is deliberately outside its typed boundaries.
    data = [list(headers), *[list(row) for row in rows]]
    insert_error = 'TABLE_INSERT_FAILED'
    style_error = 'TABLE_STYLE_APPLY_FAILED'
    merge_error = 'TABLE_MERGE_APPLY_FAILED'
    try:
        yield SemanticTableCommand('create', (start, end, len(data), len(headers)), insert_error)
    except Exception:
        raise NativeWriterObjectError(insert_error) from None
    try:
        alignment_codes = {'left': 0, 'center': 1, 'right': 2}
        for row_index, row in enumerate(data, 1):
            for column_index, value in enumerate(row, 1):
                text = str(value)
                # Check the script text boundary while still in the style phase.
                _literal(text)
                indent = _finite(cell_indent_pt)
                alignment = alignment_codes[alignments[column_index - 1]]
                yield SemanticTableCommand('cell', (row_index, column_index, text, indent, alignment), style_error)
        yield SemanticTableCommand('row_split', (bool(allow_row_split),), style_error)
        if repeat_header:
            yield SemanticTableCommand('repeat_header', (True,), style_error)
        # Body width must already reflect the frozen last-section -> document
        # fallback. A failed measurement skips fitting rather than inventing one.
        if available_width_pt is not None:
            available_width = max(72.0, _finite(available_width_pt))
            widths = _content_column_widths(data, len(data[0]), available_width)
            yield SemanticTableCommand('fit', (available_width,), style_error)
            for index, width in enumerate(widths, 1):
                yield SemanticTableCommand('column_width', (index, float(width)), style_error)
        mapping = {'top': -1, 'left': -2, 'bottom': -3, 'right': -4,
                   'insideHorizontal': -5, 'insideVertical': -6}
        for key, border_id in mapping.items():
            points = _finite(border_spec[key])
            width = {0.75: 6, 1.5: 12}.get(points, 2) if points else None
            yield SemanticTableCommand('border', ('table', border_id, int(points != 0), width), style_error)
        points = _finite(border_spec['headerBottom'])
        width = {0.75: 6, 1.5: 12}.get(points, 2) if points else None
        yield SemanticTableCommand('border', ('header', -3, int(points != 0), width), style_error)
    except Exception:
        raise NativeWriterObjectError(style_error) from None
    try:
        for merge in merges:
            values = tuple(merge[key] for key in ('top', 'left', 'bottom', 'right'))
            # Do not validate overlap, dimensions, or reorder: those are native
            # operations in the oracle. Only forbid executable coordinate text.
            if any(type(value) is not int for value in values):
                raise ValueError('Expected integer merge coordinates')
            yield SemanticTableCommand('merge', values, merge_error)
    except Exception:
        raise NativeWriterObjectError(merge_error) from None
    yield SemanticTableCommand('finish', (), insert_error)


def build_semantic_table_commands(**kwargs):
    """Materialize the pure command sequence; does not open or mutate Word."""
    return tuple(iter_semantic_table_commands(**kwargs))


def _best_effort(lines, fallback=()):
    # Mirror only the frozen optional width-setting try blocks. Transport
    # failures remain uncertain and must escape to the session quarantine path.
    return ['try', *lines, 'on error semanticMessage number semanticError',
            'if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError',
            *fallback, 'end try']


def render_semantic_table_commands(commands):
    """Emit a body for the existing guarded boundDoc/boundWindow script scope.

    No handler, app binding, selection lookup outside boundWindow, catch-to-
    fallback, success ACK, or document-end positioning is introduced here.
    The table reference returned by ``make`` is retained even for insertion
    before an existing table; ``count tables`` is never used as its identity.
    """
    lines = []
    for command in commands:
        action, args = command.action, command.args
        if action == 'create':
            start, end, rows, columns = args
            lines += [f'set semanticInsertion to create range boundDoc start {start} end {end}',
                      f'set semanticTable to make new table at boundDoc with properties {{text object:semanticInsertion, number of rows:{rows}, number of columns:{columns}}}']
        elif action == 'cell':
            row, column, text, indent, alignment = args
            align = {0: 'left', 1: 'center', 2: 'right'}[alignment]
            lines += [f'set semanticCell to get cell from table semanticTable row {row} column {column}',
                      f'set content of text object of semanticCell to {_literal(text)}',
                      f'set first line indent of paragraph format of text object of semanticCell to {indent}',
                      'set paragraph format left indent of paragraph format of text object of semanticCell to 0.0',
                      'set paragraph format right indent of paragraph format of text object of semanticCell to 0.0',
                      f'set alignment of paragraph format of text object of semanticCell to align paragraph {align}']
        elif action == 'row_split':
            lines += [f'set allow break across pages of row options of semanticTable to {str(args[0]).lower()}']
        elif action == 'repeat_header':
            lines += ['set heading format of row 1 of semanticTable to true']
        elif action == 'fit':
            lines += _best_effort(['auto fit behavior semanticTable behavior auto fit fixed'])
            lines += ['set allow auto fit of semanticTable to false']
            lines += _best_effort(['set preferred width type of semanticTable to preferred width points',
                                  f'set preferred width of semanticTable to {args[0]}'])
        elif action == 'column_width':
            index, width = args
            lines += [f'set semanticColumn to column {index} of semanticTable']
            lines += _best_effort([f'set table item width semanticColumn column width {width} ruler style adjust none'],
                                 [f'set width of semanticColumn to {width}'])
        elif action == 'border':
            scope, position, style, width = args
            target = 'semanticTable' if scope == 'table' else 'row 1 of semanticTable'
            side = {-1: 'top', -2: 'left', -3: 'bottom', -4: 'right', -5: 'horizontal', -6: 'vertical'}[position]
            lines += [f'set semanticBorder to get border {target} which border border {side}',
                      'set line style of semanticBorder to line style ' + ('single' if style else 'none')]
            if width is not None:
                lines += [f'set line width of semanticBorder to line width{ {2: 25, 6: 75, 12: 150}[width]} point']
        elif action == 'merge':
            top, left, bottom, right = args
            lines += [f'set semanticMergeStart to get cell from table semanticTable row {top} column {left}',
                      f'set semanticMergeEnd to get cell from table semanticTable row {bottom} column {right}',
                      'merge cell semanticMergeStart with semanticMergeEnd']
        elif action == 'finish':
            lines += ['set semanticTableEnd to end of content of text object of semanticTable',
                      'set semanticSelection to selection of boundWindow',
                      'set selection start of semanticSelection to semanticTableEnd',
                      'set selection end of semanticSelection to semanticTableEnd',
                      'set semanticTail to create range boundDoc start semanticTableEnd end semanticTableEnd',
                      'set content of semanticTail to return',
                      'set selection start of semanticSelection to semanticTableEnd + 1',
                      'set selection end of semanticSelection to semanticTableEnd + 1']
        else:
            raise ValueError('Unknown semantic table command')
    return lines
