"""Pure request-contract preflight for automatic macOS WPS PPT edit routing.

Mirrors presentation.js's edit handlers, not the broader COM patch grammar.
A passing result means every requested property reaches an implemented handler;
it does not prove a shape exists, has a text frame, or accepts the native setter.
Explicit WPS calls deliberately do not use this automatic-selection filter.
"""
from __future__ import annotations

import math
import re
from collections.abc import Mapping


def _number(value):
    if type(value) not in (int, float):
        return False
    try:
        return math.isfinite(value)
    except OverflowError:
        return False


def _nonnegative(value):
    return _number(value) and value >= 0


def _positive(value):
    return _number(value) and value > 0


def _tristate(value):
    return type(value) is bool or (_number(value) and value in (-3, -2, -1, 0, 1))


def _enum(*values):
    return lambda value: _number(value) and value in values


def _text(value):
    return isinstance(value, str)


def _text_number(value):
    # JSON numbers become JavaScript doubles before String(text). Beyond the
    # safe integer range, integer text can change while WPS reports success.
    return _number(value) and (type(value) is not int or abs(value) <= 2**53 - 1)


def _color(value):
    return ((_number(value) and 0 <= value <= 0xFFFFFF and value == int(value)) or
            (isinstance(value, str) and re.fullmatch(r'#?[0-9a-fA-F]{6}', value) is not None))


_FILL = {'color': _color, 'visible': _tristate,
         'transparency': lambda value: _number(value) and 0 <= value <= 1}
_GROUPS = {
    'font': {'name': _text, 'size': _positive, 'bold': _tristate,
             'italic': _tristate, 'underline': _tristate,
             'strikethrough': _tristate, 'color': _color},
    'paragraph': {'alignment': _enum(-2, 1, 2, 3, 4, 5, 6, 7),
                  'left_indent': _number, 'first_line_indent': _number,
                  'line_spacing': _nonnegative, 'line_rule_within': _tristate,
                  'space_before': _nonnegative, 'space_after': _nonnegative},
    'geometry': {'left': _number, 'top': _number, 'width': _nonnegative,
                 'height': _nonnegative, 'rotation': _number},
    'fill': _FILL,
    'background': _FILL,
    'line': {'color': _color, 'visible': _tristate, 'weight': _nonnegative},
    'text_frame': {'margin_left': _nonnegative, 'margin_right': _nonnegative,
                   'margin_top': _nonnegative, 'margin_bottom': _nonnegative,
                   'word_wrap': _tristate, 'auto_size': _enum(-2, 0, 1),
                   'vertical_anchor': _enum(-2, 1, 2, 3, 4, 5)},
}
_FIELDS = {
    'slide': {'name', 'follow_master_background', 'background'},
    'shape': {'text', 'font', 'paragraph', 'geometry', 'fill', 'line',
              'text_frame', 'vertical_alignment'},
    'text': {'text', 'font', 'paragraph'},
}
_INDEX = r'[0-9]+'
_SHAPE = rf'slide:({_INDEX})/shape:(?:@id=)?({_INDEX})'


def _target_kind(target):
    if not isinstance(target, str):
        return None
    patterns = (
        (rf'{_SHAPE}/table/cell:({_INDEX}),({_INDEX})', 'shape'),
        (rf'{_SHAPE}/paragraph:({_INDEX})(?:/run:({_INDEX}))?', 'text'),
        (_SHAPE, 'shape'),
        (rf'slide:({_INDEX})', 'slide'),
    )
    for pattern, kind in patterns:
        match = re.fullmatch(pattern, target)
        if match:
            # Avoid enormous-integer conversions while preserving leading zeroes.
            return kind if all(value is None or value.lstrip('0') for value in match.groups()) else None
    named = re.fullmatch(rf'slide:({_INDEX})/shape:@name=(.+)', target)
    if named and named[1].lstrip('0'):
        # The native regex greedily treats suffixes as part of Name. Do not
        # advertise unimplemented @name paragraph/table descendants.
        if not re.search(r'/(?:paragraph:|run:|table/)', named[2]):
            return 'shape'
    return None


def explain_presentation_set_ops(operations):
    """Return deterministic unsupported field paths without inspecting documents."""
    problems = []
    try:
        batch = tuple(operations)
    except TypeError:
        return ('operations',)
    for index, operation in enumerate(batch):
        prefix = f'operations[{index}]'
        if not isinstance(operation, Mapping):
            problems.append(prefix)
            continue
        if operation.get('op') != 'set':
            problems.append(prefix + '.op')
        kind = _target_kind(operation.get('target'))
        if kind is None:
            problems.append(prefix + '.target')
            continue
        for key, value in operation.items():
            if key in ('op', 'target'):
                continue
            path = prefix + '.' + str(key)
            if key not in _FIELDS[kind]:
                problems.append(path)
            elif key in _GROUPS:
                if value is None:  # JS patch.group || {} is an intentional no-op.
                    continue
                if not isinstance(value, Mapping):
                    problems.append(path)
                    continue
                for child, item in value.items():
                    check = _GROUPS[key].get(child)
                    if check is None or not check(item):
                        problems.append(path + '.' + str(child))
                if (key in ('fill', 'line', 'background') and
                        'visible' in value and 'color' in value and
                        value['visible'] not in (True, -1, 1) and
                        tuple(value).index('visible') < tuple(value).index('color')):
                    # The shipped color setter forces Visible=-1. JSON key
                    # order is preserved; a later explicit visibility setter
                    # still works and must remain eligible for WPS.
                    problems.append(path + '.visible')
            elif key == 'text':
                if value is not None and not (_text(value) or type(value) is bool or _text_number(value)):
                    problems.append(path)
            elif key == 'name':
                if not _text(value):
                    problems.append(path)
            elif key == 'follow_master_background':
                if not _tristate(value):
                    problems.append(path)
            elif key == 'vertical_alignment':
                if value is not None and not _GROUPS['text_frame']['vertical_anchor'](value):
                    problems.append(path)
        frame = operation.get('text_frame')
        vertical = operation.get('vertical_alignment')
        if (isinstance(frame, Mapping) and 'vertical_anchor' in frame and
                vertical is not None and frame['vertical_anchor'] != vertical):
            problems.append(prefix + '.vertical_alignment')
    return tuple(problems)


def supports_presentation_set_ops(operations):
    return not explain_presentation_set_ops(operations)
