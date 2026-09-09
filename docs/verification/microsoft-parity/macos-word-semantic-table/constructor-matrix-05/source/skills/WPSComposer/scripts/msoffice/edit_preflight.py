"""Pure Microsoft edit capability validation; never opens Office or resolves live IDs.

The Mac patch compilers are also used by their session executors. Static target
forms are distinct from live existence, selection and ownership validation.
Windows keeps its COM grammar and does not inherit dictionary restrictions.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Literal


WindowsCapability = Literal['supported', 'unsupported', 'unknown']

_FONT_FIELDS = {
    'name', 'size', 'bold', 'italic', 'underline', 'strikethrough', 'color',
}
_PARAGRAPH_FIELDS = {
    'alignment', 'left_indent', 'right_indent', 'first_line_indent',
    'space_before', 'space_after', 'line_spacing', 'line_spacing_rule',
    'keep_together', 'keep_with_next', 'page_break_before', 'widow_control',
}
_GEOMETRY_FIELDS = {'left', 'top', 'width', 'height', 'rotation'}
_FILL_FIELDS = {'color', 'back_color', 'visible', 'transparency'}
_LINE_FIELDS = {'color', 'visible', 'weight', 'dash_style', 'transparency'}
_TEXT_FRAME_FIELDS = {
    'margin_left', 'margin_right', 'margin_top', 'margin_bottom',
    'word_wrap', 'auto_size', 'vertical_anchor',
}
_WORD_PAGE_FIELDS = {
    'orientation', 'page_width', 'page_height', 'top_margin', 'bottom_margin',
    'left_margin', 'right_margin', 'header_distance', 'footer_distance',
    'gutter', 'paper_size',
}
_SHEET_PAGE_FIELDS = {
    'orientation', 'top_margin', 'bottom_margin', 'left_margin',
    'right_margin', 'header_margin', 'footer_margin', 'paper_size', 'zoom',
    'fit_to_pages_wide', 'fit_to_pages_tall', 'print_area',
    'print_title_rows', 'print_title_columns',
}
_SLIDE_PAGE_FIELDS = {'slide_width', 'slide_height'}

_WINDOWS_SIGNATURE_FIELDS = {
    'writer': {
        'text', 'font', 'paragraph', 'geometry', 'fill', 'line', 'style',
        'wrap', 'vertical_alignment', 'page_setup', 'columns',
    },
    'sheet': {
        'value', 'formula', 'font', 'fill', 'line', 'geometry',
        'number_format', 'horizontal_alignment', 'vertical_alignment',
        'wrap_text', 'indent', 'row_height', 'column_width', 'borders',
        'page_setup', 'name', 'chart_type', 'chart_title',
    },
    'slide': {
        'text', 'font', 'paragraph', 'geometry', 'fill', 'line', 'text_frame',
        'name', 'shape_type', 'background', 'follow_master_background',
        'page_setup', 'vertical_alignment',
    },
}

_WINDOWS_TARGET_FIELDS = {
    'writer': {
        'range': {'text', 'font', 'paragraph', 'style'},
        'cell': {'text', 'font', 'paragraph', 'style', 'fill', 'vertical_alignment'},
        'shape': {'text', 'font', 'geometry', 'fill', 'line', 'wrap'},
        'section': {'page_setup', 'columns'},
    },
    'sheet': {
        'range': {
            'value', 'formula', 'font', 'fill', 'number_format',
            'horizontal_alignment', 'vertical_alignment', 'wrap_text',
            'indent', 'row_height', 'column_width', 'borders',
        },
        'shape': {'geometry', 'fill', 'line', 'name'},
        'chart': {'geometry', 'chart_type', 'chart_title'},
        'sheet': {'name', 'page_setup'},
    },
    'slide': {
        'presentation': {'page_setup'},
        'slide': {'name', 'background', 'follow_master_background'},
        'text': {'text', 'font', 'paragraph'},
        'cell': {
            'text', 'font', 'paragraph', 'fill', 'line', 'text_frame',
            'vertical_alignment',
        },
        'shape': {
            'text', 'font', 'paragraph', 'geometry', 'fill', 'line',
            'text_frame', 'vertical_alignment', 'name',
        },
        # A live selection is either a text range or a shape range. The
        # concrete kind remains an identity-bound session check.
        'selection': {
            'text', 'font', 'paragraph', 'geometry', 'fill', 'line',
            'text_frame', 'vertical_alignment',
        },
    },
}


def _windows_target_kind(family, target):
    if family == 'writer':
        if target == 'selection':
            return 'range'
        if not isinstance(target, str):
            return None
        if re.fullmatch(r'paragraph:(?:[1-9]\d*|@paraId=[0-9a-fA-F]+)', target):
            return 'range'
        if re.fullmatch(r'range:\d+-\d+', target):
            return 'range'
        if re.fullmatch(r'table:[1-9]\d*/cell:[1-9]\d*,[1-9]\d*', target):
            return 'cell'
        if re.fullmatch(r'shape:[1-9]\d*', target):
            return 'shape'
        if re.fullmatch(r'section:[1-9]\d*', target):
            return 'section'
        return None
    if family == 'sheet':
        if target == 'selection':
            return 'range'
        if not isinstance(target, str):
            return None
        if re.fullmatch(r'sheet:[1-9]\d*/(?:cell|range):.+', target):
            return 'range'
        if re.fullmatch(r'sheet:[1-9]\d*/shape:(?:[1-9]\d*|@id=[1-9]\d*|@name=.+)', target):
            return 'shape'
        if re.fullmatch(r'sheet:[1-9]\d*/chart:[1-9]\d*', target):
            return 'chart'
        if re.fullmatch(r'sheet:[1-9]\d*', target):
            return 'sheet'
        return None
    if family == 'slide':
        if target == 'selection':
            return 'selection'
        if target == 'presentation':
            return 'presentation'
        if not isinstance(target, str):
            return None
        if re.fullmatch(r'slide:[1-9]\d*', target):
            return 'slide'
        shape = r'(?:[1-9]\d*|@id=[1-9]\d*|@name=.+)'
        stable_nested = r'(?:[1-9]\d*|@id=[1-9]\d*)'
        if re.fullmatch(rf'slide:[1-9]\d*/shape:{stable_nested}/table/cell:[1-9]\d*,[1-9]\d*', target):
            return 'cell'
        if re.fullmatch(rf'slide:[1-9]\d*/shape:{stable_nested}/paragraph:[1-9]\d*(?:/run:[1-9]\d*)?', target):
            return 'text'
        if re.fullmatch(rf'slide:[1-9]\d*/shape:{shape}', target):
            return 'shape'
    return None


def _has_requested_value(value):
    return value is not None and not (isinstance(value, dict) and not value)


def _valid_color(value):
    from .._colors import hex_to_rgb_long
    try:
        hex_to_rgb_long(value)
        return True
    except (TypeError, ValueError, OverflowError):
        return False


def _valid_nested(value, allowed, *, color_fields=()):
    if not isinstance(value, dict) or set(value) - set(allowed):
        return False
    return all(_valid_color(value[key]) for key in color_fields if key in value)


def _valid_borders(value):
    if not isinstance(value, dict):
        return False
    for edge, spec in value.items():
        try:
            int(edge)
        except (TypeError, ValueError, OverflowError):
            return False
        if not isinstance(spec, dict) or set(spec) - {'style', 'weight', 'color'}:
            return False
        if 'color' in spec and not _valid_color(spec['color']):
            return False
    return True


def _windows_sheet_values_are_screened(patch):
    """Reuse the existing Microsoft spreadsheet automation-string screen."""
    from .windows_office_runtime import validate_plan

    def strings(value):
        if isinstance(value, dict):
            for nested in value.values():
                yield from strings(nested)
        elif isinstance(value, (list, tuple)):
            for nested in value:
                yield from strings(nested)
        elif isinstance(value, str):
            yield value

    cells = [[value] for value in strings(patch)]
    if cells:
        validate_plan({
            'component': 'spreadsheet',
            'operations': [
                {'op': 'sheet.reset', 'args': {}},
                {'op': 'sheet.write_table', 'args': {
                    'startRow': 1, 'startCol': 1, 'values': cells,
                }},
            ],
        }, {})


def classify_windows_set_op(family, op) -> WindowsCapability:
    """Classify deterministic Python-side support for one Windows set op.

    A supported result means the current composer has an implemented target
    branch and conversion path. Native COM acceptance and live target
    existence remain deferred to the identity-bound session.
    """
    try:
        from ..document_api import validate_op
        if not validate_op(op, family).get('valid') or op.get('op', 'set') != 'set':
            return 'unsupported'
        signature = _WINDOWS_SIGNATURE_FIELDS.get(family)
        target_kind = _windows_target_kind(family, op.get('target'))
        if signature is None or target_kind is None:
            return 'unsupported'
        patch = {key: value for key, value in op.items() if key not in {'op', 'target'}}
        if set(patch) - signature:
            return 'unsupported'
        requested = {key for key, value in patch.items() if _has_requested_value(value)}
        if requested - _WINDOWS_TARGET_FIELDS[family][target_kind]:
            return 'unsupported'

        nested = {
            'font': (_FONT_FIELDS, {'color'}),
            'paragraph': (_PARAGRAPH_FIELDS, set()),
            'geometry': (_GEOMETRY_FIELDS, set()),
            'line': (_LINE_FIELDS, {'color'}),
            'text_frame': (_TEXT_FRAME_FIELDS, set()),
            'background': (_FILL_FIELDS, {'color', 'back_color'}),
        }
        for key, (allowed, colors) in nested.items():
            if key in patch and _has_requested_value(patch[key]):
                if not _valid_nested(patch[key], allowed, color_fields=colors):
                    return 'unsupported'
        if 'fill' in patch and _has_requested_value(patch['fill']):
            allowed = {'color'} if target_kind in {'cell', 'range'} else _FILL_FIELDS
            if not _valid_nested(
                    patch['fill'], allowed,
                    color_fields={'color', 'back_color'} & allowed):
                return 'unsupported'
        if 'page_setup' in patch and _has_requested_value(patch['page_setup']):
            allowed = {
                'writer': _WORD_PAGE_FIELDS,
                'sheet': _SHEET_PAGE_FIELDS,
                'slide': _SLIDE_PAGE_FIELDS,
            }[family]
            if not _valid_nested(patch['page_setup'], allowed):
                return 'unsupported'
        if 'borders' in patch and _has_requested_value(patch['borders']):
            if not _valid_borders(patch['borders']):
                return 'unsupported'
        if family == 'sheet':
            _windows_sheet_values_are_screened(patch)
        return 'supported'
    except (ValueError, TypeError, KeyError, AttributeError, OverflowError):
        return 'unsupported'


def _positive_int(value):
    return type(value) is int and value >= 1


def _windows_slide_structural(op):
    verb = op.get('op')
    if verb == 'insert':
        kind = op.get('type')
        props = op.get('props') or {}
        if not isinstance(props, dict):
            return False
        if kind == 'slide':
            try:
                int(props.get('layout', 12))
            except (TypeError, ValueError, OverflowError):
                return False
            return None if set(props) - {'layout'} else True
        if kind not in {'textbox', 'image'}:
            return False
        if not re.fullmatch(r'slide:[1-9]\d*', str(op.get('parent', ''))):
            return False
        allowed = {'left', 'top', 'width', 'height', 'text'}
        if kind == 'image':
            allowed.add('path')
        if set(props) - allowed:
            return None
        try:
            for key in {'left', 'top', 'width', 'height'} & set(props):
                float(props[key])
        except (TypeError, ValueError, OverflowError):
            return False
        return kind != 'image' or isinstance(props.get('path'), (str, Path))
    target = str(op.get('target', ''))
    match = re.fullmatch(
        r'slide:[1-9]\d*(?:/shape:(?:[1-9]\d*|@id=[1-9]\d*|@name=.+))?',
        target,
    )
    if verb == 'remove':
        return match is not None
    if verb not in {'move', 'clone'} or match is None:
        return False
    to = op.get('to', 'end')
    if '/shape:' in target:
        if to in (None, 'start', 'end'):
            return True
        return (isinstance(to, dict) and set(to) == {'slide'} and
                _positive_int(to['slide']))
    if to in (None, 'start', 'end'):
        return True
    if not isinstance(to, dict) or len(to) != 1:
        return False
    key, value = next(iter(to.items()))
    if key == 'index':
        return _positive_int(value)
    return key in {'before', 'after'} and bool(
        re.fullmatch(r'slide:[1-9]\d*', str(value)))


def _windows_sheet_structural(op, engine):
    verb = op.get('op')
    props = op.get('props') or {}
    if not isinstance(props, dict):
        return False
    if verb == 'insert':
        kind = op.get('type')
        if kind == 'sheet':
            return None if set(props) - {'name'} else True
        if kind not in {'row', 'column'}:
            return False
        if set(props) - {'values'}:
            return None
        if not re.fullmatch(r'sheet:[1-9]\d*(?:/(?:cell|range):.+)?', str(op.get('parent', ''))):
            return False
        position = op.get('position', 'end')
        if position in (None, 'end'):
            return True
        return (isinstance(position, dict) and set(position) == {'index'} and
                _positive_int(position['index']))
    target = str(op.get('target', ''))
    sheet = re.fullmatch(r'sheet:[1-9]\d*', target)
    cell = re.fullmatch(r'sheet:[1-9]\d*/(?:cell|range):.+', target)
    object_target = re.fullmatch(
        r'sheet:[1-9]\d*/(?:shape:(?:[1-9]\d*|@id=[1-9]\d*|@name=.+)|chart:[1-9]\d*)',
        target,
    )
    if verb == 'remove':
        return bool(sheet or cell or object_target) and (
            not cell or op.get('axis', 'row') in {'row', 'column'})
    if verb not in {'move', 'clone'}:
        return False
    to = op.get('to')
    if sheet:
        if to is None:
            return True
        if to == 'end':
            return verb == 'move' or engine == 'msoffice'
        if to == 'start':
            return engine == 'msoffice'
        return (isinstance(to, dict) and len(to) == 1 and
                next(iter(to)) in {'before', 'after'} and
                _positive_int(next(iter(to.values()))))
    if not cell or not isinstance(to, dict) or set(to) != {'index'}:
        return False
    if not _positive_int(to['index']):
        return False
    if verb == 'clone' and op.get('axis', 'row') != 'row':
        return False
    return op.get('axis', 'row') in {'row', 'column'}


def _windows_writer_structural(op):
    verb = op.get('op')
    props = op.get('props') or {}
    if not isinstance(props, dict):
        return False
    if verb == 'insert':
        if op.get('parent', 'body') != 'body':
            return False
        allowed = {
            'paragraph': {'text', 'style', 'level'},
            'heading': {'text', 'style', 'level'},
            'page_break': set(),
            'table': {'rows', 'cols', 'data'},
            'textbox': {'text', 'left', 'top', 'width', 'height'},
            'image': {'path'},
        }
        kind = op.get('type')
        if kind not in allowed:
            return False
        if set(props) - allowed[kind]:
            return None
        position = op.get('position', 'end')
        if position not in (None, 'start', 'end'):
            if not isinstance(position, dict) or len(position) != 1:
                return False
            key, value = next(iter(position.items()))
            if key == 'index':
                if not _positive_int(value):
                    return False
            elif key not in {'before', 'after'} or not re.fullmatch(
                    r'paragraph:(?:[1-9]\d*|@paraId=[0-9a-fA-F]+)', str(value)):
                return False
        try:
            if kind == 'heading':
                int(props.get('level', 1))
            if kind == 'table':
                rows, cols = int(props.get('rows', 2)), int(props.get('cols', 2))
                if rows < 1 or cols < 1:
                    return False
            if kind == 'textbox':
                for key in {'left', 'top', 'width', 'height'} & set(props):
                    float(props[key])
        except (TypeError, ValueError, OverflowError):
            return False
        return kind != 'image' or isinstance(props.get('path'), (str, Path))
    target = str(op.get('target', ''))
    if not re.fullmatch(
            r'(?:paragraph:(?:[1-9]\d*|@paraId=[0-9a-fA-F]+)|'
            r'(?:table|shape|inline_shape):[1-9]\d*)', target):
        return False
    if verb == 'remove':
        return True
    if verb not in {'move', 'clone'}:
        return False
    to = op.get('to', 'end')
    if to in (None, 'start', 'end'):
        return True
    if not isinstance(to, dict) or len(to) != 1:
        return False
    key, value = next(iter(to.items()))
    return (_positive_int(value) if key == 'index' else
            key in {'before', 'after'} and bool(re.fullmatch(
                r'paragraph:(?:[1-9]\d*|@paraId=[0-9a-fA-F]+)', str(value))))


def classify_windows_structural_op(
        family, op, *, engine='msoffice') -> WindowsCapability:
    try:
        from ..document_api import validate_op
        if not validate_op(op, family).get('valid') or op.get('op') == 'set':
            return 'unsupported'
        supported = {
            'writer': lambda: _windows_writer_structural(op),
            'sheet': lambda: _windows_sheet_structural(op, engine),
            'slide': lambda: _windows_slide_structural(op),
        }[family]()
        if supported is None:
            return 'unknown'
        return 'supported' if supported else 'unsupported'
    except (ValueError, TypeError, KeyError, AttributeError, OverflowError):
        return 'unsupported'


def rejects_windows_edit_ops(family, operations, *, engine='msoffice'):
    for op in operations:
        state = (classify_windows_set_op(family, op)
                 if op.get('op', 'set') == 'set'
                 else classify_windows_structural_op(family, op, engine=engine))
        if state == 'unsupported':
            return True
    return False


def rejects_windows_common_edit_ops(family, operations):
    """Reject only invariants shared by both Windows candidate engines."""
    from ..document_api import validate_op
    for op in operations:
        if not validate_op(op, family).get('valid'):
            return True
        if (op.get('op', 'set') == 'set' and
                classify_windows_set_op(family, op) == 'unsupported'):
            return True
    return False


def word_target_kind(target):
    if target == 'selection': return 'range'
    if not isinstance(target, str): raise ValueError('Unsupported Writer target')
    if re.fullmatch(r'paragraph:(?:[1-9]\d*|@paraId=[0-9a-fA-F]+)', target): return 'range'
    match = re.fullmatch(r'range:(\d+)-(\d+)', target)
    if match and int(match[1]) <= int(match[2]): return 'range'
    if re.fullmatch(r'table:[1-9]\d*/cell:[1-9]\d*,[1-9]\d*', target): return 'cell'
    if re.fullmatch(r'shape:[1-9]\d*', target): return 'shape'
    if re.fullmatch(r'section:[1-9]\d*', target): return 'section'
    raise ValueError('Unsupported Writer target')


def compile_word_patch(target, patch):
    from .macos_word_session import _FONT, _PARA, _GEOMETRY, _PAGE, _WRAP, _value
    word_target_kind(target)
    if set(patch)-{'text','font','paragraph','geometry','fill','line','style','wrap','vertical_alignment','page_setup','columns'}:
        raise ValueError('Unsupported Writer patch fields')
    patch = {key:value for key,value in patch.items() if value is not None}
    accepted, rejected, mutations = [], [], []
    shape = re.fullmatch(r'shape:([1-9]\d*)', target or '')
    section = re.fullmatch(r'section:([1-9]\d*)', target or '')
    cell = target.startswith('table:') if isinstance(target, str) else False
    mappings = {'font': (_FONT, 'font object of targetRange'), 'paragraph': (_PARA, 'paragraph format of targetRange')}
    if shape:
        mappings.update({'geometry': (_GEOMETRY, 'targetShape'), 'fill': ({'color': 'fore color', 'back_color': 'back color', 'visible': 'visible', 'transparency': 'transparency'}, 'fill format of targetShape'), 'line': ({'color': 'fore color', 'weight': 'weight', 'visible': 'visible', 'transparency': 'transparency'}, 'line format of targetShape')})
    if cell:
        mappings['fill'] = ({'color': 'background pattern color'}, 'shading of targetCell')
    if section:
        mappings = {'page_setup': (_PAGE, 'targetSetup')}
    for group, values in patch.items():
        if group in mappings and isinstance(values, dict):
            names, obj = mappings[group]
            for key, value in values.items():
                label = group + '.' + key
                try:
                    if key not in names:
                        raise ValueError()
                    rendered = _value(key, value)
                    mutations.append(f'set {names[key]} of {obj} to {rendered}')
                    accepted.append(label)
                except (ValueError, TypeError):
                    rejected.append(label)
        elif group in ('text', 'style') and not section:
            try:
                rendered = _value(group, values)
                if group == 'text' and (target.startswith('paragraph:') or cell):
                    mutations += ['set replacementStart to start of content of targetRange', 'set replacementEnd to end of content of targetRange', 'set replacementRange to create range boundDoc start replacementStart end (replacementEnd - 1)', f'set content of replacementRange to {rendered}']
                else:
                    mutations.append(f'set {"content" if group == "text" else "style"} of targetRange to {rendered}')
                accepted.append(group)
            except (TypeError, ValueError):
                rejected.append(group)
        elif group == 'vertical_alignment' and cell:
            try:
                mutations.append('set vertical alignment of targetCell to ' + _value(group, values))
                accepted.append(group)
            except (TypeError, ValueError):
                rejected.append(group)
        elif group == 'columns' and section and type(values) is int and 1 <= values <= 45:
            mutations.append(f'set number of text columns targetSetup number of columns {values}')
            accepted.append(group)
        elif group == 'wrap' and shape and type(values) is int and 0 <= values < len(_WRAP):
            mutations.append(f'set wrap type of wrap format of targetShape to {_WRAP[values]}')
            accepted.append(group)
        else:
            rejected.extend(group + '.' + k for k in values) if isinstance(values, dict) else rejected.append(group)
    return mutations, {'accepted': [] if rejected else accepted, 'rejected': rejected}


def compile_excel_patch(target, patch):
    from .macos_excel_session import (parse_target, _native_value, _number, _enum,
        _color, _quote, _sheet_name, _UNDERLINE, _CHART, _BORDER, _BORDER_STYLE,
        _BORDER_WEIGHT, _H_ALIGN, _V_ALIGN, validate_spreadsheet_formula)
    _, kind, _ = parse_target(target)
    operations=[]
    def add(key,prop,value): operations.append((key,f'set {prop} to {value}'))
    range_kind=kind in ('cell','range','selection')
    allowed={'value','formula','font','fill','number_format','horizontal_alignment','vertical_alignment','wrap_text','indent','row_height','column_width','borders'} if range_kind else ({'name','page_setup'} if kind=='sheet' else {'geometry','name','chart_type','chart_title'} if kind=='chart' else {'geometry','name','fill','line'})
    if set(patch)-allowed: raise ValueError('Unsupported Excel patch fields')
    for key,value in patch.items():
        if value is None: continue
        if key in ('value','formula'):
            if key=='formula':validate_spreadsheet_formula(value)
            add(key,key+' of obj',_native_value(value))
        elif key=='font':
            mapping={'name':'name','size':'font size','bold':'bold','italic':'italic','underline':'underline','strikethrough':'strikethrough','color':'color'}
            if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel font fields')
            for sub,item in value.items():
                val=_color(item) if sub=='color' else _number(item,1,409) if sub=='size' else _enum(item,_UNDERLINE) if sub=='underline' else _quote(item) if sub=='name' else _native_value(item) if type(item) is bool else None
                if val is None: raise ValueError('Invalid Excel font value')
                add('font.'+sub,mapping[sub]+' of font object of obj',val)
        elif key in ('fill','line'):
            if range_kind:
                if key!='fill' or not isinstance(value,dict) or set(value)-{'color'}: raise ValueError('Unsupported Excel range fill fields')
                if 'color' in value:add('fill.color','color of interior object of obj',_color(value['color']))
            else:
                mapping={'color':'fore color','back_color':'back color','transparency':'transparency','visible':'visible'}
                if key=='line':mapping['weight']='weight'
                if not isinstance(value,dict) or set(value)-set(mapping):raise ValueError('Unsupported Excel shape format fields')
                for sub,item in value.items():
                    if sub in ('color','back_color'):val=_color(item)
                    elif sub=='visible':
                        if type(item) is not bool:raise ValueError('Shape visibility must be boolean')
                        val=_native_value(item)
                    else:val=_number(item,0,1 if sub=='transparency' else 1000)
                    add(key+'.'+sub,mapping[sub]+' of '+key+' format of obj',val)
        elif key=='geometry':
            mapping={'left':'left position','top':'top','width':'width','height':'height','rotation':'rotation'}
            if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel geometry fields')
            for sub,item in value.items(): add('geometry.'+sub,mapping[sub]+' of obj',_number(item))
        elif key=='name': add(key,'name of obj',_sheet_name(value) if kind=='sheet' else _quote(value))
        elif key=='chart_type': add(key,'chart type of chart of obj',_enum(value,_CHART))
        elif key=='chart_title':
            operations.append((key,'set has title of chart of obj to true\nset chart title text of chart title of chart of obj to '+_quote(value)))
        elif key=='borders':
            if not isinstance(value,dict): raise ValueError('Invalid Excel borders')
            for edge,spec in value.items():
                try: native_edge=_enum(int(edge),_BORDER)
                except (TypeError,ValueError): raise ValueError('Invalid border edge') from None
                if not isinstance(spec,dict) or set(spec)-{'style','weight','color'}: raise ValueError('Unsupported border properties')
                commands=[f'set ownedBorder to get border obj which border {native_edge}']
                for sub,item in spec.items(): commands += [f'set '+{'style':'line style','weight':'weight','color':'color'}[sub]+' of ownedBorder to '+(_color(item) if sub=='color' else _enum(item,_BORDER_STYLE if sub=='style' else _BORDER_WEIGHT))]
                operations.append(('borders.'+str(edge),'\n'.join(commands)))
        elif key=='page_setup':
            mapping={'orientation':'page orientation','top_margin':'top margin','bottom_margin':'bottom margin','left_margin':'left margin','right_margin':'right margin','header_margin':'header margin','footer_margin':'footer margin','zoom':'zoom','fit_to_pages_wide':'fit to pages wide','fit_to_pages_tall':'fit to pages tall','print_area':'print area','print_title_rows':'print title rows','print_title_columns':'print title columns'}
            if not isinstance(value,dict) or set(value)-set(mapping): raise ValueError('Unsupported Excel page setup fields')
            for sub,item in value.items():
                val=_enum(item,{1:'portrait',2:'landscape'}) if sub=='orientation' else _quote(item) if sub.startswith('print_') else 'false' if sub=='zoom' and item is False else _number(item)
                add('page_setup.'+sub,mapping[sub]+' of page setup object of obj',val)
        else:
            prop={'number_format':'number format','horizontal_alignment':'horizontal alignment','vertical_alignment':'vertical alignment','wrap_text':'wrap text','indent':'indent level','row_height':'row height','column_width':'column width'}[key]
            if key=='number_format': val=_quote(value)
            elif key=='horizontal_alignment': val=_enum(value,_H_ALIGN)
            elif key=='vertical_alignment': val=_enum(value,_V_ALIGN)
            elif key=='wrap_text':
                if type(value) is not bool: raise ValueError('wrap_text must be boolean')
                val=_native_value(value)
            else: val=_number(value,0,255 if key=='column_width' else 15 if key=='indent' else 409)
            add(key,prop+' of obj',val)
    return operations, range_kind


def _word_position(position):
    if position in (None, 'start', 'end'): return
    if not isinstance(position, dict) or len(position) != 1:
        raise ValueError('Unsupported insertion position')
    key, value = next(iter(position.items()))
    if key == 'index':
        if type(value) is not int or value < 1: raise ValueError('Invalid paragraph index')
        return
    if key not in ('before', 'after') or word_target_kind(value) not in ('range','cell'):
        raise ValueError('Unsupported insertion anchor')


def materialize_word_structural(op):
    """Snapshot one-shot table rows for routing and execution to share.

    This is explicit request preparation, not a supports predicate. Row
    iterators remain unsupported, as in the original Word executor. Read at
    most the declared row count plus one: an excess row already proves the
    operation invalid, without exhausting an unbounded producer.
    """
    if not isinstance(op, dict) or op.get('op') != 'insert' or op.get('type') != 'table':
        return op
    props = op.get('props')
    if not isinstance(props, dict): return op
    data, rows = props.get('data'), props.get('rows', 2)
    if not data or isinstance(data, (list, tuple)) or type(rows) is not int or not 1 <= rows <= 1000:
        return op
    try:
        iterator = iter(data)
    except TypeError:
        return op  # Validation reports invalid data at the normal boundary.
    from itertools import islice
    return {**op, 'props': {**props, 'data': tuple(islice(iterator, rows + 1))}}


def validate_word_structural(op):
    from .macos_word_session import _number, apple_string
    verb, target, kind, props = op.get('op'), op.get('target',''), op.get('type'), op.get('props') or {}
    if not isinstance(props, dict): raise ValueError('Invalid insertion properties')
    if verb == 'insert':
        if op.get('parent','body') != 'body': raise ValueError('Unsupported Writer insert parent')
        _word_position(op.get('position','end'))
        allowed = {'paragraph':{'text','style','level'}, 'heading':{'text','style','level'},
                   'page_break':set(), 'table':{'rows','cols','data'},
                   'textbox':{'text','left','top','width','height'}, 'image':{'path'}}
        if kind not in allowed or set(props)-allowed[kind]: raise ValueError('Unsupported Writer insertion attributes')
        if kind in ('paragraph','heading'):
            apple_string(props.get('text',''))
            if props.get('style'): apple_string(props['style'])
            if kind == 'heading' and (type(props.get('level',1)) is not int or not 1 <= props.get('level',1) <= 9):
                raise ValueError('Invalid heading level')
        elif kind == 'table':
            rows, cols = props.get('rows',2), props.get('cols',2)
            if any(type(n) is not int or not 1 <= n <= 1000 for n in (rows,cols)) or rows*cols > 10000:
                raise ValueError('Invalid table dimensions')
            data = props.get('data') or []
            if not isinstance(data, (list, tuple)):
                raise ValueError('Table data must be materialized before validation')
            for index, row in enumerate(data):
                if index >= rows or not isinstance(row,(list,tuple)) or len(row)>cols:
                    raise ValueError('Table data exceeds dimensions')
                for value in row: apple_string(str(value))
        elif kind == 'textbox':
            apple_string(str(props.get('text', '')))
            for key in ('left','top','width','height'):
                if key in props: _number(props[key])
        elif kind == 'image':
            if set(props) != {'path'} or not isinstance(props['path'], (str,Path)):
                raise ValueError('Invalid image path')
        return
    if verb not in ('remove','move','clone'): raise ValueError('Unsupported structural verb')
    if re.fullmatch(r'(table|shape|inline_shape):[1-9]\d*', target):
        if verb != 'remove': raise ValueError('Mac Word object move/clone is not implemented')
        return
    if not target.startswith(('paragraph:','range:')): raise ValueError('Unsupported Writer structural target')
    word_target_kind(target)
    if verb != 'remove': _word_position(op.get('to','end'))


def validate_excel_structural(op):
    from .macos_excel_session import parse_target, _sheet_name, _native_value
    verb, kind, props = op.get('op'), op.get('type'), op.get('props') or {}
    if not isinstance(props,dict): raise ValueError('Invalid insertion properties')
    if verb == 'insert':
        if kind == 'sheet':
            if set(props)-{'name'}: raise ValueError('Unsupported worksheet insert fields')
            if 'name' in props: _sheet_name(props['name'])
            return
        if kind not in ('row','column') or set(props)-{'values'}: raise ValueError('Unsupported Excel insert type/fields')
        _, parent, ref = parse_target(op.get('parent'))
        if parent not in ('sheet','cell','range'): raise ValueError('Invalid Excel structural parent')
        position=op.get('position','end')
        if isinstance(position,dict):
            if set(position)!={'index'} or type(position['index']) is not int or not 1<=position['index']<=(1048576 if kind=='row' else 16384):
                raise ValueError('Invalid insert position')
        elif not ref and position!='end': raise ValueError('Invalid insert position')
        values=props.get('values')
        if values:
            if not isinstance(values,(list,tuple)) or len(values)>10000: raise ValueError('Invalid inserted values')
            for value in values: _native_value(value)
        return
    _, kind, _ = parse_target(op.get('target'))
    if verb in ('move','clone'):
        to=op.get('to')
        if kind=='sheet':
            if to is None or to=='end': return
            if not isinstance(to,dict) or len(to)!=1 or next(iter(to)) not in ('before','after'):
                raise ValueError('Invalid worksheet destination')
            value=next(iter(to.values()))
            if type(value) is not int or value<1: raise ValueError('Invalid worksheet destination')
            return
        axis=op.get('axis','row')
        if kind not in ('cell','range') or axis not in ('row','column') or not isinstance(to,dict) or set(to)!={'index'}:
            raise ValueError('Unsupported Excel move/clone target or destination')
        if type(to['index']) is not int or not 1<=to['index']<=(1048576 if axis=='row' else 16384):
            raise ValueError('Destination exceeds Excel bounds')
    elif verb=='remove':
        # Whether a sheet is freshly inserted and still empty is live session
        # state. Never infer that property from an arbitrary positional ID.
        if kind=='selection': raise ValueError('Unsupported Excel removal target')
        if kind in ('cell','range') and op.get('axis','row') not in ('row','column'):
            raise ValueError('Invalid Excel removal axis')
    else: raise ValueError('Unsupported structural verb')


def validate_powerpoint_structural(op):
    from .macos_powerpoint_session import _target, _number, _enum, LAYOUT, MacPowerPointSession, apple_string
    verb, props = op.get('op'), op.get('props') or {}
    if not isinstance(props,dict): raise ValueError('Invalid insertion properties')
    if verb=='insert':
        kind=op.get('type')
        for key, default in [('left',100),('top',100),('width',300),('height',100)]:
            _number(props.get(key,default),positive=key in ('width','height'))
        if kind=='slide':
            _enum(props.get('layout',12),LAYOUT)
            position=op.get('position','end')
            if position not in (None,'start','end'): MacPowerPointSession._position(position)
        else:
            _, parent = _target(op.get('parent') or '')
            if parent!='slide' or kind not in ('textbox','shape','image'): raise ValueError('Unsupported PowerPoint insert')
            if 'text' in props: apple_string(props['text'])
            if kind=='image' and (not isinstance(props.get('path'),(str,Path)) or Path(props['path']).suffix.lower() not in {'.png','.jpg','.jpeg','.gif','.bmp','.tiff','.tif'}):
                raise ValueError('Invalid image resource')
        return
    _,kind=_target(op.get('target') or '')
    if kind not in ('slide','shape'): raise ValueError('Structural target must be slide or shape')
    if verb=='remove': return
    if verb not in ('move','clone'): raise ValueError('Unsupported structural verb')
    to=op.get('to','end')
    if kind=='slide': MacPowerPointSession._slide_movement(to)
    else:
        # Shape destinations have a component-specific numeric slide contract.
        destination=to.get('slide',1) if isinstance(to,dict) else 1
        if type(destination) is not int or destination<1: raise ValueError('Invalid destination slide')


def supports_edit_ops(family, operations, *, platform, engine='msoffice'):
    from ..document_api import validate_op
    try:
        may_have_fresh_sheet = False
        for op in operations:
            if not validate_op(op,family).get('valid'): return False
            if platform != 'darwin':
                state = (classify_windows_set_op(family, op)
                         if op.get('op', 'set') == 'set'
                         else classify_windows_structural_op(
                             family, op, engine=engine))
                if state == 'unsupported':
                    return False
                continue
            if op.get('op','set')=='set':
                patch={key:value for key,value in op.items() if key not in ('op','target')}
                if family=='writer':
                    _, result=compile_word_patch(op['target'],patch)
                elif family=='sheet':
                    compile_excel_patch(op['target'],patch)
                    result={'rejected':[]}
                else:
                    from .macos_powerpoint_session import compile_patch
                    _,result=compile_patch(op['target'],**patch)
                    if op['target']=='selection' and result['rejected']:
                        # The live selection can be one shape or a text range.
                        # Compile a shape form only to test that capability;
                        # this does not assert any slide/shape exists or select it.
                        _,result=compile_patch('slide:1/shape:1',**patch)
                if result['rejected']: return False
            else:
                if family=='sheet':
                    if op.get('op')=='remove' and re.fullmatch(r'sheet:[1-9]\d*',op.get('target','')) and not may_have_fresh_sheet:
                        return False
                    if op.get('op')=='insert' and op.get('type')=='sheet':
                        may_have_fresh_sheet = True
                    elif op.get('op')!='remove' or re.fullmatch(r'sheet:[1-9]\d*',op.get('target','')):
                        may_have_fresh_sheet = False
                {'writer':validate_word_structural,'sheet':validate_excel_structural,
                 'slide':validate_powerpoint_structural}[family](op)
            if family=='sheet' and op.get('op','set')=='set': may_have_fresh_sheet = False
        return True
    except (ValueError,TypeError,KeyError,AttributeError,NotImplementedError,OverflowError):
        return False
