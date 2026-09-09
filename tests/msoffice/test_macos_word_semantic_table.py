"""Pure table construction contracts; these tests do not execute AppleEvents."""
from __future__ import annotations

import ast
import importlib
import subprocess
import unicodedata
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.writer import NativeWriterObjectError


def implementation():
    try:
        return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_semantic_table')
    except ModuleNotFoundError:
        pytest.fail('Pure semantic table command builder is missing')


def arguments(**overrides):
    result = dict(start=8, end=11, headers=['编号', '内容'],
                  rows=[[None, 0], [True, '汉字😀\nlong narrative']],
                  alignments=['center', 'left'], cell_indent_pt=3.5,
                  repeat_header=True, allow_row_split=False,
                  available_width_pt=450,
                  border_spec=dict(top=1.5, left=0, bottom=.75, right=.5,
                                   insideHorizontal=2, insideVertical=0, headerBottom=1.5),
                  merges=[dict(top=2, left=1, bottom=3, right=1),
                          dict(top=1, left=1, bottom=1, right=2)])
    result.update(overrides)
    return result


@pytest.fixture(scope='module')
def frozen():
    source = subprocess.check_output(['git', 'show', '6dd3a00:skills/WPSComposer/scripts/writer.py'], text=True)
    tree = ast.parse(source)
    cls = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == 'WriterComposer')
    cls.bases = []
    cls.body = [node for node in cls.body if isinstance(node, ast.FunctionDef) and node.name in {
        '_create_native_table', '_apply_native_table_borders', '_native_border_width',
        '_fit_native_table_to_body', '_current_section_page_setup'}]
    helpers = [node for node in tree.body if isinstance(node, ast.FunctionDef)
               and node.name in {'_content_column_widths', '_visual_text_width'}]
    namespace = dict(NativeWriterObjectError=NativeWriterObjectError, unicodedata=unicodedata)
    exec(compile(ast.Module(body=helpers + [cls], type_ignores=[]), '<frozen-6dd3a00>', 'exec'), namespace)
    return namespace['WriterComposer']


def oracle(frozen, args):
    """Record real frozen COM-property requests, without importing pywin32."""
    state = dict(cells={}, widths=[], borders={}, merges=[])
    class Border:
        def __init__(self, key): self.key = key
        def __setattr__(self, key, value):
            if key == 'key': object.__setattr__(self, key, value)
            else: state['borders'].setdefault(self.key, {})[key] = value
    class Cell:
        def __init__(self, row, column):
            self.key = row, column
            self.Range = SimpleNamespace(ParagraphFormat=SimpleNamespace())
            state['cells'][self.key] = self.Range
        def Merge(self, other): state['merges'].append((*self.key, *other.key))
    class Rows:
        def __call__(self, row): return self
        def Borders(self, border): return Border(('header', border))
    class Column:
        def SetWidth(self, width, rule): state['widths'].append(width)
    class Table:
        def __init__(self):
            self.Rows = Rows()
            self.Range = SimpleNamespace(End=99)
        def Cell(self, row, column):
            # Native lookup does not overwrite existing cell content during merges.
            if (row, column) not in cells: cells[row, column] = Cell(row, column)
            return cells[row, column]
        def Columns(self, index): return Column()
        def AutoFitBehavior(self, mode): state['autofit_behavior'] = mode
        def Borders(self, border): return Border(('table', border))
    cells = {}
    table = Table()
    def add(rng, rows, cols):
        state['create'] = (rng.Start, rng.End, rows, cols)
        return table
    class Selection:
        Range = SimpleNamespace(Start=args['start'], End=args['end'])
        def SetRange(self, start, end): state['cursor'] = (start, end)
        def TypeParagraph(self): state['paragraph'] = True
    writer = frozen()
    setup = SimpleNamespace(PageWidth=args['available_width_pt'] or 0, LeftMargin=0, RightMargin=0)
    if args['available_width_pt'] is None: setup = SimpleNamespace()
    writer._doc = SimpleNamespace(Tables=SimpleNamespace(Add=add), PageSetup=setup)
    writer.selection = Selection()
    writer._create_native_table(*(args[key] for key in ('headers', 'rows', 'alignments', 'border_spec',
                                'repeat_header', 'allow_row_split', 'cell_indent_pt', 'merges')))
    state['cells'] = {key: (value.Text, vars(value.ParagraphFormat)) for key, value in state['cells'].items()}
    state['repeat'] = getattr(table.Rows, 'HeadingFormat', None)
    state['split'] = table.Rows.AllowBreakAcrossPages
    state['fit'] = getattr(table, 'AllowAutoFit', None)
    return state


@pytest.mark.parametrize('overrides', [{}, {'repeat_header': False, 'allow_row_split': True},
    {'headers': ['x'], 'rows': [[]], 'alignments': ['right'], 'merges': []},
    {'available_width_pt': 20}, {'available_width_pt': None},
    {'headers': ['x','y'], 'rows': [['', False], ['very long text '*80, '短']], 'merges': []}])
def test_command_semantics_match_frozen_table_factory(frozen, overrides):
    args = arguments(**overrides)
    expected = oracle(frozen, args)
    commands = implementation().build_semantic_table_commands(**args)
    by_action = lambda action: [command.args for command in commands if command.action == action]
    assert by_action('create') == [expected['create']]
    cells = {(row, col): (text, dict(FirstLineIndent=indent, LeftIndent=0.0, RightIndent=0.0, Alignment=align))
             for row, col, text, indent, align in by_action('cell')}
    assert cells == expected['cells']
    assert [width for _, width in by_action('column_width')] == expected['widths']
    assert by_action('repeat_header') == ([(True,)] if expected['repeat'] else [])
    assert by_action('row_split') == [(bool(expected['split']),)]
    assert by_action('merge') == expected['merges']
    assert by_action('finish') == [()]
    borders = {(scope, position): dict(LineStyle=style, **({'LineWidth': width} if width is not None else {}))
               for scope, position, style, width in by_action('border')}
    assert borders == expected['borders']
    assert bool(by_action('fit')) == (expected['fit'] is not None)


@pytest.mark.parametrize('field,value,code', [
    ('alignments', ['evil', 'left'], 'TABLE_STYLE_APPLY_FAILED'),
    ('cell_indent_pt', 'no', 'TABLE_STYLE_APPLY_FAILED'),
    ('border_spec', {}, 'TABLE_STYLE_APPLY_FAILED'),
    ('merges', [dict(top=1)], 'TABLE_MERGE_APPLY_FAILED')])
def test_invalid_descriptor_preserves_failure_phase_and_prior_commands(field, value, code):
    iterator = implementation().iter_semantic_table_commands(**arguments(**{field: value}))
    seen = []
    with pytest.raises(NativeWriterObjectError) as error:
        while True: seen.append(next(iterator))
    assert error.value.code == code
    assert seen[0].action == 'create'
    assert not any(command.action == 'finish' for command in seen)
    if code == 'TABLE_MERGE_APPLY_FAILED': assert any(command.action == 'border' for command in seen)


@pytest.mark.parametrize('start,end', [(True, 3), (-1, 3), (4, 3), (1, '4')])
def test_bound_coordinates_reject_unsafe_or_ambiguous_ranges(start, end):
    with pytest.raises(ValueError):
        implementation().build_semantic_table_commands(**arguments(start=start, end=end))


def test_emission_uses_bound_range_and_preserves_text_and_cursor_contract():
    module = implementation()
    commands = module.build_semantic_table_commands(**arguments(headers=['"\\😀\r\t', 'B'], rows=[], merges=[]))
    script = '\n'.join(module.render_semantic_table_commands(commands))
    assert 'create range boundDoc start 8 end 11' in script
    assert 'set semanticTable to make new table at boundDoc' in script
    assert 'set semanticTableEnd to end of content of text object of semanticTable' in script
    assert 'create range boundDoc start semanticTableEnd end semanticTableEnd' in script
    assert 'set content of semanticTail to return' in script
    assert 'set selection start of semanticSelection to semanticTableEnd + 1' in script
    assert 'set selection end of semanticSelection to semanticTableEnd + 1' in script
    assert '\\"\\\\😀" & return & "" & tab & "' in script
    assert 'line width150 point' in script and 'line width75 point' in script and 'line width25 point' in script
    assert 'get border row 1 of semanticTable which border border bottom' in script
    for unrelated in ('padding', 'font size', 'font object', 'space before', 'space after', 'character unit', 'active document', 'ownedDoc'):
        assert unrelated not in script
    assert 'if semanticError is -1712 or semanticError is -609 or semanticError is -128 then error semanticMessage number semanticError' in script
    assert [command.error_code for command in commands if command.action in ('create', 'cell', 'merge', 'finish')] == [
        'TABLE_INSERT_FAILED', 'TABLE_STYLE_APPLY_FAILED', 'TABLE_STYLE_APPLY_FAILED', 'TABLE_INSERT_FAILED']


def test_nonfinite_measurement_and_merge_script_injection_fail_closed():
    module = implementation()
    for overrides, code in [({'cell_indent_pt': float('nan')}, 'TABLE_STYLE_APPLY_FAILED'),
                            ({'merges': [dict(top='1\nquit', left=1, bottom=2, right=1)]}, 'TABLE_MERGE_APPLY_FAILED')]:
        with pytest.raises(NativeWriterObjectError) as error:
            module.build_semantic_table_commands(**arguments(**overrides))
        assert error.value.code == code


def test_control_text_remains_literal_data_like_frozen_cell_assignment(frozen):
    args = arguments(headers=['before\x07after\x0b', 'B'], rows=[], merges=[])
    expected = oracle(frozen, args)
    module = implementation()
    commands = module.build_semantic_table_commands(**args)
    assert next(command.args[2] for command in commands if command.action == 'cell') == expected['cells'][(1, 1)][0]
    script = '\n'.join(module.render_semantic_table_commands(commands))
    assert '(character id 7)' in script and '(character id 11)' in script
    assert '\x07' not in script and '\x0b' not in script
