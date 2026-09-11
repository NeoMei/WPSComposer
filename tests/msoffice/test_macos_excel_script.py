"""Closed-plan Excel compiler validation, without launching native applications."""
from __future__ import annotations

import importlib
import math
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.generation_plan import GenerationOperation, GenerationPlan


def compiler():
    return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_excel_script')


def plan(*ops):
    return GenerationPlan('spreadsheet', tuple(GenerationOperation(name, args) for name, args in ops))


def table_plan():
    return plan(('sheet.reset', {}), ('sheet.rename', {'index': 1, 'name': 'Data'}),
                ('sheet.write_table', {'startRow': 2, 'startCol': 2,
                 'values': [['名称', 'Value'], ['Alpha', 10], ['Total', '=SUM(C3:C3)']],
                 'headerBold': True, 'headerShade': '#4472C4', 'headerFontColor': '#FFFFFF', 'fontSize': 12}),
                ('sheet.add', {'name': 'Summary'}), ('sheet.select', {'index': 1}),
                ('sheet.set_column_width', {'column': 'B:C', 'width': 22}), ('sheet.autofit', {}))


def test_compiles_native_owned_workbook_and_all_existing_plan_operations(tmp_path):
    source = compiler().compile_plan(table_plan(), {}, tmp_path / 'native.xlsx', tmp_path / 'native.pdf')
    assert 'make new workbook' in source
    assert 'worksheet 1 of ownedBook' in source
    assert 'range "B2:C4"' in source
    assert '=SUM(C3:C3)' in source
    assert 'font size of font object' in source
    assert '{68, 114, 196}' in source
    assert 'column width of range "B:C"' in source
    assert 'autofit entire column of used range of ownedSheet' in source
    assert 'Excel XML file format' in source and 'PDF file format' in source
    assert 'close ownedBook saving no' in source
    assert source.rstrip().endswith('return "WPSCOMPOSER_MS_OFFICE_OK:spreadsheet"')
    assert 'quit' not in source.lower()
    assert 'do visual basic' not in source.lower()
    assert 'display alerts' not in source.lower()


def test_validation_happens_even_for_directly_constructed_unknown_operation(tmp_path):
    with pytest.raises(ValueError):
        compiler().compile_plan(plan(('sheet.run_macro', {'name': 'bad'})), {}, tmp_path / 'x.xlsx')


@pytest.mark.parametrize('ops', [
    [('sheet.select', {'index': 2})],
    [('sheet.reset', {}), ('sheet.reset', {})],
    [('sheet.add', {'name': 'A'}), ('sheet.rename', {'index': 1, 'name': 'a'})],
    [('sheet.rename', {'index': 1, 'name': 'bad/name'})],
    [('sheet.rename', {'index': 1, 'name': "'bad"})],
    [('sheet.rename', {'index': 1, 'name': 'x' * 32})],
    [('sheet.set_column_width', {'column': 'A1', 'width': 10})],
    [('sheet.set_column_width', {'column': 'XFE', 'width': 10})],
    [('sheet.set_column_width', {'column': 'C:A', 'width': 10})],
    [('sheet.set_column_width', {'column': 'A', 'width': 256})],
    [('sheet.write_table', {'startRow': 1048576, 'startCol': 1, 'values': [[1], [2]]})],
    [('sheet.write_table', {'startRow': 1, 'startCol': 16384, 'values': [[1, 2]]})],
    [('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [['x' * 32768]]})],
])
def test_rejects_unexecutable_excel_arguments_before_launch(tmp_path, ops):
    with pytest.raises(ValueError):
        compiler().compile_plan(plan(*ops), {}, tmp_path / 'x.xlsx')


@pytest.mark.parametrize('timeout', [0, -1, True, math.inf, math.nan, 901])
def test_invalid_timeout(tmp_path, timeout):
    with pytest.raises(ValueError):
        compiler().compile_plan(table_plan(), {}, tmp_path / 'x.xlsx', timeout=timeout)


def test_paths_and_control_characters_are_validated_and_data_is_quoted(tmp_path):
    mod = compiler()
    with pytest.raises(ValueError):
        mod.compile_plan(table_plan(), {}, Path('relative.xlsx'))
    with pytest.raises(ValueError):
        mod.compile_plan(table_plan(), {}, tmp_path / 'x.docx')
    with pytest.raises(ValueError):
        mod.compile_plan(plan(('sheet.add', {'name': '\x00bad'})), {}, tmp_path / 'x.xlsx')
    text = 'quote" and \\ slash\nline\ttab'
    source = mod.compile_plan(plan(('sheet.write_table', {'startRow': 1, 'startCol': 1, 'values': [[text]]})), {}, tmp_path / 'quote"file.xlsx')
    assert 'quote\\" and \\\\ slash" & linefeed & "line" & tab & "tab' in source
    assert 'quote\\"file.xlsx' in source


def test_conversion_pins_source_rejects_existing_open_name_and_saves_pdf_only(tmp_path):
    source = compiler().compile_conversion(tmp_path / 'source.xlsx', tmp_path / 'result.pdf')
    assert 'open workbook workbook file name' in source
    assert 'with read only and editable' in source
    assert 'full name of ownedBook' in source
    assert 'Source workbook is already open' in source
    assert 'PDF file format' in source
    assert 'Excel XML file format' not in source
    assert source.rstrip().endswith('return "WPSCOMPOSER_MS_OFFICE_OK:spreadsheet"')


def test_conversion_rejects_same_path_and_non_excel_source(tmp_path):
    mod = compiler()
    with pytest.raises(ValueError):
        mod.compile_conversion(tmp_path / 'source.pdf', tmp_path / 'source.pdf')
    with pytest.raises(ValueError):
        mod.compile_conversion(Path('source.xlsx'), tmp_path / 'result.pdf')


def test_automatic_external_formula_rejected_before_native_emission():
    from skills.WPSComposer.scripts.msoffice.macos_excel_script import _cell
    import pytest
    with pytest.raises(ValueError):
        _cell('=RTD("server",,"topic")')
    assert _cell('=SUM(A1:A3)') == '"=SUM(A1:A3)"'
