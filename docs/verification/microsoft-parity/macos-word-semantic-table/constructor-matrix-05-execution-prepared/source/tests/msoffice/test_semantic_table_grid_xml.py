"""Strict outside XML splice: preserve blank paragraphs and layout properties."""
import copy
import importlib
import xml.etree.ElementTree as ET
import pytest
from fixtures.microsoft_parity.macos_word_semantic_table_grid_probe import TOKENS
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def mod():
    try:return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table_grid_xml')
    except ModuleNotFoundError:pytest.fail('Grid XML gate missing')


def pair():
    old=ET.Element(W+'document');body=ET.SubElement(old,W+'body')
    def p(parent,text=''):
        q=ET.SubElement(parent,W+'p')
        if text:ET.SubElement(ET.SubElement(q,W+'r'),W+'t').text=text
        return q
    p(body,'PREFIX 中文😀');p(body,'REPLACE');p(body,'SUFFIX preserved 中文😀')
    table=ET.SubElement(body,W+'tbl');cell=ET.SubElement(ET.SubElement(table,W+'tr'),W+'tc');p(cell,'EXISTING TABLE 原样')
    p(body);ET.SubElement(body,W+'sectPr')
    new=copy.deepcopy(old);body=new.find(W+'body');body.remove(body[1]);table=ET.Element(W+'tbl');body.insert(1,table)
    grid=ET.SubElement(table,W+'tblGrid')
    for _ in range(3):ET.SubElement(grid,W+'gridCol',{W+'w':'2000'})
    for r in range(4):
        row=ET.SubElement(table,W+'tr')
        for c in range(3):p(ET.SubElement(row,W+'tc'),TOKENS[r*3+c])
    body.insert(2,ET.Element(W+'p'))
    return old,new


def test_exact_grid_and_only_one_terminal_paragraph_pass():
    old,new=pair();assert all(mod().checks(ET.tostring(old),ET.tostring(new)).values())


@pytest.mark.parametrize('mutation',['extra_empty','outside_pagebreak','outside_section','tail_pagebreak','tail_section','missing_old_blank','placeholder_wrong_cell','one_row'])
def test_outside_layout_or_grid_mutations_fail(mutation):
    old,new=pair();body=new.find(W+'body')
    if mutation=='extra_empty':body.insert(3,ET.Element(W+'p'))
    if mutation=='outside_pagebreak':ET.SubElement(ET.SubElement(body[-2],W+'pPr'),W+'pageBreakBefore')
    if mutation=='outside_section':ET.SubElement(ET.SubElement(body[-2],W+'pPr'),W+'sectPr')
    if mutation=='tail_pagebreak':ET.SubElement(ET.SubElement(body[2],W+'pPr'),W+'pageBreakBefore')
    if mutation=='tail_section':ET.SubElement(ET.SubElement(body[2],W+'pPr'),W+'sectPr')
    if mutation=='missing_old_blank':body.remove(body[-2])
    if mutation=='placeholder_wrong_cell':body[1].find('.//'+W+'t').text='wrong'
    if mutation=='one_row':
        for row in body[1].findall(W+'tr')[1:]:body[1].remove(row)
    assert not all(mod().checks(ET.tostring(old),ET.tostring(new)).values())
