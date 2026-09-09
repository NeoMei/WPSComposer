"""One range-representation change; successful native conversion is not assumed."""
import importlib
import pytest
from fixtures.microsoft_parity import macos_word_semantic_table_conversion_probe as previous


def mod():
    try:return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table_explicit_range_probe')
    except ModuleNotFoundError:pytest.fail('Explicit range preparation is missing')


def test_only_range_binding_changes_and_one_conversion_is_emitted():
    a=previous.commands('/owned.docx','sentinel','token',end=19)
    b=mod().commands('/owned.docx','sentinel','token')
    diff=[(x,y) for x,y in zip(a,b) if x!=y]
    assert len(a)==len(b) and diff==[('set diagRange to text object of diagSelection','set diagRange to create range boundDoc start 12 end 19')]
    assert sum('convert to table diagRange' in line for line in b)==1
    assert 'set selection end of selection of boundWindow to 19' in b


def rows():
    return [['before','/owned.docx',12,19,1,'body','EXISTING TABLE 原样','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r',True,'owned',True],
            ['attempt','created',0,''],['after',2,'body after',False],
            ['table',1,12,23,1,3,'REPLACE'],['table',2,46,65,1,1,'EXISTING TABLE 原样'],
            ['conversion',True,12,23,1,3,'REPLACE','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r'],
            ['surroundings','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r','EXISTING TABLE 原样']]


def test_actual_one_by_three_result_cannot_claim_requested_four_by_three():
    r=rows();assert mod().valid(r,'/owned.docx')
    actual=mod().assess(r)
    assert actual['returned_dimensions']==[1,3]
    assert actual['requested_dimensions_observed'] is False
    assert actual['document_table_ordinal']==1
    assert actual['table_count_delta']==1
    assert actual['marker_in_new_table'] is True


def test_document_ordinal_mismatch_is_not_a_complete_ack():
    r=rows();r[3][2]=13
    assert not mod().valid(r,'/owned.docx')


def test_requested_dimensions_are_positive_only_when_observed():
    r=rows();r[3][4]=4;r[-2][4]=4
    assert mod().assess(r)['requested_dimensions_observed'] is True
