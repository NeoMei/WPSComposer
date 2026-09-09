"""No native execution: strict range-conversion proposal and evidence checks."""
import importlib
import pytest


def mod():
    try:return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table_conversion_probe')
    except ModuleNotFoundError:pytest.fail('Conversion preparation is missing')


@pytest.mark.parametrize('end',[12,19])
def test_conversion_uses_exact_range_once_and_keeps_owner_guards(end):
    lines=mod().commands('/owned.docx','sentinel','token',end=end)
    s='\n'.join(lines)
    assert f'set selection end of selection of boundWindow to {end}' in lines
    assert f'end of content of diagRange is not {end}' in s
    assert 'set diagTable to convert to table diagRange number of rows 4 number of columns 3' in lines
    assert s.count('convert to table diagRange')==1 and 'make new table' not in s
    for guard in ('DIAG_ACTIVE_OWNER_CHANGED','DIAG_SELECTION_OWNER_CHANGED','QUALITY_BOUND_PATH_CHANGED','DIAG_RANGE_CHANGED'):
        assert guard in s
    assert 'on error diagCaughtMessage number diagCaughtCode' in s
    assert 'set conversionBefore to create range boundDoc start 0 end conversionStart' in lines
    assert 'set conversionAfter to create range boundDoc start conversionEnd end (end of content of text object of boundDoc)' in lines


def test_only_collapsed_and_noncollapsed_seed_ranges_are_allowed():
    with pytest.raises(ValueError):mod().commands('/owned.docx','sentinel','token',end=20)


def test_marker_inside_converted_table_is_separate_from_outside_retention():
    row=['conversion',True,12,28,4,3,'REPLACE\r\x07','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r']
    assessment=mod().assess(row)
    assert assessment['starts_at_requested_range'] is True
    assert assessment['marker_in_new_table'] is True
    assert assessment['marker_outside_new_table'] is False


def test_wrong_location_with_marker_outside_never_meets_range_location():
    row=['conversion',True,20,36,4,3,'\r\x07','PREFIX 中文😀\rREPLACE\r','SUFFIX preserved 中文😀\r']
    assessment=mod().assess(row)
    assert assessment['starts_at_requested_range'] is False
    assert assessment['marker_in_new_table'] is False
    assert assessment['marker_outside_new_table'] is True


def test_ordinary_error_cannot_claim_conversion_or_location():
    assert mod().assess(['conversion',False,-1,-1,0,0,'','',''])=={
        'conversion_succeeded':False,'new_table_bounds':None,'starts_at_requested_range':False,
        'marker_in_new_table':False,'marker_outside_new_table':None}


def ack_rows():
    return [['before','/owned.docx',12,19,1,'body','EXISTING TABLE 原样','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r',True,'owned',True],
            ['attempt','created',0,''],['after',2,'body after',False],
            ['table',1,12,28,4,3,'REPLACE'],['table',2,59,78,1,1,'EXISTING TABLE 原样'],
            ['conversion',True,12,28,4,3,'REPLACE','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r'],
            ['surroundings','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r','EXISTING TABLE 原样']]


def test_complete_ack_preserves_verified_surroundings():
    rows=ack_rows()
    assert mod().valid(rows,'/owned.docx',end=19)
    rows[-1][1]='changed prefix'
    assert not mod().valid(rows,'/owned.docx',end=19)  # Root requires preservation as an ACK gate


@pytest.mark.parametrize('mutation',['missing_surroundings','boolean_bounds','foreign_path','unrecognized_error','false_success'])
def test_incomplete_or_mistyped_ack_stops(mutation):
    rows=ack_rows()
    if mutation=='missing_surroundings':rows.pop()
    if mutation=='boolean_bounds':rows[-2][2]=True
    if mutation=='foreign_path':rows[0][1]='/foreign.docx'
    if mutation=='unrecognized_error':rows[1]=['attempt','ordinary-error',-1700,'']
    if mutation=='false_success':rows[1]=['attempt','ordinary-error',-2710,'']
    assert not mod().valid(rows,'/owned.docx',end=19)


@pytest.mark.parametrize('mutation',['returned_bounds','returned_text','returned_dimensions','prefix_changed','suffix_changed','oldtable_changed'])
def test_returned_table_must_match_one_document_table_and_preserve_surroundings(mutation):
    rows=ack_rows()
    if mutation=='returned_bounds':rows[-2][2]=13
    if mutation=='returned_text':rows[-2][6]='different'
    if mutation=='returned_dimensions':rows[-2][4]=3
    if mutation=='prefix_changed':rows[-1][1]='changed'
    if mutation=='suffix_changed':rows[-1][2]='changed'
    if mutation=='oldtable_changed':rows[-1][3]='changed'
    assert not mod().valid(rows,'/owned.docx',end=19)
