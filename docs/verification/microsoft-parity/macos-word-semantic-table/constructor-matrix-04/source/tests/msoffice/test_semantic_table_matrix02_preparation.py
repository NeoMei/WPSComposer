"""Offline diagnostic result interpretation and handler-binding contracts."""
import importlib
import pytest


def module():
    try:return importlib.import_module('fixtures.microsoft_parity.macos_word_semantic_table_matrix02')
    except ModuleNotFoundError:pytest.fail('Matrix02 diagnostic preparation is missing')


def rows(*,outcome='created',start=100,replace=True):
    return [['before','/owned.docx',12,19,1,'PREFIX REPLACE SUFFIX','OLD','PREFIX 中文😀\r','SUFFIX preserved 中文😀\r',True,'owned',True],
            ['attempt',outcome,0 if outcome=='created' else -2710,''],
            ['after',2 if outcome=='created' else 1,'PREFIX '+('REPLACE ' if replace else '')+'SUFFIX',False],
            ['table',1,30,45,1,1,'EXISTING TABLE 原样\r\x07'],
            *([['table',2,start,start+50,4,3,'\r\x07'*12]] if outcome=='created' else [])]


def test_appended_constructor_success_is_not_replacement_success():
    observed=module().assess_placement(rows(),4,3,19)
    assert observed=={'constructor_succeeded':True,'created_table_bounds':[100,150],
        'requested_insertion_observed':False,'replacement_removed':False}


def test_exact_middle_replacement_has_separate_positive_location_contract():
    observed=module().assess_placement(rows(start=12,replace=False),4,3,19)
    assert observed=={'constructor_succeeded':True,'created_table_bounds':[12,62],
        'requested_insertion_observed':True,'replacement_removed':True}


def test_ordinary_error_has_no_created_bounds_or_location_success():
    assert module().assess_placement(rows(outcome='ordinary-error'),4,3,19)=={
        'constructor_succeeded':False,'created_table_bounds':None,
        'requested_insertion_observed':False,'replacement_removed':False}


def test_collapsed_insertion_must_not_claim_text_replacement():
    observed=module().assess_placement(rows(start=12),4,3,12)
    assert observed['requested_insertion_observed'] is True
    assert observed['replacement_removed'] is None


def test_error_handler_uses_separate_catch_bindings_and_copies_only_caught_error():
    m=module();lines=m.commands('/owned.docx','sentinel','token',4,3,19,True)
    assert 'set diagCode to 0' in lines and 'set diagMessage to ""' in lines
    assert 'on error diagCaughtMessage number diagCaughtCode' in lines
    assert 'on error diagMessage number diagCode' not in lines
    assert 'set diagCode to diagCaughtCode' in lines and 'set diagMessage to diagCaughtMessage' in lines
    assert 'if diagCaughtCode is not -2710 then error diagCaughtMessage number diagCaughtCode' in lines
    assert sum('set diagTable to make new table' in line for line in lines)==1


def test_exactly_two_location_variants_reselect_after_activation():
    m=module()
    assert m.CASES==(('A-reselect-document-location',4,3,19,True,'boundDoc'),
                     ('B-reselect-range-location',4,3,19,True,'diagRange'))
    a=m.commands('/owned.docx','sentinel','token',4,3,19,True,location='boundDoc')
    b=m.commands('/owned.docx','sentinel','token',4,3,19,True,location='diagRange')
    assert a.index('activate object boundWindow')<a.index('set selection start of selection of boundWindow to 12')
    changes=[(x,y) for x,y in zip(a,b) if x!=y]
    assert len(a)==len(b) and len(changes)==1
    assert changes[0][0].replace('at boundDoc with','at diagRange with')==changes[0][1]
    assert 'DIAG_RANGE_CHANGED' in '\n'.join(a)
    assert 'DIAG_ACTIVE_OWNER_CHANGED' in '\n'.join(a)


def test_variant_location_is_closed():
    with pytest.raises(ValueError):
        module().commands('/owned.docx','sentinel','token',4,3,19,True,location='selection')
