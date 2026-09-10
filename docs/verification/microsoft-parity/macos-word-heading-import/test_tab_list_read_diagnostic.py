from pathlib import Path
import importlib.util
import pytest
PATH=Path(__file__).with_name('tab-list-read-diagnostic.py')
def module():
    assert PATH.exists(),'tab list diagnostic absent'
    spec=importlib.util.spec_from_file_location('tab_list_diag',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def sample(m):
    return [['tabs','style','标题 1',m.CLONE,1,1],['source',1,40.0,'align tab right','tab leader spaces',True],['clone',1,40.0,'align tab right','tab leader spaces',True]]

def test_explicit_list_get_precedes_local_count_and_keeps_all_properties():
    m=module();code='\n'.join(m.tab_commands('style'))
    assert code.index('get every tab stop of sourceParagraph')<code.index('count sourceTabs')
    assert 'count tab stops of sourceParagraph' not in code
    assert 'class of sourceTabs is not list' in code
    assert all(p in code for p in ['alignment of ownTab','tab stop position of ownTab','tab leader of ownTab','custom tab of ownTab'])

def test_style_and_effective_paragraph_modes_remain_distinct():
    m=module();code='\n'.join(m.tab_commands('paragraph'))
    assert 'paragraph 1 of sourceRange' in code and 'paragraph 1 of cloneRange' in code
    assert 'count tab stops of sourceParagraph' in code
    assert 'WPSC_TAB_PARAGRAPH_STYLE' in code

def test_valid_nonempty_and_explicit_empty_lists():
    m=module();assert m.tabs_valid(sample(m),'style')
    assert m.tabs_valid([['tabs','style','标题 1',m.CLONE,0,0]],'style')

@pytest.mark.parametrize('change',['missing','wrong_count','wrong_type','property','mode','index'])
def test_tab_observation_rejects_malformed_or_different_rows(change):
    m=module();rows=sample(m)
    if change=='missing':rows.pop()
    elif change=='wrong_count':rows[0][4]=0
    elif change=='wrong_type':rows[1][-1]='true'
    elif change=='property':rows[2][2]=41
    elif change=='mode':rows[0][1]='paragraph'
    else:rows[2][1]=2
    assert not m.tabs_valid(rows,'style')

def test_no_execute_guard(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
