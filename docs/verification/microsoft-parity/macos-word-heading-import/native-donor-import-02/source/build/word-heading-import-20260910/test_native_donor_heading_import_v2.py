from pathlib import Path
import importlib.util
import pytest
PATH=Path(__file__).with_name('native-donor-heading-import-v2.py')
def module():
    assert PATH.exists(),'corrected attributed heading probe absent'
    spec=importlib.util.spec_from_file_location('donor_v2',PATH)
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def test_style_scalars_keep_all_dimensions_and_exclude_only_unsupported_tabs():
    m=module();code='\n'.join(m.style_scalar_readback())
    labels=m.style_scalar_labels()
    assert labels==[x for x in m.heading.expected_labels('detached-properties') if x!='paragraph:tabs']
    assert 'count tab stops' not in code and 'every tab stop' not in code
    assert all(m.apple_string(label) in code for label in labels)
    for prop in ['font size','paragraph format left indent','background pattern color','outside line style']:
        assert prop in code
    assert 'clone-base-normal' in code

def test_effective_paragraph_tabs_preserve_proven_identity_and_all_properties():
    m=module();code='\n'.join(m.effective_tab_readback())
    assert code=='\n'.join(m.tab_diagnostic.tab_commands('paragraph'))
    assert 'WPSC_TAB_PARAGRAPH_STYLE' in code
    assert 'paragraph 1 of boundDoc' in code and 'paragraph 2 of boundDoc' in code
    assert all(p in code for p in ['tab stop position','alignment of ownTab','tab leader','custom tab'])

def test_style_scalar_oracle_rejects_missing_or_false_dimension():
    m=module();rows=[[label,True] for label in m.style_scalar_labels()]
    assert m.style_scalars_valid(rows)
    assert not m.style_scalars_valid(rows[:-1])
    rows[4][1]=False;assert not m.style_scalars_valid(rows)

def test_paragraph_oracle_accepts_complete_nonempty_and_rejects_changed_tabs():
    m=module();rows=[['tabs','paragraph','标题 1',m.CLONE,1,1],
       ['source',1,40,'align tab right','tab leader spaces',True],
       ['clone',1,40,'align tab right','tab leader spaces',True]]
    assert m.tab_diagnostic.tabs_valid(rows,'paragraph')
    rows[2][4]='tab leader dots';assert not m.tab_diagnostic.tabs_valid(rows,'paragraph')

def test_missing_execute_never_runs_native(tmp_path,monkeypatch):
    m=module();monkeypatch.setattr(m,'run',lambda *a:pytest.fail('native called'))
    assert m.main(['--output',str(tmp_path/'absent')])==2
    assert not (tmp_path/'absent').exists()
