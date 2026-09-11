from pathlib import Path
import importlib.util
import pytest

def m():
 p=Path(__file__).with_name('style-install-refresh-v2.py');s=importlib.util.spec_from_file_location('install_v2',p);v=importlib.util.module_from_spec(s);s.loader.exec_module(v);return v

def test_parent_setup_discards_make_return_and_resolves_exact_name():
 v=m();lines=v.parent_setup_commands()
 assert lines[0].startswith('make new Word style')
 assert not any('to make new Word style' in line for line in lines)
 assert 'set base style of Word style "WPSC Install Parent" of boundDoc to style normal' in lines
 assert 'set paragraph format left indent of paragraph format of Word style "WPSC Install Parent" of boundDoc to 17' in lines

def test_source_setup_uses_exact_lhs_and_text_hypothesis_not_object_rhs():
 v=m();assert v.setup_commands()[0]=='set base style of Word style (style heading1) of boundDoc to "WPSC Install Parent"'
 assert 'set base style of sourceStyle to parentStyle' not in '\n'.join(v.setup_commands())

def test_typed_chain_checks_both_stage_independent_readbacks():
 v=m();a=['parent','Normal','标题 1',v.PARENT,'Normal','Normal',17,0,True,False]
 b=['source','Normal','标题 1',v.PARENT,'Normal',v.PARENT,17,17,True,False]
 assert v.chain_valid([a],'parent') and v.chain_valid([b],'source')
 for index,value in [(4,'标题 1'),(5,'Normal'),(6,True),(7,0),(8,1),(9,0),(3,'Wrong')]:
  bad=list(b);bad[index]=value;assert not v.chain_valid([bad],'source')

def test_original_native_failed_chain_is_rejected():
 v=m();styles,_=v.base.package(Path(__file__).with_name('style-install-refresh-01')/'setup.docx')
 assert not v.parent_chain_xml_valid(styles)
 with pytest.raises(ValueError,match='Unsupported source chain'):v.expected_clone(styles)

def test_original_install_oracles_retained():
 old=Path(__file__).with_name('style-install-refresh.py').read_text();new=Path(__file__).with_name('style-install-refresh-v2.py').read_text()
 for name,next_name in [('expected_clone','target_style'),('other_styles_preserved','clone_matches'),('clear_commands','carrier_valid'),('source_change_valid','run')]:
  assert old[old.index('def '+name):old.index('def '+next_name)]==new[new.index('def '+name):new.index('def '+next_name)]
