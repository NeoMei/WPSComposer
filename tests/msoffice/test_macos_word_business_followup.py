"""Second direct Writer slice; compilation is distinct from native execution."""
import inspect
from pathlib import Path
import subprocess
import sys
import pytest
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession,NativeWordCapabilityError
from skills.WPSComposer.scripts.writer import WriterComposer

METHODS=['ensure_styles','ensure_heading_styles','apply_heading_text_color','add_code_lines','set_columns','add_section','set_page_number_in_footer','compact_terminal_paragraph','reset','refresh_indexes','update_fields','finalize_fields','refresh_fields']
@pytest.fixture
def session(monkeypatch):
    s=MacWordSession();calls=[]
    monkeypatch.setattr(s,'_execute',lambda lines,**kw:calls.append(lines) or [['stats',3,1]])
    return s,calls

@pytest.mark.parametrize('name',METHODS)
def test_exact_signatures(name):
    assert inspect.signature(getattr(MacWordSession,name))==inspect.signature(getattr(WriterComposer,name))

def test_styles_normalize_camelcase_and_use_document_local_native_styles(session):
    s,calls=session
    assert s.ensure_styles({'Custom':{'fontName':'仿宋','fontNameAscii':'Times New Roman','fontSize':12,'indentFirst':24,'keepWithNext':True,'based_on':'Normal'}}) is None
    script='\n'.join(calls[0]);assert 'make new Word style at boundDoc' in script
    assert 'ascii name' in script and 'first line indent' in script
    assert 'base style' in script and 'to 24' in script
    assert len(calls)==1

@pytest.mark.parametrize('styles',[
 {'Okay':{'font_size':12},'Bad':{'type':'character'}},
 {'Okay':{'font_size':12},'Bad':{'unknown':True}},
 {'Bad':{'fontSize':-2}}, {'Bad':{'outline_level':99}},
 {'Bad':{'font_size':12,'fontSize':14}},
])
def test_whole_style_batch_rejected_before_mutation(session,styles):
    s,calls=session
    with pytest.raises((ValueError,TypeError)):s.ensure_styles(styles)
    assert not calls

def test_heading_styles_and_color_bind_only_heading_one_through_six(session):
    s,calls=session;s.ensure_heading_styles({1:{'font_size':16,'bold':True},6:{'font_size':12}});s.apply_heading_text_color('#123456')
    script='\n'.join(calls[0]);assert 'style heading1' in script and 'style heading6' in script
    assert 'make new Word style' not in script
    assert sum('set color of font object' in line for line in calls[1])==6

def test_code_lines_compile_all_inputs_before_one_mutation(session):
    s,calls=session
    assert s.add_code_lines(['x','']) is None
    script='\n'.join(calls[0]);assert 'Source Code' in script and 'to 4' in script
    assert len(calls)==1
    count=len(calls)
    with pytest.raises(ValueError):s.add_code_lines(['good',None])
    assert len(calls)==count

@pytest.mark.parametrize('method,args,kwargs',[
 ('set_columns',(0,),{}),('set_columns',(True,),{}),
 ('add_section',(),{'landscape':'bad'}),('add_section',(),{'continuous':'bad'}),
 ('finalize_fields',(),{'max_rounds':0}),('refresh_fields',(-1,),{}),
 ('ensure_heading_styles',({7:{'font_size':12}},),{}),
])
def test_invalid_input_before_transport(session,method,args,kwargs):
    s,calls=session
    with pytest.raises((ValueError,TypeError)):getattr(s,method)(*args,**kwargs)
    assert not calls

def test_sections_columns_and_footer_page_semantics(session):
    s,calls=session;s.set_columns(2);s.add_section(True,continuous=True);s.set_page_number_in_footer()
    assert 'number of columns 2' in '\n'.join(calls[0])
    assert 'section break continuous' in '\n'.join(calls[1])
    assert 'orient landscape' in '\n'.join(calls[1])
    page='\n'.join(calls[2]);assert 'section 1 of boundDoc' in page
    assert '"Page "' in page and 'field type field page' in page
    assert 'character 5 of text object of pagePart' in page

def test_compaction_guard_precedes_all_native_format_changes(session):
    s,calls=session;s.compact_terminal_paragraph()
    script='\n'.join(calls[0]);assert 'last paragraph of boundDoc' in script
    assert script.index('if compactable then')<script.index('set font size')
    assert 'line space exactly' in script and 'to 1' in script

def test_refresh_contract_separates_index_and_all_field_pass(session):
    s,calls=session;s.refresh_indexes();s.update_fields()
    assert 'tables of figures' in '\n'.join(calls[0]) and 'update field' not in '\n'.join(calls[0])
    assert 'update field' in '\n'.join(calls[1])
    snap=s.refresh_fields(2)
    assert len(snap)==1 and snap[0].stable_key==('doc:finalize','PAGE',0)
    assert snap[0].total_pages==3 and snap[0].toc_page_count==1 and snap[0].result_hash=='3-1'
    count=len(calls);assert s.finalize_fields(max_rounds=3) is None;assert len(calls)==count+1

def test_reset_preserves_owned_document_and_never_calls_native(session):
    s,calls=session;assert s.reset() is None;assert not calls

def test_readonly_rejects_followup_mutations(session):
    s,calls=session;s._read_only=True
    for method,args in [('set_columns',(2,)),('add_section',()),('ensure_styles',({'X':{}},)),('refresh_indexes',()),('update_fields',()),('compact_terminal_paragraph',())]:
        with pytest.raises(ValueError):getattr(s,method)(*args)
    assert not calls

@pytest.mark.skipif(sys.platform!='darwin' or not Path('/Applications/Microsoft Word.app').is_dir(),reason='Installed Word dictionary for compile-only checks')
def test_every_new_native_branch_compiles_without_execution(session,tmp_path):
    s,calls=session
    s.ensure_styles({'Custom':{'based_on':'Normal','font_size':12,'outline_level':10,'shading':'#EEEEEE','left_border':True,'border_color':'#999999'}})
    s.ensure_heading_styles({1:{'font_size':16,'keep_with_next':True,'based_on':'semanticStyleBase'}});s.apply_heading_text_color('#112233')
    s.add_code_lines(['code','']);s.set_columns(2);s.add_section(True);s.add_section(False,continuous=True)
    s.set_page_number_in_footer();s.compact_terminal_paragraph();s.refresh_indexes();s.update_fields();s.finalize_fields();s.refresh_fields(0)
    for index,lines in enumerate(calls):
        script=tmp_path/f'{index}.applescript';script.write_text('tell application "/Applications/Microsoft Word.app"\n'+'\n'.join(lines)+'\nend tell\n')
        result=subprocess.run(['/usr/bin/osacompile','-o',str(script.with_suffix('.scpt')),str(script)],capture_output=True,text=True,timeout=15)
        assert result.returncode==0,result.stderr


def test_footer_page_insertion_preserves_footer_story(session):
    s,calls=session;s.set_page_number_in_footer()
    script='\n'.join(calls[0])
    assert 'create range boundDoc' not in script
    assert 'collapse range (character 5 of text object of pagePart) direction collapse end' in script


def test_style_dependencies_preflight_and_creation_order(session):
    s,calls=session;s.ensure_styles({'Derived':{'based_on':'Base'},'Base':{'based_on':'Normal'}})
    script='\n'.join(calls[0])
    assert script.index('Word style (style normal)')<script.index('make new Word style')
    assert script.index('name local:"Base"')<script.index('name local:"Derived"')
    count=len(calls)
    with pytest.raises(ValueError,match='Cyclic'):s.ensure_styles({'A':{'based_on':'B'},'B':{'based_on':'A'}})
    assert len(calls)==count


def test_followup_fixture_import_cannot_start_office(monkeypatch):
    import importlib.util
    def forbidden(*args,**kwargs):raise AssertionError('Native launch during import')
    monkeypatch.setattr(MacWordSession,'_prepare',forbidden)
    path=Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_business_followup.py'
    spec=importlib.util.spec_from_file_location('followup_native_draft',path)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    assert callable(module.run) and callable(module.main)


def test_compaction_matches_unicode_whitespace_baseline(session):
    s,calls=session;s.compact_terminal_paragraph()
    script='\n'.join(calls[0])
    assert 'character id 160' in script and 'character id 12288' in script


def test_style_border_width_matches_frozen_two_and_quarter_points(session):
    s,calls=session;s.ensure_styles({'Border':{'left_border':True}})
    assert 'set line width of semanticBorder to line width225 point' in calls[0]


def test_heading_defaults_keep_regression_guardrail_sizes(session):
    from skills.WPSComposer.scripts.reference_styles import get_heading_style
    s,calls=session;s.ensure_heading_styles({i:get_heading_style(i) for i in range(1,7)})
    sizes=[line.rsplit(' ',1)[1] for line in calls[0] if line.startswith('set font size')]
    assert sizes==['16','15','15','14','14','12']
    assert sum('keep with next' in line for line in calls[0])==6


def test_heading_level_string_keys_remain_compatible_with_json_mappings(session):
    s,calls=session;s.ensure_heading_styles({'1':{'font_size':16},'6':{'font_size':12}})
    assert any('style heading1' in line for line in calls[0])
    assert any('style heading6' in line for line in calls[0])


def test_style_lookup_preserves_names_containing_script_variable(session):
    s,calls=session
    s.ensure_styles({'semanticStyleBody':{'based_on':'semanticStyleBase'}})
    assert 'set requestedStyle0 to Word style "semanticStyleBody" of boundDoc' in calls[0]
    assert 'set baseStyle0 to Word style "semanticStyleBase" of boundDoc' in calls[0]
    s.ensure_heading_styles({1:{'based_on':'semanticStyleBase'}})
    assert 'set headingBase1 to Word style "semanticStyleBase" of boundDoc' in calls[1]
    s.ensure_styles({'semanticStyleBody':{'based_on':'semanticStyleBase'},'semanticStyleBase':{}})
    assert 'set requestedStyle0 to Word style "semanticStyleBase" of boundDoc' in calls[2]
    assert 'set requestedStyle1 to Word style "semanticStyleBody" of boundDoc' in calls[2]


def test_heading_acyclic_dependencies_apply_base_before_derived(session):
    s,calls=session
    s.ensure_heading_styles({1:{'based_on':'Heading 2'},2:{'based_on':'Normal'}})
    lines=calls[0]
    assert lines.index('set semanticStyle to requestedHeading2')<lines.index('set semanticStyle to requestedHeading1')


@pytest.mark.parametrize('styles',[
    {1:{'font_size':16},2:{'based_on':'Heading 2'}},
    {1:{'based_on':'Heading 2'},2:{'based_on':'Heading 1'}},
    {1:{'based_on':'Heading 3'},3:{'based_on':'Heading 2'},2:{'based_on':'Heading 1'}},
    {1:{'font_size':16},'1':{'font_size':12}},
])
def test_heading_inheritance_rejected_as_whole_batch(session,styles):
    s,calls=session
    with pytest.raises(ValueError):s.ensure_heading_styles(styles)
    assert not calls


@pytest.mark.parametrize('method,styles,refs',[
    ('ensure_styles',{'Custom':{'based_on':'Base'}},['requestedStyle0','baseStyle0']),
    ('ensure_heading_styles',{1:{'based_on':'Base'}},['requestedHeading1','headingBase1']),
])
def test_paragraph_style_type_whitelist_covers_targets_and_bases_before_mutation(session,method,styles,refs):
    s,calls=session;getattr(s,method)(styles)
    lines=calls[0]
    first_mutation=next(i for i,line in enumerate(lines) if line.startswith('set semanticStyle to '))
    for ref in refs:
        guard=f'if (style type of {ref}) is not in {{style type paragraph, style type paragraph only, style type linked}} then error "WPSC_STYLE_TYPE_MISMATCH"'
        assert guard in lines[:first_mutation]


def test_native_fixture_preserves_two_columns_and_checks_native_xml():
    import ast
    source=(Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_business_followup.py').read_text()
    calls=[node for node in ast.walk(ast.parse(source)) if isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and node.func.attr=='set_columns']
    assert [node.args[0].value for node in calls]==[2]
    assert 'native_two_column_sections' in source
    assert 'w:sectPr/w:cols' in source


@pytest.mark.parametrize('method,styles',[
    ('ensure_styles',{'Foo':{'based_on':'foo'}}),
    ('ensure_styles',{'Foo':{},'foo':{}}),
    ('ensure_styles',{'Foo':{'based_on':'bar'},'Bar':{'based_on':'FOO'}}),
    ('ensure_heading_styles',{1:{'font_size':16},2:{'based_on':'heading 2'}}),
])
def test_style_graph_uses_case_insensitive_native_identity(session,method,styles):
    s,calls=session
    with pytest.raises(ValueError):getattr(s,method)(styles)
    assert not calls


def test_style_case_variant_dependency_uses_one_declared_base(session):
    s,calls=session;s.ensure_styles({'Derived':{'based_on':'base'},'Base':{}})
    lines=calls[0]
    assert 'set base style of semanticStyle to requestedStyle0' in lines
    assert not any('to Word style "base"' in line for line in lines)
    assert 'set requestedStyle0 to Word style "Base" of boundDoc' in lines
