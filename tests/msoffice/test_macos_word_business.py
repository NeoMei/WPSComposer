"""Direct Writer API parity; native acceptance is a separate, explicit gate."""
import inspect
import pytest
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, NativeWordCapabilityError
from skills.WPSComposer.scripts.writer import WriterComposer
from skills.WPSComposer.scripts.document_model import Span

METHODS=['add_paragraph','add_heading','add_heading2','add_heading_level','add_centered','add_rich_paragraph','add_styled_paragraph','add_bullet_list','add_numbered_list','add_page_break','set_margins','set_page_size','set_orientation','set_header','set_footer','set_header_footer','save_docx']
@pytest.fixture
def session(monkeypatch):
    s=MacWordSession();calls=[]
    monkeypatch.setattr(s,'_execute',lambda lines,**kwargs:calls.append(lines) or [['ok']])
    return s,calls

@pytest.mark.parametrize('name',METHODS)
def test_exact_direct_signatures(name):
    assert inspect.signature(getattr(MacWordSession,name))==inspect.signature(getattr(WriterComposer,name))

def test_paragraph_boundary_reset_and_direct_points(session):
    s,calls=session
    assert s.add_paragraph('A😀',font_name='仿宋',font_name_ascii='Times New Roman',line_spacing=18) is None
    script='\n'.join(calls[0])
    assert 'set precedingRange to create range boundDoc' in script
    assert 'insertionPoint + 4' in script
    assert 'reset font object of semanticRange' in script
    assert 'set line spacing of paragraph format of semanticRange to 18' in script
    assert 'line space multiple' not in script
    assert 'ascii name' in script and 'east asian name' in script
    assert s._structural_changed

def test_multiple_spacing_converts_lines_but_exact_keeps_points(session):
    s,calls=session
    s.add_paragraph('Multiple',line_spacing=1.25,line_spacing_rule='multiple')
    s.add_paragraph('Exact',line_spacing=17,line_spacing_rule='exact')
    assert 'to 15.0' in '\n'.join(calls[0]) and 'line space multiple' in '\n'.join(calls[0])
    assert 'to 17' in '\n'.join(calls[1]) and 'line space exactly' in '\n'.join(calls[1])

def test_heading_defaults_clamp_and_overrides(session):
    s,calls=session
    assert s.add_heading_level('Heading',99,bold=False,size=22) is None
    script='\n'.join(calls[0]);assert 'style heading6' in script
    assert 'outline level6' in script and 'to 22' in script and 'to false' in script
    s.add_heading('H1');s.add_heading2('H2')
    assert 'style heading1' in '\n'.join(calls[1]);assert 'style heading2' in '\n'.join(calls[2])

@pytest.mark.parametrize('method,args,kwargs',[
 ('add_paragraph',('x',),{'size':-1}),('add_paragraph',('x',),{'line_spacing_rule':'bad'}),
 ('add_heading_level',('x',),{'color':'bad'}),('add_rich_paragraph',([Span('ok'),Span('bad',link='https://example.com')],'Body Text'),{}),
 ('add_bullet_list',(['ok',object()],),{}),('add_bullet_list',(['ok'],),{'indent':-1}),
 ('set_margins',(72,72,-1,72),{}),('set_page_size',(0,800),{}),('set_header_footer',(),{'header':'ok','link_to_previous_footer':'bad'}),
 ('add_styled_paragraph',('x',None),{})])
def test_invalid_input_rejected_before_any_native_call(session,method,args,kwargs):
    s,calls=session
    with pytest.raises((ValueError,TypeError)):getattr(s,method)(*args,**kwargs)
    assert not calls

def test_rich_ranges_use_utf16_and_preserve_per_span_font(session):
    s,calls=session
    assert s.add_rich_paragraph([Span('😀',bold=True),Span('代码',code=True,strikethrough=True)],'Body Text') is None
    script='\n'.join(calls[0])
    assert 'start (insertionPoint + 2) end (insertionPoint + 4)' in script
    assert 'Consolas' in script and 'strike through' in script

def test_lists_preserve_literal_prefix_tab_and_hanging_indent(session):
    s,calls=session
    s.add_bullet_list(['one','two'],glyph='→',indent=28)
    s.add_numbered_list(['first','second'])
    first='\n'.join(calls[0]);second='\n'.join(calls[1])
    assert '→' in first and 'to -28' in first and 'tab stop position:28' in first
    assert '1.' in second and '2.' in second
    assert 'apply number default' not in second and 'apply bullet default' not in first
    assert len(calls)==2

def test_named_style_resolves_before_appending(session):
    s,calls=session;s.add_styled_paragraph('x','Custom')
    script='\n'.join(calls[0]);assert script.index('Word style "Custom"')<script.index('set content of semanticRange')

def test_header_section_semantics_and_link_before_text(session):
    s,calls=session
    s.set_header('First');s.set_footer('First footer');s.set_header_footer(header='Last',footer='Footer',link_to_previous_header=False)
    assert 'section 1 of boundDoc' in '\n'.join(calls[0])
    assert 'section 1 of boundDoc' in '\n'.join(calls[1])
    last='\n'.join(calls[2]);assert 'section (count sections of boundDoc)' in last
    assert last.index('set link to previous')<last.index('set content')
    assert 'border bottom' in last and 'line width75 point' in last

def test_page_setup_and_save_alias(session,monkeypatch):
    s,calls=session;s.set_margins(1,2,3,4);s.set_page_size(600,800);s.set_orientation(True);s.add_page_break()
    assert all('boundDoc' in '\n'.join(lines) for lines in calls)
    assert 'orient landscape' in '\n'.join(calls[2])
    monkeypatch.setattr(s,'save',lambda path,fmt=None:(path,fmt))
    assert s.save_docx('/tmp/out.docx')==('/tmp/out.docx',12)

@pytest.mark.parametrize('method,args',[('add_paragraph',('x',)),('add_bullet_list',(['x'],)),('set_header',('x',)),('set_page_size',(600,800))])
def test_readonly_rejects_business_mutations(session,method,args):
    s,calls=session;s._read_only=True
    with pytest.raises(ValueError):getattr(s,method)(*args)
    assert not calls


def test_known_first_paragraph_style_is_created_locally_before_content(session):
    s,calls=session;s.add_rich_paragraph([Span('First')],'First Paragraph')
    script='\n'.join(calls[0])
    assert 'make new Word style at boundDoc' in script
    assert script.index('make new Word style')<script.index('set content of semanticRange')
    assert 'base style of semanticStyle to style body text' in script


def test_list_later_invalid_item_prevents_earlier_insert(session):
    s,calls=session
    with pytest.raises(ValueError):s.add_numbered_list(['Valid',3])
    assert calls==[]


def test_append_suffix_resets_only_empty_terminal_paragraph(session):
    s,calls=session;s.add_heading('H')
    script='\n'.join(calls[0])
    assert 'trailingPoint end trailingPoint' in script
    assert script.index('outline level1')<script.index('set style of trailingRange to style normal')


def test_native_fixture_import_has_no_office_side_effects(monkeypatch):
    import importlib.util
    from pathlib import Path
    def forbidden(*args,**kwargs):raise AssertionError('Native call during fixture import')
    monkeypatch.setattr(MacWordSession,'_prepare',forbidden)
    path=Path(__file__).resolve().parents[2]/'fixtures/microsoft_parity/macos_word_business.py'
    spec=importlib.util.spec_from_file_location('native_business_draft',path)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    assert callable(module.run) and callable(module.main)


def test_rich_code_uses_existing_character_style_before_font_fallback(session):
    s,calls=session;s.add_rich_paragraph([Span('code',code=True)],'Body Text')
    script='\n'.join(calls[0])
    assert 'Word style "Verbatim Char" of boundDoc' in script
    assert script.index('Word style "Verbatim Char"')<script.index('set content of semanticRange')
    assert 'set style of spanRange to semanticCodeStyle' in script


def test_header_border_matches_frozen_word_enum_value(session):
    # WriterComposer uses wdLineWidth075pt (6), not 0.5pt (4).
    s,calls=session;s.set_header_footer(header='Header')
    assert 'set line width of pageBorder to line width75 point' in calls[0]


def test_header_footer_applescript_compiles_without_execution(session,tmp_path):
    import subprocess
    import sys
    from pathlib import Path
    if sys.platform!='darwin' or not Path('/Applications/Microsoft Word.app').is_dir():
        pytest.skip('Installed macOS Word dictionary required for syntax-only compilation')
    s,calls=session;s.set_header_footer(header='Header',footer='Footer',link_to_previous_header=False,link_to_previous_footer=True)
    source='tell application "/Applications/Microsoft Word.app"\n'+'\n'.join(calls[0])+'\nend tell\n'
    script=tmp_path/'header.applescript';script.write_text(source)
    result=subprocess.run(['/usr/bin/osacompile','-o',str(tmp_path/'header.scpt'),str(script)],capture_output=True,text=True,timeout=15)
    assert result.returncode==0,result.stderr


@pytest.mark.parametrize('style',['Body Text','First Paragraph'])
def test_rich_body_clears_inherited_left_and_right_indent(session,style):
    s,calls=session;s.add_rich_paragraph([Span('Body')],style)
    script='\n'.join(calls[0])
    assert 'set paragraph format left indent of paragraph format of semanticRange to 0' in script
    assert 'set paragraph format right indent of paragraph format of semanticRange to 0' in script


def test_consolas_preserves_inherited_east_asian_font_and_sets_latin_slots():
    lines=MacWordSession._business_format('r',{'font_name':'Consolas','font_name_ascii':'Consolas'})
    assert not any('east asian name' in line for line in lines)
    for slot in ['ascii name','other name','complex script name']:
        assert f'set {slot} of font object of r to "Consolas"' in lines
