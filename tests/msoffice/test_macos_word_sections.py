"""Direct section/metadata contracts; transport capture is syntax-only evidence."""
import inspect
import subprocess
import sys
import pytest
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, NativeWordError
from skills.WPSComposer.scripts.writer import WriterComposer

METHODS=['configure_section','set_page_numbering','set_page_role','set_document_metadata','add_landscape_section_before_pending_heading']
@pytest.fixture
def session(monkeypatch):
    s=MacWordSession();calls=[]
    def execute(lines,**kwargs):
        calls.append(lines)
        return [['heading',17,23,2]] if any('"heading",' in line for line in lines) else [['ok']]
    monkeypatch.setattr(s,'_execute',execute)
    return s,calls

@pytest.mark.parametrize('name',METHODS)
def test_exact_signatures(name):
    assert inspect.signature(getattr(MacWordSession,name))==inspect.signature(getattr(WriterComposer,name))

@pytest.mark.parametrize('kwargs',[
 {'page_size':'bad'}, {'page_number_format':'bad'}, {'start_page_number':True},
 {'margins':{'top':20,'left':-1}}, {'margins':{'unknown':1}}, {'landscape':'yes'},
 {'role':object()}, {'header_text':object()}, {'restart_page_numbering':'yes'},
 {'margins':{'top':float('inf')}}, {'margins':[]},
])
def test_full_config_preflight(session,kwargs):
    s,calls=session
    with pytest.raises((ValueError,TypeError)):s.configure_section(**kwargs)
    assert not calls and not getattr(s,'_first_section_configured',False)

@pytest.mark.parametrize('method,args,kwargs',[
 ('configure_section',(),{}), ('set_page_numbering',('arabic',),{}),
 ('set_page_role',('body',),{}), ('set_document_metadata',(),{'title':'a','author':'b'}),
 ('add_landscape_section_before_pending_heading',(),{}),
])
def test_read_only_before_transport(session,method,args,kwargs):
    s,calls=session;s._read_only=True
    with pytest.raises(ValueError,match='read.only'):getattr(s,method)(*args,**kwargs)
    assert not calls

def test_config_first_later_and_acknowledged_state(session,monkeypatch):
    s,calls=session;s.configure_section(page_size='a4',margins={'top':60},footer_text='Keep',page_number_format='roman',start_page_number=3,restart_page_numbering=True)
    assert s._first_section_configured
    first='\n'.join(calls[0]);assert 'insert break' not in first and '595.28' in first
    assert 'top margin' in first and 'to 60' in first
    s.configure_section(landscape=True,link_to_previous_footer=False,footer_text='Next')
    later='\n'.join(calls[1]);assert later.count('insert break')==1
    assert later.index('set link to previous')<later.index('set content of text object of ownFooter')
    s2=MacWordSession();monkeypatch.setattr(s2,'_execute',lambda *_a,**_k:(_ for _ in ()).throw(RuntimeError('setter rejected')))
    with pytest.raises(RuntimeError):s2.configure_section()
    assert not s2._first_section_configured

@pytest.mark.parametrize('fmt,enum',[('roman','lowercase roman'),('roman-upper','uppercase roman'),('roman-lower','lowercase roman'),('arabic','arabic'),('continue','arabic')])
def test_numbering_exact_styles_and_footer_bound_field(session,fmt,enum):
    s,calls=session;s.set_page_numbering(fmt,3,True)
    code='\n'.join(calls[0])
    assert 'page number style '+enum in code
    assert 'field type of ownField is field page' in code
    assert 'character lastCharacter of text object of ownFooter' in code
    assert 'starting number' in code and 'WPSC_NUMBERING_READBACK_FAILED' in code
    assert s._structural_changed

def test_none_deletes_only_true_page_fields(session):
    s,calls=session;s.set_page_numbering('none')
    code='\n'.join(calls[0]);assert 'delete ownField' in code
    assert 'field type of ownField is field page' in code
    assert 'set content of text object of ownFooter' not in code
    assert 'field num pages' not in code and 'field page ref' not in code

def test_role_update_is_document_local_idempotent(session):
    s,calls=session;s.set_page_role('正文😀')
    code='\n'.join(calls[0]);assert 'WpsComposerSectionRole_' in code
    assert code.index('count variables of boundDoc')<code.index('make new variable at boundDoc')
    assert 'variable value of ownRole' in code and 'WPSC_ROLE_READBACK_FAILED' in code
    assert not s._structural_changed

def test_metadata_preflights_both_and_reads_document_properties(session):
    s,calls=session
    s.set_document_metadata(title='标题😀',author=None)
    code='\n'.join(calls[0]);assert 'document property "Title" of boundDoc' in code
    assert 'document property "Author" of boundDoc' in code and 'WPSC_METADATA_READBACK_FAILED' in code
    assert 'user name' not in code and not s._structural_changed

def test_pending_heading_native_position_and_intervening_content(session):
    s,calls=session;s.add_heading_level('中文😀',2)
    assert s._pending_heading== (17,23,2)
    s.add_landscape_section_before_pending_heading()
    code='\n'.join(calls[-1]);assert 'start 17 end 17' in code
    assert code.index('WPSC_PENDING_HEADING_STALE')<code.index('insert break')
    assert s._pending_heading is None
    s.add_heading_level('中文😀',2);s.add_paragraph('later')
    before=len(calls)
    with pytest.raises(ValueError,match='pending heading'):s.add_landscape_section_before_pending_heading()
    assert len(calls)==before

def test_malformed_heading_ack_is_rejected(session,monkeypatch):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda *_a,**_kw:[['heading',True,23,2]])
    with pytest.raises(NativeWordError):s.add_heading_level('中文😀',2)
    assert s._pending_heading is None

@pytest.mark.skipif(sys.platform!='darwin',reason='AppleScript compiler required')
def test_compile_all_section_branches(session,tmp_path):
    s,calls=session
    for fmt in ('none','arabic','roman-upper','roman-lower','continue'):
        s.set_page_numbering(fmt,3,True)
    s.configure_section(page_size='letter',margins={'top':60},landscape=True,header_text='中文😀',footer_text='Footer',link_to_previous_header=False,link_to_previous_footer=False)
    s.configure_section(role='body',page_size='a3');s.set_document_metadata(title='标题😀',author='作者😀')
    s.add_heading_level('中文😀',2);s.add_landscape_section_before_pending_heading()
    for i,lines in enumerate(calls):
        source=tmp_path/f'{i}.applescript';source.write_text('tell application "Microsoft Word"\n'+'\n'.join(lines)+'\nend tell')
        result=subprocess.run(['osacompile','-o',str(source.with_suffix('.scpt')),str(source)],capture_output=True,text=True)
        assert result.returncode==0,result.stderr

def test_page_size_normalizes_inherited_landscape_before_dimensions(session):
    s,calls=session;s.configure_section(page_size='a4',landscape=False)
    code='\n'.join(calls[0])
    assert code.index('set orientation of page setup of ownSection to orient portrait') < code.index('set page width of page setup of ownSection')

@pytest.mark.parametrize('operation', ['structural', 'format'])
def test_other_content_mutations_clear_pending_heading(session,operation):
    s,calls=session;s.add_heading_level('中文😀',2)
    if operation=='structural':s.apply_structural_op({'op':'insert','type':'paragraph','text':'later','at':'end'})
    else:s.apply_format_patch('paragraph:1',text='replacement')
    assert s._pending_heading is None

def test_configure_section_does_not_advance_after_malformed_ack(session,monkeypatch):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda *_a,**_kw:[['wrong']])
    with pytest.raises(NativeWordError):s.configure_section()
    assert not s._first_section_configured and s._quarantined

@pytest.mark.parametrize('method,kwargs',[
 ('set_document_metadata',{'title':'a','author':'b'}),('set_page_role',{'role':'body'}),
 ('set_page_numbering',{'format':'none'}),
])
def test_new_methods_require_exact_native_ack(session,monkeypatch,method,kwargs):
    s,calls=session;monkeypatch.setattr(s,'_execute',lambda *_a,**_kw:[])
    with pytest.raises(NativeWordError):getattr(s,method)(**kwargs)
    assert s._quarantined

def test_implicit_landscape_section_continues_page_numbers(session):
    s,calls=session;s.add_heading_level('中文😀',2);s.add_landscape_section_before_pending_heading()
    code='\n'.join(calls[-1])
    assert 'set restart numbering at section of page number options of ownFooter to false' in code
    assert 'if restart numbering at section of page number options of ownFooter is not false' in code

def test_inherited_orientation_restored_before_explicit_margins(session):
    s,calls=session;s.configure_section(page_size='legal',margins={'top':50,'left':70})
    code='\n'.join(calls[0])
    assert code.index('set orientation of page setup of ownSection to inheritedOrientation') < code.index('set top margin of page setup of ownSection')

@pytest.mark.parametrize('format',[False,0,'',[],{}])
def test_supplied_invalid_format_cannot_hide_behind_default(session,format):
    s,calls=session
    with pytest.raises(ValueError):s.configure_section(page_number_format=format)
    assert not calls


class MetadataText:
    def __str__(self):
        return 'object "中文😀"'


@pytest.mark.parametrize('value', [7, MetadataText(), [1, '中文😀'], 0, False, None, '', [], {}])
def test_metadata_matches_frozen_value_coercion(session,value):
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    s,calls=session
    assert s.set_document_metadata(title=value,author=value) is None
    code='\n'.join(calls[0]);expected=apple_string(str(value or ''))
    assert f'set value of titleProperty to {expected}' in code
    assert f'set value of authorProperty to {expected}' in code
    assert code.index('get value of authorProperty')<code.index('set value of titleProperty')
    assert len(calls)==1 and not s._structural_changed


class MetadataConversionFailure:
    def __str__(self):
        raise RuntimeError('metadata conversion failed')


class MetadataTruthFailure:
    def __bool__(self):
        raise RuntimeError('metadata truth evaluation failed')


class MetadataControlCharacter:
    def __str__(self):
        return 'invalid\x00text'


@pytest.mark.parametrize('value,error', [
    (MetadataConversionFailure(),RuntimeError),
    (MetadataTruthFailure(),RuntimeError),
    (MetadataControlCharacter(),ValueError),
])
def test_second_metadata_conversion_or_quoting_failure_prevents_all_native_calls(session,value,error):
    s,calls=session
    with pytest.raises(error):s.set_document_metadata(title='valid title',author=value)
    assert not calls and not s._structural_changed
