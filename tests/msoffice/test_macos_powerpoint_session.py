from pathlib import Path
import pytest


def api():
    from skills.WPSComposer.scripts.msoffice import macos_powerpoint_session
    return macos_powerpoint_session


def test_session_factory_does_not_launch_or_write(tmp_path):
    s = api().MacPowerPointSession.open_document(tmp_path/'missing.pptx', read_only=True)
    assert s.kind == 'slide'
    assert s.engine == 'msoffice'
    assert not s.supports_attached_save_copy
    assert not (tmp_path/'missing.pptx').exists()


def test_unsafe_or_unsupported_patch_rejected_before_native_mutation():
    m=api()
    for patch in [{'font':{'size':float('nan')}}, {'text':'bad\0'}, {'geometry':{'width':-1}}, {'font':{'size':True}}]:
        with pytest.raises(ValueError):
            m.compile_patch('slide:1/shape:1', **patch)
    body, result=m.compile_patch('slide:1/shape:1',font={'unknown':'x'})
    assert not body
    assert result == {'accepted':[], 'rejected':['font.unknown']}


def test_nested_patch_handles_table_paragraph_and_native_text_flows():
    m=api()
    source,result=m.compile_patch('slide:2/shape:3/table/cell:1,2',text='中文 "text"',font={'bold':True})
    assert 'get cell from' in source and 'row 1 column 2' in source
    assert result['accepted']==['text','font.bold']
    source,_=m.compile_patch('slide:1/shape:1/paragraph:1/run:2',text='x')
    assert 'text flow 2 of targetText' in source


def test_invalid_structural_request_preflights_before_call(monkeypatch):
    s=api().MacPowerPointSession()
    s._entered=True
    s._read_only=False
    calls=[]
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    with pytest.raises(ValueError):
        s.apply_structural_op({'op':'insert','type':'textbox','parent':'slide:1','props':{'width':-3}})
    assert not calls


def test_readonly_patch_and_structural_operations_are_closed():
    s=api().MacPowerPointSession(read_only=True)
    with pytest.raises(PermissionError): s.apply_format_patch('presentation',page_setup={'slide_width':720})
    with pytest.raises(PermissionError): s.apply_structural_op({'op':'remove','target':'slide:1'})


def test_attached_close_never_closes_native_document(monkeypatch):
    s=api().MacPowerPointSession()
    s._attached=True;s._entered=True
    calls=[]
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    s.close()
    assert calls==[]
    with pytest.raises(api().PowerPointSessionCapabilityError):s.save_copy('/tmp/other.pptx')


def test_shape_names_escaped_and_numeric_ids_not_guessed():
    m=api()
    body,_=m.compile_patch('slide:1/shape:@name=x"y',text='safe')
    assert 'x\\"y' in body
    with pytest.raises(m.PowerPointSessionCapabilityError,match='ID'):
        m.compile_patch('slide:1/shape:@id=3',text='safe')


def test_snapshot_parser_does_not_invent_stable_numeric_ids():
    m=api()
    rows=[['doc','one','/tmp/one.pptx',True,1,720],['slide',1,256,'slide layout blank','note',True,[255,255,255]], ['shape',1,1,'shape-name','shape type text box',40,50,300,80,0,1,[255,0,0],False,0,'Text','Arial',18,True,False,False,[0,0,0]]]
    snap=m.parse_snapshot(rows)
    assert snap['kind']=='slide'
    assert snap['slides'][0]['native_slide_id']==256
    shape=snap['slides'][0]['shapes'][0]
    assert shape['shape_id'] is None
    assert shape['id']=='slide:1/shape:@name=shape-name'
    assert shape['fill']['color']=='#FF0000'


def test_max_shapes_is_global_and_include_text_false_suppresses_notes():
    m=api()
    rows=[['doc','one','/tmp/one.pptx',True,1,720],['slide',1,256,'blank','private',True,[255,255,255]]]
    snap=m.parse_snapshot(rows,include_text=False,max_shapes=0)
    assert snap['slides'][0]['notes'] is None


def test_session_total_deadline_is_not_refreshed(monkeypatch):
    s=api().MacPowerPointSession(timeout=1)
    s._deadline=10
    monkeypatch.setattr(api().time,'monotonic',lambda:11)
    with pytest.raises(api().NativeOfficeError):s._remaining()


def test_line_visible_is_explicitly_rejected_and_weight_uses_dictionary_name():
    source,result=api().compile_patch('slide:1/shape:1',line={'visible':False,'weight':2})
    assert result=={'accepted':[],'rejected':['line.visible']}
    assert not source
    source,_=api().compile_patch('slide:1/shape:1',line={'weight':2})
    assert 'line weight of line format' in source


def test_picture_geometry_binding_does_not_require_text_frame():
    source,_=api().compile_patch('slide:1/shape:3',geometry={'width':120})
    assert 'text range of text frame' not in source


def test_publication_uses_session_deadline(monkeypatch,tmp_path):
    s=api().MacPowerPointSession();s._deadline=123;s._path=str(tmp_path/'native.pptx')
    monkeypatch.setattr(s,'_run',lambda *a,**k:'SAVED')
    monkeypatch.setattr(api(),'validate_before_deadline',lambda *a,**k:None)
    calls=[]
    monkeypatch.setattr(api(),'publish_artifact',lambda a,b,**k:(calls.append(k) or b))
    s.save(tmp_path/'result.pptx')
    assert calls[0]['deadline']==123


def test_dependent_native_properties_emit_before_final_colors_and_geometry():
    source,_=api().compile_patch('slide:1',background={'color':'#112233'},follow_master_background=False)
    assert source.index('follow master background')<source.index('fore color')
    source,_=api().compile_patch('slide:1/shape:1',geometry={'height':70},line={'color':'#112233','weight':2},text_frame={'auto_size':0})
    assert source.index('line weight')<source.index('fore color')
    assert source.index('auto size')<source.index('set height')


def test_attached_preflight_prevents_mutation_before_unsaved_or_copy_failure():
    s=api().MacPowerPointSession();s._attached=True;s._path='Untitled';s._native_read_only=False
    with pytest.raises(api().PowerPointSessionCapabilityError):s.preflight_save()
    with pytest.raises(api().PowerPointSessionCapabilityError):s.preflight_save('/tmp/copy.pptx')
    s._path='/tmp/owned.pptx';s._native_read_only=True
    with pytest.raises(PermissionError):s.preflight_save()


def test_strikethrough_uses_native_strike_enum():
    source,result=api().compile_patch('slide:1/shape:1',font={'strikethrough':True})
    assert 'strike type of font of targetText to single strike' in source
    assert result['accepted']==['font.strikethrough']


def test_expired_owned_session_is_quarantined_before_cleanup(monkeypatch,tmp_path):
    s=api().MacPowerPointSession();s._deadline=1;s._entered=True;s._job=tmp_path;s._path=str(tmp_path/'bound.pptx')
    monkeypatch.setattr(api().time,'monotonic',lambda:2)
    with pytest.raises(api().NativeOfficeError):s._remaining()
    assert s._uncertain
    assert (tmp_path/'recovery.json').exists()


def test_duplicate_shape_names_update_all_descendant_addresses():
    rows=[['doc','one','/tmp/one',True,1,720],['slide',1,256,'blank','',True,[255,255,255]]]
    for i in (1,2):
        rows.append(['shape',1,i,'same','text box',0,0,100,50,0,i,[255,255,255],False,0,'x','Arial',12,False,False,False,[0,0,0]])
        rows.append(['paragraph',1,i,1,'x'])
    snapshot=api().parse_snapshot(rows)
    for shape in snapshot['slides'][0]['shapes']:
        assert shape['paragraphs'][0]['id'].startswith(shape['id']+'/')


def test_staged_copy_is_validated_before_native_open(monkeypatch,tmp_path):
    m=api();source=tmp_path/'source.pptx';source.write_bytes(b'fixture');job=tmp_path/'job';job.mkdir()
    s=m.MacPowerPointSession(source);events=[]
    monkeypatch.setattr(m,'validate_native_input',lambda p,*a,**k:events.append(('validated',str(p))))
    monkeypatch.setattr(s,'_prepare',lambda:setattr(s,'_job',job))
    monkeypatch.setattr(s,'_run',lambda *a,**k:(events.append(('native',a[0])) or 'OPEN'))
    with s:pass
    assert [event[0] for event in events[:3]]==['validated','validated','native']
    assert events[1][1]!=events[0][1]
    assert Path(events[1][1]).read_bytes()==b'fixture'


def test_readonly_close_save_rejection_keeps_session_releasable(monkeypatch):
    s=api().MacPowerPointSession(read_only=True);s._entered=True
    monkeypatch.setattr(s,'_run',lambda *a,**k:'CLOSED')
    with pytest.raises(PermissionError):s.close(save_changes=True)
    assert s._entered
    s.close()
    assert not s._entered


def test_snapshot_native_enums_match_public_numeric_schema():
    rows=[['doc','x','/tmp/x',True,1,720],['slide',1,256,'slide layout blank','',True,[255,255,255]],['shape',1,1,'title','shape type text box',0,0,100,50,0,1,[0,0,0],False,0,'x','Arial',12,False,False,False,[0,0,0]],['detail',1,1,[0,0,0],1,None,0,1,2,3,4,True,'paragraph align center',0,0,1]]
    result=api().parse_snapshot(rows)
    assert result['slides'][0]['layout']==12
    assert result['slides'][0]['shapes'][0]['type']==17
    assert result['slides'][0]['shapes'][0]['paragraph']['alignment']==2


def test_attached_preflight_validates_existing_native_input(monkeypatch,tmp_path):
    source=tmp_path/'bound.pptx';source.write_bytes(b'x')
    s=api().MacPowerPointSession();s._attached=True;s._path=str(source);s._deadline=123
    calls=[]
    monkeypatch.setattr(api(),'validate_native_input',lambda *a,**k:calls.append((a,k)))
    s.preflight_save(overwrite=False)
    assert calls[0][0]==(source,'presentation')
    assert calls[0][1]['deadline']==123


def test_missing_native_success_envelope_quarantines_mutation(monkeypatch,tmp_path):
    from types import SimpleNamespace
    s=api().MacPowerPointSession();s._entered=True;s._job=tmp_path;s._path=str(tmp_path/'bound.pptx');s._name='bound.pptx';s._deadline=9999999999
    monkeypatch.setattr(api().subprocess,'run',lambda *a,**k:SimpleNamespace(returncode=0,stdout='arbitrary output',stderr=''))
    with pytest.raises(api().NativeOfficeError):s._run('return "PATCHED"',mutation=True)
    assert s._uncertain
    assert (tmp_path/'001.stdout').read_text()=='arbitrary output'


@pytest.mark.parametrize('op',[
 {'op':'clone','target':'slide:1'},
 {'op':'clone','target':'slide:1/shape:1','to':{'slide':2}},
 {'op':'move','target':'slide:1/shape:1','to':{'slide':2}},
])
def test_clipboard_structural_ops_disclose_side_effect_and_bind_destination(monkeypatch,op):
    s=api().MacPowerPointSession();captured=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(captured.append(body) or '3'))
    result=s.apply_structural_op(op)
    assert result['clipboard_changed'] is True
    assert 'paste object' in captured[0]
    assert 'ownedDoc' in captured[0]
    assert 'the clipboard' not in captured[0]


def test_index_slide_move_uses_final_position_and_skips_noop(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or 'MOVED'))
    s.apply_structural_op({'op':'move','target':'slide:1','to':{'index':3}})
    assert 'after slide destinationIndex' in calls[0]
    assert 'before slide destinationIndex' in calls[0]


def test_move_preserves_original_until_destination_paste_is_confirmed(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or '2'))
    s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':2}})
    assert 'cut shape' not in calls[0]
    assert calls[0].index('set destinationSlide')<calls[0].index('copy shape')
    assert calls[0].index('paste count mismatch')<calls[0].index('delete targetShape')


def test_slide_insert_at_count_plus_one_has_explicit_append_branch(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or '3'))
    s.apply_structural_op({'op':'insert','type':'slide','position':{'index':3}})
    assert 'count of slides of ownedDoc' in calls[0]
    assert 'end of ownedDoc' in calls[0]


def test_mixed_unsupported_patch_is_rejected_before_any_mutation(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    result=s.apply_format_patch('slide:1/shape:1',text='Changed',font={'unknown':True})
    assert result=={'accepted':[],'rejected':['font.unknown']}
    assert not calls


def test_office_artifact_validation_uses_same_deadline(monkeypatch,tmp_path):
    s=api().MacPowerPointSession();s._deadline=123;s._path=str(tmp_path/'native.pptx')
    monkeypatch.setattr(s,'_run',lambda *a,**k:'SAVED')
    calls=[]
    monkeypatch.setattr(api(),'validate_before_deadline',lambda spec,path,deadline:calls.append(deadline),raising=False)
    monkeypatch.setattr(api(),'publish_artifact',lambda a,b,**k:b)
    s.save(tmp_path/'out.pptx')
    assert calls==[123]


def test_text_selection_snapshot_contains_font_and_paragraph(monkeypatch):
    import json
    s=api().MacPowerPointSession()
    monkeypatch.setattr(s,'_run',lambda *a,**k:json.dumps(['text','Selected','Arial',18,True,False,False,[18,52,86],'paragraph align center',0,6,1]))
    result=s.inspect_selection()
    assert result['font']['name']=='Arial'
    assert result['font']['color']=='#123456'
    assert result['paragraph']['alignment']==2


def test_shape_selection_returns_bound_native_shape_snapshot(monkeypatch):
    import json
    s=api().MacPowerPointSession();shape={'index':2,'name':'selected'}
    monkeypatch.setattr(s,'_run',lambda *a,**k:json.dumps(['shapes',1,[2]]))
    monkeypatch.setattr(s,'inspect_document',lambda:{'slides':[{'index':1,'shapes':[shape]}]})
    result=s.inspect_selection()
    assert result=={'kind':'slide_selection','count':1,'shapes':[shape]}


def test_extended_snapshot_keeps_native_cell_font_and_textframe_enums():
    rows=[['doc','x','/tmp/x',True,1,720],['slide',1,256,'slide layout blank','',True,[255,255,255]],['shape',1,1,'table','shape type table',0,0,100,50,0,1,[0,0,0],False,0,'','Arial',12,False,False,False,[0,0,0]],['cell',1,1,1,1,1,1,'Cell','Arial',16,True,False,False,[18,52,86],'paragraph align center',1,2,1,[255,255,255]]]
    cell=api().parse_snapshot(rows)['slides'][0]['shapes'][0]['table']['cells'][0]
    assert cell['font']['size']==16
    assert cell['paragraph']['alignment']==2
    assert cell['fill']['color']=='#FFFFFF'


def test_single_shape_selection_patch_binds_inspected_shape_and_checks_selection(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'inspect_selection',lambda:{'kind':'slide_selection','count':1,'shapes':[{'id':'slide:2/shape:3','index':3}]})
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or 'PATCHED'))
    result=s.apply_format_patch('selection',geometry={'left':60})
    assert result['accepted']==['geometry.left']
    assert 'Selection changed before mutation' in calls[0]
    assert 'shape 3 of targetSlide' in calls[0]


def test_multiple_shape_selection_patch_rejected_before_mutation(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'inspect_selection',lambda:{'kind':'slide_selection','count':2,'shapes':[{},{}]})
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    with pytest.raises(api().PowerPointSessionCapabilityError):s.apply_format_patch('selection',font={'bold':True})
    assert not calls


def test_attached_native_read_only_blocks_direct_mutation():
    s=api().MacPowerPointSession();s._native_read_only=True
    with pytest.raises(PermissionError):s.apply_format_patch('slide:1/shape:1',text='No')


@pytest.mark.parametrize('method,suffix',[('save','.pptx'),('export_pdf','.pdf')])
def test_session_outputs_preflight_existing_and_wrong_suffix_before_native(monkeypatch,tmp_path,method,suffix):
    s=api().MacPowerPointSession();s._job=tmp_path;s._path=str(tmp_path/'bound.pptx');calls=[]
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    output=tmp_path/('existing'+suffix);output.write_bytes(b'prior')
    with pytest.raises(FileExistsError):getattr(s,method)(output)
    with pytest.raises(ValueError):getattr(s,method)(tmp_path/'bad.txt')
    assert not calls and output.read_bytes()==b'prior'


def test_unknown_native_strike_does_not_become_true():
    rows=[['doc','x','/tmp/x',True,1,720],['slide',1,256,'slide layout blank','',True,[255,255,255]],['shape',1,1,'text','shape type text box',0,0,100,50,0,1,[0,0,0],False,0,'X','Arial',12,False,False,False,[0,0,0]],['paragraph',1,1,1,'X'],['run',1,1,1,1,'X','Arial',12,False,False,False,[0,0,0],'__NATIVE_UNAVAILABLE__']]
    run=api().parse_snapshot(rows)['slides'][0]['shapes'][0]['paragraphs'][0]['runs'][0]
    assert run['font']['strikethrough'] is None


def test_save_current_refuses_changed_source_before_native(monkeypatch,tmp_path):
    source=tmp_path/'source.pptx';source.write_bytes(b'changed')
    s=api().MacPowerPointSession(source);s._path=str(tmp_path/'bound.pptx');s._source_digest='old';s._deadline=9999999999
    calls=[];monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    with pytest.raises(ValueError,match='changed'):s.save_current()
    assert not calls


def test_new_source_session_creates_native_owned_document(monkeypatch,tmp_path):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_prepare',lambda:setattr(s,'_job',tmp_path))
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or 'CREATED'))
    with s:pass
    assert 'make new presentation' in calls[0]
    assert 'save as Open XML presentation' in calls[0]
    assert 'full name of ownedDoc' in calls[0]


def test_semantic_slide_methods_preserve_signatures_and_compile_native_text(monkeypatch):
    import inspect
    from skills.WPSComposer.scripts.slide import SlideComposer
    names=['set_slide_size','add_title_slide','add_section_slide','add_text_slide','add_bullets_slide','add_blank_slide','set_background_color','add_textbox','add_shape','add_image','add_table','set_notes','save_pptx','apply_design_preset','apply_layout_template']
    for name in names:assert inspect.signature(getattr(api().MacPowerPointSession,name))==inspect.signature(getattr(SlideComposer,name))
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or '1'))
    assert s.add_text_slide('Title',['A','B'],bullets=True)==1
    assert 'slide layout text slide' in calls[0]
    assert api().apple_string('A\rB') in calls[0]


def test_semantic_invalid_shape_and_size_rejected_before_mutation(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda *a,**k:calls.append(a))
    with pytest.raises(ValueError):s.add_shape(1,999999,0,0,100,100)
    with pytest.raises(ValueError):s.set_slide_size(960,-1)
    with pytest.raises(ValueError):s.add_table(1,2,2,0,0,100,100,[['x']],font_size=-1)
    assert not calls


def test_direct_arbitrary_size_guards_empty_presentation_before_any_setter(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**kwargs:(calls.append((body,kwargs)) or 'SIZED'))
    s.set_slide_size(720,405)
    source,kwargs=calls[0]
    guard='if (count of slides of ownedDoc) is not 0 then return "EXISTING_SLIDES"'
    assert source.index(guard)<source.index('set slide orientation')<source.index('set slide width')
    assert kwargs=={'mutation':True}


def test_direct_arbitrary_size_existing_slides_is_clean_capability_rejection(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**kwargs:(calls.append(body) or 'EXISTING_SLIDES'))
    with pytest.raises(api().PowerPointSessionCapabilityError,match='empty presentation'):
        s.set_slide_size(720,405)
    assert calls[0].index('return "EXISTING_SLIDES"')<calls[0].index('set slide orientation')
    assert not s._uncertain


def test_direct_540_size_keeps_existing_session_behavior(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**kwargs:(calls.append(body) or 'SIZED'))
    s.set_slide_size(960,540)
    assert calls == [
        'set slide size of page setup of ownedDoc to slide size on screen\n'
        'set slide width of page setup of ownedDoc to 960\nreturn "SIZED"'
    ]


def test_native_table_uses_editable_cells_and_layout_lines_use_native_line(monkeypatch):
    from skills.WPSComposer.scripts.layout_templates import LayoutTemplate
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or '1'))
    s.add_table(1,2,2,0,0,300,200,[['Name','Value'],['Item',42]])
    assert 'make new shape table' in calls[0] and 'get cell from table object' in calls[0]
    s.apply_layout_template(LayoutTemplate('x','x','x',[{'type':'line','x':0,'y':10,'w':100,'h':3,'color':'#123456'}]))
    assert any('make new line shape' in source for source in calls)


def test_new_document_factory_is_lazy_and_publicly_consistent():
    session=api().MacPowerPointSession.new_document(visible=False)
    assert session._source is None and not session._entered


def test_preflight_checks_source_conflict_before_public_mutation(monkeypatch,tmp_path):
    source=tmp_path/'source.pptx';source.write_bytes(b'changed')
    s=api().MacPowerPointSession(source);s._path=str(tmp_path/'bound.pptx');s._source_digest='old';s._deadline=9999999999
    with pytest.raises(ValueError,match='changed'):s.preflight_save()


def test_close_save_publishes_source_before_native_close(monkeypatch):
    s=api().MacPowerPointSession();s._entered=True;events=[]
    monkeypatch.setattr(s,'save_current',lambda:events.append('save_current'))
    monkeypatch.setattr(s,'_run',lambda body,**k:events.append(body))
    s.close(save_changes=True)
    assert events==['save_current','close ownedDoc saving no\nreturn "CLOSED"']


def test_preflight_save_destination_rejects_existing(monkeypatch,tmp_path):
    s=api().MacPowerPointSession();output=tmp_path/'existing.pptx';output.write_bytes(b'existing')
    with pytest.raises(FileExistsError):s.preflight_save(output)


def test_close_save_conflict_keeps_owned_session_available_for_discard(monkeypatch):
    s=api().MacPowerPointSession();s._entered=True
    def conflict():raise ValueError('Source changed')
    monkeypatch.setattr(s,'save_current',conflict)
    with pytest.raises(ValueError):s.close(save_changes=True)
    assert s._entered


def test_semantic_creation_does_not_require_changing_native_view(monkeypatch):
    s=api().MacPowerPointSession();calls=[]
    monkeypatch.setattr(s,'_run',lambda body,**k:(calls.append(body) or '1'))
    s.add_blank_slide()
    assert 'make new slide at end of ownedDoc' in calls[0]
    assert 'set slide of view' not in calls[0]


def test_truncated_snapshot_duplicate_name_uses_positional_target():
    rows=[['doc','x','/tmp/x',True,1,720],['slide',1,256,'slide layout blank','',True,[255,255,255]]]
    for i in (1,2):rows.append(['shape',1,i,'same','shape type text box',0,0,100,50,0,i,[0,0,0],False,0,'X','Arial',12,False,False,False,[0,0,0]])
    snapshot=api().parse_snapshot(rows,max_shapes=1)
    assert snapshot['slides'][0]['shapes'][0]['id']=='slide:1/shape:1'
