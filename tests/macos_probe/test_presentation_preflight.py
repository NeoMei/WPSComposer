"""Pure contract checks for automatic WPS PPT edit routing; no native host."""
from __future__ import annotations

import importlib
import math

import pytest


def supports(operations):
    module=importlib.import_module('skills.WPSComposer.scripts.macos_probe.presentation_preflight')
    return module.supports_presentation_set_ops(operations)


def op(target='slide:1/shape:2',**patch):
    return {'op':'set','target':target,**patch}


@pytest.mark.parametrize('target', ['slide:1/shape:2','slide:1/shape:@id=12','slide:1/shape:@name=标题',
    'slide:1/shape:2/table/cell:3,4','slide:1/shape:@id=12/table/cell:3,4'])
def test_shape_contract_accepts_all_implemented_groups(target):
    assert supports([op(target,text='Native text',font={'name':'Arial','size':14,'bold':True,'italic':False,'underline':-1,'strikethrough':0,'color':0x112233},
        paragraph={'alignment':2,'left_indent':0,'first_line_indent':-10,'line_spacing':1.2,'line_rule_within':-1,'space_before':0,'space_after':6},
        geometry={'left':-5,'top':0,'width':100,'height':20,'rotation':360},fill={'color':'112233','visible':True,'transparency':0.2},line={'color':'#445566','visible':-1,'weight':2},
        text_frame={'margin_left':1,'margin_right':2,'margin_top':3,'margin_bottom':4,'word_wrap':-1,'auto_size':1,'vertical_anchor':3},vertical_alignment=3)])


@pytest.mark.parametrize('target',['slide:1/shape:2/paragraph:1','slide:1/shape:2/paragraph:1/run:2','slide:1/shape:@id=12/paragraph:1','slide:1/shape:@id=12/paragraph:1/run:2'])
def test_text_subtargets_only_support_text_font_paragraph(target):
    assert supports([op(target,text='Text',font={'bold':True},paragraph={'alignment':1})])
    assert not supports([op(target,fill={'color':'#FFFFFF'})])


def test_slide_contract_is_distinct_from_shape_contract():
    assert supports([op('slide:1',name='Title',follow_master_background=False,background={'color':'#FFFFFF','transparency':0,'visible':True})])
    assert not supports([op('slide:1',text='Ignored text')])
    assert not supports([op(name='Ignored shape rename')])


@pytest.mark.parametrize('patch',[{'font':{'character_spacing':2}}, {'paragraph':{'bullet':True}},
 {'geometry':{'z_order':2}}, {'fill':{'back_color':'#FFFFFF'}}, {'line':{'dash_style':1}},
 {'text_frame':{'orientation':1}}, {'notes':'ignored'}, {'chart_type':4}, {'unknown':None}])
def test_unknown_top_level_and_nested_fields_reject_entire_request(patch):
    assert not supports([op(text='valid'),op(**patch)])


@pytest.mark.parametrize('target',['selection','presentation','slide:0','slide:-1','slide:1/shape:0',
 'slide:1/shape:@id=0','slide:1/shape:1/table/cell:0,2','slide:1/shape:@name=Title/paragraph:1',
 'slide:1/shape:1/paragraph:0','slide:1/shape:1/paragraph:1/run:0','slide:1/chart:1'])
def test_unknown_or_invalid_targets_do_not_select_wps(target):
    assert not supports([op(target,text='x')])


@pytest.mark.parametrize('patch',[{'font':'bad'}, {'font':{'size':True}}, {'font':{'size':-1}},
 {'font':{'color':'red'}}, {'fill':{'transparency':1.1}}, {'geometry':{'width':-1}},
 {'geometry':{'left':math.inf}}, {'paragraph':{'alignment':99}}, {'line':{'weight':-1}},
 {'text_frame':{'auto_size':99}}, {'vertical_alignment':99}, {'text':{'nested':'object'}},
 {'vertical_alignment':2,'text_frame':{'vertical_anchor':3}}])
def test_value_validation_prevents_known_native_rejections_and_coercion(patch):
    assert not supports([op(**patch)])


def test_none_and_empty_groups_are_known_noops_and_input_not_mutated():
    patch=op(text=None,font=None,fill={},text_frame={'vertical_anchor':3},vertical_alignment=3)
    before=repr(patch)
    assert supports([patch]) and repr(patch)==before
    assert supports([])
    assert not supports([{'op':'clone','target':'slide:1'}])
    assert not supports([None])
    assert not supports([{'target':'slide:1','name':'missing verb'}])


def test_auto_skips_incomplete_wps_request_but_explicit_wps_is_unchanged(monkeypatch):
    from skills.WPSComposer.scripts import document_api,office_engines
    monkeypatch.setattr(office_engines.sys,'platform','darwin')
    monkeypatch.setattr(office_engines,'engine_executable',lambda *a:'/installed')
    assert document_api._document_engine('input.pptx',None,'auto',action='edit',operations=[op(name='New shape name')])=='msoffice'
    assert document_api._document_engine('input.pptx',None,'auto',action='edit',operations=[op(text='Allowed')])=='wps'
    assert document_api._document_engine('input.pptx',None,'wps',action='edit',operations=[op(name='New shape name')])=='wps'


def test_auto_no_capable_installed_engine_rejects_before_native(monkeypatch):
    from skills.WPSComposer.scripts import document_api,office_engines
    monkeypatch.setattr(office_engines.sys,'platform','darwin')
    monkeypatch.setattr(office_engines,'engine_executable',lambda engine,*a:'/WPS' if engine=='wps' else None)
    with pytest.raises(office_engines.EngineUnavailableError):
        document_api._document_engine('input.pptx',None,'auto',action='edit',operations=[op(name='Ignored')])


def test_real_addin_silently_ignores_shape_name_but_helper_does_not_claim_it():
    """Execute the shipped JS on plain objects; no Office/bridge/UI is involved."""
    import json
    from pathlib import Path
    import shutil
    import subprocess
    node=shutil.which('node')
    if node is None:
        pytest.skip('Node is required for the shipped-handler contract check')
    script=Path(__file__).resolve().parents[2]/'macos/wps-jsapi-probe/addin/presentation.js'
    harness=r'''
const fs=require('fs');global.window={};
const shape={Name:'Original',Left:0,Fill:{ForeColor:{}},Line:{ForeColor:{}},TextFrame:{TextRange:{Text:'old',Font:{Color:{}},ParagraphFormat:{}}}};
const slide={Name:'Slide',Shapes:{Count:1,Item(){return shape;}}};
const presentation={Slides:{Item(){return slide;}},SaveAs(){},Close(){}};
global.Application={DisplayAlerts:7,Presentations:{Open(){return presentation;}}};
eval(fs.readFileSync(process.argv[1],'utf8'));
const result=window.WPSComposerProbe.handleCommand({method:'edit_presentation',params:{sourcePath:'/synthetic/input.pptx',outputPath:'/synthetic/output.pptx',patches:[{target:'slide:1/shape:1',name:'Ignored',text:'Applied',geometry:{left:10}}]}});
process.stdout.write(JSON.stringify({result,name:shape.Name,text:shape.TextFrame.TextRange.Text,left:shape.Left}));
'''
    result=subprocess.run([node,'-e',harness,str(script)],capture_output=True,text=True,timeout=10,check=True)
    state=json.loads(result.stdout)
    assert state['result']['patches'][0]['ok'] is True
    assert state['name']=='Original' and state['text']=='Applied' and state['left']==10
    assert not supports([op('slide:1/shape:1',name='Ignored',text='Applied',geometry={'left':10})])
    assert supports([op('slide:1/shape:1',text='Applied',geometry={'left':10})])


def test_geometry_accepts_zero_size_lines_and_textframe_uses_ppautosize_enum():
    assert supports([op(geometry={'width':0,'height':0})])
    assert not supports([op(text_frame={'auto_size':2})])


def test_unrepresentable_json_integer_is_unsupported_without_overflow():
    assert not supports([op(geometry={'left':10**400})])


def test_integral_json_numeric_enum_values_preserve_native_number_semantics():
    assert supports([op(paragraph={'alignment':2.0},font={'bold':-1.0},text_frame={'auto_size':1.0})])


def test_integral_json_color_preserves_native_number_semantics():
    assert supports([op(fill={'color':1122867.0})])
    assert not supports([op(fill={'color':1122867.5})])


def _run_real_addin_patch(patch):
    """Run actual handlers against plain objects, including JSON wire coercion."""
    import json
    from pathlib import Path
    import shutil
    import subprocess
    node = shutil.which('node')
    if node is None:
        pytest.skip('Node is required for the shipped-handler contract check')
    script = Path(__file__).resolve().parents[2] / 'macos/wps-jsapi-probe/addin/presentation.js'
    harness = r'''
const fs=require('fs');global.window={};
const shape={Fill:{ForeColor:{}},Line:{ForeColor:{}},TextFrame:{TextRange:{Text:'old',Font:{Color:{}},ParagraphFormat:{}}}};
const slide={Background:{Fill:{ForeColor:{}}},Shapes:{Count:1,Item(){return shape;}}};
const presentation={Slides:{Item(){return slide;}},SaveAs(){},Close(){}};
global.Application={DisplayAlerts:7,Presentations:{Open(){return presentation;}}};
eval(fs.readFileSync(process.argv[1],'utf8'));
const result=window.WPSComposerProbe.handleCommand({method:'edit_presentation',params:{sourcePath:'/synthetic/input.pptx',outputPath:'/synthetic/output.pptx',patches:[JSON.parse(process.argv[2])]}});
process.stdout.write(JSON.stringify({result,text:shape.TextFrame.TextRange.Text,fill:shape.Fill,line:shape.Line,background:slide.Background.Fill}));
'''
    result = subprocess.run([node, '-e', harness, str(script), json.dumps(patch)],
                            capture_output=True, text=True, timeout=10, check=True)
    return json.loads(result.stdout)


@pytest.mark.parametrize('group', ['fill', 'line', 'background'])
@pytest.mark.parametrize('visible', [False, 0])
def test_real_addin_color_overrides_earlier_hidden_visibility(group, visible):
    target = 'slide:1' if group == 'background' else 'slide:1/shape:1'
    patch = op(target, **{group: {'visible': visible, 'color': '#112233'}})
    state = _run_real_addin_patch(patch)
    assert state['result']['patches'][0]['ok'] is True
    assert state[group]['Visible'] == -1
    assert not supports([patch])


@pytest.mark.parametrize('group', ['fill', 'line', 'background'])
def test_color_then_hidden_visibility_and_visible_colors_remain_supported(group):
    target = 'slide:1' if group == 'background' else 'slide:1/shape:1'
    patch = op(target, **{group: {'color': '#112233', 'visible': False}})
    assert _run_real_addin_patch(patch)[group]['Visible'] is False
    assert supports([patch])
    assert supports([op(target, **{group: {'visible': True, 'color': 0x112233}})])


@pytest.mark.parametrize('value', [9007199254740993, -9007199254740993])
def test_real_addin_lossy_numeric_text_is_not_auto_supported(value):
    patch = op('slide:1/shape:1', text=value)
    state = _run_real_addin_patch(patch)
    assert state['result']['patches'][0]['ok'] is True
    assert state['text'] != str(value)
    assert not supports([patch])
    assert supports([op(text=str(value))])
    assert supports([op(text=9007199254740991)])
