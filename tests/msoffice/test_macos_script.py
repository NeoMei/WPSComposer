from pathlib import Path
import pytest
from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation


def test_text_serializer_escapes_script_injection():
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert apple_string('a"\\\n\r\t中') == '"a\\"\\\\" & linefeed & "" & return & "" & tab & "中"'
    with pytest.raises(ValueError):
        apple_string('bad\x00')


def test_compiler_preserves_body_styles_and_uses_real_pagination(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    build = build_longform_generation('# 标题\n\n## 第一节\n\n中文正文。')
    compiled = compile_plan(build.plan, {}, tmp_path / 'owned.docx', timeout=20)
    assert 'east asian name' in compiled.source
    assert 'first line indent' in compiled.source
    assert 'outline level body text' in compiled.source
    assert 'outline level10' not in compiled.source
    assert 'get range information' in compiled.source
    assert 'active end page number' in compiled.source
    assert 'active document' not in compiled.source.lower()
    assert 'quit' not in compiled.source.lower()
    assert compiled.nodes


def test_unsupported_equation_fails_before_script_creation(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, MacWordCapabilityError
    build = build_longform_generation('# 标题\n\n$$\nx^2\n$$')
    with pytest.raises(MacWordCapabilityError, match='equation'):
        compile_plan(build.plan, {}, tmp_path/'owned.docx', timeout=10)


def test_native_pagination_parser_rejects_missing_duplicate_invalid_records(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import parse_result
    expected = {'n1': ('b1', 'body')}
    result = parse_result('WPSC_NODE\tn1\t1\t2\t0\t20\nWPSC_OK\t0\n', expected, tmp_path/'x.docx', 3)
    assert result.pagination_map.version == 'M5-v1'
    assert [x.page for x in result.pagination_map.nodes[0].fragments] == [1, 2]
    for raw in ('WPSC_OK\t0\n', 'WPSC_NODE\tn1\t0\t2\t0\t20\nWPSC_OK\t0\n', 'WPSC_NODE\tn1\t1\t2\t0\t20\n'*2+'WPSC_OK\t0\n'):
        with pytest.raises(ValueError):
            parse_result(raw, expected, tmp_path/'x.docx', 3)


def test_tables_and_lists_compile_native_objects(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('# Report\n\n## Chapter\n\n- Alpha\n- Beta\n\n| A | B |\n|---|---|\n| one | two |')
    c=compile_plan(b.plan, {}, tmp_path/'x.docx', timeout=30)
    assert 'make new table' in c.source
    assert 'heading format of row 1' in c.source
    assert 'apply bullet default' in c.source


def test_figure_compiles_only_bound_normalized_image(tmp_path):
    from PIL import Image
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, MacWordCapabilityError
    p=tmp_path/'image.png'
    Image.new('RGB', (640, 480), 'red').save(p)
    b=build_longform_generation('# Report\n\n## Chapter\n\n:::figure {caption="a chart"}\n![a chart](image.png)\n:::\n', base_dir=str(tmp_path))
    resources={r.resource_id:p for r in b.preflight.resources}
    c=compile_plan(b.plan, resources, tmp_path/'x.docx', timeout=30)
    assert 'make new inline picture' in c.source
    assert 'save with document:true' in c.source
    with pytest.raises(MacWordCapabilityError, match='resource'):
        compile_plan(b.plan, {}, tmp_path/'x.docx', timeout=30)


def test_bibliography_and_citation_preserve_resolved_text(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    md='''# Report

## Chapter

Source [@alpha].

:::bibliography
alpha: Research publication.
:::
'''
    b=build_longform_generation(md)
    c=compile_plan(b.plan,{},tmp_path/'x.docx',timeout=20)
    assert 'Research publication' in c.source


def test_preflight_rejects_unknown_paragraph_style_and_unsafe_text(tmp_path):
    from dataclasses import replace
    from skills.WPSComposer.scripts.generation_plan import GenerationOperation
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, MacWordCapabilityError
    b=build_longform_generation('# Report\n\nBody')
    ops=list(b.plan.operations)
    index=next(i for i,o in enumerate(ops) if o.op == 'writer.add_paragraph')
    ops[index]=replace(ops[index],args={**dict(ops[index].args),'style':'ArbitraryStyle'})
    with pytest.raises((ValueError,MacWordCapabilityError)):
        compile_plan(replace(b.plan,operations=tuple(ops)),{},tmp_path/'x.docx',timeout=10)


def test_header_reference_is_consumed_before_footer_is_acquired(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('# Report\n\nBody')
    c=compile_plan(b.plan,{},tmp_path/'x.docx',timeout=20)
    start=c.source.index('set ownHeader to get header')
    assert c.source.index('set link to previous of ownHeader',start)<c.source.index('set ownFooter to get footer',start)


def test_native_snapshot_enumerates_document_indices_not_lazy_collection(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import wrap_owned
    source=wrap_owned('save ownedDoc', tmp_path/'owned.docx',20)
    assert 'repeat with documentIndex from 1 to (count of documents)' in source
    assert 'repeat with d in documents' not in source


def test_relayout_compacts_terminal_paragraph(tmp_path):
    from skills.WPSComposer.scripts.longform.platform_runtime import _apply_relayout
    from skills.WPSComposer.scripts.longform.relayout import RelayoutDirective
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('# Report\n\nBody')
    derived=_apply_relayout(b.plan,(RelayoutDirective('remove-unexpected-blank',None,{'maximumPages':1}),))
    c=compile_plan(derived,{},tmp_path/'x.docx',timeout=20)
    assert 'font size of font object of terminalRange to 1' in c.source


def test_unsupported_style_attribute_rejected(tmp_path):
    from dataclasses import replace
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, MacWordCapabilityError
    b=build_longform_generation('# Report\n\nBody')
    ops=[]
    for o in b.plan.operations:
        if o.op=='writer.ensure_styles':
            styles=[dict(s) for s in o.args['styles']];styles[0]['color']='#ff0000'
            o=replace(o,args={'styles':styles})
        ops.append(o)
    with pytest.raises(MacWordCapabilityError,match='color'):
        compile_plan(replace(b.plan,operations=tuple(ops)),{},tmp_path/'x.docx',timeout=20)


def test_missing_image_degradation_is_visible_once_and_reported(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('# Report\n\n## Section\n\n:::figure {caption="Missing source"}\n![important image](missing.png)\n:::\n',base_dir=str(tmp_path))
    c=compile_plan(b.plan,{},tmp_path/'x.docx',timeout=20)
    assert any(i.code == 'RESOURCE_NOT_FOUND' for i in c.issues)
    assert '[RESOURCE_NOT_FOUND] [RESOURCE_NOT_FOUND]' not in c.source


def test_footer_page_field_replaces_owned_placeholder_not_paragraph(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('---\ntitle: Report\nauthor: Test\ntoc: true\n---\n# Report\n\n## Chapter\n\nBody')
    source=compile_plan(b.plan,{},tmp_path/'x.docx',timeout=20).source
    assert 'set r to character 1 of text object of ownFooter' in source
    assert 'collapse range (text object of ownFooter)' not in source


def test_sequence_transparent_heading_uses_body_outline(tmp_path):
    from dataclasses import replace
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    b=build_longform_generation('# Report\n\n## Chapter\n\nBody')
    ops=[]
    for o in b.plan.operations:
        if o.op == 'writer.add_heading':
            o=replace(o,args={**dict(o.args),'level':1,'numbering':False,'sequenceTransparent':True})
        ops.append(o)
    c=compile_plan(replace(b.plan,operations=tuple(ops)),{},tmp_path/'x.docx',timeout=20)
    assert 'set outline level of paragraph format of r to outline level body text' in c.source


@pytest.mark.parametrize('collection, item', [
    ('tables of contents', 'table of contents'),
    ('tables of figures', 'table of figures'),
    ('indexes', 'index'),
])
def test_refresh_rebuilds_every_native_index_inside_each_convergence_round(collection, item):
    import re
    from skills.WPSComposer.scripts.msoffice.macos_script import refresh_source
    source = refresh_source()
    round_start = source.index('repeat with refreshRound from 1 to 3')
    ordinary_fields = source.index('repeat with fieldIndex')
    loop = re.search(r'repeat with (\w+) from 1 to \(count of ' + re.escape(collection) +
                     r' of ownedDoc\)\s+update \(' + re.escape(item) + r' \1 of ownedDoc\)\s+end repeat', source)
    assert loop is not None, 'Generic update field can report success while the native index remains stale'
    assert round_start < loop.start() < loop.end() < ordinary_fields
    assert 'update page numbers' not in source
    assert 'try' not in source, 'Dedicated update errors must propagate to owned-document cleanup'
    assert source.index('set fieldState') < source.index('if fieldState is priorFields')
    assert 'if fieldsStable is false then error' in source


def test_compiler_saves_refreshed_indexes_and_real_pagination_before_owned_close(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, refresh_source, pagination_source
    build = build_longform_generation('---\ntitle: Report\ntoc: true\n---\n# Chapter\n\nBody')
    compiled = compile_plan(build.plan, {}, tmp_path/'owned.docx', timeout=20)
    refresh = compiled.source.index(refresh_source())
    assert 'update (table of contents' in refresh_source()
    pagination = compiled.source.index(pagination_source(compiled.nodes), refresh)
    saved = compiled.source.rindex('save as ownedDoc file name')
    closed = compiled.source.index('close document closeName saving no')
    assert refresh < pagination < saved < closed


@pytest.mark.parametrize('scheme', ['decimal', 'chinese-formal', 'hybrid-bid'])
def test_heading_levels_bind_localized_styles_to_one_owned_outline(tmp_path, scheme):
    from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
    build = build_longform_generation(
        f'---\nheading_numbering: {scheme}\n---\n# Report\n\n'
        '## First\n\n### Second\n\n#### Third\n\n##### Fourth\n\nBody')
    source = compile_plan(build.plan, {}, tmp_path/'owned.docx', timeout=20).source
    assert source.count('make new list template at ownedDoc') == 1
    for level in range(1, 5):
        start = source.index(f'set lvl to list level {level} of ownList')
        localized_name = f'set headingName to (name local of (Word style (style heading{level}) of ownedDoc)) as text'
        binding = source.index(localized_name, start)
        assert source.index('set linked style of lvl to headingName', binding) > binding
    # Word clones the outline for each style-directed link, splitting chapter
    # counters. Range-level apply can mask this but destroys paragraph breaks.
    assert 'link to list template (' not in source
    assert 'apply list format template' not in source
    assert 'set lvl to list level 5 of ownList' not in source
