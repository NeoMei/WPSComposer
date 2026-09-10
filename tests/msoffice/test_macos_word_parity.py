"""Closed Mac Word advanced operations: compilation and native fixture contracts."""
from dataclasses import replace
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan, MacWordCapabilityError


def change(build, name, **args):
    return replace(build.plan, operations=tuple(
        replace(o, args={**dict(o.args), **args}) if o.op == name else o
        for o in build.plan.operations))


def compile_changed(tmp_path, markdown, name, **args):
    build = build_longform_generation(markdown)
    return compile_plan(change(build, name, **args), {}, tmp_path/'owned.docx', timeout=30)


def test_equation_builds_one_native_math_inside_numbered_owned_container(tmp_path):
    build = build_longform_generation('# Report\n\n$$\nx^2+\\frac{a}{b}\n$$\n\nAFTER-EQUATION')
    compiled = compile_plan(build.plan, {}, tmp_path/'owned.docx', timeout=30)
    assert 'create new equation from range r in document ownedDoc' in compiled.source
    assert 'build up ownMath' in compiled.source
    assert 'nativeMathCount + 1' in compiled.source
    assert 'number of columns:3' in compiled.source
    equation = next(o for o in build.plan.operations if o.op == 'writer.add_equation')
    assert compiled.nodes[equation.node_id][0] != equation.args['bookmarkName']
    assert 'field sequence' in compiled.source and 'WPSC_EQ' in compiled.source
    assert not compiled.issues


def test_unsafe_or_unknown_math_is_rejected_without_embedded_source(tmp_path):
    build = build_longform_generation('# Report\n\n$$\nx^2\n$$')
    plan = change(build, 'writer.add_equation', source='\\unsupported PRIVATE-SOURCE', fallbackText='redacted')
    with pytest.raises(MacWordCapabilityError, match='equation') as error:
        compile_plan(plan, {}, tmp_path/'owned.docx', timeout=30)
    assert 'PRIVATE-SOURCE' not in str(error.value)


def test_merge_executes_after_grid_formatting_and_keeps_outside_cells(tmp_path):
    compiled = compile_changed(tmp_path,
        '# Report\n\n| A | B | C |\n|---|---|---|\n| MERGE | | OUTSIDE |',
        'writer.add_semantic_table', merges=[{'top': 2, 'left': 1, 'bottom': 2, 'right': 2}])
    assert compiled.source.index('OUTSIDE') < compiled.source.index('merge cell mergeStart with mergeEnd')
    assert compiled.source.index('heading format of row 1') < compiled.source.index('merge cell mergeStart with mergeEnd')


def test_landscape_media_restores_portrait_without_restarting_numbering(tmp_path):
    compiled = compile_changed(tmp_path,
        '# Report\n\n## Table section\n\n| A | B |\n|---|---|\n| a | b |\n\nAFTER-LANDSCAPE',
        'writer.add_semantic_table', orientation='landscape', includePreviousHeading=True, continuousExit=True)
    assert 'orient landscape' in compiled.source and 'orient portrait' in compiled.source
    assert 'pendingHeadingStart' in compiled.source
    assert 'section break continuous' in compiled.source
    assert 'set restart numbering at section of page number options of mediaFooter to false' in compiled.source


def test_cell_degradation_is_visible_once_and_styled(tmp_path):
    fallback = '[REFERENCE_UNRESOLVED 引用目标未解析]'
    compiled = compile_changed(tmp_path, '# Report\n\n| A | B |\n|---|---|\n| '+fallback+' | outside |',
        'writer.add_semantic_table', cellDegradations=[{'row':2,'column':1,'code':'REFERENCE_UNRESOLVED','fallbackText':fallback}])
    assert compiled.source.count(fallback) == 1
    assert 'background pattern color of shading of text object of ownCell' in compiled.source
    assert any(i.code == 'REFERENCE_UNRESOLVED' for i in compiled.issues)


def test_custom_style_color_indents_and_spacing_are_faithfully_emitted(tmp_path):
    build = build_longform_generation('# Report\n\nFORMATTED')
    styles = [dict(s) for o in build.plan.operations if o.op == 'writer.ensure_styles' for s in o.args['styles']]
    styles.append({'name':'TaskColored','basedOn':'Body Text','type':'paragraph','color':'#115599','shading':'#E6F0FA',
                   'underline':True,'strikethrough':True,'leftIndent':18,'rightIndent':9,'lineSpacing':20,'lineSpacingRule':'exactly',
                   'leftBorder':True,'borderColor':'#115599'})
    plan = change(build, 'writer.ensure_styles', styles=styles)
    plan = replace(plan, operations=tuple(replace(o,args={**dict(o.args),'style':'TaskColored'}) if o.op=='writer.add_paragraph' else o for o in plan.operations))
    compiled = compile_plan(plan, {}, tmp_path/'owned.docx', timeout=30)
    assert 'make new Word style' in compiled.source
    assert '{4369, 21845, 39321}' in compiled.source  # AppleScript RGB is 16-bit
    assert 'paragraph format left indent' in compiled.source and 'paragraph format right indent' in compiled.source
    assert 'line space exactly' in compiled.source
    assert 'set line spacing of paragraph format of ownStyle to 20' in compiled.source


@pytest.mark.parametrize('rule', ['single', 'one_and_half', 'double', 'exactly', 'at_least', 'multiple'])
def test_all_closed_spacing_rules_compile_without_value_guessing(tmp_path, rule):
    compiled = compile_changed(tmp_path, '# Report\n\nBody', 'writer.add_paragraph', lineSpacing=2, lineSpacingRule=rule)
    expected = {'single':'single','one_and_half':'1 pt5','double':'double','exactly':'exactly','at_least':'at least','multiple':'multiple'}[rule]
    assert 'line space' + ('' if expected.startswith('1') else ' ') + expected in compiled.source


def test_unknown_spacing_rule_rejects_before_native_script(tmp_path):
    with pytest.raises(MacWordCapabilityError, match='lineSpacingRule'):
        compile_changed(tmp_path, '# Report\n\nBody', 'writer.add_paragraph', lineSpacing=12, lineSpacingRule='unknown')


def test_two_column_figure_uses_native_cell_bound_pictures(tmp_path):
    from PIL import Image
    for name in ('left.png','right.png'):
        Image.new('RGB',(120,80),'red').save(tmp_path/name)
    build = build_longform_generation('# Report\n\n:::figure {caption="Two pictures" layout="columns"}\n![Left](left.png)\n![Right](right.png)\n:::\n', base_dir=str(tmp_path))
    plan = change(build, 'writer.add_captioned_figure', layout='columns', columns=2)
    resources = {r.resource_id:tmp_path/'left.png' for r in build.preflight.resources}
    compiled = compile_plan(plan, resources, tmp_path/'owned.docx', timeout=30)
    assert 'number of columns:2' in compiled.source
    assert compiled.source.count('make new inline picture at beginning of text object of pictureCell') == 2


def test_footer_text_and_page_field_both_survive_configuration(tmp_path):
    compiled = compile_changed(tmp_path, '# Report\n\nBody', 'writer.configure_section',
        footerText='CONFIDENTIAL', pageNumberFormat='arabic', linkToPreviousFooter=False)
    assert 'CONFIDENTIAL' in compiled.source and 'field type field page' in compiled.source


def test_configure_page_and_section_measurements_are_not_dropped(tmp_path):
    build = build_longform_generation('# Report\n\nBody')
    plan = change(build, 'writer.configure_page', pageWidth=612,pageHeight=792,landscape=False,columns=2,header='HEAD',footer='FOOT')
    plan = replace(plan, operations=tuple(replace(o,args={**dict(o.args),'pageSize':'Letter','margins':dict(top=50,bottom=51,left=52,right=53)}) if o.op=='writer.configure_section' else o for o in plan.operations))
    source = compile_plan(plan,{},tmp_path/'owned.docx',timeout=30).source
    assert 'page width of page setup of ownedDoc to 612' in source
    assert 'number of columns 2' in source
    assert 'top margin of page setup of ownSection to 50' in source
    assert 'HEAD' in source and 'FOOT' in source


def test_spans_have_independent_native_ranges_and_formats(tmp_path):
    source = compile_changed(tmp_path,'# Report\n\nBody','writer.add_paragraph',text='ABCD',spans=[dict(text='AB',bold=True),dict(text='CD',italic=True,strikethrough=True)]).source
    assert 'start (nodeStart + 0) end (nodeStart + 2)' in source
    assert 'start (nodeStart + 2) end (nodeStart + 4)' in source
    assert 'strike through of font object of spanRange to true' in source


def m4_equation_plan(content):
    build = build_longform_generation('# Report\n\n$$\nx+y\n$$')
    operations = []
    for op in build.plan.operations:
        if op.op == 'writer.add_equation':
            args = {k:v for k,v in op.args.items() if k != 'source'}
            args.update(renderMode='native-m4',content=content)
            op = replace(op,args=args,failure_policy={'mode':'degrade','recoverableCodes':['EQUATION_INSERT_FAILED'],'fallback':'explicit-image-then-source-notice'})
        operations.append(op)
    return replace(build.plan,operations=tuple(operations))


def test_planned_equation_source_fallback_keeps_numeric_bookmark(tmp_path):
    descriptor = dict(code='FORMULA_MALFORMED',placement='block',objectLabel='formula',reason='Formula could not be emitted',fallbackText='x+y',fallbackKind='source')
    plan = m4_equation_plan({'plannedDegradation':descriptor})
    compiled = compile_plan(plan,{},tmp_path/'owned.docx',timeout=30)
    assert 'create new equation' not in compiled.source
    assert 'field sequence' in compiled.source and 'WPSC_EQ' in compiled.source
    assert 'FORMULA_MALFORMED' in compiled.source
    assert any(i.code=='FORMULA_MALFORMED' for i in compiled.issues)


def test_m4_native_failure_has_explicit_source_fallback_and_runtime_issue(tmp_path):
    plan = m4_equation_plan({'nativeMath':{'syntax':'wps-linear-v1','linearText':'x+y','sourceHash':'a'*64}})
    compiled = compile_plan(plan,{},tmp_path/'owned.docx',timeout=30)
    assert 'WPSC_EQUATION_FALLBACK' in compiled.source
    assert 'set content of text object of formulaCell to' in compiled.source
    assert 'EQUATION_INSERT_FAILED' in compiled.source


def test_runtime_equation_fallback_is_reported_and_records_are_closed(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_script import parse_result
    nodes={'eq:one':('bookmark','body')}
    raw='WPSC_EQUATION_FALLBACK\teq:one\nWPSC_NODE\teq:one\t1\t1\t0\t20\nWPSC_OK\t0'
    outcome=parse_result(raw,nodes,tmp_path/'owned.docx',1)
    assert len(outcome.issues)==1 and outcome.issues[0].code=='EQUATION_INSERT_FAILED'
    with pytest.raises(ValueError,match='fallback record'):
        parse_result(raw.replace('WPSC_EQUATION_FALLBACK\teq:one','WPSC_EQUATION_FALLBACK\tunknown'),nodes,tmp_path/'owned.docx',1)
    with pytest.raises(ValueError,match='fallback record'):
        parse_result('WPSC_EQUATION_FALLBACK\teq:one\n'+raw,nodes,tmp_path/'owned.docx',1)


def test_formula_reserves_body_width_for_math_instead_of_equal_thirds(tmp_path):
    build=build_longform_generation('# Report\n\n$$\nx^2\n$$')
    source=compile_plan(build.plan,{},tmp_path/'owned.docx',timeout=30).source
    assert 'set width of column 1 of formulaTable to 36' in source
    assert 'set width of column 2 of formulaTable to formulaWidth - 72' in source
    assert 'set width of column 3 of formulaTable to 36' in source


def test_native_fixture_rejects_the_previously_false_passing_cropped_columns():
    from fixtures.microsoft_parity.macos_word_advanced import inspect_package
    evidence=Path(__file__).resolve().parents[2]/'docs/verification/microsoft-parity/macos-word-advanced'
    bad=evidence/'integrated-04/advanced.docx'
    good=evidence/'integrated-08/advanced.docx'
    if not bad.is_file() or not good.is_file():
        pytest.skip('Native evidence artifacts are not present in this source checkout')
    assert not inspect_package(bad)['pictures_in_two_real_cells']
    assert inspect_package(good)['pictures_in_two_real_cells']


def test_span_text_cannot_format_outside_its_paragraph(tmp_path):
    with pytest.raises(MacWordCapabilityError,match='span text'):
        compile_changed(tmp_path,'# Report\n\nBody','writer.add_paragraph',text='AB',spans=[dict(text='ABC',bold=True)])


@pytest.mark.parametrize('value',[-1,4,100])
def test_invalid_alignment_is_rejected_before_native_word(tmp_path,value):
    with pytest.raises(MacWordCapabilityError,match='alignment'):
        compile_changed(tmp_path,'# Report\n\nBody','writer.add_paragraph',align=value)
