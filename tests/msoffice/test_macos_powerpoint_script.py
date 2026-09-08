import math
from pathlib import Path
import pytest
from skills.WPSComposer.scripts.generation_plan import GenerationPlan, GenerationOperation


def plan(*ops):
    return GenerationPlan('presentation', tuple(GenerationOperation(op, args) for op, args in ops))


def compiler():
    from skills.WPSComposer.scripts.msoffice import macos_powerpoint_script
    return macos_powerpoint_script


def test_title_generation_is_native_and_binds_saved_path_before_close(tmp_path):
    c = compiler()
    s = c.compile_plan(plan(('slide.reset', {}), ('slide.add_title', {'title': 'Title "中文"\nLine', 'subtitle': 'Subtitle'})), {}, tmp_path/'unique.pptx', tmp_path/'quality.pdf')
    assert 'make new slide at end of ownedDoc' in s
    assert 'Title \\"中文\\"" & linefeed & "Line' in s
    assert s.index('Owned output path mismatch') < s.index('close ownedDoc saving no')
    assert s.index('close ownedDoc saving no') < s.index('return "WPSCOMPOSER_MS_OFFICE_OK:presentation"')
    assert 'quit' not in s.lower()
    assert 'active presentation' not in s.lower()


@pytest.mark.parametrize('timeout', [0, -1, True, float('inf'), float('nan')])
def test_invalid_timeout_fails_before_native_compilation(tmp_path, timeout):
    with pytest.raises(ValueError):
        compiler().compile_plan(plan(('slide.add_blank', {})), {}, tmp_path/'x.pptx', timeout=timeout)


def test_unbound_image_and_stale_slide_are_rejected(tmp_path):
    c = compiler()
    with pytest.raises(ValueError, match='resource'):
        c.compile_plan(plan(('slide.add_blank', {}), ('slide.add_image', {'slide':1,'imageId':'i','left':0,'top':0})), {}, tmp_path/'x.pptx')
    with pytest.raises(ValueError, match='slide'):
        c.compile_plan(plan(('slide.add_image', {'slide':1,'imageId':'i','left':0,'top':0})), {'i':tmp_path/'staged.png'}, tmp_path/'x.pptx')


def test_paths_need_not_exist_but_must_be_absolute(tmp_path):
    c = compiler()
    s = c.compile_plan(plan(('slide.add_blank', {}), ('slide.add_image', {'slide':1,'imageId':'i','left':4,'top':8,'width':30})), {'i':tmp_path/'not-created.png'}, tmp_path/'not-created.pptx')
    assert 'make new picture' in s
    with pytest.raises(ValueError, match='absolute'):
        c.compile_plan(plan(('slide.add_blank', {})), {}, Path('relative.pptx'))


def test_initial_arbitrary_size_compiles_verified_orientation_swap_sequence(tmp_path):
    c = compiler()
    source = c.compile_plan(
        plan(('slide.set_size', {'width':720,'height':405}), ('slide.add_blank', {})),
        {}, tmp_path/'x.pptx')
    commands = [
        'if (count of slides of ownedDoc) is not 0 then return "EXISTING_SLIDES"',
        'set slide orientation of page setup of ownedDoc to vertical orientation',
        'set slide width of page setup of ownedDoc to 405',
        'set slide orientation of page setup of ownedDoc to horizontal orientation',
        'set slide width of page setup of ownedDoc to 720',
    ]
    assert [source.index(command) for command in commands] == sorted(source.index(command) for command in commands)
    assert 'slide height' not in source

    tall = c.compile_initial_slide_size(405, 720)
    assert tall.splitlines() == [
        commands[0],
        'set slide orientation of page setup of ownedDoc to horizontal orientation',
        'set slide width of page setup of ownedDoc to 720',
        'set slide orientation of page setup of ownedDoc to vertical orientation',
        'set slide width of page setup of ownedDoc to 405',
    ]


def test_arbitrary_size_after_content_is_rejected_before_native_compilation(tmp_path):
    c = compiler()
    with pytest.raises(c.MacPowerPointCapabilityError, match='before adding slides'):
        c.compile_plan(
            plan(('slide.add_blank', {}), ('slide.set_size', {'width':800,'height':600})),
            {}, tmp_path/'x.pptx')


@pytest.mark.parametrize('width,height', [
    (True, 405), (720, False), (0, 405), (720, -1),
    (math.inf, 405), (720, math.nan), ('720', 405),
])
def test_initial_size_script_rejects_invalid_dimensions(width, height):
    with pytest.raises(ValueError):
        compiler().compile_initial_slide_size(width, height)


def test_existing_540_height_sequence_is_unchanged(tmp_path):
    source = compiler().compile_plan(
        plan(('slide.add_blank', {}), ('slide.set_size', {'width':960,'height':540})),
        {}, tmp_path/'x.pptx')
    assert 'set slide size of page setup of ownedDoc to slide size on screen' in source
    assert 'set slide width of page setup of ownedDoc to 960' in source
    assert 'Arbitrary initial slide size requires an empty presentation' not in source


def test_conversion_requires_macro_free_native_source_and_distinct_output(tmp_path):
    c = compiler()
    for suffix in ['.pptm', '.potm', '.ppsm', '.ppt']:
        with pytest.raises(c.MacPowerPointCapabilityError):
            c.compile_conversion(tmp_path/('input'+suffix), tmp_path/'out.pdf')
    s = c.compile_conversion(tmp_path/'unique.pptx', tmp_path/'out.pdf')
    assert 'set sourceFile to (POSIX file sourcePath) as alias' in s
    assert 'open sourceFile' in s
    assert 'do shell script' not in s
    assert 'Owned source path mismatch' in s
    assert s.index('Existing presentation collision') < s.index('open sourceFile')


def test_native_image_with_one_dimension_preserves_aspect_ratio(tmp_path):
    s = compiler().compile_plan(plan(('slide.add_blank', {}), ('slide.add_image', {'slide':1,'imageId':'i','left':4,'top':8,'width':30})), {'i':tmp_path/'staged.png'}, tmp_path/'x.pptx')
    assert 'set lock aspect ratio of currentPicture to true' in s
    assert 'set width of currentPicture to 30' in s


def test_preset_changes_master_and_default_font_sizes(tmp_path):
    preset = {'name':'sample', 'colors':{'primary':'#123456','dark':'#111111','background':'#FFFFFF'},
              'fonts':{'title':{'family':'Arial','size':44,'color':'#123456'}, 'body':{'family':'Arial','size':22,'color':'#111111'}, 'subtitle':{'family':'Arial','size':24,'color':'#111111'}}}
    s = compiler().compile_plan(plan(('slide.add_blank',{}), ('slide.apply_preset',{'preset':preset}), ('slide.add_title',{'title':'Title','subtitle':'Sub','titleSize':40,'subtitleSize':20}), ('slide.add_bullets',{'title':'Body','items':['a','b'],'bodySize':18})),{},tmp_path/'x.pptx')
    assert 'background of slide master of ownedDoc' in s
    assert 'to 24' in s
    assert 'to 22' in s
