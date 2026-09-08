from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice.windows_host import NativeWordComposer
from skills.WPSComposer.scripts.writer import WriterComposer


def test_word_heading_styles_share_one_list_template():
    styles = {i: SimpleNamespace(NameLocal='本地标题 ' + str(i), template=None) for i in range(1, 5)}
    for style in styles.values():
        # Word's Style.LinkToListTemplate clones the supplied template.
        style.LinkToListTemplate = lambda template, level, style=style: setattr(style, 'template', object())
    class Level:
        @property
        def LinkedStyle(self):
            return self._name
        @LinkedStyle.setter
        def LinkedStyle(self, name):
            self._name = name
            next(style for style in styles.values() if style.NameLocal == name).template = template
    levels = {i: Level() for i in styles}
    template = SimpleNamespace(ListLevels=lambda i: levels[i])
    composer = object.__new__(NativeWordComposer)
    composer._doc = SimpleNamespace(Styles=lambda identifier: styles[-identifier - 1])
    composer._link_heading_list_styles(template)
    assert all(style.template is template for style in styles.values())
    assert [levels[i].LinkedStyle for i in styles] == [styles[i].NameLocal for i in styles]


def test_wps_keeps_its_supported_style_binding_path():
    bound = {}
    styles = {i: SimpleNamespace(LinkToListTemplate=lambda template, level, i=i: bound.update({i: (template, level)})) for i in range(1, 5)}
    composer = object.__new__(WriterComposer)
    composer._doc = SimpleNamespace(Styles=lambda identifier: styles[-identifier - 1])
    template = object()  # No ListLevels.LinkedStyle API required by the WPS path.
    composer._link_heading_list_styles(template)
    assert bound == {i: (template, i) for i in range(1, 5)}


@pytest.mark.parametrize('composer_type', [WriterComposer, NativeWordComposer])
@pytest.mark.parametrize('scheme,formats,styles', [
    ('decimal', ['%1', '%1.%2', '%1.%2.%3', '%1.%2.%3.%4'], [0, 0, 0, 0]),
    ('chinese-formal', ['第%1章', '第%2节', '%3、', '（%4）'], [37, 37, 37, 37]),
    ('hybrid-bid', ['第%1章', '%1.%2', '%1.%2.%3', '关键工法%4：'], [37, 253, 253, 22]),
])
def test_native_heading_scheme_contract(composer_type, scheme, formats, styles):
    levels = {i: SimpleNamespace() for i in range(1, 5)}
    template = SimpleNamespace(ListLevels=lambda i: levels[i])
    allocations = []
    def allocate(outline):
        allocations.append(outline)
        return template
    native_range = SimpleNamespace(ListFormat=SimpleNamespace(ListString='native number'))
    composer = object.__new__(composer_type)
    composer._native_position = lambda: 0
    composer.add_heading_level = lambda text, level: None
    composer._doc = SimpleNamespace(
        ListTemplates=SimpleNamespace(Add=allocate),
        Styles=lambda identifier: SimpleNamespace(
            NameLocal=str(identifier), LinkToListTemplate=lambda template, level: None),
        Range=lambda start, end: native_range,
    )
    for chapter in range(2):
        for level in range(1, 5):
            composer.add_heading_level_native('Chapter heading', level, numbering=True, scheme=scheme)
    assert allocations == [True], 'Both chapters must use the same outline template'
    assert [levels[i].NumberFormat for i in levels] == formats
    assert [levels[i].NumberStyle for i in levels] == styles
    assert [levels[i].ResetOnHigher for i in levels] == [0, 1, 2, 3]
    assert [levels[i].StartAt for i in levels] == [1, 1, 1, 1]
