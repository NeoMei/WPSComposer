from types import SimpleNamespace

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
