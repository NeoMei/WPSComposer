from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from skills.WPSComposer.scripts.writer import NativeWriterObjectError, WriterComposer


class _PageNumbers:
    """Model Word's observed Boolean setter: VT_I4 -1 is silently ignored."""

    def __init__(self):
        self._restart = False
        self._start = 0
        self.NumberStyle = 0

    @property
    def RestartNumberingAtSection(self):
        return self._restart

    @RestartNumberingAtSection.setter
    def RestartNumberingAtSection(self, value):
        if isinstance(value, bool):
            self._restart = value

    @property
    def StartingNumber(self):
        return self._start

    @StartingNumber.setter
    def StartingNumber(self, value):
        if self._restart:
            self._start = value


def _writer():
    fields = Mock(Count=0)
    footer = SimpleNamespace(PageNumbers=_PageNumbers(), Range=Mock(Fields=fields), LinkToPrevious=False)
    footer.Range.Duplicate = footer.Range
    header = SimpleNamespace(Range=Mock(), LinkToPrevious=False)
    section = SimpleNamespace(Footers=lambda index: footer, Headers=lambda index: header)
    writer = object.__new__(WriterComposer)
    writer._doc = SimpleNamespace(Sections=Mock(return_value=section, Count=1))
    return writer, footer, header


def test_page_number_restart_uses_com_boolean_and_sets_requested_start():
    writer, footer, _ = _writer()
    writer.set_page_numbering('roman', start=1, restart=True)
    assert footer.PageNumbers.RestartNumberingAtSection is True
    assert footer.PageNumbers.StartingNumber == 1
    assert footer.PageNumbers.NumberStyle == 2


def test_section_detaches_footer_before_inserting_page_field():
    writer, footer, _ = _writer()
    footer.LinkToPrevious = True
    previous_footer = []
    footer.Range.Fields.Add.side_effect = lambda *args: previous_footer.append('PAGE') if footer.LinkToPrevious else None
    writer.set_page_role = Mock()
    writer._current_section_page_setup = Mock(return_value=SimpleNamespace())
    writer.configure_section(role='toc', page_number_format='roman', restart_page_numbering=True,
                             start_page_number=1, link_to_previous_footer=False)
    assert previous_footer == [], 'A new section must not add PAGE to the preceding cover'


def test_explicit_empty_footer_clears_inherited_content():
    writer, footer, _ = _writer()
    footer.Range.Text = 'Inherited footer text'
    writer.set_header_footer(footer='', link_to_previous_footer=False)
    assert footer.Range.Text == ''


@pytest.mark.parametrize('property_name, ignored_value', [
    ('RestartNumberingAtSection', False), ('StartingNumber', 0), ('NumberStyle', 0),
])
def test_silently_ignored_page_number_setting_is_an_execution_failure(property_name, ignored_value):
    writer, footer, _ = _writer()
    class IgnoredSetting(_PageNumbers):
        def __setattr__(self, name, value):
            if name != property_name:
                super().__setattr__(name, value)
        def __getattribute__(self, name):
            return ignored_value if name == property_name else super().__getattribute__(name)
    footer.PageNumbers = IgnoredSetting()
    with pytest.raises(NativeWriterObjectError, match='page number'):
        writer.set_page_numbering('roman', start=1, restart=True)


def test_page_field_failure_is_not_swallowed():
    writer, footer, _ = _writer()
    footer.Range.Fields.Add.side_effect = RuntimeError('host refused PAGE field')
    with pytest.raises(NativeWriterObjectError, match='page number'):
        writer.set_page_numbering('arabic', start=1, restart=True)


def test_existing_page_field_is_not_duplicated():
    writer, footer, _ = _writer()
    footer.Range.Fields.Count = 1
    footer.Range.Fields.return_value.Type = 33
    writer.set_page_numbering('arabic', start=1, restart=True)
    footer.Range.Fields.Add.assert_not_called()


def test_cover_numbering_none_removes_page_field_content():
    writer, footer, _ = _writer()
    footer.Range.Text = 'PAGE'
    writer.set_page_numbering('none')
    assert footer.Range.Text == ''
    footer.Range.Fields.Add.assert_not_called()


def test_failed_footer_unlink_stops_before_any_content_mutation():
    writer, footer, _ = _writer()
    class RefusedUnlink:
        Range = footer.Range
        @property
        def LinkToPrevious(self):
            return True
        @LinkToPrevious.setter
        def LinkToPrevious(self, value):
            raise RuntimeError('unlink refused')
    section = SimpleNamespace(Footers=lambda index: RefusedUnlink(), Headers=Mock())
    writer._doc.Sections.return_value = section
    with pytest.raises(NativeWriterObjectError, match='header/footer'):
        writer.set_header_footer(footer='New text', link_to_previous_footer=False)
    assert footer.Range.Text != 'New text'
