"""Portable failure injection for the opt-in Excel acceptance fixture."""
from types import SimpleNamespace

import pytest

from fixtures.microsoft_parity import macos_excel_delete_selection as fixture


@pytest.mark.parametrize('matches_marker', [True, False])
def test_creation_timeout_cleanup_requires_exact_added_sentinel_marker(
    tmp_path, monkeypatch, matches_marker
):
    source = tmp_path / 'source.xlsx'
    source.write_bytes(b'not opened in the injected creation failure')
    marker = 'WPSC-DELETE-SENTINEL-' + 'a' * 32
    monkeypatch.setattr(fixture, 'uuid4', lambda: SimpleNamespace(hex='a' * 32))
    original = {'name': 'Untouched', 'path': 'Untouched', 'saved': False}
    added = {'name': 'NewBook', 'path': 'NewBook', 'saved': False}
    closed = []

    def inventory(output, label):
        if label == 'inventory-before' or closed:
            return [original]
        return [original, added]

    def native(output, label, body):
        if label == 'create-sentinel':
            raise TimeoutError('creation acknowledgement lost')
        if label.startswith('verify-'):
            assert 'workbook "NewBook"' in body
            return 'false\x1f' + (marker if matches_marker else 'foreign content')
        assert label == 'close-exact-sentinel'
        assert matches_marker
        assert marker in body and 'Sentinel marker changed' in body
        closed.append('NewBook')
        return 'closed exact sentinel'

    monkeypatch.setattr(fixture, '_inventory', inventory)
    monkeypatch.setattr(fixture, '_osascript', native)
    monkeypatch.setattr(fixture, 'open_document', lambda *a, **kw: pytest.fail('must stop after failed creation'))
    result = fixture.run(source, tmp_path / 'evidence')

    assert result['passed'] is False  # Cleanup cannot turn a failed run into PASS.
    assert result['error']['type'] == 'TimeoutError'
    assert closed == (['NewBook'] if matches_marker else [])
    assert result['inventory_restored'] is matches_marker
    if matches_marker:
        assert result['unacknowledged_sentinel_identified'] is True
    else:
        assert 'sentinel_identity_unresolved' in result
