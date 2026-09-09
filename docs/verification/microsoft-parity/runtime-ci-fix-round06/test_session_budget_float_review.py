"""Read-only review repro: real worker, original tests, isolated proxy clock."""
import os
from types import SimpleNamespace

import pytest
import test_windows_session_proxy as existing
from test_windows_session_proxy import transport


@pytest.mark.parametrize('tick', [1558.546, 424.005])
@pytest.mark.parametrize('scenario', [
    existing.test_actual_worker_serve_subprocess_preserves_error_metadata_and_close,
    existing.test_handles_roundtrip_real_worker_with_com_object_return_doubles,
])
def test_original_real_worker_contract_at_same_tick(transport, monkeypatch, tick, scenario):
    module, _, _ = transport
    assert (tick + 600) - tick > 600
    monkeypatch.setattr(module, 'time', SimpleNamespace(monotonic=lambda: tick))
    if os.environ.get('REVIEW_CAP_PROTOTYPE') == '1':
        original = module._SessionProxy._remaining
        monkeypatch.setattr(module._SessionProxy, '_remaining', lambda self: min(600, original(self)))
    scenario(transport, monkeypatch)


def test_cap_preserves_expiry_and_never_increases_positive_budget(monkeypatch):
    module, _ = existing.modules()
    original = module._SessionProxy._remaining
    proxy = object.__new__(module._SessionProxy)
    proxy._deadline = 2158.5460000000003
    for tick in [1558.546, 1558.547, 2000.0, 2158.545]:
        monkeypatch.setattr(module, 'time', SimpleNamespace(monotonic=lambda: tick))
        value = original(proxy)
        assert 0 < min(600, value) <= value
        assert proxy._deadline == 2158.5460000000003
    for tick in [proxy._deadline, proxy._deadline + 1]:
        monkeypatch.setattr(module, 'time', SimpleNamespace(monotonic=lambda: tick))
        with pytest.raises(TimeoutError, match='Session deadline expired'):
            min(600, original(proxy))
