"""Independent guard review: public normalization must match pre-open checks."""
import pytest
from skills.WPSComposer.scripts import document_api as api, office_engines as engines
from test_edit_preflight import mac_ms

@pytest.mark.parametrize('atomic', [True, False])
def test_patch_supplied_structural_op_is_rejected_before_open(mac_ms, monkeypatch, atomic):
    opened=[]
    def open_native(*args, **kwargs):
        opened.append(True)
        raise RuntimeError('native opener reached')
    monkeypatch.setattr(api, 'open_document', open_native)
    patch={'op':'insert','type':'table','position':'start',
           'props':{'rows':1,'cols':1,'data':[['keep']]}}
    with pytest.raises(ValueError, match='Mac Word.*table.*end'):
        api.edit('/source.docx', engine='msoffice', patches=iter([patch]),
                 output='/other.docx', atomic=atomic)
    assert opened == []
